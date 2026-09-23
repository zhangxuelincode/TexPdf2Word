#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Audit and repair table grid structure in a DOCX.

Covers the failure modes seen when LaTeX tables with multicolumn/multirow
headers are converted without proper OOXML span markup:

  Mode A  row has FEWER cells than grid columns because spanning cells were
          emitted without <w:gridSpan> (e.g. a "NAB" header that must cover
          three Pre/Rec/F1 columns). Cells carry a tcW equal to the sum of
          the grid columns they should cover -> infer the span from the
          widths and insert <w:gridSpan>.
  Mode B  row has MORE cells than grid columns because a multirow filler
          cell was emitted in addition to the vMerge continuation cell ->
          remove content-less filler cells until the row matches the grid.
  Mode C  tblGrid has phantom trailing columns that no row covers (Word
          normalizes such tables by stretching the grid) -> truncate them.

The repairer is intentionally conservative:
  * cells with text/images/OMath are never removed,
  * vMerge restart/continue cells are never removed (merges stay intact),
  * tcBorders are never touched, so three-line / full-grid / hline styling
    survives (regression-tested in scripts/tests/test_fix_table_structure.py),
  * rows whose widths cannot be matched confidently are left unchanged and
    reported.

Usage:
  python fix_table_structure.py doc.docx                # audit + repair in place
  python fix_table_structure.py doc.docx --audit-only   # report without changes
  python fix_table_structure.py doc.docx --out new.docx
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile

from lxml import etree

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = '{%s}' % W_NS


def qn(tag: str) -> str:
    return W + tag


# ------------------------------------------------------------------ helpers
def cell_text(tc) -> str:
    return ''.join(t.text or '' for t in tc.iter(qn('t')))


def cell_has_content(tc) -> bool:
    if cell_text(tc).strip():
        return True
    if tc.find('.//' + qn('drawing')) is not None:
        return True
    if tc.find('.//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath') is not None:
        return True
    return False


def cell_span(tc) -> int:
    tcPr = tc.find(qn('tcPr'))
    if tcPr is None:
        return 1
    gs = tcPr.find(qn('gridSpan'))
    return int(gs.get(qn('val'))) if gs is not None else 1


def cell_width_dxa(tc):
    tcPr = tc.find(qn('tcPr'))
    if tcPr is None:
        return None
    w = tcPr.find(qn('tcW'))
    if w is None or w.get(qn('type')) != 'dxa':
        return None
    try:
        return int(w.get(qn('w')))
    except (TypeError, ValueError):
        return None


def cell_is_vmerge(tc) -> bool:
    tcPr = tc.find(qn('tcPr'))
    return tcPr is not None and tcPr.find(qn('vMerge')) is not None


def row_span_sum(tr) -> int:
    return sum(cell_span(tc) for tc in tr.findall(qn('tc')))


def grid_widths(tbl):
    grid = tbl.find(qn('tblGrid'))
    if grid is None:
        return []
    out = []
    for gc in grid.findall(qn('gridCol')):
        try:
            out.append(int(gc.get(qn('w'))))
        except (TypeError, ValueError):
            out.append(0)
    return out


# ------------------------------------------------------------------ audit
def audit_table(tbl):
    """Return a dict describing grid inconsistencies of one w:tbl."""
    widths = grid_widths(tbl)
    ncols = len(widths)
    rows = tbl.findall(qn('tr'))
    issues = []
    for ri, tr in enumerate(rows):
        s = row_span_sum(tr)
        if ncols >= 0 and s != ncols:
            kind = 'extra-cells' if s > ncols else 'missing-span'
            issues.append({'row': ri, 'kind': kind, 'span_sum': s, 'ncols': ncols})
    return {'ncols': ncols, 'nrows': len(rows), 'issues': issues,
            'widths': widths}


def audit_document(root):
    """Audit every table; returns list of {index, first_text, audit}."""
    body = root.find(qn('body'))
    report = []
    for idx, tbl in enumerate(body.iter(qn('tbl'))):
        a = audit_table(tbl)
        if a['issues']:
            report.append({
                'index': idx,
                'first_text': cell_text(tbl)[:40],
                'audit': a,
            })
    return report


# ------------------------------------------------------------------ repair
def _match_span(width, grid_w, pos, remaining, tol=40):
    """Smallest k (1..remaining) with sum(grid_w[pos:pos+k]) within tol of width."""
    if width is None:
        return None
    total = 0
    for k in range(1, remaining + 1):
        if pos + k > len(grid_w):
            break
        total += grid_w[pos + k - 1]
        if abs(total - width) <= max(tol, int(width * 0.02)):
            return k
    return None


def _repair_missing_span(tr, grid_w, ncols):
    """Mode A: add gridSpan inferred from cell widths. Returns #cells fixed."""
    cells = tr.findall(qn('tc'))
    changes = 0
    pos = 0
    plan = []
    for tc in cells:
        if cell_span(tc) != 1:
            return 0  # already has explicit spans -> do not guess
        w = cell_width_dxa(tc)
        k = _match_span(w, grid_w, pos, ncols - pos)
        if k is None:
            return 0
        plan.append((tc, k))
        pos += k
    if pos != ncols:
        return 0
    for tc, k in plan:
        if k > 1:
            tcPr = tc.find(qn('tcPr'))
            if tcPr is None:
                tcPr = etree.SubElement(tc, qn('tcPr'))
                tc.insert(0, tcPr)
            gs = etree.SubElement(tcPr, qn('gridSpan'))
            gs.set(qn('val'), str(k))
            changes += 1
    return changes


def _repair_extra_cells(tr, ncols):
    """Mode B: remove content-less filler cells until span-sum == ncols.
    vMerge cells are never removed. Returns #cells removed."""
    removed = 0
    while row_span_sum(tr) > ncols:
        candidates = []
        for tc in tr.findall(qn('tc')):
            if cell_has_content(tc):
                continue
            if cell_is_vmerge(tc):
                continue  # never break a vertical merge
            if cell_span(tc) != 1:
                continue
            candidates.append(tc)
        if not candidates:
            break
        # drop the LAST safe candidate (filler cells trail the merge cell)
        tr.remove(candidates[-1])
        removed += 1
    return removed


def _truncate_phantom_cols(tbl, widths):
    """Mode C: drop trailing grid columns that no row covers."""
    rows = tbl.findall(qn('tr'))
    if not rows:
        return 0
    max_cover = max(row_span_sum(tr) for tr in rows)
    ncols = len(widths)
    if max_cover >= ncols or max_cover <= 0:
        return 0
    grid = tbl.find(qn('tblGrid'))
    cols = grid.findall(qn('gridCol'))
    for gc in cols[max_cover:]:
        grid.remove(gc)
    return ncols - max_cover


def repair_table(tbl):
    """Repair one w:tbl in place; returns a change report dict."""
    widths = grid_widths(tbl)
    ncols = len(widths)
    report = {'gridspan_added': 0, 'cells_removed': 0, 'gridcols_truncated': 0}
    rows = tbl.findall(qn('tr'))

    # Mode B first: extra cells
    for tr in rows:
        if row_span_sum(tr) > ncols:
            report['cells_removed'] += _repair_extra_cells(tr, ncols)

    # Mode A: missing spans (width-based)
    for tr in rows:
        if row_span_sum(tr) < ncols:
            report['gridspan_added'] += _repair_missing_span(tr, widths, ncols)

    # Mode C: phantom grid columns
    report['gridcols_truncated'] += _truncate_phantom_cols(tbl, widths)
    return report


def repair_document(root):
    body = root.find(qn('body'))
    out = []
    for idx, tbl in enumerate(body.iter(qn('tbl'))):
        rep = repair_table(tbl)
        if any(rep.values()):
            out.append({'index': idx,
                        'first_text': cell_text(tbl)[:40],
                        'changes': rep})
    return out


# ------------------------------------------------------------------ cli
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('--out', default=None)
    ap.add_argument('--audit-only', action='store_true')
    args = ap.parse_args()

    src = os.path.abspath(args.docx)
    dst = os.path.abspath(args.out) if args.out else src

    tmp = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(src) as z:
            names = z.namelist()
            z.extractall(tmp)
        doc_path = os.path.join(tmp, 'word', 'document.xml')
        tree = etree.parse(doc_path)
        root = tree.getroot()

        problems = audit_document(root)
        print(f'[audit] tables with grid issues: {len(problems)}')
        for p in problems:
            print('   -', p['index'], repr(p['first_text']),
                  [(i['row'], i['kind']) for i in p['audit']['issues']])

        if args.audit_only:
            return 0

        changes = repair_document(root)
        print(f'[repair] tables changed: {len(changes)}')
        for c in changes:
            print('   -', c['index'], repr(c['first_text']), c['changes'])

        tree.write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)
        if os.path.exists(dst) and os.path.abspath(dst) != os.path.abspath(src):
            os.remove(dst)
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                zout.write(os.path.join(tmp, n), n)
        print('[info] saved', dst)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return 0


if __name__ == '__main__':
    sys.exit(main())
