#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Add internal cross-reference hyperlinks to a thesis-style DOCX.

Implements clickable navigation for the numbering conventions used by
Chinese-thesis Word templates (and this skill's pipelines):

  in-text token                 ->  jumps to bookmark at target
  --------------                    ---------------------------
  图N-M / 图N-M(a)              ->  figure caption paragraph (figref_N_M)
  表N-M                         ->  table caption paragraph  (tabref_N_M)
  式N-M / (N-M) / 式(N-M)       ->  equation paragraph       (eqref_N_M)
  代码清单N-M                   ->  listing title paragraph  (lstref_N_M)
  superscript/plain [N]         ->  reference entry paragraph (ref_N)

Pass 1 bookmarks every target paragraph (captions starting with 图N-M /
表N-M / 代码清单N-M, equation paragraphs whose visible text is "(N-M)",
reference entries "[N] ..."). Pass 2 merges adjacent plain runs, finds the
tokens in body text and table cells, splits the runs and wraps each token
in <w:hyperlink w:anchor="...">. Tokens already inside a hyperlink (e.g.
citations produced by a pipeline) and tokens inside OMML math are left
untouched. Bookmark names use "_" because Word bookmark names cannot
contain "-".

Usage:
  python add_crossref_hyperlinks.py final.docx            # in place
  python add_crossref_hyperlinks.py final.docx --out new.docx

Run BEFORE finalize_word_docx.py (Word converts the inserted
w:hyperlink elements into HYPERLINK fields on save).
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
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def qn(tag: str) -> str:
    return W + tag


TOKEN_RE = re.compile(
    r'图\d+-\d+(?:\([a-h]\))?'
    r'|表\d+-\d+(?:\([a-h]\))?'
    r'|式\(\d+-\d+\)'
    r'|式\d+-\d+'
    r'|代码清单\d+-\d+'
    r'|\(\d+-\d+\)'
)

CITE_RE = re.compile(r'\[\d+(?:-\d+|,\d+)*\]')


def para_text(p) -> str:
    return ''.join(t.text or '' for t in p.iter(qn('t')))


def is_pure_text_run(r) -> bool:
    """Run whose content is a single w:t (no br/tab/drawing/etc.)."""
    if r.tag != qn('r'):
        return False
    kids = [c for c in r if c.tag != qn('rPr')]
    return len(kids) == 1 and kids[0].tag == qn('t')


def run_text(r) -> str:
    t = r.find(qn('t'))
    return t.text or '' if t is not None else ''


def make_text_run(text: str, rPr=None):
    r = etree.Element(qn('r'))
    if rPr is not None:
        r.append(etree.fromstring(etree.tostring(rPr)))
    t = etree.SubElement(r, qn('t'))
    t.set(XML_SPACE, 'preserve')
    t.text = text
    return r


def make_hyperlink(anchor: str, text: str, rPr=None):
    hl = etree.Element(qn('hyperlink'))
    hl.set(qn('anchor'), anchor)
    hl.set(qn('history'), '1')
    hl.append(make_text_run(text, rPr))
    return hl


def canonical_rpr(r):
    rPr = r.find(qn('rPr'))
    if rPr is None:
        return b'<none/>'
    return etree.tostring(rPr)


def merge_adjacent_runs(p):
    """Merge adjacent pure-text runs with identical rPr so tokens that span
    runs (e.g. 图 + 2-2) become matchable."""
    changed = 0
    children = list(p)
    i = 0
    while i < len(children) - 1:
        a, b = children[i], children[i + 1]
        if (a.tag == qn('r') and b.tag == qn('r')
                and is_pure_text_run(a) and is_pure_text_run(b)
                and canonical_rpr(a) == canonical_rpr(b)):
            ta, tb = a.find(qn('t')), b.find(qn('t'))
            ta.text = (ta.text or '') + (tb.text or '')
            p.remove(b)
            children = list(p)
            changed += 1
            # stay at i, try to extend the merge
        else:
            i += 1
    return changed


def next_bookmark_id(root) -> int:
    ids = [int(b.get(qn('id'))) for b in root.iter(qn('bookmarkStart'))
           if b.get(qn('id')) and b.get(qn('id')).isdigit()]
    return (max(ids) + 1) if ids else 1


class Targets:
    def __init__(self):
        self.by_key = {}      # (kind, 'N-M') -> bookmark name
        self.paras = {}       # id(paragraph element) -> (kind, 'N-M')

    def add(self, kind, num, p, name):
        # first occurrence wins: the real caption/entry precedes stray mentions
        if (kind, num) in self.by_key:
            return
        self.by_key[(kind, num)] = name
        self.paras[id(p)] = (kind, num)


def kind_num_from_token(tok: str):
    """token -> (kind, num) following the naming rules above."""
    m = re.match(r'^图(\d+-\d+)', tok)
    if m:
        return 'fig', m.group(1)
    m = re.match(r'^表(\d+-\d+)', tok)
    if m:
        return 'tab', m.group(1)
    m = re.match(r'^式\(?(\d+-\d+)\)?$', tok)
    if m:
        return 'eq', m.group(1)
    m = re.match(r'^代码清单(\d+-\d+)$', tok)
    if m:
        return 'lst', m.group(1)
    m = re.match(r'^\((\d+-\d+)\)$', tok)
    if m:
        return 'eq', m.group(1)
    return None, None


def bookmark_name(kind, num):
    prefix = {'fig': 'figref', 'tab': 'tabref', 'eq': 'eqref',
              'lst': 'lstref', 'bib': 'ref'}[kind]
    return f'{prefix}_{num.replace("-", "_")}'


def add_bookmark(p, bm_id, name):
    bs = etree.Element(qn('bookmarkStart'))
    bs.set(qn('id'), str(bm_id))
    bs.set(qn('name'), name)
    be = etree.Element(qn('bookmarkEnd'))
    be.set(qn('id'), str(bm_id))
    pPr = p.find(qn('pPr'))
    idx = 1 if pPr is not None else 0
    p.insert(idx, bs)
    p.append(be)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('--out', default=None,
                    help='write result to a new file (default: in place)')
    ap.add_argument('--report', default=None,
                    help='write a JSON report of created links')
    ap.add_argument('--link-citations', action='store_true',
                    help='also wrap plain [N] citation tokens in hyperlinks '
                         'to the ref_N bookmark of reference entry [N]')
    ap.add_argument('--fix-citation-fields', action='store_true',
                    help='repair HYPERLINK fields that Word serialized as '
                         '`HYPERLINK \\l "x" \\h`: the stray \\h is parsed as '
                         "an external Address='h' and silently breaks "
                         'Ctrl+Click navigation. Safe for PAGEREF \\h fields, '
                         'whose \\h switch is legitimate.')
    args = ap.parse_args()

    src = os.path.abspath(args.docx)
    dst = os.path.abspath(args.out) if args.out else src

    tmpdir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(src) as z:
            names = z.namelist()
            z.extractall(tmpdir)
        doc_path = os.path.join(tmpdir, 'word', 'document.xml')
        tree = etree.parse(doc_path)
        root = tree.getroot()
        body = root.find(qn('body'))

        # ---------------- pass 0: existing bookmarks registry
        existing = {b.get(qn('name')) for b in root.iter(qn('bookmarkStart'))}
        targets = Targets()
        bm_id = next_bookmark_id(root)

        # ---------------- pass -1: repair broken HYPERLINK fields (stray \h)
        fixed_fields = 0
        if args.fix_citation_fields:
            for it in root.iter(qn('instrText')):
                txt = it.text or ''
                if 'HYPERLINK' not in txt:
                    continue
                new = re.sub(
                    r'(HYPERLINK\s+\\l\s+(?:&quot;|")[^&"]*?(?:&quot;|"))\s*\\h\s*',
                    r'\1 ', txt)
                if new != txt:
                    it.text = new
                    fixed_fields += 1
        print(f'[info] HYPERLINK fields repaired (stray \\h removed): {fixed_fields}')

        # ---------------- pass 1: bookmark targets
        # A caption target must have whitespace right after the number
        # ("图2-1 标题"); a body mention like "图2-1从波形形态…" does not.
        # Reference entries must live outside tables.
        CAP_NUM_RE = {
            'fig': re.compile(r'^(图)(\d+-\d+)(?=[\s\u3000])'),
            'tab': re.compile(r'^(表)(\d+-\d+)(?=[\s\u3000])'),
            'lst': re.compile(r'^(代码清单)(\d+-\d+)(?=[\s\u3000])'),
        }
        stats = {'fig': 0, 'tab': 0, 'eq': 0, 'lst': 0, 'bib': 0}
        for p in list(body.iter(qn('p'))):
            text = para_text(p).strip()
            if not text:
                continue
            hit = None
            for kind, rx in CAP_NUM_RE.items():
                m = rx.match(text)
                if m:
                    hit = (kind, m.group(2))
                    break
            if hit is None and re.match(r'^\(\d+-\d+\)$', text):
                hit = ('eq', text.strip('()'))
            if hit is None and text.startswith('['):
                m = re.match(r'^\[(\d+)\]', text)
                if m:
                    in_table = any(anc.tag == qn('tbl')
                                   for anc in p.iterancestors())
                    if not in_table:
                        hit = ('bib', m.group(1))
            if hit is None:
                continue
            kind, num = hit
            name = bookmark_name(kind, num)
            if name not in existing:
                add_bookmark(p, bm_id, name)
                existing.add(name)
                bm_id += 1
            targets.add(kind, num, p, name)
            stats[kind] += 1

        # ---------------- pass 2: link tokens
        linked = {'fig': 0, 'tab': 0, 'eq': 0, 'lst': 0, 'bib': 0}
        missed = []
        for p in list(body.iter(qn('p'))):
            merge_adjacent_runs(p)
            tgt = targets.paras.get(id(p))
            for r in list(p):
                if r.tag != qn('r') or not is_pure_text_run(r):
                    continue
                text = run_text(r)
                if not text:
                    continue
                matches = list(TOKEN_RE.finditer(text))
                cites = list(CITE_RE.finditer(text)) if args.link_citations else []
                if not matches and not cites:
                    continue
                # build combined edit list
                edits = []
                for mch in matches:
                    tok = mch.group(0)
                    kind, num = kind_num_from_token(tok)
                    if tgt and (kind, num) == tgt:
                        continue  # caption's own number = the target itself
                    bm = targets.by_key.get((kind, num))
                    edits.append((mch.start(), mch.end(), tok, bm, kind, num))
                if args.link_citations:
                    for mch in cites:
                        if any(not (mch.end() <= s or mch.start() >= e)
                               for s, e, *_ in edits):
                            continue
                        num = mch.group(0)[1:-1]
                        bm = targets.by_key.get(('bib', num.split(',')[0].split('-')[0]))
                        edits.append((mch.start(), mch.end(), mch.group(0), bm, 'bib', num))
                if not edits:
                    continue
                edits.sort()
                rPr = r.find(qn('rPr'))
                parent = p
                idx = parent.index(r)
                new_elems = []
                pos = 0
                for s, e, tok, bm, kind, num in edits:
                    if s > pos:
                        new_elems.append(make_text_run(text[pos:s], rPr))
                    if bm:
                        new_elems.append(make_hyperlink(bm, tok, rPr))
                        linked[kind] = linked.get(kind, 0) + 1
                    else:
                        new_elems.append(make_text_run(tok, rPr))
                        missed.append(tok)
                    pos = e
                if pos < len(text):
                    new_elems.append(make_text_run(text[pos:], rPr))
                for el in new_elems:
                    parent.insert(idx, el)
                    idx += 1
                parent.remove(r)

        tree.write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)

        # ---------------- rezip
        if os.path.exists(dst) and os.path.abspath(dst) != os.path.abspath(src):
            os.remove(dst)
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                zout.write(os.path.join(tmpdir, name), name)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    print(f'[info] targets bookmarked: {stats}')
    print(f'[info] hyperlinks created: {linked}')
    if missed:
        print(f'[warn] {len(missed)} tokens without targets (left as plain text): '
              f'{sorted(set(missed))[:20]}')
    print('[info] saved', dst)
    return 0


if __name__ == '__main__':
    sys.exit(main())
