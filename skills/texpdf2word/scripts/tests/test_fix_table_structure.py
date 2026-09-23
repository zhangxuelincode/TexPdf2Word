# -*- coding: utf-8 -*-
"""Unit tests for fix_table_structure.py.

Run:  python scripts/tests/test_fix_table_structure.py
or:   python -m unittest discover -s scripts/tests

The regression focus (per real-world failure): repairing gridSpan / extra
filler cells / phantom grid columns must NEVER alter cell borders (three-line,
full-grid) nor break vMerge restart/continue pairing.
"""
import os
import sys
import copy
import unittest
from lxml import etree

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from fix_table_structure import (  # noqa: E402
    W, qn, audit_table, repair_table, row_span_sum, cell_text, cell_span,
)


def el(parent, tag, attrs=None):
    e = etree.SubElement(parent, W + tag)
    if attrs:
        for k, v in attrs.items():
            e.set(W + k, v)
    return e


def make_cell(text='', width=None, span=None, vmerge=None, borders=None,
              tbl=None):
    """borders: {'top': 12, 'bottom': 8, ...} -> single-line tcBorders."""
    tc = el(tbl, 'tc')
    tcPr = el(tc, 'tcPr')
    if width is not None:
        el(tcPr, 'tcW', {'w': str(width), 'type': 'dxa'})
    if span is not None:
        el(tcPr, 'gridSpan', {'val': str(span)})
    if vmerge is not None:
        if vmerge == 'restart':
            el(tcPr, 'vMerge', {'val': 'restart'})
        else:
            el(tcPr, 'vMerge')
    if borders:
        b = el(tcPr, 'tcBorders')
        for side, sz in borders.items():
            el(b, side, {'val': 'single', 'sz': str(sz), 'space': '0',
                         'color': 'auto'})
    p = el(tc, 'p')
    if text:
        el(p, 't').text = text
    return tc


def make_table(grid_widths, rows_spec):
    """rows_spec: list of rows; each row = list of cell kwargs dicts."""
    tbl = etree.Element(W + 'tbl')
    grid = el(tbl, 'tblGrid')
    for w in grid_widths:
        el(grid, 'gridCol', {'w': str(w)})
    for row_spec in rows_spec:
        tr = el(tbl, 'tr')
        for cell_kw in row_spec:
            make_cell(tbl=tr, **cell_kw)
    return tbl


def borders_of(tc):
    tcPr = tc.find(W + 'tcPr')
    b = tcPr.find(W + 'tcBorders') if tcPr is not None else None
    if b is None:
        return {}
    return {child.tag.replace(W, ''): child.get(W + 'sz') for child in b}


def spans_of(tbl, row=0):
    tr = tbl.findall(W + 'tr')[row]
    return [cell_span(tc) for tc in tr.findall(W + 'tc')]


GRID7 = [1295, 1295, 1295, 1295, 1295, 1295, 1300]
THREE_LINE = ({'top': 12}, {'bottom': 8}, {}, {'bottom': 12})


class TestAudit(unittest.TestCase):

    def test_audit_flags_missing_span(self):
        tbl = make_table(GRID7, [
            [{'text': '方法', 'width': 1295}, {'text': 'NAB', 'width': 3885},
             {'text': 'UCR', 'width': 3890}],
            [{'text': 'a'}, {'text': 'b'}, {'text': 'c'}, {'text': 'd'},
             {'text': 'e'}, {'text': 'f'}, {'text': 'g'}],
        ])
        a = audit_table(tbl)
        self.assertEqual(a['issues'][0]['kind'], 'missing-span')

    def test_audit_flags_extra_cells(self):
        tbl = make_table(GRID7, [
            [{'text': 'x'} for _ in range(8)],
        ])
        a = audit_table(tbl)
        self.assertEqual(a['issues'][0]['kind'], 'extra-cells')

    def test_audit_clean_table_no_issues(self):
        tbl = make_table(GRID7, [
            [{'text': chr(97 + i)} for i in range(7)],
        ])
        self.assertEqual(audit_table(tbl)['issues'], [])


class TestRepairMissingSpan(unittest.TestCase):

    def test_gridspan_inferred_from_widths(self):
        tbl = make_table(GRID7, [
            [{'text': '方法', 'width': 1295, 'borders': {'top': 12}},
             {'text': 'NAB', 'width': 3885, 'borders': {'bottom': 8}},
             {'text': 'UCR', 'width': 3890, 'borders': {'bottom': 8}}],
            [{'text': 'Pre', 'width': 1295}, {'text': 'Rec', 'width': 1295},
             {'text': 'F1', 'width': 1295}, {'text': 'Pre', 'width': 1295},
             {'text': 'Rec', 'width': 1295}, {'text': 'F1', 'width': 1295},
             {'text': 'F1', 'width': 1300}],
        ])
        rep = repair_table(tbl)
        self.assertEqual(rep['gridspan_added'], 2)
        self.assertEqual(spans_of(tbl, 0), [1, 3, 3])
        # whole document remains grid-consistent
        for tr in tbl.findall(W + 'tr'):
            self.assertEqual(row_span_sum(tr), 7)

    def test_inference_aborts_on_unmatchable_widths(self):
        # widths that cannot map onto the grid -> row left unchanged
        tbl = make_table(GRID7, [
            [{'text': 'a', 'width': 999}, {'text': 'b', 'width': 12345}],
        ])
        rep = repair_table(tbl)
        self.assertEqual(rep['gridspan_added'], 0)
        self.assertEqual(spans_of(tbl, 0), [1, 1])


class TestRepairExtraCells(unittest.TestCase):

    def test_filler_cell_removed_vmerge_kept(self):
        tbl = make_table(GRID7, [
            [{'text': '方法', 'width': 1295, 'vmerge': 'restart'},
             {'text': 'NAB', 'width': 3885, 'span': 3},
             {'text': 'UCR', 'width': 3890, 'span': 3}],
            [{'width': 1295, 'vmerge': 'continue'},
             {'width': 1295},
             {'text': 'Pre'}, {'text': 'Rec'}, {'text': 'F1'},
             {'text': 'Pre'}, {'text': 'Rec'}, {'text': 'F1'}],
        ])
        rep = repair_table(tbl)
        self.assertEqual(rep['cells_removed'], 1)
        row1 = tbl.findall(W + 'tr')[1]
        cells = row1.findall(W + 'tc')
        self.assertEqual(len(cells), 7)
        # vMerge continuation survived
        self.assertTrue(cells[0].find(W + 'tcPr').find(W + 'vMerge') is not None)
        self.assertEqual(row_span_sum(row1), 7)

    def test_content_cells_never_removed(self):
        tbl = make_table(GRID7, [
            [{'text': str(i)} for i in range(8)],
        ])
        rep = repair_table(tbl)
        self.assertEqual(rep['cells_removed'], 0)
        self.assertEqual(len(tbl.findall(W + 'tr')[0].findall(W + 'tc')), 8)


class TestPhantomColumns(unittest.TestCase):

    def test_trailing_phantom_gridcol_truncated(self):
        tbl = make_table(GRID7 + [346], [
            [{'text': chr(97 + i)} for i in range(7)],
        ])
        rep = repair_table(tbl)
        self.assertEqual(rep['gridcols_truncated'], 1)
        grid = tbl.find(W + 'tblGrid')
        self.assertEqual(len(grid.findall(W + 'gridCol')), 7)


class TestBordersSurvive(unittest.TestCase):
    """The user-facing regression: span fixes must not touch 横线/竖线."""

    def test_three_line_borders_preserved(self):
        tbl = make_table(GRID7, [
            [{'text': '方法', 'width': 1295, 'borders': {'top': 12}},
             {'text': 'NAB', 'width': 3885, 'borders': {'bottom': 8}},
             {'text': 'UCR', 'width': 3890, 'borders': {'bottom': 8}}],
            [{'text': 'Pre', 'width': 1295}, {'text': 'Rec', 'width': 1295},
             {'text': 'F1', 'width': 1295}, {'text': 'Pre', 'width': 1295},
             {'text': 'Rec', 'width': 1295}, {'text': 'F1', 'width': 1295},
             {'text': 'F1', 'width': 1300, 'borders': {'bottom': 12}}],
        ])
        before = [borders_of(tc) for tc in tbl.iter(W + 'tc')]
        repair_table(tbl)
        after = [borders_of(tc) for tc in tbl.iter(W + 'tc')]
        self.assertEqual(before, after)
        # heavy top line still on first cell, heavy bottom on last
        self.assertEqual(after[0].get('top'), '12')
        self.assertEqual(after[-1].get('bottom'), '12')

    def test_full_grid_borders_preserved(self):
        full = {s: 4 for s in ('top', 'left', 'bottom', 'right')}
        tbl = make_table([1000, 1000, 1000], [
            [{'text': 'h1', 'width': 1000, 'borders': full},
             {'text': 'h2', 'width': 1000, 'borders': full},
             {'text': 'h3', 'width': 1000, 'borders': full}],
            [{'text': 'a', 'width': 1000, 'borders': full},
             {'width': 1000, 'borders': full},
             {'text': 'c', 'width': 1000, 'borders': full},
             {'width': 1000, 'borders': full}],
        ])
        before = [borders_of(tc) for tc in tbl.iter(W + 'tc')]
        rep = repair_table(tbl)
        after = [borders_of(tc) for tc in tbl.iter(W + 'tc')]
        self.assertEqual(rep['cells_removed'], 1)
        self.assertEqual(before[:3], after[:3])  # untouched row keeps borders

    def test_vmerge_pairing_preserved(self):
        tbl = make_table(GRID7, [
            [{'text': '方法', 'width': 1295, 'vmerge': 'restart',
              'borders': {'top': 12}},
             {'text': 'NAB', 'width': 3885, 'span': 3},
             {'text': 'UCR', 'width': 3890, 'span': 3}],
            [{'width': 1295, 'vmerge': 'continue',
              'borders': {'bottom': 8}},
             {'text': 'Pre'}, {'text': 'Rec'}, {'text': 'F1'},
             {'text': 'Pre'}, {'text': 'Rec'},
             {'text': 'F1', 'width': 1300, 'borders': {'bottom': 12}}],
        ])
        repair_table(tbl)
        rows = tbl.findall(W + 'tr')
        r0c0 = rows[0].findall(W + 'tc')[0]
        r1c0 = rows[1].findall(W + 'tc')[0]
        self.assertEqual(r0c0.find(W + 'tcPr').find(W + 'vMerge').get(W + 'val'),
                         'restart')
        self.assertTrue(r1c0.find(W + 'tcPr').find(W + 'vMerge') is not None)
        self.assertEqual(borders_of(r1c0).get('bottom'), '8')


class TestIdempotency(unittest.TestCase):

    def test_correct_table_untouched(self):
        tbl = make_table(GRID7, [
            [{'text': '方法', 'width': 1295, 'vmerge': 'restart'},
             {'text': 'NAB', 'width': 3885, 'span': 3},
             {'text': 'UCR', 'width': 3890, 'span': 3}],
            [{'width': 1295, 'vmerge': 'continue'},
             {'text': 'Pre'}, {'text': 'Rec'}, {'text': 'F1'},
             {'text': 'Pre'}, {'text': 'Rec'}, {'text': 'F1'}],
        ])
        before = etree.tostring(tbl)
        rep = repair_table(tbl)
        self.assertEqual(rep, {'gridspan_added': 0, 'cells_removed': 0,
                               'gridcols_truncated': 0})
        self.assertEqual(etree.tostring(tbl), before)


if __name__ == '__main__':
    unittest.main(verbosity=2)
