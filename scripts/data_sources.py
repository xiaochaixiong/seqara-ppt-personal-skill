"""Auditable table-cell ledgers and comparable metric calculations.

This module does not infer a business meaning from a nearby table heading.  A
metric author must provide the population, measure, period and value type;
the extractor only preserves what the source file actually contains.
"""
from __future__ import annotations

import csv
from datetime import date, datetime, time
from hashlib import sha256
import io
import json
import math
from pathlib import Path
import re
from zipfile import ZipFile

from lxml import etree

VALUE_TYPES = {'actual', 'target', 'estimate', 'comparison'}
OPERATIONS = {'sum', 'mean', 'difference', 'ratio_percent', 'growth_percent',
              'percentage_point_change', 'target_progress_percent'}
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _sha(path):
    h = sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def _json_value(value):
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def _numeric_text(value):
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value) if math.isfinite(value) else None
    text = str(value).strip().replace(',', '').replace('，', '')
    text = text.replace('−', '-').replace('％', '%')
    if not re.fullmatch(r'[+-]?(?:\d+(?:\.\d*)?|\.\d+)%?', text):
        return None
    return float(text[:-1] if text.endswith('%') else text)


def _cell_state(raw, effective, formula=False):
    if formula and effective is None:
        return 'formula_uncached'
    if effective is None or effective == '':
        return 'blank'
    number = _numeric_text(effective)
    if number is not None:
        return 'zero' if number == 0 else 'number'
    return 'text'


def _extract_csv(path):
    raw = Path(path).read_bytes()
    try:
        content = raw.decode('utf-8-sig')
        encoding = 'utf-8-sig'
    except UnicodeDecodeError:
        content = raw.decode('gb18030')
        encoding = 'gb18030'
    try:
        dialect = csv.Sniffer().sniff(content[:8192], delimiters=',;\t')
        delimiter = dialect.delimiter
    except csv.Error:
        delimiter = ','
    rows = list(csv.reader(io.StringIO(content), delimiter=delimiter))
    cells = []
    for ri, row in enumerate(rows, 1):
        for ci, value in enumerate(row, 1):
            cells.append({'id': f'r{ri:06d}c{ci:04d}', 'location': {'row': ri, 'column': ci},
                          'raw': value, 'value': value, 'formula': None, 'cached_value': None,
                          'number_format': None, 'state': _cell_state(value, value),
                          'numeric_value': _numeric_text(value)})
    return cells, {'encoding': encoding, 'delimiter': delimiter, 'row_count': len(rows)}


def _extract_xlsx(path):
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError('XLSX data extraction requires openpyxl from requirements.txt') from exc
    formulas = load_workbook(path, data_only=False, read_only=False)
    caches = load_workbook(path, data_only=True, read_only=False)
    cells = []
    try:
        for ws in formulas.worksheets:
            cache = caches[ws.title]
            # The stored cell map avoids walking a vast formatted-but-empty rectangle.
            for (ri, ci), cell in sorted(ws._cells.items()):
                raw = cell.value
                formula = raw if cell.data_type == 'f' else None
                cached = cache.cell(ri, ci).value if formula else None
                effective = cached if formula else raw
                raw = _json_value(raw)
                effective = _json_value(effective)
                cached = _json_value(cached)
                state = _cell_state(raw, effective, formula is not None)
                numeric = _numeric_text(effective) if state in ('zero', 'number') else None
                fmt = cell.number_format
                # Excel stores 20% as 0.2.  Keep that raw value, while exposing
                # the visible percentage as the metric value for a % unit.
                if numeric is not None and '%' in fmt and isinstance(effective, (int, float)):
                    numeric *= 100
                cells.append({'id': f'{ws.title}!{cell.coordinate}',
                              'location': {'sheet': ws.title, 'cell': cell.coordinate,
                                           'row': ri, 'column': ci},
                              'raw': raw, 'value': effective, 'formula': formula,
                              'cached_value': cached, 'number_format': fmt,
                              'state': state, 'numeric_value': numeric})
    finally:
        formulas.close()
        caches.close()
    return cells, {'sheets': list(caches.sheetnames)}


def _extract_docx_tables(path):
    with ZipFile(path) as z:
        doc = etree.fromstring(z.read('word/document.xml'))
    body = doc.find(f'{{{W}}}body')
    cells = []
    table_no = 0
    # Iterate all tables, including those embedded in a cell; each occurrence
    # receives its own deterministic table/row/column address.
    for table in body.iter(f'{{{W}}}tbl'):
        table_no += 1
        for ri, row in enumerate(table.findall(f'{{{W}}}tr'), 1):
            for ci, cell in enumerate(row.findall(f'{{{W}}}tc'), 1):
                paragraphs = []
                for para in cell.findall(f'{{{W}}}p'):
                    text = ''.join(para.itertext(f'{{{W}}}t')).strip()
                    if text:
                        paragraphs.append(text)
                value = '\n'.join(paragraphs)
                span = cell.find(f'{{{W}}}tcPr/{{{W}}}gridSpan')
                cells.append({'id': f't{table_no:03d}r{ri:04d}c{ci:03d}',
                              'location': {'table': table_no, 'row': ri, 'column': ci,
                                           'grid_span': int(span.get(f'{{{W}}}val')) if span is not None else 1},
                              'raw': value, 'value': value, 'formula': None,
                              'cached_value': None, 'number_format': None,
                              'state': _cell_state(value, value),
                              'numeric_value': _numeric_text(value)})
    return cells, {'table_count': table_no}


def extract_data(path, source_id):
    """Extract source cells without interpreting headings or recasting blanks as zero."""
    path = Path(path).resolve()
    suffix = path.suffix.lower()
    if suffix == '.xlsx':
        cells, detail = _extract_xlsx(path)
    elif suffix == '.csv':
        cells, detail = _extract_csv(path)
    elif suffix == '.docx':
        cells, detail = _extract_docx_tables(path)
    else:
        raise ValueError('Data sources support XLSX, CSV, and DOCX tables')
    if not isinstance(source_id, str) or not source_id.strip() or ':' in source_id:
        raise ValueError('Data source ID must be nonempty and cannot contain a colon')
    return {'schema_version': 1, 'kind': 'data_source', 'id': source_id, 'role': 'data',
            'path': str(path), 'sha256': _sha(path), 'format': suffix[1:],
            'detail': detail, 'cells': cells}


def load_data_sources(paths, base, existing_ids=()):
    """Re-extract the original data to detect a changed or forged ledger."""
    ledgers = {}
    refs = {}
    cells = {}
    inputs = []
    used_ids = set(existing_ids)
    for ledger_path in paths:
        path = (Path(base) / ledger_path).resolve()
        ledger = json.loads(path.read_text(encoding='utf-8'))
        sid = ledger.get('id')
        if sid in used_ids:
            raise ValueError('Duplicate source id: ' + str(sid))
        used_ids.add(sid)
        original = (path.parent / ledger['path']).resolve()
        actual = extract_data(original, sid)
        for field in ('schema_version', 'kind', 'id', 'role', 'sha256', 'format', 'detail', 'cells'):
            if ledger.get(field) != actual[field]:
                raise ValueError(f'Data source ledger differs from source: {sid} ({field})')
        if len({c['id'] for c in ledger['cells']}) != len(ledger['cells']):
            raise ValueError('Duplicate data cell id in ' + sid)
        ledgers[sid] = ledger
        for cell in ledger['cells']:
            ref = f"{sid}:{cell['id']}"
            cells[ref] = cell
            refs[ref] = '' if cell['value'] is None else str(cell['value'])
        inputs.extend([{'path': str(original), 'sha256': ledger['sha256']},
                       {'path': str(path), 'sha256': _sha(path)}])
    return ledgers, refs, cells, inputs


def _range_cell_ids(spec):
    """Expand a declared contiguous column, never a hand-picked cell list."""
    source = spec.get('source_id')
    first, last, column = spec.get('first_row'), spec.get('last_row'), spec.get('column')
    if not isinstance(source, str) or not source:
        raise ValueError('Table query needs source_id')
    if any(isinstance(x, bool) or not isinstance(x, int) or x < 1 for x in (first, last, column)) or first > last:
        raise ValueError('Table query needs valid row range and column')
    if last - first > 100000:
        raise ValueError('Table query exceeds 100001 rows')
    kind = spec.get('format')
    if kind == 'csv':
        return [f'{source}:r{row:06d}c{column:04d}' for row in range(first, last + 1)]
    if kind == 'docx':
        table = spec.get('table')
        if isinstance(table, bool) or not isinstance(table, int) or table < 1:
            raise ValueError('DOCX table query needs table number')
        return [f'{source}:t{table:03d}r{row:04d}c{column:03d}' for row in range(first, last + 1)]
    if kind == 'xlsx':
        sheet = spec.get('sheet')
        if not isinstance(sheet, str) or not sheet:
            raise ValueError('XLSX table query needs sheet name')
        from openpyxl.utils import get_column_letter
        return [f'{source}:{sheet}!{get_column_letter(column)}{row}' for row in range(first, last + 1)]
    raise ValueError('Table query format must be csv, docx or xlsx')


def evaluate_metrics(specs, cells):
    """Evaluate explicit metrics; no business meaning is inferred from a table."""
    by_id = {m['id']: m for m in specs}
    if len(by_id) != len(specs):
        raise ValueError('Duplicate metric id')
    for metric in specs:
        if not isinstance(metric.get('id'), str) or not metric['id'].strip():
            raise ValueError('Metric needs an id')
    done = {}
    active = set()

    def required(m, key):
        value = m.get(key)
        if not isinstance(value, str) or not value.strip():
            raise ValueError(f"Metric {m['id']} requires {key}")
        return value

    def finite(value, mid):
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
            raise ValueError('Metric ' + mid + ' requires a finite numeric value')
        return float(value)

    def same(operands, keys, mid):
        for key in keys:
            if len({m[key] for m in operands}) != 1:
                raise ValueError(f'Metric {mid}: incompatible {key} among operands')

    def evaluate(mid):
        if mid in done:
            return done[mid]
        if mid in active:
            raise ValueError('Cyclic derived metrics')
        if mid not in by_id:
            raise ValueError('Unknown metric ' + mid)
        active.add(mid)
        m = by_id[mid]
        for key in ('unit', 'period', 'population', 'measure', 'value_type'):
            required(m, key)
        if m['value_type'] not in VALUE_TYPES:
            raise ValueError('Invalid metric value_type: ' + mid)
        kind = m.get('kind')
        if kind == 'source':
            ref = required(m, 'source_ref')
            if ref not in cells:
                raise ValueError('Unknown metric source cell: ' + ref)
            cell = cells[ref]
            if m.get('value') is None:
                if cell['state'] not in ('blank', 'formula_uncached', 'text') or not m.get('missing_reason'):
                    raise ValueError('Missing metric needs blank/uncached/non-numeric source and reason: ' + mid)
                value = None
            else:
                value = finite(m['value'], mid)
                numeric = cell['numeric_value']
                if numeric is None or not math.isclose(value, float(numeric), rel_tol=1e-9, abs_tol=1e-9):
                    raise ValueError('Metric differs from source cell: ' + mid)
            if m['unit'] == '%' and not (m.get('denominator') or m.get('denominator_description')):
                raise ValueError('Source percentage requires denominator: ' + mid)
        elif kind == 'table_count':
            query = m.get('query')
            if not isinstance(query, dict) or not m.get('count_basis'):
                raise ValueError('Table count needs contiguous query and count_basis: ' + mid)
            refs = _range_cell_ids(query)
            missing = [ref for ref in refs if ref not in cells]
            if missing:
                raise ValueError('Table query contains missing cells: ' + ', '.join(missing[:3]))
            equals = query.get('equals')
            starts = query.get('starts_with')
            nonblank = query.get('nonblank') is True
            if sum((equals is not None, starts is not None, nonblank)) != 1:
                raise ValueError('Table count needs one of equals, starts_with or nonblank: ' + mid)
            if equals is not None or starts is not None:
                choices = equals if equals is not None else starts
                if not isinstance(choices, list) or not choices or any(not isinstance(x, str) or not x for x in choices):
                    raise ValueError('Table count match values must be a nonempty string list: ' + mid)
                if len(set(choices)) != len(choices):
                    raise ValueError('Table count has duplicate match values: ' + mid)
                if equals is not None:
                    result = sum(str(cells[ref]['value']).strip() in equals for ref in refs)
                else:
                    result = sum(any(str(cells[ref]['value']).strip().startswith(term) for term in starts)
                                 for ref in refs)
            else:
                result = sum(cells[ref]['state'] not in ('blank', 'formula_uncached') for ref in refs)
            value = finite(m.get('value'), mid)
            if value != result:
                raise ValueError('Incorrect table count: ' + mid)
        elif kind == 'derived':
            ids = m.get('operands')
            if not isinstance(ids, list) or not ids or any(not isinstance(x, str) for x in ids):
                raise ValueError('Derived metric needs operand IDs: ' + mid)
            operands = [evaluate(x) for x in ids]
            if any(x['value'] is None for x in operands):
                raise ValueError('Missing operand in derived metric: ' + mid)
            op = m.get('operation')
            if op not in OPERATIONS:
                raise ValueError('Unsupported metric operation: ' + str(op))
            if op in ('sum', 'mean'):
                if len(operands) < 2:
                    raise ValueError('Aggregation needs at least two operands: ' + mid)
                same(operands, ('unit', 'period', 'population', 'measure', 'value_type'), mid)
                slices = [x.get('slice') for x in operands]
                if not m.get('aggregation_basis') or any(not x for x in slices) or len(set(slices)) != len(slices):
                    raise ValueError('Aggregation needs basis and distinct operand slices: ' + mid)
                if any(m[k] != operands[0][k] for k in ('unit', 'period', 'population', 'measure', 'value_type')):
                    raise ValueError('Aggregation output scope differs from operands: ' + mid)
                result = sum(x['value'] for x in operands)
                if op == 'mean':
                    result /= len(operands)
            elif op == 'difference':
                if len(operands) != 2 or m.get('comparison') not in ('category', 'period', 'target'):
                    raise ValueError('Difference needs two operands and comparison kind: ' + mid)
                same(operands, ('unit', 'population', 'measure'), mid)
                if m['population'] != operands[0]['population'] or m['measure'] != operands[0]['measure'] or m['value_type'] != 'comparison':
                    raise ValueError('Difference output needs matching scope and comparison value_type: ' + mid)
                if m['comparison'] == 'period':
                    same(operands, ('value_type',), mid)
                    if not m.get('comparison_basis') or operands[0]['period'] == operands[1]['period'] or m['period'] != operands[0]['period']:
                        raise ValueError('Period difference requires distinct periods and current period output: ' + mid)
                elif m['comparison'] == 'target':
                    if [x['value_type'] for x in operands] != ['actual', 'target'] or operands[0]['period'] != operands[1]['period']:
                        raise ValueError('Target difference requires actual and target in one period: ' + mid)
                    if m['period'] != operands[0]['period']:
                        raise ValueError('Target difference output period mismatch: ' + mid)
                else:
                    same(operands, ('period', 'value_type'), mid)
                    if m['period'] != operands[0]['period'] or not operands[0].get('slice') or not operands[1].get('slice') or operands[0]['slice'] == operands[1]['slice']:
                        raise ValueError('Category difference requires distinct slices in one period: ' + mid)
                if m['unit'] != operands[0]['unit']:
                    raise ValueError('Difference unit mismatch: ' + mid)
                result = operands[0]['value'] - operands[1]['value']
            elif op == 'ratio_percent':
                if len(operands) != 2 or m.get('denominator') != ids[1] or m.get('relation') not in ('subset', 'rate'):
                    raise ValueError('Ratio needs numerator, explicit denominator and relation: ' + mid)
                same(operands, ('unit', 'period', 'population', 'value_type'), mid)
                if operands[1]['value'] <= 0:
                    raise ValueError('Zero denominator: ' + mid)
                result = operands[0]['value'] / operands[1]['value'] * 100
                if m['relation'] == 'subset' and not 0 <= operands[0]['value'] <= operands[1]['value']:
                    raise ValueError('Subset ratio outside 0–100%: ' + mid)
                if (m['unit'] != '%' or m['period'] != operands[0]['period'] or
                    m['population'] != operands[0]['population'] or m['value_type'] != operands[0]['value_type']):
                    raise ValueError('Ratio output scope mismatch: ' + mid)
            elif op in ('growth_percent', 'percentage_point_change'):
                if len(operands) != 2 or not m.get('comparison_basis'):
                    raise ValueError('Period comparison needs two operands and comparison_basis: ' + mid)
                same(operands, ('unit', 'population', 'measure', 'value_type'), mid)
                if (operands[0]['period'] == operands[1]['period'] or m['period'] != operands[0]['period'] or
                    m['population'] != operands[0]['population'] or m['value_type'] != 'comparison'):
                    raise ValueError('Period comparison needs distinct periods and current-period output: ' + mid)
                if op == 'growth_percent':
                    if operands[1]['value'] <= 0 or m['unit'] != '%':
                        raise ValueError('Growth needs positive baseline and % unit: ' + mid)
                    result = (operands[0]['value'] / operands[1]['value'] - 1) * 100
                else:
                    if operands[0]['unit'] != '%' or m['unit'] != '百分点':
                        raise ValueError('Percentage-point change requires % operands and point unit: ' + mid)
                    result = operands[0]['value'] - operands[1]['value']
            else:  # target_progress_percent
                if len(operands) != 2 or m.get('denominator') != ids[1]:
                    raise ValueError('Target progress needs actual and explicit target denominator: ' + mid)
                same(operands, ('unit', 'period', 'population', 'measure'), mid)
                if [x['value_type'] for x in operands] != ['actual', 'target'] or operands[1]['value'] <= 0:
                    raise ValueError('Target progress needs actual, positive target: ' + mid)
                if (m['unit'] != '%' or m['period'] != operands[0]['period'] or
                    m['population'] != operands[0]['population'] or m['value_type'] != 'comparison'):
                    raise ValueError('Target progress output scope mismatch: ' + mid)
                result = operands[0]['value'] / operands[1]['value'] * 100
            value = finite(m.get('value'), mid)
            decimals = m.get('decimals', 1)
            if not isinstance(decimals, int) or isinstance(decimals, bool) or not 0 <= decimals <= 8:
                raise ValueError('Metric decimals must be 0–8: ' + mid)
            if not math.isclose(round(result, decimals), value, rel_tol=0, abs_tol=1e-8):
                raise ValueError('Incorrect derived metric value: ' + mid)
        else:
            raise ValueError('Metric kind must be source or derived: ' + mid)
        result = dict(m)
        result['value'] = value
        result['id'] = mid
        done[mid] = result
        active.remove(mid)
        return result

    for mid in by_id:
        evaluate(mid)
    for mid, m in done.items():
        denominator = m.get('denominator')
        if denominator and denominator not in done:
            raise ValueError('Unknown metric denominator: ' + mid)
        if denominator and done[denominator]['value'] in (None, 0):
            raise ValueError('Missing or zero metric denominator: ' + mid)
        if denominator and (done[denominator]['period'] != m['period'] or done[denominator]['population'] != m['population']):
            raise ValueError('Metric denominator scope mismatch: ' + mid)
    return done


def metric_report(metrics):
    return [{'id': m['id'], 'value': m['value'], 'unit': m['unit'],
             'period': m['period'], 'population': m['population'],
             'measure': m['measure'], 'value_type': m['value_type'],
             'source_ref': m.get('source_ref'), 'operation': m.get('operation'),
             'operands': m.get('operands', []), 'denominator': m.get('denominator'),
             'query': m.get('query'), 'count_basis': m.get('count_basis'),
             'missing_reason': m.get('missing_reason')}
            for m in metrics.values()]
