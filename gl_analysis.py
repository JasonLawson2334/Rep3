import argparse
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from pathlib import Path
import re

DATE_NUMFMT_IDS = {14, 15, 16, 17, 22, 27, 30, 36, 45, 46, 47, 50, 57}
NUMERIC_COLUMN_NAMES = {'debit', 'credit', 'amount', 'absoluteamount'}


def col_to_index(col):
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c.upper()) - 64)
    return idx - 1


def load_shared_strings(z):
    shared_strings = []
    if 'xl/sharedStrings.xml' not in z.namelist():
        return shared_strings
    sst = ET.fromstring(z.read('xl/sharedStrings.xml'))
    ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    for si in sst.findall('ns:si', ns):
        text_parts = []
        for t in si.findall('.//ns:t', ns):
            text_parts.append(t.text or '')
        shared_strings.append(''.join(text_parts))
    return shared_strings


def load_styles(z):
    numfmt_map = {}
    style_numfmts = []
    if 'xl/styles.xml' not in z.namelist():
        return numfmt_map, style_numfmts
    styles = ET.fromstring(z.read('xl/styles.xml'))
    ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    numFmts = styles.find('ns:numFmts', ns)
    if numFmts is not None:
        for numFmt in numFmts.findall('ns:numFmt', ns):
            numfmt_map[int(numFmt.get('numFmtId'))] = numFmt.get('formatCode')
    cellXfs = styles.find('ns:cellXfs', ns)
    if cellXfs is not None:
        for xf in cellXfs.findall('ns:xf', ns):
            style_numfmts.append(int(xf.get('numFmtId')))
    return numfmt_map, style_numfmts


def is_date_format(numfmt_id, numfmt_code):
    if numfmt_id in DATE_NUMFMT_IDS:
        return True
    if not numfmt_code:
        return False
    fmt = numfmt_code.lower()
    return any(token in fmt for token in ['yy', 'mm', 'dd', 'hh', 'ss'])


def excel_to_datetime(serial):
    base = datetime(1899, 12, 30)
    return base + timedelta(days=serial)


def load_sheet(z, sheet_path):
    shared_strings = load_shared_strings(z)
    numfmt_map, style_numfmts = load_styles(z)
    sheet = ET.fromstring(z.read(sheet_path))
    sheet_ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    rows = []
    for row in sheet.findall('ns:sheetData/ns:row', sheet_ns):
        cells = {}
        for c in row.findall('ns:c', sheet_ns):
            r = c.get('r')
            match = re.match(r'([A-Z]+)', r)
            if not match:
                continue
            col = col_to_index(match.group(1))
            t = c.get('t')
            s = c.get('s')
            v = c.find('ns:v', sheet_ns)
            value = None
            if v is not None:
                v_text = v.text
                if t == 's':
                    value = shared_strings[int(v_text)]
                elif t == 'b':
                    value = v_text == '1'
                else:
                    try:
                        num = float(v_text)
                    except (TypeError, ValueError):
                        value = v_text
                    else:
                        is_date = False
                        if s is not None:
                            style_index = int(s)
                            fmt_id = style_numfmts[style_index] if style_index < len(style_numfmts) else None
                            fmt_code = numfmt_map.get(fmt_id)
                            if fmt_id is not None and is_date_format(fmt_id, fmt_code):
                                is_date = True
                        value = excel_to_datetime(num) if is_date else num
            cells[col] = value
        rows.append(cells)

    max_col = max((max(r.keys()) for r in rows if r), default=-1)
    table = []
    for row in rows:
        table.append([row.get(i) for i in range(max_col + 1)])
    return table


def select_sheet(z, sheet_name=None):
    workbook = ET.fromstring(z.read('xl/workbook.xml'))
    ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sheets = workbook.find('ns:sheets', ns)
    sheet_info = []
    for sheet in sheets.findall('ns:sheet', ns):
        sheet_info.append((sheet.get('name'), sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')))

    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rel_ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
    rel_map = {rel.get('Id'): rel.get('Target') for rel in rels.findall('r:Relationship', rel_ns)}

    if sheet_name:
        for name, rel_id in sheet_info:
            if name == sheet_name:
                return name, 'xl/' + rel_map[rel_id]
        raise ValueError(f'Sheet {sheet_name} not found. Options: {[name for name, _ in sheet_info]}')
    name, rel_id = sheet_info[0]
    return name, 'xl/' + rel_map[rel_id]


def normalize_rows(headers, rows):
    numeric_indices = {i for i, header in enumerate(headers) if str(header).strip().lower() in NUMERIC_COLUMN_NAMES}
    normalized = []
    for row in rows:
        new_row = list(row)
        for idx in numeric_indices:
            if idx >= len(new_row):
                continue
            value = new_row[idx]
            if isinstance(value, str):
                try:
                    new_row[idx] = float(value)
                except ValueError:
                    pass
        normalized.append(new_row)
    return normalized


def detect_column_types(rows, headers):
    types = {}
    for idx, header in enumerate(headers):
        header_key = str(header).strip().lower()
        values = [row[idx] for row in rows if idx < len(row)]
        non_null = [v for v in values if v not in (None, '')]
        is_date = any(isinstance(v, datetime) for v in non_null)
        if header_key in NUMERIC_COLUMN_NAMES:
            types[header] = 'number'
            continue
        is_number = all(isinstance(v, (int, float)) for v in non_null) if non_null else False
        if is_date:
            types[header] = 'date'
        elif is_number:
            types[header] = 'number'
        else:
            types[header] = 'text'
    return types


def svg_bar_chart(title, x_labels, values, width=960, height=540):
    margin = 80
    chart_width = width - margin * 2
    chart_height = height - margin * 2
    max_val = max(values) if values else 1
    bar_width = chart_width / max(len(values), 1)

    lines = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}">',
        f'<rect width="100%" height="100%" fill="#ffffff"/>',
        f'<text x="{width/2}" y="{margin/2}" text-anchor="middle" font-size="20" font-family="Arial">{title}</text>',
        f'<line x1="{margin}" y1="{height - margin}" x2="{width - margin}" y2="{height - margin}" stroke="#333"/>',
        f'<line x1="{margin}" y1="{margin}" x2="{margin}" y2="{height - margin}" stroke="#333"/>'
    ]
    for i, (label, value) in enumerate(zip(x_labels, values)):
        x = margin + i * bar_width + bar_width * 0.1
        bar_h = chart_height * (value / max_val) if max_val else 0
        y = height - margin - bar_h
        lines.append(f'<rect x="{x:.2f}" y="{y:.2f}" width="{bar_width * 0.8:.2f}" height="{bar_h:.2f}" fill="#4a90e2"/>')
        lines.append(f'<text x="{x + bar_width * 0.4:.2f}" y="{height - margin + 18}" text-anchor="middle" font-size="10" font-family="Arial" transform="rotate(45 {x + bar_width * 0.4:.2f},{height - margin + 18})">{label}</text>')
    lines.append('</svg>')
    return '\n'.join(lines)


def select_amount_column(headers, rows):
    for candidate in ['AbsoluteAmount', 'Amount', 'Debit', 'Credit']:
        if candidate in headers:
            idx = headers.index(candidate)
            has_numeric = any(isinstance(row[idx], (int, float)) for row in rows if idx < len(row))
            if has_numeric:
                return candidate
    return None


def main():
    parser = argparse.ArgumentParser(description='Basic GL Excel analysis without external dependencies.')
    parser.add_argument('xlsx_path', type=Path)
    parser.add_argument('--sheet', default=None)
    parser.add_argument('--output-dir', type=Path, default=Path('analysis_outputs'))
    args = parser.parse_args()

    output_dir = args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(args.xlsx_path) as z:
        sheet_name, sheet_path = select_sheet(z, args.sheet)
        table = load_sheet(z, sheet_path)

    headers = table[0]
    raw_rows = table[1:]
    rows = normalize_rows(headers, raw_rows)
    row_count = len(rows)

    column_types = detect_column_types(rows, headers)

    column_stats = []
    for idx, header in enumerate(headers):
        values = [row[idx] for row in rows if idx < len(row)]
        non_null = [v for v in values if v not in (None, '')]
        unique_vals = set(non_null)
        sample_vals = list(unique_vals)[:5]
        stats = {
            'column': header,
            'type': column_types[header],
            'non_null': len(non_null),
            'unique': len(unique_vals),
            'sample': sample_vals,
        }
        if column_types[header] == 'number':
            numeric_vals = [v for v in non_null if isinstance(v, (int, float))]
            if numeric_vals:
                stats['min'] = min(numeric_vals)
                stats['max'] = max(numeric_vals)
                stats['sum'] = sum(numeric_vals)
        if column_types[header] == 'date' and non_null:
            stats['min'] = min(non_null)
            stats['max'] = max(non_null)
        column_stats.append(stats)

    date_columns = [col for col, typ in column_types.items() if typ == 'date']
    date_ranges = {}
    for col in date_columns:
        idx = headers.index(col)
        values = [row[idx] for row in rows if idx < len(row) and isinstance(row[idx], datetime)]
        if values:
            date_ranges[col] = (min(values), max(values))

    primary_date_col = None
    for candidate in headers:
        if 'date' in str(candidate).lower() and column_types.get(candidate) == 'date':
            primary_date_col = candidate
            break
    if primary_date_col is None and date_columns:
        primary_date_col = date_columns[0]

    monthly_counts = []
    if primary_date_col:
        idx = headers.index(primary_date_col)
        counter = Counter()
        for row in rows:
            value = row[idx] if idx < len(row) else None
            if isinstance(value, datetime):
                counter[value.strftime('%Y-%m')] += 1
        monthly_counts = sorted(counter.items())

    account_column = None
    for candidate in ['GLAccountNumber', 'GLAccountName', 'Account', 'AccountNumber']:
        if candidate in headers:
            account_column = candidate
            break
    amount_column = select_amount_column(headers, rows)

    top_accounts = []
    if account_column and amount_column:
        account_idx = headers.index(account_column)
        amount_idx = headers.index(amount_column)
        totals = defaultdict(float)
        for row in rows:
            account = row[account_idx] if account_idx < len(row) else None
            amount = row[amount_idx] if amount_idx < len(row) else None
            if account not in (None, '') and isinstance(amount, (int, float)):
                totals[str(account).strip()] += abs(amount)
        top_accounts = sorted(totals.items(), key=lambda x: x[1], reverse=True)[:10]

    summary_path = output_dir / 'je_analysis.txt'
    with summary_path.open('w', encoding='utf-8') as f:
        f.write(f'Workbook: {args.xlsx_path.name}\n')
        f.write(f'Sheet analyzed: {sheet_name}\n')
        f.write(f'Row count (excluding header): {row_count}\n')
        f.write(f'Column count: {len(headers)}\n\n')

        f.write('Date ranges:\n')
        if date_ranges:
            for col, (min_date, max_date) in date_ranges.items():
                f.write(f'  - {col}: {min_date.date()} to {max_date.date()}\n')
        else:
            f.write('  (No date columns detected)\n')

        f.write('\nColumn profiles:\n')
        for stats in column_stats:
            f.write(f"- {stats['column']} ({stats['type']}): non-null={stats['non_null']}, unique={stats['unique']}\n")
            if 'min' in stats and 'max' in stats:
                f.write(f"  min={stats['min']}, max={stats['max']}\n")
            if 'sum' in stats:
                f.write(f"  sum={stats['sum']}\n")
            f.write(f"  sample={stats['sample']}\n")

        if primary_date_col:
            f.write(f"\nMonthly counts based on {primary_date_col}: {len(monthly_counts)} months\n")

        if top_accounts:
            f.write(f"Top accounts by {amount_column} (absolute):\n")
            for account, total in top_accounts:
                f.write(f"  {account}: {total:.2f}\n")

    if monthly_counts:
        labels = [label for label, _ in monthly_counts]
        values = [value for _, value in monthly_counts]
        chart_svg = svg_bar_chart(f'Entries by Month ({primary_date_col})', labels, values)
        (output_dir / 'entries_by_month.svg').write_text(chart_svg, encoding='utf-8')

    if top_accounts:
        labels = [label for label, _ in top_accounts]
        values = [value for _, value in top_accounts]
        chart_svg = svg_bar_chart(f'Top Accounts by {amount_column} (Absolute)', labels, values)
        (output_dir / 'top_accounts_by_amount.svg').write_text(chart_svg, encoding='utf-8')


if __name__ == '__main__':
    main()
