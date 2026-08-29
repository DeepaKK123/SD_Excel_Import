#!/usr/bin/env python3
"""
sd_file_fetch.py  <source_ref>  <import-archive-dir>

Resolves an incoming file reference to a local plain CSV/tab file that SD
BASIC (SR.LOAD.FILE / SR.DATA.PARSE) can then process. Called from
SR.FILE.FETCH via EXECUTE/PERFORM.

<source_ref> is one of:
  - a plain local path (email attachment already saved to the server)
  - 'DRIVE:<file-name>' - fetched from Google Drive via rclone first

Last line of stdout is always a status sentinel so the BASIC caller doesn't
have to rely on the OS exit code (unreliable through EXECUTE's '!' escape):
    OK:<local-file-path>      on success (single text/CSV source)
    OKLIST:<manifest-path>    on success (Excel workbook split to per-sheet CSVs)
  ERR:<message>          on failure

Prerequisites (fill in / install before the DRIVE: path will work):
  - rclone installed and configured with a remote named RCLONE_REMOTE below
    (one-time: `rclone config`, then set RCLONE_REMOTE to that remote name
    and RCLONE_DRIVE_PATH to the Drive folder the client drops files into).
  - `pip install openpyxl` if the client ever sends true binary .xlsx files
    that need converting to CSV. Plain .csv/.tsv/.txt sources need no extra
    package (stdlib csv module only).
"""
import csv
import os
import re
import shutil
import contextlib
import io
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

if '__file__' in globals():
    SCRIPT_DIR = Path(globals()['__file__']).resolve().parent
else:
    SCRIPT_DIR = Path('scripts').resolve()
BUNDLED_PACKAGE_DIR = SCRIPT_DIR / 'python_packages'
if BUNDLED_PACKAGE_DIR.is_dir():
    sys.path.insert(0, str(BUNDLED_PACKAGE_DIR))

RCLONE_REMOTE = 'gdrive'            # TODO: set to your configured rclone remote name
RCLONE_DRIVE_PATH = 'incoming'      # TODO: set to the Drive folder path the client uses
SHOW_INFO_LOGS = False


def info(msg):
    if SHOW_INFO_LOGS:
        print(msg, file=sys.stderr)


def fail(msg):
    print(msg, file=sys.stderr)
    print(f'ERR:{msg}')
    sys.exit(1)


def fetch_from_drive(file_name, dest_dir):
    if not shutil.which('rclone'):
        fail("rclone not installed - install/configure rclone, or save the file "
             "locally and pass its path instead of DRIVE:<name>.")
    remote_src = f'{RCLONE_REMOTE}:{RCLONE_DRIVE_PATH}/{file_name}'
    result = subprocess.run(
        ['rclone', 'copy', remote_src, dest_dir],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        fail(f"rclone copy failed for {remote_src}: {result.stderr.strip()}")
    downloaded = Path(dest_dir) / file_name
    if not downloaded.exists():
        fail(f"rclone reported success but {downloaded} not found")
    return downloaded


def ensure_python_dependencies():
    try:
        import openpyxl
        return
    except ImportError:
        fail('Bundled openpyxl package is missing from: ' + str(BUNDLED_PACKAGE_DIR))


def xlsx_to_csvs_with_manifest(xlsx_path, staging_dir):
    ensure_python_dependencies()
    import openpyxl

    def _clean_cell(value):
        if value is None:
            return ''
        text = str(value).strip()
        # BASIC parser is line-based; embedded newlines would split one row
        # into multiple physical lines and corrupt Record IDs.
        text = text.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
        return text

    def _norm_header_cell(text):
        return ''.join(ch for ch in str(text).strip().lower() if ch.isalnum())

    def _find_header_index(norm_headers, candidates):
        for idx, name in enumerate(norm_headers):
            if name in candidates:
                return idx
        return -1

    def _sheet_key_plan(sheet_name, header):
        norm_headers = [_norm_header_cell(col) for col in header]
        sheet_key = _norm_header_cell(sheet_name)

        sheet_mode = 'records'
        if 'split' in sheet_key:
            sheet_mode = 'splits'
        elif 'line' in sheet_key or 'detail' in sheet_key:
            sheet_mode = 'lines'

        po_idx = _find_header_index(
            norm_headers,
            {
                'ponumber', 'po', 'purchaseorder', 'purchaseorderno',
                'orderid', 'poid', 'poidentifier'
            },
        )
        line_idx = _find_header_index(
            norm_headers,
            {
                'linenumber', 'line', 'lineitem', 'lineitemnumber',
                'lineid', 'line#', 'polinenumber', 'lineitems',
                'itemnumber', 'item'
            },
        )
        split_idx = _find_header_index(
            norm_headers,
            {
                'split', 'splitnumber', 'splitid', 'distribution',
                'distributionnumber', 'distnumber'
            },
        )

        needs_line = sheet_mode in ('lines', 'splits')
        needs_split = sheet_mode == 'splits'
        return po_idx, line_idx, split_idx, needs_line, needs_split, sheet_mode

    def _build_rec_id(row, po_idx, line_idx, split_idx, needs_line, needs_split, sheet_mode):
        def _value_at(idx):
            if idx < 0 or idx >= len(row):
                return ''
            return row[idx].strip()

        first_col = _value_at(0)
        po_val = _value_at(po_idx)
        line_val = _value_at(line_idx)
        split_val = _value_at(split_idx)

        if po_val == '' and first_col.upper().startswith('PO-'):
            match = re.search(r'(PO-[A-Z0-9-]+)', first_col.upper())
            if match:
                po_val = match.group(1)
            if line_val == '':
                line_match = re.search(r'LINE\s*([0-9A-Z.-]+)', first_col.upper())
                if line_match:
                    line_val = line_match.group(1)
            if split_val == '':
                split_match = re.search(r'SPLIT\s*([0-9A-Z.-]+)', first_col.upper())
                if split_match:
                    split_val = split_match.group(1)

        # Prefer stable composite keys when the source exposes their parts.
        # Diffing indexes rows by ID, so omitting line/split values collapses
        # otherwise distinct rows during update imports.
        if sheet_mode == 'splits' and po_val != '' and line_val != '' and split_val != '':
            return f'{po_val}|{line_val}|{split_val}'
        if sheet_mode == 'lines' and po_val != '' and line_val != '':
            return f'{po_val}|{line_val}'

        # Preserve an explicit source key when the composite parts are absent.
        if first_col != '':
            return first_col

        if po_val != '':
            if sheet_mode == 'splits' and line_val != '':
                return f'{po_val}|{line_val}'
            return po_val

        return first_col

    def _extract_sheet_table(ws):
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return None

        scan_limit = min(len(rows), 50)
        header_index = -1
        best_score = -1

        for idx in range(scan_limit):
            row = rows[idx]
            width = sum(1 for value in row if value not in (None, ''))
            if width < 2:
                continue

            non_empty_next = 0
            for probe in rows[idx + 1:idx + 6]:
                probe_width = sum(1 for value in probe[:len(row)] if value not in (None, ''))
                if probe_width > 0:
                    non_empty_next += 1

            score = (non_empty_next * 1000) + width
            if score > best_score:
                best_score = score
                header_index = idx

        if header_index < 0:
            return None

        raw_header = rows[header_index]
        last_non_blank = -1
        for idx, value in enumerate(raw_header):
            if value not in (None, ''):
                last_non_blank = idx
        if last_non_blank < 1:
            return None

        header = [_clean_cell(value) for value in raw_header[:last_non_blank + 1]]

        data_rows = []
        blank_streak = 0
        for row in rows[header_index + 1:]:
            values = [_clean_cell(value) for value in row[:len(header)]]
            if all(value == '' for value in values):
                blank_streak += 1
                if blank_streak >= 5 and data_rows:
                    break
                continue

            blank_streak = 0
            data_rows.append(values)

        if not data_rows:
            return None
        return header, data_rows

    def _safe_sheet_slug(sheet_name):
        cleaned = ''.join(ch if ch.isalnum() else '_' for ch in str(sheet_name))
        cleaned = cleaned.strip('_')
        return cleaned or 'sheet'

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    included_sheets = []

    for ws in wb.worksheets:
        extracted = _extract_sheet_table(ws)
        if extracted is None:
            info(f"Skipping sheet '{ws.title}': no tabular data detected")
            continue

        header, data_rows = extracted
        header_key = tuple(_norm_header_cell(column) for column in header)
        if not any(header_key):
            info(f"Skipping sheet '{ws.title}': header row appears empty")
            continue

        included_sheets.append((ws.title, header, data_rows))

    wb.close()

    if not included_sheets:
        fail(f"No compatible table sheets found in workbook: {xlsx_path}")

    workbook_stem = Path(xlsx_path).stem
    manifest_path = Path(staging_dir) / f'{workbook_stem}__sheet_manifest.txt'
    manifest_lines = []
    total_rows = 0

    for sheet_index, (sheet_name, header, data_rows) in enumerate(included_sheets, start=1):
        total_rows += len(data_rows)
        sheet_slug = _safe_sheet_slug(sheet_name)
        csv_path = Path(staging_dir) / f'{workbook_stem}__sheet{sheet_index}_{sheet_slug}.csv'

        po_idx, line_idx, split_idx, needs_line, needs_split, sheet_mode = _sheet_key_plan(sheet_name, header)
        seen_rec_ids = set()
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['REC_ID'] + header)
            for row_no, row in enumerate(data_rows, start=1):
                rec_id = _build_rec_id(
                    row,
                    po_idx,
                    line_idx,
                    split_idx,
                    needs_line,
                    needs_split,
                    sheet_mode,
                )
                if rec_id == '':
                    continue
                if rec_id in seen_rec_ids:
                    fail(f"Duplicate record ID '{rec_id}' in sheet '{sheet_name}'")
                seen_rec_ids.add(rec_id)

                writer.writerow([rec_id] + row)
        info(f"Prepared sheet '{sheet_name}' -> {csv_path} ({len(data_rows)} rows)")
        manifest_lines.append(f"{sheet_name}|{csv_path}|{len(data_rows)}")

    with open(manifest_path, 'w', encoding='utf-8') as f:
        for line in manifest_lines:
            f.write(line + '\n')

    info(
        f"Workbook split complete: {len(included_sheets)} sheet CSV files, {total_rows} rows total"
    )
    return manifest_path


def main():
    if len(sys.argv) == 2 and sys.argv[1] == '--install-dependencies':
        ensure_python_dependencies()
        print('PYTHON_DEP_OK')
        return

    if len(sys.argv) == 4 and sys.argv[1] == '--parse':
        parse_csv(sys.argv[2], sys.argv[3])
        return

    if len(sys.argv) != 3:
        fail(f"Usage: {sys.argv[0]} <source_ref> <import-archive-dir>")

    source_ref, staging_dir = sys.argv[1], sys.argv[2]
    Path(staging_dir).mkdir(parents=True, exist_ok=True)

    if source_ref.startswith('DRIVE:'):
        file_name = source_ref[len('DRIVE:'):]
        with tempfile.TemporaryDirectory() as tmp:
            local_src = fetch_from_drive(file_name, tmp)
            local_src = _materialize(local_src, staging_dir)
    else:
        local_src = Path(source_ref)
        if not local_src.exists():
            fail(f"Source file not found: {local_src}")

    out_kind, out_path = materialize_text_or_xlsx(local_src, staging_dir)

    if out_kind == 'list':
        print(f'OKLIST:{out_path}')
    else:
        _read_csv_rows(out_path)
        print(f'OK:{out_path}')


def _to_sd_text(value):
    """Return text representable by SD's Latin-1 embedded Python bridge."""
    text = str(value)
    text = text.replace(chr(0x2013), '-').replace(chr(0x2014), '-')
    text = text.replace(chr(0x2018), "'").replace(chr(0x2019), "'")
    text = text.replace(chr(0x201C), '"').replace(chr(0x201D), '"')
    return text.encode('latin-1', 'replace').decode('latin-1')


def parse_csv(path, delimiter_name):
    delimiters = {'COMMA': ',', 'TAB': '\t', 'PIPE': '|'}
    delimiter = delimiters.get(delimiter_name.upper(), delimiter_name)
    try:
        with open(path, newline='', encoding='utf-8-sig') as source:
            rows = list(csv.reader(source, delimiter=delimiter))
    except Exception as exc:
        fail(f'CSV parse failed for {path}: {exc}')

    if not rows:
        fail(f'CSV file is empty: {path}')
    header = [_to_sd_text(cell) for cell in rows[0]]
    data_rows = [
        [_to_sd_text(cell) for cell in row]
        for row in rows[1:]
    ]
    data_rows = [_normalise_export_row(header, row) for row in data_rows]
    seen_ids = set()
    parsed_rows = []
    for row_number, row in enumerate(data_rows, start=2):
        if not row or row[0].strip() == '':
            continue
        record_id = row[0].strip()
        if record_id in seen_ids:
            fail(f"Duplicate record ID '{record_id}' in {path} at row {row_number}")
        seen_ids.add(record_id)
        parsed_rows.append(row)

    globals()['SD_HEADER'] = chr(253).join(header)
    globals()['SD_ROWS'] = chr(254).join(chr(253).join(row) for row in parsed_rows)
    globals()['SD_ROW_COUNT'] = str(len(parsed_rows))
    globals()['SD_RESULT'] = 'PARSE_OK'
    print('PARSE_OK')


def _materialize(downloaded_path, staging_dir):
    # Move the rclone temp-dir download into staging_dir before the temp dir is cleaned up.
    dest = Path(staging_dir) / downloaded_path.name
    shutil.move(str(downloaded_path), dest)
    return dest


def materialize_text_or_xlsx(local_src, staging_dir):
    if local_src.suffix.lower() in ('.xlsx', '.xls'):
        manifest_path = xlsx_to_csvs_with_manifest(local_src, staging_dir)
        return 'list', manifest_path
    else:
        # Keep delimited text as-is and normalize the output extension to CSV.
        out_path = Path(staging_dir) / (local_src.stem + '.csv')
        if local_src.resolve() != out_path.resolve():
            shutil.copy2(local_src, out_path)
    return 'single', out_path


def _read_csv_rows(path):
    with open(path, newline='', encoding='utf-8-sig') as f:
        rows = list(csv.reader(f))
    if not rows:
        return [], {}
    header, data_rows = rows[0], rows[1:]
    by_id = {}
    for row_number, row in enumerate(data_rows, start=2):
        if not row or row[0] == '':
            continue
        if row[0] in by_id:
            fail(f"Duplicate record ID '{row[0]}' in {path} at row {row_number}")
        by_id[row[0]] = row
    return header, by_id


def _normalise_header(header):
    return ''.join(character for character in header.lower() if character.isalnum())


def _normalise_export_value(header, value):
    text = str(value)
    header_key = _normalise_header(header)
    if header_key in ('lastcheckin', 'lasteassynctime'):
        if text.strip() == '0001-01-01 00:00:00.0000000':
            return ''
    if header_key in ('phonenumber', 'phone') and text.startswith('1'):
        return text[1:]
    return text


def _normalise_export_row(header, row):
    return [
        _normalise_export_value(header[index] if index < len(header) else '', value)
        for index, value in enumerate(row)
    ]


def _is_excel_text_header(header):
    return _normalise_header(header) in {
        'deviceid', 'imei', 'meid', 'serialnumber', 'phonenumber', 'phone'
    }


def _excel_text_value(value):
    return "'" + str(value) if value != '' else ''


def _excel_report_row(header, row):
    return [
        _excel_text_value(value) if _is_excel_text_header(header[index]) else value
        for index, value in enumerate(row)
    ]


def diff_csv(old_csv, new_csv, staging_dir, change_log_dir=None):
    """Compares yesterday's archived CSV against today's CSV, keyed by
    Column A (RecID), and writes a single dated report with a Status column
    (ADD/UPDATE/DELETE) - reviewable directly in Excel, and small enough that
    the BASIC caller only has to act on flagged rows instead of re-reading
    every record from the live MV file to find out what changed.
    """
    old_header, old_rows = _read_csv_rows(old_csv)
    new_header, new_rows = _read_csv_rows(new_csv)

    old_rows = {
        rec_id: _normalise_export_row(old_header, row)
        for rec_id, row in old_rows.items()
    }
    new_rows = {
        rec_id: _normalise_export_row(new_header, row)
        for rec_id, row in new_rows.items()
    }

    added_ids = new_rows.keys() - old_rows.keys()
    deleted_ids = old_rows.keys() - new_rows.keys()
    changed_ids = {
        rec_id for rec_id in (new_rows.keys() & old_rows.keys())
        if new_rows[rec_id] != old_rows[rec_id]
    }

    # Diagnostic only (stderr, not part of the DIFF_OK sentinel line): if a
    # huge fraction of rows are flagged changed, show which column indices
    # actually differ and how often - almost always a formatting/rounding
    # drift in one column (timestamps, booleans, etc.), not a real change.
    if changed_ids:
        header_for_diag = new_header or old_header
        col_diff_counts = [0] * len(header_for_diag)
        for rec_id in changed_ids:
            old_row, new_row = old_rows[rec_id], new_rows[rec_id]
            for idx in range(min(len(old_row), len(new_row), len(col_diff_counts))):
                if old_row[idx] != new_row[idx]:
                    col_diff_counts[idx] += 1
        ranked = sorted(enumerate(col_diff_counts), key=lambda x: -x[1])
        print(f'DIFF_DIAG: {len(changed_ids)} rows flagged changed; '
              f'columns differing most often:', file=sys.stderr)
        for idx, count in ranked[:5]:
            if count > 0:
                name = header_for_diag[idx] if idx < len(header_for_diag) else f'col{idx}'
                print(f'  {name!r}: differs in {count}/{len(changed_ids)} changed rows',
                      file=sys.stderr)

    header = new_header or old_header
    stem = Path(new_csv).stem
    # Include time, not just date - a second same-day import would otherwise
    # silently overwrite the earlier run's report with a date-only name.
    run_time = datetime.now()
    timestamp = run_time.strftime('%Y-%m-%d_%H%M%S')
    report_path = Path(staging_dir) / f'{stem}_diff_{timestamp}.csv'
    if change_log_dir is None:
        change_log_dir = staging_dir
    Path(change_log_dir).mkdir(parents=True, exist_ok=True)
    change_log_path = Path(change_log_dir) / f'{stem}_change_log_{timestamp}.csv'

    with open(report_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Status'] + header)
        for rec_id in new_rows:
            if rec_id in added_ids:
                writer.writerow(['ADD'] + _excel_report_row(header, new_rows[rec_id]))
            elif rec_id in changed_ids:
                writer.writerow(['UPDATE'] + _excel_report_row(header, new_rows[rec_id]))
        for rec_id in deleted_ids:
            writer.writerow(['DELETE'] + _excel_report_row(header, old_rows[rec_id]))

    # Keep detailed before/after values separate from the import report. The
    # BASIC loader expects the report's original column layout after Status.
    with open(change_log_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'TrackID', 'Date', 'Time', 'Action', 'RecordID',
            'FieldDescription', 'BeforeValue', 'AfterValue'
        ])
        track_id = _excel_text_value(run_time.strftime('%Y%m%d%H%M%S'))
        date_text = run_time.strftime('%Y-%m-%d')
        time_text = run_time.strftime('%H:%M:%S')

        for rec_id in sorted(added_ids):
            writer.writerow([
                track_id, date_text, time_text, 'ADD', _excel_text_value(rec_id),
                'Entire record', '', 'Added'
            ])

        for rec_id in sorted(changed_ids):
            old_row, new_row = old_rows[rec_id], new_rows[rec_id]
            max_columns = max(len(old_row), len(new_row), len(header))
            for index in range(max_columns):
                old_value = old_row[index] if index < len(old_row) else ''
                new_value = new_row[index] if index < len(new_row) else ''
                if old_value == new_value:
                    continue
                field_name = header[index] if index < len(header) else f'Column{index + 1}'
                if _is_excel_text_header(field_name):
                    old_value = _excel_text_value(old_value)
                    new_value = _excel_text_value(new_value)
                writer.writerow([
                    track_id, date_text, time_text, 'UPDATE', _excel_text_value(rec_id),
                    field_name, old_value, new_value
                ])

        for rec_id in sorted(deleted_ids):
            writer.writerow([
                track_id, date_text, time_text, 'DELETE', _excel_text_value(rec_id),
                'Entire record', 'Present', 'Deleted'
            ])

    # Report the new file's header/row count too, so the BASIC caller never
    # has to re-read the full source CSV just to get what it already has.
    print(f'DIFF_OK:{report_path}:'
          f'{len(added_ids)}:{len(changed_ids)}:{len(deleted_ids)}:'
            f'{len(new_rows)}:{"|".join(header)}:{change_log_path}')


def run_embedded():
    source_ref = globals().get('SD_SOURCE_REF')
    staging_dir = globals().get('SD_STAGING_DIR', '')
    if source_ref is None:
        return False

    original_argv = sys.argv
    captured = io.StringIO()
    if source_ref == '--diff':
        sys.argv = [
            'sd_file_fetch.py', '--diff', globals().get('SD_OLD_CSV', ''),
            globals().get('SD_NEW_CSV', ''), staging_dir
        ]
        if globals().get('SD_CHANGELOG_DIR'):
            sys.argv.append(globals()['SD_CHANGELOG_DIR'])
    elif source_ref == '--parse':
        sys.argv = [
            'sd_file_fetch.py', '--parse', globals().get('SD_PARSE_PATH', ''),
            globals().get('SD_PARSE_DELIM', 'COMMA')
        ]
    else:
        sys.argv = ['sd_file_fetch.py', source_ref]
    if source_ref not in ('--install-dependencies', '--diff', '--parse'):
        sys.argv.append(staging_dir)
    try:
        with contextlib.redirect_stdout(captured), contextlib.redirect_stderr(captured):
            main()
    except SystemExit:
        pass
    finally:
        sys.argv = original_argv

    lines = captured.getvalue().splitlines()
    for line in reversed(lines):
        if line.startswith(('OK:', 'OKLIST:', 'ERR:', 'PYTHON_DEP_OK', 'PARSE_OK')):
            globals()['SD_RESULT'] = line
            return True
    globals()['SD_RESULT'] = 'ERR:Python script produced no result'
    return True


if __name__ == '__main__':
    if run_embedded():
        pass
    elif len(sys.argv) in (5, 6) and sys.argv[1] == '--diff':
        Path(sys.argv[4]).mkdir(parents=True, exist_ok=True)
        change_log_dir = sys.argv[5] if len(sys.argv) == 6 else None
        diff_csv(sys.argv[2], sys.argv[3], sys.argv[4], change_log_dir)
    else:
        main()
