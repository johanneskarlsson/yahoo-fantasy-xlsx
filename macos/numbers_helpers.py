"""Helper functions for interacting with Apple Numbers via AppleScript.

These were originally private instance methods on MacOSDraftExporter but have been
extracted so they can be reused or tested in isolation. They remain thin wrappers
around osascript calls; higher-level code should handle data preparation.
"""

from __future__ import annotations

import os
import subprocess
import logging
from typing import List, Any, Iterable, Tuple


def create_sheets(
    filename: str,
    logger: logging.Logger,
    sheets: Iterable[Tuple[str, List[Any]]],
    force: bool = False,
    timeout: int = 30,
) -> None:
    """Create multiple sheets in a single Numbers open/save cycle.

    Parameters
    ----------
    filename : str
        Path to the .numbers file (created if first save in GUI previously).
    logger : logging.Logger
        Logger instance.
    sheets : iterable[(name, headers)]
        Sequence of (sheet_name, headers_list) definitions.
    force : bool, default False
        If True, will always (re)apply column count & headers even if sheet exists.
    timeout : int, default 30
        Seconds before AppleScript execution is aborted.

    Notes
    -----
    - This replaces repeated open/save/close cycles with a single operation.
    - If a sheet already exists and force is False, its headers are left untouched.
    - Header / sheet names have quotes escaped; other characters are passed through.
    """
    sheets = list(sheets)
    if not sheets:
        return

    numbers_abs = os.path.abspath(filename)

    snippet_list: List[str] = []
    for sheet_name, headers in sheets:
        safe_sheet = str(sheet_name).replace('"', '\\"')
        header_cmds = []
        for i, header in enumerate(headers, 1):
            escaped_header = str(header).replace('"', '\\"')
            header_cmds.append(f'set value of cell {i} of row 1 to "{escaped_header}"')
        headers_script = '\n                        '.join(header_cmds)

        # AppleScript snippet per sheet
        # targetSheet variable is reused per iteration safely (scoped inside tell doc)
        snippet = f'''
            -- Ensure sheet "{safe_sheet}"
            set sheetExists to false
            set targetSheet to missing value
            repeat with s in sheets
                if name of s is "{safe_sheet}" then
                    set sheetExists to true
                    set targetSheet to s
                    exit repeat
                end if
            end repeat
            if targetSheet is missing value then
                set targetSheet to make new sheet
                set name of targetSheet to "{safe_sheet}"
            end if
            tell targetSheet
                -- Ensure predictable table name
                set name of table 1 to "{safe_sheet}"
                tell table 1
                    if {str(force).lower()} or (not sheetExists) then
                        set column count to {len(headers)}
                        {headers_script}
                    end if
                end tell
            end tell
        '''
        snippet_list.append(snippet)

    all_snippets = '\n'.join(snippet_list)

    script = f'''
tell application "Numbers"
    try
        set targetPath to "{numbers_abs}"
        set doc to missing value
        -- Iterate open documents; some may be unsaved (file is missing value)
        repeat with d in documents
            set matched to false
            try
                set f to file of d
                if f is not missing value then
                    set p to POSIX path of (f as alias)
                    if p is targetPath then
                        set matched to true
                    end if
                end if
            end try
            if matched then
                set doc to d
                exit repeat
            end if
        end repeat
        if doc is missing value then
            -- Open (returns existing if already open with same file, otherwise opens)
            set doc to open (POSIX file targetPath)
        end if

        tell doc
{all_snippets}
        end tell
        save doc
        return "OK"
    on error errorMessage
        return "ERROR: " & errorMessage
    end try
end tell
'''

    try:
        res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        logger.error("Timeout creating sheets in bulk")
        return
    except Exception as e:  # pragma: no cover
        logger.error(f"Unexpected error creating sheets: {e}")
        return

    stdout = (res.stdout or "").strip()
    if stdout.startswith("ERROR:"):
        logger.error(f"Bulk sheet creation failed: {stdout}")
    elif res.returncode != 0:
        logger.error(f"Bulk sheet creation non-zero exit: {res.stderr}")
    else:
        logger.debug(f"Bulk sheet creation/ensure completed for {[name for name, _ in sheets]}")


def _write_sheet_chunk(
    sheet_name: str,
    data_rows,
    start_row: int,
    numbers_abs: str,
    logger: logging.Logger,
    timeout: int = 20,
) -> bool:
    """Internal helper to write a contiguous chunk of ``data_rows`` starting at ``start_row``.

    Improvements vs earlier version:
    - Returns bool success indicator instead of always None.
    - Expands table column count if incoming data has more columns than existing.
    - Escapes sheet name defensively.
    - Adds explicit "OK" / "ERROR" return strings from AppleScript for clearer diagnostics.
    - Parameterised timeout.
    - Slightly clearer separation of Python-side row serialization.
    """
    # Fast-path: nothing to write
    if not data_rows:
        return True

    # Calculate max column count we will need for this chunk
    try:
        max_cols = max(len(r) for r in data_rows if r is not None)
    except ValueError:  # all rows None/empty
        return True

    safe_sheet = sheet_name.replace('"', '\\"')

    # Encode rows for AppleScript list-of-lists syntax. Truncate each cell to 100 chars as before.
    script_rows: list[str] = []
    for row in data_rows:
        if row is None:
            row = []
        row_str: list[str] = []
        for cell in row:
            if cell is None or cell == "":
                row_str.append('""')
            else:
                cell_str = str(cell)[:100]
                # Basic escaping for quotes + normalise newlines -> space (AppleScript doesn't like raw newlines inside string literal)
                escaped = cell_str.replace('"', '\\"').replace('\n', ' ').replace('\r', ' ')
                row_str.append(f'"{escaped}"')
        script_rows.append('{' + ', '.join(row_str) + '}')

    rows_applescript = '{' + ', '.join(script_rows) + '}'

    # AppleScript: ensure sheet/table, expand columns if required, then write cell values row-by-row.
    script = f'''
tell application "Numbers"
    try
        set doc to open (POSIX file "{numbers_abs}")
        tell doc
            tell sheet "{safe_sheet}"
                tell table 1
                    if (column count) < {max_cols} then
                        set column count to {max_cols}
                    end if
                    set dataRows to {rows_applescript}
                    set rowIndex to {start_row}
                    repeat with rowData in dataRows
                        if rowIndex > (row count) then
                            add row below last row
                        end if
                        set colIndex to 1
                        repeat with cellValue in rowData
                            try
                                if column count >= colIndex then
                                    set value of cell colIndex of row rowIndex to cellValue
                                end if
                            end try
                            set colIndex to colIndex + 1
                        end repeat
                        set rowIndex to rowIndex + 1
                    end repeat
                end tell
            end tell
        end tell
        save doc
        close doc
        return "OK"
    on error errorMessage
        return "ERROR: " & errorMessage
    end try
end tell
'''

    try:
        res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        logger.error(f"Timeout updating {sheet_name} chunk (rows {start_row}-{start_row + len(data_rows) - 1})")
        return False
    except Exception as e:  # pragma: no cover - defensive
        logger.error(f"Error invoking osascript for {sheet_name} chunk: {e}")
        return False

    stdout = (res.stdout or "").strip()
    if stdout.startswith("ERROR:"):
        logger.error(f"AppleScript error updating {sheet_name} chunk: {stdout}")
        return False
    if res.returncode != 0:
        logger.error(f"osascript non-zero exit updating {sheet_name} chunk: rc={res.returncode} stderr={res.stderr.strip()}")
        return False

    logger.debug(
        f"Updated {sheet_name} chunk (rows {start_row}-{start_row + len(data_rows) - 1}, cols 1-{max_cols})"
    )
    return True


def update_sheet(filename: str, logger: logging.Logger, sheet: str, rows: List[List[Any]]) -> None:
    """Update the given sheet with the provided rows (after headers).

    Handles chunking to avoid overly large AppleScript payloads.
    Row 1 is assumed to contain headers already; data starts at row 2.
    """
    if not rows:
        return

    numbers_abs = os.path.abspath(filename)
    chunk_size = 100
    total_rows = len(rows)

    if total_rows > chunk_size:
        logger.debug(f"Processing {total_rows} rows in chunks of {chunk_size}")
        for i in range(0, total_rows, chunk_size):
            chunk = rows[i:i + chunk_size]
            chunk_start = i + 2  # +2: headers + 1-indexed
            _write_sheet_chunk(sheet, chunk, chunk_start, numbers_abs, logger)
    else:
        _write_sheet_chunk(sheet, rows, 2, numbers_abs, logger)


def apply_formulas(
    filename,
    logger,
    sheet: str,
    per_row: list = None,
    start_row: int = 2,
    end_row: int = None,
    static: list = None,
    table_index: int = 1,
    timeout_sec: int = 300,
):
    """
    Generic AppleScript-based formula applier.

    per_row: list of (column_letter, formula_template) where formula_template may contain "{row}" placeholder.
             Example: [("E", "=IF(ISERROR(INDEX('Teams'::D;MATCH(D{row};'Teams'::A;0)));\"\";INDEX('Teams'::D;MATCH(D{row};'Teams'::A;0)))")]
    static: list of (cell_ref, formula_string) for one-off formulas (e.g., [("A1", "=1+1")])
    If end_row is None it uses current table row count.
    All formulas must include the leading "=".
    """
    if not per_row and not static:
        return

    numbers_abs = os.path.abspath(filename)

    # Build AppleScript-safe lines
    static_cmds = []
    if static:
        for cell_ref, formula in static:
            if not formula.startswith("="):
                formula = "=" + formula
            esc = formula.replace('"', '\\"')
            static_cmds.append(f'set value of cell "{cell_ref}" to "{esc}"')

    per_row_cmds = []
    if per_row:
        for col, formula in per_row:
            if not formula.startswith("="):
                formula = "=" + formula
            # We'll substitute {row} at runtime inside AppleScript
            esc = formula.replace('"', '\\"')
            per_row_cmds.append(f'''
                        set fml to "{esc}"
                        set fml to my replace_text(fml, "{{row}}", r as text)
                        set value of cell ("{col}" & r) to fml''')

    per_row_block = ""
    if per_row_cmds:
        per_row_block = f'''
                        repeat with r from {start_row} to rowLimit
{''.join(per_row_cmds)}
                        end repeat'''

    static_block = ""
    if static_cmds:
        static_block = "\n                        " + "\n                        ".join(static_cmds)

    script = f'''tell application "Numbers"
    with timeout of {timeout_sec} seconds
        try
            set doc to open (POSIX file "{numbers_abs}")
            tell doc
                if (every sheet whose name is "{sheet}") = {{}} then return "ERROR: Missing sheet {sheet}"
                tell sheet "{sheet}"
                    tell table {table_index}
                        set rowCount to row count
                        set rowLimit to { 'rowCount' if end_row is None else end_row }
                        if rowLimit < {start_row} then
                            return "OK"
                        end if{static_block}{per_row_block}
                    end tell
                end tell
            end tell
            save doc
            close doc
            return "OK"
        on error errMsg
            return "ERROR: " & errMsg
        end try
    end timeout
end tell

on replace_text(t, find, repl)
    set {{otid, AppleScript's text item delimiters}} to {{AppleScript's text item delimiters, find}}
    set parts to text items of t
    set AppleScript's text item delimiters to repl
    set newText to parts as text
    set AppleScript's text item delimiters to otid
    return newText
end replace_text
'''

    try:
        res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=timeout_sec + 30)
        out = (res.stdout or "").strip()
        if out.startswith("ERROR:"):
            logger.error(f"apply_formulas AppleScript error ({sheet}): {out}")
        elif res.returncode != 0:
            logger.error(f"apply_formulas non-zero exit ({sheet}) rc={res.returncode} stderr={res.stderr.strip()}")
        else:
            logger.debug(f"Formulas applied to {sheet} (per_row={bool(per_row)} static={bool(static)})")
    except subprocess.TimeoutExpired:
        logger.error(f"apply_formulas timeout on sheet {sheet}")
    except Exception as e:  # pragma: no cover
        logger.error(f"apply_formulas unexpected error on {sheet}: {e}")


__all__ = ["create_sheets", "update_sheet", "append_rows", "apply_formulas"]
