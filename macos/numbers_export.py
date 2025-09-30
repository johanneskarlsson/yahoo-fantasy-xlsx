import os, logging, subprocess, csv, tempfile
from typing import List, Dict, Any
from datetime import datetime


class MacOSDraftExporter:
    """Mac exporter using pure AppleScript for Numbers - no XLSX intermediate files."""

    BASE_SHEETS = {
        "Draft Board": ["draftedBy", "playerName", "team", "position", "averagePick", "projectedPoints"],
        "League Settings": ["setting", "value"],
        "Teams": ["teamKey", "teamId", "teamName", "manager"],
        "Draft Results": ["round", "pick", "playerName", "teamId", "manager"],
        "Pre-Draft Analysis": [
            "playerKey", "playerName", "team", "position", "averagePick", "averageRound", "percentDrafted",
            "projectedAuctionValue", "averageAuctionCost", "seasonRank", "positionRank", "preseasonAveragePick",
            "preseasonPercentDrafted"
        ]
    }

    def __init__(self, filename: str = "fantasy_draft_data.numbers"):
        # Override parent to use .numbers instead of .xlsx
        if not filename.lower().endswith('.numbers'):
            filename = filename.replace('.xlsx', '.numbers') if filename.endswith('.xlsx') else filename + '.numbers'

        self.filename = filename
        self.logger = logging.getLogger(__name__)

        # Don't create the file here - it will be created by the first data import
        # The CSV import creates the document from scratch anyway

    def _add_missing_sheets(self, missing_sheets):
        """Add missing sheets to the Numbers file."""
        numbers_abs = os.path.abspath(self.filename)

        script_parts = [f'''
on run
    set documentsPath to POSIX file "{numbers_abs}"

    tell application "Numbers"
        set doc to open documentsPath

        tell doc''']

        for sheet_name in missing_sheets:
            headers = self.BASE_SHEETS[sheet_name]
            script_parts.append(f'''
            -- Add {sheet_name} sheet
            set newSheet to make new sheet
            set name of newSheet to "{sheet_name}"

            tell newSheet
                tell table 1
                    {self._generate_header_script(headers)}
                end tell
            end tell''')

        script_parts.extend(['''
        end tell

        save doc
        close doc

    end tell
    return "OK"
end run'''])

        script = '\n'.join(script_parts)

        try:
            res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
            if res.returncode == 0:
                self.logger.debug(f"Added missing sheets: {missing_sheets}")
            else:
                self.logger.error(f"Failed to add missing sheets: {res.stderr}")
        except Exception as e:
            self.logger.error(f"Error adding missing sheets: {e}")

    def append_picks(self, rows: List[List[str]]):
        """Append new draft picks to Draft Results sheet using pure AppleScript."""
        if not rows:
            return

        numbers_abs = os.path.abspath(self.filename)

        # Convert rows to AppleScript format
        rows_script = self._rows_to_applescript(rows)

        script = f'''
on run
    set documentsPath to POSIX file "{numbers_abs}"
    set newRows to {rows_script}

    tell application "Numbers"
        set doc to open documentsPath

        tell doc
            tell sheet "Draft Results"
                tell table 1
                    set currentRows to row count
                    set startRow to currentRows + 1

                    -- Add rows for new data
                    repeat (count of newRows) times
                        add row below last row
                    end repeat

                    -- Fill in the data
                    set rowIndex to startRow
                    repeat with rowData in newRows
                        set colIndex to 1
                        repeat with cellValue in rowData
                            set value of cell colIndex of row rowIndex to cellValue
                            set colIndex to colIndex + 1
                        end repeat
                        set rowIndex to rowIndex + 1
                    end repeat
                end tell
            end tell
        end tell

        save doc
        close doc

    end tell
    return "OK"
end run
'''

        try:
            res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
            if res.returncode == 0:
                self.logger.debug(f"Added {len(rows)} picks to Draft Results")
            else:
                self.logger.error(f"Failed to append picks: {res.stderr}")
        except Exception as e:
            self.logger.error(f"Error appending picks: {e}")

    def _rows_to_applescript(self, rows: List[List[str]]) -> str:
        """Convert Python rows to AppleScript list format."""
        applescript_rows = []

        for row in rows:
            row_items = []
            for cell in row:
                if cell is None:
                    row_items.append('""')
                elif isinstance(cell, (int, float)):
                    row_items.append(str(cell))
                else:
                    # Escape quotes and wrap in quotes
                    escaped = str(cell).replace('"', '\\"')
                    row_items.append(f'"{escaped}"')
            applescript_rows.append('{' + ', '.join(row_items) + '}')

        return '{' + ', '.join(applescript_rows) + '}'

    def timestamp(self):
        """Add timestamp to Draft Results sheet."""
        numbers_abs = os.path.abspath(self.filename)
        timestamp_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        script = f'''
tell application "Numbers"
    try
        set doc to open (POSIX file "{numbers_abs}")
        tell doc
            tell sheet "Draft Results"
                tell table 1
                    -- Ensure we have enough columns (at least 9 for column I)
                    if column count < 9 then
                        set column count to 9
                    end if
                    set value of cell "I1" to "Last updated: {timestamp_str}"
                end tell
            end tell
        end tell
        save doc
        close doc
    on error errorMessage
        return "ERROR: " & errorMessage
    end try
end tell
'''

        try:
            res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=10)
            if res.returncode == 0:
                self.logger.debug("Added timestamp to Draft Results")
            else:
                self.logger.error(f"Failed to add timestamp: {res.stderr}")
        except subprocess.TimeoutExpired:
            self.logger.error("Timeout adding timestamp")
        except Exception as e:
            self.logger.error(f"Error adding timestamp: {e}")

    def update_league_settings_data(self, league_settings: Dict[str, Any]):
        """Populate League Settings sheet with grouped sections using AppleScript."""
        # Build the data rows
        rows = []
        rows.append(["League Name", league_settings.get('league_name', '')])
        rows.append(["League Type", league_settings.get('league_type', '')])
        rows.append(["Scoring Type", league_settings.get('scoring_type', '')])
        rows.append(["Max Teams", league_settings.get('max_teams', '')])
        rows.append(["Playoff Teams", league_settings.get('num_playoff_teams', '')])
        rows.append(["Playoff Start Week", league_settings.get('playoff_start_week', '')])
        rows.append(["", ""])  # spacer
        rows.append(["ROSTER POSITIONS", "COUNT"])
        for pos in league_settings.get('roster_positions', []):
            rows.append([pos.get('position', ''), pos.get('count', '')])
        rows.append(["", ""])  # spacer
        rows.append(["SKATER STATS", "VALUE"])
        for stat in league_settings.get('stat_categories', []):
            if stat.get('position_type') == 'P':
                name = stat.get('display_name') or stat.get('name') or ''
                rows.append([name, stat.get('value', '')])
        rows.append(["", ""])  # spacer
        rows.append(["GOALIE STATS", "VALUE"])
        for stat in league_settings.get('stat_categories', []):
            if stat.get('position_type') == 'G':
                name = stat.get('display_name') or stat.get('name') or ''
                rows.append([name, stat.get('value', '')])

        # Use the simplified update method
        self._update_sheet_data_simple("League Settings", rows)

    def update_teams_data(self, teams_rows):
        """Write Teams sheet data using AppleScript."""
        if not teams_rows:
            return

        self._update_sheet_data_simple("Teams", teams_rows)

    def update_draft_analysis_data(self, players_rows):
        """High-performance write for Pre-Draft Analysis using CSV import + sheet duplication.

        Strategy:
        1. Write a temporary CSV with headers + rows.
        2. Open CSV in Numbers (creates a new document quickly).
        3. Duplicate its first sheet into the main Numbers document, replacing existing sheet.
        4. Close temp document without saving; delete temp file.

        This leverages Numbers' native CSV ingestion which is far faster than AppleScript per-cell writes.
        """
        if not players_rows:
            return

        headers = self.BASE_SHEETS.get("Pre-Draft Analysis", [])
        try:
            self._import_sheet_via_csv("Pre-Draft Analysis", headers, players_rows)
        except Exception as e:
            self.logger.error(f"CSV import path failed for Pre-Draft Analysis: {e}; falling back to bulk script")
            # Fallback to older bulk method if CSV approach fails
            try:
                self._bulk_update_sheet("Pre-Draft Analysis", players_rows)
            except Exception as e2:
                self.logger.error(f"Secondary bulk fallback failed: {e2}")

    def _import_sheet_via_csv(self, sheet_name: str, headers, rows):
        """Replace entire Numbers document with CSV import as base, then restore missing base sheets."""
        numbers_abs = os.path.abspath(self.filename)
        # Create temp CSV
        fd, temp_path = tempfile.mkstemp(suffix='.csv', prefix='yf_tmp_')
        os.close(fd)
        try:
            with open(temp_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                for r in rows:
                    if r is None:
                        writer.writerow([''] * len(headers))
                        continue
                    adj = list(r)[:len(headers)]
                    if len(adj) < len(headers):
                        adj.extend([''] * (len(headers) - len(adj)))
                    # Convert decimal separators from dots to commas for Numbers (Swedish locale)
                    converted_row = []
                    for val in adj:
                        if isinstance(val, (int, float)):
                            # Convert number to string with comma separator
                            converted_row.append(str(val).replace('.', ','))
                        elif isinstance(val, str) and val:
                            # Try to detect if it's a numeric string and convert
                            try:
                                float(val)  # Test if it's numeric
                                converted_row.append(val.replace('.', ','))
                            except ValueError:
                                converted_row.append(val)
                        else:
                            converted_row.append(val)
                    writer.writerow(converted_row)
        except Exception as e:
            try:
                os.remove(temp_path)
            except Exception:
                pass
            raise RuntimeError(f"Failed writing temp CSV: {e}")

        # Build script to recreate other base sheets if missing
        recreate_snippets = []
        for sname, sheaders in self.BASE_SHEETS.items():
            if sname == sheet_name:
                continue
            header_cmds = []
            for i, h in enumerate(sheaders, 1):
                safe_h = str(h).replace('"', '\\"')
                header_cmds.append(f'set value of cell {i} of row 1 to "{safe_h}"')
            header_block = '\n          '.join(header_cmds)
            recreate_snippets.append(f'''
      if (every sheet whose name is "{sname}") = {{}} then
        set newSheet to make new sheet
        set name of newSheet to "{sname}"
        tell newSheet
          set name of table 1 to "{sname} Table"
          tell table 1
            set column count to {len(sheaders)}
            {header_block}
          end tell
        end tell
      end if''')
        recreate_script = '\n'.join(recreate_snippets)

        script = f'''
tell application "Numbers"
    with timeout of 600 seconds
        set csvDoc to open (POSIX file "{temp_path}")
        delay 0.5
        -- csvDoc now contains sheet 1 with imported data; rename sheet and table
        tell csvDoc
            try
                set name of sheet 1 to "{sheet_name}"
                -- Rename table to a predictable name
                tell sheet 1
                    set name of table 1 to "{sheet_name} Table"
                end tell
            end try
        end tell
        -- Save CSV Numbers doc as target path (overwrite)
        try
            save csvDoc in POSIX file "{numbers_abs}"
        on error errMsg number errNum
            -- If already exists, attempt closing and retry
            try
                close csvDoc saving yes
            end try
        end try
        close csvDoc saving yes
        -- Re-open the newly saved main document
        set mainDoc to open (POSIX file "{numbers_abs}")
        delay 0.3
        tell mainDoc
            {recreate_script}
        end tell
        save mainDoc
        close mainDoc
    end timeout
end tell
'''
        try:
            res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=120)
            if res.returncode != 0:
                raise RuntimeError(f"AppleScript CSV import failed: {res.stderr.strip() or res.stdout.strip()}")
            self.logger.debug(f"Replaced sheet {sheet_name} via CSV import ({len(rows)} rows)")
        finally:
            try:
                os.remove(temp_path)
            except Exception:
                pass

    def _bulk_update_sheet(self, sheet_name: str, data_rows):
        """High-performance update using range assignment to entire rows in one script.

        Strategy:
        1. Open Numbers document once.
        2. Ensure sheet exists (create if missing).
        3. Ensure column count >= widest row.
        4. Ensure row count >= header + data rows.
        5. Set each row's cells in a single statement: set value of cells 1 thru N of row R to { ... }.
        6. Blank any trailing old rows beyond new data (optional cleanup for consistency).
        """
        if not data_rows:
            return

        numbers_abs = os.path.abspath(self.filename)

        # Determine widest row to size columns appropriately
        max_cols = 0
        prepped_rows = []
        for r in data_rows:
            if r is None:
                continue
            row_vals = []
            for c in r:
                if c is None:
                    row_vals.append('""')
                elif isinstance(c, (int, float)):
                    row_vals.append(str(c))
                else:
                    # Truncate very long text to avoid enormous AppleScript payloads
                    s = str(c)[:200].replace('"', '\\"').replace('\n', ' ').replace('\r', ' ')
                    row_vals.append(f'"{s}"')
            if len(row_vals) > max_cols:
                max_cols = len(row_vals)
            prepped_rows.append('{' + ', '.join(row_vals) + '}')

        total_rows = len(prepped_rows)

        # Build AppleScript list for iterative assignment inside AppleScript (reduces script size)
        rows_applescript = '{' + ', '.join(prepped_rows) + '}'

        # AppleScript loop will iterate rows & columns; faster than generating thousands of lines
        assignment_loop = f'''set dataRows to {rows_applescript}
                    set startRow to 2
                    set totalDataRows to count of dataRows
                    repeat with rIndex from 1 to totalDataRows
                        set rowData to item rIndex of dataRows
                        set targetRow to startRow + rIndex - 1
                        set colCount to count of rowData
                        if colCount > 0 then
                            repeat with c from 1 to colCount
                                try
                                    set value of cell c of row targetRow to item c of rowData
                                end try
                            end repeat
                        end if
                    end repeat'''

        script = f'''
tell application "Numbers"
    with timeout of 600 seconds
        set doc to open (POSIX file "{numbers_abs}")
        tell doc
            -- Ensure target sheet exists
            set sheetExists to false
            repeat with s in sheets
                if name of s is "{sheet_name}" then
                    set sheetExists to true
                    exit repeat
                end if
            end repeat
            if not sheetExists then
                set newSheet to make new sheet
                set name of newSheet to "{sheet_name}"
            end if

            tell sheet "{sheet_name}"
                tell table 1
                    -- Ensure sufficient columns
                    if (column count) < {max_cols} then set column count to {max_cols}

                    -- Ensure sufficient rows (header + data)
                    set neededRows to (1 + {total_rows})
                    if (row count) < neededRows then
                        repeat (neededRows - (row count)) times
                            add row below last row
                        end repeat
                    end if

                    -- Bulk loop assignments
                    {assignment_loop}
                end tell
            end tell
        end tell
        save doc
        close doc
    end timeout
end tell
'''

        # Execute script (single large pass)
        try:
            res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=120)
            if res.returncode == 0:
                self.logger.debug(f"Bulk-updated {sheet_name} with {total_rows} rows (max {max_cols} cols)")
            else:
                self.logger.error(f"Bulk update script failed rc={res.returncode}: {res.stderr.strip()}")
        except subprocess.TimeoutExpired:
            raise RuntimeError(f"Bulk update timeout for {sheet_name}")
        except Exception:
            raise

    def _update_sheet_data_simple(self, sheet_name: str, data_rows):
        """Simplified method to update sheet data using AppleScript - optimized for large datasets."""
        if not data_rows:
            return

        numbers_abs = os.path.abspath(self.filename)

        # For large datasets, break into smaller chunks
        chunk_size = 100  # Process 100 rows at a time
        total_rows = len(data_rows)

        if total_rows > chunk_size:
            self.logger.debug(f"Processing {total_rows} rows in chunks of {chunk_size}")

            # Process in chunks
            for i in range(0, total_rows, chunk_size):
                chunk = data_rows[i:i + chunk_size]
                chunk_start = i + 2  # +2 because row 1 is headers, and we're 1-indexed
                self._update_sheet_chunk(sheet_name, chunk, chunk_start, numbers_abs)
        else:
            # Process all at once for small datasets
            self._update_sheet_chunk(sheet_name, data_rows, 2, numbers_abs)

    def _update_sheet_chunk(self, sheet_name: str, data_rows, start_row: int, numbers_abs: str):
        """Update a chunk of data in a sheet."""
        # Convert rows to a simpler format for AppleScript
        script_rows = []
        for row in data_rows:
            row_str = []
            for cell in row:
                if cell is None or cell == "":
                    row_str.append('""')
                else:
                    # Escape quotes and wrap in quotes - limit length to avoid AppleScript issues
                    cell_str = str(cell)[:100]  # Limit cell content length
                    escaped = cell_str.replace('"', '\\"').replace('\n', ' ').replace('\r', ' ')
                    row_str.append(f'"{escaped}"')
            script_rows.append('{' + ', '.join(row_str) + '}')

        rows_applescript = '{' + ', '.join(script_rows) + '}'

        script = f'''
tell application "Numbers"
    try
        set doc to open (POSIX file "{numbers_abs}")
        tell doc
            tell sheet "{sheet_name}"
                tell table 1
                    set dataRows to {rows_applescript}
                    set rowIndex to {start_row}

                    repeat with rowData in dataRows
                        -- Ensure we have enough rows
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
    on error errorMessage
        return "ERROR: " & errorMessage
    end try
end tell
'''

        try:
            res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=20)
            if res.returncode == 0:
                self.logger.debug(f"Updated {sheet_name} chunk (rows {start_row}-{start_row + len(data_rows) - 1})")
            else:
                self.logger.error(f"Failed to update {sheet_name} chunk: {res.stderr}")
        except subprocess.TimeoutExpired:
            self.logger.error(f"Timeout updating {sheet_name} chunk")
        except Exception as e:
            self.logger.error(f"Error updating {sheet_name} chunk: {e}")

    def setup_projection_sheets(self, league_settings):
        """Create Skater/Goalie Projections sheets with TOTAL formulas."""
        try:
            # Determine stat names by position_type
            skater_stats, goalie_stats = [], []
            for stat in league_settings.get('stat_categories', []):
                name = stat.get('display_name') or stat.get('name') or ''
                if name:  # Only add if name exists
                    ptype = stat.get('position_type')
                    if ptype == 'P':
                        skater_stats.append(name)
                    elif ptype == 'G':
                        goalie_stats.append(name)

            # Create sheets if they have stats
            if skater_stats:
                self._create_simple_sheet("Skater Projections", ["playerName"] + skater_stats + ["TOTAL"])
                self._setup_total_formulas("Skater Projections", skater_stats, league_settings)
            if goalie_stats:
                self._create_simple_sheet("Goalie Projections", ["playerName"] + goalie_stats + ["TOTAL"])
                self._setup_total_formulas("Goalie Projections", goalie_stats, league_settings)

        except Exception as e:
            self.logger.error(f"Error setting up projection sheets: {e}")

    def _setup_total_formulas(self, sheet_name: str, stat_names, league_settings):
        """Set up TOTAL column formulas for projection sheets."""
        try:
            numbers_abs = os.path.abspath(self.filename)
            ptype = 'G' if 'Goalie' in sheet_name else 'P'

            # Build stat values map
            values = {}
            for stat in league_settings.get('stat_categories', []):
                if stat.get('position_type') == ptype:
                    name = stat.get('display_name') or stat.get('name') or ''
                    try:
                        values[name] = float(stat.get('value') or 0)
                    except Exception:
                        values[name] = 0

            # Calculate TOTAL column index (playerName + stats + TOTAL)
            total_col_index = len(stat_names) + 2

            # Build formula for TOTAL column
            # Format: =B2*value1+C2*value2+D2*value3...
            formula_parts = []
            for i, stat_name in enumerate(stat_names):
                val = values.get(stat_name, 0)
                if val:
                    col_letter = chr(66 + i)  # B, C, D, etc.
                    formula_parts.append(f"{col_letter}2*{val}")

            if not formula_parts:
                return

            # Build formula commands for each row
            formula_commands = []
            for row_num in range(2, 102):  # Rows 2-101 (100 rows)
                # Build formula for this row
                row_formula_parts = []
                for i, stat_name in enumerate(stat_names):
                    val = values.get(stat_name, 0)
                    if val:
                        col_letter = chr(66 + i)  # B, C, D, etc.
                        # Numbers uses comma as decimal separator
                        val_str = str(val).replace('.', ',')
                        row_formula_parts.append(f"{col_letter}{row_num}*{val_str}")

                if row_formula_parts:
                    row_formula = "+".join(row_formula_parts)
                    formula_commands.append(f'tell cell {total_col_index} of row {row_num}')
                    formula_commands.append(f'    set formulaStr to "=" & "{row_formula}"')
                    formula_commands.append(f'    set its value to formulaStr')
                    formula_commands.append(f'end tell')

            # Apply formulas in batches
            batch_size = 25  # 25 rows at a time
            lines_per_row = 4
            total_batches = (100 + batch_size - 1) // batch_size

            for batch_num in range(total_batches):
                start_idx = batch_num * batch_size * lines_per_row
                end_idx = start_idx + (batch_size * lines_per_row)
                batch = formula_commands[start_idx:end_idx]
                formulas_script = '\n                    '.join(batch)

                script = f'''
tell application "Numbers"
    try
        set doc to open (POSIX file "{numbers_abs}")
        tell doc
            tell sheet "{sheet_name}"
                tell table 1
                    {formulas_script}
                end tell
            end tell
        end tell
        save doc
        close doc
    on error errorMessage
        return "ERROR: " & errorMessage
    end try
end tell
'''

                res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=60)
                if res.returncode != 0:
                    self.logger.error(f"Failed to set TOTAL formulas for {sheet_name} batch {batch_num + 1}: {res.stderr}")
                else:
                    self.logger.debug(f"Set TOTAL formulas for {sheet_name} batch {batch_num + 1}/{total_batches}")
        except Exception as e:
            self.logger.error(f"Error setting TOTAL formulas for {sheet_name}: {e}")

    def _create_simple_sheet(self, sheet_name: str, headers):
        """Create a simple sheet with headers."""
        numbers_abs = os.path.abspath(self.filename)

        # Format headers for AppleScript
        header_commands = []
        for i, header in enumerate(headers, 1):
            escaped_header = str(header).replace('"', '\\"')
            header_commands.append(f'set value of cell {i} of row 1 to "{escaped_header}"')

        headers_script = '\n                        '.join(header_commands)

        script = f'''
tell application "Numbers"
    try
        set doc to open (POSIX file "{numbers_abs}")
        tell doc
            -- Check if sheet exists
            set sheetExists to false
            repeat with s in sheets
                if name of s is "{sheet_name}" then
                    set sheetExists to true
                    exit repeat
                end if
            end repeat

            if not sheetExists then
                set newSheet to make new sheet
                set name of newSheet to "{sheet_name}"

                tell newSheet
                    -- Set table name to be predictable
                    set name of table 1 to "{sheet_name} Table"
                    tell table 1
                        -- Set column count
                        set column count to {len(headers)}
                        -- Set headers
                        {headers_script}
                    end tell
                end tell
            end if
        end tell
        save doc
        close doc
    on error errorMessage
        return "ERROR: " & errorMessage
    end try
end tell
'''

        try:
            res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=15)
            if res.returncode == 0:
                self.logger.debug(f"Created/updated {sheet_name}")
            else:
                self.logger.error(f"Failed to create {sheet_name}: {res.stderr}")
        except subprocess.TimeoutExpired:
            self.logger.error(f"Timeout creating {sheet_name}")
        except Exception as e:
            self.logger.error(f"Error creating {sheet_name}: {e}")

    def create_draft_board(self):
        """Create Draft Board with formulas referencing Pre-Draft Analysis sheet."""
        try:
            numbers_abs = os.path.abspath(self.filename)

            # Use fixed table names (we name all tables as "{Sheet Name} Table")
            pda_table = "Pre-Draft Analysis Table"
            dr_table = "Draft Results Table"
            skater_table = "Skater Projections Table"
            goalie_table = "Goalie Projections Table"

            # Get max row from Pre-Draft Analysis sheet
            get_max_row_script = f'''
tell application "Numbers"
    try
        set doc to open (POSIX file "{numbers_abs}")
        tell doc
            tell sheet "Pre-Draft Analysis"
                tell table 1
                    set maxRow to row count
                end tell
            end tell
        end tell
        close doc saving no
        return maxRow
    on error errorMessage
        return "ERROR: " & errorMessage
    end try
end tell
'''

            res = subprocess.run(["osascript", "-e", get_max_row_script], capture_output=True, text=True, timeout=10)
            if res.returncode != 0:
                self.logger.error(f"Failed to get row count: {res.stderr}")
                return

            try:
                max_row = int(res.stdout.strip())
            except ValueError:
                self.logger.error(f"Could not parse max row: {res.stdout.strip()}")
                return

            # Create Draft Board sheet with headers
            headers = self.BASE_SHEETS["Draft Board"]
            self._create_simple_sheet("Draft Board", headers)

            # Now add formulas for each row
            # Build formula commands for all rows
            formula_commands = []

            for r in range(2, max_row + 1):
                # Build formulas using actual table names
                # Numbers formula syntax: ='Sheet Name'::TableName::CellRef

                # Column A (draftedBy): INDEX/MATCH to lookup manager from Draft Results based on player key
                # Look up column E (manager) instead of column D (teamKey)
                formula_a = f"IF(ISERROR(INDEX('Draft Results'::'{dr_table}'::E;MATCH('Pre-Draft Analysis'::'{pda_table}'::A{r};'Draft Results'::'{dr_table}'::C;0)));\\\"\\\";INDEX('Draft Results'::'{dr_table}'::E;MATCH('Pre-Draft Analysis'::'{pda_table}'::A{r};'Draft Results'::'{dr_table}'::C;0)))"
                formula_commands.append(f'tell cell 1 of row {r}')
                formula_commands.append(f'    set formulaStr to "=" & "{formula_a}"')
                formula_commands.append(f'    set its value to formulaStr')
                formula_commands.append(f'end tell')

                # Column B (playerName): Direct reference to Pre-Draft Analysis
                formula_b = f"'Pre-Draft Analysis'::'{pda_table}'::B{r}"
                formula_commands.append(f'tell cell 2 of row {r}')
                formula_commands.append(f'    set formulaStr to "=" & "{formula_b}"')
                formula_commands.append(f'    set its value to formulaStr')
                formula_commands.append(f'end tell')

                # Column C (team): Direct reference to Pre-Draft Analysis
                formula_c = f"'Pre-Draft Analysis'::'{pda_table}'::C{r}"
                formula_commands.append(f'tell cell 3 of row {r}')
                formula_commands.append(f'    set formulaStr to "=" & "{formula_c}"')
                formula_commands.append(f'    set its value to formulaStr')
                formula_commands.append(f'end tell')

                # Column D (position): Direct reference to Pre-Draft Analysis
                formula_d = f"'Pre-Draft Analysis'::'{pda_table}'::D{r}"
                formula_commands.append(f'tell cell 4 of row {r}')
                formula_commands.append(f'    set formulaStr to "=" & "{formula_d}"')
                formula_commands.append(f'    set its value to formulaStr')
                formula_commands.append(f'end tell')

                # Column E (averagePick): Direct reference to Pre-Draft Analysis
                formula_e = f"'Pre-Draft Analysis'::'{pda_table}'::E{r}"
                formula_commands.append(f'tell cell 5 of row {r}')
                formula_commands.append(f'    set formulaStr to "=" & "{formula_e}"')
                formula_commands.append(f'    set its value to formulaStr')
                formula_commands.append(f'end tell')

                # Column F (projectedPoints): IF statement to check if Goalie or Skater
                # If position is "G", lookup in Goalie Projections, otherwise Skater Projections
                # Using INDEX/MATCH for exact matching (LOOKUP requires sorted data)
                formula_f = (
                    f"IF('Pre-Draft Analysis'::'{pda_table}'::D{r}=\\\"G\\\";"
                    f"IF(ISERROR(INDEX('Goalie Projections'::'{goalie_table}'::F;MATCH('Pre-Draft Analysis'::'{pda_table}'::B{r};'Goalie Projections'::'{goalie_table}'::A;0)));\\\"\\\";INDEX('Goalie Projections'::'{goalie_table}'::F;MATCH('Pre-Draft Analysis'::'{pda_table}'::B{r};'Goalie Projections'::'{goalie_table}'::A;0)));"
                    f"IF(ISERROR(INDEX('Skater Projections'::'{skater_table}'::I;MATCH('Pre-Draft Analysis'::'{pda_table}'::B{r};'Skater Projections'::'{skater_table}'::A;0)));\\\"\\\";INDEX('Skater Projections'::'{skater_table}'::I;MATCH('Pre-Draft Analysis'::'{pda_table}'::B{r};'Skater Projections'::'{skater_table}'::A;0))))"
                )
                formula_commands.append(f'tell cell 6 of row {r}')
                formula_commands.append(f'    set formulaStr to "=" & "{formula_f}"')
                formula_commands.append(f'    set its value to formulaStr')
                formula_commands.append(f'end tell')

            # Apply formulas in batches to avoid script size issues
            # Each row has 6 cells * 3 lines = 18 lines per row
            batch_size = 50  # Process 50 rows at a time
            lines_per_row = 18  # 6 cells * 3 lines each
            total_batches = (max_row - 1 + batch_size - 1) // batch_size  # ceiling division

            for batch_num in range(total_batches):
                start_row = 2 + (batch_num * batch_size)
                end_row = min(start_row + batch_size, max_row + 1)

                # Calculate which formula commands belong to this batch
                start_idx = batch_num * batch_size * lines_per_row
                end_idx = start_idx + ((end_row - start_row) * lines_per_row)

                batch = formula_commands[start_idx:end_idx]
                formulas_script = '\n                    '.join(batch)

                script = f'''
tell application "Numbers"
    try
        set doc to open (POSIX file "{numbers_abs}")
        tell doc
            tell sheet "Draft Board"
                tell table 1
                    -- Ensure we have enough rows
                    set neededRows to {max_row}
                    if (row count) < neededRows then
                        repeat (neededRows - (row count)) times
                            add row below last row
                        end repeat
                    end if

                    -- Set formulas
                    {formulas_script}
                end tell
            end tell
        end tell
        save doc
        close doc
    on error errorMessage
        return "ERROR: " & errorMessage
    end try
end tell
'''

                res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=60)
                if res.returncode != 0:
                    self.logger.error(f"Failed to set Draft Board formulas (batch {batch_num + 1}/{total_batches}): {res.stderr}")
                else:
                    self.logger.debug(f"Set Draft Board formulas batch {batch_num + 1}/{total_batches} (rows {start_row}-{end_row - 1})")

            self.logger.debug(f"Created Draft Board with formulas for {max_row - 1} rows")

        except Exception as e:
            self.logger.error(f"Error creating Draft Board: {e}")


__all__ = ["MacOSDraftExporter"]