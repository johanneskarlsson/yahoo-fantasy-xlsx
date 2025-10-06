import re
import os, logging, subprocess, csv, tempfile
from typing import List, Dict, Any, Iterable, Tuple
from datetime import datetime
from .numbers_helpers import create_sheets, update_sheet, apply_formulas


class MacOSDraftExporter:
    """Mac exporter using pure AppleScript for Numbers - no XLSX intermediate files."""

    BASE_SHEETS = {
        "Draft Board": ["", "draftedBy", "playerKey", "playerName", "team", "position", "averagePick", "projectedPoints", "vorp"],
        "Positions": ["playerKey", "playerName", "team", "position", "averagePick", "projectedPoints", "rank", "vorp"],
        "League Settings": ["setting", "value"],
        "Teams": ["teamKey", "teamId", "teamName", "manager"],
        "Draft Results": ["round", "pick", "playerKey", "teamKey", "manager"],
    }

    def __init__(self, filename: str = "fantasy_draft_data.numbers"):
        # Override parent to use .numbers instead of .xlsx
        if not filename.lower().endswith('.numbers'):
            filename = filename.replace('.xlsx', '.numbers') if filename.endswith('.xlsx') else filename + '.numbers'

        self.filename = filename
        self.logger = logging.getLogger(__name__)

    # ---------------------------- Sheet helpers ----------------------------

    def create_draft_board(self, players_rows):
        if not players_rows:
            return
        headers = self.BASE_SHEETS.get("Draft Board", [])
        try:
            self._create_draft_board_with_csv("Draft Board", headers, players_rows)
        except Exception as e:
            self.logger.error(f"Error creating Draft Board via CSV import: {e}")


    def _create_draft_board_with_csv(self, sheet_name: str, headers, rows):
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
                    writer.writerow(['', ''] + converted_row)
        except Exception as e:
            try:
                os.remove(temp_path)
            except Exception:
                pass
            raise RuntimeError(f"Failed writing temp CSV: {e}")

        script = f'''
tell application "Numbers"
    with timeout of 3600 seconds
        -- Close any existing document with same path
        set targetPath to POSIX file "{numbers_abs}"
        repeat with doc in documents
            try
                if path of doc is (targetPath as text) then
                    close doc saving no
                end if
            end try
        end repeat

        -- Phase 1: import CSV -> Numbers doc (no formulas yet)
        set csvDoc to open (POSIX file "{temp_path}")
        delay 0.5
        tell csvDoc
            set name of sheet 1 to "{sheet_name}"
            tell sheet 1
                set name of table 1 to "{sheet_name}"
            end tell
        end tell

        -- Save as target .numbers file and close temp doc
        try
            save csvDoc in POSIX file "{numbers_abs}"
        on error errMsg number errNum
            try
                close csvDoc saving yes
            end try
        end try
        close csvDoc saving yes
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



    def create_pos_sheets(self, players_rows):
            """Create Position sheets (C, LW, RW, D, G) from players_rows.

            Supports multi-position strings like 'LW/RW', 'C-LW', 'LW,RW', 'C RW' etc.
            """
            if not players_rows:
                return
            pos_map = {'C': [], 'LW': [], 'RW': [], 'D': [], 'G': []}
            try:
                create_sheets(
                    self.filename,
                    self.logger,
                    [(f"{pos} Players", self.BASE_SHEETS.get("Positions", [])) for pos in pos_map.keys()],
                    force=False,
                )
            except Exception as e:  # pragma: no cover - defensive
                self.logger.debug(f"Position sheets ensure skipped/failed (may already exist): {e}")

            for row in players_rows:
                try:
                    position_field = str(row[3] or "")
                except Exception:
                    continue
                raw_tokens = re.split(r'[\/,;+\-\s]+', position_field)
                tokens = {t.strip() for t in raw_tokens if t.strip()}
                for token in tokens:
                    if token in pos_map:
                        pos_map[token].append(row)

            # Formulas (col F) pulling TOTAL from projection sheets
            skater_formula = "=IF(ISERROR(INDEX('Skater Projections'::TOTAL;MATCH(B{row};'Skater Projections'::playerName;0)));\"\";INDEX('Skater Projections'::TOTAL;MATCH(B{row};'Skater Projections'::playerName;0)))"
            goalie_formula = "=IF(ISERROR(INDEX('Goalie Projections'::TOTAL;MATCH(B{row};'Goalie Projections'::playerName;0)));\"\";INDEX('Goalie Projections'::TOTAL;MATCH(B{row};'Goalie Projections'::playerName;0)))"
            for pos, rows in pos_map.items():
                if not rows:
                    continue
                sheet_name = f"{pos} Players"
                update_sheet(self.filename, self.logger, sheet_name, rows)
                proj_formula = goalie_formula if pos == 'G' else skater_formula
                # Rank within the sheet (descending). Blank if no projected points.
                rank_formula = "=IF(F{row}=\"\";\"\";RANK(F{row};F$2:F$1000;0))"
                vorp_formula = f"=IF(F{{row}}=\"\";\"\";IF(ISERROR(MATCH(\"VORP_{pos}\";'League Settings'::A;0));\"\";IFERROR(F{{row}}-INDEX(F$2:F$1000;MATCH(INDEX('League Settings'::B;MATCH(\"VORP_{pos}\";'League Settings'::A;0));G$2:G$1000;0));\"\")))"
                try:
                    apply_formulas(
                        self.filename,
                        self.logger,
                        sheet=sheet_name,
                        per_row=[("F", proj_formula), ("G", rank_formula), ("H", vorp_formula)],
                        start_row=2,
                    )
                except Exception as e:  # pragma: no cover
                    self.logger.debug(f"Apply projectedPoints/rank formulas failed for {sheet_name}: {e}")
                self.logger.debug(f"Updated position sheet {sheet_name} with {len(rows)} players and applied projectedPoints + rank formulas")


    def update_league_settings_data(self, league_settings: Dict[str, Any]):
        """Populate League Settings sheet with grouped sections using AppleScript."""
        # Ensure the sheet exists (headers applied once) before attempting updates
        try:
            create_sheets(
                self.filename,
                self.logger,
                [("League Settings", self.BASE_SHEETS["League Settings"])],
                force=False,
            )
        except Exception as e:  # pragma: no cover - defensive
            self.logger.debug(f"League Settings sheet ensure skipped/failed (may already exist): {e}")
        rows = []
        rows.append(["League Name", league_settings.get('league_name', '')])
        rows.append(["League Type", league_settings.get('league_type', '')])
        rows.append(["Scoring Type", league_settings.get('scoring_type', '')])
        rows.append(["Max Teams", league_settings.get('max_teams', '')])
        rows.append(["Playoff Teams", league_settings.get('num_playoff_teams', '')])
        rows.append(["Playoff Start Week", league_settings.get('playoff_start_week', '')])
        rows.append(["", ""])
        rows.append(["ROSTER POSITIONS", "COUNT"])
        for pos in league_settings.get('roster_positions', []):
            rows.append([pos.get('position', ''), pos.get('count', '')])
        rows.append(["", ""])
        rows.append(["SKATER STATS", "VALUE"])
        for stat in league_settings.get('stat_categories', []):
            if stat.get('position_type') == 'P':
                name = stat.get('display_name') or stat.get('name') or ''
                rows.append([name, stat.get('value', '')])
        rows.append(["", ""])
        rows.append(["GOALIE STATS", "VALUE"])
        for stat in league_settings.get('stat_categories', []):
            if stat.get('position_type') == 'G':
                name = stat.get('display_name') or stat.get('name') or ''
                rows.append([name, stat.get('value', '')])

        rows.append(["", ""])
        # VORP baselines section (total roster slots per position = count * max_teams)
        rows.append(["VORP BASELINES", "VALUE"])
        max_teams_raw = league_settings.get('max_teams', 0)

        def _as_int(v):
            try:
                if v is None or v == '':
                    return 0
                return int(str(v).strip())
            except (ValueError, TypeError):
                return 0

        max_teams = _as_int(max_teams_raw)
        for pos in league_settings.get('roster_positions', []):
            p_code = "VORP_" + pos.get('position', '')
            count = _as_int(pos.get('count', 0))
            total_slots = count * max_teams if count and max_teams else ''
            rows.append([p_code, total_slots])
        update_sheet(self.filename, self.logger, "League Settings", rows)

    def update_teams_data(self, teams_rows):
        """Write Teams sheet data using AppleScript."""
        if not teams_rows:
            return
        # Ensure Teams sheet exists first
        try:
            create_sheets(
                self.filename,
                self.logger,
                [("Teams", self.BASE_SHEETS["Teams"])],
                force=False,
            )
        except Exception as e:  # pragma: no cover - defensive
            self.logger.debug(f"Teams sheet ensure skipped/failed (may already exist): {e}")
        update_sheet(self.filename, self.logger, "Teams", teams_rows)

    def update_draft_results_data(self, draft_results):
        """Write Draft Results sheet data (round, pick, playerKey, teamKey, manager(lookup)).

        draft_results example:
        {'pick': '1', 'round': '1', 'team_key': '465.l.118719.t.4', 'player_key': '465.p.6743'}
        Manager is now populated via formula lookup (Teams sheet) for consistency.
        """
        if not draft_results:
            try:
                self._preallocate_draft_results_rows()
            except Exception as e:  # pragma: no cover - defensive
                self.logger.debug(f"Preallocation of Draft Results rows skipped/failed: {e}")
            return
        try:
            create_sheets(
                self.filename,
                self.logger,
                [("Draft Results", self.BASE_SHEETS["Draft Results"])],
                force=False,
            )
        except Exception as e:  # pragma: no cover
            self.logger.debug(f"Draft Results sheet ensure skipped/failed (may already exist): {e}")

        rows = []
        for entry in draft_results:
            rnd = entry.get("round", "")
            pick = entry.get("pick", "")
            player_key = entry.get("player_key") or entry.get("playerKey") or ""
            team_key = entry.get("team_key") or entry.get("teamKey") or ""
            # Manager left blank here; filled by formula later
            rows.append([rnd, pick, player_key, team_key, ""])

        if rows:
            update_sheet(self.filename, self.logger, "Draft Results", rows)
            # Apply manager lookup formulas after data insert
            self._apply_draft_results_formulas()


    def _apply_draft_results_formulas(self):
        """Use generic helper to apply manager lookup (col E)."""
        manager_formula = "=IF(ISERROR(INDEX('Teams'::D;MATCH(D{row};'Teams'::A;0)));\"\";INDEX('Teams'::D;MATCH(D{row};'Teams'::A;0)))"
        apply_formulas(
            self.filename,
            self.logger,
            sheet="Draft Results",
            per_row=[("E", manager_formula)],
            start_row=2,
        )

    def _preallocate_draft_results_rows(self, target_rows: int = 1000):
        """Ensure Draft Results sheet has header + target_rows data rows (default 1000).

        This reduces the need for the runtime monitor to add rows one-by-one (AppleScript add row
        operations are relatively slow and can momentarily shift focus). We only ever append picks;
        having extra empty rows is harmless.
        """
        numbers_abs = os.path.abspath(self.filename)
        # +1 because row 1 is headers
        desired_total = target_rows + 1
        script = f'''
tell application "Numbers"
    try
        -- Open document if not already open
        set doc to missing value
        repeat with d in documents
            if path of d is "{numbers_abs}" then
                set doc to d
                exit repeat
            end if
        end repeat
        if doc is missing value then
            set doc to open (POSIX file "{numbers_abs}")
        end if

        tell doc
            if (every sheet whose name is "Draft Results") = {{}} then return "ERROR: Missing Draft Results"
            tell sheet "Draft Results"
                tell table 1
                    set currentRows to row count
                    if currentRows < {desired_total} then
                        repeat ({desired_total} - currentRows) times
                            add row below last row
                        end repeat
                    end if
                end tell
            end tell
        end tell
        save doc
        return "OK"
    on error errMsg
        return "ERROR: " & errMsg
    end try
end tell
'''
        try:
            res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=40)
            out = (res.stdout or "").strip()
            if out.startswith("ERROR:"):
                self.logger.debug(f"Draft Results preallocation AppleScript error: {out}")
            elif res.returncode != 0:
                self.logger.debug(f"Draft Results preallocation non-zero exit rc={res.returncode} stderr={res.stderr.strip()}")
            else:
                self.logger.debug(f"Draft Results preallocated to >= {target_rows} data rows")
        except subprocess.TimeoutExpired:
            self.logger.debug("Draft Results preallocation timeout (continuing without it)")
        except Exception as e:  # pragma: no cover
            self.logger.debug(f"Draft Results preallocation unexpected error: {e}")

    def apply_draft_board_formulas(self):
        """Apply Draft Board base formulas (A,B,H) then per-row dynamic VORP (I)."""
        a_formula = "=LEN(B{row})>0"
        b_formula = "=IF(ISERROR(INDEX('Draft Results'::E;MATCH(C{row};'Draft Results'::C;0)));\"\";INDEX('Draft Results'::E;MATCH(C{row};'Draft Results'::C;0)))"
        # Fix TOTAL casing (previously '::Total' caused lookup failure -> empty projections)
        h_formula = "=IF(F{row}=\"G\";IF(ISERROR(INDEX('Goalie Projections'::TOTAL;MATCH(D{row};'Goalie Projections'::playerName;0)));\"\";INDEX('Goalie Projections'::TOTAL;MATCH(D{row};'Goalie Projections'::playerName;0)));IF(ISERROR(INDEX('Skater Projections'::TOTAL;MATCH(D{row};'Skater Projections'::playerName;0)));\"\";INDEX('Skater Projections'::TOTAL;MATCH(D{row};'Skater Projections'::playerName;0))))"

        apply_formulas(
            self.filename,
            self.logger,
            sheet="Draft Board",
            per_row=[("A", a_formula), ("B", b_formula), ("H", h_formula)],
            start_row=2,
        )

        try:
            self._apply_row_specific_vorp()
        except Exception as e:  # pragma: no cover
            self.logger.debug(f"Per-row VORP application failed: {e}")

    def _apply_row_specific_vorp(self):
        numbers_abs = os.path.abspath(self.filename)
        read_script = f'''
tell application "Numbers"
    with timeout of 3600 seconds
        set doc to missing value
        set targetFile to POSIX file "{numbers_abs}"
        repeat with d in documents
            try
                if (path of d) is targetFile then
                    set doc to d
                    exit repeat
                end if
            end try
        end repeat
        if doc is missing value then
            set doc to open targetFile
        end if
        if (every sheet of doc whose name is "Draft Board") = {{}} then return ""
        set outList to {{}}
        tell sheet "Draft Board" of doc
            tell table 1
                set rc to row count
                repeat with r from 2 to rc
                    set pk to value of cell 3 of row r
                    set posStr to value of cell 6 of row r
                    if (pk is missing value or pk = "") and (posStr is missing value or posStr = "") then
                        -- skip empty line
                    else
                        if pk is missing value then set pk to ""
                        if posStr is missing value then set posStr to ""
                        copy (r as text) & "||" & pk & "||" & posStr to end of outList
                    end if
                end repeat
            end tell
        end tell
        return outList
    end timeout
end tell
'''
        res = subprocess.run(["osascript", "-e", read_script], capture_output=True, text=True, timeout=90)
        if res.returncode != 0:
            self.logger.debug(f"Read Draft Board rows AppleScript stderr={res.stderr.strip()}")
            return

        raw = (res.stdout or "").strip()
        if raw.startswith("{") and raw.endswith("}"):
            raw = raw[1:-1]
        entries = [e.strip() for e in raw.split(", ") if e.strip()]
        if not entries:
            self.logger.debug("No Draft Board rows found for VORP formula generation")
            return

        row_formula_map: Dict[int, str] = {}
        for line in entries:
            parts = line.split("||")
            if len(parts) != 3:
                continue
            try:
                row_idx = int(parts[0])
            except ValueError:
                continue
            player_key = parts[1].strip()
            pos_str = parts[2].strip()
            if not player_key or not pos_str:
                continue
            norm = pos_str.replace("/", ",").replace(";", ",").replace(" ", ",")
            tokens_raw = [t.strip() for t in norm.split(",") if t.strip()]
            seen = set()
            positions = []
            for t in tokens_raw:
                if t in ("C", "LW", "RW", "D", "G") and t not in seen:
                    seen.add(t)
                    positions.append(t)
            if not positions:
                continue
            row_formula_map[row_idx] = self._build_vorp_formula_for_positions(positions, row_idx)

        if not row_formula_map:
            self.logger.debug("No per-row VORP formulas constructed (positions missing)")
            return

        write_chunks = []
        for r, fmla in row_formula_map.items():
            esc = fmla.replace('"', '\\"')
            write_chunks.append(f'''
                try
                    tell row {r}
                        tell cell 9
                            set value to "{esc}"
                        end tell
                    end tell
                end try''')

        write_body = "\n".join(write_chunks)
        write_script = f'''
tell application "Numbers"
    with timeout of 3600 seconds
        set targetFile to POSIX file "{numbers_abs}"
        set doc to missing value
        repeat with d in documents
            try
                if (path of d) is targetFile then
                    set doc to d
                    exit repeat
                end if
            end try
        end repeat
        if doc is missing value then
            set doc to open targetFile
        end if
        if (every sheet of doc whose name is "Draft Board") = {{}} then return "ERROR: Missing Draft Board"
        tell sheet "Draft Board" of doc
            tell table 1
{write_body}
            end tell
        end tell
        save doc
        return "OK"
    end timeout
end tell
'''
        res2 = subprocess.run(["osascript", "-e", write_script], capture_output=True, text=True, timeout=180)
        if res2.returncode != 0:
            self.logger.debug(f"Write VORP formulas AppleScript stderr={res2.stderr.strip()}")
            return
        self.logger.debug(f"Applied VORP formulas to {len(row_formula_map)} rows (col I)")

    def _build_vorp_formula_for_positions(self, positions: List[str], row: int) -> str:
        """Return per‑row VORP formula choosing VORP from the sheet where rank (col G) is lowest.
        Uses INDEX/MATCH everywhere (no LOOKUP) for exact match reliability.
        Tie-break: earlier position in the player's position string wins.
        """
        if len(positions) == 1:
            p = positions[0]
            # Single position: directly pull that sheet's VORP (col H) via exact match on playerKey (col A)
            return (
                f"=IF(C{row}=\"\";\"\";"
                f"IFERROR(INDEX('{p} Players'::H;MATCH(C{row};'{p} Players'::A;0));\"\"))"
            )

        # For multi-position: build rank expressions with sentinel 9999
        rank_exprs = [
            f"IFERROR(INDEX('{p} Players'::G;MATCH(C{row};'{p} Players'::A;0));9999)"
            for p in positions
        ]
        min_list = ";".join(rank_exprs)

        # Nested IF chain (reverse order so first position in original list has priority on ties)
        nested = "\"\""
        for p, r_expr in reversed(list(zip(positions, rank_exprs))):
            vorp_expr = f"IFERROR(INDEX('{p} Players'::H;MATCH(C{row};'{p} Players'::A;0));\"\")"
            nested = f"IF({r_expr}=MIN({min_list});{vorp_expr};{nested})"

        return f"=IF(C{row}=\"\";\"\";{nested})"



    def setup_projection_sheets(self, league_settings):
        """Create Skater/Goalie Projections sheets with TOTAL formulas using bulk creation."""
        try:
            skater_stats: List[str] = []
            goalie_stats: List[str] = []
            for stat in league_settings.get('stat_categories', []):
                name = stat.get('display_name') or stat.get('name') or ''
                if not name:
                    continue
                ptype = stat.get('position_type')
                if ptype == 'P':
                    skater_stats.append(name)
                elif ptype == 'G':
                    goalie_stats.append(name)

            sheets_to_create: List[Tuple[str, List[str]]] = []
            if skater_stats:
                sheets_to_create.append(("Skater Projections", ["playerName"] + skater_stats + ["TOTAL"]))
            if goalie_stats:
                sheets_to_create.append(("Goalie Projections", ["playerName"] + goalie_stats + ["TOTAL"]))

            if sheets_to_create:
                create_sheets(self.filename, self.logger, sheets_to_create)

            if skater_stats:
                self._setup_total_formulas("Skater Projections", skater_stats, league_settings)
            if goalie_stats:
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




__all__ = ["MacOSDraftExporter"]