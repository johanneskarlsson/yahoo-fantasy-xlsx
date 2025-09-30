#!/usr/bin/env python3

"""macOS-specific draft monitor that appends picks to an OPEN Numbers document without switching sheets."""
import os
import sys
import time
import logging
import subprocess
from dotenv import load_dotenv

# Add parent directory to path to import yahoo_api
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from yahoo_api import YahooFantasyAPI

load_dotenv()

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
LOG = logging.getLogger("draft_monitor")

INTERVAL = 10  # Check for new picks every 10 seconds

filename = os.getenv('FILENAME', 'fantasy_draft_data.numbers')
if not filename.lower().endswith('.numbers'):
    # Remove .xlsx if present and add .numbers
    if filename.lower().endswith('.xlsx'):
        filename = filename[:-5]
    filename = filename + '.numbers'

numbers_abs = os.path.abspath(filename)
api = YahooFantasyAPI()
seen_picks: set[int] = set()


def _scalar(v):
    if isinstance(v, dict):
        return v.get('#text') or v.get('full') or v.get('name') or ''
    return v


def _player_key(dr):
    return dr.get('player_key', '')


def collect_new(draft_results):
    """Collect new draft picks that haven't been seen yet."""
    rows = []
    for dr in draft_results:
        try:
            pick_raw = _scalar(dr.get('pick'))
            if pick_raw is None:
                continue
            pick_num = int(pick_raw)
            if pick_num in seen_picks:
                continue
            rnd = _scalar(dr.get('round')) or ''
            team_key = _scalar(dr.get('team_key')) or ''
            player_key = _player_key(dr) or ''
            rows.append([rnd, pick_num, player_key, team_key, ""])
            seen_picks.add(pick_num)
        except Exception:
            continue
    rows.sort(key=lambda r: r[1])
    return rows


def _rows_to_applescript(rows):
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
                escaped = str(cell).replace('"', '\\"')
                row_items.append(f'"{escaped}"')
        applescript_rows.append('{' + ', '.join(row_items) + '}')
    return '{' + ', '.join(applescript_rows) + '}'


def append_picks_silently(rows):
    """
    Append draft picks to Draft Results WITHOUT changing active sheet or closing document.

    Key differences from standard append:
    - Uses document 1 (assumes already open)
    - No open/close commands
    - No save command (user can save manually or auto-save will handle it)
    - Works on background sheet while user views Draft Board
    """
    if not rows:
        return True

    rows_script = _rows_to_applescript(rows)

    script = f'''
on run
    set newRows to {rows_script}

    tell application "Numbers"
        -- Work with the first open document (assumes user has Numbers open)
        if (count of documents) is 0 then
            return "ERROR: No Numbers document is open"
        end if

        tell document 1
            -- Check if Draft Results sheet exists
            set sheetExists to false
            repeat with s in sheets
                if name of s is "Draft Results" then
                    set sheetExists to true
                    exit repeat
                end if
            end repeat

            if not sheetExists then
                return "ERROR: Draft Results sheet not found"
            end if

            tell sheet "Draft Results"
                tell table 1
                    -- Find first empty row (skip header row 1)
                    set startRow to 2
                    set currentRows to row count
                    repeat with i from 2 to currentRows
                        set cellVal to value of cell 1 of row i
                        if cellVal is missing value or cellVal is "" then
                            set startRow to i
                            exit repeat
                        end if
                    end repeat

                    -- If all rows are full, we need to add new rows
                    if startRow > currentRows then
                        set startRow to currentRows + 1
                    end if

                    -- Calculate how many rows we need total
                    set neededRows to startRow + (count of newRows) - 1
                    if neededRows > currentRows then
                        repeat (neededRows - currentRows) times
                            add row below last row
                        end repeat
                    end if

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
    end tell
    return "OK"
end run
'''

    try:
        res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=10)
        if res.returncode == 0:
            output = res.stdout.strip()
            if output.startswith("ERROR"):
                LOG.error(f"AppleScript error: {output}")
                return False
            LOG.debug(f"Added {len(rows)} picks to Draft Results")

            # Now set formulas for manager column in a separate script
            _set_manager_formulas()

            return True
        else:
            LOG.error(f"Failed to append picks: {res.stderr}")
            return False
    except subprocess.TimeoutExpired:
        LOG.error("Timeout appending picks")
        return False
    except Exception as e:
        LOG.error(f"Error appending picks: {e}")
        return False


def _set_manager_formulas():
    """Set formulas for manager column in Draft Results sheet."""
    # First get row count
    get_rows_script = '''
tell application "Numbers"
    tell document 1
        tell sheet "Draft Results"
            tell table 1
                return row count
            end tell
        end tell
    end tell
end tell
'''
    try:
        res = subprocess.run(["osascript", "-e", get_rows_script], capture_output=True, text=True, timeout=5)
        if res.returncode != 0:
            LOG.error(f"Failed to get row count: {res.stderr}")
            return
        row_count = int(res.stdout.strip())
    except Exception as e:
        LOG.error(f"Failed to get row count: {e}")
        return

    # Build formula commands like we do in numbers_export
    formula_commands = []
    for r in range(2, row_count + 1):
        # Column E (manager): INDEX/MATCH lookup from Teams sheet
        formula_e = f"IF(D{r}=\\\"\\\";\\\"\\\";INDEX('Teams'::'Teams Table'::D;MATCH(D{r};'Teams'::'Teams Table'::A;0)))"
        formula_commands.append(f'tell cell 5 of row {r}')
        formula_commands.append(f'    set formulaStr to "=" & "{formula_e}"')
        formula_commands.append(f'    set its value to formulaStr')
        formula_commands.append(f'end tell')

    formulas_script = '\n                    '.join(formula_commands)

    script = f'''
tell application "Numbers"
    tell document 1
        tell sheet "Draft Results"
            tell table 1
                {formulas_script}
            end tell
        end tell
    end tell
end tell
'''
    try:
        res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=10)
        if res.returncode != 0:
            LOG.error(f"Failed to set manager formulas: {res.stderr}")
        else:
            LOG.debug("Set manager formulas successfully")
    except Exception as e:
        LOG.error(f"Failed to set manager formulas: {e}")


def main():
    print(f"🏒 Yahoo Fantasy Draft Monitor (macOS)")
    print(f"📊 Monitoring: {filename}")
    print(f"⏱️  Checking every {INTERVAL} seconds")
    print()
    print("⚠️  IMPORTANT: Keep the Numbers document OPEN while monitoring!")
    print("    New picks will be added silently to Draft Results sheet.")
    print("    You can keep Draft Board active - it won't switch sheets.")
    print()

    # Test API connection
    try:
        print("🔗 Testing Yahoo API connection...")
        api.ensure_authenticated()
        print("✅ Connected to Yahoo API")
    except Exception as e:
        print(f"❌ Failed to connect to Yahoo API: {e}")
        print("Please run 'python setup.py' first to authenticate")
        return

    # Check if Numbers document is open
    check_open = '''
tell application "Numbers"
    if (count of documents) is 0 then
        return "CLOSED"
    else
        return "OPEN"
    end if
end tell
'''
    try:
        res = subprocess.run(["osascript", "-e", check_open], capture_output=True, text=True)
        if res.returncode == 0 and res.stdout.strip() == "CLOSED":
            print()
            print(f"⚠️  WARNING: No Numbers document appears to be open!")
            print(f"    Please open {filename} before starting the monitor.")
            response = input("\n    Continue anyway? (y/n): ").lower().strip()
            if response != 'y':
                print("Exiting...")
                return
    except Exception:
        pass

    print()
    print("🔄 Monitoring... (Press Ctrl+C to stop)")
    print()

    polls = 0
    try:
        while True:
            start = time.time()
            try:
                results = api.get_draft_results() or []
                polls += 1

                # Log draft status on first poll
                if polls == 1:
                    if results:
                        print(f"📋 Found {len(results)} total draft picks so far")
                    else:
                        print("📋 No draft picks found yet")

                new_rows = collect_new(results)

                if new_rows:
                    success = append_picks_silently(new_rows)
                    if success:
                        for r in new_rows:
                            print(f"✅ Pick {r[1]}: {r[2]} (Team: {r[3]})")
                    else:
                        print(f"⚠️  Error saving {len(new_rows)} draft picks")
                else:
                    # Show periodic status so user knows it's working
                    if polls % 6 == 1:  # Every ~60 seconds (6 polls × 10s)
                        print(f"⏳ Still monitoring... ({polls} checks completed)")

            except Exception as e:
                print(f"⚠️  Error during check #{polls}: {e}")
                # Continue monitoring even if one check fails

            remaining = INTERVAL - (time.time() - start)
            if remaining > 0:
                time.sleep(remaining)
    except KeyboardInterrupt:
        print(f"\n🛑 Stopped by user after {polls} checks.")


if __name__ == "__main__":
    main()