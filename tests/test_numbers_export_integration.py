import os
import platform
import shutil
import subprocess
import tempfile
import pytest


pytestmark = pytest.mark.integration


@pytest.mark.skipif(platform.system() != "Darwin", reason="Numbers integration requires macOS")
@pytest.mark.skipif(shutil.which("osascript") is None, reason="osascript command not available")
@pytest.mark.skipif(os.environ.get("RUN_NUMBERS_IT") != "1", reason="Set RUN_NUMBERS_IT=1 to enable Numbers integration tests")
def test_numbers_end_to_end_creation_and_formulas():
    """End-to-end smoke test that exercises actual Numbers file creation.

    Steps:
    - Create exporter and write a small Pre-Draft Analysis dataset
    - Setup projection sheets with weights (Goals=4, Assists=3)
    - Create draft board
    - Inject sample stat values into Skater Projections row 2 (Goals=10, Assists=5)
    - Recalculate TOTAL (Numbers does automatically) and read it back
    - Assert sheets exist & sample cells populated
    """
    from macos.numbers_export import MacOSDraftExporter

    tmpdir = tempfile.mkdtemp(prefix="yf_numbers_it_")
    try:
        path = os.path.join(tmpdir, "itest.numbers")
        exp = MacOSDraftExporter(path)

        # Minimal Pre-Draft Analysis row (matches header order)
        pda_rows = [[
            "k1",  # playerKey
            "Player 1",  # playerName
            "NYR",  # team
            "C",    # position
            "10",   # averagePick
            "1",    # averageRound
            "100",  # percentDrafted
            "20",   # projectedAuctionValue
            "18",   # averageAuctionCost
            "5",    # seasonRank
            "2",    # positionRank
            "12",   # preseasonAveragePick
            "90",   # preseasonPercentDrafted
        ]]
        exp.update_draft_analysis_data(pda_rows)

        league_settings = {
            "stat_categories": [
                {"position_type": "P", "display_name": "Goals", "value": "4"},
                {"position_type": "P", "display_name": "Assists", "value": "3"},
            ]
        }
        exp.setup_projection_sheets(league_settings)
        exp.create_draft_board()
        exp.timestamp()

        # AppleScript to inject stat values and read back data.
        script = f'''
tell application "Numbers"
    set theDoc to open (POSIX file "{path}")
    tell theDoc
        -- Ensure Skater Projections sheet exists
        tell sheet "Skater Projections"
            tell table 1
                -- Insert stat values in row 2 (Goals=B2, Assists=C2)
                set value of cell 2 of row 2 to 10
                set value of cell 3 of row 2 to 5
                set totalVal to value of cell 4 of row 2 -- TOTAL is 4th column (playerName + 2 stats + TOTAL)
            end tell
        end tell
        -- Read back some Pre-Draft Analysis & Draft Board values
        tell sheet "Pre-Draft Analysis"
            tell table 1
                set pdaName to value of cell 2 of row 2
            end tell
        end tell
        tell sheet "Draft Board"
            tell table 1
                set dbName to value of cell 2 of row 2
            end tell
        end tell
        -- Collect sheet names
    set sheetNames to {{}}
        repeat with s in sheets
            copy (name of s) to end of sheetNames
        end repeat
        close theDoc saving no
    end tell
end tell
return (sheetNames as string) & "|" & pdaName & "|" & dbName & "|" & totalVal
'''

        res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=60)
        assert res.returncode == 0, f"AppleScript failed: {res.stderr or res.stdout}"
        output = res.stdout.strip()
        parts = output.split("|")
        assert len(parts) == 4, f"Unexpected output: {output}"
        sheet_list, pda_name, db_name, total_val = parts

        # Assertions
        assert "Pre-Draft Analysis" in sheet_list
        assert "Skater Projections" in sheet_list
        assert "Draft Board" in sheet_list
        assert pda_name == "Player 1"
        assert db_name == "Player 1"  # Draft Board playerName reference
        # TOTAL should be Goals*4 + Assists*3 = 10*4 + 5*3 = 55
        digits = ''.join(ch for ch in total_val if ch.isdigit())
        assert digits.startswith("55"), f"Expected total 55, got {total_val}"
    finally:
        try:
            shutil.rmtree(tmpdir)
        except Exception:
            pass
