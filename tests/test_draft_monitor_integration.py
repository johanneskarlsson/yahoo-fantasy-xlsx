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
def test_append_picks_inserts_rows_in_draft_results(monkeypatch):
    """End-to-end: ensure append_picks_silently writes rows into an OPEN Draft Results sheet.

    We intentionally bypass manager formula setup to focus on insertion correctness.
    """
    from macos.numbers_export import MacOSDraftExporter
    from macos import draft_monitor

    # Skip manager formulas (we may not have Teams sheet in this minimal scenario)
    monkeypatch.setattr(draft_monitor, '_set_manager_formulas', lambda: None)

    tmpdir = tempfile.mkdtemp(prefix="yf_dm_it_")
    try:
        path = os.path.join(tmpdir, "itest.numbers")
        exporter = MacOSDraftExporter(path)

        # Create base document by injecting a minimal Pre-Draft Analysis sheet (ensures file exists)
        exporter.update_draft_analysis_data([
            ["k1", "Player 1", "NYR", "C", "10", "1", "100", "20", "18", "5", "2", "12", "90"],
        ])

        # Create Draft Results sheet with expected headers using exporter helper
        headers = exporter.BASE_SHEETS['Draft Results']
        exporter._create_simple_sheet('Draft Results', headers)

        # Open the document and leave it open
        open_script = f'''tell application "Numbers" to open (POSIX file "{path}")'''
        res = subprocess.run(["osascript", "-e", open_script], capture_output=True, text=True, timeout=10)
        assert res.returncode == 0, f"Failed to open Numbers doc: {res.stderr}"

        # Prepare two picks (round, pick, player_key, team_key, manager placeholder)
        rows = [["1", 1, "k1", "t1", ""], ["1", 2, "k2", "t2", ""]]
        # Clear seen picks to ensure they register
        draft_monitor.seen_picks.clear()
        ok = draft_monitor.append_picks_silently(rows)
        assert ok is True

        # Read back inserted rows via AppleScript (values from rows 2 and 3)
        read_script = f'''
tell application "Numbers"
    tell document 1
        tell sheet "Draft Results"
            tell table 1
                set r2_pick to value of cell 2 of row 2
                set r2_player to value of cell 3 of row 2
                set r3_pick to value of cell 2 of row 3
                set r3_player to value of cell 3 of row 3
            end tell
        end tell
        close saving no
    end tell
end tell
return (r2_pick as string) & "," & (r2_player as string) & ";" & (r3_pick as string) & "," & (r3_player as string)
'''
        res = subprocess.run(["osascript", "-e", read_script], capture_output=True, text=True, timeout=15)
        assert res.returncode == 0, f"Failed to read back rows: {res.stderr or res.stdout}"
        out = res.stdout.strip()
        first, second = out.split(';')
        assert first.startswith('1,')  # pick number 1
        assert ',k1' in first or ',k1' == first.split(',')[1]  # player key present
        assert second.startswith('2,')
        assert ',k2' in second or ',k2' == second.split(',')[1]
    finally:
        try:
            shutil.rmtree(tmpdir)
        except Exception:
            pass
