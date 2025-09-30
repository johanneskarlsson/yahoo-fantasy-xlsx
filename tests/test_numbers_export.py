import os
import subprocess
from types import SimpleNamespace
import pytest


@pytest.fixture()
def exporter(tmp_path):
    from macos.numbers_export import MacOSDraftExporter
    # Use temp filename to avoid any accidental collisions
    fname = tmp_path / "test_export.numbers"
    return MacOSDraftExporter(str(fname))


class RunRecorder:
    """Utility to capture subprocess.run calls for assertions."""
    def __init__(self, returncode=0, stdout="", stderr=""):
        self.calls = []
        self.returncode = returncode
        self.stdout = stdout
        self.stderr = stderr

    def __call__(self, *args, **kwargs):  # mimic subprocess.run signature subset
        self.calls.append(SimpleNamespace(args=args, kwargs=kwargs))
        # Provide object with attributes used by code
        return SimpleNamespace(returncode=self.returncode, stdout=self.stdout, stderr=self.stderr)


def test_rows_to_applescript_format(exporter):
    rows = [["Alpha", 'He said "Hello"', 42, None, 3.14]]
    script = exporter._rows_to_applescript(rows)
    # Expect outer braces and escaped quote
    assert script.startswith('{') and script.endswith('}')
    assert '\\"Hello\\"' in script  # escaped quotes
    # Numeric 42 should appear unquoted
    assert '42' in script
    # None should become empty quotes
    assert '{"Alpha", "He said \\"Hello\\"", 42, "", 3.14}' in script or '""' in script


def test_update_league_settings_data_builds_expected_rows(exporter, monkeypatch):
    captured = {}

    def fake_update(sheet_name, data_rows):
        captured['sheet'] = sheet_name
        captured['rows'] = data_rows

    monkeypatch.setattr(exporter, '_update_sheet_data_simple', fake_update)

    league_settings = {
        'league_name': 'Test League',
        'league_type': 'H2H',
        'scoring_type': 'points',
        'max_teams': '12',
        'num_playoff_teams': '6',
        'playoff_start_week': '20',
        'roster_positions': [
            {'position': 'C', 'count': '2'},
            {'position': 'G', 'count': '2'},
        ],
        'stat_categories': [
            {'position_type': 'P', 'display_name': 'G', 'value': '4'},
            {'position_type': 'P', 'display_name': 'A', 'value': '3'},
            {'position_type': 'G', 'display_name': 'W', 'value': '5'},
        ],
    }

    exporter.update_league_settings_data(league_settings)

    assert captured['sheet'] == 'League Settings'
    # Ensure roster positions & stat headers present
    flat = ['|'.join(map(str, r)) for r in captured['rows']]
    assert any('ROSTER POSITIONS' in r for r in flat)
    assert any('SKATER STATS' in r for r in flat)
    assert any('GOALIE STATS' in r for r in flat)
    # Check a skater stat row and goalie stat row included
    assert any(row[0] == 'G' and row[1] == '4' for row in captured['rows'])
    assert any(row[0] == 'W' and row[1] == '5' for row in captured['rows'])


def test_setup_projection_sheets_calls_helpers(exporter, monkeypatch):
    created = []
    formulas = []

    def fake_create(sheet_name, headers):
        created.append((sheet_name, headers))

    def fake_setup_formulas(sheet_name, stat_names, league_settings):
        formulas.append((sheet_name, tuple(stat_names)))

    monkeypatch.setattr(exporter, '_create_simple_sheet', fake_create)
    monkeypatch.setattr(exporter, '_setup_total_formulas', fake_setup_formulas)

    league_settings = {
        'stat_categories': [
            {'position_type': 'P', 'display_name': 'Goals', 'value': '4'},
            {'position_type': 'P', 'display_name': 'Assists', 'value': '3'},
            {'position_type': 'G', 'display_name': 'Wins', 'value': '5'},
        ]
    }

    exporter.setup_projection_sheets(league_settings)

    # Assert skater & goalie sheets created with TOTAL column
    assert ('Skater Projections', ['playerName', 'Goals', 'Assists', 'TOTAL']) in created
    assert ('Goalie Projections', ['playerName', 'Wins', 'TOTAL']) in created
    # Formulas configured
    assert any(s == 'Skater Projections' for s, _ in formulas)
    assert any(s == 'Goalie Projections' for s, _ in formulas)


def test_update_draft_analysis_data_calls_import(exporter, monkeypatch):
    called = {}

    def fake_import(sheet_name, headers, rows):
        called['sheet'] = sheet_name
        called['headers'] = headers
        called['rows'] = rows

    monkeypatch.setattr(exporter, '_import_sheet_via_csv', fake_import)

    rows = [
        ['key1', 'Player 1', 'NYR', 'C', '10', '1', '100', '20', '18', '5', '2', '12', '90'],
        ['key2', 'Player 2', 'BOS', 'G', '20', '2', '80', '15', '12', '10', '1', '18', '70'],
    ]
    exporter.update_draft_analysis_data(rows)
    assert called['sheet'] == 'Pre-Draft Analysis'
    assert len(called['headers']) == 13  # expected header count
    assert len(called['rows']) == 2


def test_import_sheet_via_csv_executes_applescript(exporter, monkeypatch, tmp_path):
    # Real path interactions, but mock subprocess.run so no AppleScript executes.
    recorder = RunRecorder(stdout="", returncode=0)
    monkeypatch.setattr(subprocess, 'run', recorder)

    # Provide minimal rows to pass guards
    headers = exporter.BASE_SHEETS['Pre-Draft Analysis']
    rows = [
        ['k', 'Name', 'T', 'C', '1', '1', '90', '15', '12', '10', '1', '20', '80']
    ]

    exporter._import_sheet_via_csv('Pre-Draft Analysis', headers, rows)

    # Ensure subprocess.run was called at least once with osascript
    assert any('osascript' in call.args[0][0] for call in recorder.calls)
    # Confirm temp file cleaned up (deleted) -> we cannot rely on deletion because we cannot predict path
    # Instead ensure script contained sheet name (in args)
    found_sheet_name = False
    for call in recorder.calls:
        if any('Pre-Draft Analysis' in str(arg) for arg in call.args):
            found_sheet_name = True
            break
    assert found_sheet_name


def test_create_simple_sheet_runs_osascript(exporter, monkeypatch):
    recorder = RunRecorder(stdout="", returncode=0)
    monkeypatch.setattr(subprocess, 'run', recorder)

    exporter._create_simple_sheet('Unit Test Sheet', ['H1', 'H2'])
    # Validate an osascript invocation happened
    assert any(call.args[0][0] == 'osascript' for call in recorder.calls)
    # Rough sanity: script should include sheet name
    assert any('Unit Test Sheet' in call.args[0][2] for call in recorder.calls)


def test_bulk_update_sheet_builds_expected_script(exporter, monkeypatch):
    # Provide minimal data rows and intercept subprocess.run
    recorder = RunRecorder(stdout="", returncode=0)
    monkeypatch.setattr(subprocess, 'run', recorder)

    data_rows = [
        ['a', 'b', 'c'],
        ['d', 'e', 'f'],
    ]
    exporter._bulk_update_sheet('Some Sheet', data_rows)
    assert any('osascript' in call.args[0][0] for call in recorder.calls)
    # Ensure we reference sheet name and at least one data value
    combined_script = '\n'.join(call.args[0][2] for call in recorder.calls)
    assert 'Some Sheet' in combined_script
    assert 'a' in combined_script and 'f' in combined_script


def test_update_sheet_data_simple_small_dataset(exporter, monkeypatch):
    # For small dataset path: single call to _update_sheet_chunk
    invoked = {}

    def fake_chunk(sheet_name, rows, start_row, numbers_abs):
        invoked['sheet'] = sheet_name
        invoked['rows'] = rows
        invoked['start'] = start_row

    monkeypatch.setattr(exporter, '_update_sheet_chunk', fake_chunk)
    exporter._update_sheet_data_simple('Teams', [[1,2,3]])
    assert invoked['sheet'] == 'Teams'
    assert invoked['start'] == 2  # expected start row


def test_update_sheet_data_simple_chunked(exporter, monkeypatch):
    # Force chunking by using >100 rows
    calls = []

    def fake_chunk(sheet_name, rows, start_row, numbers_abs):
        calls.append((sheet_name, start_row, len(rows)))

    monkeypatch.setattr(exporter, '_update_sheet_chunk', fake_chunk)
    big_rows = [[i] for i in range(150)]
    exporter._update_sheet_data_simple('Teams', big_rows)
    # Expect two chunks: first start at 2, second at 102
    starts = [c[1] for c in calls]
    assert 2 in starts and 102 in starts
    assert sum(c[2] for c in calls) == 150
