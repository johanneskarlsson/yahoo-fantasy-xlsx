import types
import subprocess
import pytest


@pytest.fixture(autouse=True)
def reset_seen_picks():
    # Reset global seen_picks before each test to avoid cross-test coupling
    from macos import draft_monitor
    draft_monitor.seen_picks.clear()
    yield
    draft_monitor.seen_picks.clear()


def test_collect_new_basic_and_ordering():
    from macos import draft_monitor
    draft_results = [
        {"pick": {"#text": "3"}, "round": {"#text": "2"}, "team_key": {"#text": "t2"}, "player_key": "p3"},
        {"pick": {"#text": "1"}, "round": {"#text": "1"}, "team_key": {"#text": "t1"}, "player_key": "p1"},
        {"pick": {"#text": "2"}, "round": {"#text": "1"}, "team_key": {"#text": "t1"}, "player_key": "p2"},
    ]
    rows = draft_monitor.collect_new(draft_results)
    # Sorted by pick number ascending
    picks = [r[1] for r in rows]
    assert picks == [1, 2, 3]
    # Seen picks recorded
    assert draft_monitor.seen_picks == {1, 2, 3}


def test_collect_new_skips_duplicates_and_invalid():
    from macos import draft_monitor
    first = [{"pick": {"#text": "1"}, "round": {"#text": "1"}, "team_key": {"#text": "t"}, "player_key": "px"}]
    draft_monitor.collect_new(first)
    # Duplicate plus invalid pick
    second = [
        {"pick": {"#text": "1"}, "round": {"#text": "1"}, "team_key": {"#text": "t"}, "player_key": "dup"},
        {"pick": {"#text": "not_int"}},
        {"pick": None},
        {"round": {"#text": "2"}},  # missing pick
    ]
    rows = draft_monitor.collect_new(second)
    assert rows == []  # No new valid picks


def test_rows_to_applescript_formatting_quotes_and_numbers():
    from macos import draft_monitor
    rows = [["Pick 1", 5, None, 'He said "Hi"']]
    script = draft_monitor._rows_to_applescript(rows)
    assert script.startswith('{') and script.endswith('}')
    assert '\\"Hi\\"' in script  # escaped inner quotes
    assert '5' in script
    assert '""' in script  # None becomes empty string


def test_append_picks_silently_success_triggers_manager_formulas(monkeypatch):
    from macos import draft_monitor

    calls = []

    def fake_run(args, capture_output, text, timeout):  # mimic subprocess.run
        # First call is append, second call is manager formulas row count, third call formulas assignment
        if len(calls) == 0:
            # Successful append
            res = types.SimpleNamespace(returncode=0, stdout="OK\n", stderr="")
        elif len(calls) == 1:
            # Row count retrieval for _set_manager_formulas
            res = types.SimpleNamespace(returncode=0, stdout="3\n", stderr="")
        else:
            # Formulas set script
            res = types.SimpleNamespace(returncode=0, stdout="", stderr="")
        calls.append(args)
        return res

    monkeypatch.setattr(subprocess, "run", fake_run)
    rows = [["1", 1, "p1", "t1", ""], ["1", 2, "p2", "t1", ""]]
    ok = draft_monitor.append_picks_silently(rows)
    assert ok is True
    # Expect at least 3 subprocess invocations: append + get row count + set formulas
    assert len(calls) >= 3


def test_append_picks_silently_error_output(monkeypatch):
    from macos import draft_monitor

    def fake_run(args, capture_output, text, timeout):
        return types.SimpleNamespace(returncode=0, stdout="ERROR: Draft Results sheet not found\n", stderr="")

    monkeypatch.setattr(subprocess, "run", fake_run)
    rows = [["1", 1, "p1", "t1", ""]]
    ok = draft_monitor.append_picks_silently(rows)
    assert ok is False


def test_set_manager_formulas_builds_expected_script(monkeypatch):
    from macos import draft_monitor

    run_calls = []

    def fake_run(args, capture_output, text, timeout):
        # First call returns row count 4, next call for formulas
        if not run_calls:
            res = types.SimpleNamespace(returncode=0, stdout="4\n", stderr="")
        else:
            # Capture formula script to inspect
            res = types.SimpleNamespace(returncode=0, stdout="", stderr="")
        run_calls.append(args)
        return res

    monkeypatch.setattr(subprocess, "run", fake_run)
    draft_monitor._set_manager_formulas()
    assert len(run_calls) >= 2
    # The second call script arg should contain a cell reference for row 2 manager formula
    script_arg = run_calls[1][2]
    assert 'cell 5 of row 2' in script_arg
    assert 'INDEX(' in script_arg


def test_set_manager_formulas_handles_rowcount_failure(monkeypatch, caplog):
    from macos import draft_monitor

    def fake_run_fail(args, capture_output, text, timeout):
        return types.SimpleNamespace(returncode=1, stdout="", stderr="boom")

    monkeypatch.setattr(subprocess, "run", fake_run_fail)
    with caplog.at_level("ERROR"):
        draft_monitor._set_manager_formulas()
    assert any("Failed to get row count" in r.message for r in caplog.records)
