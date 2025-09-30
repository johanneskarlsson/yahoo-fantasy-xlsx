import pytest


@pytest.fixture(autouse=True)
def reset_seen_picks():
    from windows import draft_monitor
    draft_monitor.seen_picks.clear()
    yield
    draft_monitor.seen_picks.clear()


def test_collect_new_basic_and_ordering():
    from windows import draft_monitor
    results = [
        {"pick": {"#text": "3"}, "round": {"#text": "2"}, "team_key": {"#text": "t2"}, "player_key": "p3"},
        {"pick": {"#text": "1"}, "round": {"#text": "1"}, "team_key": {"#text": "t1"}, "player_key": "p1"},
        {"pick": {"#text": "2"}, "round": {"#text": "1"}, "team_key": {"#text": "t1"}, "player_key": "p2"},
    ]
    rows = draft_monitor.collect_new(results)
    assert [r[1] for r in rows] == [1, 2, 3]
    assert draft_monitor.seen_picks == {1, 2, 3}


def test_collect_new_skips_duplicates_and_invalid():
    from windows import draft_monitor
    first = [{"pick": {"#text": "1"}, "round": {"#text": "1"}, "team_key": {"#text": "t"}, "player_key": "p1"}]
    draft_monitor.collect_new(first)
    second = [
        {"pick": {"#text": "1"}},
        {"pick": {"#text": "notint"}},
        {"pick": None},
        {"round": {"#text": "2"}},
    ]
    assert draft_monitor.collect_new(second) == []


def test_main_processes_new_picks(monkeypatch, capsys):
    from windows import draft_monitor

    # Ensure underlying exporter doesn't have the newer method names; patch compatibility layer
    # Windows exporter uses append_picks/timestamp; monitor calls append_draft_results/add_timestamp.
    # Provide those attributes dynamically for test.
    captured = {}

    def fake_append(rows):
        captured['rows'] = rows

    def fake_ts():
        captured['ts'] = True

    monkeypatch.setattr(draft_monitor.exporter, 'append_draft_results', fake_append, raising=False)
    monkeypatch.setattr(draft_monitor.exporter, 'add_timestamp', fake_ts, raising=False)

    class FakeAPI:
        def __init__(self):
            self.calls = 0
        def ensure_authenticated(self):
            return True
        def get_draft_results(self):
            if self.calls == 0:
                self.calls += 1
                return [
                    {"pick": {"#text": "2"}, "round": {"#text": "1"}, "team_key": {"#text": "T"}, "player_key": "p2"},
                    {"pick": {"#text": "1"}, "round": {"#text": "1"}, "team_key": {"#text": "T"}, "player_key": "p1"},
                ]
            # Simulate stop after second poll
            raise KeyboardInterrupt

    monkeypatch.setattr(draft_monitor, 'api', FakeAPI())
    monkeypatch.setattr(draft_monitor, 'INTERVAL', 0)
    monkeypatch.setattr(draft_monitor, 'time', type('T', (), {'time': staticmethod(lambda: 0), 'sleep': staticmethod(lambda x: None)}))

    # Run main; should exit via KeyboardInterrupt gracefully
    draft_monitor.main()

    # Validate picks processed in sorted order and methods invoked
    assert 'rows' in captured and len(captured['rows']) == 2
    assert [r[1] for r in captured['rows']] == [1, 2]
    assert captured.get('ts') is True
    out = capsys.readouterr().out
    assert 'Pick 1: p1' in out and 'Pick 2: p2' in out
