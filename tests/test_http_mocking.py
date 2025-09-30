import requests


def test_get_game_key_with_requests_mock(api_instance, requests_mock):
    """Validate that get_game_key parses XML correctly via an HTTP mocked response."""
    api_instance.session = requests.Session()  # provide a real Session so requests-mock hooks in
    mock_xml = """
    <fantasy_content>
        <game>
            <game_key>2025</game_key>
        </game>
    </fantasy_content>
    """.strip()
    requests_mock.get(
        "https://fantasysports.yahooapis.com/fantasy/v2/game/nhl",
        text=mock_xml,
        status_code=200,
    )

    game_key = api_instance.get_game_key()
    assert game_key == "2025"
    assert api_instance.year_id == "2025"
    assert len(requests_mock.request_history) == 1


def test_get_player_name_with_requests_mock_and_cache(api_instance, requests_mock):
    """Ensure player name fetched once over HTTP then served from cache."""
    api_instance.session = requests.Session()
    api_instance.year_id = "2025"  # prevent triggering game key call
    player_key = "2025.p.9999"
    player_xml = f"""
    <fantasy_content>
        <player>
            <name>
                <full>Test Player</full>
            </name>
        </player>
    </fantasy_content>
    """.strip()
    requests_mock.get(
        f"https://fantasysports.yahooapis.com/fantasy/v2/player/{player_key}",
        text=player_xml,
        status_code=200,
    )

    first = api_instance.get_player_name(player_key)
    second = api_instance.get_player_name(player_key)
    assert first == "Test Player" == second
    # Only one HTTP request should have been issued
    assert len(requests_mock.request_history) == 1
