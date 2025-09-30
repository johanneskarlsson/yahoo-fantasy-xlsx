import pytest


def build_fake_player(player_key: str, full_name: str, team_abbr: str, position: str):
    return {
        "player_key": {"#text": player_key},
        "name": {"full": {"#text": full_name}},
        "editorial_team_abbr": {"#text": team_abbr},
        "display_position": {"#text": position},
        "draft_analysis": {
            "average_pick": {"#text": "45.6"},
            "average_round": {"#text": "4.2"},
            "percent_drafted": {"#text": "98"},
            "preseason_average_pick": {"#text": "50.1"},
            "preseason_percent_drafted": {"#text": "97"},
        },
        "projected_auction_value": {"#text": "25"},
        "average_auction_cost": {"#text": "23"},
        "player_ranks": {
            "player_rank": [
                {"rank_type": "S", "rank_season": "2025", "rank_value": {"#text": "30"}},
                {"rank_type": "S", "rank_season": "2025", "rank_position": "C", "rank_value": {"#text": "8"}},
            ]
        },
    }


def test_extract_player_draft_info_full(api_instance, faker):
    player_key = f"{faker.random_number(digits=3)}.p.{faker.random_number(digits=5)}"
    name = faker.name()
    team = faker.lexify(text="???").upper()
    position = "C"
    fake_player = build_fake_player(player_key, name, team, position)

    row = api_instance._extract_player_draft_info(fake_player)
    assert row[0] == player_key
    assert row[1] == name
    assert row[2] == team
    assert row[3] == position
    # average_pick, average_round, percent_drafted
    assert row[4] == "45.6"
    assert row[5] == "4.2"
    assert row[6] == "98"
    # projected_auction_value, average_auction_cost
    assert row[7] == "25"
    assert row[8] == "23"
    # season_rank, position_rank
    assert row[9] == "30"
    assert row[10] == "8"


def test_extract_player_draft_info_handles_missing(api_instance):
    # Missing many fields should not raise and should produce sensible fallbacks
    minimal_player = {"player_key": {"#text": "1.p.1"}, "name": {"full": {"#text": "Test Player"}}}
    row = api_instance._extract_player_draft_info(minimal_player)
    assert row[0] == "1.p.1"
    assert row[1] == "Test Player"
    # The rest may be empty strings
    assert len(row) == 13
