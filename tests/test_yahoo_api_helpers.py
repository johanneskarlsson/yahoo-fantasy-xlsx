import pytest


def test_extract_dict_value_basic(api_instance):
    data = {"#text": "VALUE"}
    assert api_instance._extract_dict_value(data) == "VALUE"


def test_extract_dict_value_nested_with_key(api_instance):
    data = {"outer": {"#text": "InnerVal"}}
    assert api_instance._extract_dict_value(data, "outer") == "InnerVal"


def test_extract_dict_value_fallback_primitive(api_instance):
    assert api_instance._extract_dict_value("plain") == "plain"
    assert api_instance._extract_dict_value(123) == "123"


@pytest.mark.parametrize(
    "input_data,expected",
    [
        (None, []),
        ({"a": 1}, [{"a": 1}]),
        ([1, 2, 3], [1, 2, 3]),
    ],
)
def test_ensure_list(api_instance, input_data, expected):
    assert api_instance._ensure_list(input_data) == expected


def test_get_stat_modifier_value(api_instance):
    settings = {
        "stat_modifiers": {
            "stats": {
                "stat": [
                    {"stat_id": "1", "value": "0.5"},
                    {"stat_id": "2", "value": "2"},
                ]
            }
        }
    }
    assert api_instance._get_stat_modifier_value(settings, "1") == "0.5"
    assert api_instance._get_stat_modifier_value(settings, "2") == "2"
    assert api_instance._get_stat_modifier_value(settings, "999") == ""


def test_get_player_name_caching(api_instance, monkeypatch, faker):
    # Arrange: create a synthetic player key & name
    player_key = f"{faker.random_number(digits=4)}.p.{faker.random_number(digits=6)}"
    full_name = faker.name()

    call_counter = {"count": 0}

    # Bypass internal call to get_game_key (avoids requiring game response structure)
    api_instance.year_id = "2025"

    def fake_request(url):
        call_counter["count"] += 1
        # Provide both possible structures to exercise fallback logic
        return {
            "fantasy_content": {
                "player": {
                    "name": {"full": full_name}
                }
            }
        }

    # Monkeypatch network layer
    monkeypatch.setattr(api_instance, "_make_api_request", fake_request)

    # Act: first call triggers network
    name1 = api_instance.get_player_name(player_key)
    # Second call should hit cache only
    name2 = api_instance.get_player_name(player_key)

    # Assert
    assert name1 == full_name == name2
    assert call_counter["count"] == 1, "Expected only one underlying API call due to caching"


def test_get_player_name_missing_key_returns_empty(api_instance):
    assert api_instance.get_player_name("") == ""


def test_get_player_name_handles_missing_structure(api_instance, monkeypatch):
    def fake_request(url):
        return {"fantasy_content": {}}  # No player / players keys

    monkeypatch.setattr(api_instance, "_make_api_request", fake_request)
    api_instance.year_id = "2025"  # Prevent get_player_name from calling get_game_key
    name = api_instance.get_player_name("999.p.111")
    assert name == "(unknown)" or name == ""  # Accept either fallback per implementation
