import pytest
from faker import Faker


@pytest.fixture(scope="session")
def faker():  # override name to avoid needing pytest-faker plugin
    return Faker()


class _DummySession:
    """Minimal dummy session to prevent accidental real HTTP during tests."""
    def get(self, url, *_, **__):  # pragma: no cover - safeguard
        raise RuntimeError(f"Network call blocked in test: {url}")


@pytest.fixture()
def api_instance():
    """Provide a fresh YahooFantasyAPI instance with a dummy session (no real HTTP)."""
    from yahoo_api import YahooFantasyAPI
    api = YahooFantasyAPI()
    api.session = _DummySession()
    return api


def assert_formula_tokens(formula: str, *tokens: str):
    """Assert each token appears (order-independent) in a formula string; raise AssertionError otherwise.

    Normalizes surrounding whitespace and uppercase for robustness.
    """
    if formula is None:
        raise AssertionError("Formula is None")
    norm = formula.replace(' ', '').upper()
    for t in tokens:
        if t.replace(' ', '').upper() not in norm:
            raise AssertionError(f"Token '{t}' not in formula '{formula}'")


@pytest.fixture()
def formula_assert():
    return assert_formula_tokens
