import gc

import pytest


@pytest.fixture(autouse=True)
def _collect_tk_objects_on_main_thread():
    """Destroyed windows' Tk variables must be collected on the main thread, not
    later inside a background worker (Tk would complain about the wrong thread)."""
    yield
    gc.collect()


@pytest.fixture(autouse=True)
def _decisions_in_a_temporary_folder(tmp_path_factory, monkeypatch):
    """Nothing is ever written to the real user profile while testing."""
    monkeypatch.setenv("BINAA_COMPARE_DIR", str(tmp_path_factory.mktemp("decisions")))
