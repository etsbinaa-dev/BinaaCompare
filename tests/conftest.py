import gc

import pytest


@pytest.fixture(autouse=True)
def _collect_tk_objects_on_main_thread():
    """Destroyed windows' Tk variables must be collected on the main thread, not
    later inside a background worker (Tk would complain about the wrong thread)."""
    yield
    gc.collect()
