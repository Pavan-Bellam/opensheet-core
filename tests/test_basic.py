import opensheet_core


def test_version():
    assert opensheet_core.__version__ == "0.1.0"


def test_import():
    """Verify the native module loads correctly."""
    from opensheet_core._native import version
    assert isinstance(version(), str)
