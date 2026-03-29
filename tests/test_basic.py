import opensheet_core


def test_version():
    # Verify version is a valid semver string
    parts = opensheet_core.__version__.split(".")
    assert len(parts) == 3
    assert all(p.isdigit() for p in parts)


def test_import():
    """Verify the native module loads correctly."""
    from opensheet_core._native import version
    assert isinstance(version(), str)
