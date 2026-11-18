"""Keep writable application data and configurable caches with this installation."""
import os
import tempfile

from .paths import state_dir


def configure_runtime():
    home = state_dir()
    home.mkdir(parents=True, exist_ok=True)
    temporary = home / ".tmp"
    temporary.mkdir(exist_ok=True)
    os.environ["TEMP"] = str(temporary)
    os.environ["TMP"] = str(temporary)
    tempfile.tempdir = str(temporary)
    # pywin32's generated wrappers can otherwise use the roaming/temp profile.
    try:
        import win32com
        cache = home / ".cache" / "gen_py"
        cache.mkdir(parents=True, exist_ok=True)
        win32com.__gen_path__ = str(cache)
    except ImportError:
        pass
