"""Activate a Python environment relative to python.exe."""
import os
from pathlib import Path
import sys


def activate():
    """Add environment variables."""
    python_exe = Path(sys.executable)
    python_env = python_exe.parent
    env = os.environ

    env['PATH'] = ('{python_env};'
                   '{python_env}\\Library\\mingw-w64\\bin;'
                   '{python_env}\\Library\\usr\\bin;'
                   '{python_env}\\Library\\bin;'
                   '{python_env}\\Scripts;'
                   '{python_env}\\bin;'
                   '{path}').format(path=env['PATH'],
                                    python_env=python_env
                                    )
    env['VIRTUAL_ENV'] = str(python_env)
