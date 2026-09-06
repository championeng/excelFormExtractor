"""Compatibility entry point for build frontends that execute setup.py."""

from pathlib import Path

exec(Path(__file__).with_name("setup_ci.py").read_text())
