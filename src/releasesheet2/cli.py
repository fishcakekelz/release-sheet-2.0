"""Command-line entry for Release sheet 2.0."""

from __future__ import annotations

import argparse

from releasesheet2 import __version__


def main() -> None:
    parser = argparse.ArgumentParser(prog="releasesheet2", description="Release sheet 2.0")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.parse_args()
    print("releasesheet2: no command yet. Use --version for the package version.")


if __name__ == "__main__":
    main()
