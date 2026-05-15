"""Compatibility wrapper for the ISO assessment CLI.

The active implementation lives in `assessment_runner_core.py`.
This wrapper exists so older commands and documentation keep working.
"""

from assessment_runner_core import main


if __name__ == "__main__":
    main()
