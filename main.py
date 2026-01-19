"""Main entry point for the data quality control tool.

This is a convenience wrapper around the orsa_analysis package.
You can run this file directly:

    python main.py

Or use the installed command:

    orsa-qc

Or run as a module:

    python -m orsa_analysis
"""

from orsa_analysis.__main__ import main

if __name__ == "__main__":
    main()
