#!/usr/bin/env python3
"""
Entry point for the schedule-to-PDF converter.

Usage:
    python run.py

Then open http://localhost:5000 in your browser.
"""

from app.web import create_app

if __name__ == "__main__":
    app = create_app()
    print("=" * 50)
    print("  Расписание → PDF")
    print("  Откройте в браузере: http://localhost:5000")
    print("=" * 50)
    app.run(host="0.0.0.0", port=5000, debug=True)
