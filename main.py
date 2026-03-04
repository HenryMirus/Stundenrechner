"""Einstiegspunkt für den Stundenrechner (root-level)."""
import os
import sys

# src/ zum Suchpfad hinzufügen, damit die internen Imports in app.py funktionieren
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from src.app import StundenrechnerApp  # noqa: E402

if __name__ == "__main__":
    app = StundenrechnerApp()
    app.run()
