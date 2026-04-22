#!/usr/bin/env python3
"""
Erstellt eine portable .exe mit PyInstaller.

Verwendung:
  pip install pyinstaller
  python build_exe.py

Die fertige .exe findet sich danach unter:
  dist/IT-Dokumentationsassistent.exe
"""

import subprocess
import sys
import os

def main():
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller wird installiert ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                          # alles in eine .exe
        "--windowed",                         # kein Konsolenfenster
        "--name", "IT-Dokumentationsassistent",
        "--clean",
        "screen_doc_recorder.py",
    ]

    print("Build wird gestartet ...")
    result = subprocess.run(cmd, cwd=os.path.dirname(os.path.abspath(__file__)))

    if result.returncode == 0:
        print("\nErfolgreich erstellt: dist/IT-Dokumentationsassistent.exe")
        print("Die .exe kann ohne Python-Installation ausgeführt werden.")
    else:
        print("\nBuild fehlgeschlagen. Bitte Fehlermeldung prüfen.")
        sys.exit(1)

if __name__ == "__main__":
    main()
