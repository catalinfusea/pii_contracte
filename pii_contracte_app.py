"""Punctul de intrare pentru .exe (PyInstaller) și pentru rularea din sursă.

    pii_contracte.exe                 interfața
    pii_contracte.exe f1.docx f2.pdf  interfața, cu fișierele deja adăugate
    pii_contracte.exe --selftest      autotest pe date sintetice (raport în %LOCALAPPDATA%\\pii_contracte)
"""
from pii_gui.app import main

if __name__ == "__main__":
    main()
