# -*- coding: utf-8 -*-
"""
:author: Tim Häberlein
:version: 1.0
:date: 20.10.2025
:organisation: TU Dresden, FZM
"""

# -------- start import block ---------
import argparse
from excel_protection_remover.logger import LoggerFactory
from excel_protection_remover.remover import ExcelProtectionRemover

# -------- /import block ---------
"""Kommandozeileninterface für excel_protection_remover.

Dieses Modul ermöglicht es, den Schutz von Excel-Dateien direkt über
die Konsole aufzuheben. Es nutzt die :class:`ExcelProtectionRemover`
und bietet Parameter zur Steuerung von Logging und Ausgabedateien.

Beispiel:
    Entfernen des Blattschutzes mit Standard-Logging:

        $ excel-unlock meine_datei.xlsx

    Mit Debug-Level und Logdatei:

        $ excel-unlock meine_datei.xlsx --log DEBUG --logfile logs/output.log
"""



def main():
    """Startpunkt der Kommandozeilenanwendung.

    Diese Funktion:
      * parst die Kommandozeilenargumente,
      * konfiguriert das Logging-System über :class:`LoggerFactory`,
      * führt den Entsperrprozess mit :class:`ExcelProtectionRemover` aus.

    Returns:
        None
    """
    parser = argparse.ArgumentParser(
        description="Entfernt Blattschutz und Arbeitsmappenschutz aus Excel-Dateien (.xlsx / .xlsm)."
    )
    parser.add_argument(
        "input",
        help="Pfad zur geschützten Excel-Datei",
    )
    parser.add_argument(
        "output",
        nargs="?",
        help="Pfad für die entsperrte Ausgabe-Datei (optional)",
    )
    parser.add_argument(
        "--log",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="INFO",
        help="Log-Level (Standard: INFO)",
    )
    parser.add_argument(
        "--logfile",
        help="Optionaler Pfad für Logdatei (z. B. logs/output.log)",
    )

    args = parser.parse_args()

    # Logging konfigurieren
    LoggerFactory.configure(level=args.log, logfile=args.logfile)

    remover = ExcelProtectionRemover(args.input, args.output)
    remover.remove_protection()


if __name__ == "__main__":
    main()