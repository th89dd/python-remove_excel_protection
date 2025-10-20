# -*- coding: utf-8 -*-
"""
:author: Tim Häberlein
:version: 1.0
:date: 20.10.2025
:organisation: TU Dresden, FZM
"""

# -------- start import block ---------

from .logger import LoggerFactory
from .base import Loggable
from .archive import ExcelArchive
from .cleaner import ProtectionCleaner
from .remover import ExcelProtectionRemover

# -------- /import block ---------

"""Excel Protection Remover Package.

Dieses Package bietet Werkzeuge zum Entfernen von Blattschutz
und Arbeitsmappenschutz aus Excel-Dateien (.xlsx, .xlsm),
indem es die Datei entpackt, XML-Tags entfernt und sie
anschließend wieder zusammenpackt.

Hauptklassen:
    * :class:`ExcelProtectionRemover` – Fassade für den gesamten Prozess.
    * :class:`ExcelArchive` – Handhabt Entpacken und Wiederpacken.
    * :class:`ProtectionCleaner` – Entfernt Schutz-Tags aus XML-Dateien.
    * :class:`LoggerFactory` – Zentrale Logging-Verwaltung.
    * :class:`Loggable` – Basisklasse mit automatischem Logger.

Beispiel:
    >>> from excel_protection_remover import ExcelProtectionRemover, LoggerFactory
    >>> LoggerFactory.configure(level="DEBUG")
    >>> remover = ExcelProtectionRemover("test.xlsx")
    >>> remover.remove_protection()
    >>> print(remover.status_message)
"""


__all__ = [
    "LoggerFactory",
    "Loggable",
    "ExcelArchive",
    "ProtectionCleaner",
    "ExcelProtectionRemover",
]

__version__ = "1.2.0"
__author__ = "Tim Häberlein"
__license__ = "MIT"

