# -*- coding: utf-8 -*-
"""
:author: Tim Häberlein
:version: 1.0
:date: 20.10.2025
:organisation: TU Dresden, FZM
"""

# -------- start import block ---------
import os
import re
from excel_protection_remover.base import Loggable, logging
# -------- /import block ---------



class ProtectionCleaner(Loggable):
    """Entfernt Schutz-Tags aus den XML-Dateien einer Excel-Arbeitsmappe.

    Aufgaben:
      * Entfernt `<sheetProtection .../>`- und `<workbookProtection .../>`-Tags.
      * Zählt, wie viele Einträge entfernt wurden.
      * Arbeitet direkt auf den entpackten XML-Dateien im Arbeitsverzeichnis.

    Erbt von :class:`Loggable` für Logging-Unterstützung.
    """

    __SHEET_PATTERN = re.compile(r"<sheetProtection[^>]*?/?>")
    __WORKBOOK_PATTERN = re.compile(r"<workbookProtection[^>]*?/?>")

    def __init__(self, base_dir: str):
        """Initialisiert den ProtectionCleaner.

        Args:
            base_dir: Pfad zum entpackten Arbeitsverzeichnis (z. B. von ExcelArchive).
        """
        super().__init__()
        self.__base_dir = base_dir
        self.__removed_count = 0

    @property
    def base_dir(self) -> str:
        """Pfad zum Arbeitsverzeichnis."""
        return self.__base_dir

    @property
    def removed_count(self) -> int:
        """Anzahl der insgesamt entfernten Schutz-Tags."""
        return self.__removed_count

    def __collect_target_files(self):
        """Sammelt alle relevanten XML-Dateien.

        Returns:
            list[str]: Pfade zu den zu bereinigenden XML-Dateien.
        """
        targets = [os.path.join(self.__base_dir, "xl", "workbook.xml")]
        ws_dir = os.path.join(self.__base_dir, "xl", "worksheets")
        if os.path.isdir(ws_dir):
            for file in os.listdir(ws_dir):
                if file.endswith(".xml"):
                    targets.append(os.path.join(ws_dir, file))
        self._logger.debug(f"{len(targets)} XML-Dateien gefunden.")
        return targets

    def __clean_file(self, xml_file: str) -> int:
        """Bereinigt eine einzelne XML-Datei.

        Args:
            xml_file: Pfad zur XML-Datei.

        Returns:
            int: Anzahl der entfernten Schutz-Tags.
        """
        with open(xml_file, "r", encoding="utf-8") as f:
            content = f.read()

        new_content, c1 = self.__SHEET_PATTERN.subn("", content)
        new_content, c2 = self.__WORKBOOK_PATTERN.subn("", new_content)

        if c1 or c2:
            with open(xml_file, "w", encoding="utf-8") as f:
                f.write(new_content)
            self._logger.debug(f"{xml_file}: {c1 + c2} Schutz-Tags entfernt.")
        return c1 + c2

    def clean(self) -> int:
        """Durchsucht alle relevanten XML-Dateien und entfernt Schutz-Tags.

        Returns:
            int: Gesamtzahl der entfernten Schutz-Einträge.
        """
        total = 0
        xml_files = self.__collect_target_files()
        if self._logger.level == logging.DEBUG:
            self.xml_files = xml_files  # für Debug-Zwecke
        for xml_file in xml_files:
            if os.path.exists(xml_file):
                total += self.__clean_file(xml_file)
        self.__removed_count = total
        self._logger.info(f"Insgesamt {total} Schutz-Tags entfernt.")
        return total


# ausführung nur bei direktem Script-Aufruf
if __name__ == '__main__':
    cleaner = ProtectionCleaner("C:/Users/s7285521/AppData/Local/Temp/tmp2spyxwd7")
    cleaner._logger.setLevel('DEBUG')
    cleaner.clean()
    pass
