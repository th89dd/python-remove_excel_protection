# -*- coding: utf-8 -*-
"""
:author: Tim Häberlein
:version: 1.0
:date: 20.10.2025
:organisation: TU Dresden, FZM
"""

# -------- start import block ---------
import os
import zipfile
import shutil
import tempfile
from excel_protection_remover.base import Loggable
# -------- /import block ---------

class ExcelArchive(Loggable):
    """Verwaltet das Entpacken und Wiederpacken einer Excel-Datei (.xlsx oder .xlsm).

    Diese Klasse ist zuständig für:
      * das sichere Entpacken einer Excel-Datei in ein temporäres Verzeichnis,
      * das spätere Wiederpacken nach Modifikation,
      * und das Aufräumen des temporären Ordners.

    Sie erbt von :class:`Loggable` und nutzt somit automatisch Logging.
    """

    def __init__(self, input_path: str, output_path: str = None):
        """Initialisiert die ExcelArchive-Instanz.

        Args:
            input_path: Pfad zur zu bearbeitenden Excel-Datei.
            output_path: Optionaler Pfad für die entpackte und neu gepackte Ausgabe-Datei.
        """
        super().__init__()  # Initialisiert den Logger
        if not os.path.isfile(input_path):
            raise FileNotFoundError(f"Datei nicht gefunden: {input_path}")

        self.__input_path = input_path
        self.__output_path = output_path or self.__generate_output_path(input_path)
        self.__temp_dir = tempfile.mkdtemp()

        self._logger.debug(f"Temporäres Verzeichnis erstellt: {self.__temp_dir}")

    @property
    def input_path(self) -> str:
        """Pfad zur Eingabedatei."""
        return self.__input_path

    @property
    def output_path(self) -> str:
        """Pfad zur Ausgabedatei."""
        return self.__output_path

    @property
    def temp_dir(self) -> str:
        """Pfad zum temporären Arbeitsverzeichnis."""
        return self.__temp_dir

    @staticmethod
    def __generate_output_path(path: str) -> str:
        """Erzeugt den Standardnamen der Ausgabe-Datei.

        Args:
            path: Ursprünglicher Dateipfad.

        Returns:
            str: Neuer Pfad mit Suffix "_unlocked".
        """
        base, ext = os.path.splitext(path)
        return f"{base}_unlocked{ext}"

    def extract(self):
        """Entpackt die Excel-Datei in das temporäre Arbeitsverzeichnis."""
        self._logger.info(f"Entpacke: {self.__input_path}")
        with zipfile.ZipFile(self.__input_path, "r") as zip_ref:
            zip_ref.extractall(self.__temp_dir)
        self._logger.debug(f"Datei erfolgreich entpackt nach temp_dir: {self.__temp_dir}.")

    def repack(self):
        """Packt die bearbeiteten Dateien wieder zu einer neuen Excel-Datei."""
        self._logger.info(f"Packe wieder zusammen → {self.__output_path}")
        with zipfile.ZipFile(self.__output_path, "w", zipfile.ZIP_DEFLATED) as zip_out:
            for root, _, files in os.walk(self.__temp_dir):
                for file in files:
                    abs_path = os.path.join(root, file)
                    rel_path = os.path.relpath(abs_path, self.__temp_dir)
                    zip_out.write(abs_path, rel_path)
        self._logger.debug("Repack abgeschlossen.")

    def cleanup(self):
        """Löscht das temporäre Arbeitsverzeichnis."""
        shutil.rmtree(self.__temp_dir, ignore_errors=True)
        self._logger.debug("Temporäres Verzeichnis gelöscht.")



# ausführung nur bei direktem Script-Aufruf
if __name__ == '__main__':
    print('run')
    archive = ExcelArchive("../working_dir/AZ_Nachweis_2025_haeberlein.xlsx")
    archive._logger.setLevel('DEBUG')
    archive._logger.debug(archive.temp_dir)
    archive.extract()
    archive.repack()
    archive.cleanup()
