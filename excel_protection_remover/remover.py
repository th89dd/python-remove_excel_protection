# -*- coding: utf-8 -*-
"""
:author: Tim Häberlein
:version: 1.0
:date: 20.10.2025
:organisation: TU Dresden, FZM
"""

# -------- start import block ---------
from excel_protection_remover.archive import ExcelArchive
from excel_protection_remover.cleaner import ProtectionCleaner
from excel_protection_remover.base import Loggable
# -------- /import block ---------


class ExcelProtectionRemover(Loggable):
    """Koordiniert den gesamten Prozess zur Entfernung von Excel-Blattschutz.

    Schritte:
      1. Entpacken der Datei mit :class:`ExcelArchive`.
      2. Entfernen der Schutz-Tags mit :class:`ProtectionCleaner`.
      3. Wiederpacken der Datei.
      4. Aufräumen des temporären Arbeitsverzeichnisses.

    Erbt von :class:`Loggable` für einheitliches Logging.

    Attributes:
        status_message (str): Letzte Statusmeldung des Prozesses.

    Args:
        input_path (str): Pfad zur zu bearbeitenden Excel-Datei.
        output_path (str, optional): Pfad für die bearbeitete Ausgabe-Datei.
    """

    def __init__(self, input_path: str, output_path: str = None):
        """Initialisiert den ExcelProtectionRemover.

        Args:
            input_path: Pfad zur zu bearbeitenden Excel-Datei.
            output_path: Optionaler Pfad für die bearbeitete Ausgabe-Datei.
        """
        super().__init__()
        self.__archive = ExcelArchive(input_path, output_path)
        self.__cleaner = ProtectionCleaner(self.__archive.temp_dir)
        self.__status_message = ""

    @property
    def status_message(self) -> str:
        """Letzte Statusmeldung des Prozesses."""
        return self.__status_message

    def remove_protection(self):
        """Führt den gesamten Entsperrprozess aus.

        Der Ablauf:
            1. Entpackt die Datei.
            2. Entfernt Schutz-Tags.
            3. Packt die Datei wieder zusammen.
            4. Löscht das temporäre Verzeichnis.

        Exceptions:
            Alle Ausnahmen werden geloggt und im Status gespeichert.
        """
        try:
            self.__archive.extract()
            removed = self.__cleaner.clean()
            self.__archive.repack()
            self.__status_message = (
                f"✅ Schutz entfernt (insgesamt {removed} Einträge) → output-file: {self.__archive.output_path}"
            )
            self._logger.info(self.__status_message)
        except Exception as e:
            self.__status_message = f"❌ Fehler: {e}"
            self._logger.error(self.__status_message)
        finally:
            self.__archive.cleanup()


# ausführung nur bei direktem Script-Aufruf
if __name__ == '__main__':
    input_path = "working_dir/AZ_Nachweis_2025_haeberlein.xlsx"
    remover = ExcelProtectionRemover(input_path)
    pass
