# -*- coding: utf-8 -*-
"""
:author: Tim Häberlein
:version: 1.0
:date: 20.10.2025
:organisation: TU Dresden, FZM
"""

# -------- start import block ---------
from excel_protection_remover.logger import LoggerFactory, logging

# -------- /import block ---------

class Loggable:
    """Basisklasse für alle Komponenten mit Logging-Unterstützung.

    Jede von Loggable abgeleitete Klasse erhält automatisch
    einen konfigurierten Logger mit hierarchischem Namen im Format:

        <modulname>.<klassenname>

    Beispiel:
        >>> class MyComponent(Loggable):
        ...     def __init__(self):
        ...         super().__init__()
        ...         self._logger.info("Komponente gestartet")

    Nach der Initialisierung steht der Logger über die Property
    :pyattr:`_logger` zur Verfügung.
    """

    def __init__(self):
        """Erzeugt automatisch einen hierarchischen Logger für die abgeleitete Klasse."""
        full_name = f"{self.__module__}.{self.__class__.__name__}"
        self.__logger = LoggerFactory.get_logger(full_name)
        self.__logger.debug(f"{self.__class__.__name__} initialisiert.")

    @property
    def _logger(self) -> logging.Logger:
        """Geschützter Zugriff auf den Klassen-Logger.

        Returns:
            logging.Logger: Der Logger der aktuellen Klasse.
        """
        return self.__logger




# ausführung nur bei direktem Script-Aufruf
if __name__ == '__main__':
    print('run')
    pass
