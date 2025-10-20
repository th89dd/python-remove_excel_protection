![version](https://img.shields.io/badge/version-1.2.0-blue.svg)
![date](https://img.shields.io/badge/date-20.10.2025-green.svg)
![status](https://img.shields.io/badge/status-running-green.svg)
![license](https://img.shields.io/badge/license-MIT-lightgrey.svg)
![python](https://img.shields.io/badge/python-%E2%89%A53.10-blue.svg)

# 🧩 Remove Excel Protection

Python-Tool zum **Entfernen von Blattschutz und Arbeitsmappenschutz** in Excel-Dateien (`.xlsx`, `.xlsm`), wenn das Passwort nicht mehr bekannt ist.
Das Tool **entpackt** die Datei, **entfernt** Schutz-Tags aus den XML-Dateien und **packt sie anschließend wieder zusammen** – vollautomatisch.

---

## 🚀 Features

* 🔓 Entfernt **sheetProtection** und **workbookProtection** Tags
* 💻 **CLI-Tool** (`excel-unlock`) und **Python-API**

---

## 📙 Inhaltsverzeichnis

1. [Getting Started](#-getting-started)
2. [Installation](#-installation)
3. [Paket bauen](#-paket-bauen)
4. [Benutzung](#-benutzung)
5. [Verwendung im Python-Code](#-verwendung-im-python-code)
7. [Entwicklung](#-entwicklung)
8. [Lizenz](#-lizenz)
9. [Versionshistorie](#-versionshistorie)
10. [Kontakt / Mitwirkung](#-kontakt--mitwirkung)

---

## 🧽 Getting Started

Voraussetzungen:

* Python **≥ 3.10**
* `pip` (aktuell empfohlen)

Dieses Projekt kann als Paket installiert oder direkt aus dem Quellcode verwendet werden.

---

## ⚙️ Installation

### 🔹 Variante 1 – Installation über Wheel-Datei

Wenn bereits gebaut (siehe Abschnitt „Paket bauen“):

```bash
pip install dist/excel_protection_remover-1.2.0-py3-none-any.whl
```

### 🔹 Variante 2 – Installation direkt aus Git

```bash
pip install git+https://github.com/th89dd/python-remove_excel_protection.git
```

### 🔹 Variante 3 – Entwicklung (editable mode)

Für aktive Entwicklung, Änderungen greifen sofort:

```bash
pip install -e .
```

### 🔹 Variante 4 – Isolierte CLI-Installation mit `pipx` (empfohlen)

```bash
pipx install .
```

Nach der Installation steht das Kommando `excel-unlock` systemweit zur Verfügung (ohne venv).

Falls `pipx` noch nicht installiert ist:
```bash
pip install pipx
pipx ensurepath
```

#### Was macht pipx?:

pipx ist eine Erweiterung von pip, die speziell für Befehlszeilenprogramme (CLI-Tools) geschrieben in Python gedacht ist.
Es installiert jedes Tool in einer eigenen, isolierten virtuellen Umgebung und macht den Kommandozeilenbefehl systemweit verfügbar.

Wenn dein Projekt auf GitHub liegt (z. B. öffentlich oder privat mit Token), kannst du direkt so installieren:
pipx install git+https://github.com/th89dd/python-remove_excel_protection.git

oder über ein wheel, wenn du das Paket gebaut hast:
pipx install excel_protection_remover-1.2.0-py3.none-any.whl

Wenn du später ein Update willst:

pipx upgrade excel-protection-remover

---

## 🔨 Paket bauen

Dieses Projekt nutzt **PEP 621** / `pyproject.toml`.

1. Build-Tool installieren:

```bash
pip install build
```

2. Wheel- und Source-Paket erzeugen:

```bash
python -m build
```

Ergebnis (Beispiel):

```
dist/
├── excel_protection_remover-0.1-py3-none-any.whl
└── excel_protection_remover-0.1.tar.gz
```

3. Optional: lokal testen

```bash
pip install dist/excel_protection_remover-0.1-py3-none-any.whl
```

4. Weitere nützliche Befehle:
* Pakete listen: `pip list`
* Paket deinstallieren: `pip uninstall excel_protection_remover`
* Paket in Test-PyPI hochladen: `twine upload --repository testpypi dist/*`
* Paket in PyPI hochladen: `twine upload dist/*`
* Dokumentation mit Sphinx erstellen (optional):
  **TODO**

---

## 💻 Benutzung

### Nach der Installation:

```bash
excel-unlock "C:\Pfad\zu\geschuetzt.xlsx"
```

#### Optionen

| Parameter   | Beschreibung                                           | Standard                |
| ----------- | ------------------------------------------------------ | ----------------------- |
| `input`     | Pfad zur geschützten Excel-Datei                       | —                       |
| `output`    | Pfad für die Ausgabedatei                              | `<input>_unlocked.xlsx` |
| `--log`     | Log-Level: `DEBUG`, `INFO`, `WARNING`, `ERROR`         | `INFO`                  |
| `--logfile` | Optionaler Pfad für Logdatei (z. B. `logs/output.log`) | *keine Datei*           |

**Beispiel mit Datei-Logging:**

```bash
excel-unlock "test.xlsx" --log DEBUG --logfile logs/output.log
```

**Hilfe anzeigen:**

```bash
excel-unlock --help
```

---

### 🧩 Verwendung im Python-Code

```python
from excel_protection_remover import ExcelProtectionRemover, LoggerFactory

# Optional: Logging zentral konfigurieren
LoggerFactory.configure(level="INFO")  # oder "DEBUG", "WARNING", ...

remover = ExcelProtectionRemover("test.xlsx")
remover.remove_protection()

print(remover.status_message)
# -> ✅ Schutz entfernt (X Einträge) → test_unlocked.xlsx
```

---

### 🗾 Beispielausgabe

```
[2025-10-20 22:30:45] [INFO] excel_protection_remover.archive.ExcelArchive: Entpacke test.xlsx
[2025-10-20 22:30:46] [DEBUG] excel_protection_remover.cleaner.ProtectionCleaner: 3 Schutz-Tags entfernt
[2025-10-20 22:30:47] [INFO] excel_protection_remover.remover.ExcelProtectionRemover: ✅ Schutz entfernt (3 Einträge) → test_unlocked.xlsx
```

---

## 🛠️ Entwicklung

**TODO**: Hinweise zur Entwicklung, Tests, Code-Style, CI/CD, etc.

---

## 🪪 Lizenz

Dieses Projekt steht unter der [MIT License](LICENSE).

---

## 🕒 Versionshistorie

| Datum / Version     | Autor         | Bemerkung  |
|---------------------| ------------- |------------|
| 20.10.2025 / v1.2.0 | Tim Häberlein | Erstellung |
| – / –               | –             | –          |

---

## 📬 Kontakt / Mitwirkung

**Autor:** Tim Häberlein
**Lizenz:** MIT
**GitHub:** [https://github.com/](https://github.com/th89dd/python-remove_excel_protection)

Pull Requests, Bug Reports und Feature-Vorschläge sind willkommen! ✌️
