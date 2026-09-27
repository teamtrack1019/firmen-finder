# Firmen-Finder

Sucht kleine Betriebe von Würzburg über Aschaffenburg und Hanau bis Frankfurt am Main, öffnet die öffentliche Webseite und speichert Firma, E-Mail und Geschäftsführer.

Quelle der Betriebe ist OpenStreetMap. Webseiten werden mit Pause gelesen, `robots.txt` wird beachtet. Ergebnisdateien liegen in diesem Ordner: `leads.xlsx` und `leads.csv`.

Vor dem Anschreiben gilt das deutsche Recht (UWG, DSGVO). Das Skript verschickt keine E-Mails.

## Nächtlicher Lauf

Ein Lauf ohne Aufsicht setzt die vorhandene `leads.xlsx` fort und liest höchstens 40 neue Webseiten:

```powershell
python firmen_finder.py --nacht
```

Bereits gespeicherte Firmen werden übersprungen. Die Tabelle wird danach auf dem Branch `main` erwartet, damit sie am Morgen auf dem Handy über GitHub geöffnet werden kann.

## Start auf dem eigenen Rechner

```powershell
.venv\Scripts\python firmen_finder.py
```

Die ganze Strecke kann länger dauern. Ein unterbrochener Lauf setzt beim nächsten Start fort. Mit `--neu` beginnt die Datei von vorn.

```powershell
.venv\Scripts\python firmen_finder.py --orte "Aschaffenburg,Hanau" --limit 30
.venv\Scripts\python firmen_finder.py --self-test
```

In Excel filtert die Spalte **Status** auf `vollständig`, wenn E-Mail und Geschäftsführer gefunden wurden.

Das Repository kann öffentlich sein. Sobald `leads.xlsx` auf GitHub liegt, sind die darin stehenden Namen und E-Mail-Adressen für jeden sichtbar.
