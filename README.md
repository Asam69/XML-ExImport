# XML-Lernwerkstatt

Eine kleine Lernanwendung für **Vorname, Nachname, ID, Produkt, Preis und Datum**.
Python 3.10 oder neuer genügt; keine zusätzlichen Pakete oder Datenbankserver erforderlich.

## Starten

Im Repository-Verzeichnis:

```sh
python3 app.py
```

Öffne lokal im Browser `http://127.0.0.1:8000`. Beenden mit Strg+C.
Die Anwendung ist für lokale Lernzwecke vorgesehen, ohne Benutzeranmeldung.
Für eine extern erreichbare Installation wären Authentifizierung und weiterer Schutz nötig.

1. Einen Datensatz ausfüllen. Beispiel: ID `1001`, Vorname `Anna`, Nachname `Müller`, Produkt `Notizbuch`, Preis `12,90`, Datum `2026-10-11`.
2. **In Datenbank speichern** wählen. Die ID muss eindeutig sein.
3. **XML herunterladen** wählen und die Datei in einem Texteditor betrachten.

Die Daten bleiben nach einem Neustart in `data/records.sqlite3` erhalten. Diese lokale Datenbank wird nicht in Git aufgenommen. Zum Sichern bei gestoppter Anwendung diese Datei kopieren.

## Wie die Verarbeitung funktioniert

`save()` prüft die Eingaben und schreibt sie mit parametrisiertem SQL in SQLite.
Preise werden als ganze Cent-Beträge gespeichert, damit keine Rundungsfehler durch Gleitkommazahlen entstehen. Das Datum wird als ISO-Datum gespeichert.

`records()` liest die Daten. `export_xml()` erzeugt daraus eine XML-Struktur:

```xml
<?xml version='1.0' encoding='utf-8'?>
<datensaetze version="1">
  <datensatz id="1001">
    <vorname>Anna</vorname>
    <nachname>Müller</nachname>
    <produkt>Notizbuch</produkt>
    <preis waehrung="EUR">12.90</preis>
    <datum>2026-10-11</datum>
  </datensatz>
</datensaetze>
```

## Später: XML in Word verwenden

XML definiert Struktur und Inhalte, aber noch kein Word-Layout. Diese erste Version exportiert nur XML. Für den nächsten Schritt kann ein Generator das XML lesen und jeden Datensatz auf eine Tabellenzeile in einer Word-Vorlage abbilden. Alternativ bietet Word Inhaltssteuerelemente mit XML-Datenbindung. Einfaches Öffnen einer beliebigen XML-Datei in Word erzeugt noch kein gestaltetes Dokument.

Ein später erzeugtes Word-Dokument wäre eine Momentaufnahme der Daten; automatische Aktualisierung braucht eine zusätzliche Anbindung.

## Tests

```sh
python3 -m unittest -v
```

Die Tests verwenden eine temporäre Datenbank und prüfen Speicherung, Validierung, XML-Sonderzeichen, alle exportierten Felder und leere Exporte.

## Word-Vorlage und XML-Zuordnung

Im Ordner `word/` liegen eine Word-Vorlage (`.dotx`), ein Beispieldokument und fiktive Beispiel-XML-Daten. Die [Anleitung](word/Anleitung.md) erklärt die XML-Zuordnung in Word Desktop unter Windows. Die Anwendung selbst exportiert weiterhin nur XML.

Zum Herunterladen auf GitHub **Code → Download ZIP** wählen und das Archiv entpacken. Alternativ enthält `Word-Lernpaket.zip` nur die Word-Lerndateien.
