# XML-Daten in Word kennenlernen

## Dateien

- `Produktbeleg-Vorlage.dotx`: echte Word-Vorlage mit sechs Klartext-Inhaltssteuerelementen und XML-Datenbindungen.
- `Produktbeleg-Beispiel.docx`: ausgefülltes Beispieldokument derselben Struktur.
- `Beispieldaten.xml`: ein fiktiver Datensatz im Format des App-Exports.

Die Vorlage enthält bewusst Beispielwerte und einen eingebetteten benutzerdefinierten XML-Teil. Es besteht keine Verbindung zur SQLite-Datenbank und keine automatische Verbindung zur externen XML-Datei.

## Erster Lernschritt: die vorhandene Zuordnung ansehen

Diese Anleitung richtet sich an Microsoft Word als Desktop-App unter Windows. Word im Browser bietet diese XML-Zuordnungsfunktionen nicht vollständig. Menüs und Verfügbarkeit können je nach Word-Version abweichen.

1. Lade die Vorlage herunter und öffne sie in Word. Ein Doppelklick auf `.dotx` erstellt normalerweise ein neues Dokument auf Basis der Vorlage. Speichere dieses als `.docx`, damit die Vorlage unverändert bleibt.
2. Aktiviere **Datei → Optionen → Menüband anpassen → Entwicklertools**.
3. Wähle **Entwicklertools → XML-Zuordnungsbereich**.
4. Wähle im Dropdown den benutzerdefinierten XML-Teil mit dem Wurzelknoten `datensaetze`. Er enthält die eingebetteten Beispieldaten.
5. Klappe `datensaetze → datensatz` auf. Die ID ist ein Attribut, die übrigen Werte sind Elemente.
6. Klicke im Dokument auf ein Datenfeld. Über **Entwicklertools → Eigenschaften** siehst du den Titel und das Tag des Inhaltssteuerelements.

Die vorhandenen Bindungen verwenden diese XPath-Ausdrücke:

| Feld | XPath |
|---|---|
| ID | `/datensaetze/datensatz[1]/@id` |
| Vorname | `/datensaetze/datensatz[1]/vorname` |
| Nachname | `/datensaetze/datensatz[1]/nachname` |
| Produkt | `/datensaetze/datensatz[1]/produkt` |
| Preis | `/datensaetze/datensatz[1]/preis` |
| Datum | `/datensaetze/datensatz[1]/datum` |

`[1]` bedeutet: erster Datensatz. Der Preis bleibt zunächst eine Dezimalzahl mit Punkt, das Datum ein ISO-Datum. Eine formatierte Darstellung kann später ergänzt werden.

## Zweiter Lernschritt: deine exportierte XML-Datei zuordnen

1. Erstelle eine neue Dokumentkopie aus der Vorlage.
2. Im XML-Zuordnungsbereich wähle im Dropdown **Neuen Teil hinzufügen…** und öffne deine exportierte XML-Datei.
3. Wähle anschließend diesen neu hinzugefügten Teil. Es können jetzt zwei ähnlich benannte Teile vorhanden sein: Beispiel und Import. Kontrolliere die tatsächlichen Werte im Baum.
4. Markiere das bestehende Inhaltssteuerelement für den Vornamen im Dokument. Klicke im XML-Baum mit der rechten Maustaste auf `vorname` unter dem ersten `datensatz` und wähle **Dem ausgewählten Inhaltssteuerelement zuordnen**. Wiederhole dies für Nachname, Produkt, Preis und Datum sowie das Attribut `id`.
5. Prüfe, dass alle sechs Felder die Daten aus deinem ersten Datensatz anzeigen, und speichere das Dokument als `.docx`.

Falls Word nur **Inhaltssteuerelement einfügen** anbietet, ist das bestehende Steuerelement noch nicht ausgewählt. Wähle es über seinen Rahmen aus; alternativ ersetze es an derselben Stelle durch ein neues Klartext-Steuerelement aus dem XML-Baum.

Wichtig: **Neuen Teil hinzufügen** ersetzt die bisherigen Bindungen nicht automatisch. Die Felder müssen auf den neuen Teil zugeordnet werden. Der Import bettet eine Kopie der XML-Daten in das Dokument ein; spätere Änderungen an der externen Datei aktualisieren das Dokument nicht automatisch. Die eingebetteten Daten bleiben im Dokument enthalten, auch wenn nicht alle davon sichtbar sind. Gib deshalb keine Dokumente mit vertraulichen zusätzlichen Datensätzen weiter.

## Ein eigenes Dokument pro Datensatz

Word erzeugt durch XML-Zuordnung allein keine getrennten Dokumente. Die Vorlage zeigt zunächst genau einen Datensatz.

Für zwei Datensätze kannst du zum Lernen zwei Dokumentkopien erstellen und in jeder die sechs Steuerelemente auf den passenden `datensatz` im XML-Baum zuordnen. Achte auf die ID, damit du den richtigen Datensatz wählst.

Für beliebig viele Datensätze ist der nächste Schritt ein Generator:

1. Export-XML lesen.
2. Für jeden `datensatz` eine Kopie der Vorlage erstellen.
3. Den eingebetteten XML-Teil durch XML mit genau diesem einen Datensatz ersetzen und die angezeigten Feldwerte aktualisieren.
4. Als `Produktbeleg-<ID>.docx` speichern; Dateinamen müssen bereinigt und doppelte Namen abgefangen werden.

Damit können die bestehenden `[1]`-Bindungen erhalten bleiben, und jedes Dokument enthält nur seinen eigenen Datensatz. Der automatische Generator steht als separates Kommandozeilenwerkzeug bereit; die Webanwendung selbst exportiert weiterhin nur XML. Die Bedienung steht weiter unten.

## Prüfung

Die erzeugten Office-Dateien wurden als ZIP-Pakete und auf wohlgeformtes XML geprüft. Alle sechs Bindungen wurden gegen die Beispiel-XML geprüft. Eine visuelle Prüfung bzw. das Aktualisierungsverhalten in Microsoft Word muss auf einem Rechner mit Word erfolgen; Word ist in der Cloud-Umgebung nicht installiert.

## Vorlage erneut erstellen

Der optionale Generator `tools/create_word_template.py` erstellt Vorlage und Beispieldokument. Er benötigt `python-docx` (bei dieser Erstellung Version 1.2.0); die XML-Webanwendung braucht dieses Paket nicht. Aus dem Repository-Verzeichnis:

```sh
python -m pip install python-docx==1.2.0
python tools/create_word_template.py
python -m unittest discover -s tools -p 'test_*.py' -v
```

Bei der ersten Vorlagenversion waren Inhalts-Typ und Verknüpfung des benutzerdefinierten XML-Teils doppelt enthalten. Die korrigierte Version verwendet die vorhandenen Einträge und prüft diese auf Eindeutigkeit. Bitte alte Downloads durch die neue Version ersetzen.

## Automatisch ein Word-Dokument je Datensatz erzeugen

Der Generator ist jetzt unter `tools/generate_word_documents.py` verfügbar. Er benötigt nur Python 3.10 oder neuer, keine zusätzlichen Pakete und kein installiertes Word.

1. Lade auf GitHub die aktuelle Repository-Version über **Code → Download ZIP** herunter und entpacke sie. Verwende die enthaltene korrigierte Vorlage.
2. Kopiere deinen XML-Export in den entpackten Projektordner und benenne ihn beispielsweise `datensaetze.xml`.
3. Öffne in diesem Projektordner ein Terminal. Unter Windows kannst du im Explorer in die Adresszeile `powershell` eingeben und Enter drücken.
4. Führe aus:

```sh
python tools/generate_word_documents.py datensaetze.xml --output ausgabe
```

Falls bei dir Python über `py` statt `python` gestartet wird, ersetze das erste Wort durch `py`.

Bei drei Datensätzen entstehen drei Dateien im Ordner `ausgabe`, zum Beispiel `Produktbeleg-1001.docx`, `Produktbeleg-1002.docx` und `Produktbeleg-1003.docx`. Die Namen richten sich nach deinen IDs. Die Ausgabe im Terminal nennt die tatsächlich erzeugten Dateien.

Öffne jedes Dokument und vergleiche ID, Namen, Produkt, Preis und Datum mit dem entsprechenden XML-Datensatz. Unter **Entwicklertools → XML-Zuordnungsbereich** kannst du außerdem den eingebetteten XML-Teil untersuchen: Er enthält genau einen Datensatz. Alle Bindungen beziehen sich weiterhin auf `datensatz[1]`, weil jede Dokumentkopie nur ihren eigenen Datensatz enthält.

Vorhandene Ausgabedateien werden nicht überschrieben. Für einen weiteren Lauf wähle beispielsweise `--output ausgabe-neu`.

### Die Struktur in deiner Vorlage nachvollziehen

- Die linke Tabellenspalte enthält feste Beschriftungen.
- Die rechte enthält sechs Klartext-Inhaltssteuerelemente.
- Jedes Steuerelement besitzt einen Titel und ein Tag: `id`, `vorname`, `nachname`, `produkt`, `preis` oder `datum`.
- Die XML-Bindung verbindet das Steuerelement mit dem entsprechenden XPath aus der Tabelle oben. Die ID wird aus dem Attribut `@id` gelesen.

Du kannst Überschrift, Beschriftungen, Schriftarten und Tabellenlayout in Word ändern. Öffne dafür die `.dotx` über **Datei → Öffnen**, damit du die Vorlage selbst bearbeitest, und speichere sie. Behalte die sechs Steuerelemente und ihre Tags; entferne oder dupliziere sie nicht. Gib bei einer anderen Vorlage den Pfad ausdrücklich an:

```sh
python tools/generate_word_documents.py datensaetze.xml --template word/Meine-Vorlage.dotx --output ausgabe-neu
```

Der Generator ist auf die Struktur unserer Vorlage ausgelegt. Er ersetzt den eingebetteten XML-Teil und die sichtbaren Inhalte der sechs Steuerelemente. Eine beliebige andere Word-Vorlage oder zusätzliche XML-Teile unterstützt er nicht automatisch. Die Formatierung innerhalb der Datenfelder wird zunächst auf einfache Textausgabe zurückgesetzt; Tabellen- und Dokumentlayout bleiben bestehen.
