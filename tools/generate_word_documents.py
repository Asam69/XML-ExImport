"""Ein Word-Dokument pro XML-Datensatz, ohne zusätzliche Python-Pakete."""
import argparse
from copy import deepcopy
from pathlib import Path
import re
import tempfile
import xml.etree.ElementTree as ET
from xml.dom import minidom
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parents[1]
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
FIELDS = ('vorname', 'nachname', 'produkt', 'preis', 'datum')


def generate(xml_path, template_path, output_dir):
    source = ET.parse(xml_path).getroot()
    if source.tag != 'datensaetze':
        raise ValueError('Erwartet wird der XML-Wurzelknoten datensaetze.')
    rows = source.findall('datensatz')
    if not rows:
        raise ValueError('Die XML-Datei enthält keine Datensätze.')
    with ZipFile(template_path) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    # DOM preserves namespace declarations also referenced inside attribute values.
    # ElementTree drops these, leaving mc:Ignorable prefixes undefined for Word.
    base_doc = minidom.parseString(files['word/document.xml'])
    types = ET.fromstring(files['[Content_Types].xml'])
    for node in types:
        if node.get('PartName') == '/word/document.xml':
            node.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')
    files['[Content_Types].xml'] = ET.tostring(types, encoding='utf-8', xml_declaration=True)
    controls = base_doc.getElementsByTagNameNS(W[1:-1], 'sdt')
    tags = [control.getElementsByTagNameNS(W[1:-1], 'tag')[0].getAttributeNS(W[1:-1], 'val') for control in controls]
    if sorted(tags) != sorted(('id',)+FIELDS):
        raise ValueError('Die Vorlage muss genau die sechs vorgesehenen Inhaltssteuerelemente enthalten.')
    planned = []
    used = set()
    for index, row in enumerate(rows, 1):
        values = {'id': row.get('id', ''), **{key: row.findtext(key, '') for key in FIELDS}}
        if any(not value.strip() for value in values.values()):
            raise ValueError(f'Datensatz {index}: ID oder Pflichtfeld fehlt.')
        safe_id = re.sub(r'[^A-Za-z0-9_-]', '_', values['id'])[:80]
        filename = f'Produktbeleg-{safe_id}.docx'
        if filename.casefold() in used:
            raise ValueError(f'Dateiname für ID {values["id"]} ist nicht eindeutig.')
        used.add(filename.casefold())
        target = Path(output_dir) / filename
        if target.exists():
            raise ValueError(f'{target} existiert bereits. Bitte einen neuen Ausgabeordner wählen.')
        planned.append((row, values, target))
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    # Erst alle Pakete vorbereiten, dann veröffentlichen. Vorhandene Dateien bleiben erhalten.
    with tempfile.TemporaryDirectory(dir=output_dir) as staging:
        for row, values, target in planned:
            doc = base_doc.cloneNode(deep=True)
            for control in doc.getElementsByTagNameNS(W[1:-1], 'sdt'):
                key = control.getElementsByTagNameNS(W[1:-1], 'tag')[0].getAttributeNS(W[1:-1], 'val')
                content = control.getElementsByTagNameNS(W[1:-1], 'sdtContent')[0]
                while content.firstChild:
                    content.removeChild(content.firstChild)
                run = doc.createElementNS(W[1:-1], 'w:r')
                text = doc.createElementNS(W[1:-1], 'w:t')
                text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve')
                text.appendChild(doc.createTextNode(values[key]))
                run.appendChild(text)
                content.appendChild(run)
            single = ET.Element('datensaetze', source.attrib)
            single.append(deepcopy(row))
            replacements = {
                'word/document.xml': doc.toxml(encoding='utf-8'),
                'customXml/item1.xml': ET.tostring(single, encoding='utf-8', xml_declaration=True),
            }
            with ZipFile(Path(staging)/target.name, 'w', ZIP_DEFLATED) as archive:
                for name, data in files.items():
                    archive.writestr(name, replacements.get(name, data))
        for _, _, target in planned:
            # Exklusives Anlegen verhindert auch Überschreiben bei parallelen Aufrufen.
            with target.open('xb') as destination:
                destination.write((Path(staging)/target.name).read_bytes())
    return [target for _, _, target in planned]


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('xml', type=Path, help='XML-Export aus der Lernanwendung')
    parser.add_argument('--template', type=Path, default=ROOT/'word'/'Produktbeleg-Vorlage.dotx')
    parser.add_argument('--output', type=Path, default=ROOT/'ausgabe')
    args = parser.parse_args()
    try:
        paths = generate(args.xml, args.template, args.output)
    except (ValueError, OSError, ET.ParseError) as error:
        parser.exit(1, f'Fehler: {error}\n')
    for path in paths:
        print(f'Erstellt: {path}')
    print(f'{len(paths)} Word-Dokument(e) erstellt.')
