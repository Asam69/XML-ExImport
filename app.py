"""Lernanwendung: Formular → SQLite → XML (nur Python-Standardbibliothek)."""
from datetime import date
from decimal import Decimal, InvalidOperation
from html import escape
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs
import argparse
import sqlite3
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent
DB = ROOT / 'data' / 'records.sqlite3'
FIELDS = ('id', 'vorname', 'nachname', 'produkt', 'preis', 'datum')
LABELS = ('ID', 'Vorname', 'Nachname', 'Produkt', 'Preis (€)', 'Datum')


def connect():
    DB.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    con.execute('''CREATE TABLE IF NOT EXISTS records (
        id TEXT PRIMARY KEY, vorname TEXT NOT NULL, nachname TEXT NOT NULL,
        produkt TEXT NOT NULL, preis_cent INTEGER NOT NULL CHECK(preis_cent >= 0),
        datum TEXT NOT NULL)''')
    return con


def save(values):
    cleaned = {key: values.get(key, '').strip() for key in FIELDS}
    for key in FIELDS:
        if not cleaned[key] or len(cleaned[key]) > 200:
            raise ValueError(f'{key}: Bitte 1 bis 200 Zeichen eingeben.')
    try:
        price = Decimal(cleaned['preis'].replace(',', '.'))
        if not price.is_finite() or price < 0 or price > Decimal('999999999.99') or price != price.quantize(Decimal('.01')):
            raise ValueError('Preis: maximal zwei Nachkommastellen und kein negativer Betrag.')
        cents = int(price * 100)
    except InvalidOperation:
        raise ValueError('Bitte einen gültigen Preis eingeben.') from None
    try:
        cleaned['datum'] = date.fromisoformat(cleaned['datum']).isoformat()
    except ValueError:
        raise ValueError('Bitte ein gültiges Datum eingeben.') from None
    with connect() as con:
        try:
            con.execute('INSERT INTO records VALUES (?, ?, ?, ?, ?, ?)',
                        (cleaned['id'], cleaned['vorname'], cleaned['nachname'], cleaned['produkt'], cents, cleaned['datum']))
        except sqlite3.IntegrityError:
            raise ValueError('Diese ID ist bereits vergeben. Bitte eine andere ID wählen.') from None


def records():
    with connect() as con:
        rows = con.execute('SELECT * FROM records ORDER BY datum, id').fetchall()
    return [dict(id=r['id'], vorname=r['vorname'], nachname=r['nachname'],
                 produkt=r['produkt'], preis=f"{r['preis_cent'] // 100}.{r['preis_cent'] % 100:02d}",
                 datum=r['datum']) for r in rows]


def export_xml(rows):
    root = ET.Element('datensaetze', version='1')
    for row in rows:
        element = ET.SubElement(root, 'datensatz', id=row['id'])
        for key in FIELDS[1:]:
            child = ET.SubElement(element, key, {'waehrung': 'EUR'} if key == 'preis' else {})
            child.text = row[key]
    ET.indent(root)
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def page(message='', values=None):
    values = values or {}
    inputs = ''
    for key, label in zip(FIELDS, LABELS):
        kind = 'date' if key == 'datum' else 'text'
        value = values.get(key, date.today().isoformat() if key == 'datum' else '')
        inputs += f'<label>{label}<input name="{key}" type="{kind}" maxlength="200" required value="{escape(value, quote=True)}"></label>'
    rows = records()
    table = ''.join('<tr>' + ''.join(f'<td>{escape(row[key])}</td>' for key in FIELDS) + '</tr>' for row in rows)
    return f'''<!doctype html><html lang="de"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>XML-Lernwerkstatt</title>
    <style>body{{font:16px system-ui;background:#f3f5fa;color:#18243b;max-width:1050px;margin:40px auto;padding:20px}}section{{background:white;padding:24px;border-radius:16px;margin:20px 0}}form{{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:16px}}label{{display:grid;gap:8px}}input,button{{font:inherit;padding:12px;border:1px solid #b6c3d5;border-radius:8px}}button{{background:#244dc4;color:white;cursor:pointer}}a{{color:#244dc4;margin-right:20px}}table{{width:100%;border-collapse:collapse}}td,th{{padding:12px;text-align:left;border-bottom:1px solid #ddd}}.scroll{{overflow:auto}}.message{{color:#9f2525}}</style>
    <h1>XML-Lernwerkstatt</h1><p>Daten erfassen → SQLite-Datenbank → XML</p>
    <section><h2>Datensatz erfassen</h2><p>Die ID wird von dir vergeben und muss eindeutig sein. Preis in Euro, z. B. 19,90.</p><p class="message" role="status">{escape(message)}</p><form method="post" action="/records">{inputs}<button>In Datenbank speichern</button></form></section>
    <section><h2>Gespeicherte Daten ({len(rows)})</h2><p><a href="/export.xml">XML herunterladen</a></p><div class="scroll"><table><thead><tr>{''.join(f'<th>{label}</th>' for label in LABELS)}</tr></thead><tbody>{table}</tbody></table></div>{'' if rows else '<p>Noch keine Daten. Erfasse deinen ersten Datensatz.</p>'}</section>
    <p>XML beschreibt die Datenstruktur. Eine spätere Word-Erzeugung kann diese XML-Daten lesen und auf eine Dokumentvorlage abbilden.</p></html>'''.encode('utf-8')


class Handler(BaseHTTPRequestHandler):
    def respond(self, data, content_type, status=200, filename=None):
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        if filename:
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if self.path == '/':
            self.respond(page(), 'text/html; charset=utf-8')
        elif self.path == '/export.xml':
            self.respond(export_xml(records()), 'application/xml; charset=utf-8', filename='datensaetze.xml')
        else:
            self.respond(b'Nicht gefunden', 'text/plain', 404)

    def do_POST(self):
        if self.path != '/records':
            self.respond(b'Nicht gefunden', 'text/plain', 404)
            return
        try:
            size = int(self.headers.get('Content-Length', '0'))
            if size < 1 or size > 16384:
                raise ValueError('Formular ist leer oder zu groß.')
            values = {k: v[0] for k, v in parse_qs(self.rfile.read(size).decode('utf-8')).items()}
            save(values)
        except (ValueError, UnicodeError) as error:
            self.respond(page(str(error), locals().get('values', {})), 'text/html; charset=utf-8', 400)
            return
        self.send_response(303)
        self.send_header('Location', '/')
        self.send_header('Content-Length', '0')
        self.end_headers()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--host', default='127.0.0.1')
    parser.add_argument('--port', type=int, default=8000)
    args = parser.parse_args()
    with connect():
        pass
    print(f'XML-Lernwerkstatt gestartet auf {args.host}:{args.port}', flush=True)
    ThreadingHTTPServer((args.host, args.port), Handler).serve_forever()
