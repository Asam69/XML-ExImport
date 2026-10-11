from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
import app


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.old_db = app.DB
        app.DB = Path(self.temp.name) / 'test.sqlite3'
        self.sample = dict(id='1001', vorname='Anna', nachname='Müller', produkt='Buch & Stift <Set>', preis='12,90', datum='2026-10-11')

    def tearDown(self):
        app.DB = self.old_db
        self.temp.cleanup()

    def test_database_xml_roundtrip(self):
        app.save(self.sample)
        rows = app.records()
        self.assertEqual(rows[0]['preis'], '12.90')
        xml = app.export_xml(rows)
        item = ET.fromstring(xml).find('datensatz')
        self.assertEqual(item.attrib['id'], '1001')
        for key in ('vorname', 'nachname', 'produkt', 'datum'):
            self.assertEqual(item.findtext(key), self.sample[key])
        self.assertEqual(item.find('preis').attrib['waehrung'], 'EUR')

    def test_invalid_data_and_duplicate(self):
        for key, value in [('preis', '-1'), ('preis', 'NaN'), ('preis', '1.001'), ('datum', '2026-02-30'), ('vorname', '')]:
            with self.subTest(key=key, value=value), self.assertRaises(ValueError):
                app.save(dict(self.sample, **{key: value}))
        self.assertEqual(app.records(), [])
        app.save(self.sample)
        with self.assertRaisesRegex(ValueError, 'ID'):
            app.save(self.sample)
        self.assertEqual(len(app.records()), 1)

    def test_empty_exports(self):
        xml = app.export_xml(app.records())
        self.assertEqual(len(ET.fromstring(xml)), 0)



if __name__ == '__main__':
    unittest.main()
