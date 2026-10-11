from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZipFile
from generate_word_documents import generate, ROOT, W


class GenerationTests(unittest.TestCase):
    def test_three_records_create_separate_bound_documents(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            root = ET.Element('datensaetze', version='1')
            for index in range(1, 4):
                row = ET.SubElement(root, 'datensatz', id=str(index))
                for key, val in dict(vorname=f'Anna {index}', nachname='Müller', produkt='Buch & Stift', preis='12.90', datum='2026-10-11').items():
                    ET.SubElement(row, key).text = val
            xml = folder/'daten.xml'
            ET.ElementTree(root).write(xml, encoding='utf-8', xml_declaration=True)
            paths = generate(xml, ROOT/'word'/'Produktbeleg-Vorlage.dotx', folder/'out')
            self.assertEqual(len(paths), 3)
            for index, path in enumerate(paths, 1):
                with ZipFile(path) as archive:
                    self.assertIsNone(archive.testzip())
                    single = ET.fromstring(archive.read('customXml/item1.xml'))
                    self.assertEqual(len(single.findall('datensatz')), 1)
                    self.assertEqual(single.find('datensatz').get('id'), str(index))
                    doc = ET.fromstring(archive.read('word/document.xml'))
                    values = [node.text for node in doc.findall('.//'+W+'sdtContent/'+W+'r/'+W+'t')]
                    self.assertIn(f'Anna {index}', values)
                    self.assertIn('Buch & Stift', values)
                    ct = ET.fromstring(archive.read('[Content_Types].xml'))
                    self.assertTrue(any(node.get('ContentType', '').endswith('wordprocessingml.document.main+xml') for node in ct))
            with self.assertRaisesRegex(ValueError, 'existiert'):
                generate(xml, ROOT/'word'/'Produktbeleg-Vorlage.dotx', folder/'out')
