"""Run: python3 -m unittest discover -s tools -p 'test_*.py' -v"""
from pathlib import Path
import unittest
from zipfile import ZipFile
import xml.etree.ElementTree as ET


class WordPackageTests(unittest.TestCase):
    def test_unique_content_types_and_relationships(self):
        for path in (Path(__file__).resolve().parents[1] / 'word').iterdir():
            if path.suffix not in ('.docx', '.dotx'):
                continue
            with self.subTest(file=path.name), ZipFile(path) as archive:
                self.assertIsNone(archive.testzip())
                content_types = ET.fromstring(archive.read('[Content_Types].xml'))
                parts = [node.get('PartName') for node in content_types if node.get('PartName')]
                self.assertEqual(len(parts), len(set(parts)))
                for name in archive.namelist():
                    if name.endswith(('.xml', '.rels')):
                        root = ET.fromstring(archive.read(name))
                        if name.endswith('.rels'):
                            identities = [node.get('Id') for node in root]
                            targets = [(node.get('Type'), node.get('Target')) for node in root]
                            self.assertEqual(len(identities), len(set(identities)))
                            self.assertEqual(len(targets), len(set(targets)))
