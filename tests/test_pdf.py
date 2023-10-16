import unittest
from classes.PDF import PDF

class TestPDF(unittest.TestCase):
    def setUp(self):
        self.pdf1 = PDF("pdf1.pdf")
        self.pdf2 = PDF("pdf2.pdf", "Tax PDF", "C:/pdf2", "TaxDocs.com")

    def test_initialization_defaults(self):
        self.assertEqual(self.pdf1.url, "pdf1.pdf")
        self.assertEqual(self.pdf1.title, "")
        self.assertEqual(self.pdf1.fileLocation, "")
        self.assertEqual(self.pdf1.source, "TransAmerica")

    def test_intialization_with_custom_values(self):
        self.assertEqual(self.pdf2.url, "pdf2.pdf")
        self.assertEqual(self.pdf2.title, "Tax PDF")
        self.assertEqual(self.pdf2.fileLocation, "C:/pdf2")
        self.assertEqual(self.pdf2.source, "TaxDocs.com")

    

    #Add more tests as needed