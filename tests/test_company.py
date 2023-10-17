import unittest
from classes.Company import Company
from classes.PDF import PDF

class TestCompany(unittest.TestCase):
    def setUp(self):
        self.company1 = Company("Geico")
        self.company2 = Company("Cintas", 150, 45000000)

    def test_initialization_defaults(self):
        self.assertEqual(self.company1.name, "Geico")
        self.assertEqual(self.company1.pdfs, [])
        self.assertEqual(self.company1.planParticipants, 0)
        self.assertEqual(self.company1.assets, 0)

    def test_intialization_with_custom_values(self):
        self.assertEqual(self.company2.name, "Cintas")
        self.assertEqual(self.company2.planParticipants, 150)
        self.assertEqual(self.company2.assets, 45000000)

    def test_adding_pdf_links(self):
        pdf1 = PDF("pdf1.pdf")
        pdf2 = PDF("pdf2.pdf")
        self.company1.add_pdf(pdf1)
        self.assertEqual(self.company1.pdfs, [pdf1])

        self.company2.add_pdf(pdf1)
        self.company2.add_pdf(pdf2)
        self.assertEqual(self.company2.pdfs, [pdf1, pdf2])

    

    #Add more tests as needed