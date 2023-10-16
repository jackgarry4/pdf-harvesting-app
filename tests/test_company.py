import unittest
from classes.Company import Company

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

    

    #Add more tests as needed