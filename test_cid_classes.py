import unittest

from cid_classes import *

class CidClassesTest(unittest.TestCase):

    def test_alphabetic_revs(self):
        self.assertTrue(Rev("G") > Rev("A"))
        self.assertTrue(Rev("G") == Rev("G"))
        self.assertTrue(Rev("Y") < Rev("AA"))

    def test_numeric_revs(self):
        self.assertTrue(Rev("1") < Rev("5"))
        self.assertTrue(Rev("5") == Rev("5"))

    def test_alphanumeric_revs(self):
        self.assertTrue(Rev("B1") < Rev("B2"))
        self.assertTrue(Rev("C5") > Rev("B2"))

    def test_mixed_revs(self):
        self.assertTrue(Rev("B") > Rev("1"))
        self.assertTrue(Rev("1") < Rev("B2"))
        self.assertTrue(Rev("B1") > Rev("B"))
        self.assertTrue(Rev("B") < Rev("B2"))
        self.assertTrue(Rev("-") < Rev("1"))
        self.assertTrue(Rev("-") < Rev("A"))
        self.assertTrue(Rev("C5") > Rev("-"))

    def test_invalid_revs(self):
        for bad_rev in ["BO", "I", "AZ", "11", "B21", "1B1", "-A"]:
            with self.assertRaises(ValueError):
                print 'Verifying that Rev. {} is properly flagged as "bad"...'.format(bad_rev)
                next_rev=Rev(bad_rev)

if __name__ == "__main__":
    unittest.main()