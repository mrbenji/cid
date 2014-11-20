import unittest

from cid_classes import *

class CID_ClassesTest(unittest.TestCase):

    def setUp(self):
        self.revs = {}
        for rev in ["-", "1", "5", "A", "B", "G", "AA", "BC", "B1", "B2", "C5", "Y"]:
            self.revs[rev] = Rev(rev)
        self.revs["Gdupe"] = Rev("G")
        self.revs["5dupe"] = Rev("5")

    def test_alphabetic_revs(self):
        self.assertTrue(self.revs["G"] > self.revs["A"])
        self.assertTrue(self.revs["G"] == self.revs["Gdupe"])
        self.assertTrue(self.revs["Y"] < self.revs["AA"])

    def test_numeric_revs(self):
        self.assertTrue(self.revs["1"] < self.revs["5"])
        self.assertTrue(self.revs["5"] == self.revs["5dupe"])

    def test_alphanumeric_revs(self):
        self.assertTrue(self.revs["B1"] < self.revs["B2"])
        self.assertTrue(self.revs["C5"] > self.revs["B2"])

    def test_mixed_revs(self):
        self.assertTrue(self.revs["B"] > self.revs["1"])
        self.assertTrue(self.revs["1"] < self.revs["B2"])
        self.assertTrue(self.revs["B1"] > self.revs["B"])
        self.assertTrue(self.revs["B"] < self.revs["B2"])
        self.assertTrue(self.revs["-"] < self.revs["1"])
        self.assertTrue(self.revs["-"] < self.revs["A"])
        self.assertTrue(self.revs["C5"] > self.revs["-"])

    def test_invalid_revs(self):
        for bad_rev in ["BO", "I", "AZ", "11", "B21", "1B1", "-A"]:
            with self.assertRaises(ValueError):
                print 'Verifying that Rev. {} is properly flagged as "bad"...'.format(bad_rev)
                next_rev=Rev(bad_rev)

if __name__ == "__main__":
    unittest.main()