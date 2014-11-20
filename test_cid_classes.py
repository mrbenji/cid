import unittest

from cid_classes import *

class CidClassesTest(unittest.TestCase):

    def test_alphabetic_revs(self):
        self.assertTrue(Rev("G") > Rev("A"))
        self.assertTrue(Rev("G") == Rev("G"))
        self.assertTrue(Rev("Y") < Rev("AA"))
        self.assertTrue(Rev("AAB") > Rev("AAA"))

    def test_numeric_revs(self):
        self.assertTrue(Rev("1") < Rev("5"))
        self.assertTrue(Rev("5") == Rev("5"))

    def test_alphanumeric_revs(self):
        self.assertTrue(Rev("B1") < Rev("B2"))
        self.assertTrue(Rev("C5") > Rev("B2"))
        self.assertTrue(Rev("B1") == Rev("B1"))

    def test_mixed_revs(self):
        self.assertTrue(Rev("B") > Rev("1"))
        self.assertTrue(Rev("1") < Rev("B2"))
        self.assertTrue(Rev("B1") > Rev("B"))
        self.assertTrue(Rev("B") < Rev("B2"))
        self.assertTrue(Rev("-") < Rev("1"))
        self.assertTrue(Rev("-") < Rev("A"))
        self.assertTrue(Rev("C5") > Rev("-"))
        self.assertTrue(Rev("B2") != Rev("B"))

    def test_invalid_revs(self):
        print 'Verifying that revs are being properly flagged as invalid...'
        for bad_rev in ["I", "AZ", "11", "B21", "1B1", "-A", "1C"]:
            print " {}".format(bad_rev),
            with self.assertRaises(ValueError):
                Rev(bad_rev)
        print "\n"

    def test_next_rev(self):
        self.assertEqual(Rev("A").next_rev(), Rev("B"))
        self.assertEqual(Rev("H").next_rev(), Rev("J"))
        self.assertEqual(Rev("-").next_rev(), Rev("A"))
        self.assertEqual(Rev("1").next_rev(), Rev("A"))
        self.assertEqual(Rev("5").next_rev(), Rev("A"))

if __name__ == "__main__":
    unittest.main()