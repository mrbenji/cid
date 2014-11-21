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
        self.assertEqual(Rev("Y").next_rev(), Rev("AA"))
        self.assertEqual(Rev("B1").next_rev(), Rev("C"))
        self.assertEqual(Rev("CA7").next_rev(), Rev("CB"))

    def test_valid_part_numbers(self):
        self.assertEqual(Part("123-456789-01").number, "123-456789-01")
        self.assertEqual(Part("145-123456-00").number, "145-123456-00")

    def test_invalid_part_numbers(self):
        with self.assertRaises(ValueError):
            Part("123")
        with self.assertRaises(ValueError):
            Part("123456789")
        with self.assertRaises(ValueError):
            Part("987-654-3210")

    def test_add_rev(self):
        my_part = Part("123-456789-01")

        # we haven't added any parts... these should all return False
        self.assertFalse(my_part.revs.has_key("A"))
        self.assertFalse(my_part.revs.has_key("B1"))
        self.assertFalse(my_part.revs.has_key("C"))

        # add_rev should return True if add is successful
        self.assertTrue(my_part.add_rev("A"))
        self.assertTrue(my_part.add_rev("B1"))
        self.assertTrue(my_part.add_rev("C"))

        # add_rev should return False if rev was previously added
        self.assertFalse(my_part.add_rev("A"))
        self.assertFalse(my_part.add_rev("B1"))
        self.assertFalse(my_part.add_rev("C"))

        # if the rev was added, has_key(rev) should return True
        self.assertTrue(my_part.revs.has_key("A"))
        self.assertTrue(my_part.revs.has_key("B1"))
        self.assertTrue(my_part.revs.has_key("C"))

        # add_rev should raise a ValueError if we attempt to add an invalid rev.
        with self.assertRaises(ValueError):
            my_part.add_rev("11")

        with self.assertRaises(ValueError):
            my_part.add_rev("AO")


if __name__ == "__main__":
    unittest.main()