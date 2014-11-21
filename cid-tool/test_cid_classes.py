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
        self.assertEqual(Rev("A").next_rev, Rev("B"))
        self.assertEqual(Rev("H").next_rev, Rev("J"))
        self.assertEqual(Rev("-").next_rev, Rev("A"))
        self.assertEqual(Rev("1").next_rev, Rev("A"))
        self.assertEqual(Rev("5").next_rev, Rev("A"))
        self.assertEqual(Rev("Y").next_rev, Rev("AA"))
        self.assertEqual(Rev("B1").next_rev, Rev("C"))
        self.assertEqual(Rev("CA7").next_rev, Rev("CB"))

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

        # we haven't added any revs... these should all return False
        self.assertFalse(my_part.has_rev("A"))
        self.assertFalse("B1" in my_part.revs)
        self.assertFalse(my_part.has_rev("C"))

        # add_rev should return True if add was successful
        self.assertTrue(my_part.add_rev("A"))
        self.assertTrue(my_part.add_rev("B1"))
        self.assertTrue(my_part.add_rev("C"))

        # add_rev should return False if rev was previously added
        self.assertFalse(my_part.add_rev("A"))
        self.assertFalse(my_part.add_rev("B1"))
        self.assertFalse(my_part.add_rev("C"))

        # if the rev was added, these should return True
        self.assertTrue(my_part.has_rev("A"))
        self.assertTrue("B1" in my_part.revs)
        self.assertTrue(my_part.has_rev("C"))

        # add_rev should raise a ValueError if we attempt to add an invalid rev.
        with self.assertRaises(ValueError):
            my_part.add_rev("11")

        with self.assertRaises(ValueError):
            my_part.add_rev("AO")

        # test max_rev
        self.assertEqual(my_part.max_rev.name, "C")
        self.assertEqual(my_part.max_rev, Rev("C"))
        self.assertFalse(my_part.has_rev("D"))

    def test_add_part(self):
        my_list = ListOfParts()

        # we haven't added any parts... this should return False
        self.assertFalse(my_list.has_part("123-456789-01", "A"))

        # add_part should return True if add was successful
        self.assertTrue(my_list.add_part("123-456789-01", "A"))

        # add_part should return False if part/rev combo was previously added
        self.assertFalse(my_list.add_part("123-456789-01", "A"))

        # if the part/rev combo was truly added, has_part(pn, rev) should return True
        self.assertTrue(my_list.has_part("123-456789-01", "A"))

    def test_next_rev(self):
        my_list = ListOfParts()

        my_list.add_part("987-654321-01", "A")
        my_list.add_part("123-123123-11", "Y")

        self.assertEqual(my_list.next_rev("987-654321-01").name, "B")
        self.assertEqual(my_list.next_rev("123-123123-11").name, "AA")


if __name__ == "__main__":
    unittest.main()