#!/usr/bin/env python3

from functools import total_ordering
import string
import re

# The chars in VALID_REV_CHARS are all the valid options for positions in the rev
# IMPORTANT:  if the "-i" flag is used, VALID_REV_CHARS is overwritten with val of VALID_AND_INVALID_REV_CHARS
VALID_REV_CHARS = "-123456789ABCDEFGHJKLMNPRTUVWY"
VALID_AND_INVALID_REV_CHARS = "-123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
INVALID_REV_CHARS = "IOQSXZ"


def is_valid_rev(rev_text, mode=1):
    """
    :param rev_text: text of the rev being checked for validity
    :param mode:  mode 2 checks that revs contain only chars in VALID_REV_CHARS, while
    default mode 1 checks that revs contain only chars NOT in INVALID_REV_CHARS.  The
    reason for the two modes is that VALID_REV_CHARS is reset if the -i/--invalid-revs
    argument is present... if it's set we want to allow the Rev object to be created,
    but still need to be able to detect that the Rev is invalid per company standards,
    so we can issue a warning.
    :return: True or False, depending on whether the selected mode flags the rev as valid
    """

    # Numeric revisions are legal for redline releases, but must be converted to string.
    if isinstance(rev_text, int):
        rev_text = str(rev_text)

    # valid revs must be non-zero-length strings
    if not (isinstance(rev_text, str) or isinstance(rev_text, str)) or not len(rev_text):
        return False

    # we start by assuming there are no digits in this rev
    rev_has_digit = False

    for char in rev_text:

        if char in INVALID_REV_CHARS and mode == 1:
            return False

        if not char in VALID_REV_CHARS and mode == 2:
            return False

        # the dash character is only valid if it's the only character in the rev
        if char == "-" and len(rev_text) > 1:
            return False

        if char.isdigit():
            rev_has_digit = True

        # letters can't follow digits in a revision
        if char.isalpha() and rev_has_digit:
            return False

    return True


@total_ordering
class Rev(object):
    def __init__(self, name, eco=None):
        self.name = str(name).strip()
        self.eco = str(eco)

        # mode 2 checks that revs contain only chars in VALID_REV_CHARS, allowing this
        # Rev object to be created if the Rev is invalid per the CM standards but the
        # -i/--invalid-rev argument has been used.  We can't use mode 1 (checking that
        # revs contain only chars NOT in INVALID_REV_CHARS, because that list isn't
        # affected by the -i/--invalid-rev argument, i.e. Rev creation would fail on an
        # invalid rev even if that argument were set.
        if not is_valid_rev(self.name, mode=2):
            raise ValueError(self.name + " is not a valid rev!")

    def __eq__(self, other):
        return self.name == other.name

    def __gt__(self, other):

        # if these revs match per __eq__() above, we can short circuit, because neither is greater
        if self == other:
            return False

        # catches an oddity of revisions, i.e. that the rev after Y is AA, so Y < AA, and AY < BA
        if self.name.isalpha() and other.name.isalpha() and len(self.name) != len(other.name):
            return len(self.name) > len(other.name)

        for position in range(len(self.name)):
            # if we've moved past the first position and only one operand has a value, we know that the operand
            # of greater length is the larger one.  Ex. B1 is greater than B, and C is not greater than CA.
            if len(other.name) <= position < len(self.name):
                return True
            if len(self.name) <= position:
                if len(other.name) > position:
                    return False

                # if the first letters match, skip to the next letter
            if self.name[position] == other.name[position]:
                continue

            # if the current letters don't match, the one with the greater index number is greater
            return VALID_REV_CHARS.find(self.name[position]) > VALID_REV_CHARS.find(other.name[position])

        return False

    @property
    def next_rev(self):

        if VALID_REV_CHARS.find(self.name) in range(0, 10):
            return Rev("A")

        if VALID_REV_CHARS.find(self.name[-1]) in range(0, 10):
            new_name = self.name[:-2] + VALID_REV_CHARS[VALID_REV_CHARS.find(self.name[-2:-1])+1]

            # we never want to suggest an invalid next_rev, even if we're in -i/--invalid-revs mode
            if not is_valid_rev(new_name):
                return Rev(new_name).next_rev
            else:
                return Rev(new_name)

        if VALID_REV_CHARS.find(self.name[-1]) < 29:
            new_name = self.name[0:-1] + VALID_REV_CHARS[VALID_REV_CHARS.find(self.name[-1]) + 1]
            if not is_valid_rev(new_name):
                return Rev(new_name).next_rev
            else:
                return Rev(new_name)

        if self.name[-1] == "Y":
            new_name = self.name[0:-1] + "AA"
            return Rev(new_name)


# a compiled regular expression for the RAST part number format
PN_RE = re.compile(r'^\d\d\d\-\d\d\d\d\d\d-\d\d$')


def is_valid_part(pn_text):
    if not (isinstance(pn_text, str) or isinstance(pn_text, str)):
        return False

    if not PN_RE.match(pn_text):
        return False

    return True


class Part(object):
    def __init__(self, number, revs=None):
        self.number = str(number)

        if not revs:
            self.revs = {}
        else:
            self.revs = dict(revs)

        self.max_rev = None

        if not is_valid_part(number):
            raise ValueError(str(number).strip() + " is not a valid part number!")

    def has_rev(self, rev_text):
        # returns True or False, based on whether or not rev_text is one of self.rev's keys
        return rev_text in self.revs

    def add_rev(self, rev_text, eco=None):
        if self.has_rev(rev_text):
            return False

        self.revs[rev_text] = Rev(rev_text, eco)

        if not self.max_rev or (self.revs[rev_text] > self.max_rev):
            self.max_rev = Rev(rev_text)

        return True


class ListOfParts(object):
    def __init__(self, parts=None):
        if not parts:
            self.parts = {}
        else:
            self.parts = parts

    def add_part(self, pn, rev, eco=None):
        if not pn in list(self.parts.keys()):
            self.parts[pn] = Part(pn)

        return self.parts[pn].add_rev(rev, eco)

    def has_part(self, pn, rev):
        if not pn in list(self.parts.keys()):
            return False

        return self.parts[pn].has_rev(rev)

    def next_rev(self, pn):
        if not is_valid_part(pn):
            return None

        if not pn in self.parts:
            return Rev("-")

        return self.parts[pn].max_rev.next_rev

