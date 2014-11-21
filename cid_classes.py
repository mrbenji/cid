from functools import total_ordering
import string
import re

# the chars in VALID_REV_CHARS are all the valid options for positions in the rev
VALID_REV_CHARS = "-123456789ABCDEFGHJKLMNPRTUVWY"

@total_ordering
class Rev(object):
    def __init__(self, name):
        self.name = str(name)
        if not self.is_valid_rev(name):
            raise ValueError(str(name).strip() + " is not a valid rev!")

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
            if len(other.name) <= position and len(self.name) > position:
                return True
            if len(self.name) <= position and len(other.name) > position:
               return False

            # if the first letters match, skip to the next letter
            if self.name[position] == other.name[position]:
                continue

            # if the current letters don't match, the one with the greater index number is greater
            return VALID_REV_CHARS.find(self.name[position]) > VALID_REV_CHARS.find(other.name[position])

        return False

    def is_valid_rev(self, rev_text):

        # valid revs must be non-zero-length strings
        if not len(rev_text) or not isinstance(rev_text, str):
            return False

        # we start by assuming there are no digits in this rev
        rev_has_digit = False

        for char in rev_text:

            if not char in VALID_REV_CHARS:
                return False

            # the dash character is only valid if it's the only character in the rev
            if char == "-" and len(rev_text)>1:
                return False

            # non-redline revisions cannot contain more than one digit
            if char.isdigit():
                if rev_has_digit:
                    return False
                else: rev_has_digit = True

            # letters can't follow digits in a revision
            if char.isalpha() and rev_has_digit:
                return False

        return True

    def next_rev(self):

        if VALID_REV_CHARS.find(self.name) in range (0,10):
            return Rev("A")

        if VALID_REV_CHARS.find(self.name[-1]) in range (0,10):
            new_name = self.name[:-2] + VALID_REV_CHARS[VALID_REV_CHARS.find(self.name[-2:-1])+1]
            return Rev(new_name)

        if VALID_REV_CHARS.find(self.name[-1]) < 29:
            new_name = self.name[0:-1] + VALID_REV_CHARS[VALID_REV_CHARS.find(self.name[-1]) + 1]
            return Rev(new_name)

        if self.name[-1] == "Y":
            new_name = self.name[0:-1] + "AA"
            return Rev(new_name)


# a compiled regular expression for the RAST part number format
PN_RE = re.compile(r'^\d\d\d\-\d\d\d\d\d\d-\d\d$')

def is_valid_part(pn_text):
    if not isinstance(pn_text, str):
        return False

    if not PN_RE.match(pn_text):
        return False

    return True


class Part(object):
    def __init__(self, number, revs={}):
        self.number = str(number)
        self.revs = dict(revs)
        self.max_rev = None
        if not is_valid_part(number):
            raise ValueError(str(number).strip() + " is not a valid part number!")

    def has_rev(self, rev_text):
        return self.revs.has_key(rev_text)

    def add_rev(self, rev_text):
        if self.has_rev(rev_text):
            return False

        self.revs[rev_text] = Rev(rev_text)

        if not self.max_rev or (self.revs[rev_text] > self.max_rev):
            self.max_rev = Rev(rev_text)

        return True


class ListOfParts(object):
    def __init__(self, parts={}):
        self.parts = parts

    def add_part(self, pn, rev):
        if not pn in self.parts.keys():
            self.parts[pn] = Part(pn)

        return self.parts[pn].add_rev(rev)

    def has_part(self, pn, rev):
        if not pn in self.parts.keys():
            return False

        return self.parts[pn].has_rev(rev)

    def next_rev(self, pn):
        if not is_valid_part(pn):
            return ""

        if not pn in self.parts.keys():
            return "-"

        return self.parts[pn].max_rev.next_rev()

