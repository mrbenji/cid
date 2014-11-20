from functools import total_ordering
import string

# the following revision letters are not valid
INVALID_REVS = ["I", "O", "Q", "S", "X", "Z"]
BETA_REVS = [str(x) for x in range (1,10)]

REV_INDEX = {'-':0}
for x in range(1,10):
    REV_INDEX[str(x)] = x

i = 10
for x in string.ascii_uppercase:
    if x not in INVALID_REVS:
        REV_INDEX[x] = i
    i += 1

def is_valid_rev(rev_text):

    # valid revs must be non-zero-length strings
    if not isinstance(rev_text, str) and len(rev_text)>0:
        return False

    # we start by assuming there are no digits in this rev
    rev_has_digit = False

    for char in rev_text:

        # the keys of REV_INDEX are all the valid options for positions in the rev
        if not char in REV_INDEX.keys():
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

@total_ordering
class Rev(object):
    def __init__(self, name):
        self.name = str(name)
        if not is_valid_rev(name):
            raise ValueError(str(name).strip() + " is not a valid rev!")

    def __eq__(self, other):
        return self.name == other.name

    def __gt__(self, other):

        # if these revs match per __eq__() above, we can short circuit, because neither is greater
        if self == other:
            return False

        # catches an oddity of revisions, i.e. that the rev after Y is AA, so Y < AA
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

            #
            if REV_INDEX[self.name[position]] > REV_INDEX[other.name[position]]:
                return True

        return False

    def next_rev(self):
        pass


class Part(object):
    def __init__(self, number, revs, version, description):
        self.number = str(number)
        self.revs = list(revs)
        self.version = str(version)
        self.description = str(description)
        self.max_rev = None

class ListOfParts(object):
    def __init__(self, parts):
        self.parts = list(parts)

    def add_part(self, part, rev, version="", description=""):
        new_part = Part(part, rev, version, description)
        self.parts.append(new_part)

    def part_max_rev(self, part):
        pass

