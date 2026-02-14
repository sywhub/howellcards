#!/usr/bin/env python3
# Given # of pairs in a tournament, produce the table seating
# for each round suitable for Howell movements

# Howell movement general guideline is one pair (the highest) stays
# stationary at table one.  All other pairs move to where the lower-numbered
# pair sat preivously.  (Pair #4 goes to where pair #3 used to sit, for example).
# Further, it is best if pairs sit NS or EW alternatively for each round.

# We have previously generated all possible "seatings" for 6 to 12 pairs.

# Implement as a Python iterator
from maininit import setlog
import logging

# The constructor and the iterator members (__iter__ and __next__) are "for real".
# Other member functions are experimental code.
class HowellSeats:
    # Known good initial seating by # of tables
    # The tuple is seat assignments from the 2nd to the last table
    # The first table is the highest pair number and the 1st
    # This seatings allow "movements" to be consistent: a pair will always move to
    # where the previous numbered pair sat.
    GoodTables = {
        3: [(3, 4, 5, 2),(2, 5, 3, 4), (4, 3, 2, 5)],
        4: [(6, 5, 4, 2, 7, 3),
            (2, 4, 3, 7, 5, 6), (4, 2, 3, 7, 5, 6), (6, 2, 3, 4, 5, 7)],
        5: [(9, 8, 6, 4, 3, 7, 5, 2),
            (2, 3, 4, 7, 5, 9, 6, 8), (4, 2, 3, 6, 5, 9, 7, 8), (6, 2, 3, 4, 5, 8, 7, 9), (8, 2, 3, 5, 4, 9, 6, 7)],
        6: [(3, 10, 8, 7, 5, 11, 9, 6, 4, 2),
            (2, 3, 4, 7, 5, 9, 6, 11, 8, 10), (4, 2, 3, 6, 5, 10, 7, 11, 8, 9), (6, 2, 3, 4, 5, 10, 7, 9, 8, 11),
            (8, 2, 3, 6, 4, 11, 5, 7, 9, 10), (10, 2, 3, 7, 4, 6, 5, 11, 8, 9)],
        7: [(2, 3, 4, 7, 5, 10, 6, 13, 8, 12, 9, 11),
            (4, 2, 3, 6, 5, 12, 7, 11, 8, 13, 9, 10),
            (6, 2, 3, 4, 5, 12, 7, 10, 8, 13, 9, 11),
            (8, 2, 3, 4, 5, 13, 6, 9, 7, 11, 10, 12),
            (10, 2, 3, 4, 5, 7, 6, 12, 8, 11, 9, 13),
            (12, 2, 3, 4, 5, 9, 6, 11, 7, 13, 8, 10)]}

    def __init__(self, npairs, log = None, idx = 0):
        self.log = setlog('tables', log)
        self.log.info(f'Arrange table for {npairs} pairs')
        self.choice = idx
        odd = npairs % 2
        self.counter = npairs + odd - 1
        tables = (npairs + odd) // 2
        if tables in self.GoodTables:
            # for now, just pick the 1st good seating
            # we will experiment with others later
            if self.choice >= len(self.GoodTables[tables]):
                self.choice = -1
            self.seats = list(self.GoodTables[tables][self.choice])
            self.seats.insert(0, 1)
            self.seats.insert(0, npairs if not odd else 0)
        else:
            self.seats =  None

    def __iter__(self):
        return self

    def __len__(self):
        return len(self.seats)

    # Simply "move" all pairs to their next seating.
    # By Howell rules, it is where the one-lower numbered pair
    # used to sit.
    # This function does not assume the move is legit. That must be pre-valideated.
    # See initset.py for details.
    # Implement as a Python iterator
    def __next__(self):
        if self.counter <= 0:
              self.log.info('End of iteration')
              raise StopIteration

        self.counter -= 1
        if self.counter == 0:
            return self.seats

        # prepare for the next call
        dup = self.seats.copy()
        maxPair = self.seats[0]
        if maxPair == 0:
            maxPair = len(self.seats)
        for p in range(1,len(self.seats)):
            self.seats[p] += 1
            if self.seats[p] >= maxPair:
                self.seats[p] = 1
        return dup

    # discard existing seating and use the provided one
    # "tryData" is a list of pairs (as numbers) except for the first table
    # which is always the higest pair and "1"
    def resetSeat(self, tryData):
        self.log.info(f'Using provided seating: {tryData}')
        self.seats = list(tryData)
        # Add the first table for further validations
        self.seats.insert(0, 1)
        self.seats.insert(0, len(tryData) + 2)

    # "Tournament" is all the pairings for all the rounds
    # This function makes sure no pairs will meet twice with any other pairs
    def validateTournament(self, tournament):
        matches = {}
        valid = True
        for r in tournament:
            for i in range(0, len(r), 2):
                if not self.checkMember(matches, r[i], r[i+1]) or \
                    not self.checkMember(matches, r[i+1], r[i]):
                    self.log.info(f'{r[i]} and {r[i+1]} met before')
                    valid = False
                    break
        return valid

    # has we seen this pair before?
    def checkMember(self, members, k, v):
        if k not in members:
            members[k] = set()
        if v not in members[k]:
            members[k].add(v)
            return True
        return False

# List all validated seatings
# Test iterable implementation
def listAllSeatings():
    for k in sorted(HowellSeats.GoodTables.keys()):
        for i in range(len(HowellSeats.GoodTables[k])):
            tournament = []
            print(f'For {k} tables, seating #{i}:')
            howellSeats = HowellSeats(k*2, None, i)
            for j, tbl in enumerate(howellSeats):
                tournament.append(tbl)
                print(f'Round {j+1:>2}: {tbl}')
            howellSeats.validateTournament(tournament)

if __name__ == '__main__':
    listAllSeatings()
