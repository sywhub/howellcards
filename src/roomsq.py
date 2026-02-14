#!/usr/bin/env python3
# Room Square Problem
#
# A Room square, named after Thomas Gerald Room, is an n-by-n array filled with n + 1 different symbols in such a way that:
#
# 1. Each cell of the array is either empty or contains an unordered pair from the set of symbols
# 2. Each symbol occurs exactly once in each row and column of the array
# 3. Every unordered pair of symbols occurs in exactly one cell of the array.
# It is known that no solution exists for n == 3 or 5
# It has been solved up to n == 127

# For Bridge's Howell movement, we are interested in n in [5,7]
# Edwin Howell "invented" the Howell movements for the game of Whist, a predecessor or modern bridge.
# if n == 5 (6 pairs), we hack by designing Room Sq for n == 4 and insert a "shared board" for the last round

import argparse
import logging
import os
from maininit import setlog
import tables as Moves
import jsonIO
import itertools

class RoomSq:
    def __init__(self, n, tableIdx=0, log=None):
        self.log = setlog('roomsq', log, False)
        self.npairs = n
        self.tableIdx = tableIdx
        self.jIO = jsonIO.JsonIO(n, log)
        roundEven = n + n % 2
        self.boardSet = list(range(roundEven - 1))
        self.nTables = roundEven // 2
        self.jIO.meta(roundEven - 1, self.nTables)

    # sequence is how the "relay tables" are setup
    def assignTables(self, npairs, sequence):
        tblIter = Moves.HowellSeats(npairs, log, self.tableIdx)    # iterator for table seating
        n = len(self.boardSet)
        for round in tblIter:   # get the sitting for this round
            # seat each table
            tbls = [{'NS': round[i], 'EW': round[i+1]} for i in range(0,len(round),2)]
            # The 1st table always get the board equal to the round number
            # which is also the pair number sitting EW
            bFirst = tbls[0]['EW'] - 1
            tbls[0]['Board'] = bFirst
            # The subsequent tables get the board governed by the sequencing.
            for idx,t in enumerate(sequence):
                b = bFirst + t
                if b >= n:
                    b -= n
                tbls[idx+1]['Board'] = b
            self.jIO.addRound(tbls) # capture into JSON structure
        return

    def roomsq(self, fname):
        # 5x5 Room Square has no known solution
        if self.npairs == 6:
            return self.roomsq5by5(fname)

        # permutation to pick boards for the round
        # For each round, pick a board for each table.
        # First table is always the same arrangment, so skip it
        boards = self.boardSet[1:]
        boardForTables = itertools.permutations(boards, self.nTables - 1)
        valid = False
        # permute through all relay-table scenario till finding one
        for idx, seq in enumerate(boardForTables):
            self.log.info(f'{"-"*5}: {idx} {seq}')
            self.assignTables(self.npairs, seq)
            valid = self.jIO.validateBoards()
            if valid:
                self.log.info(f'Found Room Sq Solution')
                self.jIO.boardMovement(sorted(seq))
                self.jIO.sortByBoard()
                self.jIO.showArrangement()
                self.save2file(fname)
                break
            self.jIO.resetTournament()
        return valid

    def save2file(self, fname):
        mode = 'a' if os.path.exists(fname) else 'w'
        with open(fname, mode) as f:
            rm.jIO.dump2File(f)

    def roomsq5by5(self, fname):
        self.jIO.tournament = {'Rounds': 5, 'Tables': 3, 'BoardMovement': None, 'Arrangement':
            [[{'NS': 6, 'EW': 1, 'Board': 0}, {'NS': 3, 'EW': 4, 'Board': 1}, {'NS': 5, 'EW': 2, 'Board': 3}],
            [{'NS': 6, 'EW': 2, 'Board': 1}, {'NS': 4, 'EW': 5, 'Board': 2}, {'NS': 1, 'EW': 3, 'Board': 3}],
            [{'NS': 6, 'EW': 3, 'Board': 2}, {'NS': 5, 'EW': 1, 'Board': 1}, {'NS': 2, 'EW': 4, 'Board': 0}],
            [{'NS': 6, 'EW': 4, 'Board': 3}, {'NS': 1, 'EW': 2, 'Board': 2}, {'NS': 3, 'EW': 5, 'Board': 0}],
            [{'NS': 6, 'EW': 5, 'Board': 4}, {'NS': 2, 'EW': 3, 'Board': 4}, {'NS': 4, 'EW': 1, 'Board': 4}]]}
        self.jIO.showArrangement()
        self.save2file(fname)


if __name__ == '__main__':
    log = setlog('roomsq', None, False)
    logLevels = {'INFO': logging.INFO, 'DEBUG': logging.DEBUG, 'ERROR': logging.ERROR}
    parser = argparse.ArgumentParser()
    parser.add_argument('-p', '--pair', type=int, default=8)
    parser.add_argument('-i', '--index', type=int, default=0)
    parser.add_argument('-f', '--file', type=str, default='roomsq.txt')
    parser.add_argument('-d', '--debug', type=str, default='INFO')
    args = parser.parse_args()
    if args.debug.upper() in logLevels:
        log.setLevel(logLevels[args.debug.upper()])

    rm = RoomSq(args.pair, args.index, log)
    rm.roomsq(args.file)