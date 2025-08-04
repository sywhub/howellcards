#!/usr/bin/env python3
# Find working Howell table assignments for popular # of pairs
# Pairs are numbered from 1 to n.
# Pair n sits NS at table one and stay stationary
# All other pairs move to where the next higher pair sat in the previous round
# Pair one moves to where highest pair used to sit
# The goal is find all initial seat assignment so not no pairs will meet twice during n-1 rounds
import argparse
from tables import HowellSeats
import itertools
import json

def genSeats(nTbl):
    noMore = nTbl - 1
    ret = []
    perm = itertools.permutations(list(range(2,nTbl*2)))
    seenFirst = [x for x in range(2,nTbl*2) if x % 2]
    for trySeat in perm:
        hSeats = HowellSeats(nTbl*2)
        hSeats.resetSeat(trySeat)
        tourney = []
        if trySeat[0] not in seenFirst:
            for next in hSeats:
                tourney.append(next)
            goodHowell = hSeats.validateTournament(tourney)
            if goodHowell:
                ret.append(tuple(trySeat))
                seenFirst.append(trySeat[0])
                noMore -= 1
        if noMore <= 0:
            break
    return ret


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--table', type=int)
    args = parser.parse_args()
    allStr = ''
    if args.table:
        tList = [args.table]
    else:
        tList = list(range(3,8))
    for n in tList[:-1]:
        seatings = genSeats(n)
        jstr = json.dumps(seatings)
        allStr += str(n) + ': {' +jstr + '},\n'
    seatings = genSeats(tList[-1])
    allStr += str(tList[-1]) + ': {' + json.dumps(seatings) + '}\n'
    with open('inittable.txt', 'w') as f:
        print(allStr)
        print(allStr, file=f)