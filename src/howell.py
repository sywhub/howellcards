#!/usr/bin/env python3
# Generate Howell movement placards and scoring spreadsheet based on the number of pairs
# The tournament setups were pre-generated and stored as JSON.
# The algorithm of generating those data is a separate program, as those generations take time.

# --pair #: generate sheets for the specific pair #, if absent do all of them
# --fake: fake results to test the scoring mechanism in the spreadsheet
# --debug <DEBUG LEVEL>: used only by the developer

# to do: do Google sheet instead of Microsoft Excel
#        Smooother board transitions

import argparse
import logging
from maininit import setlog
import docset
import jsonIO

def howellFromJson(log, pairs, fake, jsonfile):
    jIO = jsonIO.JsonIO(pairs, log)
    tourney = jIO.load(jsonfile)
    if tourney:
        doc = docset.HowellDocSet(log, fake)
        doc.init(pairs, tourney['Rounds'])
        doc.saveByRound(tourney['Arrangement'])
        doc.saveByTable(tourney['Arrangement'])
        doc.saveByBoard(tourney['Arrangement'])
        doc.save()

if __name__ == '__main__':
    log = setlog('howell', None)
    parser = argparse.ArgumentParser()
    parser.add_argument('-p', '--pair', type=int, choices=range(5,15), help='# of pairs in the tournament')
    parser.add_argument('-f', '--fake', type=bool, default=False, help='Fake scores to test the spreadsheet')
    parser.add_argument('-d', '--debug', type=str, default='INFO')
    parser.add_argument('-j', '--jsonfile', type=str)
    args = parser.parse_args()
    for l in [['INFO', logging.INFO], ['DEBUG', logging.DEBUG], ['ERROR', logging.ERROR]]:
        if args.debug.upper() == l[0]:
            log.setLevel(l[1])
            break

    if args.pair: 
        howellFromJson(log, args.pair, args.fake, args.jsonfile)
    elif args.pair is None:
        for p in range(5,15):
            howellFromJson(log, p, args.fake, args.jsonfile)
