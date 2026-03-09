#!/usr/bin/env python3
# Generate PDF file for generic scoring forms.
# Meant to be printed, copied, and cut into slips
# Travelers good for 8 rounds of play
# Pickup slip good for 6 boards per round
# Play Record good for 30 boards for the tournament
#
# These numbers were reasonable for normal amateur tournaments and optimal for 8x11 paper and human friendly font size.
#
import pdf
import datetime
from docset import PairGames

class GenericPDF(PairGames):
    def __init__(self):
        super().__init__(None)
        self.pdf = pdf.PDF()
        self.notice = 'For public domain. No rights reserved. Generated on'
        self.pdf.HeaderFooterText(f'{self.notice} {datetime.date.today().strftime("%b %d, %Y")}.',' ')
        self.nPerPg = 4
    
    # placeholders to facilitate reusing code in PairGames
    def pairN(self, n):
        return '' if n == None else n

    def pairID(self, n):
        return 'Pair:'+' '* 4

    # Fake data to reuse code
    # Print one sheet of 8 travelers each for 8 boards
    def printTravler(self):
        bData = {}
        for b in range(8):
            bData[b] = []
            for t in range(8):
                bData[b].append((None,None,str(t+1),None))    # round, table, NS, EW
        self.TravelersWithData(bData)

    # Fake data to reuse code
    # PairGames' function will print extra sheet of 4 pickup slips.  That's the only thing we want.
    def Pickups(self):
        tables = {0: {0: []}}
        for i in range(6):
            tables[0][0].append({'NS': None, 'EW': None, 'Board': None})
        self.PickupsWithData(tables)

    # Fake data to reuse code
    # Just enough to fit one page, 4 slips of 30 boards each
    def printRecords(self):
        jData = {}
        for pairNum in range(4):
            jData[pairNum] = []
            for b in range(30):
                jData[pairNum].append((b, None, None, None, None))
        self.JournalWithData(jData)

    def save(self):
        import os
        here = os.path.dirname(os.path.abspath(__file__))
        fn = f'{here}/../generic.pdf'
        self.pdf.output(fn)
        print(f'Saved {fn}')


    def printPDF(self):
        self.pdf.instructions(None, 'generic.txt')
        self.printTravler()
        self.Pickups()
        self.printRecords()
        self.save()
        return


if __name__ == '__main__':
    generic = GenericPDF()
    generic.printPDF()