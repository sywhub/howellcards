#!/usr/bin/env python3
import pdf
import logging
import datetime
from maininit import setlog

class GenericPDF:
    def __init__(self, log = None):
        self.log = log
        self.pdf = pdf.PDF()
        self.notice = 'For public domain. No rights reserved. Generated on'
        self.pdf.HeaderFooterText(f'{self.notice} {datetime.date.today().strftime("%b %d, %Y")}.',' ')
    
    def printTravler(self):
        self.log.debug(f'print Travelers')
        tblCols = []
        xMargin = 0.5
        hdrs = ['NS','Contract', 'By', 'Made', 'Down', '8'*4, '8'*4, 'vs.']
        self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.headerPt)
        self.pdf.setHeaders(xMargin, hdrs, tblCols)
        hdrs[5] = 'NS'
        hdrs[6] = 'EW'
        nPerPage = 4
        bIdx = 0
        for b in range(4):
            if bIdx % nPerPage == 0:
                self.pdf.add_page()
                y = 0.5
            y = self.printOneTraveler(xMargin, tblCols, hdrs, y)
            bIdx += 1
            if nPerPage > 1:
                y = self.pdf.sectionDivider(nPerPage, bIdx, xMargin)
        return

    def printOneTraveler(self, leftSide, tblCols, hdrs, y):
        self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.linePt)
        y = self.pdf.headerRow(leftSide, y, tblCols, hdrs, 'Board:', 'Traveler')
        y += self.pdf.lineHeight(self.pdf.font_size_pt)
        self.pdf.set_font(self.pdf.sansSerifFont, size=self.pdf.notePt)
        h = self.pdf.lineHeight(self.pdf.font_size_pt)
        for r in range(8):
            self.pdf.set_xy(leftSide, y)
            self.pdf.cell(tblCols[0], h, text=f'{r+1}', align='C', border=1)
            for c in range(len(hdrs)-1):
                self.pdf.cell(tblCols[c+1], h, text='', align='C', border=1)
            y += h
        return y

    def printPickup(self):
        self.log.debug(f'print Pickup Slips')
        tblCols = []
        xMargin = self.pdf.margin
        hdrs = ['NS Score', 'Made', 'Down', 'NS Contract', 'By', 'Board', 'EW Contract', 'By', 'Made', 'Down', 'EW Socre']
        self.pdf.set_font(self.pdf.sansSerifFont, style='B', size=self.pdf.notePt)
        self.pdf.setHeaders(xMargin, hdrs, tblCols)
        bIdx = 0
        for t in range(4):
            if bIdx % 4 == 0:
                self.pdf.add_page()
                y = 2 * self.pdf.margin
            self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.headerPt)
            self.printOnePickup(tblCols, hdrs, xMargin, y)
            bIdx += 1
            y = self.pdf.sectionDivider(4, bIdx, xMargin)
        return

    def printOnePickup(self, tblCols, hdrs, xMargin, y):
        y += self.pdf.lineHeight(self.pdf.font_size_pt)
        self.pdf.set_font(self.pdf.sansSerifFont, style='B', size=self.pdf.notePt)
        y = self.pdf.headerRow(xMargin, y, tblCols, hdrs, 'Table:        Round:        NS:       EW:', 'Pickup Slip')
        self.pdf.set_font(size=self.pdf.notePt)
        h = self.pdf.lineHeight(self.pdf.font_size_pt)
        y += h
        self.pdf.set_xy(xMargin, y)
        for b in range(6):
            for i in range(len(hdrs)):
                self.pdf.cell(tblCols[i], h, text=f'', align='C', border=1)
            y += h
            self.pdf.set_xy(xMargin, y)

    def printRecords(self):
        self.log.debug(f'print Play Records')
        tblCols = []
        pIdx = 0
        xMargin = self.pdf.margin * 2
        hdrs = ['Board', 'vs.', 'Contract', 'By', 'Made', 'Down', '8'*4, '8'*4]
        self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.linePt)
        self.pdf.setHeaders(xMargin, hdrs, tblCols)
        hdrs[6] = 'NS'
        hdrs[7] = 'EW'
        for p in range(4):
            if pIdx % 2 == 0:
                self.pdf.add_page()
                y = self.pdf.margin*2
            self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.linePt)
            y = self.pdf.headerRow(xMargin, y, tblCols, hdrs ,"Pair:", "Play Records")
            y += self.pdf.lineHeight(self.pdf.font_size_pt)
            self.pdf.set_font(size=self.pdf.smallPt)
            h = self.pdf.lineHeight(self.pdf.font_size_pt)
            self.pdf.set_xy(xMargin, y)
            for v in range(24):
                self.pdf.cell(tblCols[0], h, text=f'{v+1}', align='C', border=1)
                for c in range(1,len(hdrs)):
                    self.pdf.cell(tblCols[c], h, text='', align='C', border=1)
                y += h
                self.pdf.set_xy(xMargin, y)
            pIdx += 1
            y = self.pdf.sectionDivider(2, pIdx, self.pdf.margin)
        return

    def save(self):
        import os
        here = os.path.dirname(os.path.abspath(__file__))
        fn = f'{here}/../generic.pdf'
        self.log.debug(f'Save files: {fn}')
        self.pdf.output(fn)
        print(f'Saved {fn}')


    def printPDF(self):
        self.pdf.instructions(self.log, 'generic.txt')
        self.printTravler()
        self.printPickup()
        self.printRecords()
        self.save()
        return

def makeGenericPDF(log):
    generic = GenericPDF(log)
    generic.printPDF()

if __name__ == '__main__':
    log = setlog('generic', None)
    makeGenericPDF(log)