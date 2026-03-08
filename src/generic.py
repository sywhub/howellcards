#!/usr/bin/env python3
# Generate PDF file for generic scoring forms.
# Meant to be printed, copied, and cut into slips
# Travelers good for 8 rounds of play
# Pickup slip good for 6 boards per round
# Play Record good for 24 boards for the tournament
#
# These numbers were reasonable for normal amateur tournaments and optimal for 8x11 paper and human friendly font size.
#
import pdf

class GenericPDF:
    def __init__(self):
        import datetime
        self.pdf = pdf.PDF()
        self.notice = 'For public domain. No rights reserved. Generated on'
        self.pdf.HeaderFooterText(f'{self.notice} {datetime.date.today().strftime("%b %d, %Y")}.',' ')
        self.nPerPg = 4
        self.orgNotice = 'Sunnyvale Senior Center Bridge Club'
    
    def printOrg(self, shift, width, y):
        return
        w = self.pdf.get_string_width(self.orgNotice) - self.pdf.c_margin
        x = (width - w) / 2
        fSize = self.pdf.font_size_pt
        self.pdf.set_font(style='I', size=self.pdf.tinyPt)
        y += self.pdf.lineHeight(self.pdf.tinyPt) / 2
        self.pdf.set_xy(x+shift*width, y)
        self.pdf.cell(text=self.orgNotice)
        self.pdf.set_font_size(fSize)

    def printTravler(self):
        tblCols = []
        hdrs = ['NS','Bid'*2, 'By', 'M', 'M', 'NS', 'EW', 'vs.']
        self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.headerPt+1)
        self.pdf.setHeaders(0, hdrs, tblCols)
        xMargin = (self.pdf.w - 2*sum(tblCols)) / 4
        hdrs[1] = 'Bid'
        hdrs[3] = 'Made'
        hdrs[4] = 'Down'
        self.pdf.add_page()
        y = self.pdf.margin
        for b in range(self.nPerPg):
            y = self.printOneTraveler(xMargin, tblCols, hdrs, y)
            y = self.pdf.sectionDivider(self.nPerPg, b+1, xMargin)
        return

    def printOneTraveler(self, leftSide, tblCols, hdrs, y):
        halfW = self.pdf.w / 2
        startY = y
        for i in range(2):
            y = startY
            self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.notePt+1)
            y = self.pdf.headerRow(leftSide + halfW * i, y, tblCols, hdrs, 'Board:', 'Traveler')
            y += self.pdf.lineHeight(self.pdf.font_size_pt)
            self.pdf.set_font(self.pdf.sansSerifFont, size=self.pdf.notePt+1)
            h = self.pdf.lineHeight(self.pdf.font_size_pt)
            for r in range(8):
                self.pdf.set_xy(leftSide+halfW*i, y)
                self.pdf.cell(tblCols[0], h, text=f'{r+1}', align='C', border=1)
                for c in range(len(hdrs)-1):
                    self.pdf.cell(tblCols[c+1], h, text='', align='C', border=1)
                y += h
            self.printOrg(i, halfW, y)
        return y

    def printPickup(self):
        tblCols = []
        hdrs = ['NS Score', 'M', 'D', 'NS Bid', 'By', 'Board', 'EW Bid', 'By', 'M', 'M', 'EW Socre']
        self.pdf.set_font(self.pdf.sansSerifFont, style='B', size=self.pdf.linePt+1)
        hdrs[1] = hdrs[8] = 'Made'
        hdrs[2] = hdrs[9] = 'Down'
        self.pdf.setHeaders(0, hdrs, tblCols)
        xMargin = (self.pdf.w - sum(tblCols))/2
        self.pdf.add_page()
        y = self.pdf.margin
        for bIdx in range(8):
            self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.tinyPt)
            self.printOnePickup(tblCols, hdrs, xMargin, y)
            y = self.pdf.sectionDivider(8, bIdx+1, xMargin)
            y -= self.pdf.pt2in(self.pdf.tinyPt)
        return

    def printOnePickup(self, tblCols, hdrs, xMargin, y):
        y += self.pdf.lineHeight(self.pdf.font_size_pt)
        self.pdf.set_font(self.pdf.sansSerifFont, style='B', size=self.pdf.tinyPt)
        y = self.pdf.headerRow(xMargin, y, tblCols, hdrs, 'Table:'+' '*20+'Round:'+' '*20+'NS:'+' '*20+'EW:', 'Pickup Slip')
        self.pdf.set_font(size=self.pdf.tinyPt)
        h = self.pdf.lineHeight(self.pdf.font_size_pt)
        y += h
        self.pdf.set_xy(xMargin, y)
        for b in range(6):
            for i in range(len(hdrs)):
                self.pdf.cell(tblCols[i], h, text=f'', align='C', border=1)
            y += h
            self.pdf.set_xy(xMargin, y)

    def printRecords(self):
        tblCols = []
        hdrs = ['Board', 'vs.', 'Bid'*2, 'By', 'M', 'M', 'NS', 'EW']
        self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.linePt+2)
        self.pdf.setHeaders(0, hdrs, tblCols)
        xMargin = (self.pdf.w - 2*sum(tblCols)) / 4
        hdrs[2] = 'Bid'
        hdrs[4] = 'Made'
        hdrs[5] = 'Down'
        self.pdf.add_page()
        nPerPage = 2
        halfW = self.pdf.w / 2
        y = self.pdf.margin
        startY = y
        for p in range(nPerPage):  # two sets each page
            for i in range(2):
                y = startY
                self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.notePt)
                y = self.pdf.headerRow(xMargin+halfW*i, y, tblCols, hdrs ,"Pair:"+" "*20+"Names:", "Play Records")
                y += self.pdf.lineHeight(self.pdf.font_size_pt)
                self.pdf.set_font(size=self.pdf.smallPt-1)
                h = self.pdf.lineHeight(self.pdf.font_size_pt)
                self.pdf.set_xy(xMargin+halfW*i, y)
                for v in range(30): # 24 boards for the tournament
                    self.pdf.cell(tblCols[0], h, text=f'{v+1}', align='C', border=1)
                    for c in range(1,len(hdrs)):
                        self.pdf.cell(tblCols[c], h, text='', align='C', border=1)
                    y += h
                    self.pdf.set_xy(xMargin+halfW*i, y)
                self.printOrg(i, halfW, y)
            startY = self.pdf.sectionDivider(nPerPage, p+1, self.pdf.margin)
        return

    def save(self):
        import os
        here = os.path.dirname(os.path.abspath(__file__))
        fn = f'{here}/../generic.pdf'
        self.pdf.output(fn)
        print(f'Saved {fn}')


    def printPDF(self):
        self.pdf.instructions(None, 'generic.txt')
        self.printTravler()
        self.printPickup()
        self.printRecords()
        self.save()
        return


if __name__ == '__main__':
    generic = GenericPDF()
    generic.printPDF()