#!/usr/bin/env python3
# Generate a team match setup
#   A PDF with Roster and Score sheets
#   An Excel spreadsheet to enter the results and calculate the scores
#
# This program arrange a tournament for 4 pairs to play in 1 to 3 rounds of "team matches".
# Each match is formally a match of 2 teams of 2 pairs.  At the end of each match, we change the composition
# of both teams.  In 3 matches, therefore, each pair has played with the other 2.
#
import numbers
import argparse
import logging
from maininit import setlog
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
import pdf
from docset import DupBridge
import datetime

class Mitchell(DupBridge):
    def __init__(self, log, p, b, f):
        super().__init__(log)
        self.pairs = p
        self.boards = b
        self.tables = (self.pairs + 1) // 2
        self.boardSet = self.tables
        if self.tables % 2 == 0:
            self.boardSet += 1
        self.oddPairs = self.pairs % 2 == 1
        self.SITOUT = "Sit-Out"
        self.fake = f
        self.footer = f'Mitchell Tournament: {(self.pairs+1)//2} Tables, {self.boards} Boards per round'
        self.noChangeFont = Font(bold=True, italic=True, color='FF0000')
        self.pdf = pdf.PDF()
        self.wb = Workbook()
    
    def go(self):
        self.roster()
        self.roundTab()
        self.boardTab()
        self.ScoreTable()
        self.Pickups()
        self.Journal()
        self.save()
        return

    def roster(self):
        self.rosterSheet()
        self.rosterPDF()
    
    def rosterSheet(self):
        ws = self.wb.active # the first tab
        ws.title = 'Roster'
        # First simple list of names
        row = 1
        ws.cell(row, 1).value = self.footer
        ws.cell(row, 1).font = self.HeaderFont
        ws.merge_cells(f'{ws.cell(row,1).coordinate}:{ws.cell(row,5).coordinate}')
        ws.cell(row, 1).alignment = self.centerAlign
        row += 2
        
        ws.cell(row, 1).value = 'Players'
        ws.cell(row, 1).font = self.HeaderFont
        ws.cell(row, 1).alignment = self.centerAlign
        ws.merge_cells(f'{ws.cell(row,1).coordinate}:{ws.cell(row,3).coordinate}')
        ws.cell(row, 4).value = '%'
        ws.cell(row, 4).font = self.HeaderFont
        ws.cell(row, 4).alignment = self.centerAlign
        ws.cell(row, 5).value = 'Score'
        ws.cell(row, 5).font = self.HeaderFont
        ws.cell(row, 5).alignment = self.centerAlign
        row += 1
        for p in range(self.pairs):
            ws.cell(row, 1).font = self.HeaderFont
            ws.cell(row, 1).value = f'Pair {p+1}'
            ws.cell(row, 2).value = self.placeHolderName()
            ws.cell(row, 3).value = self.placeHolderName()
            row += 1
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 30
        
    def rosterPDF(self):
        self.headerFooter()
        y = self.meta()
        self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.rosterPt) 
        h = self.pdf.lineHeight(self.pdf.font_size_pt)
        title = 'Player Pairs'
        x = self.pdf.setHCenter(self.pdf.get_string_width(title))
        y += 2 * h
        self.pdf.set_xy(x, y)
        self.pdf.cell(text=title)
        widths = [1, 2, 2]
        y +=  h
        leftM = (self.pdf.w - sum(widths)) / 2
        self.pdf.set_xy(leftM, y)
        self.pdf.set_font(self.pdf.sansSerifFont, size=self.pdf.bigPt) 
        h = self.pdf.lineHeight(self.pdf.font_size_pt)
        for p in range(self.pairs):
            self.pdf.cell(widths[0], h, text=f'Pair {p+1}', align='C', border=1)
            self.pdf.cell(widths[1], h, text='', align='C', border=1)
            self.pdf.cell(widths[2], h, text='', align='C', border=1)
            y += h
            self.pdf.set_xy(leftM, y)
        return

    def meta(self):
        self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.rosterPt) 
        h = self.pdf.lineHeight(self.pdf.font_size_pt)
        title = 'Mitchell Tournament'
        x = self.pdf.setHCenter(self.pdf.get_string_width(title))
        y = self.pdf.margin + h
        self.pdf.set_xy(x, y)
        self.pdf.cell(text=title)
        self.pdf.set_font(self.pdf.sansSerifFont, size=self.pdf.bigPt) 
        h = self.pdf.lineHeight(self.pdf.font_size_pt)
        x += h
        y += h
        self.pdf.set_xy(x, y)
        self.pdf.cell(text=f'{self.pairs} pairs')
        y += h
        self.pdf.set_xy(x, y)
        self.pdf.cell(text=f'{self.pairs // 2 - 1} rounds')
        y += h
        self.pdf.set_xy(x, y)
        self.pdf.cell(text=f'{self.boards} boards per round')
        y += h
        self.pdf.set_xy(x, y)
        return self.pdf.get_y()

    def boardTab(self):
        self.log.debug('Saving by Board')
        headers = ['Board', 'Round', 'Table', 'NS', 'EW', 'Vul', 'Contract', 'By', 'Result', 'NS', 'EW']
        sh, row = self.contractHeaders(headers, 'By Board', ['Scores','%', 'Pts', 'Net'], 1)
        calcs = ['NS', 'EW'] * 3
        col = len(headers) + 1
        for h in calcs:
            sh.cell(row-1, col).value = h
            sh.cell(row-1, col).font = self.HeaderFont
            sh.cell(row-1, col).alignment = self.centerAlign
            col += 1
        rGap = self.tables * self.boards
        for b in self.boardData.keys():
            sh.cell(row, 1).value = b+1
            sh.cell(row, 1).alignment = self.centerAlign
            for r in self.boardData[b]: # (round, table, NS, EW)
                sh.cell(row, 2).value = f"='By Round'!{self.rc2a1(r[0] * rGap + 3, 1)}"
                tBase = r[0] * rGap + r[1] * self.boards + 3
                sh.cell(row, 3).value = f"='By Round'!{self.rc2a1(tBase, 2)}"
                sh.cell(row, 4).value = f"='By Round'!{self.rc2a1(tBase, 3)}"
                sh.cell(row, 5).value = f"='By Round'!{self.rc2a1(tBase, 4)}"
                cIdx = len(headers)
                rawNS = self.rc2a1(row, cIdx - 1)
                rawEW = self.rc2a1(row, cIdx)
                sh.cell(row, cIdx+5).value = f'=IF(ISNUMBER({rawNS}),{rawNS},IF(ISNUMBER({rawEW}),-{rawEW},""))'
                sh.cell(row, cIdx+6).value = f'=IF(ISNUMBER({rawEW}),{rawEW},IF(ISNUMBER({rawNS}),-{rawNS},""))'
                tBase += b % self.boards
                for i in range(6, len(headers)+1):
                    c = f"'By Round'!{self.rc2a1(tBase, i)}"
                    sh.cell(row, i).value = f'=IF(ISBLANK({c}),"",{c})'
                for i in range(2,7):
                    sh.cell(row, i).alignment = self.centerAlign
                row += 1
        return

    def roundTab(self):
        self.log.debug('Saving by Round')
        headers = ['Round', 'Table', 'NS', 'EW', 'Board', 'Vul', 'Contract', 'By', 'Result', 'NS', 'EW']
        sh, row = self.contractHeaders(headers, 'By Round', ['Scores'])
        rounds = self.tables
        self.boardData = {}
        f = 0
        for r in range(rounds): # round
            sh.cell(row, 1).value = r+1
            sh.cell(row, 1).alignment = self.centerAlign
            for t in range(rounds): # table
                sh.cell(row, 2).value = t+1
                sh.cell(row, 3).value = self.NSPair(r, t)
                sh.cell(row, 4).value = self.EWPair(r, t)
                for i in range(2,5):
                    sh.cell(row, i).alignment = self.centerAlign
                b = self.boardIdx(r, t)
                for bset in range(self.boards):
                    if (b + bset) not in self.boardData:
                        self.boardData[b+bset] = []
                    sh.cell(row, 5).value = b+bset+1
                    self.boardData[b+bset].append((r, t, self.NSPair(r, t), self.EWPair(r, t)))
                    sh.cell(row, 6).value = f"{self.vulLookup(b+bset)}"
                    sh.cell(row, 5).alignment = self.centerAlign
                    sh.cell(row, 6).alignment = self.centerAlign
                    if self.fake:
                        f += 1
                        sh.cell(row, 10 if r < self.tables // 2 else 11).value = f
                    row += 1
        return

    def NSPair(self, r, t):
        if t == self.tables and self.oddPairs:
            return self.SITOUT
        return t * 2 + 1

    def EWPair(self, r, t):
        ew = self.tables - (self.tables - t + r - 1) % self.tables
        ew *= 2
        return ew
    
    def boardIdx(self, r, t):
        b = (r + t) % self.boardSet
        b *= self.boards
        return b

    def contractHeaders(self, hdrs, tabName, merges, tabIdx=2):
        sh = self.wb.create_sheet(tabName, tabIdx)
        sCol = hdrs.index('Result')+2
        for i in range(len(merges)):
            sh.cell(1, sCol).value = merges[i]
            sh.merge_cells(f'{self.rc2a1(1, sCol)}:{self.rc2a1(1, sCol+1)}')
            sh.cell(1, sCol).font = self.HeaderFont
            sh.cell(1, sCol).alignment = self.centerAlign
            sCol += 2
        row = self.headerRow(sh, hdrs, 2)
        sh.column_dimensions[chr(hdrs.index('Contract')+ord('A'))].width = 30
        return sh, row

    def Pickups(self):
        return

    def Journal(self):
        return

    # Output into filesystem
    def save(self):
        import os
        here = os.path.dirname(os.path.abspath(__file__))
        fn = f'{here}/../mitchell{self.pairs}x{self.boards}'
        self.wb.save(f'{fn}.xlsx')
        self.pdf.output(f'{fn}.pdf')
        print(f'Saved {fn}.{{xlsx,pdf}}')

    def headerFooter(self):
        notice = f'For public domain. No rights reserved. {datetime.date.today().strftime("%Y")}.'
        self.pdf.set_font(size=self.pdf.tinyPt)
        h = self.pdf.lineHeight(self.pdf.font_size_pt)
        w = self.pdf.get_string_width(notice)
        x = self.pdf.setHCenter(w)
        self.pdf.set_xy(x, h)
        self.pdf.cell(text=notice)
        w = self.pdf.get_string_width(self.footer)
        x = self.pdf.setHCenter(w)
        y = self.pdf.eph - h * 2
        self.pdf.set_xy(x, y)
        self.pdf.cell(text=self.footer)


if __name__ == '__main__':
    log = setlog('mitchell', None)
    def even_type(value):
        ivalue = int(value)
        if ivalue % 2 != 0:
            raise argparse.ArgumentTypeError(f"{value} is not an even number")
        return ivalue

    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--debug', type=str, default='INFO', help='Debug level, INFO, DEBUG, ERROR')
    parser.add_argument('-b', '--boards', type=int, choices=range(2,7), default=4, help='Number of pairs')
    parser.add_argument('-p', '--pair', type=even_type, choices=range(7,25), default=8, help='Number of pairs')
    parser.add_argument('-f', '--fake', type=bool, default=False, help='Fake scores to test the spreadsheet')
    args = parser.parse_args()
    for l in [['INFO', logging.INFO], ['DEBUG', logging.DEBUG], ['ERROR', logging.ERROR]]:
        if args.debug.upper() == l[0]:
            log.setLevel(l[1])
            break
    mitchell = Mitchell(log, args.pair, args.boards, args.fake)
    # A match has n rounds, each round has m boards, divided into two halves, each half of the boards
    mitchell.go()