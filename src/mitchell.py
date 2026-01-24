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
from docset import PairGames
import datetime

class Mitchell(PairGames):
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
        self.noChangeFont = Font(bold=True, italic=True, color='FF0000')
        self.pdf = pdf.PDF()
        self.wb = Workbook()

        notice = f'For public domain. No rights reserved. {datetime.date.today().strftime("%Y")}.'
        footer = f'Mitchell Tournament: {(self.pairs+1)//2} Tables, {self.boards} Boards per round'
        self.pdf.HeaderFooterText(notice, footer)
    
    def go(self):
        # the sequence of calls is important
        self.roster()
        self.roundTab() # build data structure for late
        self.boardTab()
        self.results()
        self.ScoreTable()
        self.Pickups()  # PDF only
        self.Tables()
        self.Travelers()  # PDF only
        self.Journal()  # PDF only
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
        ws.cell(row, 1).value = self.pdf.footerText
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
        self.pdf.headerFooter()
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
        calcs.append('Calculations')
        col = len(headers) + 1
        for h in calcs:
            sh.cell(row-1, col).value = h
            sh.cell(row-1, col).font = self.HeaderFont
            sh.cell(row-1, col).alignment = self.centerAlign
            col += 1
        for i in range(len(headers)+1, col+1):
            sh.cell(row-1,i).font = self.noChangeFont
        rGap = self.tables * self.boards
        for b in self.boardData.keys():
            sh.cell(row, 1).value = b+1
            sh.cell(row, 1).alignment = self.centerAlign
            cursorRow = 0
            for r in self.boardData[b]: # (round, table, NS, EW)
                nPlayed = len(self.boardData[b])    # # of times this board was played
                sh.cell(row, 2).value = f"='By Round'!{self.rc2a1(r[0] * rGap + 3, 1)}"
                tBase = r[0] * rGap + r[1] * self.boards + 3
                sh.cell(row, 3).value = f"='By Round'!{self.rc2a1(tBase, 2)}"
                sh.cell(row, 4).value = f"='By Round'!{self.rc2a1(tBase, 3)}"
                sh.cell(row, 5).value = f"='By Round'!{self.rc2a1(tBase, 4)}"
                cIdx = len(headers)
                rawNS = self.rc2a1(row, cIdx - 1)
                rawEW = self.rc2a1(row, cIdx)
                tBase += b % self.boards
                for i in range(6, len(headers)+1):
                    c = f"'By Round'!{self.rc2a1(tBase, i)}"
                    sh.cell(row, i).value = f'=IF(ISBLANK({c}),"",{c})'
                for i in range(2,7):
                    sh.cell(row, i).alignment = self.centerAlign

                # Computing MPs
                sh.cell(row, cIdx+1).value = f"={self.rc2a1(row, cIdx+3)}/{nPlayed-1}"
                sh.cell(row, cIdx+2).value = f"={self.rc2a1(row, cIdx+4)}/{nPlayed-1}"
                sh.cell(row, cIdx+1).number_format = sh.cell(row, cIdx+2).number_format = "0.00%"
                sh.cell(row, cIdx+3).value = f"=SUM({self.rc2a1(row, cIdx+7)}:{self.rc2a1(row,cIdx+5+nPlayed)})"
                sh.cell(row, cIdx+4).value = f"=SUM({self.rc2a1(row, cIdx+6+nPlayed)}:{self.rc2a1(row,cIdx+4+2*nPlayed)})"
                sh.cell(row, cIdx+3).number_format = sh.cell(row, cIdx+4).number_format = "#0.00"
                sh.cell(row, cIdx+5).value = f'=IF(ISNUMBER({rawNS}),{rawNS},IF(ISNUMBER({rawEW}),-{rawEW},""))'
                sh.cell(row, cIdx+6).value = f'=IF(ISNUMBER({rawEW}),{rawEW},IF(ISNUMBER({rawNS}),-{rawNS},""))'
                opponents = [x - cursorRow for x in range(nPlayed) if x != cursorRow]
                for i in range(2):
                    n = nPlayed - 1
                    for rCmp in range(n):
                        cmpF = f"=IF(AND(ISNUMBER({self.rc2a1(row, cIdx+5+i)}),ISNUMBER({self.rc2a1(row+opponents[rCmp], cIdx+5+i)})),"
                        cmpF += f"IF({self.rc2a1(row, cIdx+5+i)}>{self.rc2a1(row+opponents[rCmp], cIdx+5+i)},1,"
                        cmpF += f"IF({self.rc2a1(row, cIdx+5+i)}={self.rc2a1(row+opponents[rCmp], cIdx+5+i)},0.5,0)),0.5)"
                        targetC = cIdx+7+rCmp+i*n
                        sh.cell(row, targetC).value = cmpF

                row += 1
                cursorRow += 1
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
                        self.fakeScore(sh, row, 10)
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
        # rearrange by tables
        tables = {}
        for b,r in self.boardData.items():
            for v in r:
                if v[1] not in tables:
                    tables[v[1]] = {}
                if v[0] not in tables[v[1]]:
                    tables[v[1]][v[0]] = []
                tables[v[1]][v[0]].append({'NS': v[2], 'EW': v[3], 'Board': b})
        tblCols = []
        xMargin = 0.5
        hdrs = ['Board', 'NS', 'EW', 'Contract', 'By', 'Result', 'NS', 'EW']
        self.pdf.setHeaders(xMargin, hdrs, tblCols)
        bIdx = 0
        for t in sorted(tables.keys()):
            if bIdx % 4 == 0:
                self.pdf.add_page()
                self.pdf.headerFooter()
                y = 0.5
            for r in sorted(tables[t].keys()):
                y = self.pdf.headerRow(xMargin, y, tblCols, hdrs, f"Pickup: Table {t+1}, Round {r+1}")
                h = self.pdf.lineHeight(self.pdf.font_size_pt)
                self.pdf.set_font(size=self.pdf.linePt)
                y += h
                self.pdf.set_xy(xMargin, y)
                for b in tables[t][r]:
                    self.pdf.cell(tblCols[0], h, text=f'{b["Board"]+1}', align='C', border=1)
                    self.pdf.cell(tblCols[1], h, text=f'{b["NS"]}', align='C', border=1)
                    self.pdf.cell(tblCols[2], h, text=f'{b["EW"]}', align='C', border=1)
                    for c in range(3,len(hdrs)):
                        self.pdf.cell(tblCols[c], h, text='', align='C', border=1)
                    y += h
                    self.pdf.set_xy(xMargin, y)
                bIdx += 1
                y = self.pdf.sectionDivider(4, bIdx, xMargin)
        return

    def Tables(self):
        tables = {}
        for b,r in self.boardData.items():
            for v in r:
                if v[1] not in tables:
                    tables[v[1]] = {}
                if v[0] not in tables[v[1]]:
                    tables[v[1]][v[0]] = []
                tables[v[1]][v[0]].append({'NS': v[2], 'EW': v[3], 'Board': b})
        hdrs = ['Round', 'NS', 'EW', 'Boards']
        tblCols = []
        xMargin = 0.5
        self.pdf.setHeaders(xMargin, hdrs, tblCols)
        tblCols[3] = self.pdf.get_string_width('8'*8+','*3) + 0.25
        w = sum(tblCols)
        xMargin = (self.pdf.w - w) / 2

        for t in sorted(tables.keys()):
            self.pdf.add_page()
            self.pdf.movementSheet()
            self.pdf.compass()
            self.pdf.set_font(self.pdf.sansSerifFont, style='B', size=self.pdf.linePt)
            self.pdf.headerRow(xMargin, 2, tblCols, hdrs, f"Table {t+1}")
            self.pdf.set_font(size=self.pdf.linePt)
            y = self.pdf.get_y()
            h = self.pdf.lineHeight(self.pdf.font_size_pt);
            self.pdf.set_xy(xMargin, y + h)
            for r in sorted(tables[t].keys()):
                tRound = tables[t][r]
                self.pdf.cell(tblCols[0], h, text=f'{r+1}', align='C', border=1)
                self.pdf.cell(tblCols[1], h, text=f'{tables[t][r][0]['NS']}', align='C', border=1)
                self.pdf.cell(tblCols[2], h, text=f'{tables[t][r][0]['EW']}', align='C', border=1)
                bds = ""
                for b in tRound:
                    bds += f'{b['Board']+1},'
                self.pdf.cell(tblCols[3], h, text=bds[:-1], align='C', border=1)
                y += h
                self.pdf.set_xy(xMargin, y + h)



    def Travelers(self):
        tblCols = []
        xMargin = 0.5
        hdrs = ['Round', 'NS', 'EW', 'Contract', 'By', 'Result', 'NS', 'EW']
        self.pdf.setHeaders(xMargin, hdrs, tblCols)
        bIdx = 0
        for b,r in self.boardData.items():
            if bIdx % 4 == 0:
                self.pdf.add_page()
                self.pdf.headerFooter()
                y = 0.5
            y = self.pdf.headerRow(xMargin, y, tblCols, hdrs, f'Traveler for Board: {b+1}')
            h = self.pdf.lineHeight(self.pdf.font_size_pt)
            self.pdf.set_font(size=self.pdf.linePt)
            for v in r:
                y += h
                self.pdf.set_xy(xMargin, y)
                self.pdf.cell(tblCols[0], h, text=f'{v[0]+1}', align='C', border=1)
                self.pdf.cell(tblCols[1], h, text=f'{v[2]}', align='C', border=1)
                self.pdf.cell(tblCols[2], h, text=f'{v[3]}', align='C', border=1)
                for c in range(3,len(hdrs)):
                    self.pdf.cell(tblCols[c], h, text='', align='C', border=1)
            bIdx += 1
            y = self.pdf.sectionDivider(4, bIdx, xMargin)
        return

    def Journal(self):
        pairData = {}
        for b,r in self.boardData.items():
            for v in r:
                for p in range(2,4):
                    if v[p] not in pairData:
                        pairData[v[p]] = []
                    pairData[v[p]].append((v[0], b, v[1], v[2], v[3])) # (round, board, table, NS, EW)
        tblCols = []
        xMargin = self.pdf.margin * 2
        hdrs = ['Round', 'Board', 'NS', 'EW', 'Contract', 'By', 'Result', 'NS', 'EW']
        self.pdf.setHeaders(xMargin, hdrs, tblCols)
        self.pdf.set_xy(self.pdf.margin, self.pdf.margin)
        for p in sorted(pairData.keys()):
            self.pdf.add_page()
            self.pdf.headerFooter()
            y = self.pdf.headerRow(xMargin, self.pdf.margin, tblCols, hdrs ,f"Pair {p} Play Journal")
            self.pdf.set_font(size=self.pdf.linePt)
            h = self.pdf.lineHeight(self.pdf.font_size_pt)
            y += h;
            self.pdf.set_xy(xMargin, y)
            for v in sorted(pairData[p], key=lambda x: x[0]):
                self.pdf.cell(tblCols[0], h, text=f'{v[0]+1}', align='C', border=1)
                self.pdf.cell(tblCols[1], h, text=f'{v[1]+1}', align='C', border=1)
                self.pdf.cell(tblCols[2], h, text=f'{v[3]}', align='C', border=1)
                self.pdf.cell(tblCols[3], h, text=f'{v[4]}', align='C', border=1)
                for c in range(4,len(hdrs)):
                    self.pdf.cell(tblCols[c], h, text='', align='C', border=1)
                y += h
                self.pdf.set_xy(xMargin, y)
        return

    def results(self):
        sh = self.wb['Roster']
        nRows = 0
        for b in self.boardData.values():
            nRows += len(b)
        for i in range(self.pairs):
            ifRange = f"'By Board'!{self.rc2a1(3, 4+i%2)}:{self.rc2a1(3+nRows,4+i%2)}"
            sumRange = f"'By Board'!{self.rc2a1(3, 12+i%2)}:{self.rc2a1(3+nRows,12+i%2)}"
            ptsRange = f"'By Board'!{self.rc2a1(3, 12+i%2)}:{self.rc2a1(3+nRows,13+i%2)}"
            sh.cell(4+i,4).value=f"=SUMIF({ifRange},\"={i+1}\",{sumRange})/{len(self.boardData)}"
            sh.cell(4+i,5).value=f"=SUMIF({ifRange},\"={i+1}\",{ptsRange})"
            sh.cell(4+i,4).number_format = "0.00%"
            sh.cell(4+i,5).number_format = "#0.0"

    # Output into filesystem
    def save(self):
        import os
        here = os.path.dirname(os.path.abspath(__file__))
        fn = f'{here}/../mitchell{self.pairs}x{self.boards}'
        self.wb.save(f'{fn}.xlsx')
        self.pdf.output(f'{fn}.pdf')
        print(f'Saved {fn}.{{xlsx,pdf}}')


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