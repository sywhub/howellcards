#!/usr/bin/env python3
# Generate a team match setup
#   A PDF with Roster and Score sheets
#   An Excel spreadsheet to enter the results and calculate the scores
#
# This program arrange a tournament for 4 pairs to play in 1 to 3 rounds of "team matches".
# Each match is formally a match of 2 teams of 2 pairs.  At the end of each match, we change the composition
# of both teams.  In 3 matches, therefore, each pair has played with the other 2.
#
import argparse
import logging
from maininit import setlog
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
import datetime
from mitchell import Mitchell

class Square(Mitchell):
    def __init__(self, log, b, f):
        super().__init__(log, 8, b, f)

    def go(self):
        # the sequence of calls is important
        self.loadSetupSheet()
        self.roster()
        self.roundTab() # build data structure for late
        self.boardTab()
        self.results()
        self.ScoreTable()
        self.idTags()
        self.Pickups()  # PDF only
        self.Tables(True)
        self.Travelers()  # PDF only
        self.Journal()  # PDF only
        self.save()
        return

    def roundTab(self):
        self.roundByData(self.sqSetup)
    
    def roundByData(self, data):
        self.log.debug('Saving by Round')
        headers = ['Round', 'Table', 'NS', 'EW', 'Board', 'Vul', 'Contract', 'By', 'Result', 'NS', 'EW']
        sh, row = self.contractHeaders(headers, 'By Round', ['Scores'])
        roundData = {}
        self.boardData = {}
        for t,tData in data.items():
            for r in tData:
                if r['Round'] not in roundData:
                    roundData[r['Round']] = []
                roundData[r['Round']].append((t, r['NS'], r['EW'], r['Board']))
        for r in sorted(roundData.keys()):
            sh.cell(row, 1).value = r + 1
            for t in sorted(roundData[r], key=lambda x: x[0]):
                sh.cell(row, 2).value = t[0] + 1
                sh.cell(row, 3).value = t[1]
                sh.cell(row, 4).value = t[2]
                for j in range(1,5):
                    sh.cell(row, j).alignment = self.centerAlign
                for i in t[3]:
                    sh.cell(row, 5).value = i+1
                    sh.cell(row, 6).value = self.vulLookup(i)
                    sh.cell(row, 5).alignment = sh.cell(row, 6).alignment = self.centerAlign
                    if i not in self.boardData:
                        self.boardData[i] = []
                    if self.fake:
                        self.fakeScore(sh, row, 10)
                    self.boardData[i].append((r, t[0], t[1], t[2]))
                    row += 1


    def loadSetupSheet(self):
        fn = "SqMtichellx8.xlsx"
        setupWb = load_workbook(filename=fn)
        sh = setupWb.active
        self.sqSetup = {}
        row = 2
        while row <= sh.max_row:
            t = sh.cell(row, 1).value
            if t != None and t not in self.sqSetup:
                tbl = t - 1
                self.sqSetup[tbl] = []
            ns = sh.cell(row, 3).value
            ew = sh.cell(row, 4).value
            ns = (ns - 1) * 2 + 1
            ew *= 2
            bSet = sh.cell(row,5).value 
            b = [(bSet - 1)*self.boards + x for x in range(self.boards)]

            self.sqSetup[tbl].append({'Round': sh.cell(row, 2).value -1, 'NS': ns, 'EW': ew, 'Board': b})
            row += 1
        self.tables = len(self.sqSetup.keys())

    def idTags(self):
        idData = {}
        for t, d in self.sqSetup.items():
            for r in d:
                if r['NS'] not in idData:
                    idData[r['NS']] = []
                if r['EW'] not in idData:
                    idData[r['EW']] = []
                idData[r['NS']].append((r['Round'], t, r['NS'], r['EW'], r['Board']))
                idData[r['EW']].append((r['Round'], t, r['NS'], r['EW'], r['Board']))
        self.idTagsByData(idData)

    def idTagsByData(self, data):
        nTagsPage = len(data) if len(data) <= 4 else 4
        cHeight = self.pdf.eph / nTagsPage
        cWidth = self.pdf.w / 2
        leftMargin = self.pdf.margin * 2
        tags = 0
        colW = []
        hdrs = ['Round', 'Table', 'NS', 'EW']
        self.pdf.setHeaders(leftMargin, hdrs, colW)
        for id in sorted(data.keys()):
            if tags % nTagsPage == 0:
                self.pdf.add_page()
                x = leftMargin
                y = self.pdf.margin
            rData = sorted(data[id], key=lambda x: x[0])
            for half in range(2):
                self.pdf.set_font(self.pdf.serifFont, size=self.pdf.headerPt)
                self.pdf.set_xy(leftMargin+cWidth*half, y)
                self.pdf.cell(text=f"Pair {id}")
                h = self.pdf.lineHeight(self.pdf.font_size_pt)
                ty = y + h
                self.pdf.set_xy(leftMargin+cWidth*half, ty)
                self.pdf.set_font(self.pdf.sansSerifFont, style='B', size=8)
                h = self.pdf.lineHeight(self.pdf.font_size_pt)
                for i in range(len(hdrs)):
                    self.pdf.cell(colW[i], h, text=hdrs[i], align='C', border=1)
                ty +=  h
                for r in rData:
                    self.pdf.set_xy(leftMargin+cWidth*half, ty)
                    for i in range(len(hdrs)):
                        self.pdf.cell(colW[i], h, text=f"{r[i]+(1 if i <= 1 else 0)}", align='C', border=1)
                    ty += h
            y += cHeight
            tags += 1
        return
           
    # Output into filesystem
    def save(self):
        import os
        here = os.path.dirname(os.path.abspath(__file__))
        fn = f'{here}/../squarex{self.boards}'
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
    parser.add_argument('-b', '--boards', type=int, choices=range(2,7), default=3, help='Boards per round')
    parser.add_argument('-f', '--fake', type=bool, default=False, help='Fake scores to test the spreadsheet')
    args = parser.parse_args()
    for l in [['INFO', logging.INFO], ['DEBUG', logging.DEBUG], ['ERROR', logging.ERROR]]:
        if args.debug.upper() == l[0]:
            log.setLevel(l[1])
            break
    sq = Square(log, args.boards, args.fake)
    # A match has n rounds, each round has m boards, divided into two halves, each half of the boards
    sq.go()