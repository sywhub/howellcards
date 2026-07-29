#!/usr/bin/env python3
# Generate a team match setup
#   A PDF with Roster and Score sheets
#   An Excel spreadsheet to enter the results and calculate the scores
#
# 4 Pairs (1, 2, 3, 4) are arranged into 2 teams (1 and 2, 3 and 4) and play each other.
#   They sit at 2 tables: {NS: Team 1, EW: 3} {NS: Team 4, EW: Team 2}.
#       Each table play n boards, then they exchange boards to play again.
#       Then they exchage the oppoenents:{NS: Team 1, EW: 4} {NS: Team 3, EW: 2}.
#       Play as before again.
#       Technically, that's 4 rounds.
#       The scoring diffreences are converted into IMPs and the one with higher IMP wins.
#
#
import argparse
import logging
import pdf
import datetime
from openpyxl import Workbook
from docset import PairGames
from maininit import setlog

class TeamMatch(PairGames):
    def __init__(self, log):
        super().__init__(log, 4)
        self.pdf = pdf.PDF()
        self.wb = Workbook()

    # record metadata
    def setup(self, boards, nameFile, fake):
        self.decks = boards
        self.fake = fake
        self.loadNames(nameFile, {'File': f'teammatchx{self.decks}{"xF" if self.fake else ""}',
                    'Tournament': f'Team Match, {self.decks} boards round'})
        self.pdf.HeaderFooterText(f'{self.notice} {datetime.date.today().strftime("%b %d, %Y")}.',
            self.nameObj['Tournament'])

        self.initData()
        try:
            self.checkBoardData()
        except ValueError as ex:
            if ex.args[0] != 'Same Pair':
                raise ex

    def pairN(self, n):
        return n

    def pairID(self, n):
        idStr = f'{self.pairN(n)}'
        if len(self.nameObj['Players']) > 0:
            idStr += f' ({self.nameObj['Players'][n-1]})'
        return idStr

    # Setup boardData and roundData for parent class methods
    def initData(self):
        setupData = {
            0: {0: {'NS': 1, 'EW': 4, 'Board': 0}, 1: {'NS': 3, 'EW': 2, 'Board': 1}},
            1: {0: {'NS': 1, 'EW': 4, 'Board': 1}, 1: {'NS': 3, 'EW': 2, 'Board': 0}},
            2: {0: {'NS': 1, 'EW': 3, 'Board': 2}, 1: {'NS': 4, 'EW': 2, 'Board': 3}},
            3: {0: {'NS': 1, 'EW': 3, 'Board': 3}, 1: {'NS': 4, 'EW': 2, 'Board': 2}}}

        for r in range(4):
            self.roundData[r] = {}
            for t,tbl in setupData[r].items():
                self.roundData[r][t] = {}
                for k,v in tbl.items():
                    self.roundData[r][t][k] = v
                self.roundData[r][t]['Board'] = self.boardList(self.roundData[r][t]['Board'])
        for r,t in self.roundData.items():
            for tbl,tData in self.roundData[r].items():
                for b in tData['Board']:
                    if b not in self.boardData:
                        self.boardData[b] = []
                    self.boardData[b].append([r, tbl, tData['NS'], tData['EW']])

    # Roster sheet
    # The roster tab also shows the tournament results
    def rosterSheet(self):
        ws = self.wb.active # the first tab
        ws.title = 'Roster'
        metaData = {'Title': self.nameObj['Tournament'],
                    'Info': [["Rounds", len(self.roundData)], ["Boards per Round", self.decks]]}
        row = self.sheetMeta(ws, metaData)
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 30
        row += 2
        nSum = len(self.boardData) * len(self.boardData[0]) - 1
        for t in range(2):
            ws.cell(row, 2).font = self.HeaderFont
            ws.cell(row, 2).alignment = self.centerAlign
            ws.cell(row, 2).value = f'Team {t+1}'
            ws.merge_cells(f'{ws.cell(row,2).coordinate}:{ws.cell(row,3).coordinate}')
            row += 1
            for p in range(2):
                ws.cell(row, 1).font = self.HeaderFont
                ws.cell(row, 1).alignment = self.centerAlign
                ws.cell(row, 1).value = 2 * t + p + 1
                ws.cell(row, 2).value = self.placeHolderName()
                ws.cell(row, 3).value = self.placeHolderName()
                row += 1
            for c in range(4):
                ws.cell(row-1,c+1).border = self.bottomLine
            ws.cell(row, 3).font = self.HeaderFont
            ws.cell(row, 3).alignment = self.centerAlign
            ws.cell(row, 3).value = 'IMP Sum'
            sum = f"=SUMIF('By Board'!{self.rc2a1(3,4)}:{self.rc2a1(3+nSum,4)},\"=\"&{self.rc2a1(row-2,1)},'By Board'!{self.rc2a1(3,12)}:{self.rc2a1(3+nSum,12)})"
            sum += f"+SUMIF('By Board'!{self.rc2a1(3,4)}:{self.rc2a1(3+nSum,4)},\"=\"&{self.rc2a1(row-1,1)},'By Board'!{self.rc2a1(3,12)}:{self.rc2a1(3+nSum,12)})"
            ws.cell(row, 4).value = sum
            ws.cell(row, 4).font = self.noChangeFont
            row += 2
            
    # simple sign-up sheet, PDF
    def rosterPDF(self):
        self.pdf.headerFooter()
        self.pdf.set_y(self.pdf.margin + self.pdf.lineHeight(self.pdf.font_size_pt) * 2)
        for t in range(2):
            self.pdf.set_font(style='BI', size=self.pdf.rosterPt, family=self.pdf.serifFont)
            self.pdf.cell(w=self.pdf.epw, text=f'Team {t+1}', align='C')
            self.pdf.set_y(self.pdf.get_y() + self.pdf.lineHeight(self.pdf.font_size_pt))

            self.pdf.set_font(style='', family=self.pdf.sansSerifFont)
            pw = self.pdf.get_string_width('Pair'+'8'*4) + 0.25
            nameW = (self.pdf.w - pw - 4 * self.pdf.margin) / 2
            ht = self.pdf.lineHeight(self.pdf.font_size_pt)
            for p in range(2):
                names = self.pairNames(t * 2 + p)
                self.pdf.set_x(self.pdf.margin*2)
                self.pdf.cell(w=pw, h=ht, text=f'Pair {t * 2 + p+1}', align='C', border=1)
                self.pdf.cell(w=nameW, h=ht, text=names[0], align='C', border=1)
                self.pdf.cell(w=nameW, h=ht, text=names[1], align='C', border=1)
                self.pdf.ln()
            self.pdf.set_y(self.pdf.get_y() + self.pdf.lineHeight(self.pdf.font_size_pt))

    def boardSheetHeaders(self, sh, nTbl):
        # first row setup some spanning column headers
        mergeHdrs = [['Score', 2], ['', 1], ['Net', 2]]

        headers = ['Board', 'Round', 'Table', 'NS', 'EW', 'Vul', 'Contract', 'By', 'Result', 'NS', 'EW', 'IMP', 'NS', 'EW']
        cStart = headers.index('Result') + 2
        for h in mergeHdrs:
            sh.cell(1, cStart).value = h[0]
            sh.cell(1, cStart).font = self.noChangeFont
            sh.cell(1, cStart).alignment = self.centerAlign
            sh.merge_cells(f'{sh.cell(1,cStart).coordinate}:{sh.cell(1,cStart+h[1]-1).coordinate}')
            cStart += h[1]
        row = self.headerRow(sh, headers, 2)
        return (row, headers)
        
    # Table of boards played, no PDF equivalent
    # Team matches are always IMP and 2 tables.  It always uses traveler.
    def Boards(self):
        self.log.debug('Saving by Board')
        sh = self.wb.create_sheet('By Board', 1)
        row, headers = self.boardSheetHeaders(sh, 2)
        for b in sorted(self.boardData.keys()):
            sh.cell(row, 1).value = b+1     # board #
            sh.cell(row, 1).alignment = self.centerAlign
            cursorRow = 0
            for r in sorted(self.boardData[b], key=lambda x: x[2]): # (round, table, NS, EW)
                sh.cell(row, 2).value = r[0]+1  # round
                sh.cell(row, 3).value = r[1]+1  # table
                sh.cell(row, 4).value = r[2]    # NS
                sh.cell(row, 5).value = r[3]    # EW
                sh.cell(row, 6).value = self.vulLookup(b)
                for i in range(2,7):
                    sh.cell(row, i).alignment = self.centerAlign

                cIdx = headers.index('Result')+3
                nIdx = cIdx + 2
                self.computeIMP(sh, cIdx, 2, row, cursorRow, nIdx, -1)  # put *here*
                self.computeNet(sh, row, cIdx-1, nIdx)
                row += 1
                cursorRow += 1
            if self.fake:
                self.fakeScore(sh, row-2, cIdx-1, 1.0)
                self.fakeScore(sh, row-1, cIdx-1, 1.0)
            for c in range(nIdx+1):
                sh.cell(row-1,c+1).border = self.bottomLine
        
    def VPTable(self):
        sh = self.wb['IMP Table']
        sh.cell(1, 4).value = 'VP'
        tau = (5**0.5 - 1)/2
        tau3 = 1 - tau**3
        vpb = 15 * len(self.boardData)**0.6
        for row in range(sh.max_row-1):
            sh.cell(row+2, 4).value = f"=10+10*(1-{tau}^(3*{self.rc2a1(row+2,3)}/{vpb}))/{tau3}"

    # Output into filesystem
    def save(self):
        import os
        here = os.path.dirname(os.path.abspath(__file__))
        fn = f'{here}/../{self.nameObj['File']}'
        self.wb.save(f'{fn}.xlsx')
        self.pdf.output(f'{fn}.pdf')
        print(f'Saved {fn}.{{xlsx,pdf}}')

    # Orchestrator
    def match(self):
        #self.pdf.instructions(self.log, "teaminstructions.txt")
        self.rosterSheet()
        self.rosterPDF()
        self.Boards()
        self.IMPTable()
        self.ScoreTable()
        self.VPTable()
        #self.idTags()
        #self.Travelers()  # PDF only
        self.Journal()  # pdf only
        self.save()
        return

if __name__ == '__main__':
    log = setlog('team', None)

    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--debug', type=str, default='INFO', help='Debug level, INFO, DEBUG, ERROR')
    parser.add_argument('-b', '--boards', type=int, choices=range(1,9), default=4, help='Number of boards per round')
    parser.add_argument('-n', '--names', type=str, default="", help='Names in the tournament')
    parser.add_argument('-f', '--fake', action='store_true', help='Fake scores to test the spreadsheet')
    args = parser.parse_args()
    for l in [['INFO', logging.INFO], ['DEBUG', logging.DEBUG], ['ERROR', logging.ERROR]]:
        if args.debug.upper() == l[0]:
            log.setLevel(l[1])
            break
    team = TeamMatch(log)
    # A match has n rounds, each round has m boards, divided into two halves, each half of the boards
    team.setup(boards=args.boards, nameFile=args.names, fake=args.fake)
    team.match()