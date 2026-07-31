#!/usr/bin/env python3
# Generate a team match setup
#   A PDF with Roster and Score sheets
#   An Excel spreadsheet to enter the results and calculate the scores
#
#   4 pairs form 2 teams.  Team 1 = pair 1 & 2, Team 2 = pair 3 & 4
#   They play two matches.  Each two rounds of "decks" boards.
#   For each round, table 1 play a set of boards and table 2 the other
#   After ward, playes stay at the same table. And play the boards the other table just played.
#   That's a "match".
#
#   Second match, players change table to meet the "other pair" of the other taam.
#   Repeat as above.  That's the 2nd and final match.
#   
#   Minimal boards to play is 4, meximal is 32.  Remember, always 4 rounds or 2 matches.
#
import argparse
import logging
import pdf
import datetime
from openpyxl import Workbook
from openpyxl.styles import Border
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
        self.roundData = {
            0: {0: {'NS': 1, 'EW': 4, 'Board': self.boardList(0)}, 1: {'NS': 3, 'EW': 2, 'Board': self.boardList(1)}},
            1: {0: {'NS': 1, 'EW': 4, 'Board': self.boardList(1)}, 1: {'NS': 3, 'EW': 2, 'Board': self.boardList(0)}},
            2: {0: {'NS': 1, 'EW': 3, 'Board': self.boardList(2)}, 1: {'NS': 4, 'EW': 2, 'Board': self.boardList(3)}},
            3: {0: {'NS': 1, 'EW': 3, 'Board': self.boardList(3)}, 1: {'NS': 4, 'EW': 2, 'Board': self.boardList(2)}}}

        for r,t in self.roundData.items():
            for tbl,tData in t.items():
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
                    'Info': [["Matches", len(self.roundData) // 2], ["Boards per Match", self.decks*2]]}
        row = self.sheetMeta(ws, metaData)
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 30
        ws.column_dimensions['D'].width = 30
        ws.column_dimensions['E'].width = 30
        row += 2
        nSum = len(self.boardData) * len(self.boardData[0]) # number of rows to pick from
        tRows = []
        for t in range(2):  # Team 1 & 2
            ws.cell(row, 2).font = self.HeaderFont
            ws.cell(row, 2).alignment = self.centerAlign
            ws.cell(row, 2).value = f'Team {t+1}'
            ws.merge_cells(f'{ws.cell(row,2).coordinate}:{ws.cell(row,3).coordinate}')
            row += 1
            for p in range(2):  # Pair n & n+1 for each team
                names = self.pairNames(t * 2 + p)   # load from "--name" command line option, if provided
                ws.cell(row, 1).font = self.HeaderFont
                ws.cell(row, 1).alignment = self.centerAlign
                ws.cell(row, 1).value = 2 * t + p + 1
                ws.cell(row, 2).value = names[0]
                ws.cell(row, 3).value = names[1]
                tRows.append(row)   # remember which pairs are for which team
                row += 1
            for c in range(3):
                ws.cell(row-1,c+1).border = self.bottomLine
            row += 2
        # Display tournament results
        # First the headers
        ws.cell(row, 2).font = self.HeaderFont
        ws.cell(row, 2).alignment = self.centerAlign
        ws.cell(row, 2).value = 'IMP Summary'
        ws.merge_cells(f'{ws.cell(row,2).coordinate}:{ws.cell(row,3).coordinate}')
        ws.cell(row, 4).font = self.HeaderFont
        ws.cell(row, 4).alignment = self.centerAlign
        ws.cell(row, 4).value = 'VP Summary'
        ws.merge_cells(f'{ws.cell(row,4).coordinate}:{ws.cell(row,5).coordinate}')
        row += 1
        i = 1
        for h in ['Match'] + ['Team 1', 'Team 2'] * 2:
            ws.cell(row, i).font = self.HeaderFont
            ws.cell(row, i).alignment = self.centerAlign
            ws.cell(row, i).value = h
            i += 1

        # Now compute VP sums and use them to compute VPs
        row += 1
        nSum //= 2
        rStart = 3
        for m in range(2):
            ws.cell(row, 1).value = m+1
            for t in range(2):
                # IMP for each "match"
                # Team match needs only to sum the N-S pair's IMPs, the other table are their teammates
                # The pairs of the team took turn sitting at N-S position.  We remembered their rounds above
                sum =  f"=SUMIF('By Board'!{self.rc2a1(rStart,4)}:{self.rc2a1(rStart+nSum-1,4)},\"=\"&{self.rc2a1(tRows[2*t],1)},'By Board'!{self.rc2a1(rStart,13)}:{self.rc2a1(rStart+nSum-1,13)})"
                sum += f"+SUMIF('By Board'!{self.rc2a1(rStart,4)}:{self.rc2a1(rStart+nSum-1,4)},\"=\"&{self.rc2a1(tRows[2*t+1],1)},'By Board'!{self.rc2a1(rStart,13)}:{self.rc2a1(rStart+nSum-1,13)})"
                # Compute the VP for the winning side.  The losing side gets the remainder
                opp = [5,4][t]
                vp = f"=IF({self.rc2a1(row, t+2)}>=0,{self.VPFormula(self.rc2a1(row, t+2),nSum//2)},20-{self.rc2a1(row,opp)})"
                ws.cell(row, t+2).value = sum
                ws.cell(row, t+2).font = self.noChangeFont
                ws.cell(row, t+2).number_format = "#0.0"
                ws.cell(row, t+4).value = vp
                ws.cell(row, t+4).font = self.noChangeFont
                ws.cell(row, t+4).number_format = "#0.0"
            rStart += nSum
            row += 1
        for c in range(5):
            ws.cell(row-1,c+1).border = self.bottomLine
        # Add up both IMP and VP
        for c in range(4):
            ws.cell(row, 2+c).value=f'=SUM({self.rc2a1(row-2,2+c)}:{self.rc2a1(row-1,2+c)})'
            ws.cell(row, 2+c).font = self.noChangeFont
            ws.cell(row, 2+c).number_format = "#0.0"
                
    # simple sign-up sheet, PDF
    # First the roster names
    # Then the table information.
    def rosterPDF(self):
        self.pdf.headerFooter()
        self.pdf.set_y(self.pdf.margin + self.pdf.lineHeight(self.pdf.font_size_pt) * 4)
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
        self.pdf.set_font_size(self.pdf.bigPt)
        ht = self.pdf.lineHeight(self.pdf.font_size_pt)
        xMargin = self.pdf.margin
        hdrs = ['Round', 'NS', 'EW', 'Boards']
        tblCols=[]
        self.pdf.setHeaders(xMargin, hdrs, tblCols)
        tblCols[3] = self.pdf.get_string_width('8'*self.decks*2+','*(self.decks-1))+0.25
        xMargin = (self.pdf.w - sum(tblCols)) / 2
        for t in range(2):
            self.pdf.set_x(xMargin)
            self.pdf.set_font(style='BI')
            self.pdf.cell(h=ht, text=f'Table {t+1}')
            self.pdf.headerRow(xMargin, self.pdf.get_y(), tblCols, hdrs)
            self.pdf.set_font(style='')
            self.pdf.ln()
            for r in sorted(self.roundData.keys()):
                pos = self.roundData[r][t]
                self.pdf.set_x(xMargin)
                self.pdf.cell(w=tblCols[0], h=ht, text=f'{r+1}', align='C', border=1)
                self.pdf.cell(w=tblCols[1], h=ht, text=f'{pos['NS']}', align='C', border=1)
                self.pdf.cell(w=tblCols[2], h=ht, text=f'{pos['EW']}', align='C', border=1)
                bds = ','.join([str(x+1) for x in pos['Board']])
                self.pdf.cell(w=tblCols[3], h=ht, text=bds, align='C', border=1)
                self.pdf.ln()
            self.pdf.ln()

    # Shreadsheet header
    def boardSheetHeaders(self, sh, nTbl):
        # first row setup some spanning column headers
        mergeHdrs = [['Score', 2], ['', 1], ['Net', 2]]

        headers = ['Board', 'Round', 'Table', 'NS', 'EW', 'Vul', 'Contract', 'By', 'Made', 'Down', 'NS', 'EW', 'IMP', 'NS', 'EW']
        cStart = headers.index('Down') + 2
        for h in mergeHdrs:
            sh.cell(1, cStart).value = h[0]
            sh.cell(1, cStart).font = self.noChangeFont
            sh.cell(1, cStart).alignment = self.centerAlign
            sh.merge_cells(f'{sh.cell(1,cStart).coordinate}:{sh.cell(1,cStart+h[1]-1).coordinate}')
            cStart += h[1]
        row = self.headerRow(sh, headers, 2)
        return (row, headers)
        
    # Table of boards played, no PDF equivalent
    # Team matches are always IMP and 2 tables.
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

                cIdx = headers.index('Down')+3
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
        for c in [11, 13,14]:
            for r in range(2,sh.max_row+1):
                bd = sh.cell(r, c).border
                sh.cell(r, c).border = Border(left=self.thinLine, bottom=bd.bottom)

    # Generate a string that's an Excel VP formula
    def VPFormula(self, impCell, bNum):
        tau = (5.0**.5 - 1)/2
        vpb = 15*(bNum**.5)
        vpF = f'(10+10*((1-{tau}^(3*{impCell}/{vpb}))/(1-{tau}^3)))'
        vp = f'IF({vpF}>=20,20,{vpF})'
        return vp

    # Directly compute VP
    def VPCompute(self, impDiff, bNum):
        tau = (5.0**.5 - 1)/2
        vpb = 15*(bNum**.5)
        vp = 10+10*((1-tau**(3*impDiff/vpb))/(1-tau**3))
        if vp >= 20:
            vp = 20
        return vp

    # Only for team matches, generate Victory Point scale table
    # The roster actually does not use it.  It's for human references.
    def VPTable(self):
        sh = self.wb.create_sheet('VP Table')
        sh.cell(1, 2).value = '20 Victory Point Scales'
        sh.cell(1, 2).font = self.HeaderFont
        sh.cell(1, 2).alignment = self.centerAlign
        sh.merge_cells(f'{self.rc2a1(1,2)}:{self.rc2a1(1,10)}')
        row = 2
        col = 1
        sh.cell(row, col).value = 'IMP'
        for b in range(8,17):
            col += 1
            sh.cell(row, col).value = b
        for c in range(col):
            sh.cell(row, c+1).font = self.HeaderFont
            sh.cell(row, c+1).alignment = self.centerAlign
        row += 1
        col = 1
        for imp in range(61):
            sh.cell(row, col).value = imp
            for b in range(8, 17):
                col += 1
                sh.cell(row, col).value = self.VPCompute(imp, b)
                sh.cell(row, col).number_format = "#0.00"
            row += 1
            col = 1
        
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
        self.VPTable()
        self.ScoreTable()
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