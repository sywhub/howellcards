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
from openpyxl.styles import Font, Alignment, Border, Side
import pdf
from docset import DupBridge
import datetime

class TeamMatch(DupBridge):
    def __init__(self, log):
        super().__init__(log)
        self.pdf = pdf.PDF()
        self.wb = Workbook()
        self.boards = 4
        self.matches = 0
        self.rounds = 2
        self.vulTbl = ['None', 'NS', 'EW', 'Both']
        self.TeamPairs = {1: [1, 2, 3, 4], 2: [1, 3, 2, 4], 3: [1, 4, 2, 3]}
        self.Fake = True

    # record metadata
    def setup(self, rounds, boards, match, fake):
        self.boards = boards
        self.matches = match
        self.rounds = rounds
        self.Fake = fake

    # Roster sheet
    # The roster tab also shows the tournament results
    def Roster(self):
        ws = self.wb.active # the first tab
        ws.title = 'Roster'
        # First simple list of names
        row = 1
        ws.cell(row, 1).value = 'Players'
        ws.cell(row, 1).font = self.HeaderFont
        ws.cell(row, 1).alignment = self.centerAlign
        ws.merge_cells(f'{ws.cell(row,1).coordinate}:{ws.cell(row,3).coordinate}')
        ws.cell(row, 4).value = 'IMP'
        ws.cell(row, 4).font = self.HeaderFont
        ws.cell(row, 4).alignment = self.centerAlign
        row += 1
        for pair in range(4):
            ws.cell(row, 1).font = self.HeaderFont
            ws.cell(row, 1).value = f'Pair {pair+1}'
            ws.cell(row, 2).value = self.placeHolderName()
            ws.cell(row, 3).value = self.placeHolderName()
            row += 1
        row += 1
        
        # The sum of all earned IMP for each pair. The pair ranking.
        for pair in range(4):
            sumImp = "=+"
            for s in range(self.matches):
                seat = self.TeamPairs[s+1].index(pair+1) // 2
                if s > 0 and s < self.matches:
                    sumImp += '+'
                sumImp += self.rc2a1(9+s*3, 5+seat)
            ws.cell(2+pair, 4).value = sumImp

        matchIdx = 1
        while self.matches >= matchIdx:
            ws.cell(row, 1).font = self.HeaderFont
            ws.cell(row, 1).alignment = self.centerAlign
            ws.cell(row, 1).value = f'Match #{matchIdx}'
            ws.merge_cells(f'{ws.cell(row,1).coordinate}:{ws.cell(row,6).coordinate}')
            row += 1

            # Each match is a competition of 2 teams.
            # The team composition changes for each match.
            for t in range(2):
                ws.cell(row, 2*t+1).value = f'Team {t+(matchIdx-1)*2+1}'
                ws.cell(row, 2*t+1).font = self.HeaderFont
                ws.cell(row, 2*t+1).alignment = self.centerAlign
                ws.merge_cells(f'{ws.cell(row,2*t+1).coordinate}:{ws.cell(row,2*t+2).coordinate}')
            # The IMP for each team
            ws.cell(row, 5).font = self.HeaderFont
            ws.cell(row, 5).alignment = self.centerAlign
            ws.cell(row, 5).value = 'IMP'
            ws.merge_cells(f'{ws.cell(row,5).coordinate}:{ws.cell(row,6).coordinate}')
            row += 1
            idx = matchIdx % 3
            if idx == 0:
                idx = 3
            for i in range(4):
                ws.cell(row, i+1).value = f'Pair {self.TeamPairs[idx][i]}'
            # The sum of IMP won/lost with each board these two team played during the match
            # Which is n rounds of m boards, each board played twice on each table
            nRows = self.rounds * self.boards * 2   # rows for each match
            rStart = 3 + (matchIdx - 1) * nRows # the starting board row
            rEnd = rStart + nRows - 1
            # Excel SUMIF function 
            formula = f'=SUMIF(Boards!C{rStart}:C{rEnd},{self.rc2a1(row,1)},Boards!K{rStart}:K{rEnd})'
            formula +=f'+SUMIF(Boards!C{rStart}:C{rEnd},{self.rc2a1(row,2)},Boards!K{rStart}:K{rEnd})'
            ws.cell(row, 5).value = formula
            formula = f'=SUMIF(Boards!D{rStart}:D{rEnd},{self.rc2a1(row,3)},Boards!L{rStart}:L{rEnd})'
            formula +=f'+SUMIF(Boards!D{rStart}:D{rEnd},{self.rc2a1(row,4)},Boards!L{rStart}:L{rEnd})'
            ws.cell(row, 6).value = formula
            matchIdx += 1
            row += 1

    # simple sign-up sheet, PDF
    def Signup(self):
        headers = {'Team': self.pdf.epw * 0.05, 'Pairs': self.pdf.epw * 0.1, 'Name': self.pdf.epw * 0.7}
        self.pdf.add_page()
        self.headerFooter()
        self.pdf.set_font(style='BI', size=self.pdf.rosterPt, family=self.pdf.serifFont)
        self.pdf.set_y(self.pdf.margin + self.pdf.lineHeight(self.pdf.font_size_pt) * 2)
        self.pdf.cell(w=self.pdf.epw, text='Pair Signup', align='C')
        self.pdf.set_y(self.pdf.get_y() + self.pdf.lineHeight(self.pdf.font_size_pt) * 2)

        for p in range(4):
            self.pdf.set_font(style='', size=self.pdf.linePt, family=self.pdf.sansSerifFont)
            self.pdf.set_x(self.pdf.margin+headers['Team'])
            self.pdf.cell(w=headers['Pairs'], h=self.pdf.lineHeight(self.pdf.font_size_pt), text=f'Pair {p+1}', align='C', border=1)
            self.pdf.cell(w=headers['Name']/2, h=self.pdf.lineHeight(self.pdf.font_size_pt), text='', border=1)
            self.pdf.cell(w=headers['Name']/2, h=self.pdf.lineHeight(self.pdf.font_size_pt), text='', border=1)
            self.pdf.ln()

    # a card for each pair for navigation and general info
    # Since these are exactly 4 pairs, we fit them into one page
    def PairCard(self):
        tblHeaders = ['Round', 'Table', 'NS', 'EW', 'Boards']
        cols = []
        self.pdf.setHeaders(self.pdf.margin, tblHeaders, cols)
        cols[4] = max(self.pdf.get_string_width(f"{'8'*6}{','*2}")+0.2,cols[4]) # make it wide enough to fit 3 boards
        w = sum(cols)
        self.pdf.add_page()
        self.headerFooter()
        for p in range(4):
            if p % 2 == 0:
                xStart = (self.pdf.w - 2*w) / 3
            else:
                xStart += xStart + w
            if p // 2 == 0:
                yStart = 0.5
            else:
                yStart = self.pdf.eph / 2 + 0.5
            y = self.pdf.headerRow(xStart, yStart, cols, tblHeaders, f"Pair {p+1} Play Info")
            self.pdf.set_font(size=self.pdf.linePt)
            h = self.pdf.lineHeight(self.pdf.font_size_pt);
            for r in range(0, len(self.boardData), self.boards // 2):
                y += h
                rb = r // (self.boards // 2)
                self.pdf.set_xy(xStart, y)
                self.pdf.cell(cols[0], h, text=f'{rb+1}', align='C', border=1)
                t = 0 if (p+1) in self.boardData[r]['Tables'][0] else 1
                self.pdf.cell(cols[1], h, text=f'{t+1}', align='C', border=1)
                self.pdf.cell(cols[2], h, text=f'{self.boardData[r]['Tables'][t][0]}', align='C', border=1)
                self.pdf.cell(cols[3], h, text=f'{self.boardData[r]['Tables'][t][1]}', align='C', border=1)
                bds = self.boardSet(t, r)
                txt = ""
                for b in bds:
                    txt += f"{b+1},"
                self.pdf.cell(cols[4], h, text=txt[:-1], align='C', border=1)

    # board #s for each half of round
    def boardSet(self, t, r):
        bds = []
        if t == 0:
            bN = r
        else:
            bN = r + self.boards // 2 + 1
            if (bN % self.boards) == 1:
                bN -= self.boards
            bN -= 1
        for b in range(self.boards//2):
            bds.append(bN + b)
        return bds

    # score keeping for each pair
    def Journal(self):
        hdrs = ['Board', 'Table', 'NS', 'EW', 'Vul', 'Contract', 'By', 'Result', 'NS', 'EW']
        cols = []
        self.pdf.setHeaders(self.pdf.margin, hdrs, cols)
        w = sum(cols)
        leftMargin = (self.pdf.w - w) / 2
        for p in range(4):
            self.pdf.add_page()
            self.headerFooter()
            y = self.pdf.headerRow(leftMargin, self.pdf.margin, cols, hdrs, f"Pair {p+1} Journal")
            self.pdf.set_font(size=self.pdf.linePt)
            h = self.pdf.lineHeight(self.pdf.font_size_pt);
            for r in range(len(self.boardData)):
                y += h
                self.pdf.set_xy(leftMargin, y)
                self.pdf.cell(cols[0], h, text=f'{r+1}', align='C', border=1)
                t = 0 if (p+1) in self.boardData[r]['Tables'][0] else 1
                self.pdf.cell(cols[1], h, text=f'{t+1}', align='C', border=1)
                self.pdf.cell(cols[2], h, text=f'{self.boardData[r]['Tables'][t][0]}', align='C', border=1)
                self.pdf.cell(cols[3], h, text=f'{self.boardData[r]['Tables'][t][1]}', align='C', border=1)
                vulIdx = (r + r // 4) % 4
                self.pdf.set_font_size(self.pdf.notePt - 2);
                self.pdf.cell(cols[4], h, text=f'{self.vulTbl[vulIdx]}', align='C', border=1)
                self.pdf.set_font_size(self.pdf.linePt);
                for i in range(5,len(hdrs)):
                    self.pdf.cell(cols[i], h, text='', align='C', border=1)

        return

    # Table of boards played, no PDF equivalent
    def Boards(self):
        self.wb.create_sheet('Boards')
        ws = self.wb['Boards']
        # merged columns
        headers = ['Board', 'Table', 'NS', 'EW', 'Vul', 'Contract', 'By', 'Result',
                   'NS', 'EW',  # raw score from play
                   'NS', 'EW',  # IMP columns
                   'NS', 'EW',  # Diffs
                   'NS', 'EW',  # Net scores
                   ]
        scoreCol = headers.index('NS', 4)   # raw scores, zero based
        col = scoreCol + 1
        for h in ['Score', 'IMP', 'Diff', 'Net Score']:
            ws.cell(1, col).value = h
            ws.cell(1, col).font = self.HeaderFont
            ws.cell(1, col).alignment = self.centerAlign
            ws.merge_cells(f'{ws.cell(1,col).coordinate}:{ws.cell(1,col+1).coordinate}')
            col += 2

        # next row is the headers
        row = self.headerRow(ws, headers, 2)
        pairCol = headers.index('NS')+1
        self.boardData = []
        for board in range(self.boards * self.rounds * self.matches):
            self.boardData.append({})
            col = 1
            ws.cell(row, col).value = board+1
            vulIdx = (board + board // 4) % 4
            self.boardData[board]['Vul'] = vulIdx
            self.boardData[board]['Tables'] = []
            matchIdx = board // (self.rounds * self.boards) + 1
            # change opponents for each table
            swap = (board % self.boards) // (self.boards // 2)   # 1st or 2nd half of the round
            for t in range(2):  # 2 tables
                if self.Fake:
                    self.fakeContracts(ws, row)
                ws.cell(row, col+headers.index('Vul')).value = f'{self.vulTbl[vulIdx]}'
                ws.cell(row, col+headers.index('Vul')).alignment = self.centerAlign
                ws.cell(row, col+headers.index('Table')).value = t+1

                # IMP Columns
                lookCell = self.rc2a1(row, scoreCol+5)
                lookUp =f"=IF(ISNUMBER({lookCell}),VLOOKUP(ABS({lookCell}),'IMP Table'!$A$2:$C$26,3)*SIGN({lookCell}),\"\")"
                ws.cell(row, scoreCol+3).value = lookUp
                lookCell = self.rc2a1(row, scoreCol+6)
                lookUp =f"=IF(ISNUMBER({lookCell}),VLOOKUP(ABS({lookCell}),'IMP Table'!$A$2:$C$26,3)*SIGN({lookCell}),\"\")"
                ws.cell(row, scoreCol+4).value = lookUp

                # Set which pair against which
                # Compute the differences of the net scores
                if t == 0:
                    ws.cell(row, pairCol).value = f'Pair {self.TeamPairs[matchIdx][0]}'
                    ws.cell(row, pairCol+1).value = f'Pair {self.TeamPairs[matchIdx][2 + swap]}'
                    checkNum = f'=IF(AND(ISNUMBER({self.rc2a1(row, scoreCol+7)}),ISNUMBER({self.rc2a1(row+1, scoreCol+7)})),'
                    ws.cell(row, scoreCol+5).value = f'{checkNum}{self.rc2a1(row, scoreCol+7)}-{self.rc2a1(row+1, scoreCol+7)},"")'
                    checkNum = f'=IF(AND(ISNUMBER({self.rc2a1(row, scoreCol+8)}),ISNUMBER({self.rc2a1(row+1, scoreCol+8)})),'
                    ws.cell(row, scoreCol+6).value = f'{checkNum}{self.rc2a1(row, scoreCol+8)}-{self.rc2a1(row+1, scoreCol+8)},"")'
                else:
                    ws.cell(row, pairCol).value = f'Pair {self.TeamPairs[matchIdx][3 - swap]}'
                    ws.cell(row, pairCol+1).value = f'Pair {self.TeamPairs[matchIdx][1]}'
                    checkNum = f'=IF(AND(ISNUMBER({self.rc2a1(row, scoreCol+7)}),ISNUMBER({self.rc2a1(row-1, scoreCol+7)})),'
                    ws.cell(row, scoreCol+5).value = f'{checkNum}{self.rc2a1(row, scoreCol+7)}-{self.rc2a1(row-1, scoreCol+7)},"")'
                    checkNum = f'=IF(AND(ISNUMBER({self.rc2a1(row, scoreCol+8)}),ISNUMBER({self.rc2a1(row-1, scoreCol+8)})),'
                    ws.cell(row, scoreCol+6).value = f'{checkNum}{self.rc2a1(row, scoreCol+8)}-{self.rc2a1(row-1, scoreCol+8)},"")'
                self.boardData[board]['Tables'].append((int(ws.cell(row, pairCol).value[-1]),
                                                        int(ws.cell(row, pairCol+1).value[-1])))
                # setup net scores
                NSraw=self.rc2a1(row, scoreCol+1)
                EWraw=self.rc2a1(row, scoreCol+2)
                ws.cell(row, scoreCol+7).value = f'=IF(ISNUMBER({NSraw}),{NSraw},IF(ISNUMBER({EWraw}),-{EWraw},""))'
                ws.cell(row, scoreCol+8).value = f'=IF(ISNUMBER({EWraw}),{EWraw},IF(ISNUMBER({NSraw}),-{NSraw},""))'
                row += 1
            bd = Side(style='thin', color='000000')
            for i in range(1,len(headers)+1):
                ws.cell(row-1, i).border = Border(bottom=bd)
        # hint that this section is not to touch
        bd = Side(style='thin', color='f08000')
        for board in range(self.boards * self.rounds * self.matches*2+2):
            bexist = ws.cell(board+1, 13).border 
            ws.cell(board+1, 13).border = Border(left=bd, bottom=bexist.bottom)
        

    # Fake a score to test IMP calculation
    def fakeContracts(self, ws, row):
        import random
        col = random.choice([9,10])
        ws.cell(row, col).value = random.randint(5,12)*100

    # Paper scoring sheet for each table
    # Excel tabs are not really needed.
    def scoreSheet(self, headers):
        for t in range(2):
            self.wb.create_sheet(f'Table{t+1} Scores')
            row = 1
            startBoard = 1
            for r in range(1, self.Rounds + 1):
                ws = self.wb[f'Table{t+1} Scores']
                row = self.headerRow(ws, [f'Table {t+1}, Round {r}'], row)
                ws.merge_cells(f'{ws.cell(row-1,1).coordinate}:{ws.cell(row-1,len(headers)+1).coordinate}')
                row = self.headerRow(ws, headers, row)
                for board in range(startBoard, self.boardPerSwitch * self.matchPerRound + startBoard):
                    ws.cell(row, 1).value = board
                    row += 1
                startBoard = board + 1
                row += 1

    def Score(self):
        headers = {'Board': .1, 'Contract': .4, 'By': .15, 'Result': .25}
        headers = {k: v * self.pdf.epw for k,v in headers.items()}
        self.scoreSheet(list(headers.keys()))
        
        # How many rounds per page?
        boardsPerRound = self.matches * self.boards
        if boardsPerRound >= 8:
            roundPerPage = 2
        elif boardsPerRound <= 4:
            roundPerPage = 4
        else:
            roundPerPage = self.rounds
        roundHeight = self.pdf.eph / roundPerPage
        for t in range(2):
            startBoard = 1
            for r in range(1, self.Rounds + 1):
                if r % roundPerPage == 1:
                    self.pdf.add_page()
                    self.headerFooter()
                self.pdf.set_y(self.pdf.margin + (r % roundPerPage - 1) * roundHeight)
                self.pdf.set_font(style='B', size=self.pdf.bigPt, family=self.pdf.serifFont)
                self.pdf.cell(text=f'Table {t+1}, Round {r}')
                if r > 1:
                    x = self.pdf.get_x()
                    ydiff = self.pdf.get_y() + self.pdf.lineHeight(self.pdf.font_size_pt)
                    self.pdf.set_font(style='B', size=self.pdf.linePt, family=self.pdf.serifFont)
                    ydiff -= self.pdf.lineHeight(self.pdf.font_size_pt*1.25)
                    self.pdf.set_xy(x, ydiff)
                    self.pdf.cell(text='(EW Change Table)')
                self.pdf.set_y(self.pdf.get_y() + self.pdf.lineHeight(self.pdf.font_size_pt)*.5)
                self.pdf.set_font(style='B', size=self.pdf.headerPt, family=self.pdf.sansSerifFont)
                self.pdf.set_y(self.pdf.get_y() + self.pdf.lineHeight(self.pdf.font_size_pt))
                for h,pdfW in headers.items():
                    self.pdf.cell(w=pdfW, h=self.pdf.lineHeight(self.pdf.font_size_pt), text=h, align='C', border=1)
                self.pdf.ln()
                self.pdf.set_font(style='', size=self.pdf.linePt)

                self.pdf.set_font(size=self.pdf.linePt, family=self.pdf.sansSerifFont)
                for board in range(startBoard, self.boardPerSwitch * self.matchPerRound + startBoard):
                    self.pdf.cell(w=headers['Board'], h=self.pdf.lineHeight(self.pdf.font_size_pt), text=f'{board}', align='C', border=1)
                    for col in [v for k,v in headers.items() if k != 'Board']:
                        self.pdf.cell(w=col, h=self.pdf.lineHeight(self.pdf.font_size_pt), text='', border=1)
                    self.pdf.ln()

    # Output into filesystem
    def save(self):
        import os
        here = os.path.dirname(os.path.abspath(__file__))
        fn = f'{here}/../teammatch{self.rounds}x{self.boards}x{self.matches}'
        self.wb.save(f'{fn}.xlsx')
        self.pdf.output(f'{fn}.pdf')
        print(f'Saved {fn}.{{xlsx,pdf}}')

    # Some text for the TD/Organizer
    def Instructions(self):
        tourneyMeta = [['Rounds',self.rounds * 2], ['Boards per round',self.boards // 2], ['Number of Matches', self.matches]]
        self.pdf.meta(None, "Team Match", tourneyMeta)
        self.pdf.instructions(None, "teaminstructions.txt")

    def headerFooter(self):
        notice = f'For public domain. No rights reserved. {datetime.date.today().strftime("%Y")}.'
        footer = f'{f"{self.matches} Matches of " if self.matches > 1 else ""}{self.rounds * 2} {self.boards // 2}-Boards Rounds '
        self.pdf.set_font(size=self.pdf.tinyPt)
        h = self.pdf.lineHeight(self.pdf.font_size_pt)
        w = self.pdf.get_string_width(notice)
        x = self.pdf.setHCenter(w)
        self.pdf.set_xy(x, h)
        self.pdf.cell(text=notice)
        w = self.pdf.get_string_width(footer)
        x = self.pdf.setHCenter(w)
        y = self.pdf.eph - h * 2
        self.pdf.set_xy(x, y)
        self.pdf.cell(text=footer)

    def rc2a1(self, r, c):
        return f"{chr(c-1+ord('A'))}{r}"

    # Orchestrator
    def match(self):
        self.Instructions()
        self.Signup()
        self.Roster()
        self.Boards()
        self.IMPTable()
        self.ScoreTable()
        self.PairCard() # pdf only
        self.Journal()  # pdf only
        self.save()
        return

if __name__ == '__main__':
    log = setlog('team', None)
    def even_type(value):
        ivalue = int(value)
        if ivalue % 2 != 0:
            raise argparse.ArgumentTypeError(f"{value} is not an even number")
        return ivalue

    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--debug', type=str, default='INFO', help='Debug level, INFO, DEBUG, ERROR')
    parser.add_argument('-r', '--round', type=even_type, default=2, help='Number of rounds')
    parser.add_argument('-b', '--boards', type=even_type, default=4, help='Number of boards per round')
    parser.add_argument('-m', '--match', type=int, choices=range(1,4), default=1, help='Number of matches')
    parser.add_argument('-f', '--fake', type=bool, default=False, help='Fake scores to test the spreadsheet')
    args = parser.parse_args()
    for l in [['INFO', logging.INFO], ['DEBUG', logging.DEBUG], ['ERROR', logging.ERROR]]:
        if args.debug.upper() == l[0]:
            log.setLevel(l[1])
            break
    team = TeamMatch(log)
    # A match has n rounds, each round has m boards, divided into two halves, each half of the boards
    team.setup(rounds=args.round, boards=args.boards, match=args.match, fake=args.fake)
    team.match()