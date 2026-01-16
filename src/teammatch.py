#!/usr/bin/env python3
# Generate a team match setup
#   A PDF with Roster and Score sheets
#   An Excel spreadsheet to enter the results and calculate the scores
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
        self.switches = 0
        self.rounds = 2
        self.vulTbl = ['None', 'NS', 'EW', 'Both']
        self.Fake = True

    # record metadata
    def setup(self, rounds, boards, switch, fake):
        self.boards = boards
        self.switches = switch
        self.rounds = rounds
        self.Fake = fake
        self.TeamPairs = {1: [1, 2, 3, 4], 2: [1, 3, 2, 4], 3: [1, 4, 2, 3]}

    # Roster changes with each "switch"
    def rosterSheet(self):
        ws = self.wb.active
        ws.title = 'Roster'
        row = 1
        ws.cell(row, 1).value = 'Players'
        ws.cell(row, 1).font = self.HeaderFont
        ws.cell(row, 1).alignment = self.centerAlign
        ws.merge_cells(f'{ws.cell(row,1).coordinate}:{ws.cell(row,3).coordinate}')
        row += 1
        for pair in range(4):
            ws.cell(row, 1).font = self.HeaderFont
            ws.cell(row, 1).value = f'Pair {pair+1}'
            ws.cell(row, 2).value = self.placeHolderName()
            ws.cell(row, 3).value = self.placeHolderName()
            row += 1
        row += 1

        switchIdx = 1
        while self.switches >= switchIdx:
            ws.cell(row, 1).font = self.HeaderFont
            ws.cell(row, 1).alignment = self.centerAlign
            ws.cell(row, 1).value = f'Rotation #{switchIdx}'
            ws.merge_cells(f'{ws.cell(row,1).coordinate}:{ws.cell(row,6).coordinate}')
            row += 1

            for t in range(2):
                ws.cell(row, 2*t+1).value = f'Team {t+(switchIdx-1)*2+1}'
                ws.cell(row, 2*t+1).font = self.HeaderFont
                ws.cell(row, 2*t+1).alignment = self.centerAlign
                ws.merge_cells(f'{ws.cell(row,2*t+1).coordinate}:{ws.cell(row,2*t+2).coordinate}')
            ws.cell(row, 5).font = self.HeaderFont
            ws.cell(row, 5).alignment = self.centerAlign
            ws.cell(row, 5).value = 'IMP'
            ws.merge_cells(f'{ws.cell(row,5).coordinate}:{ws.cell(row,6).coordinate}')
            row += 1
            idx = switchIdx % 3
            if idx == 0:
                idx = 3
            for i in range(4):
                ws.cell(row, i+1).value = f'Pair {self.TeamPairs[idx][i]}'
            #ws.cell(row, 5).value = f'=SUM(Boards!I3:I{self.switches*self.switches*self.rounds*2+2})'
            switchIdx += 1
            row += 1

    def Roster(self):
        self.rosterSheet()
        return

        # convert percents into inches
        headers = {'Team': self.pdf.epw * 0.05, 'Pairs': self.pdf.epw * 0.1, 'Name': self.pdf.epw * 0.7}

        self.pdf.add_page()
        self.headerFooter()
        self.pdf.set_font(style='BI', size=self.pdf.rosterPt, family=self.pdf.serifFont)
        self.pdf.set_y(self.pdf.margin + self.pdf.lineHeight(self.pdf.font_size_pt) * 2)
        self.pdf.cell(w=self.pdf.epw, text='Team Match Roster', align='C')
        self.pdf.set_y(self.pdf.get_y() + self.pdf.lineHeight(self.pdf.font_size_pt) * 2)

        for team in range(1,3):
            self.pdf.set_font(style='B', size=self.pdf.bigPt, family=self.pdf.serifFont)
            self.pdf.set_x(self.pdf.margin+headers['Team'])
            self.pdf.cell(text=f'Team {team}')
            self.pdf.ln()
            for p in range(1,3):
                self.pdf.set_font(style='', size=self.pdf.linePt, family=self.pdf.sansSerifFont)
                self.pdf.set_x(self.pdf.margin+headers['Team'])
                self.pdf.cell(w=headers['Pairs'], h=self.pdf.lineHeight(self.pdf.font_size_pt), text=f'Pair {p}', align='C', border=1)
                self.pdf.cell(w=headers['Name']/2, h=self.pdf.lineHeight(self.pdf.font_size_pt), text='', border=1)
                self.pdf.cell(w=headers['Name']/2, h=self.pdf.lineHeight(self.pdf.font_size_pt), text='', border=1)
                self.pdf.ln()
            self.pdf.ln()

    # Table of boards played, no PDF equivalent
    def Boards(self):
        self.wb.create_sheet('Boards')
        ws = self.wb['Boards']
        # merged columns
        headers = ['Board', 'Table', 'NS', 'EW', 'Vul', 'Contract', 'By', 'Result', 'NS', 'EW', 'NS', 'EW']
        NSColScoreCol = headers.index('NS', 4) + 1
        col = NSColScoreCol
        for h in ['Score', 'IMP']:
            ws.cell(1, col).value = h
            ws.cell(1, col).font = self.HeaderFont
            ws.cell(1, col).alignment = self.centerAlign
            ws.merge_cells(f'{ws.cell(1,col).coordinate}:{ws.cell(1,col+1).coordinate}')
            col += 2

        # next row is the headers
        col = 1
        row = self.headerRow(ws, headers, 2)
        for board in range(self.boards * self.rounds * self.switches):
            col = 1
            ws.cell(row, col).value = board+1
            vulIdx = board % 4
            ws.cell(row, col+headers.index('Vul')).value = f'{self.vulTbl[vulIdx]}'
            ws.cell(row, col+headers.index('Vul')).alignment = self.centerAlign
            ws.cell(row+1, col+headers.index('Vul')).value = f'{self.vulTbl[vulIdx]}'
            ws.cell(row+1, col+headers.index('Vul')).alignment = self.centerAlign
            ws.cell(row, col+headers.index('Table')).value = 1
            ws.cell(row+1, col+headers.index('Table')).value = 2
            switchIdx = board // (self.rounds * self.boards) + 1
            swap = (board % (self.rounds * self.boards)) // self.boards
            pairCol = headers.index('NS')+1
            ws.cell(row, pairCol).value = f'Pair {self.TeamPairs[switchIdx][0]}'
            ws.cell(row, pairCol+1).value = f'Pair {self.TeamPairs[switchIdx][2 + swap]}'
            ws.cell(row+1, pairCol).value = f'Pair {self.TeamPairs[switchIdx][3 - swap]}'
            ws.cell(row+1, pairCol+1).value = f'Pair {self.TeamPairs[switchIdx][1]}'
            #self.fakeCOntracts(ws, row)
            NSCol = chr(ord('A')+NSColScoreCol)
            EWCol = chr(ord(NSCol)+1)
            for t in range(2):
                NSscore = f"IF({NSCol}{row+1}>0,{NSCol}{row}-{NSCol}{row+1},{NSCol}{row}+{EWCol}{row+1})"
                EWscore = f"IF({EWCol}{row+1}>0,{EWCol}{row}-{EWCol}{row+1},{EWCol}{row}+{NSCol}{row+1})"
                #ws.cell(f'{NSCol}{row}').value = NSscore;
                #ws.cell(f'{EWCol}{row}').value = EWscore;
                ws.cell(row, col+headers.index('NS', NSColScoreCol+1)).value = f"=IF({NSCol}{row}>0,VLOOKUP(ABS({NSscore}),'IMP Table'!$A$2:$C$26,3)*SIGN({NSscore}),-J{row})"
                ws.cell(row, col+headers.index('EW', NSColScoreCol+1)).value = f"=IF({EWCol}{row}>0,VLOOKUP(ABS({EWscore}),'IMP Table'!$A$2:$C$26,3)*SIGN({EWscore}),-I{row})"
                row += 1
            bd = Side(style='thin', color='000000')
            for i in range(1,len(headers)+1):
                ws.cell(row-1, i).border = Border(bottom=bd)

    # Fake a score to test IMP calculation
    def fakeCOntracts(self, ws, row):
        import random
        if self.Fake:
            col = random.choice([7,8])
            ws.cell(row, col).value = random.randint(5,12)*100
            col = random.choice([7,8])
            ws.cell(row+1, col).value = random.randint(5,12)*100
        else:
            ws.cell(row, 7).value = 10

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
                for board in range(startBoard, self.boardPerSwitch * self.switchPerRound + startBoard):
                    ws.cell(row, 1).value = board
                    row += 1
                startBoard = board + 1
                row += 1

    def Score(self):
        headers = {'Board': .1, 'Contract': .4, 'By': .15, 'Result': .25}
        headers = {k: v * self.pdf.epw for k,v in headers.items()}
        self.scoreSheet(list(headers.keys()))
        
        # How many rounds per page?
        boardsPerRound = self.switches * self.boards
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
                for board in range(startBoard, self.boardPerSwitch * self.switchPerRound + startBoard):
                    self.pdf.cell(w=headers['Board'], h=self.pdf.lineHeight(self.pdf.font_size_pt), text=f'{board}', align='C', border=1)
                    for col in [v for k,v in headers.items() if k != 'Board']:
                        self.pdf.cell(w=col, h=self.pdf.lineHeight(self.pdf.font_size_pt), text='', border=1)
                    self.pdf.ln()

    # Output into filesystem
    def save(self):
        import os
        here = os.path.dirname(os.path.abspath(__file__))
        fn = f'{here}/../teammatch{self.rounds}x{self.boards}x{self.switches}'
        self.wb.save(f'{fn}.xlsx')
        #self.pdf.output(f'{fn}.pdf')

    # Some text for the TD/Organizer
    def Instructions(self):
        txt = '''There is a matching spreadsheet for this PDF.
               Print this PDF before the match.
               Have both teams sign in the roster page.
               Team 1 sits NS of table 1, EW of table 2.  Team 2 sits EW of table 1, NS of table 2.
               Place the scoring sheet on each of the table.
               Put boards 1 to 4 on table 1, 5 to 8 on table 2. Each table shuffle and play the boards.
               When done, swap the boards and continue.
               Team 2 pairs swap seats to their team mates of the other table.
               Put boards 9 to 12 on table 1, 13 to 16 on table 2.
               Shuffle, play, swap boards, finish all boards.
               Collect both scoring sheets, and enter the results in the spreadsheet.'''
        self.headerFooter()
        self.pdf.set_font(style='B', size=self.pdf.headerPt)
        h = self.pdf.lineHeight(self.pdf.font_size_pt)
        line = h * 3
        toWrite = 'Team Match Setup'
        w = self.pdf.get_string_width(toWrite)
        x = self.pdf.setHCenter(w)
        self.pdf.set_xy(x, line)
        self.pdf.cell(text=toWrite)
        y = self.pdf.get_y()+1
        self.pdf.set_font(size=self.pdf.linePt)
        h = self.pdf.lineHeight(self.pdf.font_size_pt)
        nLine = 1
        for t in txt.split('\n'):
            self.pdf.set_xy(1, y)
            self.pdf.cell(h, h, f'{nLine}.', align='R')
            self.pdf.set_xy(1+h, y)
            self.pdf.multi_cell(self.pdf.epw-2, h=h, text=t.strip())
            y = self.pdf.get_y()
            nLine += 1

    def headerFooter(self):
        notice = f'For public domain. No rights reserved. {datetime.date.today().strftime("%Y")}.'
        footer = f'{self.Rounds} Rounds, swap every {self.boards} boards'
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

    # Orchestrator
    def match(self):
        # self.Instructions()
        self.Roster()
        self.Boards()
        #self.Score()
        self.IMPTable()
        self.ScoreTable()
        self.save()
        return

if __name__ == '__main__':
    log = setlog('team', None)
    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--debug', type=str, default='INFO', help='Debug level, INFO, DEBUG, ERROR')
    parser.add_argument('-r', '--round', type=int, default=2, help='Number of rounds')
    parser.add_argument('-b', '--boards', type=int, default=4, help='Number of boards per round')
    parser.add_argument('-s', '--switch', type=int, default=1, help='Number of switches')
    parser.add_argument('-f', '--fake', type=bool, default=False, help='Fake scores to test the spreadsheet')
    args = parser.parse_args()
    for l in [['INFO', logging.INFO], ['DEBUG', logging.DEBUG], ['ERROR', logging.ERROR]]:
        if args.debug.upper() == l[0]:
            log.setLevel(l[1])
            break
    team = TeamMatch(log)
    team.setup(rounds=args.round, boards=args.boards, switch=args.switch, fake=args.fake)
    team.match()