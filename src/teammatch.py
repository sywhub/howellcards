#!/usr/bin/env python3
import argparse
import logging
from maininit import setlog
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
import pdf
from docset import DupBridge

class TeamMatch(DupBridge):
    def __init__(self, leg):
        super().__init__(log)
        self.pdf = pdf.PDF()
        self.wb = Workbook()
        self.boardPerSwitch = 4
        self.switchPerRound = 2
        self.Rounds = 2
        self.vulTbl = ['None', 'NS', 'EW', 'Both']
        self.Fake = True

    def setup(self, rounds, boards, switch, fake):
        self.boardPerSwitch = boards
        self.switchPerRound = switch
        self.Rounds = rounds
        self.Fake = fake

    def Roster(self):
        ws = self.wb.active
        headers = ['Team', 'Pairs']
        pdfWidths = [x * self.pdf.epw / 100 for x in [10, 80]]

        self.pdf.set_font(style='BI', size=self.pdf.rosterPt, family=self.pdf.serifFont)
        self.pdf.cell(w=self.pdf.epw, text='Team Match Roster', align='C')
        self.pdf.ln()
        self.pdf.ln()
        for t in range(2):
            ws.cell(1, 2*t+2).value = f'Team {t+1}'
            ws.cell(1, 2*t+2).font = self.HeaderFont
            ws.cell(1, 2*t+2).alignment = self.centerAlign
            ws.merge_cells(f'{ws.cell(1,2*t+2).coordinate}:{ws.cell(1,2*t+3).coordinate}')

        row = 2
        for pair in range(2):
            ws.cell(row, 1).value = f'Pair {pair+1}'
            ws.cell(row, 1).font = self.HeaderFont
            i = 0
            self.pdf.set_font(style='B', size=self.pdf.headerPt, family=self.pdf.serifFont)
            for h in headers:
                self.pdf.cell(w=pdfWidths[i], h=self.pdf.lineHeight(self.pdf.font_size_pt), text=f'{h}', align='C', border=1)
                i += 1
            self.pdf.ln()
            self.pdf.set_font(style='', size=self.pdf.linePt, family=self.pdf.sansSerifFont)
            self.pdf.cell(w=pdfWidths[0], h=self.pdf.lineHeight(self.pdf.font_size_pt), text=f'Pair 1', align='C', border=1)
            self.pdf.cell(w=pdfWidths[1]/2, h=self.pdf.lineHeight(self.pdf.font_size_pt), text='', border=1)
            self.pdf.cell(w=pdfWidths[1]/2, h=self.pdf.lineHeight(self.pdf.font_size_pt), text='', border=1)
            self.pdf.ln()
            self.pdf.cell(w=pdfWidths[0], h=self.pdf.lineHeight(self.pdf.font_size_pt), text=f'Pair 2', align='C', border=1)
            self.pdf.cell(w=pdfWidths[1]/2, h=self.pdf.lineHeight(self.pdf.font_size_pt), text='', border=1)
            self.pdf.cell(w=pdfWidths[1]/2, h=self.pdf.lineHeight(self.pdf.font_size_pt), text='', border=1)
            self.pdf.ln()
            for team in range(2):
                ws.cell(row,2*team+2).value = self.placeHolderName()
                ws.cell(row,2*team+3).value = self.placeHolderName()
            row += 1
        ws.cell(row, 1).value = 'Total IMP'
        ws.cell(row, 3).value = f'=SUM(Boards!I3:I{self.boardPerSwitch*self.switchPerRound*self.Rounds*2+2})'
        ws.cell(row, 5).value = f'=SUM(Boards!J3:J{self.boardPerSwitch*self.switchPerRound*self.Rounds*2+2})'
        bd = Side(style='double', color='000000')
        for i in [1,3,5]:
            if i != 1:
                ws.cell(row, i).border = Border(top=bd)
            ws.cell(row, i).font = self.HeaderFont
            ws.cell(row, i).alignment = self.centerAlign

    def Boards(self):
        self.wb.create_sheet('Boards')
        ws = self.wb['Boards']
        col = 7
        for h in ['Score', 'IMP']:
            ws.cell(1, col).value = h
            ws.cell(1, col).font = self.HeaderFont
            ws.cell(1, col).alignment = self.centerAlign
            ws.merge_cells(f'{ws.cell(1,col).coordinate}:{ws.cell(1,col+1).coordinate}')
            col += 2
        col = 1
        headers = ['Board', 'Vul', 'Table', 'Contract', 'By', 'Result', 'NS', 'EW', 'NS', 'EW']
        row = self.headerRow(ws, headers, 2)
        for board in range(1, self.boardPerSwitch * self.Rounds * self.switchPerRound + 1):
            col = 1
            ws.cell(row, col).value = board
            vulIdx = (board - 1) % 4 + (board - 1) // 4
            vulIdx %= 4
            ws.cell(row, col+1).value = f'{self.vulTbl[vulIdx]}'
            ws.cell(row, col+2).value = 1
            ws.cell(row+1, col+2).value = 2
            self.fakeCOntracts(ws, row)
            NSscore = f"IF(G{row+1}>0,G{row}-G{row+1},G{row}+H{row+1})"
            EWscore = f"IF(H{row+1}>0,H{row}-H{row+1},H{row}+G{row+1})"
            ws.cell(row, col+8).value = f"=IF(G{row}>0,VLOOKUP(ABS({NSscore}),'IMP Table'!$A$2:$C$26,3)*SIGN({NSscore}),-J{row})"
            ws.cell(row, col+9).value = f"=IF(H{row}>0,VLOOKUP(ABS({EWscore}),'IMP Table'!$A$2:$C$26,3)*SIGN({EWscore}),-I{row})"
            bd = Side(style='thin', color='dd0000')
            row += 1
            for i in range(1,len(headers)+1):
                ws.cell(row, i).border = Border(bottom=bd)
            row += 1

    def boardHeaders(self, ws):
        for h in headers:
            ws.cell(2, col).value = h
            ws.cell(2, col).font = Font(bold=True)
            ws.cell(2, col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            bd = Side(style='thin', color='000000')
            ws.cell(2, col).border = Border(bottom=bd)
            col += 1
        return 3;

    def fakeCOntracts(self, ws, row):
        import random
        if self.Fake:
            col = random.choice([7,8])
            ws.cell(row, col).value = random.randint(5,12)*100
            col = random.choice([7,8])
            ws.cell(row+1, col).value = random.randint(5,12)*100
        else:
            ws.cell(row, 7).value = 10

    def ScoreSheet(self):
        headers = ['Board', 'Contract', 'By', 'Result']
        pdfWidths = [x * self.pdf.epw / 100 for x in [10, 40, 15, 25]]
        roundHeight = self.pdf.eph / self.Rounds
        for t in range(2):
            self.pdf.add_page()
            self.wb.create_sheet(f'Table{t+1} Scores')
            row = 1
            startBoard = 1
            for r in range(1, self.Rounds + 1):
                ws = self.wb[f'Table{t+1} Scores']
                self.pdf.set_y(self.pdf.margin + (r - 1) * roundHeight)
                self.pdf.set_font(style='B', size=self.pdf.rosterPt, family=self.pdf.serifFont)
                self.pdf.cell(text=f'Table {t+1}: Round {r}', align='C')
                self.pdf.set_y(self.pdf.get_y() + self.pdf.lineHeight(self.pdf.font_size_pt))
                self.pdf.set_font(style='B', size=self.pdf.titlePt)
                row = self.headerRow(ws, [f'Table {t+1}: Round {r}'], row)
                ws.merge_cells(f'{ws.cell(row-1,1).coordinate}:{ws.cell(row-1,len(headers)+1).coordinate}')
                row = self.headerRow(ws, headers, row)
                self.pdf.set_font(style='B', size=self.pdf.headerPt, family=self.pdf.sansSerifFont)
                self.pdf.set_y(self.pdf.get_y() + self.pdf.lineHeight(self.pdf.font_size_pt))
                i = 0
                for h in headers:
                    self.pdf.cell(w=pdfWidths[i], h=self.pdf.lineHeight(self.pdf.font_size_pt), text=h, align='C', border=1)
                    i += 1
                self.pdf.ln()
                self.pdf.set_font(style='', size=self.pdf.linePt)
                for board in range(startBoard, self.boardPerSwitch * self.switchPerRound + startBoard):
                    ws.cell(row, 1).value = board
                    self.pdf.cell(w=pdfWidths[0], h=self.pdf.lineHeight(self.pdf.font_size_pt), text=f'{board}', align='C', border=1)
                    for col in range(1, len(headers)):
                        self.pdf.cell(w=pdfWidths[col], h=self.pdf.lineHeight(self.pdf.font_size_pt), text='', border=1)
                    self.pdf.ln()
                    row += 1
                startBoard = board + 1
                row += 1

    def save(self):
        import os
        here = os.path.dirname(os.path.abspath(__file__))
        fn = f'{here}/../teammatch{self.Rounds}x{self.boardPerSwitch*self.switchPerRound}'
        self.wb.save(f'{fn}.xlsx')
        self.pdf.output(f'{fn}.pdf')

    def match(self):
        self.Roster()
        self.Boards()
        self.ScoreSheet()
        self.IMPTable()
        self.ScoreTable()
        self.save()
        return

if __name__ == '__main__':
    log = setlog('team', None)
    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--debug', type=str, default='INFO')
    parser.add_argument('-r', '--round', type=int, default=2)
    parser.add_argument('-b', '--boards', type=int, default=4)
    parser.add_argument('-s', '--switch', type=int, default=2)
    parser.add_argument('-f', '--fake', type=bool, default=False)
    args = parser.parse_args()
    for l in [['INFO', logging.INFO], ['DEBUG', logging.DEBUG], ['ERROR', logging.ERROR]]:
        if args.debug.upper() == l[0]:
            log.setLevel(l[1])
            break
    team = TeamMatch(log)
    team.setup(rounds=args.round, boards=args.boards, switch=args.switch, fake=args.fake)
    team.match()