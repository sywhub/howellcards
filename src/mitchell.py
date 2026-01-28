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
        self.bottomLine = Border(bottom=Side(style='thin', color='000000'))
        self.pdf = pdf.PDF()
        self.wb = Workbook()

        notice = f'For public domain. No rights reserved. {datetime.date.today().strftime("%Y")}.'
        footer = f'Mitchell Tournament: {(self.pairs+1)//2} Tables, {self.boards} Boards per round'
        self.pdf.HeaderFooterText(notice, footer)
    
    # turn pair number to readable ID
    def pairSide(self, n):
        return ['EW', 'NS'][n % 2]

    def pairN(self, n):
        if n == self.pairs and self.oddPairs:
            return self.SITOUT
        return n // 2 if n % 2 == 0 else n // 2 + 1

    def pairID(self, n):
        if n == self.pairs and self.oddPairs:
            return self.SITOUT
        return f"{self.pairSide(n)} {self.pairN(n)}"

    def go(self):
        # the sequence of calls is important
        self.initData()
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

	# {Rounds: 4, Tables: 3, BoardMovement: null, Arrangement:
    #      [[{NS: 0, EW: 1, Board: 0}, {NS: 3, EW: 4, Board: 1}, {NS: 5, EW: 2, Board: 3}],
    #       [...], ...]
    def initData(self):
        self.boardData = {}
        for r in range(self.tables): # round
            for t in range(self.tables): # table
                b = self.boardIdx(r, t)
                for bset in range(self.boards):
                    if (b + bset) not in self.boardData:
                        self.boardData[b+bset] = []
                    self.boardData[b+bset].append((r, t, self.NSPair(r, t), self.EWPair(r, t)))

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
        
        for s in range(2):
            ws.cell(row, 1).value =  f'{['NS', 'EW'][s]} Pairs'
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
            toN = self.pairs + (1 if self.oddPairs else 0)
            avgStart = row
            for p in range(s, toN, 2):
                pName = self.pairN(p+1)
                if pName == self.SITOUT:
                    continue
                ws.cell(row, 1).font = self.HeaderFont
                ws.cell(row, 1).alignment = self.centerAlign
                ws.cell(row, 1).value = pName
                ws.cell(row, 2).value = self.placeHolderName()
                ws.cell(row, 3).value = self.placeHolderName()
                row += 1
            for i in range(5):
                ws.cell(row-1, i+1).border = self.bottomLine
            ws.cell(row, 3).value = 'Average'
            ws.cell(row, 4).value = f'=AVERAGE({self.rc2a1(avgStart, 4)}:{self.rc2a1(row-1,4)})'
            ws.cell(row, 5).value = f'=AVERAGE({self.rc2a1(avgStart, 5)}:{self.rc2a1(row-1,5)})'
            ws.cell(row,4).number_format = "0.00%"
            ws.cell(row,5).number_format = "#0.0"
            ws.cell(row,3).font = self.noChangeFont
            ws.cell(row,4).font = self.noChangeFont
            ws.cell(row,5).font = self.noChangeFont
            row += 2
        row += 2
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
        for s in range(2):
            self.pdf.set_font(style='BI')
            self.pdf.cell(5, h, text=f'{['NS', 'EW'][s]} Pairs', align='L')
            self.pdf.set_font(style='')
            y += h
            self.pdf.set_xy(leftM, y)
            toN = self.pairs + (1 if self.oddPairs else 0)
            for p in range(s, toN, 2):
                self.pdf.cell(widths[0], h, text=f'{self.pairN(p+1)}', align='C', border=1)
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
        calcs.append('Calculations Area')
        col = len(headers) + 1
        for h in calcs:
            sh.cell(row-1, col).value = h
            sh.cell(row-1, col).font = self.HeaderFont
            sh.cell(row-1, col).alignment = self.centerAlign
            col += 1
        for i in range(len(headers)+1, col+1):
            sh.cell(row-1,i).font = self.noChangeFont
        sh.merge_cells(f"{self.rc2a1(row-1, col-1)}:{self.rc2a1(row-1,col+self.tables)}")
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
            for c in range(len(headers)+len(calcs)+self.tables+2):
                sh.cell(row-1, c+1).border = self.bottomLine
        return

    def roundTab(self):
        self.log.debug('Saving by Round')
        self.roundData = {}
        for b,bset in self.boardData.items():
            for s in bset:
                if s[0] not in self.roundData:   # round
                    self.roundData[s[0]] = {}
                if s[1] not in self.roundData[s[0]]: # table
                    self.roundData[s[0]][s[1]] = {'NS': s[2], 'EW': s[3], 'Board': []}
                self.roundData[s[0]][s[1]]['Board'].append(b)

        headers = ['Round', 'Table', 'NS', 'EW', 'Board', 'Vul', 'Contract', 'By', 'Result', 'NS', 'EW']
        sh, startRow = self.contractHeaders(headers, 'By Round', ['Scores'])
        row = startRow
        for r in sorted(self.roundData.keys()): # round
            sh.cell(row, 1).value = r+1
            sh.cell(row, 1).alignment = self.centerAlign
            for t in sorted(self.roundData[r]): # table
                sh.cell(row, 2).value = t+1
                sh.cell(row, 3).value = self.pairN(self.roundData[r][t]['NS'])
                sh.cell(row, 4).value = self.pairN(self.roundData[r][t]['EW'])
                for b in self.roundData[r][t]['Board']:
                    sh.cell(row, 5).value = b+1
                    sh.cell(row, 6).value = f"{self.vulLookup(b)}"
                    for i in range(2,7):
                        sh.cell(row, i).alignment = self.centerAlign
                    row += 1
                for i in range(2,len(headers)+1):
                    sh.cell(row-1, i).border = self.bottomLine
            sh.cell(row-1,1).border = self.bottomLine
        if self.fake:
            for i in range(startRow, row):
                self.fakeScore(sh, i, headers.index('Result')+2)
                
        return

    def NSPair(self, r, t):
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
        xMargin = self.pdf.margin
        hdrs = ['NS Score', 'Result', 'NS Contract', 'By', 'Board', 'EW Contract', 'By', 'Result', 'EW Socre']
        self.pdf.set_font(self.pdf.sansSerifFont, style='B', size=self.pdf.linePt)
        self.pdf.setHeaders(xMargin, hdrs, tblCols)
        allW = sum(tblCols)
        extraW = (self.pdf.epw - allW) / 2
        tblCols[2] += extraW
        tblCols[5] += extraW
        bIdx = 0
        for t in sorted(tables.keys()):
            if self.oddPairs and t == len(tables.keys()) - 1:
                continue
            for r in sorted(tables[t].keys()):
                if bIdx % 4 == 0:
                    self.pdf.add_page()
                    y = 2 * self.pdf.margin
                x = tables[t][r][0]
                self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.headerPt)
                title = f"Table {t+1}, Round {r+1}, NS: {self.pairN(x["NS"])}, EW: {self.pairN(x["EW"])}"
                y += self.pdf.lineHeight(self.pdf.font_size_pt)
                self.pdf.set_font(self.pdf.sansSerifFont, style='B', size=self.pdf.linePt)
                y = self.pdf.headerRow(xMargin, y, tblCols, hdrs, title)
                self.pdf.set_font(size=self.pdf.linePt)
                h = self.pdf.lineHeight(self.pdf.font_size_pt)
                y += h
                self.pdf.set_xy(xMargin, y)
                for b in tables[t][r]:
                    for i in range(4):
                        self.pdf.cell(tblCols[i], h, text=f'', align='C', border=1)
                    self.pdf.cell(tblCols[4], h, text=f'{b["Board"]+1}', align='C', border=1)
                    for i in range(5,len(hdrs)):
                        self.pdf.cell(tblCols[i], h, text=f'', align='C', border=1)
                    y += h
                    self.pdf.set_xy(xMargin, y)
                bIdx += 1
                y = self.pdf.sectionDivider(4, bIdx, xMargin)
        return

    def Tables(self, sqMove=False):
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
        self.pdf.set_font(self.pdf.sansSerifFont, style='B', size=self.pdf.rosterPt)
        self.pdf.setHeaders(xMargin, hdrs, tblCols)
        tblCols[3] = self.pdf.get_string_width('8'*self.boards*2+','*(self.boards-1)) + 0.25
        w = sum(tblCols)
        xMargin = (self.pdf.w - w) / 2

        for t in sorted(tables.keys()):
            if self.oddPairs and t == len(tables.keys()) - 1:
                continue
            self.pdf.add_page()
            self.pdf.movementSheet()
            self.pdf.compass()
            self.pdf.tableAnchors(f"{t+1}")
            if sqMove:
                sqTxt = ['R2 to T2/EW, R3 to T3/EW, R4 to T2/EW',
                         'R2 to T1/EW, R3 to T4/EW, R4 to T1/EW',
                         'R2 to T4/EW, R3 to T1/EW, R4 to T4/EW',
                         'R2 to T3/EW, R3 to T2/EW, R4 to T3/EW']
                bdTxt = ['Stay here. Boards: R2 to T4, R3 to T2, R4 to T4',
                         'Stay here. Boards: R2 to T3, R3 to T1, R4 to T3', 
                         'Stay here. Boards: R2 to T2, R3 to T4, R4 to T2', 
                         'Stay here. Boards: R2 to T1, R3 to T3, R4 to T1'] 
                ewNext = sqTxt[t]
                nsNext = bdTxt[t]
            else:
                ewNext = f'Move to Table {t+2 if t < 3 else 1} EW'
                nsNext = f'Stay Here, Boards to T{t if t > 0 else 4}'
            self.pdf.inkEdgeText(nsNext, ewNext)
            self.pdf.set_font(self.pdf.sansSerifFont, style='B', size=self.pdf.rosterPt)
            self.pdf.headerRow(xMargin, 2, tblCols, hdrs, ' ')
            self.pdf.set_font(size=self.pdf.rosterPt)
            y = self.pdf.get_y()
            h = self.pdf.lineHeight(self.pdf.font_size_pt);
            self.pdf.set_xy(xMargin, y + h)
            for r in sorted(tables[t].keys()):
                tRound = tables[t][r]
                self.pdf.cell(tblCols[0], h, text=f'{r+1}', align='C', border=1)
                self.pdf.cell(tblCols[1], h, text=f'{self.pairN(tables[t][r][0]['NS'])}', align='C', border=1)
                self.pdf.cell(tblCols[2], h, text=f'{self.pairN(tables[t][r][0]['EW'])}', align='C', border=1)
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
        self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.headerPt)
        self.pdf.setHeaders(xMargin, hdrs, tblCols)
        if self.oddPairs:
            tblCols[1] = self.pdf.get_string_width(self.SITOUT)+0.25
        nPerPage = 4 if len(self.boardData[0]) <= 5 else 2 if len(self.boardData[0]) <= 12 else 1
        bIdx = 0
        for b,r in self.boardData.items():
            if bIdx % nPerPage == 0:
                self.pdf.add_page()
                y = 0.5
            self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.headerPt)
            y = self.pdf.headerRow(xMargin, y, tblCols, hdrs, f'Board {b+1} Traveler')
            y += self.pdf.lineHeight(self.pdf.font_size_pt)
            self.pdf.set_font(self.pdf.sansSerifFont, size=self.pdf.linePt)
            h = self.pdf.lineHeight(self.pdf.font_size_pt)
            for v in r:
                self.pdf.set_xy(xMargin, y)
                self.pdf.cell(tblCols[0], h, text=f'{v[0]+1}', align='C', border=1)
                if type(v[2]) == str:
                    self.pdf.cell(tblCols[1], h, text=v[2], align='C', border=1)
                else:
                    self.pdf.cell(tblCols[1], h, text=f'{self.pairN(v[2])}', align='C', border=1)
                self.pdf.cell(tblCols[2], h, text=f'{self.pairN(v[3])}', align='C', border=1)
                for c in range(3,len(hdrs)):
                    self.pdf.cell(tblCols[c], h, text='', align='C', border=1)
                y += h
            bIdx += 1
            if nPerPage > 1:
                y = self.pdf.sectionDivider(nPerPage, bIdx, xMargin)
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
        nPerPage = 2 if len(pairData[1]) < 24 else 1
        pIdx = 0
        xMargin = self.pdf.margin * 2
        hdrs = ['Round', 'Board', 'NS', 'EW', 'Contract', 'By', 'Result', 'NS', 'EW']
        self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.headerPt)
        self.pdf.setHeaders(xMargin, hdrs, tblCols)
        for p in sorted(pairData.keys()):
            if self.pairN(p) == self.SITOUT:
                continue
            if pIdx % nPerPage == 0:
                self.pdf.add_page()
                self.pdf.headerFooter()
                y = self.pdf.margin*2
            self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.headerPt)
            y = self.pdf.headerRow(xMargin, y, tblCols, hdrs ,f"Pair {self.pairID(p)} Play Journal")
            y += self.pdf.lineHeight(self.pdf.font_size_pt)
            self.pdf.set_font(size=self.pdf.linePt)
            h = self.pdf.lineHeight(self.pdf.font_size_pt)
            self.pdf.set_xy(xMargin, y)
            for v in sorted(pairData[p], key=lambda x: x[0]):
                self.pdf.cell(tblCols[0], h, text=f'{v[0]+1}', align='C', border=1)
                self.pdf.cell(tblCols[1], h, text=f'{v[1]+1}', align='C', border=1)
                self.pdf.cell(tblCols[2], h, text=f'{self.pairN(v[3])}', align='C', border=1)
                self.pdf.cell(tblCols[3], h, text=f'{self.pairN(v[4])}', align='C', border=1)
                for c in range(4,len(hdrs)):
                    self.pdf.cell(tblCols[c], h, text='', align='C', border=1)
                y += h
                self.pdf.set_xy(xMargin, y)
            pIdx += 1
            y = self.pdf.sectionDivider(nPerPage, pIdx, self.pdf.margin)
        return

    def results(self):
        sh = self.wb['Roster']
        nRows = 0
        row = 4
        divident = len(self.roundData) * len(self.roundData[0][0]['Board'])

        for b in self.boardData.values():
            nRows += len(b)
        for s in range(2):
            toN = self.pairs + (1 if self.oddPairs else 0)
            for p in range(s, toN, 2):
                pName = self.pairN(p+1)
                if pName == self.SITOUT:
                    continue
                ifRange = f"'By Board'!{self.rc2a1(3, 4+s)}:{self.rc2a1(3+nRows,4+s)}"
                sumRange = f"'By Board'!{self.rc2a1(3, 12+s)}:{self.rc2a1(3+nRows,12+s)}"
                ptsRange = f"'By Board'!{self.rc2a1(3, 14+s)}:{self.rc2a1(3+nRows,14+s)}"
                sh.cell(row,4).value=f"=SUMIF({ifRange},\"=\"&{self.rc2a1(row, 1)},{sumRange})/{divident}"
                sh.cell(row,5).value=f"=SUMIF({ifRange},\"=\"&{self.rc2a1(row, 1)},{ptsRange})"
                sh.cell(row,4).number_format = "0.00%"
                sh.cell(row,5).number_format = "#0.0"
                row += 1
            row += 3

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
    def mitchell_check(value):
        ivalue = int(value)
        if ivalue == 16:
            raise argparse.ArgumentTypeError(f"Cannot have even number of tables")
        return ivalue

    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--debug', type=str, default='INFO', help='Debug level, INFO, DEBUG, ERROR')
    parser.add_argument('-b', '--boards', type=int, choices=range(1,7), default=4, help='Boards per round')
    parser.add_argument('-p', '--pair', type=mitchell_check, choices=range(8,24), default=8, help='Number of pairs')
    parser.add_argument('-f', '--fake', type=bool, default=False, help='Fake scores to test the spreadsheet')
    args = parser.parse_args()
    for l in [['INFO', logging.INFO], ['DEBUG', logging.DEBUG], ['ERROR', logging.ERROR]]:
        if args.debug.upper() == l[0]:
            log.setLevel(l[1])
            break
    mitchell = Mitchell(log, args.pair, args.boards, args.fake)
    # A match has n rounds, each round has m boards, divided into two halves, each half of the boards
    mitchell.go()