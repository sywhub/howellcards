#!/usr/bin/env python3
# Mainly spreadsheet class to Howell tournaments into Excel templates
# Also produce PDF file the event
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.formula import ArrayFormula
import pdf
import random
import datetime
import os

class DupBridge:
    def __init__(self, log):
        self.log = log
        self.HeaderFont = Font(bold=True, size=14)
        self.centerAlign = Alignment(horizontal='center')
        self.trumps = ('D/C', 'H/S', 'NT')  
        self.thinLine = Side(style='thin', color="000000")
        self.mediumLine = Side(style='medium',color="000000")
        self.thinTop = Border(top=self.thinLine)
        self.thinLeft = Border(left=self.thinLine)
        self.fake = False

    # IMP conversion table
    def IMPTable(self):
        IMPRanges = [0, 20, 50, 90, 130, 170, 220, 270, 320, 370, 430, 500, 600, \
                    750, 900, 1100, 1300, 1500, 1750, 2000, 2250, 2500, 3000, 3500, 4000, 10010]
        sh = self.wb.create_sheet('IMP Table')
        row = self.headerRow(sh, ['From', 'To', 'IMP'])
        for i in range(0, len(IMPRanges)-1):
            sh.cell(row, 1).value = IMPRanges[i] 
            sh.cell(row, 2).value = IMPRanges[i+1] - 10
            sh.cell(row, 3).value = i
            row += 1
        return row

    # (Default) 1st row as the header
    # return the next row for data
    def headerRow(self, sh, headers, row = 1):
        if not hasattr(self, 'HeaderFont'):
            self.HeaderFont = Font(bold=True, size=14)
        col = 1
        for h in headers:
            sh.cell(row, col).font = self.HeaderFont
            sh.cell(row, col).alignment = self.centerAlign
            if type(h) is str:
                sh.cell(row, col).value = h
            else:
                sh.cell(row, col).value = h[0]
                mRange = f'{self.rc2a1(row, col)}:{self.rc2a1(row, col+h[1]-1)}'
                sh.merge_cells(mRange)
                col += h[1] - 1
            col += 1
        try:
            contractCol = headers.index('Contract')
            sh.column_dimensions[chr(ord('A')+contractCol)].width = 30;
        except ValueError:
            pass

        return row + 1

    # A sheet for easy scoring
    def ScoreTable(self):
        sh = self.wb.create_sheet('Scoring Table')
        for c in [3, 11]:
            v = ['Not Vulnerable', 'Vulnerable']
            for x in range(2):
                sh.cell(1, c+x*3).value = v[x]
                sh.cell(1, c+x*3).font = self.HeaderFont
                sh.merge_cells(f'{sh.cell(1, c+x*3).coordinate}:{sh.cell(1, c+x*3+2).coordinate}')
                sh.cell(1, c+x*3).alignment = self.centerAlign
        headers = ['Contract', 'Made', '', 'X', 'XX', '', 'X', 'XX']
        row = self.headerRow(sh, headers, 2)
        self.scorePenalty(sh, row, len(headers)+2, headers[2:])
        for i in range(1,8):
            for trump in self.trumps:
                sh.cell(row, 1).value = f'{i} {trump}'
                sh.cell(row, 1).font = self.HeaderFont
                sh.cell(row, 1).alignment = self.centerAlign
                for j in range(0, 8-i):
                    sh.cell(row, 2).value = j+1
                    for k in range(3):
                        sh.cell(row, 3+k).value = self.score(i, trump, j, False, k)
                        sh.cell(row, 6+k).value = self.score(i, trump, j, True, k)
                        sh.cell(row, 3+k).number_format = sh.cell(row, 6+k).number_format = "#0"
                    row += 1

    # This is based on the rules for duplicate bridge
    # We use table lookup extensively
    def score(self, level, trumpSuit, res, vul, dbl):
        baseScores = [[20],[30],[40,30]]
        overDblBonus = [100, 200]
        gameBonus = [300, 500]
        slamBonus = [500, 750]
        gSlamBonus = [1000, 1500]
        dblBonus = 50
        gameThreshold = 100
        partialBonus = 50

        dblMul = 2**dbl # 2 to the power of "dbl" which is 0, 1, or 2
        vulIdx = 1 if vul else 0
        score = 0

        # Pick the right table
        tbl = baseScores[self.trumps.index(trumpSuit)]
        lTbl = len(tbl) - 1
        for c in range(level):
            if c < lTbl:
                score += tbl[c]
            else:
                score += tbl[lTbl]
        score *= dblMul
        if score < gameThreshold:
            score += partialBonus
        else:   # made game
            score += gameBonus[vulIdx]
        if level == 6:
            score += slamBonus[vulIdx]
        elif level == 7:
            score += gSlamBonus[vulIdx]
        score += dblBonus * dbl
        overTricks = res * tbl[lTbl] * dblMul
        if overTricks > 0 and dbl > 0:
            overTricks = res * overDblBonus[vulIdx] * dbl
        score += overTricks
        return score

    # The table for failing the contract
    def scorePenalty(self, sh, row, col, headers):
        penaltyTbl = [[50], [100], [100, 200, 200, 300], [200, 300]]         
        headers.insert(0, 'Down by')
        for i in range(len(headers)):
            sh.cell(row-1, col+i).value = headers[i]
            sh.cell(row-1, col+i).font = self.HeaderFont
            sh.cell(row-1, col+i).alignment = self.centerAlign

        sh.cell(row, col).value = -1
        sh.cell(row, col+1).value = -penaltyTbl[0][0]
        sh.cell(row, col+2).value = -penaltyTbl[2][0]
        sh.cell(row, col+3).value = f'={chr(ord('B')+col)}{row}*2'
        sh.cell(row, col+4).value = -penaltyTbl[1][0]
        sh.cell(row, col+5).value = -penaltyTbl[3][0]
        sh.cell(row, col+6).value = f'={chr(ord('B')+col+3)}{row}*2'
        row += 1
        for down in range(2,14):
            sh.cell(row, col).value = -down
            sh.cell(row, col+1).value = f'={sh.cell(row-1,col+1).coordinate}-{penaltyTbl[0][0]}'
            sh.cell(row, col+2).value = f'={sh.cell(row-1,col+2).coordinate}-{penaltyTbl[2][down - 1 if down <= 4 else 3]}'
            sh.cell(row, col+3).value = f'={sh.cell(row,col+2).coordinate}*2'
            sh.cell(row, col+4).value = f'={sh.cell(row-1,col+1).coordinate}-{penaltyTbl[1][0]}'
            sh.cell(row, col+5).value = f'={sh.cell(row-1,col+2).coordinate}-{penaltyTbl[3][1]}'
            sh.cell(row, col+6).value = f'={sh.cell(row,col+2).coordinate}*2'
            row += 1

    def placeHolderName(self):
        return f'Name {random.randint(11,90)}'

    def rc2a1(self, r, c):
        col = ''
        if c > 26:
            col = 'A'
            c -= 26
        col += chr(c-1+ord('A'))
        return f"{col}{r}"

    def vulLookup(self, bidx):
        vulShift = bidx // 4
        return ['None', 'NS', 'EW', 'Both'][(bidx + vulShift) % 4]

class PairGames(DupBridge):
    def __init__(self, log):
        super().__init__(log)
        self.noChangeFont = Font(bold=True, italic=True, color='FF0000')
        self.bottomLine = Border(bottom=Side(style='thin', color='000000'))
        self.SITOUT = "Sit-Out"

    def pairN(self, n):
        return n + 1

    # turn pair number to readable ID
    def pairID(self, n):
        return f"{n+1}"
    

    def fakeScore(self, sh, row, col):
        if random.random() < 0.90:
            pickSide = col if random.random() >= 0.5 else col+1
            score = random.randint(2,80)*10
            sh.cell(row, pickSide).value = score
        else:
            sh.cell(row, col).value = 'Avg'
            sh.cell(row, col+1).value = 'Avg'

    def initRounds(self):
        # A convenient restructure
        self.roundData = {}
        for b,bset in self.boardData.items():
            for s in bset:
                if s[0] not in self.roundData:   # round
                    self.roundData[s[0]] = {}
                if s[1] not in self.roundData[s[0]]: # table
                    self.roundData[s[0]][s[1]] = {'NS': s[2], 'EW': s[3], 'Board': []}
                self.roundData[s[0]][s[1]]['Board'].append(b)

    def roundTab(self):
        self.log.debug('Saving by Round')
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
            if (hasattr(self, 'oddPairs') and self.oddPairs and t == len(tables.keys()) - 1):
                continue
            for r in sorted(tables[t].keys()):
                x = tables[t][r][0]
                if self.pairN(x['NS']) == 0:
                    continue
                if bIdx % 4 == 0:
                    self.pdf.add_page()
                    y = 2 * self.pdf.margin
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

    def Journal(self):
        pairData = {}
        for b,r in self.boardData.items():
            for v in r:
                for p in range(2,4):
                    if v[p] not in pairData:
                        pairData[v[p]] = []
                    pairData[v[p]].append((v[0], b, v[1], v[2], v[3])) # (round, board, table, NS, EW)
        tblCols = []
        nPerPage = 2 if len(pairData[1]) < 18 else 1
        pIdx = 0
        xMargin = self.pdf.margin * 2
        hdrs = ['Round', 'Board', 'Sit-Out', 'EW', 'Contract', 'By', 'Result', 'NS', 'EW']
        self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.headerPt)
        self.pdf.setHeaders(xMargin, hdrs, tblCols)
        hdrs[2] = 'NS'
        for p in sorted(pairData.keys()):
            if self.pairID(p) == self.SITOUT:
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

    # Data {pair #: [(round, table, NS, EW), ...], ...}
    def Travelers(self):
        tblCols = []
        xMargin = 0.5
        hdrs = ['Round', 'NS', 'EW', 'Contract', 'By', 'Result', 'NS', 'EW']
        self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.headerPt)
        self.pdf.setHeaders(xMargin, hdrs, tblCols)
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

    def Tables(self, nsTexts, ewTexts):
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
        tblCols[3] = self.pdf.get_string_width('8'*self.decks*2+','*(self.decks-1)) + 0.25
        w = sum(tblCols)
        xMargin = (self.pdf.w - w) / 2
        tblHeight = (len(tables[0][0])+1) * self.pdf.pt2in(self.pdf.rosterPt)
        top = self.pdf.pt2in(self.pdf.bigPt) * 2.5 + 1
        compassTop = self.pdf.h - (self.pdf.pt2in(self.pdf.bigPt) * 5 + self.pdf.starRadius + 2)

        for t in sorted(tables.keys()):
            if self.ifSitout(t, tables[t][0][0]['NS'], tables[t][0][0]['EW']):
                continue
            self.pdf.add_page()
            self.pdf.movementSheet()
            if compassTop > top + tblHeight:
                self.pdf.compass()
            self.pdf.tableAnchors(f"{t+1}")
            if nsTexts != None and ewTexts != None:
                self.pdf.inkEdgeText(nsTexts[t], ewTexts[t])
            self.pdf.set_font(self.pdf.sansSerifFont, style='B', size=self.pdf.rosterPt)
            self.pdf.headerRow(xMargin, top, tblCols, hdrs, ' ')
            self.pdf.set_font(size=self.pdf.rosterPt)
            y = self.pdf.get_y()
            h = self.pdf.lineHeight(self.pdf.font_size_pt);
            self.pdf.set_xy(xMargin, y + h)
            for r in sorted(tables[t].keys()):
                tRound = tables[t][r]
                self.pdf.cell(tblCols[0], h, text=f'{r+1}', align='C', border=1)
                self.pdf.cell(tblCols[1], h, text=f'{self.pairN(tRound[0]['NS'])}', align='C', border=1)
                self.pdf.cell(tblCols[2], h, text=f'{self.pairN(tRound[0]['EW'])}', align='C', border=1)
                bds = ""
                for b in tRound:
                    bds += f'{b['Board']+1},'
                self.pdf.cell(tblCols[3], h, text=bds[:-1], align='C', border=1)
                y += h
                self.pdf.set_xy(xMargin, y + h)

    def idTags(self):
        idData = {}
        for rd, tbl in self.roundData.items(): # (round, table, NS, EW)
            for t, r in tbl.items():
                if r['NS'] not in idData:
                    idData[r['NS']] = []
                if r['EW'] not in idData:
                    idData[r['EW']] = []
                idData[r['NS']].append((rd, t, r['NS'], r['EW']))
                idData[r['EW']].append((rd, t, r['NS'], r['EW']))
        self.idTagsByData(idData)

    def idTagsByData(self, data):
        tags = 0
        colW = []
        hdrs = ['Round', 'Table', 'Seat', 'vs']
        self.pdf.setHeaders(0, hdrs, colW)
        # page can be portrait or landscape
        w = min(self.pdf.w,self.pdf.h)
        leftMargin = (w - 2 * sum(colW)) / 4
        cWidth = w / 2

        nTagsPage = 1 if len(data[1]) > 15 else (4 if len(data[1]) <= 8 else 2)
        for id in sorted(data.keys()):
            if self.pairID(id) == self.SITOUT:
                continue
            if tags % nTagsPage == 0:
                self.pdf.add_page(orientation='P') # no header/footer
                y = self.pdf.margin * 2
            rData = sorted(data[id], key=lambda x: x[0])    # by round
            for half in range(2):   # two identical tags for each person of the pair
                self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.headerPt)
                self.pdf.set_xy(leftMargin+cWidth*half, y)
                self.pdf.cell(text=f"Pair: {self.pairID(id)}")
                ty = y + self.pdf.lineHeight(self.pdf.font_size_pt)
                self.pdf.set_xy(leftMargin+cWidth*half, ty)
                self.pdf.set_font(self.pdf.sansSerifFont, style='B', size=self.pdf.smallPt)
                h = self.pdf.lineHeight(self.pdf.font_size_pt)
                for i in range(len(hdrs)):
                    self.pdf.cell(colW[i], h, text=hdrs[i], align='C', border=1)
                ty +=  h

                for r in rData:
                    opp,seat = (self.pairN(r[3]),'NS') if id == r[2] else (self.pairN(r[2]),'EW')
                    self.pdf.set_xy(leftMargin+cWidth*half, ty)
                    self.pdf.cell(colW[0], h, text=f"{r[0]+1}", align='C', border=1)
                    self.pdf.cell(colW[1], h, text=f"{r[1]+1}", align='C', border=1)
                    self.pdf.cell(colW[2], h, text=f"{seat}", align='C', border=1)
                    self.pdf.cell(colW[3], h, text=f"{opp}", align='C', border=1)
                    ty += h
            tags += 1
            y = self.pdf.sectionDivider(nTagsPage, tags, self.pdf.margin) + self.pdf.margin * 2
        return

    # Computing MPs
    # cIdx is the column on the left (one less) of the MP% columns
    # We have 4 MP "result" columns: 2 for % and 2 for Pts.
    # Then we have another 2 columns for "net" scores.
    # That's 6 columns, therefore the magic "7" is where the calculation area start
    #
    # "Net" is a simple formula so that no cell is blank, other than Averages
    # The calculation area is pair-wise comparisons to all opponents.  It's slightly esoteric.
    def computeMP(self, sh, cIdx, nPlayed, row, cursorRow, netIdx, calcStart=7):
        Win = 1.0
        Tie = 0.5
        Lost = 0.0
        # do twice, one for NS and another for EW
        for i in range(2):
            cStart = cIdx + calcStart + i*(nPlayed - 1)
            cEnd   = cStart + nPlayed - 2
            spread=f"{self.rc2a1(row, cStart)}:{self.rc2a1(row, cEnd)}"
            # The results are the computations from the calculation areas
            # "Non-comparisions" are blank cells and skipped by "COUNT"
            # Therefore the % is based on the times the board is actually played, not counting Averages
            # The Average cell is simply assigned as 50%
            sh.cell(row, cIdx+1+i).value = f"=IF(COUNT({spread})>0,{self.rc2a1(row, cIdx+3+i)}/COUNT({spread}),{Tie})"
            sh.cell(row, cIdx+1+i).number_format = sh.cell(row, cIdx+2).number_format = "0.00%"
            sh.cell(row, cIdx+3+i).value = f"=SUM({spread})"
            sh.cell(row, cIdx+3+i).number_format = "#0.00"

            # "cursorRow" is the counter within the current "play" group.  It goes from 0 to nPlayed - 1.
            # This comprehensive create the relative rows for *this* player to compare with.
            # By definition, there's only nPlayed - 1 opponents.
            opponents = [x - cursorRow for x in range(nPlayed) if x != cursorRow]
            n = nPlayed - 1
            # each column advanced (targetC) compares this row with one of the opponents (different rows)
            # So this is like 2-dimensional movements
            for rCmp in range(n):
                # if self is not a number, then blank out all comparisions
                # if the opponent is not a number, make that comparision blank
                # Otherwise, a win is 1 pt, tie 0.5, and lost 0.o
                cmpF = f"=IF(ISNUMBER({self.rc2a1(row, netIdx+i)}),IF(ISNUMBER({self.rc2a1(row+opponents[rCmp], netIdx+i)}),"
                cmpF += f"IF({self.rc2a1(row, netIdx+i)}>{self.rc2a1(row+opponents[rCmp], netIdx+i)},{Win},"
                cmpF += f'IF({self.rc2a1(row, netIdx+i)}={self.rc2a1(row+opponents[rCmp], netIdx+i)},{Tie},{Lost})),""),"")'
                targetC = cIdx+calcStart+rCmp+i*n
                sh.cell(row, targetC).value = cmpF

    def computeNet(self, sh, row, raw, target):
        rawNS = self.rc2a1(row, raw)
        rawEW = self.rc2a1(row, raw+1)
        sh.cell(row, target).value = f'=IF(ISNUMBER({rawNS}),{rawNS},IF(ISNUMBER({rawEW}),-{rawEW},""))'
        sh.cell(row, target+1).value = f'=IF(ISNUMBER({rawEW}),{rawEW},IF(ISNUMBER({rawNS}),-{rawNS},""))'