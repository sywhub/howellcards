#!/usr/bin/env python3
# Generate Howell movement placards and scoring spreadsheet based on the number of pairs
# The tournament setups were pre-generated and stored as JSON.
# The algorithm of generating those data is a separate program, as those generations take time.

# --pair #: generate sheets for the specific pair #, if absent do all of them
# --fake: fake results to test the scoring mechanism in the spreadsheet
# --debug <DEBUG LEVEL>: used only by the developer

# to do: do Google sheet instead of Microsoft Excel
#        Smooother board transitions

import argparse
import pdf
import os
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
import logging
import jsonIO
from maininit import setlog
from docset import PairGames

class Howell(PairGames):
    def __init__(self, log, toFake=False):
        super().__init__(log)
        self.notice = 'For public domain. No rights reserved. Generated on'
        self.fakeResult = toFake
        self.pdf = pdf.PDF()
        self.wb = Workbook()
        self.here = os.path.dirname(os.path.abspath(__file__))
        return

    def save(self):
        self.Traveler()
        self.IMPTable()
        self.ScoreTable()
        self.wb.save(f'{self.here}/../howell{self.pairs}.xlsx')
        self.pdf.output(f'{self.here}/../howell{self.pairs}.pdf')

    def boardList(self, bIdx):
        return [self.decks*bIdx+x+1 for x in range(self.decks)]

    # string to enumerate a "board set" into individual decks
    def boardSet(self, bIdx):
        bds = self.boardList(bIdx)
        str = ''
        for i in bds:
            str += f'{i}'
            if i != bds[-1]:
                str += ' & '
        return str

    # probably should be part of the constructor
    # Initialize some state.
    # Create meta and roster sheets
    def init(self, pairs, nRound):
        self.pairs = pairs
        if pairs <= 6:
            self.decks = 3
        else:
            self.decks = 2

        # meta data
        tourneyMeta = [['Howell Arrangement (IMP & MP)'],
            ['Pairs',pairs], ['Tables',int((pairs + (pairs % 2))/ 2)],
            ['Rounds',nRound], ['Boards per round',self.decks], ['Total Boards to play', self.decks*nRound]]

        ws = self.wb.active
        ws.title = 'Tournament'
        ws.cell(1, 1).value = f'{self.notice} {datetime.date.today().strftime("%b %d, %Y")}.'
        # a less noticable color
        ws.cell(1, 1).font = Font(size=10, italic=True, color="5DADE2")

        for row in range(len(tourneyMeta)):
            ws.cell(row+2, 1).value = tourneyMeta[row][0]
            ws.cell(row+2, 1).font = self.HeaderFont
            if len(tourneyMeta[row]) > 1:
                ws.cell(row+2, 2).value = tourneyMeta[row][1]
                ws.cell(row+2, 2).font = self.HeaderFont
        ws.column_dimensions['A'].width = 30

        self.pdf.HeaderFooterText(f'{self.notice} {datetime.date.today().strftime("%b %d, %Y")}.',
            f'Howell Movement for {pairs} Pairs')
        self.pdf.meta(self.log, ws.title, tourneyMeta)
        self.pdf.instructions(self.log, "instructions.txt")
        self.rosterSheet()


    # A sheet to present the round-oriented data
    # This is the "native" view from data structure's point of view
    # A "round" is keyed by its number (zero based)
    # The value part is a "Table"
    # A table is also (zero) keyed as the table number
    # Its value is another dictionary of "ns", "ew", and "board"
    # which are the pair IDs and the board "set" to be play for that table at that round
    def saveByRound(self, rounds):
        self.log.debug('Saving by Round')
        headers = ['Round', 'Table', 'NS', 'EW', 'Board', 'Vul', 'Contract', 'By', 'Result', 'NS', 'EW']
        self.log.debug('Saving by Round')
        sh = self.wb.create_sheet('By Round')
        sCol = headers.index('Result')+2
        sh.cell(1, sCol).value = 'Scores'
        sh.merge_cells(f'{self.rc2a1(1, sCol)}:{self.rc2a1(1, sCol+1)}')
        sh.cell(1, sCol).font = self.HeaderFont
        sh.cell(1, sCol).alignment = self.centerAlign
        row = self.headerRow(sh, headers, 2)
        sh.column_dimensions[chr(headers.index('Contract')+ord('A'))].width = 30
        sh.column_dimensions['H'].width = 15
        sh.column_dimensions['I'].width = 15
        for rIdx, r in enumerate(rounds):
            sh.cell(row, 1).value = rIdx+1
            for tIdx, tbl in enumerate(r):
                sh.cell(row, 2).value = tIdx+1
                if tbl['NS'] == 0:
                    sh.cell(row, 3).value = "Sit-Out"
                else:
                    sh.cell(row, 3).value = tbl['NS']
                sh.cell(row, 4).value = tbl['EW']
                for b in self.boardList(tbl['Board']):
                    sh.cell(row, 5).value = b
                    sh.cell(row, 6).value = self.vulLookup(b-1)
                    sh.cell(row, 6).alignment = self.centerAlign
                    if self.fakeResult:
                        self.fakeScore(sh, row, 10)
                    row += 1
                for c in range(2,12):
                    sh.cell(row, c).border = self.thinLine
            for c in range(1,12):
                sh.cell(row, c).border = self.thinLine

    # Present the same data table-oriented
    def saveByTable(self, rounds):
        self.log.debug('Saving by Table')
        nTbl = len(rounds[0])
        nRounds = len(rounds)
        # tbl#: {'nRound': # of rounds, r: ({'NS': ns, 'EW': ew, 'Board': board set #}, "boards set")}
        pdfData = {}
        # iterate by table then by round
        for tbl in range(nTbl):
            pdfData[tbl] = {'nRound': nRounds}
            for r in range(nRounds):
                # Simply referene the "By Round" sheet
                pdfData[tbl][r] = (rounds[r][tbl], self.boardSet(rounds[r][tbl]['Board']))
                # The movement, which table/seat for the next round
                if r != nRounds - 1:
                    for side in ['NS', 'EW']:
                        # build a reverse lookup of "side: table" of next round
                        next = {v[side]: k for k,v in enumerate(rounds[r+1])}
                        # look up the pair im that side's lookup
                        if rounds[r][tbl]['NS'] in next.keys():
                            pdfData[tbl]['nsNext'] = (next[rounds[r][tbl]['NS']], side)
                        if rounds[r][tbl]['EW'] in next.keys():
                            pdfData[tbl]['ewNext'] = (next[rounds[r][tbl]['EW']], side)
        self.pdf.overview(pdfData)
        self.pdf.tableOut(pdfData)
        self.pdf.idTags(pdfData)
        self.pdf.pickupSlips(pdfData, self.decks)
        self.pdf.pairRecords(pdfData, self.decks)

    # Board-oriented view
    # A "board" is really a set of decks in the code.  The number of decks is in
    # "self.decks".  We make it 3 for 6-pair tournaments and 2 otherwise.
    # In this "by board" sheet, we also do the scoring.
    def saveByBoard(self, rounds):
        self.log.debug('Saving by Board')
        sh = self.wb.create_sheet('By Board', 2)    # insert it as the 2nd sheet
        nTbl = len(rounds[0])
        # first row setup some spanning column headers
        mergeHdrs = [['Score', 2], ['IMP', 2], ['MP %', 2], ['MP Pts', 2], ['Net', 2],
               ['IMP Calculation', nTbl*2 - 2],['MP Calculation', nTbl*2 - 2]]

        headers = ['Board', 'Round', 'Table', 'NS', 'EW', 'Vul', 'Contract', 'By', 'Result'] + ['NS', 'EW'] * 5
        pairWiseCol = len(headers) + 1
        cStart = headers.index('Result') + 2
        for h in mergeHdrs:
            sh.cell(1, cStart).value = h[0]
            sh.cell(1, cStart).font = self.noChangeFont
            sh.cell(1, cStart).alignment = self.centerAlign
            sh.merge_cells(f'{sh.cell(1,cStart).coordinate}:{sh.cell(1,cStart+h[1]-1).coordinate}')
            cStart += h[1]
        headers += [['NS Pair-Wise', nTbl - 1], ['EW Pair-Wise', nTbl - 1], ['NS MP Score', nTbl - 1], ['EW MP Score', nTbl - 1]]
        # The contract column should be wider for data entry
        row = self.headerRow(sh, headers, 2)

        # build a datastructure for ease of navigation
        # just pivotig the source data
        # board keyed by board set #, value = [(round, table, NS, EW), ...]
        boards = {}
        for r,t in enumerate(rounds):
            for tbl, p in enumerate(t):
                for b in [p['Board']*self.decks + x for x in range(self.decks)]:
                    if b not in boards:
                        boards[b] = []
                    boards[b].append((r, tbl, p['NS'], p['EW']))

        self.pdf.travelers(self.log, self.decks, boards)

        # each iteration advanceds by a set of boards, governed by self.decks
        for b in sorted(boards.keys()): # b is a set of self.decks
            sh.cell(row, 1).value = b + 1
            cursorRow = 0
            # loop through the "rounds" this board were played
            for r in sorted(boards[b], key=lambda x: x[0]):
                nPlayed = len(boards[b])    # # of times this board was played
                # always reference the "By Round" sheet for ease of editing by hand
                roundRow = r[0]*nTbl*self.decks+3
                sh.cell(row, 2).value = f"='By Round'!{self.rc2a1(roundRow, 1)}"
                for c in range(2, 11):
                    boardRow = roundRow + r[1]*self.decks
                    a1 = self.rc2a1(boardRow, c if c < 5 else c + 1)
                    cVal = f"'By Round'!{a1}"
                    if c >= 6:
                        bcheck = f'=IF(ISBLANK({cVal}),"",{cVal})'
                    else:
                        bcheck= f'={cVal}'
                    sh.cell(row, 1+c).value = bcheck 
                sh.cell(row, 6).value = self.vulLookup(b)
                sh.cell(row, 6).alignment = self.centerAlign

                # There are steps to calculate IMP for each board
                # Here are two columns for the end result
                nCmps = len(r) - 1
                avgRange = f'{sh.cell(row, pairWiseCol).coordinate}:{sh.cell(row, pairWiseCol+nCmps-1).coordinate}'
                sh.cell(row, 12).value = f'=IFERROR(AVERAGE({avgRange}),"")'
                sh.cell(row, 12).number_format = '#0.00'
                avgRange = f'{sh.cell(row, pairWiseCol+nCmps).coordinate}:{sh.cell(row, pairWiseCol+2*nCmps-1).coordinate}'
                sh.cell(row, 13).value = f'=IFERROR(AVERAGE({avgRange}),"")'
                sh.cell(row, 13).number_format = '#0.00'
                sumRange = f'{sh.cell(row, pairWiseCol+2*nCmps).coordinate}:{sh.cell(row, pairWiseCol+3*nCmps-1).coordinate}'
                sh.cell(row, 14).value = f'={self.rc2a1(row, 16)}/{nPlayed-1}'
                sh.cell(row, 14).number_format = '0.0%'
                sh.cell(row, 15).value = f'={self.rc2a1(row, 17)}/{nPlayed-1}'
                sh.cell(row, 15).number_format = '0.0%'
                sh.cell(row, 16).value = f'=IFERROR(SUM({sumRange}),"")'
                sh.cell(row, 16).number_format = '#0.00'
                sumRange = f'{sh.cell(row, pairWiseCol+3*nCmps).coordinate}:{sh.cell(row, pairWiseCol+4*nCmps-1).coordinate}'
                sh.cell(row, 17).value = f'=IFERROR(SUM({sumRange}),"")'
                sh.cell(row, 17).number_format = '#0.00'

                # IMP Computation sequence
                # 1. For each side, record their "net" raw scores.  Negative if the other side scored
                sh.cell(row, 18).value = f'=IF(ISNUMBER(J{row}),J{row},IF(ISNUMBER(K{row}),-K{row},""))'
                sh.cell(row, 19).value = f'=IF(ISNUMBER(K{row}),K{row},IF(ISNUMBER(J{row}),-J{row},""))'
                # 2. Compare to all other pairs, on the same side, and use the difference to lookup IMPs
                # competitors are all other pairs
                opponents = [x - cursorRow for x in range(nPlayed) if x != cursorRow]
                startCol = [20, 20+nPlayed-1]
                lookupCol = ['R', 'S']
                for i in range(2):
                    colInc = 0  # distance to the previous section
                    n = nPlayed - 1
                    for rCmp in range(n):
                        cond=f'AND(ISNUMBER({lookupCol[i]}{row}),ISNUMBER({lookupCol[i]}{row+opponents[rCmp]}))'
                        lookup=f"VLOOKUP(ABS({lookupCol[i]}{row}-{lookupCol[i]}{row+opponents[rCmp]}),'IMP Table'!$A$2:$C$26,3)*SIGN({lookupCol[i]}{row}-{lookupCol[i]}{row+opponents[rCmp]})"
                        formula=f'=IF({cond},{lookup},"")'
                        sh.cell(row, startCol[i]+colInc).value = formula
                        cmpF = f'=IF({cond},'
                        cmpF += f"IF({lookupCol[i]}{row}>{lookupCol[i]}{row+opponents[rCmp]},1,"
                        cmpF += f"IF({lookupCol[i]}{row}={lookupCol[i]}{row+opponents[rCmp]},0.5,0)),0.5)"
                        sh.cell(row, startCol[i]+2*nCmps+colInc).value = cmpF
                        colInc += 1
                cursorRow += 1
                row += 1
        borderCols = [12, 18]
        leftBorder = Border(left=Side(style='medium',color="000000"))
        for r in range(1, row):
            for c in borderCols:
                sh.cell(r, c).border = leftBorder
        allCols = 1
        for h in headers:
            if type(h) is str:
                allCols += 1
            else:
                allCols += h[1]
        for r in range(2, len(rounds)*nTbl*self.decks, nTbl):
            for c in range(1, allCols):
                bds = sh.cell(r, c).border
                sh.cell(r, c).border = Border(left=bds.left, bottom=self.thinLine.top)



    # Roster sheet
    # Also the final result
    def rosterSheet(self):
        self.log.debug('Creating Roster Sheet')
        headers = ['Pair #', 'Player 1', 'Player 2', 'IMP', 'MP']
        self.pdf.roster(self.log, self.pairs, headers[:-1])

        sh = self.wb.create_sheet('Roster')
        row = self.headerRow(sh, headers)
        totalPlayed = int((self.pairs + self.pairs % 2) / 2) * self.decks * (self.pairs - 1)
        for i in range(self.pairs):
            sh.cell(i+row, 1).value = i+1
            sh.cell(i+row, 2).value = self.placeHolderName()
            sh.cell(i+row, 3).value = self.placeHolderName()
            sh.column_dimensions['B'].width = 25
            sh.column_dimensions['C'].width = 25
            IMPsum1 = f"=SUMIF('By Board'!$D$3:$D${totalPlayed+2},\"={i+1}\",'By Board'!$L$3:$L${totalPlayed+2})"
            MPsum1 = f"=SUMIF('By Board'!$D$3:$D${totalPlayed+2},\"={i+1}\",'By Board'!$N$3:$N${totalPlayed+2})"
            if self.pairs % 2 != 0 or i != self.pairs - 1:
                IMPsum2 = f"SUMIF('By Board'!$E$3:$E${totalPlayed+2},\"={i+1}\",'By Board'!$M$3:$M${totalPlayed+2})"
                MPsum2 = f"SUMIF('By Board'!$E$3:$E${totalPlayed+2},\"={i+1}\",'By Board'!$O$3:$O${totalPlayed+2})"
            else:
                IMPsum2=0
                MPsum2=0
            sh.cell(i+row, 4).value = f"{IMPsum1}+{IMPsum2}"
            sh.cell(i+row, 4).number_format = '#0.00'
            sh.cell(i+row, 5).value = f"{MPsum1}/{self.decks*(self.pairs-1)}+{MPsum2}/{self.decks*(self.pairs-1)}"
            sh.cell(i+row, 5).number_format = '0.0%'
        
        IMPRow = self.pairs + row + 2
        sh.cell(IMPRow, 1).value = 'Array Formula below, remove single quote'
        IMPRow += 1
        #arrayRange = f'{self.rc2a1(IMPRow,1)}:{self.rc2a1(IMPRow+self.pairs-1,5)}'
        #formulaTxt = f"=SORT({self.rc2a1(row+1, 1)}:{self.rc2a1(row+self.pairs-1,5)},4,-1)"
        #sh[self.rc2a1(IMPRow,1)] = ArrayFormula(arrayRange, formulaTxt)
        sh.cell(IMPRow,1).value = f"'=SORT({self.rc2a1(row, 1)}:{self.rc2a1(row+self.pairs-1,5)},4,-1)"
        for i in range(self.pairs):
            sh.cell(IMPRow+i,4).number_format = "#0.00"
            sh.cell(IMPRow+i,5).number_format = "0.00%"
        MPRow = self.pairs + IMPRow + 2
        sh.cell(MPRow,1).value = f"'=SORT({self.rc2a1(row, 1)}:{self.rc2a1(row+self.pairs-1,5)},5,-1)"
        for i in range(self.pairs):
            sh.cell(MPRow+i,4).number_format = "#0.00"
            sh.cell(MPRow+i,5).number_format = "0.00%"

        # Check to make sure IMPs add up to zero
        ft = Font(bold=True,color="FF0000")
        topBorder = Border(top=Side(style='thin', color="FF0000"))
        sh.cell(self.pairs+2, 4).value=f'=SUM(D2:D{self.pairs+1})'
        sh.cell(self.pairs+2, 4).number_format = '#0.00'
        sh.cell(self.pairs+2, 5).value=f'=AVERAGE(E2:E{self.pairs+1})'
        sh.cell(self.pairs+2, 5).number_format = '0.00%'
        sh.cell(self.pairs+2, 4).font = ft
        sh.cell(self.pairs+2, 4).border = topBorder
        sh.cell(self.pairs+2, 5).font = ft
        sh.cell(self.pairs+2, 5).border = topBorder

    def Traveler(self):
        self.log.debug('Creating Traveler Template Sheet')
        headers = ['Round', 'NS', 'EW', 'Contract', 'By', 'Result', 'NS', 'EW']
        colWidthTbl = [8, 8, 8, 30, 8, 10, 8, 8]
        sh = self.wb.create_sheet('Traveler Template')
        sh.cell(1, 1).value = 'Board #'
        titleFont = Font(size=self.HeaderFont.size + 8, bold=True)
        sh.cell(1, 1).font = titleFont
        sh.merge_cells(f'{sh.cell(1,1).coordinate}:{sh.cell(1,len(headers)).coordinate}')
        row = self.headerRow(sh, headers, 3)
        side=Side(style='thin',color='000000')
        border=Border(top=side,left=side,bottom=side,right=side)
        for i in range(self.pairs - 1):
            sh.cell(i+4, 1).value = i+1
            sh.cell(i+4, 1).alignment = self.centerAlign
            sh.cell(i+4, 1).font = self.HeaderFont
            for j in range(len(headers)):
                sh.cell(i+4, j+1).border = border
        for c in range(len(headers)):
            sh.column_dimensions[chr(ord('A')+c)].width = colWidthTbl[c]

def howellFromJson(log, pairs, fake, jsonfile):
    jIO = jsonIO.JsonIO(pairs, log)
    tourney = jIO.load(jsonfile)
    if tourney:
        doc = Howell(log, fake)
        doc.init(pairs, tourney['Rounds'])
        doc.saveByRound(tourney['Arrangement'])
        doc.saveByTable(tourney['Arrangement'])
        doc.saveByBoard(tourney['Arrangement'])
        doc.save()

if __name__ == '__main__':
    log = setlog('howell', None)
    parser = argparse.ArgumentParser()
    parser.add_argument('-p', '--pair', type=int, choices=range(5,15), help='# of pairs in the tournament')
    parser.add_argument('-f', '--fake', type=bool, default=False, help='Fake scores to test the spreadsheet')
    parser.add_argument('-d', '--debug', type=str, default='INFO')
    parser.add_argument('-j', '--jsonfile', type=str)
    args = parser.parse_args()
    for l in [['INFO', logging.INFO], ['DEBUG', logging.DEBUG], ['ERROR', logging.ERROR]]:
        if args.debug.upper() == l[0]:
            log.setLevel(l[1])
            break

    if args.pair: 
        howellFromJson(log, args.pair, args.fake, args.jsonfile)
    elif args.pair is None:
        for p in range(5,15):
            howellFromJson(log, p, args.fake, args.jsonfile)

