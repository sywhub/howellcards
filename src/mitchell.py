#!/usr/bin/env python3
# Generate Mitchell movements
# A 4-table Mitchell uses "Square" arrangement, found at MIT web site
# (But for 4-table tournaments, really should use Howell)
#
# Even-numbered-table games have 2 styles: skip-rouind or "relay plus sharing".  We do skip.
# Four-table games uses MIT "square" movement.  It's not really better than Howell.
#
import argparse
import logging
from maininit import setlog
from openpyxl import Workbook
from openpyxl.styles import Font
import pdf
from docset import PairGames
import datetime

# Pairs are internally numbered 1,3,5,... for EW pairs and 2,4,6,... for NS
# Pair 0 is the sit-out phantom pair
# Externally, they are number 1 to n for both NS and EW sides
#
# There are "internal" pair numbers that go uniquely from 0 to n.
# The external ones are "NS 1 to n" and "ES 1 to m", where there are same pair numbers in NS and EW group.
# Internal even number pairs map to EW and odd to NS.  So that when there are odd number of pairs, the last EW pair
# sits at the sit-out table.
#
class Mitchell(PairGames):
    def __init__(self, log, p, b, sq, f, nameFile):
        super().__init__(log, p)
        self.decks = b
        self.tables = (self.pairs + 1) // 2
        self.oddPairs = self.pairs % 2 == 1
        self.square = sq
        self.fake = f
        self.pdf = pdf.PDF()
        self.wb = Workbook()

        self.loadNames(nameFile, {'File': f'mitchell{self.pairs}x{self.decks}{"xF" if self.fake else ""}',
                    'Tournament': f'Mitchell Movement for {self.pairs} Pairs, {self.decks} boards round'})
        self.pdf.HeaderFooterText(f'{self.notice} {datetime.date.today().strftime("%b %d, %Y")}.',
            self.nameObj['Tournament'])
        # initData must be the first one
        self.initData()
        self.meta()

    # Internal even number pairs are EW and odd NS
    def pairSide(self, n):
        return ['EW', 'NS'][n % 2]

    # translate internal pair number to external
    # Note NS and EW have same pair numbering
    def pairN(self, n):
        return n // 2 + 1

    # Full external name of a pair
    # [NS | EW] <pair number> (member 1 + member 2)
    def pairID(self, n):
        sitout = self.oddPairs != 0 and (n == self.pairs)
        idStr = f"{self.pairSide(n)} {self.pairN(n)}" if not sitout else self.SITOUT
        if not sitout and len(self.nameObj['Players']) > 0:
            idStr += f' ({self.nameObj['Players'][n]})'
        return idStr

    # assign internal pair number by table and round
    def NSPair(self, r, t):
        return t * 2 + 1

    # EW pairs internal numbers are even, zero-based.
    def EWPair(self, r, t):
        x = (r + t) % self.tables * 2
        return x
    
    # board base number for each round/table
    def boardIdx(self, r, t):
        return ((t - r + self.tables) % self.tables) * self.decks

    # The sit-out (phantom) pair is the last of the NS pairs
    def ifSitout(self, t, ns, ew):
        return (self.oddPairs) and (ns == self.pairs)

    # Initialize board and round tables for various internal code
    def initData(self):
        self.boardData = {}
        if self.pairs == 8 and self.square:
            self.loadSquare()   # square Mitchell
            self.initRounds()
        elif self.tables % 2 == 0: 
            self.loadEven()
        else:  # standard Mitchell
            for r in range(self.tables): # round
                for t in range(self.tables): # table
                    b = self.boardIdx(r, t)
                    for bset in range(self.decks):
                        if (b + bset) not in self.boardData:
                            self.boardData[b+bset] = []
                        ns = self.NSPair(r, t)
                        ew = self.EWPair(r, t)
                        if not self.ifSitout(t, ns, ew):
                            self.boardData[b+bset].append([r, t, self.NSPair(r, t), self.EWPair(r, t)])
            self.initRounds()
        self.checkBoardData()

    def roster(self):
        self.log.debug('Roster sheet and PDF')
        self.rosterSheet()
        self.rosterPDF()
    
    # a notice on public domain
    # Then the meta info about this tournament
    # Last a list of names for pairs
    def rosterSheet(self):
        ws = self.wb.active # the first tab
        ws.title = 'Roster'

        row = self.sheetMeta(ws, self.metaData) + 2
        start = 1
        for s in ['NS', 'EW']:
            ws.cell(row, 1).value =  f'{s} Pairs'
            ws.cell(row, 1).font = self.HeaderFont
            ws.cell(row, 1).alignment = self.centerAlign
            ws.merge_cells(f'{ws.cell(row,1).coordinate}:{ws.cell(row,3).coordinate}')
            ws.cell(row, 4).value = 'MP'
            ws.cell(row, 4).font = self.HeaderFont
            ws.cell(row, 4).alignment = self.centerAlign
            ws.cell(row, 5).value = 'IMP'
            ws.cell(row, 5).font = self.HeaderFont
            ws.cell(row, 5).alignment = self.centerAlign
            row += 1
            avgStart = row  # remember this row
            for p in range(start, self.pairs, 2):
                useNames = self.pairNames(p)
                ws.cell(row, 1).font = self.HeaderFont
                ws.cell(row, 1).alignment = self.centerAlign
                ws.cell(row, 1).value = self.pairN(p)
                ws.cell(row, 2).value = useNames[0]
                ws.cell(row, 3).value = useNames[1]
                row += 1
            start -= 1

            # draw a line
            for i in range(6):
                ws.cell(row-1, i+1).border = self.bottomLine

            ws.cell(row, 3).value = 'Average'
            ws.cell(row, 3).font = self.noChangeFont

            ws.cell(row, 4).value = f'=AVERAGE({self.rc2a1(avgStart, 4)}:{self.rc2a1(row-1,4)})'
            ws.cell(row, 5).value = f'=SUM({self.rc2a1(avgStart, 6)}:{self.rc2a1(row-1,6)})'
            ws.cell(row,4).number_format = "0.00%"
            ws.cell(row,5).number_format = "#0.0"
            ws.cell(row,4).font = self.noChangeFont
            ws.cell(row,5).font = self.noChangeFont
            row += 2
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 30
        
    def rosterPDF(self):
        self.pdf.add_page()
        self.pdf.headerFooter()
        self.pdf.meta(self.metaData)
        self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.rosterPt) 
        h = self.pdf.lineHeight(self.pdf.font_size_pt)
        title = 'Player Pairs'
        x = self.pdf.setHCenter(self.pdf.get_string_width(title))
        y = self.pdf.get_y() + 2 * h
        self.pdf.set_xy(x, y)
        self.pdf.cell(text=title)
        widths = [1, 2, 2]
        y +=  h
        leftM = (self.pdf.w - sum(widths)) / 2
        self.pdf.set_xy(leftM, y)
        self.pdf.set_font(self.pdf.sansSerifFont, size=(self.pdf.bigPt if self.pairs < 19 else self.pdf.linePt)) 
        h = self.pdf.lineHeight(self.pdf.font_size_pt)
        start = 1
        for s in ['NS', 'EW']:
            self.pdf.set_font(style='BI')
            self.pdf.cell(5, h, text=f'{s} Pairs', align='L')
            self.pdf.set_font(style='')
            y += h
            self.pdf.set_xy(leftM, y)
            saveFont = self.pdf.font_family
            self.pdf.set_font(self.pdf.chineseFont)
            for p in range(start, self.pairs, 2):
                useNames = self.pairNames(p)
                self.pdf.cell(widths[0], h, text=f'{self.pairN(p)}', align='C', border=1)
                self.pdf.cell(widths[1], h, text=useNames[0], align='C', border=1)
                self.pdf.cell(widths[2], h, text=useNames[1], align='C', border=1)
                y += h
                self.pdf.set_xy(leftM, y)
            start -= 1
            self.pdf.set_font(saveFont)
        self.pdf.set_font(self.pdf.serifFont, style='B', size=self.pdf.rosterPt) 
        return

    # roster shows meta info first
    # Not doing meta sheet
    def meta(self):
        self.log.debug('Meta')
        self.metaData = {'Title': 'Mitchell Tournament', 'Info': []}
        self.metaData['Info'].append(('Pairs', self.pairs))
        self.metaData['Info'].append(('Tables', self.tables))
        self.metaData['Info'].append(('Rounds', self.pairs // 2 - 1))
        self.metaData['Info'].append(('Boards per round', self.decks))
        if self.tables % 2 == 0:
            if self.tables == 4 and self.square:
                self.metaData['Info'].append(('Square Movement', ''))
            else:
                self.metaData['Info'].append((f'EW pairs skip a table after round {self.tables // 2}',""))

    def setTableTexts(self):
        self.log.debug('Setting Table borders')
        if self.pairs == 8 and self.square:
            # The trade-off of square is the iregularity of movements
            ewText = ['R2 to T2/EW, R3 to T3/EW, R4 to T2/EW',
                        'R2 to T1/EW, R3 to T4/EW, R4 to T1/EW',
                        'R2 to T4/EW, R3 to T1/EW, R4 to T4/EW',
                        'R2 to T3/EW, R3 to T2/EW, R4 to T3/EW']
            nsText = ['Stay here. Boards: R2 to T4, R3 to T2, R4 to T4',
                        'Stay here. Boards: R2 to T3, R3 to T1, R4 to T3', 
                        'Stay here. Boards: R2 to T2, R3 to T4, R4 to T2', 
                        'Stay here. Boards: R2 to T1, R3 to T3, R4 to T1'] 
        else:
            nsText = []
            ewText = []
            for t in range(self.tables):
                ewText.append(f'Move to Table {(t+self.tables-1) % self.tables + 1} EW')
                nsText.append(f'Stay Here, Boards to T{t + 2 if t < self.tables else 1}')
        self.Tables(nsText, ewText, True)

    # Square arrangement is not programatic.
    def loadSquare(self):
        self.log.debug('Load Square data')
        self.sqSetup = {
            # Primary key is the table number
            # Pair numbering in this data is separated by NS/EW
            # Each "board set" is n boards, as dedicated by command line argument
            0: [{'Round': 0, 'NS': 1, 'EW': 1, 'Board': 0},   # round, NS, EW, boardSet #
                {'Round': 1, 'NS': 1, 'EW': 2, 'Board': 3},
                {'Round': 2, 'NS': 1, 'EW': 4, 'Board': 2},
                {'Round': 3, 'NS': 1, 'EW': 3, 'Board': 1},],
            1: [{'Round': 0, 'NS': 2, 'EW': 2, 'Board': 1},
                {'Round': 1, 'NS': 2, 'EW': 1, 'Board': 2},
                {'Round': 2, 'NS': 2, 'EW': 3, 'Board': 3},
                {'Round': 3, 'NS': 2, 'EW': 4, 'Board': 0},],
            2: [{'Round': 0, 'NS': 3, 'EW': 3, 'Board': 2},
                {'Round': 1, 'NS': 3, 'EW': 4, 'Board': 1},
                {'Round': 2, 'NS': 3, 'EW': 2, 'Board': 0},
                {'Round': 3, 'NS': 3, 'EW': 1, 'Board': 3},],
            3: [{'Round': 0, 'NS': 4, 'EW': 4, 'Board': 3},
                {'Round': 1, 'NS': 4, 'EW': 3, 'Board': 0},
                {'Round': 2, 'NS': 4, 'EW': 1, 'Board': 1},
                {'Round': 3, 'NS': 4, 'EW': 2, 'Board': 2}]}
        self.boardData = {}
        for t,tbl in self.sqSetup.items():
            for r in tbl:
                r['Board'] = [r['Board']*self.decks + x for x in range(self.decks)]
                r['NS'] = (r['NS'] - 1) * 2 + 1
                r['EW'] = (r['EW'] - 1) * 2 
                for b in r['Board']:
                    if b not in self.boardData:
                        self.boardData[b] = []
                    self.boardData[b].append([r['Round'], t, r['NS'], r['EW']])

    # Even number of tables not 7 or 8 pairs
    # Basically 11, 12, 15, and 16 pairs
    # Key point is skipping a round at mid-way
    def loadEven(self):
        self.roundData = {}
        for r in range(self.tables - 1):
            self.roundData[r] = {}
            for t in range(self.tables):
                bIdx = self.boardIdx(r , t)
                blist = [bIdx+x for x in range(self.decks)]
                ns = self.NSPair(r, t)
                ew = self.EWPair(r, t)
                if not self.ifSitout(t, ns, ew):
                    self.roundData[r][t] = {'NS': ns, 'EW': ew, 'Board': blist}
                    if r >= self.tables // 2:
                        self.roundData[r][t]['EW'] = self.EWPair(r+1,t)
        self.boardData = {}
        for r,tbl in self.roundData.items():
            for t,d in tbl.items():
                for b in d['Board']:
                    if b not in self.boardData:
                        self.boardData[b] = []
                    self.boardData[b].append([r, t, d['NS'], d['EW']])


    def results(self):
        self.log.debug('Add results to Roster')
        sh = self.wb['Roster']
        lastRows = 0
        row = len(self.metaData['Info']) + 4 + 1    # Copyright, Title, a Spacer, and score table row, plus sheet is 1-based

        for b in self.boardData.values():
            lastRows += len(b)
        lastRows -= 1  # inclusive
        # Compute the max MP pts for each pair
        # Scan boardData, find the board the pair played and count how many times that board was played
        # Max MP pt is the one less than the number of times the board was played
        divMap = {x: 0 for x in range(self.pairs)}
        for v in self.boardData.values():
            for p in v:
                divMap[p[2]] += len(v) - 1
                divMap[p[3]] += len(v) - 1
        for s in [1, 0]:
            toN = self.pairs
            for p in range(s, toN, 2):
                pName = self.pairN(p+1)
                if pName == self.SITOUT:
                    continue
                ifRange = f"'By Board'!{self.rc2a1(3, 5-s)}:{self.rc2a1(3+lastRows,5-s)}"
                impRange = f"'By Board'!{self.rc2a1(3, 14-s)}:{self.rc2a1(3+lastRows,14-s)}"
                sumRange = f"'By Board'!{self.rc2a1(3, 18-s)}:{self.rc2a1(3+lastRows,18-s)}"
                sh.cell(row,4).value=f"=SUMIF({ifRange},\"=\"&{self.rc2a1(row, 1)},{sumRange})/{divMap[p]}"
                sh.cell(row,5).value=f"=SUMIF({ifRange},\"=\"&{self.rc2a1(row, 1)},{impRange})"
                sh.cell(row,4).number_format = "0.00%"
                sh.cell(row,5).number_format = "#0.0"
                row += 1
            row += 3

    # Output into filesystem
    def save(self):
        import os
        here = os.path.dirname(os.path.abspath(__file__))
        fn = f'{here}/../{self.nameObj['File']}'
        if self.pairs == 8 and self.square:
            fn += 'Sq'
        self.log.debug(f'Save files: {fn}')
        self.wb.save(f'{fn}.xlsx')
        self.pdf.output(f'{fn}.pdf')
        print(f'Saved {fn}.{{xlsx,pdf}}')

    def main(self):
        self.log.debug('Main goes')
        self.pdf.instructions(self.log, 'mitchellInstructions.txt')
        self.roster()
        self.results()
        self.roundTab()
        self.boardTab()
        self.IMPTable() # static sheet
        self.ScoreTable()   # static sheet, produced to aid human TD, not used elsewhere.
        #self.idTags()  # PDF only
        self.setTableTexts()  # PDF only
        self.Travelers()  # PDF only
        self.Journal()  # PDF only
        #self.Pickups()  # PDF only
        self.save()
        return


if __name__ == '__main__':
    log = setlog('mitchell', None)
    def mitchell_check(value):
        ivalue = int(value)
        if ivalue in [11,15,16]:
            raise argparse.ArgumentTypeError(f"Cannot have even number of tables")
        return ivalue

    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--debug', type=str, default='INFO', help='Debug level, INFO, DEBUG, ERROR')
    parser.add_argument('-b', '--boards', type=int, choices=range(1,7), default=4, help='Boards per round')
    parser.add_argument('-p', '--pair', type=int, choices=range(8,25), default=8, help='Number of pairs')
    parser.add_argument('-f', '--fake', action='store_true', help='Fake scores to test the spreadsheet')
    parser.add_argument('-n', '--names', type=str, default="", help='Names in the tournament')
    parser.add_argument('-s', '--square', action='store_true', help='Use square movement for 4 tables')
    args = parser.parse_args()
    for l in [['INFO', logging.INFO], ['DEBUG', logging.DEBUG], ['ERROR', logging.ERROR]]:
        if args.debug.upper() == l[0]:
            log.setLevel(l[1])
            break
    mitchell = Mitchell(log, args.pair, args.boards, args.square, args.fake, args.names)
    mitchell.main()
