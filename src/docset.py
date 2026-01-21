#!/usr/bin/env python3
# Mainly spreadsheet class to Howell tournaments into Excel templates
# Also produce PDF file the event
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
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
		for col in range(len(headers)):
			sh.cell(row, col+1).value = headers[col]
			sh.cell(row, col+1).font = self.HeaderFont
			sh.cell(row, col+1).alignment = self.centerAlign
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
			sh.cell(1, c).value = 'Not Vulnerable'
			sh.cell(1, c).font = self.HeaderFont
			sh.merge_cells(f'{sh.cell(1, c).coordinate}:{sh.cell(1, c+2).coordinate}')
			sh.cell(1, c).alignment = self.centerAlign
		for c in [6, 14]:
			sh.cell(1, c).value = 'Vulnerable'
			sh.cell(1, c).font = self.HeaderFont
			sh.merge_cells(f'{sh.cell(1, c).coordinate}:{sh.cell(1, c+2).coordinate}')
			sh.cell(1, c).alignment = self.centerAlign
		headers = ['Contract', 'Result', '', 'X', 'XX', '', 'X', 'XX']
		row = self.headerRow(sh, headers, 2)
		self.scorePenalty(sh, row, len(headers)+2, headers[2:])
		for i in range(1,8):
			for trump in self.trumps:
				sh.cell(row, 1).value = f'{i} {trump}'
				for j in range(0, 8-i):
					sh.cell(row, 2).value = j
					sh.cell(row, 2).number_format = '+#0;-#0;0'
					sh.cell(row, 3).value = self.score(i, trump, j, False, 0)
					sh.cell(row, 4).value = self.score(i, trump, j, False, 1)
					sh.cell(row, 5).value = self.score(i, trump, j, False, 2)
					sh.cell(row, 6).value = self.score(i, trump, j, True, 0)
					sh.cell(row, 7).value = self.score(i, trump, j, True, 1)
					sh.cell(row, 8).value = self.score(i, trump, j, True, 2)
					for c in range(3,9):
						sh.cell(row, c).number_format = '#0'
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

		dblMul = 2**dbl	# 2 to the power of "dbl" which is 0, 1, or 2
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
		else:	# made game
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
		return f"{chr(c-1+ord('A'))}{r}"


class HowellDocSet(DupBridge):
	def __init__(self, log, toFake=False):
		super().__init__(log)
		self.notice = 'For public domain. No rights reserved. Generated on'
		self.travelerText = '0 for contract made, "Avg" for incomplete'
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
		self.pairs = pairs;
		if pairs <= 6:
			self.decks = 3
		else:
			self.decks = 2

		# meta data
		tourneyMeta = [['Howell Arrangement (IMP)'],
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

		self.Instructions(ws, len(tourneyMeta)+11)

		self.pdf.noright(self.log, f'{self.notice} {datetime.date.today().strftime("%b %d, %Y")}.')
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
		vulTbl = ['None', 'NS', 'EW', 'Both']
		headers = ['Round', 'Table', 'NS', 'EW', 'Board', 'Vul', 'Contract', 'By', 'Result', 'NS Score', 'EW Score']
		self.log.debug('Saving by Round')
		sh = self.wb.create_sheet('By Round')
		row = self.headerRow(sh, headers)
		sh.column_dimensions[chr(headers.index('Contract')+ord('A'))].width = 30
		sh.column_dimensions['H'].width = 15
		sh.column_dimensions['I'].width = 15
		thinLine = Border(top=Side(style='thin', color="000000"))
		thickLine = Border(top=Side(style='medium', color="F80000"))
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
					sh.cell(row, 6).value = vulTbl[(b-1)%4]
					sh.cell(row, 6).alignment = self.centerAlign
					if self.fakeResult:
						if random.random() < 0.90:
							pickSide = 10 if random.random() >= 0.5 else 11
							score = random.randint(2,80)*10
							sh.cell(row, pickSide).value = score
						else:
							sh.cell(row, 10).value = 'Avg'
							sh.cell(row, 11).value = 'Avg'
					row += 1
				for c in range(2,12):
					sh.cell(row, c).border = thinLine
			for c in range(1,12):
				sh.cell(row, c).border = thickLine

	# Present the same data table-oriented
	def saveByTable(self, rounds):
		self.log.debug('Saving by Table')
		nTbl = len(rounds[0])
		nRounds = len(rounds)
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
		vulTbl = ['None', 'NS', 'EW', 'Both']
		self.log.debug('Saving by Board')
		sh = self.wb.create_sheet('By Board', 2)	# insert it as the 2nd sheet
		headers = ['Board', 'Round', 'Table', 'NS', 'EW', 'Vul', 'Contract', 'By', 'Result', 'NS', 'EW']
		nTbl = len(rounds[0])
		leftBorder = Border(left=Side(style='medium',color="000000"))
		noChangeFont = Font(bold=True, italic=True, color='FF0000')
		# The contract column should be wider for data entry
		sh.column_dimensions[chr(headers.index('Contract')+ord('A'))].width = 30

		# first row setup some spanning column headers
		mergeHdrs = [['Contract Pt', 2], ['IMP', 2], ['IMP Calculation', nTbl*2],
			  ['MP', 2], ['MP Calculation', nTbl*2]]
		impHdrs = ['NS', 'EW', 'NS Net', 'EW Net', 'NS Pair-Wise', 'EW Pair-Wise']
		mpHdrs = ['NS', 'EW', 'NS MP Score', 'EW MP Score']
		cStart = 10
		for h in mergeHdrs:
			sh.cell(1, cStart).value = h[0]
			sh.cell(1, cStart).font = noChangeFont
			sh.cell(1, cStart).alignment = self.centerAlign
			sh.merge_cells(f'{sh.cell(1,cStart).coordinate}:{sh.cell(1,cStart+h[1]-1).coordinate}')
			cStart += h[1]
		sh.cell(1, 10).font = self.HeaderFont
		cStart = len(headers) + 1
		for h in impHdrs[:4]:
			sh.cell(2, cStart).value = h
			sh.cell(2, cStart).font = noChangeFont
			sh.cell(2, cStart).alignment = self.centerAlign
			cStart += 1
		for h in impHdrs[4:]:
			sh.cell(2, cStart).value = h
			sh.cell(2, cStart).font = noChangeFont
			sh.cell(2, cStart).alignment = self.centerAlign
			sh.merge_cells(f'{sh.cell(2,cStart).coordinate}:{sh.cell(2,cStart+nTbl-2).coordinate}')
			cStart += nTbl - 1
		for h in mpHdrs[:2]:
			sh.cell(2, cStart).value = h
			sh.cell(2, cStart).font = noChangeFont
			sh.cell(2, cStart).alignment = self.centerAlign
			cStart += 1
		for h in mpHdrs[2:]:
			sh.cell(2, cStart).value = h
			sh.cell(2, cStart).font = noChangeFont
			sh.cell(2, cStart).alignment = self.centerAlign
			sh.merge_cells(f'{sh.cell(2,cStart).coordinate}:{sh.cell(2,cStart+nTbl-2).coordinate}')
			cStart += nTbl - 1


		row = self.headerRow(sh, headers, 2)

		# build a datastructure for ease of navigation
		# just pivotig the source data
		# board keyed by board set #, value = [(round, table, NS, EW), ...]
		boards = {}
		for r,t in enumerate(rounds):
			for tbl, p in enumerate(t):
				if p['Board'] not in boards:
					boards[p['Board']] = []
				boards[p['Board']].append((r, tbl, p['NS'], p['EW']))

		list = sorted([x for x in boards.keys()])
		self.pdf.travelers(self.log, self.decks, boards)
		for b in list:
			for i in range(self.decks):
				bIdx = b*self.decks+i
				sh.cell(row, 1).value = bIdx + 1
				# vulnerability
				vulShift = bIdx // 4
				vulIdx = (bIdx + vulShift) % 4
				for r in range(len(boards[b])):
					# tbls: [round, table, NS, EW]
					tbls = boards[b][r]
					# this part just reference the "mother sheet"
					rBase = tbls[0]*nTbl*self.decks+2
					sh.cell(row, 2).value = f"='By Round'!{self.rc2a1(rBase, 1)}"
					for c in range(2, 11):
						bBase = rBase + tbls[1]*self.decks
						if c >= 6:
							bBase += bIdx % self.decks
						a1 = self.rc2a1(bBase, c if c < 5 else c + 1)
						cVal = f"'By Round'!{a1}"
						if c >= 6:
							bcheck = f'=IF(ISBLANK({cVal}),"",{cVal})'
						else:
							bcheck= f'={cVal}'
						sh.cell(row, 1+c).value = bcheck 
					sh.cell(row, 6).value = vulTbl[vulIdx]
					sh.cell(row, 6).alignment = self.centerAlign

					# There are steps to calculate IMP for each board
					# Here are two columns for the end result
					avgRange = f'{sh.cell(row, 16).coordinate}:{sh.cell(row, 16+nTbl-2).coordinate}'
					sh.cell(row, 12).value = f'=IFERROR(AVERAGE({avgRange}),"")'
					sh.cell(row, 12).number_format = '#0.00'
					avgRange = f'{sh.cell(row, 16+nTbl-1).coordinate}:{sh.cell(row, 16+nTbl-1+nTbl-2).coordinate}'
					sh.cell(row, 13).value = f'=IFERROR(AVERAGE({avgRange}),"")'
					sh.cell(row, 13).number_format = '#0.00'
					sumRange = f'{sh.cell(row, 24).coordinate}:{sh.cell(row, 24+nTbl-2).coordinate}'
					sh.cell(row, 22).value = f'=IFERROR(SUM({sumRange})/{nTbl-1},"")'
					sh.cell(row, 22).number_format = '0.0%'
					sumRange = f'{sh.cell(row, 24+nTbl-1).coordinate}:{sh.cell(row, 24+nTbl-1+nTbl-2).coordinate}'
					sh.cell(row, 23).value = f'=IFERROR(SUM({sumRange})/{nTbl-1},"")'
					sh.cell(row, 23).number_format = '0.0%'

					# IMP Computation sequence
					# 1. For each side, record their "net" raw scores.  Negative if the other side scored
					sh.cell(row, 14).value = f'=IF(ISNUMBER(J{row}),J{row},IF(ISNUMBER(K{row}),-K{row},""))'
					sh.cell(row, 15).value = f'=IF(ISNUMBER(K{row}),K{row},IF(ISNUMBER(J{row}),-J{row},""))'
					# 2. Compare to all other pairs, on the same side, and use the difference to lookup IMPs
					# competitors are all other pairs
					competitors = [x for x in range(nTbl)]
					competitors.remove(r)
					competitors = [x - r for x in competitors]	# turn it into relative ref to "this" column
					colInc = 2	# distance to the previous section
					for j in competitors:
						# NS comparisons
						cond=f'AND(ISNUMBER(N{row}),ISNUMBER(N{row+j}))'
						lookup=f"VLOOKUP(ABS(N{row}-N{row+j}),'IMP Table'!$A$2:$C$26,3)*SIGN(N{row}-N{row+j})"
						formula=f'=IF({cond},{lookup},"")'
						sh.cell(row, 14+colInc).value = formula
						sh.cell(row, 22+colInc).value = f'=IF(ISNUMBER(N{row}),IF(ISNUMBER(N{row+j}),IF(N{row}>N{row+j},1,if(N{row}=N{row+j},0.5,0)),0.5),0.5)'
						# EW comparisons
						cond=f'AND(ISNUMBER(O{row}),ISNUMBER(O{row+j}))'
						lookup=f"VLOOKUP(ABS(O{row}-O{row+j}),'IMP Table'!$A$2:$C$26,3)*SIGN(O{row}-O{row+j})"
						formula=f'=IF({cond},{lookup},"")'
						sh.cell(row, 14+nTbl-1+colInc).value = formula
						sh.cell(row, 22+nTbl-1+colInc).value = f'=IF(ISNUMBER(O{row}),IF(ISNUMBER(O{row+j}),IF(O{row}>O{row+j},1,if(O{row}=O{row+j},0.5,0)),0.5),0.5)'
						colInc += 1
					row += 1
		borderCols = [12, 14, 14+nTbl*2, 14+nTbl*2+2]
		for r in range(1, row):
			for c in borderCols:
				sh.cell(r, c).border = leftBorder


	# Roster sheet
	# Also the final result
	def rosterSheet(self):
		self.log.debug('Creating Roster Sheet')
		headers = ['Pair #', 'Player 1', 'Player 2', 'IMP', 'MP', 'MP Sum %']
		self.pdf.roster(self.log, self.pairs, headers[:-1])

		sh = self.wb.create_sheet('Roster')
		row = self.headerRow(sh, headers)
		sh.merge_cells(f'{self.rc2a1(1,6)}:{self.rc2a1(1,7)}')
		totalPlayed = int((self.pairs + self.pairs % 2) / 2) * self.decks * (self.pairs - 1)
		for i in range(self.pairs):
			sh.cell(i+row, 1).value = i+1
			sh.cell(i+row, 2).value = self.placeHolderName()
			sh.cell(i+row, 3).value = self.placeHolderName()
			sh.column_dimensions['B'].width = 25
			sh.column_dimensions['C'].width = 25
			IMPsum1 = f"=SUMIF('By Board'!$D$3:$D${totalPlayed+2},\"={i+1}\",'By Board'!$L$3:$L${totalPlayed+2})"
			MPsum1 = f"=SUMIF('By Board'!$D$3:$D${totalPlayed+2},\"={i+1}\",'By Board'!$V$3:$V${totalPlayed+2})"
			if self.pairs % 2 != 0 or i != self.pairs - 1:
				IMPsum2 = f"SUMIF('By Board'!$E$3:$E${totalPlayed+2},\"={i+1}\",'By Board'!$M$3:$M${totalPlayed+2})"
				MPsum2 = f"SUMIF('By Board'!$E$3:$E${totalPlayed+2},\"={i+1}\",'By Board'!$W$3:$W${totalPlayed+2})"
			else:
				IMPsum2=0
				MPsum2=0
			sh.cell(i+row, 4).value = f"{IMPsum1}+{IMPsum2}"
			sh.cell(i+row, 4).number_format = '#0.00'
			sh.cell(i+row, 5).value = f"=({self.rc2a1(i+row, 6)}+{self.rc2a1(i+row, 7)})/{self.decks*(self.pairs-1)}"
			sh.cell(i+row, 5).number_format = '0.0%'
			sh.cell(i+row, 6).value = f"={MPsum1}"
			sh.cell(i+row, 6).number_format = '0.0%'
			sh.cell(i+row, 7).value = f"={MPsum2}"
			sh.cell(i+row, 7).number_format = '0.0%'
		# Check to make sure IMPs add up to zero
		ft = Font(bold=True,color="FF0000")
		topBorder = Border(top=Side(style='thin', color="FF0000"))
		sh.cell(self.pairs+2, 3).value='Sum to Zero'
		sh.cell(self.pairs+2, 4).value=f'=SUM(D2:D{self.pairs+1})'
		sh.cell(self.pairs+2, 4).number_format = '#0.00'
		sh.cell(self.pairs+2, 3).font = ft
		sh.cell(self.pairs+2, 3).border = topBorder
		sh.cell(self.pairs+2, 4).font = ft
		sh.cell(self.pairs+2, 4).border = topBorder

	def Traveler(self):
		self.log.debug('Creating Traveler Template Sheet')
		headers = ['Round', 'NS', 'EW', 'Contract', 'By', 'Result']
		colWidthTbl = [8, 8, 8, 30, 8, 10]
		sh = self.wb.create_sheet('Traveler Template')
		sh.cell(1, 1).value = 'Board #'
		sh.cell(1, 4).value = self.travelerText
		sh.merge_cells(f'{sh.cell(1,1).coordinate}:{sh.cell(1,2).coordinate}')
		titleFont = Font(size=self.HeaderFont.size + 8, bold=True)
		sh.cell(1, 1).font = titleFont
		sh.cell(1, 3).font = titleFont
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



	def Instructions(self, sh, row):
		text = ['There is a matching PDF file for this spreadsheet.  Take a look of that first.',
		  	'The PDF file has better traveler and movement instrucdtion sheet.  This is for plan B.',
			'Shuffle and deal number of boards based on the Board sheet.  Insert cards into slots.',
			'Make sure traveler sheet has board # written/printed on.  Fold with score side hidden.  Tuck it into the North slot for the corresponding board.',
			'Assign pair # to each participating pairs.  Usually by drawing.',
			'Seat each pair based on the Table sheet',
			'Assign North to be score keeper and South as the board caddy.',
			'At the end of the ternament, collect traveler and record into the spreadsheet.  Everything else has been automated.']
		sh.cell(row, 1).value = 'For Tournament Director/Organizer'
		for r in range(len(text)):
			sh.cell(row+r+1, 1).value = f'{r+1}. {text[r]}'
		for r in range(len(text)+1):
			sh.cell(row+r, 1).font = self.HeaderFont

