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
		self.thinLine = Border(top=Side(style='thin', color="000000"))

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

	def fakeScore(self, sh, row, col):
		if random.random() < 0.90:
			pickSide = col if random.random() >= 0.5 else col+1
			score = random.randint(2,80)*10
			sh.cell(row, pickSide).value = score
		else:
			sh.cell(row, col).value = 'Avg'
			sh.cell(row, col+1).value = 'Avg'
