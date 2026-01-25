#!/usr/bin/env python3
from fpdf import FPDF
# Generate PDF for Howell tournament based on the data produced elsewhere
# Driven from "docset" which was made for Excel to start

class PDF(FPDF):
    margin = 0.25
    thinLine = 1/72/2   # half a point
    starRadius = 1
    rosterPt = 20
    titlePt = 18
    bigPt = 16
    headerPt = 14
    linePt = 12
    notePt = 10
    smallPt = 8
    tinyPt = 6

    def __init__(self):
        # The contructor takes no arguemnts
        # Force choices to the super
        # Letter size paper is 8.5 by 11 in
        # Taking a quarter inch off each side as margin
        super().__init__(unit='in', format='letter')
        self.set_margin(PDF.margin)
        self.fixedWidthFont = 'Courier'
        self.sansSerifFont = 'Helvetica'
        self.serifFont = 'Times'

        # Always a default font
        self.set_font(self.sansSerifFont)
        # add the first page
        self.add_page()

    # convert "point" font size to inch
    def pt2in(self, p):
        return p / 72

    # Our line spacing is 1.5 of the font size
    def lineHeight(self, p):
        return self.pt2in(p) * 1.5
    
    def setHCenter(self, linewidth):
        x = (self.w - linewidth) / 2
        return x
    
    def HeaderFooterText(self, h, f):
        self.headerText = h
        self.footerText = f

    def headerFooter(self):
        self.set_font(size=self.tinyPt)
        h = self.lineHeight(self.font_size_pt)
        w = self.get_string_width(self.headerText)
        x = self.setHCenter(w)
        self.set_xy(x, h)
        self.cell(text=self.headerText)
        w = self.get_string_width(self.footerText)
        x = self.setHCenter(w)
        y = self.eph - h * 2
        self.set_xy(x, y)
        self.cell(text=self.footerText)
    
    # put the (none) copyright notice on top
    # (Old habit dies hard)
    def noright(self, log, txt):
        self.set_font(size=PDF.tinyPt)
        h = self.lineHeight(self.font_size_pt)
        w = self.get_string_width(txt)
        x = self.setHCenter(w)
        self.set_xy(x, h)
        self.cell(text=txt)
        return

    # Footer at each page to help sorting papers
    def footer(self):
        if not hasattr(self, 'Pairs') or not hasattr(self, 'Tables'):
            return
        footerText = f'Howell Movement for {self.Pairs} Pairs and {self.Tables} Tables'
        self.set_font(size=PDF.tinyPt)
        self.set_xy(self.margin, self.h - self.pt2in(self.font_size_pt)*4)
        self.cell(text=footerText)

    # meta information for the tournament
    def meta(self, log, title, meta):
        self.set_font(style='B', size=PDF.headerPt)
        h = self.lineHeight(self.font_size_pt)
        line = h * 3
        tab = 2
        toWrite = f'{title} Setup'
        w = self.get_string_width(toWrite)
        x = self.setHCenter(w)
        self.set_xy(x, line)
        self.cell(text=toWrite)
        line += h * 3
        x = 2.5
        h = self.lineHeight(PDF.linePt)
        for m in meta:
            self.set_xy(x, line)
            self.set_font(size=PDF.linePt)
            self.cell(2, h, text=m[0])
            if len(m) > 1:
                setattr(self, m[0], m[1])
                self.set_xy(x+tab, line)
                self.set_font(style='I', size=PDF.linePt)
                self.cell(2*h, h, text=f'{m[1]}', align='R')
            line += h
        return

    # part of the meta pagee 
    # Some text for the TD/Organizer
    def instructions(self, log, fname):
        with open(fname, "r") as f:
            txt = f.read()
            txt = txt.replace('\n',' ')
            txt = txt[1:]
        self.set_font(size=PDF.linePt)
        h = self.lineHeight(self.font_size_pt)
        y = self.get_y()+2*h
        nLine = 1
        lineBreak = txt.find('#')
        while lineBreak > 0:
            t = txt[:lineBreak]
            # wish for a lambda with closure
            self.set_xy(1, y)
            self.cell(h, h, f'{nLine}.', align='R')
            self.set_xy(1+h, y)
            self.multi_cell(self.epw-2, h=h, text=t.strip())
            y = self.get_y()
            nLine += 1
            txt = txt[lineBreak+1:]
            lineBreak = txt.find('#')
        if len(txt) > 0:
            t = txt
            self.set_xy(1, y)
            self.cell(h, h, f'{nLine}.', align='R')
            self.set_xy(1+h, y)
            self.multi_cell(self.epw-2, h=h, text=t.strip())
            y = self.get_y()

    # Sign-up sheet
    def roster(self, log, rows, headers):
        self.add_page() # new page
        self.footer()
        self.set_font(self.serifFont, style='B', size=PDF.rosterPt) 
        h = self.lineHeight(self.font_size_pt)
        title = 'Player Pairings'
        x = self.setHCenter(self.get_string_width(title))
        self.set_xy(x, 2*h)
        self.cell(text=title)
        nCol = len(headers)
        # paper width minus margins from both side, minus the 1st column width
        # divide the rest evenly.
        xstart = 1  # left edge
        colWidth = (self.epw - xstart * 3 ) / (nCol - 1)
        tblCols = [colWidth]*(nCol - 1)
        tblCols.insert(0, xstart)
        self.set_font(style='B', size=PDF.rosterPt) 
        y = self.get_y() + h * 3   # go down 3 lines
        x = xstart
        self.set_xy(x, y)
        for i in range(nCol):
            self.cell(tblCols[i], h, headers[i], align='C', border=1)
        y += h
        self.set_xy(xstart, y)
        self.set_font(self.sansSerifFont, style='', size=PDF.headerPt)
        h = self.lineHeight(self.font_size_pt)
        for i in range(rows):
            self.cell(tblCols[0], h, f'{i+1}', align='C', border=1)
            for j in range(1, nCol):
                self.cell(tblCols[j], h, '', align='C', border=1)
            y += h
            self.set_xy(xstart, y)

    def tableOut(self, data):
        def roundBoards(cols, tCol0, tCol2, tCol3, v, i, extra=''):
            self.cell(cols[0], h, text=tCol0, align='C', border=1)
            self.cell(cols[1], h, text=f"{v[i][0]['NS']}", align='C', border=1)
            self.cell(cols[2], h, text=tCol2, align='C', border=1)
            self.cell(cols[3], h, text=f'{tCol3}{extra}', align='C', border=1)

        for t,v in data.items():
            self.add_page()
            self.movementSheet()
            self.compass()
            self.moveInstruction(t, v['nsNext'], v['ewNext'])

            tblHeight = self.lineHeight(PDF.titlePt) * 2 + self.lineHeight(PDF.headerPt)  + self.lineHeight(PDF.bigPt) * v['nRound']
            bottom = self.eph - self.lineHeight(PDF.bigPt) * 3.5 - PDF.starRadius
            topEdge = self.lineHeight(PDF.bigPt) * 3.5
            top = (bottom - topEdge - tblHeight) / 2 + topEdge
                
            tblIdText = f'Table {t+1}'
            x = self.setHCenter(self.get_string_width(tblIdText))
            self.set_xy(x, top)
            self.set_font(self.serifFont, style='B', size=PDF.titlePt)
            self.cell(text=tblIdText)
            self.set_font(self.sansSerifFont)
            cols = self.tableRoundHeaders()
            allWidth = sum(cols)
            x = self.setHCenter(allWidth)
            y = self.get_y() + self.lineHeight(self.font_size_pt)
            self.set_xy(x, y)
            saveFont = self.font_family
            self.set_font(self.fixedWidthFont, size=PDF.bigPt)
            h = self.lineHeight(self.font_size_pt)

            # Save the last round for special handling
            for i in range(v['nRound'] - 1):
                roundBoards(cols, f'{i+1}', f"{v[i][0]['EW']}", f"{v[i][1]}", v, i)
                y += h
                self.set_xy(x, y)

            # Last round
            i = v['nRound'] - 1
            # case of 5 or 6 pairs, the last round has a special board arrangement
            if v['nRound'] == 5 and i == 4:
                # special case of 6 pairs
                pBoards = [x.strip() for x in v[i][1].split('&')]
                esRotate = [5, 3, 1]
                if data[0][0][0]['NS'] != 0:    # 6 pairs
                    for bIdx in range(len(pBoards)):
                        b = pBoards[(bIdx+t)%len(pBoards)]
                        roundBoards(cols, f'{i+1}', f"{v[i][0]['EW']}", b, v, i, ' (Shared)')
                        y += h
                        self.set_xy(x, y)
                else:
                    for bIdx in range(len(pBoards)):
                        b = pBoards[(bIdx-t+3)%len(pBoards)]
                        esTable = f'{esRotate[(bIdx+t)%len(esRotate)]}'
                        roundBoards(cols, f'{i+1}', esTable, b, v, i, ' (EW Moves)')    
                        y += h
                        self.set_xy(x, y)
            else:
                # normal case
                roundBoards(cols, f'{i+1}', f"{v[i][0]['EW']}", f"{v[i][1]}", v, i)
            
            self.set_font(saveFont)
        return

    def idTags(self, data):
        round1 = []
        for k,v in data.items():
            if v[0][0]['NS'] != 0:
                round1.append((v[0][0]['NS'], k, 'NS'))
            round1.append((v[0][0]['EW'], k, 'EW'))
        round1.sort(key=lambda x: x[0])

        nTagsPage = len(round1) if len(round1) <= 8 else 8
        nPage = len(round1) // 8 + 1
        cHeight = self.eph / nTagsPage
        cWidth = self.w / 2
        for p in range(nPage):
            y = PDF.margin
            x = PDF.margin
            self.add_page(orientation='portrait')
            self.set_line_width(PDF.thinLine)
            self.set_dash_pattern(dash=0.1, gap=0.1)
            for r in range(nTagsPage - 1):
                y += cHeight
                self.line(x1=0, y1=y, x2=self.w, y2=y)
            self.line(x1=cWidth, y1=0, x2=cWidth, y2=self.h)
            self.set_dash_pattern()
            x = self.margin + 0.5
            if len(round1) >= nTagsPage and p <= 0:
                enumRound = round1[:nTagsPage]
            else:
                enumRound = round1[nTagsPage:]
            for l,r in enumerate(enumRound):
                self.set_font(self.serifFont, size=PDF.rosterPt)
                pairId = f'Pair {r[0]}'
                moveInstruction = f'Round 1 go to Table {r[1]+1}, {r[2]}'
                y = cHeight * l + self.lineHeight(self.font_size_pt)
                for xPos in [x, x+cWidth]:
                    self.set_xy(xPos, y)
                    self.cell(w=cWidth, h=self.lineHeight(self.font_size_pt), text=pairId, align='L')
                y += 2 * self.pt2in(self.font_size_pt)
                self.set_font(self.sansSerifFont, size=8)
                for xPos in [x, x+cWidth]:
                    self.set_xy(xPos, y)
                    outtxt = f'{moveInstruction}\nSubsequent rounds follow the movement sheet on the table.'
                    self.multi_cell(w=cWidth, h=self.lineHeight(self.font_size_pt), text=outtxt, align='L')

    # pickup slip: one sheet per table per round
    def pickupSlips(self, pdfData, nPerSet):
        tblCols = []
        xMargin = 0.5
        hdrs = ['Board', 'NS', 'EW', 'Contract', 'By', 'Result', 'NS', 'EW']
        self.setHeaders(xMargin, hdrs, tblCols)
        bIdx = 0
        for tbl in sorted([a for a in pdfData.keys()]):
            for r in range(pdfData[tbl]['nRound']):
                if bIdx % 4 == 0:
                    self.add_page()
                    self.footer()
                    y = 0.5
                y = self.headerRow(xMargin, y, tblCols, hdrs, f"Pickup: Table {tbl+1}, Round {r+1}")
                h = self.lineHeight(self.font_size_pt)
                self.set_font(size=PDF.linePt)
                y += h
                self.set_xy(xMargin, y)
                tblRound = pdfData[tbl][r][0]
                for b in range(nPerSet):
                    self.cell(tblCols[0], h, text=f"{tblRound['Board']*nPerSet+b+1}", align='C', border=1)
                    self.cell(tblCols[1], h, text=f"{tblRound['NS']}", align='C', border=1)
                    self.cell(tblCols[2], h, text=f"{tblRound['EW']}", align='C', border=1)
                    for c in range(3,len(hdrs)):
                        self.cell(tblCols[c], h, text='', align='C', border=1)
                    y += h
                    self.set_xy(xMargin, y)
                bIdx += 1
                y = self.sectionDivider(4, bIdx, xMargin)
        return

    # Record sheet for each pair
    def pairRecords(self, pdfData, nPerSet):
        pairs = {}
        for tbl,rData in pdfData.items():
            for r in range(rData['nRound']):
                    if rData[r][0]['NS'] not in pairs:
                        pairs[rData[r][0]['NS']] = []
                    if rData[r][0]['EW'] not in pairs:
                        pairs[rData[r][0]['EW']] = []
                    pairs[rData[r][0]['NS']].append((r, rData[r][0]['NS'], rData[r][0]['EW'], rData[r][0]['Board']))
                    pairs[rData[r][0]['EW']].append((r, rData[r][0]['NS'], rData[r][0]['EW'], rData[r][0]['Board']))
        tblCols = []
        xMargin = PDF.margin * 2
        hdrs = ['Round', 'Board', 'NS', 'EW', 'Contract', 'By', 'Result', 'NS', 'EW']
        self.setHeaders(xMargin, hdrs, tblCols)
        for p in sorted(pairs.keys()):
            self.add_page()
            self.footer()
            y = self.headerRow(xMargin, self.margin, tblCols, hdrs ,f"Pair {p} Play Journal")
            h = self.lineHeight(self.font_size_pt)
            y += h
            self.set_xy(xMargin, y)
            self.set_font(size=PDF.linePt)
            for r in sorted(pairs[p],key=lambda x: x[0]):
                for b in range(nPerSet):
                    self.cell(tblCols[0], h, text=f'{r[0]+1}', align='C', border=1)
                    self.cell(tblCols[1], h, text=f'{r[3]*nPerSet+b+1}', align='C', border=1)
                    self.cell(tblCols[2], h, text=f'{r[1]}', align='C', border=1)
                    self.cell(tblCols[3], h, text=f'{r[2]}', align='C', border=1)
                    for c in range(4,len(hdrs)):
                        self.cell(tblCols[c], h, text='', align='C', border=1)
                    y += h
                    self.set_xy(xMargin, y)

        return

    def tableRoundHeaders(self):
        line = self.get_y() + self.lineHeight(self.font_size_pt) 
        self.set_font(style='B', size=PDF.headerPt)
        h = self.lineHeight(self.font_size_pt)
        headers = ['Round', 'NS', 'EW', 'Boards']
        tblCols = []
        for t in headers:
            tblCols.append(self.get_string_width(t)+0.2)
        tblCols[3] += 1
        allWidth = sum(tblCols)
        x = self.setHCenter(allWidth)
        self.set_xy(x, line)
        for w in range(len(tblCols)):
            self.cell(tblCols[w], h, text=headers[w], border=1, align='C')
        return tblCols

    # An 1-page overview of the entire tournament
    # Table-by-Round with each cell as players and boards to play
    # Won't fit well in portrait mode
    def overview(self, data):
        # font sizes are determined relatively to the amount of data to fit
        self.add_page(orientation='landscape', format='letter')
        nCol = len(data.items()) + 1
        nRow = data[1]['nRound'] + 1
        colWidth = self.w / nCol
        rowHeight = (self.h - 2 * self.lineHeight(PDF.rosterPt))/ nRow
        rowFontSize = rowHeight*72/4+2  # convert to Pt, try to fit two lines into a row
        if rowFontSize > PDF.rosterPt - 4:
            rowFontSize = PDF.rosterPt - 4
        self.set_font(self.serifFont, size=rowFontSize+4)
        title='Tournament Overview'
        self.set_xy((self.w - self.get_string_width(title)) / 2, self.margin)
        self.cell(text=title)
        self.set_font(self.serifFont, size=rowFontSize+2)
        x = 0
        y = self.get_y() + self.lineHeight(self.font_size_pt)
        self.set_font(self.sansSerifFont, size=rowFontSize)
        for i in range(len(data)):
            x += colWidth
            self.set_xy(x, y)
            self.cell(w=colWidth, h=rowHeight, text=f'Table {i+1}', align='C')

        x = 0
        for i in range(data[0]['nRound']):
            y += rowHeight
            self.set_xy(x, y)
            self.cell(w=colWidth, h=rowHeight/2, text=f'Round {i+1}', align='C')

        x = colWidth
        y = rowHeight + 2 * self.lineHeight(rowFontSize)
        line = self.lineHeight(self.font_size_pt)
        savePt = self.font_size_pt
        for t,v in data.items():    # by table
            for i in range(v['nRound']):
                self.set_xy(x, y)
                self.set_font(self.sansSerifFont, size=savePt, style='')
                self.cell(w=colWidth, h=line, text=f'{v[i][0]["NS"]} v. {v[i][0]["EW"]}', align='C')
                self.set_xy(x, y+line)
                self.set_font(self.fixedWidthFont, style='I', size=savePt - 2)
                self.cell(w=colWidth, h=line, text=f'Boards {v[i][1]}', align='C')
                y += rowHeight
            x += colWidth
            y = rowHeight + 2 * self.lineHeight(rowFontSize)
        

    def compass(self):
        # fancy compass canvas
        self.set_font(self.serifFont, size=PDF.bigPt)
        h = self.lineHeight(self.font_size_pt)
        bottomEdge = h * 4
        starCenter = (self.w/2, self.h - bottomEdge - PDF.starRadius - h/2)
        self.star(starCenter[0], starCenter[1], 0.2, PDF.starRadius, 4, 0, 'D')
        self.line(starCenter[0]-1, starCenter[1], starCenter[0]+1, starCenter[1])
        self.line(starCenter[0], starCenter[1]-1, starCenter[0], starCenter[1]+1)
        self.set_xy(starCenter[0]-self.get_string_width('N')/1.5, starCenter[1]-PDF.starRadius-h)
        self.cell(None,h,text='N')

    def moveInstruction(self, tbl, nextNS, nextEW):
        # movement instructions
        if tbl == nextNS[0]:
            nsText = 'Stationary Pair.  Just stay.'
        else:
            nsText = f'Next Round to Table {nextNS[0]+1} {nextNS[1].upper()}'
        ewText = f'Next Round to Table {nextEW[0]+1} {nextEW[1].upper()}'
        self.inkEdgeText(nsText, ewText)

    def inkEdgeText(self, nsText, ewText):
        self.set_font(style='B', size=PDF.bigPt)
        edge = self.lineHeight(self.font_size_pt) * 2.5
        self.angleText(nsText, 'N', edge)
        self.angleText(nsText, 'S', edge)
        self.angleText(ewText, 'E', edge)
        self.angleText(ewText, 'W', edge)

    def movementSheet(self):
        saveFont = self.font_family
        # Construction lines
        #self.line(self.w/2,0,self.w/2,self.h)
        #self.line(0,self.h / 2, self.w, self.h / 2)
        self.set_font(self.serifFont, style="B", size=24)
        edge = self.lineHeight(self.font_size_pt)
        self.angleText('North', 'N', edge)
        self.angleText('East',  'E', edge)
        self.angleText('West',  'W', edge)
        self.angleText('South', 'S', edge)
        self.set_font(saveFont)

    def angleText(self, txt, facing, edgeMargin):
        strW = self.get_string_width(txt)
        if strW <= 0:
            return

        match facing:
            case 'S':
                self.set_xy((self.w - strW)/2, self.eph - edgeMargin)
                rot = 0
            case 'E':
                self.set_xy(self.w - edgeMargin, (self.h + strW)/2)
                rot = 90
            case 'N':
                self.set_xy(self.w / 2 + strW / 2, edgeMargin + self.pt2in(self.font_size_pt))
                rot = 180
            case 'W':
                self.set_xy(edgeMargin, (self.h - strW)/2)
                rot = 270
        with self.rotation(angle=rot):
            self.cell(text=txt)

    def setHeaders(self, leftMargin, hdrs, cols):
        self.set_font(size=PDF.linePt)
        for i in range(len(hdrs)):
            cols.append(self.get_string_width(hdrs[i]) + 0.2)
        allW = sum(cols)
        if 'Contract' in hdrs:
            cols[hdrs.index('Contract')] += self.epw - allW - leftMargin

    def headerRow(self, leftMargin, y, cols, hdrs, title):
        self.set_xy(leftMargin, y)
        self.set_font(style='B', size=PDF.headerPt)
        h = self.lineHeight(self.font_size_pt)
        self.cell(text=title)
        y += h
        self.set_xy(leftMargin, y)
        self.set_font(style='B', size=PDF.linePt)
        h = self.lineHeight(self.font_size_pt)
        for i in range(len(hdrs)):
            self.cell(cols[i], h, text=hdrs[i], align='C', border=1)
        return y

    def sectionDivider(self, nSection, bIdx, leftMargin):
        h = self.lineHeight(self.font_size_pt)
        secIdx = bIdx % nSection
        if secIdx == 0:
            return self.get_y() + h
        secY = (self.h - 0.5) / nSection
        y = secY * (secIdx % nSection) + 0.5 - h
        self.set_line_width(PDF.thinLine)
        self.set_dash_pattern(dash=0.1, gap=0.1)
        self.line(x1=leftMargin, y1=y, x2=self.w - leftMargin, y2=y)
        self.set_dash_pattern()
        return y + h

    # boards : {board #: [(r, tbl, ns, ew), ...]] 
    def travelers(self, log, nDeck, boards):
        tblCols = []
        xMargin = 0.5
        hdrs = ['Round', 'NS', 'EW', 'Contract', 'By', 'Result', 'NS', 'EW']
        self.setHeaders(xMargin, hdrs, tblCols)
        bIdx = 0
        for i in range(len(boards)):
            l = sorted(boards[i], key=lambda x: x[0])
            for d in range(nDeck):
                if bIdx % 4 == 0:
                    self.add_page()
                    self.footer()
                    y = 0.5
                y = self.headerRow(xMargin, y, tblCols, hdrs, f'Traveler for Board: {nDeck*i+d+1}')
                h = self.lineHeight(self.font_size_pt)
                self.set_font(size=PDF.linePt)
                for x in [r for r in l if r[2] != 0]:
                    y += h
                    self.set_xy(xMargin, y)
                    self.cell(tblCols[0], h, text=f'{x[0]+1}', align='C', border=1)
                    self.cell(tblCols[1], h, text=f'{x[2]}', align='C', border=1)
                    self.cell(tblCols[2], h, text=f'{x[3]}', align='C', border=1)
                    for c in range(3,len(hdrs)):
                        self.cell(tblCols[c], h, text='', align='C', border=1)
                bIdx += 1
                y = self.sectionDivider(4, bIdx, xMargin)

        return

def experimentPDF(pdf):
    pdf.line(0, pdf.h / 2, pdf.w, pdf.h / 2)
    pdf.line(0, pdf.h / 2, pdf.w, pdf.h / 2)
    pdf.line(pdf.w / 2, 0, pdf.w / 2, pdf.h)
    pdf.set_xy(pdf.w / 2, pdf.h / 2)
    pdf.set_dash_pattern(dash=0.1, gap=0.1)
    pdf.line(0, pdf.h / 2+pdf.pt2in(pdf.font_size_pt), pdf.w, pdf.h / 2+pdf.pt2in(pdf.font_size_pt))
    pdf.line(0, pdf.h / 2-pdf.font_size, pdf.w, pdf.h / 2-pdf.font_size)
    pdf.set_dash_pattern()
    t = 'Hello'
    pdf.star(pdf.w/2, pdf.h/2, 0.2, 1, 4, 0, 'D')
    pdf.line(pdf.w/2-1, pdf.h / 2-1, pdf.w/2+1, pdf.h / 2-1)
    for a in [0, 90, 180, 270]:
        pdf.set_xy(pdf.w / 2, pdf.h / 2)
        with pdf.rotation(angle=a):
            pdf.cell(w=None, h=pdf.font_size, text=f'{t} @ {a}')


if __name__ == '__main__':
    pdf = PDF()
    experimentPDF(pdf)
    pdf.output('experment.pdf')