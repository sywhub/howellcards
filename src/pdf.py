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
                self.set_xy(x+tab, line)
                self.set_font(style='I', size=PDF.linePt)
                self.cell(2*h, h, text=f'{m[1]}', align='R')
            line += h
        return

    # part of the meta pagee 
    # Some text for the TD/Organizer
    def instructions(self, log):
        txt = '''There is a matching spreadsheet for this PDF.
               Print this before the event.  Cut travelers and ID tags along the dotted line.
               Assign pair number to each pair.  Hand them the ID tag. Encourage them to write their own names on the ID tag. Note the ID tag tell them where to sit.
               Bring writable masking tapes and pens to the  event.
               Travelers are not compatible to other Howell events. The spreadsheet version is generic.
               Arrange to shuffle and deal all boards.
               Fold and tuck the traveler for each board before the tournament begins.
               Tape the "movement sheets" on the table facing the same direction.
               Generally, people mvoe "down" the table and boards move "up."  Having relay tables strategically place helps board movements.
               Announce the direction of north to all participants.
               Generally, North keeps score.  South caddies the boards.
               Collect travelers when the tournment ends.  Record and results on the spreadsheet.  The tournament result should be at the Roster tag.'''
        # wherever we are
        y = self.get_y()+1
        self.set_font(size=PDF.linePt)
        h = self.lineHeight(self.font_size_pt)
        nLine = 1
        for t in txt.split('\n'):
            self.set_xy(1, y)
            self.cell(h, h, f'{nLine}.', align='R')
            self.set_xy(1+h, y)
            self.multi_cell(self.epw-2, h=h, text=t.strip())
            y = self.get_y()
            nLine += 1

    # Sign-up sheet
    def roster(self, log, rows, headers):
        self.add_page() # new page
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
        self.overview(data)
        self.firstRound(data)
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
            for i in range(v['nRound']):
                self.cell(cols[0], h, text=f'{i+1}', align='C', border=1)
                self.cell(cols[1], h, text=f"{v[i][0]['NS']}", align='C', border=1)
                self.cell(cols[2], h, text=f"{v[i][0]['EW']}", align='C', border=1)
                self.cell(cols[3], h, text=f"{v[i][1]}", align='C', border=1)
                y += h
                self.set_xy(x, y)
            
            self.set_font(saveFont)
        return

    def firstRound(self, data):
        round1 = []
        writeInSpaces = 20
        for k,v in data.items():
            if v[0][0]['NS'] != 0:
                round1.append((v[0][0]['NS'], k, 'NS'))
            round1.append((v[0][0]['EW'], k, 'EW'))
        round1.sort(key=lambda x: x[0])

        self.add_page(orientation='portrait')
        cHeight = self.eph / len(round1)
        cWidth = self.w / 2
        y = PDF.margin
        x = PDF.margin
        self.set_line_width(PDF.thinLine)
        self.set_dash_pattern(dash=0.1, gap=0.1)
        for r in range(len(round1)-1):
            y += cHeight
            self.line(x1=0, y1=y, x2=self.w, y2=y)
        self.line(x1=cWidth, y1=0, x2=cWidth, y2=self.h)
        self.set_dash_pattern()
        x = self.margin + 0.5
        topMargin = (cHeight - self.lineHeight(PDF.headerPt) - 2 * self.lineHeight(8)) / 2
        for l,r in enumerate(round1):
            self.set_font(self.serifFont, size=PDF.headerPt)
            pairId = f'Pair {r[0]}: ' + ' ' * writeInSpaces
            moveInstruction = f'Round 1 go to Table {r[1]+1}, {r[2]}'
            y = cHeight * l + topMargin
            y += self.lineHeight(self.font_size_pt)
            for xPos in [x, x+cWidth]:
                self.set_xy(xPos, y)
                self.cell(w=cWidth, h=self.lineHeight(self.font_size_pt), text=pairId, align='L')
            y += 2 * self.pt2in(self.font_size_pt)
            self.set_font(self.sansSerifFont, size=8)
            for xPos in [x, x+cWidth]:
                self.set_xy(xPos, y)
                self.cell(w=cWidth, h=self.lineHeight(self.font_size_pt), text=moveInstruction, align='L')
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
        self.set_font(style='B', size=PDF.bigPt)
        edge = self.lineHeight(self.font_size_pt) * 2.5
        if tbl == nextNS[0]:
            nextStr = 'Stationary Pair.  Just stay.'
        else:
            nextStr = f'Next Round to Table {nextNS[0]+1} {nextNS[1].upper()}'
        self.angleText(nextStr, 'N', edge)
        self.angleText(nextStr, 'S', edge)

        nextStr = f'Next Round to Table {nextEW[0]+1} {nextEW[1].upper()}'
        self.angleText(nextStr, 'E', edge)
        self.angleText(nextStr, 'W', edge)

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

    # boards : {board #: [(r, tbl, ns, ew), ...]] 
    def travlers(self, log, nDeck, boards, txt):
        self.set_font(size=PDF.linePt)
        headers = ['Round', 'NS', 'EW', 'Contract', 'By', 'Result']
        xMargin = 0.5
        tblCols = []
        for i in range(len(headers)):
            tblCols.append(self.get_string_width(headers[i]) + 0.2)
        allW = sum(tblCols)
        tblCols[3] += self.epw - allW - xMargin
        bIdx = 0
        secY = self.eph / 4
        for i in range(len(boards)):
            l = sorted(boards[i], key=lambda x: x[0])
            for d in range(nDeck):
                if bIdx % 4 == 0:
                    self.add_page()
                    y = 0.5
                self.set_xy(xMargin, y)
                self.set_font(style='B', size=PDF.headerPt)
                h = self.lineHeight(self.font_size_pt)
                self.cell(text=f'{len(boards)}-Round & {len(boards[0])}-Table Traveler: Board {bIdx+1}')
                x = self.get_x() + 0.5
                self.set_font(style='', size=PDF.notePt)
                self.cell(w=self.w - x, text=txt, align='R')
                y += h
                self.set_xy(xMargin, y)
                self.set_font(style='B', size=PDF.linePt)
                h = self.lineHeight(self.font_size_pt)
                for i in range(len(headers)):
                    self.cell(tblCols[i], h, text=headers[i], align='C', border=1)
                self.set_font(size=PDF.linePt)
                for x in [r for r in l if r[2] != 0]:
                    y += h
                    self.set_xy(xMargin, y)
                    self.cell(tblCols[0], h, text=f'{x[0]+1}', align='C', border=1)
                    self.cell(tblCols[1], h, text=f'{x[2]}', align='C', border=1)
                    self.cell(tblCols[2], h, text=f'{x[3]}', align='C', border=1)
                    self.cell(tblCols[3], h, text='', align='C', border=1)
                    self.cell(tblCols[4], h, text='', align='C', border=1)
                    self.cell(tblCols[5], h, text='', align='C', border=1)
                bIdx += 1
                y = secY * (bIdx % 4) + 0.5
                self.set_line_width(PDF.thinLine)
                self.set_dash_pattern(dash=0.1, gap=0.1)
                self.line(x1=xMargin, y1=y-0.5, x2=self.w - xMargin, y2=y-0.5)
                self.set_dash_pattern()

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