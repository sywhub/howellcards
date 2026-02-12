#!/usr/bin/env python3
from fpdf import FPDF
# Generate PDF for a matching spreadsheet
# PDFs are for machines, its generation is tedious...

class PDF(FPDF):
    margin = 0.25
    thinLine = 1/72/2   # half a point
    starRadius = 1
    anchorFontSize = 72
    edgePt = 24
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

    def meta(self, meta):
        self.set_font(self.serifFont, style='B', size=self.rosterPt) 
        h = self.lineHeight(self.font_size_pt)
        x = self.setHCenter(self.get_string_width(meta['Title']))
        y = self.margin + h
        self.set_xy(x, y)
        self.cell(text=meta['Title'])
        self.set_font(self.sansSerifFont, style='I', size=self.bigPt) 
        h = self.lineHeight(self.font_size_pt)
        self.set_xy(x, y)
        for t in meta['Info']:
            y += h
            self.set_xy(x, y)
            self.cell(text=f"{t[0]}: {t[1]}")
        return self.get_y()

    # Board tab references its data from the Round tab, for consistency

    # part of the meta pagee 
    # Some text for the TD/Organizer
    def instructions(self, log, fname):
        with open(fname, "r") as f:
            txt = f.read().splitlines()
        self.headerFooter()
        self.set_font(self.serifFont, style='B', size=self.rosterPt) 
        h = self.lineHeight(self.font_size_pt)
        x = self.setHCenter(self.get_string_width(txt[0]))
        y = 2 * h + self.margin
        self.set_xy(x, y)
        self.cell(text=txt[0])
        self.set_font(size=PDF.linePt)
        h = self.lineHeight(self.font_size_pt)
        y = self.get_y()+2*h
        nLine = -1
        mLines = []
        for t in txt[1:]:
            if t[0] == '#':
                mLines.append(t[1:])
                nLine += 1
            else:
                mLines[nLine] += ' ' + t
        lineNo = 1
        for t in mLines:
            self.set_xy(1, y)
            self.cell(h, h=h, text=f"{lineNo}.", align='R')
            self.set_xy(1+h, y)
            self.multi_cell(self.epw-2, h=h, text=t.strip())
            lineNo += 1
            y = self.get_y()

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
        self.set_font(self.serifFont, style="B", size=self.edgePt)
        edge = self.lineHeight(self.font_size_pt)
        self.angleText('North', 'N', edge)
        self.angleText('East',  'E', edge)
        self.angleText('West',  'W', edge)
        self.angleText('South', 'S', edge)
        self.set_font(saveFont)

    def tableAnchors(self, t):
        saveMargin = self.margin
        self.set_margin(0)
        self.set_font_size(self.anchorFontSize)
        tWidth = self.get_string_width(t)
        tHeight = self.pt2in(self.anchorFontSize)
        for corner in range(4):
            match corner:
                case 0:
                    x = tWidth
                    y = tHeight
                    rot = 180
                case 1:
                    x = self.w - tWidth
                    y = tHeight
                    rot = 180
                case 2:
                    x = 0
                    y = self.eph - tHeight
                    rot = 0
                case 3:
                    x = self.w - tWidth
                    y = self.eph - tHeight
                    rot = 0
            # self.circle(x, y, radius, 'D')
            self.set_xy(x, y)
            with self.rotation(angle=rot):
                self.cell(text=t)
            self.set_margin(saveMargin)

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

    # compute the width of each column
    # use whatever font active
    def setHeaders(self, leftMargin, hdrs, cols):
        for h in hdrs:
            cols.append(self.get_string_width(h) + 0.2)
        allW = sum(cols)
        if 'Contract' in hdrs:
            cols[hdrs.index('Contract')] += self.epw - allW - leftMargin

    def headerRow(self, leftMargin, y, cols, hdrs, title):
        self.set_xy(leftMargin, y)
        h = self.lineHeight(self.font_size_pt)
        self.cell(text=title)
        y += h
        self.set_xy(leftMargin, y)
        self.set_font(self.sansSerifFont, style='B')
        h = self.lineHeight(self.font_size_pt)
        for i in range(len(hdrs)):
            if hdrs[i] == 'Made':
                self.set_font(size=self.font_size_pt / 2)
                self.cell(cols[i]+cols[i]+1, h/2, text='Result', align='C', border=1)
            self.cell(cols[i], h, text=hdrs[i], align='C', border=1)
            if hdrs[i] == 'Down':
                self.set_font(size=self.font_size_pt)
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

def experimentPDF(pdf):
    pdf.set_margin(0)
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
    anchor = '8'
    aW = pdf.get_string_width(anchor) * 1.5
    aH = pdf.pt2in(pdf.font_size_pt)
    pdf.set_xy(aW, aH)
    with pdf.rotation(angle=180):
        pdf.cell(text='6')
    pdf.set_xy(pdf.w-aW, aH)
    with pdf.rotation(angle=180):
        pdf.cell(text='7')
    pdf.set_xy(0, pdf.h-aH)
    pdf.cell(text='8')
    pdf.set_xy(pdf.w-aW, pdf.h-aH)
    pdf.cell(text='9')



if __name__ == '__main__':
    pdf = PDF()
    experimentPDF(pdf)
    pdf.output('experment.pdf')