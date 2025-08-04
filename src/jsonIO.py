#!/usr/bin/env python3
# Load setup data, verify they are valid
# File "setup.json" was separately generated. We load from it and validate the data
# are indeed good for tournament.
import logging
import json5    # JSON5 supposedly can handle comments
from maininit import setlog
from tables import HowellSeats

class JsonIO:
    def __init__(self, pairs, log=None):
        self.fname = 'setup.json'
        self.log = log
        if self.log == None:
            self.log = setlog('jsonIO', None)
        self.pairs = pairs
        self.tournament = None

    def meta(self, rounds, tables):
        self.tournament = {'Rounds': rounds, 'Tables': tables, 'BoardMovement': None}

    def resetTournament(self):
        del self.tournament['Arrangement']

    def addRound(self, tables):
        k = 'Arrangement'
        if k not in self.tournament:
            self.tournament[k] = []
        self.tournament[k].append(tables)

    def boardToSet(self, bIdx):
        return chr(ord('A')+bIdx)

    def getFileName(self):
        return self.fname

    # data conversion from previous algorithm-based result
    # This is used really just once (or occassionally) to regenerate data
    def saveToJSON(self, rounds):
        self.tournament = {'Rounds': len(rounds), 'Tables': len(rounds[0]), 'BoardMovement': None}
        self.tournament['Arrangement'] = [None]*len(rounds)
        for r,v in rounds.items():
            self.tournament['Arrangement'][r] = [None] * len(v)
            for t,tbl in v.items():
                self.tournament['Arrangement'][r][t] = {'NS': tbl['ns'], 'EW': tbl['ew'], 'Board': tbl['board']}
        return self.tournament

    def boardMovement(self, seq):
        seq.insert(0, 0)
        self.tournament['BoardMovement'] = seq

    def dump2File(self, f):
        objStr = json5.dumps(self.tournament)
        for k in self.tournament.keys():
            objStr = objStr.replace('Arrangement: ', 'Arrangement:\n\t\t')
            objStr = objStr.replace('], ', '],\n\t\t')
            objStr = objStr.replace(']]}, ',']]},\n')
        objStr = f'"{self.pairs}": \t // {self.pairs} pairs\n\t' + objStr
        print(objStr,end=',\n',file=f)
        return

    # Load tournament data as JSON
    def load(self, fname=None):
        if fname != None:
            self.fname = fname
        self.log.info('Loading data file')
        try:
            with open(self.getFileName(), 'r') as f:
                loadObj = json5.load(f)
        except:
            self.log.error('JSON load failed')
            return None
    
        # Default keys are always string.  We don't like that.
        if str(self.pairs) not in loadObj:
            self.log.error(f'{self.pairs} not in data')
            return None

        self.tournament = loadObj[str(self.pairs)]
        isValid = self.validateData()
        self.showArrangement()
        if not isValid:
            self.tournament = None
        return self.tournament

    def validateData(self):
        self.log.info(f'Validating {self.pairs}-pair tournament data')
        ret =  self.validateMovement() and \
            self.validatePairs() and \
            self.validateBoards()
        self.log.info(f'{self.pairs}-pair data {"validated" if ret else "invalid"}')
        return ret

    # The pairs for every table must always move the same way
    def validateMovement(self):
        moves = [None] * self.tournament['Tables']
        # each element a separate copy
        for m in range(self.tournament['Tables']):
            moves[m] = {'NS': None, 'EW': None}

        for r in range(self.tournament['Rounds'] - 1): # last round has no movement
            for t in range(self.tournament['Tables']):
                tbl = self.tournament['Arrangement'][r][t]
                for s in ['NS', 'EW']:
                    for sNext in ['NS', 'EW']:
                        nextSide = [x[sNext] for x in self.tournament['Arrangement'][r+1]]
                        if tbl[s] in nextSide:
                            nextSeat = (s, nextSide.index(tbl[s]))
                            if moves[t][s] == None:
                                moves[t][s] = nextSeat
                            elif moves[t][s] != nextSeat:
                                msg = f'Round {r+1} table {t+1} {sNext.upper()} moves inconsistently'
                                self.log.error(msg)
                                print(msg)
                                return False
        self.log.info('All table movements consistent')
        return True

    def validatePairs(self):
        tourney = []
        for rIdx, r in enumerate(self.tournament['Arrangement']):
            tourney.append([])
            for t in r:
                tourney[rIdx].append(t['NS'])
                tourney[rIdx].append(t['EW'])
        hSeat = HowellSeats(self.pairs, self.log)
        good = hSeat.validateTournament(tourney)
        return good

    # Number of boards is one less than number of pairs
    # Each board played once for each pair
    # All boards played number of times as number of tables
    def validateBoards(self):
        boardSets = {}
        addOdd = self.pairs + self.pairs % 2
        playCount = [0] * (addOdd - 1)
        for r in self.tournament['Arrangement']:
            for t in r:
                try:
                    playCount[t['Board']] += 1
                except IndexError:
                    msg = f'Board {t['Board']} IndexError in play count'
                    self.log.error(msg)
                    print(msg)
                    return False
                if t['Board'] not in boardSets:
                    boardSets[t['Board']] = set()
                for side in ['NS', 'EW']:
                    if t[side] in boardSets[t['Board']]:
                        msg = f'{t[side]} already played board {self.boardToSet(t['Board'])}'
                        self.log.error(msg)
                        print(msg)
                        return False
                    boardSets[t['Board']].add(t[side])

        # Number of boards is one less than number of pairs
        if len(boardSets) != addOdd - 1:
            msg = f'# of Board {len(boardSets)} not {addOdd - 1}'
            self.log.error(msg)
            print(msg)
            return False

        # Each board played exactly the number of times as number of tables
        for p in playCount:
            if p != self.tournament['Tables']:
                msg = f'{playCount} not matched number of tables {self.tournament['Tables']}'
                self.log.error(msg)
                print(msg)
                return False

        # Each board play by all pairs
        for b in boardSets.values():
            if len(b) != addOdd:
                msg = f'Board not played # of times'
                self.log.error(msg)
                print(msg)
                return False
        self.log.info(f'All boards played {playCount[0]} times by all pairs')
        return True

    def showArrangement(self):
        if self.tournament == None:
            return

        print(f'{self.pairs} Pairs, {self.tournament['Rounds']} Rounds, {self.tournament['Tables']} Tables')
        nBoards = len(self.tournament['Arrangement'])
        print(f'{"Brd":>4}', end='')
        for r in range(nBoards):
            print(f'{self.boardToSet(r):^6}',end='')
        print('')
        rIdx = 1
        for r in self.tournament['Arrangement']:
            tbls = sorted(r, key=lambda x: x['Board'])
            print(f'{rIdx:2}: ', end='')
            rIdx += 1
            tIdx = 0
            for i in range(nBoards):
                if tIdx < len(tbls) and tbls[tIdx]['Board'] == i:
                    print(f'{tbls[tIdx]['NS']:2},{tbls[tIdx]['EW']:<2}', end=' ')
                    tIdx += 1
                else:
                    print(f'{" ":5}', end=' ')
            print('')

    def sortByBoard(self):
        p = self.tournament['Arrangement']
        n = len(p)
        m = len(p[0])
        for i in range(m-1):
            for j in range(i+1,m):
                if p[0][i]['Board'] > p[0][j]['Board']:
                    for k in range(n):
                        p[k][i], p[k][j] = p[k][j], p[k][i]
        return


if __name__ == '__main__':
    log = setlog('jsonIO', None)
    log.setLevel(logging.DEBUG)
    for p in range(6,15):
        jIO = JsonIO(p, log)
        print(f'--- {p} pairs ---')
        jObj = jIO.load()
