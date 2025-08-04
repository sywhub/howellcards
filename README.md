# Howell Movement Cards and Scoring Spreadsheet

Here we have the necessary document to conduct a Bridge Howell tournament for 6 to 14 pairs (3 to 7 tables).  They are designed for the tournament director and/or organizer.

## What are the files?

There are two files for each numbner of pairs in the tournament.  They are named howell_x.pdf and howell_x.xlsx, where "x" is the number of pairs.  Tournaments for less than 6 pairs probably should just go for team format.  Those for greater than 14 should probably adopt Mitchell style movements.

## Tournament Operations

Pairs are identified from 1 to n.  If n is an odd number (7, 9, 11, or 13 pairs), there's a ~phantom~ "pair 0" to make it even number.  Odd-number-pair tournaments have a "sit-out" table where the idling pair sits with the phatom "pair 0".  In reality, the tournament director does not need to setup a physical table to accomodate the sit-out pair.

These tournaments designate a "stationary" pair who does not move with each round.  It is either the highest numbered one or the phatom pair 0.  The stationary pair sits at table zero position North-South.  All other pairs move to where the immediate lower numbered one used to sit.  For example, pair 4 moves to where pair 3 used to sit, pair 3 to pair 2, etc.

We recommend assigning the stationary pair strategically prior to the tournament.  If such consideration is not necessary, then assign it randomly.

The boards generally move to the next less numbered table.  If you study the moving cards, you will discovered a design of one or two "relay" tables to avoid the same players playing the same boards.  Since most tournaments actually simply designate a "relay area" that players deposit the boards just played and fetch the ones to play next.  The relay tables are generally not necessary.

Prior to the tournament, print out the PDF file, and scissor the name tags and traveler sheets along the dotted line.  Tape the movement instruction sheets at the center of each table, facing the same direction.  Its best to arrange the tables either clock- or counter-clock-wise.

When the players arrive, assign pair number to each, give them the name tags which tell them which table to sit initially.  Somehow, shuffle and deal the necessary boards.  Fold and attach the traveler sheets to each board.  (The convention is to tuck it at the North slot.)

Announce to the room the general etiquette: South fetches/deposits the board before and after each round, North record the results on the traveler sheet.  Each pair moves to the next table/position according to the instructions on the table.

Begin the tournament. Hopefully, everything just go smoothly.  Collect all travelers and record them into the spreadsheet.  The result shows up on the "roster" tab immediately.

## Technical Details

There are a setup of Python files used to generate these files.  At this moment, I have not uploaded them to GitHub yet.

### Requirement #1

Except for the stationary pair, all pair move to the seat occupied by the pair numbered just below them.  Pair #1 moves to where pair N used to sit.  (N being the highest numbered pair.)

No pair meet all other pairs exactly once.

These two requirement dictates the initial seat assignment.  There are several such assignments for each total pair number. The program is capable for generating all assignments.  It picks one at random.

### Requirement #2

No pair will play a board more than once. All pairs play the same number of boards.

This requirement is a special case for the mathematical "Room Squares" problem.  (Named after T.G. Room, a mathematician.)  There is no solution for odd number of pairs or 6 pairs.

We "hack" 6-pair tournament by having all pair "sharing" the same board for the last round.  This is why there are 3 boards per round.  The last round is really 3 sub-rounds each playing one board, then just toss the board to the next table.

For odd-numbered pair tournaments, we invent a "phantom pair" to sit one each real pair for each round.  This creates a "sit-out" table.

To generate the boards assignment for each round, we first give table one the boards in ascending sequence. Then we try to give the next table the next set of boards, and so on.  If we found "duplicates" that a pair will play the same board more than once, then we create a relay table to buffer the board movement.  We try all combinations until one arragnement meeting the requirement.  This is a solution for the special case Room Square problem, with the extra resources of one to three "relay tables."

As stated, TD does not need to have physical relay tables.  A "relay area" that all boards are deposited/fetched works well.  It takes a bit searching in the area to find your boards, most players do not find it difficult or inconvenient.

Alternative, TD may have human "caddies" to transport boards to each table for each round.

### Requirement #3 (soft)

It is desirable for boards to move in a consistent direction, usually toward the less numbered table.  The movement sheets will arrange that with setting up of one to three relay tables.  If used, South will put the just-played boards on the relay table closer to the less numbered table and South will fetch/receive the new boards from the higher numbered table.