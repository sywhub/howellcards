No Right reserved.  All files are in public domain.  Author claims no responsbility for using any of the files.

# Howell Movement Cards and Scoring Spreadsheet

Here we have the necessary document to conduct a Bridge Howell tournament for 6 to 14 pairs (3 to 7 tables).  They are designed for the tournament director (TD) and/or organizer.

## What are the files?

There are two files for each numbner of pairs in the tournament.  They are named howell_x.pdf and howell_x.xlsx, where "x" is the number of pairs.  Tournaments for less than 6 pairs probably should just go for team format.  Those for greater than 14 should probably adopt Mitchell style movements.

## Tournament Operations

Pairs are identified from 1 to n.  If n is an odd number (7, 9, 11, or 13 pairs), there's a _phantom_ "pair 0" to make it even number.  Odd-number-pair tournaments have a "sit-out" table where the idling pair sits with the phatom "pair 0".  In reality, the tournament director does not need to setup a physical table to accomodate the sit-out pair.

These tournaments designate a "stationary" pair who does not move with each round.  It is either the highest numbered one or the phatom pair 0.  The stationary pair sits at table zero position North-South.  All other pairs move to where the immediate lower numbered one used to sit.  For example, pair 4 moves to where pair 3 used to sit, pair 3 to pair 2, etc.

We recommend assigning the stationary pair strategically prior to the tournament.  If such consideration is not necessary, then assign it randomly.

The boards generally move to the next less numbered table.  If you study the moving cards, you will discovered a design of one or two "relay" tables to avoid the same players playing the same boards.  Since most tournaments actually simply designate a "relay area" that players deposit the boards just played and fetch the ones to play next.  The relay tables are generally not necessary.

Prior to the tournament, print out the PDF file, and scissor the name tags and traveler sheets along the dotted line.  Tape the movement instruction sheets at the center of each table, facing the same direction.  Its best to arrange the tables either clock- or counter-clock-wise.

When the players arrive, assign pair number to each, give them the name tags which tell them which table to sit initially.  Somehow, shuffle and deal the necessary boards.  Fold and attach the traveler sheets to each board.  (The convention is to tuck it at the North slot.)

Announce to the room the general etiquette: South fetches/deposits the board before and after each round, North record the results on the traveler sheet.  Each pair moves to the next table/position according to the instructions on the table.

Begin the tournament. Hopefully, everything just go smoothly.  Collect all travelers and record them into the spreadsheet.  The result shows up on the "roster" tab immediately.

## Historical Trivials

Edwin Howell is credited for Howell movement.  He invented it for the game of Whist, the prececessor to bridge.  It is probablyu the most popular "single winner" pair-wise tournament style.  The other major pair-wise tournament styles are Mitchell and Swiss.  Mitchell is "multiple winners" and Swiss arranges each round based on the result of the previous one.

No one seems to know how Edwin Howell invented this movement. Since he was a mathematician, it is speculated that he did it based on the Room Squares https://en.wikipedia.org/wiki/Room_square problem.  The obscure mathematical problem was widely researched in the early 1900s and considered solved.  Before computer, people published the solutions for simpler cases.  Computers then have solved it for very big cases.

Howell's solved the Room Square problem quite cleverly.  We got the hint of his methodology by observing the many published tournament arrangements.  I have not discovered publication on the exact algoirthm.

## Technical Requirements

There are Python programs used to generate these files.  At this moment, I have not uploaded them to GitHub yet.  What follow are general design points.  I have not written down the algorithm in prose form.  It exists in the form of these python code.

### Requirement: Seating and Movement

**Except for the stationary pair, all pair move to the seat occupied by the pair numbered just below them.  (Pair #1 moves to where pair N used to sit.  N being the highest numbered pair.)  All pairs meet all other pairs exactly once.**

Naturally, there must be n/2 tables for n pairs to have a pair-wise tournament. For each table, one pair sits North-South (NS) and the other East-West (EW).  These directions do not need to comform their geological designations.  We follow the common convention of assigning a "stationary pair" at table 1, NS. Algorithmically, it is a random choice where to sit this pair, or even to have one at all.  In actual tournaments, having a stationary pair makes it easy to accomodate moving-challenged people or a playing TD.

To meet this requirement, n pairs must play n-1 rounds to play aginst all other pairs.  The movement requirement dictates the initial seatings.  After players have seated, they move to the next position after each round as this requirement.  I cannot find an algorithm to arrange the initial seatings.  So I just iterate all possibilities and filter out those not meeting this requirement.  It is rather simple to use Python _itertool_ and just sit back to let the computer work.  There are several such assignments for each tournament size — numerous when size is bigger.  The program is capable for generating all assignments.  It picks one at random.

The tournament floor is less chaotic if the movements follow the same direction, either always toward the lower number table or higher. To do so, pairs must initially sit at either ascending or descending order relative to the table.  Some researches indicate more "fair" play if pairs sit at NS and EW alternatively.  For this, even and odd number pairs must all sit at the same side (NS or EW) initially.  (If all even numbered pairs sit at NS and they move to the where the lower-number pair used to sit, they then must all sit at EW next round.)

I cannot find any solution if the these are also requirements.

### Requirement #2: Board Assignments

**No pair will play a board more than once. All pairs play the same number of boards.**

There is no mathematically proper solution for tournaments of 6 pairs or odd number of pairs.  Fortunately, there are easy hacks.

6-pair tournaments "share" the same boards for one of the rounds, typically the last.  The simplest way is to have 3 boards per round and have 3 "mini rounds" at the end.  During that last round, each table plays just one board at a time and just move the board to the next table when done.

For odd-numbered pair tournaments, we make up a "phantom pair" to sit as the stationary pair.  That table becomes the "sit-out" table.  WHichever pair rotates to that table just idle for a round.

With these, we proceed to solve for n = 8, 10, 12, and 14.  For a tournament of size n, there must be n-1 rounds and therefore same number of board sets.  Each table plays a set of boards for each round.  A "set of boards" are several boards in consecutive numbers.  For example, if a set has 2 boards, then the 1st set has board #1 and #2, the 2nd set has #3 and #4, etc.

For n = 6, we allocate 3 boards for a set; otherwise just 2 boards.

First we assign board set to the stationary table (#1) in ascending sequence for each round.  It gets first board set for the first round, 2nd for second, etc.  We then attempt to allocate sets serially to the consecutive tables. (Table 2 gets the 2nd board set, table 3 the 3rd, etc.)  This breaks down quickly in the ensuing rounds when, inevitably, pairs will play the same boards.  We solve this by skipping one or more sets of board during the initial assignment or change the board sequence (the next table does not get the next set of boards).  We iterate all possible arrangements until a solution emerges. For each round, we will skip in the same manner.  If table 3 got set #4, instead of #3, for the first round, while table one got set #1 and table two got #2.  Then for the second one, we also skip one for table 3: table one get #2, two #3, and three #5, etc.

The consistent skipping make it easier for putting a "relay table" between playing table.

### Requirement #3 (soft): Board Movement

There are several ways to physically get the boards to the right tables for each round.  Caddies can run around tables and move them.  Players can deposit the boards that they just finished to an area and fetch the ones they are about to play.  Lastly, players can move the boards toward a fixed direction.  For the last one, relay tables must be employed to avoid pairs playing the same board.

I recommend the 2nd way: everyone put boards back to an area after each round and get the new ones from the same place.  The drawback for this manner is the chaotic room traffic.  After each round, players are moving to the new places and also going to the same area to get their new boards.  This works well for smaller tournament sizes or smaller rooms.
