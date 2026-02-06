No Right reserved.  All files are in public domain.  Author claims no responsbility for using any of the files.
# TOC
This directory contains necessary PDF and Excel files for various styles of Bridge
tournaments.  They are designed for "small" tournaments organized by amateur directors.  For
large events, people typically use ACBL tools.
The files come in pairs with matching names differ only by their extention of "PDF" or "xlsx".  The diretor should print out the PDF file ahead of the time.

# Team Match
This is a simple 2-table team match setup.  The expectation is 2 pairs per team to play 2 rounds of 8 boards.  The tables swap boards after playing for 4 boards.  Then one team swap pairing after 8 boards.  Each table keeps a scoring sheet on the table and tally all up after the match.  The spreadsheet uses converts contract scores to IMPs.

You can change the format for # of rounds, boards per round, and switches per round.

# Mitchell Movement
Mitchell Movement is a very popular format for relatively large participants.  In this
style, the players are separated into the "North-South" group and "East-West".  At the end,
rankings are among the groups and there will be 2 winning pairs.
Mitchell movements are straight-forward: NS pairs stay stationary, EW pairs move "up" to the table one less than the current one after each round, boards move "down" to the higher numbered table.
## Even number of tables
If the number of tables is even (4, 6, 8, 10, etc.), pairs see the same board after a few rounds.  We solved this problem by placing a relay table at the middle and an extra set of boards.  For example, 6-table Mitchell has a relay table between table 3 and 4.  While players move the same way, a set of boards will rest on the relay table and not played.  The consequence of this movement is players do not play the same boards and the comparison is less fair.
## 4-table Square Mitchell
For 4-table, there's a "Square" movement that does not require the relay table, or an extra set of boards.  The trade-off is a slightly complicated movement for both players and boards.
We recommend simply do Howell movement for 4-table (8 pairs).  The movements are not that much more complex and all pairs play with all other pairs.

# Howell Movement 
Unlike Mitchell, Howell movements have all pairs play against all other pairs and generating
one winning pair. Both the player and board movements are more complex.  For Howell, the
"movement cards" are critically important for the game.

# Which files do you need?
For Howell, you need a PDF and a Spreadsheet files of matching names.  The spreadsheet makes
scoring easier.

For Mitchell, the same.  Except that not all possible Mitchell files were pre-generated.  We
only generate the necessary files for 8, 9, or 10 pairs and 3 boards per round.  For any
other combinations of number of pairs and number of rounds, you need to run the Python
program and provide it with the the command-line options.

## The Spreadsheet
Record scores in either the "By Board" or "By Round" tab, but not both.  If "By Round" is
chosen, the "By Board" tab fetches the data and compute the scores automatically.  If "By
Board" was used, the "By Round" tab is not used.

The "Roster" tab is the summary of all boards played and provide the final results in both
MP and IMP.

# Tournament Operation
Prior to the tournament, decide which way to keep score.  Most tournaments choose to use
"traveler" sheet that goes with each board.  Some use the "pickup slip" tha the director
collects at the end of each round.  There's no need to do both.

With the above choice, leave out the unnecesary pages from the PDF file.

The PDF file has several sections:
## The sign-up sheet
One single page for players to sign-up.  I recommend pre-assign the pair numbers to
participants prior to the tournament.

## ID Tags
These are for each individual players to carry with themselves as guides in the tournament.
It is provided only for the convenience for them, not used by the director.

## Movement Cards
These are to be taped on each table facing the same direction.  It provides necessary
information for the tournament.  It is the most imporant part of the document set.

## The Travelers
One slip for each board.  The round number, opposing pairs, and board numbers are
pre-printed for convinience.  Many tournaments use standard ACBL travelers.  There is no
real difference.

## The Pickup Slips
One slip per table and per eround.  The directors should collect them at the end of each
round.  There is no need to use both traveler and pickup slips.

## Player Journals
There's one for each pair to record the play history of their plays.  This is the redundant
information to corroborate with either the travelers or the pickup slips.  Whenever there
are errors, these data can help to discern the correct outcome of the play.  The director
should collect these at the end of the tournament.

The use of player journals is optional, particularly among experience players.

# History

Edwin Howell is credited for Howell movement.  In late 1800s, he invented it for the game of Whist, the prececessor to bridge.  It is probably the most popular "single winner" pair-wise tournament style.  The other major pair-wise tournament styles are Mitchell and Swiss.  Mitchell is "multiple winners" and Swiss arranges each round based on the result of the previous one.

No one seems to know how Edwin Howell invented this movement. Since he was a mathematician, it is speculated that he did it based on the Room Squares https://en.wikipedia.org/wiki/Room_square problem.  The obscure mathematical problem was widely researched in the early 1900s and considered solved.  Before computer, people published the solutions for simpler cases.  Computers then have solved it for very big cases.

Howell's solved the Room Square problem quite cleverly.  We got the hint of his methodology by observing the many published tournament arrangements.  I have not discovered publication on the exact algoirthm.

## Technical Requirements

The "src" directory has Python programs used to generate these files.  TD should just ignore them.  The ensuing text describes the general design concepts.  Non-techie please stop reading. 

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
