#!/bin/bash
rm ../mitchell*.{pdf,xlsx}
rm ../teammatch*.{pdf,xlsx}
./mitchell.py -p 8 -b 4
./mitchell.py -p 8 -s -b 3
for p in 9 10
do
./mitchell.py -p $p -b 3
done
./howell.py 
./generic.py
./teammatch.py -m 1 -b 4
./teammatch.py -m 2 -b 3
./teammatch.py -m 3 -b 2
