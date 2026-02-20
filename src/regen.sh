#!/bin/bash
rm ../mitchell*.{pdf,xlsx}
rm ../teammatch*.{pdf,xlsx}
for p in 8 9 10
do
./mitchell.py -p $p -b 3
done
./howell.py 
./generic.py
./teammatch.py -m 1 -b 4
./teammatch.py -m 2 -b 3
./teammatch.py -m 3 -b 2
