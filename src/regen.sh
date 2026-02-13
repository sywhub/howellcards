#!/bin/bash
rm ../mitchell*.{pdf,xlsx}
for p in 8 9 10
do
./mitchell.py -p $p -b 3
done
./howell.py 
