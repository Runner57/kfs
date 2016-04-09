#!/bin/sh

cd income_sheet
python3 income_sheet.py 2>&1 | tee run.log
cd -

cd revenue
python3 revenue.py 2>&1 | tee run.log
cd -

cd dividend
perl kfs.pl 2>&1 |tee run.log
cd -
