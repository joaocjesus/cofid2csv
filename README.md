# cofid2csv

Converts CoFID excel spreadsheet onto MongoDB friendly importable format.

Composition of foods integrated dataset (CoFID) spreadsheet from:
https://www.gov.uk/government/publications/composition-of-foods-integrated-dataset-cofid

fetchHeaders.py
---------------
Simple script to parse custom column naming spreadsheet into a csv file

cofid2csv.py
------------
Main script to convert CoFID spreadsheet into a csv file.

Requires columnNames.csv - contains naming format to be used by Nutrimprove project

Requires CoFID spreadsheet

Usage:
Run cofid2csv.py
(no parameters)