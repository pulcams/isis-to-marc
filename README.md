Isis Press Excel to MARC 
------------------------

Generate MaRC records from Excel spreadsheets provided by Turkish vendor Isis ( http://www.theisispress.org/ ). These come every month, pretty much.

First time, create the ./in dir and put the Excel spreadsheet (.xlsx) inside, then run e.g.:

`python isis.py -f 2014-7_inv_no_210_Prin.xlsx`

To generate multiple files after given line numbers use the `-s` flag. E.g. To split after lines 96 and 189:

`python isis.py -f 2014-7_inv_no_210_Prin.xlsx -s 96,189`

```
isis
├── archive
│       ├── .mrc
│       └── .mrk
├── in
│   └── .xlsx
├── isis.py
└── temp
```
