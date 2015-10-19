# isis_to_marc

Generate MaRC records from Excel spreadsheets provided by Turkish vendor Isis ( http://www.theisispress.org/ ). These come every month, pretty much.

If not there, create the ./in dir and put the Excel spreadsheet (.xlsx) inside, then run e.g.:

`python isis.py -f 2014-7_inv_no_210_Prin.xlsx`

To generate multiple files after given line numbers use the `-s` flag. E.g. To split after lines 96 and 189:

`python isis.py -f 2014-7_inv_no_210_Prin.xlsx -s 96,189`

The mrc files then need to be copied to the load folder. 

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

#### Requires
[xlrd](http://www.python-excel.org/) `pip install xlrd`
