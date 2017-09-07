%let pgm=utl_excel_update_rectangle;

Update excel "rectangle" within a named range without using column names

Easier with column names?

WORKING CODE
============

     PYTHON CODE


WORKING CODE
============

     PYTHON CODE

       ws = wb.get_sheet_by_name('class');
       rows = dataframe_to_rows(clas);
       for r_idx in range(3):;       * start in row 3;
           for c_idx in range(2):;   * start in column 2;
                c=c_idx+1;
                r=r_idx+1;
                ws.cell(row=r_idx+3, column=c_idx+2,value=clas.iloc[r-1,c-1]);

HAVE
====

Exel sheet

d:/xls/class.xlsx

------------------------------------------
|    A       |     B      |    C         |
|----------------------------------------+
|NAME        |SEX         |AGE           |
|------------+------------+--------------|
|Alfred      |M           |14            |
|------------+------------+--------------+
|Alice       |F           |13            |
|------------+------------+--------------+
|Barbara     |F           |13            |
|------------+------------+--------------+
|Carol       |F           |14            |
|------------+------------+--------------+
|Henry       |M           |14            |
|------------+------------+--------------+
|James       |M           |12            |
|------------+------------+--------------+
|Jane        |F           |12            |
|------------+------------+--------------+
|Janet       |F           |15            |
|------------+------------+--------------+
|Jeffrey     |M           |13            |
|------------+------------+--------------+

[CLASS]


Up to 40 obs SD1.CLASS total obs=3

Obs    NAME       SEX    AGE

 1     Alice       A     112
 2     John        J     111
 3     William     W     114

WANT
====

------------------------------------------
|    A       |     B      |    C         |
|----------------------------------------+
|NAME        |SEX         |AGE           |
|------------+------------+--------------|
|Alfred      |M           |14            |
|------------+------------+--------------+
|Alice       |                           |
|------------+                           +
|Barbara     |    UPDATE THIS INPLACE    |
|------------+                           +
|Carol       |                           |
|------------+------------+--------------+
|Henry       |M           |14            |
|------------+------------+--------------+
|James       |M           |12            |
|------------+------------+--------------+
|Jane        |F           |12            |
|------------+------------+--------------+
|Janet       |F           |15            |
|------------+------------+--------------+
|Jeffrey     |M           |13            |
|------------+------------+--------------+


------------------------------------------
|    A       |     B      |    C         |
|----------------------------------------+
|NAME        |SEX         |AGE           |
|------------+------------+--------------|
|Alfred      |M           |14            |
|------------+------------+--------------+
|Alice       |A            112           |
|------------+                           +
|Barbara     |B            111           |
|------------+                           +
|Carol       |C            114           |
|------------+------------+--------------+
|Henry       |M           |14            |
|------------+------------+--------------+
|James       |M           |12            |
|------------+------------+--------------+
|Jane        |F           |12            |
|------------+------------+--------------+
|Janet       |F           |15            |
|------------+------------+--------------+
|Jeffrey     |M           |13            |
|------------+------------+--------------+

[CLASS]
*                _              _       _
 _ __ ___   __ _| | _____    __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/ | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|

;

options validvarname=upcase;
libname sd1 "d:/sd1";
data sd1.class;
  retain name sex age;
  set sashelp.class(keep=name age where=(name in ('Alice','Barbara','Carol')));
  sex=substr(name,1,1);
  age=age+99;
  keep name sex age;
run;quit;
libname sd1 clear;


%utlfkil(d:/xls/class.xlsx);
libname xel "d:/xls/class.xlsx";
data xel.class;
  set sashelp.class(keep=name sex age);
run;quit;
libname xel clear;

/*
Up to 40 obs from sd1.class total obs=3

Obs    SEX    AGE

 1      A     112
 2      J     111
 3      W     114
*/

*          _       _   _
 ___  ___ | |_   _| |_(_) ___  _ __
/ __|/ _ \| | | | | __| |/ _ \| '_ \
\__ \ (_) | | |_| | |_| | (_) | | | |
|___/\___/|_|\__,_|\__|_|\___/|_| |_|

;

* this works;
%utl_submit_py64old("
from openpyxl.utils.dataframe import dataframe_to_rows;
from openpyxl import Workbook;
from openpyxl import load_workbook;
from sas7bdat import SAS7BDAT;
with SAS7BDAT('d:/sd1/class.sas7bdat') as m:;
.   clas = m.to_data_frame();
wb = load_workbook(filename='d:/xls/class.xlsx', read_only=False);
ws = wb.get_sheet_by_name('class');
rows = dataframe_to_rows(clas);
for r_idx in range(3):;
.   for c_idx in range(2):;
.        c=c_idx+1;
.        r=r_idx+1;
.        ws.cell(row=r_idx+3, column=c_idx+2,value=clas.iloc[r-1,c-1]);
wb.save('d:/xls/class.xlsx');
");





