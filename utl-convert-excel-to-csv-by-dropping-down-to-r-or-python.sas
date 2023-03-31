%let pgm=utl-convert-excel-to-csv-by-dropping-down-to-r-or-python;

Convert excel to csv by dropping down to r or python

  1. R
    a. Create a workbook with one sheet using sashelp.class
    b. Convert the sheet to a csv
  2. Python
    a. Create a workbook with one sheet using sashelp.class
    b. Convert the sheet to a csv
  3. Attempt failed using Poweshell

github
https://tinyurl.com/ycy88w4f
https://github.com/rogerjdeangelis/utl-convert-excel-to-csv-by-dropping-down-to-r-or-python

For bells and whistles, to quote Iqn Whitlock use these links

https://www.the-analytics.club/convert-from-excel-to-csv
https://stackoverflow.com/questions/51273978/convert-xls-to-csv-r-tried-rio-package

Could not get my dropdown to powershell to work
https://blog.powercram.com/2015/05/use-powershell-to-save-excel-worksheet-as-csv.html

/*               _                            _   ____
 _ __ ___   __ _| | _____    _____  _____ ___| | |  _ \
| `_ ` _ \ / _` | |/ / _ \  / _ \ \/ / __/ _ \ | | |_) |
| | | | | | (_| |   <  __/ |  __/>  < (_|  __/ | |  _ <
|_| |_| |_|\__,_|_|\_\___|  \___/_/\_\___\___|_| |_| \_\

*/
 /*---- create input excel file ----*/
libname sd1 "d:/sd1";

data sd1.class;
  set sashelp.class;
run;quit;

/*---- delete workbook if it exists ----*/
%utlfkil(d:/xls/class.xlsx);

%utl_submit_r64('
   library(haven);
   library(XLConnect);
   have<-read_sas("d:/sd1/class.sas7bdat");
   have;
   wb <- loadWorkbook("d:/xls/class.xlsx", create = TRUE);
   createSheet(wb, name = "class");
   writeWorksheet(wb, have, sheet = "class");
   saveWorkbook(wb);
');

/*___              _       _   _
|  _ \   ___  ___ | |_   _| |_(_) ___  _ __
| |_) | / __|/ _ \| | | | | __| |/ _ \| `_ \
|  _ <  \__ \ (_) | | |_| | |_| | (_) | | | |
|_| \_\ |___/\___/|_|\__,_|\__|_|\___/|_| |_|

*/

/*---- create csv ----*/

%utlfkil(d:/csv/class.csv);

%let inp = 'd:/xls/class.xlsx' ;
%let out = 'd:/csv/class.csv'  ;

%utl_submit_r64("
library('readxl');
library('readr');
indata <- read_excel(&inp);
write_csv(indata, file=&out);
");


/*               _                            _
 _ __ ___   __ _| | _____    _____  _____ ___| |  _ __  _   _
| `_ ` _ \ / _` | |/ / _ \  / _ \ \/ / __/ _ \ | | `_ \| | | |
| | | | | | (_| |   <  __/ |  __/>  < (_|  __/ | | |_) | |_| |
|_| |_| |_|\__,_|_|\_\___|  \___/_/\_\___\___|_| | .__/ \__, |
                                                 |_|    |___/
*/

/*---- create input excel file ----*/

%utlfkil(d:/xls/class_py.xlsx)

%let inp = 'd:/xls/class_py.xlsx' ;
%let sas = 'd:/sd1/class.sas7bdat' ;

%utl_submit_py64_310("
import pandas as pd;
import pyreadstat;
import openpyxl;
want, metaWant = pyreadstat.read_sas7bdat(&sas);
print(want);
print('COLUMNS:  ', metaWant.column_names);
print('LABELS:   ',metaWant.column_labels);
print('ROWS:     ',metaWant.number_rows);
print('FIELDS:   ',metaWant.number_columns);
print('DSN LABEL:',metaWant.file_label);
print('ENCODING: ',metaWant.file_encoding);
want.to_excel(r&inp,sheet_name='class', index=False);
");


/*           _   _                             _       _   _
 _ __  _   _| |_| |__   ___  _ __    ___  ___ | |_   _| |_(_) ___  _ __
| `_ \| | | | __| `_ \ / _ \| `_ \  / __|/ _ \| | | | | __| |/ _ \| `_ \
| |_) | |_| | |_| | | | (_) | | | | \__ \ (_) | | |_| | |_| | (_) | | | |
| .__/ \__, |\__|_| |_|\___/|_| |_| |___/\___/|_|\__,_|\__|_|\___/|_| |_|
|_|    |___/
*/

%utlfkil(d:/csv/class_py.csv);

%let inp = 'd:/xls/class_py.xlsx' ;
%let out = 'd:/csv/class_py.csv'  ;

%utl_submit_py64_310("
import os;
import pandas as pd;
import openpyxl;
df = pd.read_excel(&inp);
df.to_csv(&out,index=False);
");

/*                                _          _ _
 _ __   _____      _____ _ __ ___| |__   ___| | |
| `_ \ / _ \ \ /\ / / _ \ `__/ __| `_ \ / _ \ | |
| |_) | (_) \ V  V /  __/ |  \__ \ | | |  __/ | |
| .__/ \___/ \_/\_/ \___|_|  |___/_| |_|\___|_|_|
|_|
*/

/*---- COULD NOT GET THID TO WORK ----*/
/*---- SUSPECT NOR SUPPORT FOR LEGACY OFFICE (MAYBE ONLY 1 YEAR ONLY MS365) ----*/

%utl_submit_ps64('
$outFile   = "d:/csv/class.csv";
$excelFile = "d:/xls/class.xlsx";
$E = New-Object -ComObject Excel.Application;
$E.Visible = $false;
$E.DisplayAlerts = $false;
$wb = $E.Workbooks.Open($excelFile);
$ws.SaveAs($outFile,6);
$E.Quit();
');

/*           _               _
  ___  _   _| |_ _ __  _   _| |_
 / _ \| | | | __| `_ \| | | | __|
| (_) | |_| | |_| |_) | |_| | |_
 \___/ \__,_|\__| .__/ \__,_|\__|
                |_|
*/

NAME,SEX,AGE,HEIGHT,WEIGHT
Alfred,M,14.0,69.0,112.5
Alice,F,13.0,56.5,84.0
Barbara,F,13.0,65.3,98.0
Carol,F,14.0,62.8,102.5
....

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
