ods excel update excel sheet in place python                                                                                               
                                                                                                                                           
github                                                                                                                                     
run;quit;                                                                                                                                  
https://github.com/rogerjdeangelis/utl-ods-excel-update-excel-sheet-in-place-python                                                        
                                                                                                                                           
SAS Forum                                                                                                                                  
https://tinyurl.com/y9zltsf8                                                                                                               
https://communities.sas.com/t5/ODS-and-Base-Reporting/excel-engine-replace-the-content-of-a-sheet-created-using-ods/m-p/663764             
                                                                                                                                           
   Process                                                                                                                                 
                                                                                                                                           
       a. Use ods excel to create a class sheet of male students                                                                           
       b. drop the table inside the class sheet                                                                                            
       c. replace the male data sheet with female data                                                                                     
                                                                                                                                           
I believe this can be done in pure SAS but you will be better off                                                                          
taking the oportunity to learn a little python.                                                                                            
                                                                                                                                           
There are many in place excel update repos on my site                                                                                      
                                                                                                                                           
RELATED REPOS                                                                                                                              
https://github.com/rogerjdeangelis/utl_excel_update_rectangle                                                                              
https://github.com/rogerjdeangelis?tab=repositories&q=UPDATE+EXCEL&type=&language=                                                         
                                                                                                                                           
*_                   _                                                                                                                     
(_)_ __  _ __  _   _| |_                                                                                                                   
| | '_ \| '_ \| | | | __|                                                                                                                  
| | | | | |_) | |_| | |_                                                                                                                   
|_|_| |_| .__/ \__,_|\__|                                                                                                                  
        |_|                                                                                                                                
;                                                                                                                                          
                                                                                                                                           
%utlfkil(d:/xls/class.xlsx);                                                                                                               
ods excel file="d:/xls/classx.xlsx" options(sheet_name='Class');                                                                           
proc print data=sashelp.class(where=(sex="M")) noobs;                                                                                      
run;                                                                                                                                       
ods excel close;                                                                                                                           
*                _                                                                                                                         
 _ __ ___   __ _| | ___  ___                                                                                                               
| '_ ` _ \ / _` | |/ _ \/ __|                                                                                                              
| | | | | | (_| | |  __/\__ \                                                                                                              
|_| |_| |_|\__,_|_|\___||___/                                                                                                              
                                                                                                                                           
;                                                                                                                                          
d:/xls/class.xlsx                                                                                                                          
                                                                                                                                           
      ----------------------------------------------                                                                                       
   1  |NAME   SEX         AGE     HEIGHT     WEIGHT|                                                                                       
      |--------------------------------------------|                                                                                       
   2  |James   | M|        12|      57.3|        83|                                                                                       
      |--------+--+----------+----------+----------|                                                                                       
   3  |Thomas  | M|        11|      57.5|        85|                                                                                       
      |--------+--+----------+----------+----------|                                                                                       
   4  |John    | M|        12|        59|      99.5|                                                                                       
      |--------+--+----------+----------+----------|                                                                                       
   5  |Jeffrey | M|        13|      62.5|        84|                                                                                       
      |--------+--+----------+----------+----------|                                                                                       
   6  |Henry   | M|        14|      63.5|     102.5|                                                                                       
      |--------+--+----------+----------+----------|                                                                                       
   7  |Robert  | M|        12|      64.8|       128|                                                                                       
      |--------+--+----------+----------+----------|                                                                                       
   8  |William | M|        15|      66.5|       112|                                                                                       
      |--------+--+----------+----------+----------|                                                                                       
   9  |Ronald  | M|        15|        67|       133|                                                                                       
      |--------+--+----------+----------+----------|                                                                                       
  10  |Alfred  | M|        14|        69|     112.5|                                                                                       
      |--------+--+----------+----------+----------|                                                                                       
  11  |Philip  | M|        16|        72|       150|                                                                                       
      ----------------------------------------------                                                                                       
                                                                                                                                           
  [CLASS]                                                                                                                                  
                                                                                                                                           
*            _               _                                                                                                             
  ___  _   _| |_ _ __  _   _| |_                                                                                                           
 / _ \| | | | __| '_ \| | | | __|                                                                                                          
| (_) | |_| | |_| |_) | |_| | |_                                                                                                           
 \___/ \__,_|\__| .__/ \__,_|\__|                                                                                                          
  __            |_|       _                                                                                                                
 / _| ___ _ __ ___   __ _| | ___  ___                                                                                                      
| |_ / _ \ '_ ` _ \ / _` | |/ _ \/ __|                                                                                                     
|  _|  __/ | | | | | (_| | |  __/\__ \                                                                                                     
|_|  \___|_| |_| |_|\__,_|_|\___||___/                                                                                                     
;                                                                                                                                          
                                                                                                                                           
d:/xls/class.xlsx  (same workbook);                                                                                                        
                                                                                                                                           
This is an update in place of the class sheet                                                                                              
                                                                                                                                           
    ----------------------------------------------                                                                                         
 1  |                                            | * I leave it to you                                                                     
    |--------------------------------------------|   top populate header                                                                   
 2  |Joyce   | F|        11|      51.3|      50.5|   Just use a row for loop                                                               
    |--------+--+----------+----------+----------|   in python.                                                                            
 3  |Louise  | F|        12|      56.3|        77|                                                                                         
    |--------+--+----------+----------+----------|                                                                                         
 4  |Alice   | F|        13|      56.5|        84|                                                                                         
    |--------+--+----------+----------+----------|                                                                                         
 5  |Jane    | F|        12|      59.8|      84.5|                                                                                         
    |--------+--+----------+----------+----------|                                                                                         
 6  |Janet   | F|        15|      62.5|     112.5|                                                                                         
    |--------+--+----------+----------+----------|                                                                                         
 7  |Carol   | F|        14|      62.8|     102.5|                                                                                         
    |--------+--+----------+----------+----------|                                                                                         
 8  |Judy    | F|        14|      64.3|        90|                                                                                         
    |--------+--+----------+----------+----------|                                                                                         
 9  |Barbara | F|        13|      65.3|        98|                                                                                         
    |--------+--+----------+----------+----------|                                                                                         
10  |Mary    | F|        15|      66.5|       112|                                                                                         
    ----------------------------------------------                                                                                         
*                                                                                                                                          
 _ __  _ __ ___   ___ ___  ___ ___                                                                                                         
| '_ \| '__/ _ \ / __/ _ \/ __/ __|                                                                                                        
| |_) | | | (_) | (_|  __/\__ \__ \                                                                                                        
| .__/|_|  \___/ \___\___||___/___/                                                                                                        
|_|                                                                                                                                        
;                                                                                                                                          
                                                                                                                                           
* CREATE AN CLASS SHEET OF MALES USING OCD EXCEL;                                                                                          
                                                                                                                                           
%utlfkil(d:/xls/classx.xls);                                                                                                               
                                                                                                                                           
* CREATE EXCEL WORKBOOK USING ODS EXCEL;                                                                                                   
ods excel file="d:/xls/class.xlsx" options(sheet_name='Class');                                                                            
proc print data=sashelp.class(where=(sex="M")) noobs;                                                                                      
run;                                                                                                                                       
ods excel close;                                                                                                                           
                                                                                                                                           
* DROP THE TABLE IN SHEET CLASS;                                                                                                           
libname xel "d:/xls/class.xlsx";                                                                                                           
proc sql;                                                                                                                                  
  drop table xel.'Class$'n;;                                                                                                               
run;quit;                                                                                                                                  
libname xel clear;                                                                                                                         
                                                                                                                                           
* GET FEMALE STUDENTS;                                                                                                                     
options validvarname=upcase;                                                                                                               
libname sd1 "d:/sd1";                                                                                                                      
data sd1.class_f;                                                                                                                          
    set sashelp.class (where=(sex='F'));                                                                                                   
run;                                                                                                                                       
                                                                                                                                           
* uSE PYTHON TO UPDATE IN PLACE;                                                                                                           
%utl_submit_py64_38("                                                                                                                      
from openpyxl.utils.dataframe import dataframe_to_rows;                                                                                    
from openpyxl import Workbook;                                                                                                             
from openpyxl import load_workbook;                                                                                                        
from sas7bdat import SAS7BDAT;                                                                                                             
with SAS7BDAT('d:/sd1/class_f.sas7bdat') as m:;                                                                                            
.   clas = m.to_data_frame();                                                                                                              
wb = load_workbook(filename='d:/xls/Class.xls', read_only=False);                                                                          
ws = wb.get_sheet_by_name('Class');                                                                                                        
rows = dataframe_to_rows(clas);                                                                                                            
for r_idx in range(9):;                                                                                                                    
.   for c_idx in range(5):;                                                                                                                
.        c=c_idx+1;                                                                                                                        
.        r=r_idx+1;                                                                                                                        
.        ws.cell(row=r_idx+2, column=c_idx+1,value=clas.iloc[r-1,c-1]);                                                                    
wb.save('d:/xls/class.xlsx');                                                                                                              
");                                                                                                                                        
                                                                                                                                           
                                                                                                                                           
