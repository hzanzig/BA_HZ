Repository BA_HZ 

This repository contains every code file and data file of the bachelor thesis of Hendrik Zanzig. 
The title of the bachelor thesis: "Mapping of clinical variables to the OMOP vocabulary: Case study in a German university hospital". 
The thesis is written at the Freie Universität Berlin, in cooperation with the Charité Berlin. 



What is in the repository? 

There are seven data files in the repository. 
The following two are the original data files provided by the Charité Berlin:
icd_diagnosis_full_charite.xlsx
ops_codes_full_charite.xlsx

The follwing five data files are the result files of the ICD and OPS mapping process. As the OPS mapping process had to be redone two times, 
there are different versions of results. the ops_full_result.csv file is the final OPS result file. 

icd_full_results.csv
OPSCharite_result.xlsx
OPSCharite_result2.xlsx
OPSCharite_result3.xlsx
ops_full_result.csv


In addition to the result files, there are 12 codefiles. 

.gitignore
database.ini
config.py
icd_mapping.py
ops_mapping.py
ops_mapping_2.py
ops_mapping_3.py
icd_review.py
ops_review.py
ConcatExcelSheetsToyCsv.py
perfect_exec_icd.py
perfect_exec_ops.py

Explanation to the code files: 
database.ini holds the information for the database connection, which is initiated within the python scripzts with the help of the config.py file. 
icd_mapping.py, as well as ops_mapping with it's different versions, hold the code for the conceptual mapping process of the original 
Chrité data to the ATHENA vocabulary. The SQL querys are within those codes. 

icd_review.py and ops_review.py are the files which review the results produced by the *mapping.py code files.
As the initial coding used the python extension "xlwt" and it's function "Workbook", which limits the user to only being able to write the first 65,536 rows 
of a Excel sheet, the ConcatExcelSheetsToyCsv.py code is used to reduce the four Excel sheets within the result .xlsx file to one sheet in a .csv file. 
This is done with the help of data frames, a functionality of the python extension "pandas".

perfect_exec_icd.py and perfect_exec_ops.py is the reduced and adjusted code, which directly uses dataframes and ops_mapping issues were fixed in this version. 
In general these files are the ones, which should be taken into further consideration for future mapping processes. 

The ATHENA vocabulary is not included within this repository, as the rights remain at the Odysseus Data Services, Inc.
The OHDSI ATHENA repository can be found here: https://athena.ohdsi.org/search-terms/start

Please note that the download of additional python extensions, for example "pandas", is neccessary that the code works. 
