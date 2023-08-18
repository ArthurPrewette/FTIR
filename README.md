FTIR
=====================
Data processing for scientific instruments with data outputs in .CSV format.

Dependencies
============
The following libraries must be installed:
* pandas
* glob

How to use 
============
Currently, the script can only be run on a single folder full of .CSV files at one time. I will add nested-folder searching/filemaking capabilities when I have the time. 
The CSV files are sorted based on the numerical characters present in the filename.
* (data1.CSV, data2.CSV, data3.CSV etc.) will be titled Trial 0, Trial 1, and Trial 2 in columns B, C, and D, respectively to preserve time-dependent data collection.
* 
