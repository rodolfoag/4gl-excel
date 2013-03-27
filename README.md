4gl-excel
=========

A simple library to create an well formatted Excel file from Progress 4GL.

How it works
-------------------------

In order to optimize the process creation of the excel file, first a CSV file is generated, and then imported on excel using "Import data from a text file". After the file has been imported, the file columns are adjusted to fit their width and the head line is formatted using bold text and a filter is applied. 
