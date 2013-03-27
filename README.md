4gl-excel
=========

A simple library to create an well formatted Excel file from Progress 4GL.

How it works
------------

In order to optimize the process creation of the excel file, first a CSV file is generated, and then imported on excel using "Import data from a text file". After the file has been imported, the file columns are adjusted to fit their width and the head line is formatted using bold text and then a filter is applied. 

How to use
----------

First, you have to define a temp-table that's going to be the data source for the file, the columns data-type and column-label are used to format the excel file. After that you must include the library 4gl-excel.i in your program, and then run the procedures (passing the correct arguments) to create the files, as described below:

1. pi-cria-arquivo-csv
2. pi-cria-arquivo-xls

Example
-------

> {include/relat-excel.i}