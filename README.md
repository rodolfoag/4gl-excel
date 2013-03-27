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

    {include/relat-excel.i}

    def var c-csv-file as char no-undo.
    def var c-xls-file as char no-undo. /* will contain the XLS file path created */

    def temp-table tt-data /* Data-Source */
        field cust-name as char column-label "Name"
        field cust-age  as int  column-label "Age".

    create tt-data.
    assign tt-data.cust-name = "Customer Name 1"
           tt-data.cust-age  = 22.
    
    create tt-data.
    assign tt-data.cust-name = "Customer Name 2"
           tt-data.cust-age  = 19.

    run pi-cria-arquivo-csv(input  buffer tt-data:handle,
                            input  session:temp-directory + "file",
                            output c-csv-file).

    run pi-cria-arquivo-xls(input  buffer tt-data:handle,
                            input  c-csv-file,
                            output c-xls-file).