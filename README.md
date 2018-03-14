# tbl2latex

## About

A MATLAB function that reads content from an Excel file and converts it to LaTeX table code or converts a MATLAB table or cell array into LaTeX table code. LaTeX table code is referring to column entries separated by &amp; and rows ended with \\

## Read from Excel file

tbl2latex(filename)
Read the contents of the Excel file <filename> and convert it to LaTeX table code.

tbl2latex(filename,sheet)
Read the contents of the Excel file <filename> on sheet <sheet> and convert it to LaTeX table code.

tbl2latex(filename,sheet,range)
Read the contents of the Excel file <filename> on sheet <sheet> in range <range> and convert it to LaTeX table code.

tbl2latex(filename, sheet, range, dcl) 
dcl = A vector containing the the desired number of decimals for every column from left to right that contains numerical values. 
The default value is 2 decimals.

tbl2latex(filename, sheet, range, dcl, type)
type = A vector of characters containing the desired formatting of the numerical values for every column from left to right
using the formatSpec from the sprintf function.
The default value is 'f'.
See further in doc sprintf.

