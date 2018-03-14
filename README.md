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

### Example
  
tbl2latex('mytable.xlsx','tables', 'A1:C3', [3, 2], ['f','e'])      
Read the content in file <mytable.xlsx> on sheet <tables> in range <A1:C3> where the first column with numerical values  
will be rounded to 3 decimals and the second one rounded to 2 decimals. The first column will be displayed as a float and the second column in exponential form.  

## Convert existing MATLAB table or cell array

tbl2latex(tbl)     
Converts the contents of <tbl> to LaTeX table code. <tbl> can be a MATLAB table or cell array.  

tbl2latex(tbl, dcl)   
dcl = A vector containing the the desired number of decimals for every column from left to right that contains numerical values.  
The default value is 2 decimals.  

tbl2latex(tbl, dcl, type)     
type = A vector of characters containing the desired formatting of the numerical values for every column from left to right  
using the formatSpec from the sprintf function.  
The default value is 'f'.  
See further in doc sprintf.  

## Function History

Version 1.01  
Last updated : 2018-03-14 by Anton Lydell  

This function (version 1.0) was originally created by:  
Anton Lydell  
Student M.Sc., Energy and Environmental Technology  
Linköping University  
2018-02-25  
https://www.linkedin.com/in/antonlydell/  
