function tbl2latex(filename, sheet, range, dcl, type)
% ==========================================================================================================
% About
% ==========================================================================================================
%
% A function that reads content from an Excel file and converts it to LaTeX table code
% or converts a MATLAB table or cell array into LaTeX table code.
% LaTeX table code is referring to column entries separated by & and rows ended with \\
%
% ==========================================================================================================
% Read from Excel file
% ==========================================================================================================
%
% tbl2latex(filename)
% Read the contents of the Excel file <filename> and convert it to LaTeX table code.
%
% tbl2latex(filename,sheet)
% Read the contents of the Excel file <filename> on sheet <sheet> and convert it to LaTeX table code.
%
% tbl2latex(filename,sheet,range)
% Read the contents of the Excel file <filename> on sheet <sheet> in range <range> and convert it to LaTeX table code.
%
% tbl2latex(filename, sheet, range, dcl)
% dcl = A vector containing the the desired number of decimals for every column from left to right that contains numerical values.
% The default value is 2 decimals.
%
% tbl2latex(filename, sheet, range, dcl, type)
% type = A vector of characters containing the desired formatting of the numerical values for every column from left to right
% using the formatSpec from the sprintf function.
% The default value is 'f'.
% See further in doc sprintf.
%
% Example
% ----------------------------------------------------------------------------------------------------------
%
% tbl2latex('mytable.xlsx','tables', 'A1:C3', [3, 2], ['f','e'])
% Read the content in file <mytable.xlsx> on sheet <tables> in range <A1:C3> where the first column with numerical values
% will be rounded to 3 decimals and the second one rounded to 2 decimals. The first column will be displayed as a float and the second
% column in exponential form.
%
% Contents of the Excel file <mytable.xlsx>
% -----------------------------------------
%
%   A       B        C       D
% 1 Type    U-value  Cost    g-value
% 2 Window  0,953    6500    0,5
% 3 Wall    0,16     8425,5
% 4 Roof    0,10     7425
%
% Output from tbl2latex
% -----------------------------------------
%
% Type     &  U-value   &      Cost  \\
% Window   &    0.953   &  6.50e+03  \\
% Wall     &    0.160   &  8.43e+03  \\
%
% ==========================================================================================================
% Convert existing MATLAB table or cell array
% ==========================================================================================================
%
% tbl2latex(tbl)
% Converts the contents of <tbl> to LaTeX table code. <tbl> can be a MATLAB table or cell array.
%
% tbl2latex(tbl, dcl)
% dcl = A vector containing the the desired number of decimals for every column from left to right that contains numerical values.
% The default value is 2 decimals.
%
% tbl2latex(tbl, dcl, type)
% type = A vector of characters containing the desired formatting of the numerical values for every column from left to right
% using the formatSpec from the sprintf function.
% The default value is 'f'.
% See further in doc sprintf.
%
% ==========================================================================================================
% Function history
% ==========================================================================================================
%
% Version 1.01
% Last updated : 2018-03-18 by Anton Lydell
%
% =========================================================
% This function (version 1.0) was originally created by:
% Anton Lydell
% Student M.Sc., Energy and Environmental Technology
% Linköping University
% 2018-02-25
% https://www.linkedin.com/in/antonlydell/
% =========================================================

%================================================================================================================
% Start of the code
%================================================================================================================

dclDef = 2; % The default amount of decimals
typeDef = 'f'; % The default number representation (float)

dclErrorMsg = ''; % Variable for assigning possible error message if erranous input of the dcl vector
typeErrorMsg = ''; % Variable for assigning possible error message if erranous input of the type vector

%================================================================================================================
% Analyze the imput procedure
%================================================================================================================

switch nargin

case 1  % 1 input argument for MATLAB table and cell array
    if ischar(filename)
      [~, ~, tbl] = xlsread(filename); % Read the file contents to cell array
    else
      tbl = filename;
      if istable(tbl)
        tbl = table2cell(tbl); % Convert to cell array
      end % if istable
    end % if ischar

    % Standard values
    [rows, cols] = size(tbl); % Size of the table
    dcl = dclDef * ones(1,cols); % Assign standard values
    type(1:cols) = typeDef; % Assign standard values

case 2 % 2 input arguments
    if ischar(filename)
      [~, ~, tbl] = xlsread(filename,sheet); % Read the file contents to cell array

      % Standard values
      [rows, cols] = size(tbl); % Size of the table
      dcl = dclDef * ones(1,cols); % Assign standard values
    else
      tbl = filename;
      dcl = sheet;
      if istable(tbl)
        tbl = table2cell(tbl); % Convert to cell array
      end % if istable
    end % if ischar

    % Standard values
    [rows, cols] = size(tbl); % Size of the table
    type(1:cols) = typeDef; % Assign standard values

case 3 % 3 Input arguments
  if ischar(filename)
    [~, ~, tbl] = xlsread(filename,sheet,range); % Read the file contents to cell array

    % Standard values
    [rows, cols] = size(tbl); % Size of the table
    dcl = dclDef * ones(1,cols); % Assign standard values
    type(1:cols) = typeDef; % Assign standard values

  else % If the input is not for reading and Excel-file
    tbl = filename;
    dcl = sheet;
    type = range;
    if istable(tbl)
      tbl = table2cell(tbl); % Convert to cell array
    end % if istable
    [rows, cols] = size(tbl); % Size of the table
  end % if ischar

case {4, 5} % 4-5 Input arguments
  [~, ~, tbl] = xlsread(filename,sheet,range); % Read the file contents to cell array
  [rows, cols] = size(tbl); % Size of the table
  if nargin == 4
    type(1:cols) = typeDef; % Assign standard values
  end % if nargin

otherwise
    disp('Error')
end % switch

idx = 1; % Start index for the dcl and type vectors

% If empty inputs to dcl and type
if isempty(dcl)
    dcl = dclDef * ones(1,cols); % Assign standard values
    if nargin == 3 & ~ischar(filename) % If wanting standard values for dcl and change the type vector for MATLAB table and cell array
      
    elseif nargin ~= 5 % If wanting standard values for dcl and change the type vector for Excel file
      dclErrorMsg = sprintf(' WARNING: \n The entered <dcl> vector was empty and has been changed to the standard value <%d> to prevent the program from crashing.',dclDef);
    end % if
end % if

if isempty(type)
    type(1:cols) = typeDef; % Assign standard values
    typeErrorMsg = sprintf(' WARNING: \n The entered <type> vector was empty and has been changed to the standard value ''%s'' to prevent the program from crashing.',typeDef);
end % if

%================================================================================================================
% Convert numerics to string
%================================================================================================================

dclOriginal = length(dcl); % The original length of the dcl vector
typeOriginal = length(type); % The original length of the type vector

for row = 1:rows
  countNumCol = 0; % Reset/initiate counter of numerical columns

  for col = 1:cols
    if isnumeric(tbl{row,col}) % if that place in the cell array contains a number

      if isnan(tbl{row,col}) % If the cell is a blank cell in Excel
        tbl{row,col} = '';
      else
        countNumCol = countNumCol + 1; % Count the number of columns with numerical values
        formatDcl = ['%.', num2str(dcl(idx)),type(idx)]; % For rounding the numbers correctly
        tbl{row,col} = num2str(tbl{row,col},formatDcl); % Convert to string
        idx = idx + 1;
      end % if isnan

    end % if isnumeric

    if idx > length(dcl) && col < cols % To avoid the program to crashing if the dcl or type vector is shorter than the actual number of columns with numbers in them
      dcl = [dcl, dclDef*ones(1,cols-col)]; % Append dcl with standard values for the remaining columns to scan through
    end % if idx dcl

    if idx > length(type) && col < cols
      type(idx:cols) = typeDef; % Append type with standard values for the remaining columns to scan through
    end % if idx type

  end % for col
  storeNumCol(row) = countNumCol; % Store the number of numeric columns for current row
  idx = 1; % Reset to column 1 for next row
end % for row

% Analyze possible erranous input of dcl and type vectors

if dclOriginal < max(storeNumCol)
  dclErrorMsg = sprintf(' WARNING: \n The entered <dcl> vector had a length less than the amount of columns with numerical values in the data source. \n <dcl> has been appended with the standard value of <%d> to prevent the program from crashing.',dclDef);
end % if

if typeOriginal < max(storeNumCol)
  typeErrorMsg = sprintf(' WARNING: \n The entered <type> vector had a length less than the amount of columns with numerical values in the data source. \n <type> has been appended with the standard value of ''%s'' to prevent the program from crashing.',typeDef);
end % if


%================================================================================================================
% Find max length of each column
%================================================================================================================

maxLengths = ones(1,cols); % The starting max length for each column

for col = 1:cols
  for row = 1:rows
    if length(tbl{row,col}) > maxLengths(col);
      maxLengths(col) = length(tbl{row,col});
    end % if
  end % for row
end % for col

%================================================================================================================
% Generate the format string
%================================================================================================================

colSep = '  &  '; % For separating the columns
colEntry = '%';
str = 's ';
endOfLine = ' \\\\ ';

for j = 1:cols % Loop over the columns
    if j == 1 & cols == 1 % if the table is only one column
      formatStr = [colEntry, num2str(maxLengths(1)), str, endOfLine, '\n'];
    elseif j == 1 % Initiate formatStr
      formatStr = [colEntry, num2str(maxLengths(1)), str];
    elseif j == cols % End formatStr
      formatStr = [formatStr, colSep, colEntry, num2str(maxLengths(j)), str, endOfLine, '\n'];
    else
      formatStr = [formatStr, colSep, colEntry, num2str(maxLengths(j)), str];
  end % if
end % for j

%================================================================================================================
% Print the table
%================================================================================================================

tbl = tbl'; % To get the right format for fprintf. fprintf reads 1 column at the time
fprintf(formatStr,tbl{:})

% Possible warning message for dcl vector
if ~isempty(dclErrorMsg)
  fprintf('\n\n') % New lines
  disp(dclErrorMsg) % Display warning message
  fprintf('\n')
end % if

% Possible warning message for type vector
if ~isempty(typeErrorMsg)
  fprintf('\n') % New lines
  disp(typeErrorMsg) % Display warning message
  fprintf('\n\n') % New lines
end % if
