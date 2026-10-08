Error Code to SQL Tool
======================

Converts an Excel error-code sheet into SQL for the "internationalization" table.

Requirements
------------
Java 25 or newer (JDK or JRE) must be installed, and the "java" command
must be available in PATH. Download: https://adoptium.net/

How to use
----------
1. Double-click "StartTool.bat" to open the window.
2. Click "Browse..." to select the Excel file (.xlsx).
3. Choose or type the output SQL file path
   (defaults to errorcode.sql next to the Excel file).
4. Click "Generate SQL". A dialog shows the row count and output path.

Excel format
------------
The tool automatically finds the header row containing these columns,
then reads data from the row below it:

  param value | key | Chinese text | English text

Generation rules:
- i18n_key    = RWI. + param value
- cn value    = key + space + Chinese text
- en value    = key + space + English text
- Output contains one DELETE block, one Chinese INSERT block,
  and one English INSERT block.

Command line (optional)
-----------------------
java -jar error-code-convert-sql-1.0.0.jar <excel-file> [output-sql-file]

Options:
  -o, --output <file>   Output SQL file, default errorcode.sql
  --sheet <name>        Sheet name, default the first sheet
  --prefix <prefix>     i18n_key prefix, default RWI.
  --group <name>        i18n_group, default robot
