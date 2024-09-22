# Electronic-Sheet
The Electronic Sheet Project is a custom spreadsheet application that provides essential features for data manipulation and calculations. It is designed for ease of use and efficient processing of large datasets.

Features:
1. Cell-based data input
2. Basic mathematical functions (addition, subtraction, multiplication, division)
3. Ability to save and load spreadsheet.
4. Formula support
   
Dependencies:
1. numpy==1.26.4
2. pandas==2.2.1
3. matplotlib==3.8.3

How to Use the Electronic Sheet Program:
This program allows users to manage and manipulate spreadsheet-like data, offering the ability to add, update, calculate, and save data in multiple formats.

1. Launching the Program
You can start the program by executing the script from the command line:
```bash
python main.py
```
On startup, the program will prompt you to either:
1. Load a file: Load an existing .json file that contains saved sheet data.
2. Create a new sheet: Start with an empty sheet by providing a file name.

2. Adding Data
You can add data to the sheet by specifying whether you want to add a column, row, or individual cell.
Add a Column:
```bash
ADD COLUMN <Column Name> <Row Number> <Value> , <Row Number> <Value> , ...
```
example-
```bash
ADD COLUMN A 1 100 , 2 200
```
Add a Row:
```bash
ADD ROW <Row Number> <Column Name> <Value> , <Column Name> <Value> , ...
```
example-
```bash
ADD ROW 1 A 100 , B 200
```
Add a Cell:
```bash
ADD CELL <Column Name> <Row Number> <Value>
```
example-
```bash
ADD CELL A 1 100
```

3. Updating Data
You can also update existing data with similar commands:
Update a Column:
```bash
UPDATE COLUMN <Column Name> <Row Number> <Value> , <Row Number> <Value> , ...
```
Update a Row:
```bash
UPDATE ROW <Row Number> <Column Name> <Value> , <Column Name> <Value> , ...
```
Update a Cell:
```bash
UPDATE CELL <Column Name> <Row Number> <Value>
```

4. Calculating Data
You can perform calculations on specific cells or across a range of cells. The results will be stored in a target cell.
Calculation Commands:
```bash
CALCULATE <Target Index> = <Function> <Indexes>
```
Supported functions:
SUM
AVERAGE
MAX
MIN
COUNTIF
SQUARE
SQRT
example-
```bash
CALCULATE A3 = SUM A1:A2
```

5. Saving the Data
You can save the sheet in multiple formats. Supported formats include JSON, Excel, CSV, HTML, and PDF.
Save as JSON:
```bash
SAVE FILE <File Name>
```
Save as Excel:
```bash
SAVE to EXCEL <File Name>
```
Save as CSV:
```bash
SAVE to CSV <File Name>
```
Save as HTML:
```bash
SAVE to HTML <File Name>
```
Save as PDF:
```bash
SAVE to PDF <File Name>
```

6. Exiting the Program
To exit the program, simply type:
```bash
EXIT
```
