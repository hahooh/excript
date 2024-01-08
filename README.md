# Excel XLS

Create excel file in XLS form.
Now it only support basic types such as String and number and it does not support styles.

# Usage
 // create excel<br/>
const excel = new Excel('filename.xls');

// create work sheet<br/>
const workSheet = new WorkSheet('work_sheet_name);

// create table<br/>
const table = new Table();

// create row<br/>
const row = new Row();

// create data cell<br/>
const data = new Data('this is cell value', 'String');
const cell = new Cell(data);

// insert cells in row<br/>
row.addCell(cell);

// insert rows in table<br/>
table.addRow(row);

// insert tables in worksheet<br/>
workSheet.addTable(table);

// insert worksheets in excel<br/>
excel.addWorkSheet(worksheet);

// final excel in XML form<br/>
const excelXML = excel.exportExcel();

# NPM
[you can install from here](https://www.npmjs.com/package/excript)
