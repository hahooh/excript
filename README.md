# Excel XLS

Create excel file in XLS form.
Now it only support basic types such as String and number and it does not support styles.

# Usage
 // create excel
const excel = new Excel('filename.xls');

// create work sheet
const workSheet = new WorkSheet('work_sheet_name);

// create table 
const table = new Table();

// create row
const row = new Row();

// create data cell
const data = new Data('this is cell value', 'String');
const cell = new Cell(data);

// insert cells in row
row.addCell(cell);

// insert rows in table
table.addRow(row);

// insert tables in worksheet
workSheet.addTable(table);

// insert worksheets in excel
excel.addWorkSheet(worksheet);

// final excel in XML form
const excelXML = excel.exportExcel();
