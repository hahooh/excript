import { Data, Cell, Row, Table, WorkSheet, Excel } from '..';

const excelHeader = '<?xml version="1.0" encoding="UTF-8"?>' +
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"' +
    '          xmlns:c="urn:schemas-microsoft-com:office:component:spreadsheet" ' +
    '          xmlns:html="http://www.w3.org/TR/REC-html40" xmlns:o="urn:schemas-microsoft-com:office:office" ' +
    '          xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" ' +
    '          xmlns:x2="http://schemas.microsoft.com/office/excel/2003/xml" ' +
    '          xmlns:x="urn:schemas-microsoft-com:office:excel" ' +
    '          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '<OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office"></OfficeDocumentSettings>' +
    '<ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">' +
    '        <WindowHeight>6135</WindowHeight>' +
    '        <WindowWidth>8445</WindowWidth>' +
    '        <WindowTopX>240</WindowTopX>' +
    '        <WindowTopY>120</WindowTopY>' +
    '        <ProtectStructure>False</ProtectStructure>' +
    '        <ProtectWindows>False</ProtectWindows>' +
    '</ExcelWorkbook>';
const style: string = '<Styles>' +
    '      <Style ss:ID="Default" ss:Name="Normal">' +
    '            <Alignment ss:Vertical="Bottom" />' +
    '            <Borders />' +
    '            <Font />' +
    '            <Interior />' +
    '            <NumberFormat />' +
    '            <Protection />' +
    '      </Style>' +
    '</Styles>';

test('export string data', () => {
    const data = new Data('hello world');
    const excelData = data.exportData();
    expect(excelData).toBe('<Data ss:Type="String">hello world</Data>');
});

test('export number data', () => {
    const data = new Data(55, 'Number');
    const excelData = data.exportData();
    expect(excelData).toBe('<Data ss:Type="Number">55</Data>');
});

test('export string data with encode XML', () => {
    const data = new Data('a<>"\'b');
    const excelData = data.exportData();
    expect(excelData).toBe('<Data ss:Type="String">a&lt;&gt;&quot;&apos;b</Data>')
});

test('export cell', () => {
    const data = new Data('helloWorld');
    const cell = new Cell(data);
    expect(cell.exportCell()).toBe(`<Cell>${data.exportData()}</Cell>`);
});

test('export row', () => {
    const row = new Row();
    const cells: Cell[] = [];
    [1, 'world', 4, 5, 6, 'Hello'].forEach(el => {
        const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
        const cell = new Cell(data);
        row.addCell(cell);
        cells.push(cell);
    });
    expect(row.exportRow()).toBe(`<Row>${cells.map(cell => cell.exportCell()).join('')}</Row>`);
});

test('export table', () => {
    const table = new Table();
    const row1 = new Row();
    const row2 = new Row();
    [1, 'world', 4, 5, 6, 'Hello'].forEach(el => {
        const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
        const cell = new Cell(data);
        row1.addCell(cell);
    });

    [7, 'compute', 8, 9, 10, 'things', 'and <more>'].forEach(el => {
        const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
        const cell = new Cell(data);
        row2.addCell(cell);
    });

    table.addRow(row1);
    table.addRow(row2);

    expect(table.exportTable()).toBe(`<Table>${row1.exportRow() + row2.exportRow()}</Table>`);
});

test('export work sheet noname', () => {
    const table1 = new Table();
    const row1 = new Row();
    const row2 = new Row();
    [1, 'world', 4, 5, 6, 'Hello'].forEach(el => {
        const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
        const cell = new Cell(data);
        row1.addCell(cell);
    });

    [7, 'compute', 8, 9, 10, 'things', 'and <more>'].forEach(el => {
        const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
        const cell = new Cell(data);
        row2.addCell(cell);
    });

    table1.addRow(row1);
    table1.addRow(row2);

    const table2 = new Table();
    const row3 = new Row();
    const row4 = new Row();
    [1, 'world', 4, 5, 6, 'Hello'].forEach(el => {
        const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
        const cell = new Cell(data);
        row3.addCell(cell);
    });

    [7, 'compute', 8, 9, 10, 'things', 'and <more>'].forEach(el => {
        const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
        const cell = new Cell(data);
        row4.addCell(cell);
    });

    table2.addRow(row1);
    table2.addRow(row2);

    const workSheet = new WorkSheet();
    workSheet.addTable(table1);
    workSheet.addTable(table2);

    expect(workSheet.exportWorkSheet()).toBe(`<Worksheet ss:Name="">${table1.exportTable() + table2.exportTable()}</Worksheet>`);
});

test('export work sheet with name', () => {
    const table1 = new Table();
    const row1 = new Row();
    const row2 = new Row();
    [1, 'world', 4, 5, 6, 'Hello'].forEach(el => {
        const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
        const cell = new Cell(data);
        row1.addCell(cell);
    });

    [7, 'compute', 8, 9, 10, 'things', 'and <more>'].forEach(el => {
        const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
        const cell = new Cell(data);
        row2.addCell(cell);
    });

    table1.addRow(row1);
    table1.addRow(row2);

    const table2 = new Table();
    const row3 = new Row();
    const row4 = new Row();
    [1, 'world', 4, 5, 6, 'Hello'].forEach(el => {
        const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
        const cell = new Cell(data);
        row3.addCell(cell);
    });

    [7, 'compute', 8, 9, 10, 'things', 'and <more>'].forEach(el => {
        const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
        const cell = new Cell(data);
        row4.addCell(cell);
    });

    table2.addRow(row1);
    table2.addRow(row2);

    const workSheet = new WorkSheet('noname');
    workSheet.addTable(table1);
    workSheet.addTable(table2);

    expect(workSheet.exportWorkSheet()).toBe(`<Worksheet ss:Name="noname">${table1.exportTable() + table2.exportTable()}</Worksheet>`);
});

test('export excel without name', () => {
    const workSheets: WorkSheet[] = [];
    ['', 'noname'].forEach(rowName => {
        const table1 = new Table();
        const row1 = new Row();
        const row2 = new Row();
        [1, 'world', 4, 5, 6, 'Hello'].forEach(el => {
            const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
            const cell = new Cell(data);
            row1.addCell(cell);
        });

        [7, 'compute', 8, 9, 10, 'things', 'and <more>'].forEach(el => {
            const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
            const cell = new Cell(data);
            row2.addCell(cell);
        });

        table1.addRow(row1);
        table1.addRow(row2);

        const table2 = new Table();
        const row3 = new Row();
        const row4 = new Row();
        [1, 'world', 4, 5, 6, 'Hello'].forEach(el => {
            const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
            const cell = new Cell(data);
            row3.addCell(cell);
        });

        [7, 'compute', 8, 9, 10, 'things', 'and <more>'].forEach(el => {
            const data = new Data(el, typeof el === 'string' ? 'String' : 'Number');
            const cell = new Cell(data);
            row4.addCell(cell);
        });

        table2.addRow(row1);
        table2.addRow(row2);

        const workSheet = new WorkSheet(rowName);
        workSheet.addTable(table1);
        workSheet.addTable(table2);
        workSheets.push(workSheet);
    });

    const excel = new Excel();
    workSheets.forEach(workSheet => excel.addWorkSheets(workSheet));
    expect(excel.exportExcel()).toBe(excelHeader + style + workSheets.map((workSheet, index) => {
        if (!workSheet.workSheetName) {
            workSheet.workSheetName = 'work sheet ' + index;
        }
        return workSheet.exportWorkSheet();
    }).join('') + '</Workbook>');
});