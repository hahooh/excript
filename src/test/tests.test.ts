import { Data, Cell } from '..';
import { Row } from '../components/row';

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
    const cellExport = cell.exportCell();
    expect(`<Cell>${data.exportData()}</Cell>`).toBe(cellExport);
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
    expect(`<Row>${cells.map(cell => cell.exportCell()).join('')}</Row>`)
});