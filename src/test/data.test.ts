import { Data } from '..';

test('export string data', () => {
    const data = new Data('hello world');
    const excelData = data.exportData();
    expect('<Data ss:Type="String">hello world</Data>').toBe(excelData);
})