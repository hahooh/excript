import { Data } from '..';

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