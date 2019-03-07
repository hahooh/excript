import { WorkSheet } from "./workSheet";

export class Excel {
    private header: string = '<?xml version="1.0" encoding="UTF-8"?>' +
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
    private footer: string = '</Workbook>';
    private style: string = '<Styles>' +
        '      <Style ss:ID="Default" ss:Name="Normal">' +
        '            <Alignment ss:Vertical="Bottom" />' +
        '            <Borders />' +
        '            <Font />' +
        '            <Interior />' +
        '            <NumberFormat />' +
        '            <Protection />' +
        '      </Style>' +
        '</Styles>';

    public workSheets: WorkSheet[] = [];
    public fileName: string;

    constructor(filename: string = 'noname.xls') {
        this.fileName = filename;
    }

    public addWorkSheets(workSheet: WorkSheet) {
        this.workSheets.push(workSheet);
    }

    public exportExcel(): string {
        const workSheets = this.workSheets.map((workSheet: WorkSheet, index: number) => {
            if (!workSheet.workSheetName) {
                workSheet.workSheetName = 'work sheet ' + index;
            }
            return workSheet.exportWorkSheet()
        });
        return this.header + this.style + workSheets.join('') + this.footer;
    }
}