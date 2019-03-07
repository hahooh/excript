import { Table } from "./table";

export class WorkSheet {
    private footer: string = '</Worksheet>';

    public workSheetName: string = '';
    public tables: Table[] = [];

    constructor(workSheetName: string = '') {
        this.workSheetName = workSheetName;
    }

    public addTable(table: Table) {
        this.tables.push(table);
        return this.tables.length - 1;
    }

    public exportWorkSheet(): string {
        return this.getHeader() + this.tables.map(table => table.exportTable()).join('') + this.footer;
    }

    private getHeader(): string {
        return '<Worksheet ss:Name="' + this.workSheetName + '">';
    }
}