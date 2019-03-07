import { Row } from "./row";

export class Table {
    private header = '<Table>';
    private footer = '</Table>';

    public rows: Row[] = [];

    public addRow(row: Row) {
        this.rows.push(row);
        return this.rows.length - 1;
    }

    public exportTable(): string {
        return this.header + this.rows.map(row => row.exportRow()).join('') + this.footer;
    }

}