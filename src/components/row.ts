import { Cell } from "./cell";

export class Row {
    private header = '<Row>';
    private footer = '</Row>';

    public cells: Cell[] = [];

    public addCell(cell: Cell) {
        this.cells.push(cell);
        return this.cells.length - 1;
    }

    public exportRow(): string {
        return this.header + this.cells.map(cell => cell.exportCell()).join('') + this.footer;
    }
}