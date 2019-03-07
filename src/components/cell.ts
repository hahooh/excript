import { Data } from "./data";

export class Cell {
    private header = '<Cell>';
    private footer = '</Cell>';
    public data: Data;

    constructor(data: Data) {
        this.data = data;
    };

    public exportCell(): string {
        return this.header + this.data.exportData() + this.footer;
    }
}