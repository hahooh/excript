import { DataType } from "../types/DataTypes";


export class Data {
    private footer: string = '</Data>';
    public type: string;
    public value: any;

    constructor(value: any, type: DataType = 'String') {
        this.type = type;
        this.value = value;
    }

    public exportData(): string {
        if (!this.value) {
            return this.getHeader() + this.footer;
        }

        return this.getHeader() + this.encodeXML(this.value) + this.footer;
    }

    private getHeader(): string {
        return '<Data ss:Type="' + this.type + '">'
    }

    private encodeXML(string: string): string {
        return string.replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }
}