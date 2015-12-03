import * as Promise from "bluebird";
import Workbook = require('./Workbook');
import Saveable = require('./Saveable');
declare class SharedStrings extends Saveable {
    protected path: string;
    protected workbook: Workbook;
    xml: any;
    private cache;
    constructor(path: string, workbook: Workbook);
    load(): Promise<any>;
    getIndex(s: string): number;
    getString(n: number): string;
    private storeString(s);
    protected addRelationship(): void;
    protected addContentType(): void;
}
export = SharedStrings;
