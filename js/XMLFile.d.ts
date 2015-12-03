import * as Promise from "bluebird";
import Saveable = require('./Saveable');
declare class XMLFile extends Saveable {
    protected path: string;
    xml: any;
    constructor(path: string);
    load(): Promise<void>;
}
export = XMLFile;
