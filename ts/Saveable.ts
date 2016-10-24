import * as Util from "./Util";
import Workbook = require('./Workbook');
import * as Promise from "bluebird";

export abstract class Saveable {

    public xml:any;

    constructor(protected path:string) {
    }

    public save():Promise<string> {
        return Util.saveXML(this.xml, this.path);
    }

    public abstract load():Promise<void>;
}
