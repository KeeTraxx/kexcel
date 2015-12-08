import * as fs from "fs";
import * as xml2js from "xml2js";
import * as Promise from "bluebird";
import * as Util from "./Util";
import Workbook = require('./Workbook');
import Saveable = require('./Saveable');

class XMLFile extends Saveable {
    public xml:any;

    constructor(protected path:string) {
        super(path);
    }

    public load():Promise<void> {
        return this.xml ? Promise.resolve<any>(this.xml) : Util.loadXML(this.path).then(xml => {
            return this.xml = xml;
        });
    }

}

export = XMLFile;