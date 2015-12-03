import * as fs from "fs";
import * as xml2js from "xml2js";
import * as Promise from "bluebird";
import * as Util from "./Util";
import Workbook = require('./Workbook');
import Saveable = require('./Saveable');

var parseString: any = Promise.promisify(xml2js.parseString);
var readFile = Promise.promisify(fs.readFile);
var writeFile: any = Promise.promisify(fs.writeFile);
var builder = new xml2js.Builder();

class XMLFile extends Saveable {
    public xml: any;

    constructor(protected path: string) {
        super(path);
    }

    public load(): Promise<void> {
        return this.xml ? Promise.resolve<any>(this.xml) : Util.loadXML(this.path).then(xml => {
            this.xml = xml;
        });
    }

}

export = XMLFile;