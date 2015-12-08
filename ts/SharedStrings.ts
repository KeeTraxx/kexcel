import * as fs from "fs";
import * as xml2js from "xml2js";
import * as Promise from "bluebird";

import * as Util from "./Util";

import XMLFile = require('./XMLFile');
import Workbook = require('./Workbook');
import Saveable = require('./Saveable');
import path = require('path');

interface SharedString {
    t?: Array<string>;
    r?: Array<FormattedStrings>;
}

interface FormattedStrings {
    t: any;
}

class SharedStrings extends Saveable {
    public xml:any;
    private cache:{ [s: string]: number; } = {};

    constructor(protected path:string, protected workbook:Workbook) {
        super(path);
    }

    public load():Promise<any> {
        return this.xml ?
            Promise.resolve(this.xml) :
            Util.loadXML(this.path).then(xml => {
                this.xml = xml;
            }).catch(() => {
                return Promise.all([
                    this.addRelationship(),
                    this.addContentType()
                ]).then(() => {
                    return Util.parseXML('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"><si/></sst>');
                }).then(xml => {
                    this.xml = xml;
                });
            }).finally(() => {
                this.xml.sst.si.forEach((si, index) => {
                    if (si.t) {
                        this.cache[si.t[0]] = index;
                    } else {
                        // don't cache strange strings.
                    }
                });
                return this.xml;
            });
    }

    public getIndex(s:string):number {
        return this.cache[s] || this.storeString(s);
    }

    public getString(n:number):string {
        var sxml:SharedString = this.xml.sst.si[n];
        if (!sxml) return undefined;

        return sxml.hasOwnProperty('t') ? sxml.t[0] : _.compact(sxml.r.map((d) => {
            return _.isString(d.t[0]) ? d.t[0] : null;
        })).join(' ');
    }

    private storeString(s):number {
        var index = this.xml.sst.si.push({t: [s]}) - 1;
        this.cache[s] = index;
        return index;
    }

    protected addRelationship():void {
        var relationships = this.workbook.getXML('xl/_rels/workbook.xml.rels');
        relationships.Relationships.Relationship.push({
            '$': {
                Id: 'rId1ss',
                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
                Target: 'sharedStrings.xml'
            }
        });
    }

    protected addContentType():void {
        var contentTypes = this.workbook.getXML('[Content_Types].xml');
        contentTypes.Types.Override.push({
            '$': {
                PartName: '/xl/sharedStrings.xml',
                ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
            }
        });
    }

}

export = SharedStrings;