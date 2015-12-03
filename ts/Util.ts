import * as fs from "fs";
import * as xml2js from "xml2js";
import * as Promise from "bluebird";
import Workbook = require('./Workbook');

var parseString: any = Promise.promisify(xml2js.parseString);
var readFile = Promise.promisify(fs.readFile);
var writeFile: any = Promise.promisify(fs.writeFile);
var builder = new xml2js.Builder();

export function parseXML(input: string): Promise<{}> {
	return parseString(input);
}

export function loadXML(path: string): Promise<{}> {
	return readFile(path).then(buffer => {
		return parseString(buffer.toString());
	});
}

export function saveXML(xmlobj: {}, path: string): Promise<string> {
	var contents = builder.buildObject(xmlobj);
	return writeFile(path, contents).thenReturn(path);
}