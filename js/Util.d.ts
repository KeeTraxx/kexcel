import * as Promise from "bluebird";
export declare function parseXML(input: string): Promise<{}>;
export declare function loadXML(path: string): Promise<{}>;
export declare function saveXML(xmlobj: {}, path: string): Promise<string>;
