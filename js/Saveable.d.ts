declare abstract class Saveable {
    protected path: string;
    xml: any;
    constructor(path: string);
    save(): Promise<string>;
    abstract load(): Promise<void>;
}
export = Saveable;
