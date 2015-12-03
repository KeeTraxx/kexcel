export interface Cell {
    $: CellAttributes;
    
    // text node for values
    v?: Array<any>;
    
    // text node for functions
    f?: Array<any>;
}

export interface CellAttributes {
    t?: string;
    s?: string;
    r: string;
}

export interface Row {
    $: RowAttributes;
    v?: Array<any>;
    f?: Array<any>;
    c?: Array<Cell>;
}

export interface RowAttributes {
    r: number;
    c?: Array<Cell>;
}