import { Cell, Column, Row, Workbook, Worksheet } from "exceljs";
import { existsSync, mkdirSync, statSync, writeFileSync } from "fs";
import path from "path";

export interface HeaderData {
    key: string;
    letter: string;
    number: number;
    rowNumber: number;
}

export interface Filter {
    header: string;
    useColumnLetter?: boolean;
}

export type FilterMode = "excludes" | "includes";

export interface FilterOption {
    mode: FilterMode;
    filters: Filter[];
}

export class FilterHandler {
    private mode: FilterMode;
    private filterList: Filter[];

    constructor(filterOption?: FilterOption) {
        this.mode = filterOption?.mode ?? "includes";
        this.filterList = filterOption?.filters ?? [];
    }
    private isIncludeMode(): boolean {
        return this.mode === "includes";
    }
    private matches(column: HeaderData): boolean {
        for(var i = 0; i < this.filterList.length; i++){
            let filter: Filter = this.filterList[i];
            let header: string = filter.header.toLowerCase();
            
            if(filter.useColumnLetter && header === column.letter.toLowerCase()) {
                return this.isIncludeMode();
            }
            if(!filter.useColumnLetter && header === column.key.toLowerCase()){
                return this.isIncludeMode();
            }
        }
        return !this.isIncludeMode();
    }
    private isEmpty(): boolean {
        return this.filterList.length < 1;
    }
    public allowHeader(header: HeaderData): boolean {
        if(!this.isEmpty()) {
            return this.matches(header);
        }
        return true;
    }
    public getMode(): FilterMode {
        return this.mode;
    }
}

export interface RowEventForEachCallback {
    (key: string, value: any): void;
}

export interface RowEvent {
    getValueAsString(key: string, defaultValue?: string): string | null;
    getValueAsNumber(key: string, defaultValue?: number): number | null;
    getValueAsDate(key: string, defaultValue?: Date): Date | null;
    getValueAsBoolean(key: string, defaultValue?: boolean): boolean | null;
    hasKey(key: string): boolean;
    getRawValue(key: string): any;
    getKeys(): string[];
    foreach(callback: RowEventForEachCallback): void;
    getAll(): Record<string, any>;
}

export class RowEventHandler implements RowEvent {
    private readonly data: Record<string, any>;

    constructor(row: Record<string, any>) {
        this.data = row;
    }
    public getValueAsString(key: string, defaultValue?: string): string | null {
        return typeof this.data[key] === "string" ? (
            this.data[key]
        ): defaultValue ?? null;
    }
    public getValueAsNumber(key: string, defaultValue?: number): number | null {
        return typeof this.data[key] === "number" ? (
            this.data[key]
        ): defaultValue ?? null;
    }
    public getValueAsDate(key: string, defaultValue?: Date): Date | null {
        return this.data[key] instanceof Date ? (
            this.data[key]
        ): defaultValue ?? null;
    }
    public getValueAsBoolean(key: string, defaultValue?: boolean): boolean | null {
        return typeof this.data[key] === "boolean" ? (
            this.data[key]
        ): defaultValue ?? null;
    }
    public hasKey(key: string): boolean {
        return key in this.data;
    }
    public getRawValue(key: string): any {
        return this.hasKey(key) ? this.data[key] : null;
    }
    public getKeys(): string[] {
        return Object.keys(this.data);
    }
    public foreach(callback: RowEventForEachCallback): void {
        for(const key of Object.keys(this.data)) {
            callback(key, this.data[key]);
        }
    }
    public getAll(): Record<string, any> {
        return this.data;
    }
}

export interface ForEachController {
    abort(): void;
    currentRowIndex(): number;
    totalRows(): number;
}

export interface ForEachRowCallback {
    (row: RowEvent, controller: ForEachController): void;
}

export interface ExcelParserOptions {
    useFirstRowAsHeader?: boolean;
    filterOption?: FilterOption;
    maxKeyLength?: number;
}

export interface ExcelParserConfig extends ExcelParserOptions {
    worksheet: Worksheet;
}

export interface ExcelParserReadFileConfig extends ExcelParserOptions {
    excelFilePath: string;
    worksheet: string;
}

export class ExcelParser {
    private worksheet: Worksheet;
    private headers: Map<number, HeaderData>;
    private keySet: Set<string>;
    private useFirstRowAsHeader: boolean;
    private filters: FilterHandler;
    private excludedHeaders: Set<number>;
    private maxKeyLength: number;

    constructor(config: ExcelParserConfig) {
        this.worksheet = config.worksheet;
        this.headers = new Map();
        this.keySet = new Set();
        this.useFirstRowAsHeader = config.useFirstRowAsHeader ?? true;
        this.filters = new FilterHandler(config.filterOption);
        this.excludedHeaders = new Set();
        this.maxKeyLength = config.maxKeyLength ?? 128;
    }
    private parseCellValue(cell: Cell): any {
        let value: any = cell.result ?? cell.value;
        if(value && typeof value === "object") {
            if('result' in value && 'formula' in value) {
                return value['result'];
            }
        }
        return value;
    }
    private normalizeKey(value: string): string {
        return value
            .trim()
            .replace(/\s+/g, ' ')
            .replace(/[\r\n\t]/g, '')
            .replace(/[\x00-\x1F\x7F]/g, '')
    }
    private parseKeyAddress(value: any): string {
        if(value === null) {
            return null;
        }
        switch(typeof value) {
            case "object":
            case "function":
            case "undefined":
                return null;
            default:
                let result: string = this.normalizeKey(`${value}`);
                return result.length > 0 && result.length <= this.maxKeyLength ? (
                    result
                ): null;
        }
    }
    private parseRowRecord(row: Row): RowEvent {
        let result: Record<string, any> = {};
        let isHeading: boolean = this.headers.size === 0;
        let isIsolatedHeading: boolean = isHeading && this.useFirstRowAsHeader;
        let rowNumber: number = row.number;

        row.eachCell((cell) => {
            let column: Partial<Column> = this.worksheet.getColumn(cell.col);
            let columnNumber: number = column.number;

            if(!this.excludedHeaders.has(columnNumber)) {
                let columnLetter: string = column.letter;
                let value: any = this.parseCellValue(cell);

                let header: HeaderData = null;
                let address: string = `${columnLetter}${rowNumber}`;

                if(this.headers.has(columnNumber)) {
                    header = this.headers.get(columnNumber);
                } else {
                    let keyAddress: string = address;

                    if(this.useFirstRowAsHeader){
                        keyAddress = this.parseKeyAddress(value) ?? keyAddress;
                    }

                    if(this.keySet.has(keyAddress)) {
                        keyAddress = address;
                    } else {
                        this.keySet.add(keyAddress);
                    }

                    header = {
                        key: keyAddress,
                        letter: columnLetter,
                        number: columnNumber,
                        rowNumber
                    };
                    if(this.filters.allowHeader(header)){
                        this.headers.set(columnNumber, header);
                    } else {
                        this.excludedHeaders.add(columnNumber);
                        header = null;
                    }
                }
                if(header && !isIsolatedHeading) {
                    result[header.key] = value;
                }
            }
        });
        if(Object.keys(result).length > 0 && !isIsolatedHeading){
            return new RowEventHandler(result);
        }
        return null;
    }
    public foreach(callback: ForEachRowCallback, limit?: number): void {
        let rowCount: number = this.worksheet.rowCount;
        let counter: number = 0;

        for(var i = 1; i <= rowCount; i++) {
            if(limit && counter >= limit) {
                break;
            }
            let row: Row = this.worksheet.getRow(i);
            let isAborted: boolean = false;
            
            if(row.hasValues) {
                let event: RowEvent = this.parseRowRecord(row);
                if(event !== null) {
                    callback(event, {
                        abort: function(): void {
                            isAborted = true;
                        },
                        currentRowIndex: function(): number {
                            return row.number;
                        },
                        totalRows: function(): number {
                            return rowCount;
                        }
                    });
                    counter += 1;
                }
            }
            if(isAborted){
                break;
            }
        }
    }
    public getHeaders(): Map<number, HeaderData> {
        return this.headers;
    }
    public toJson(): Record<string, any>[] {
        let result: Record<string, any>[] = [];
        this.foreach(function(row) {
            result.push(row.getAll());
        })
        return result;
    }
    public setFilter(filters: Filter[], mode?: FilterMode): void {
        this.filters = new FilterHandler({
            filters, mode: mode ?? this.filters.getMode()
        });
        this.headers = new Map();
        this.keySet = new Set();
        this.excludedHeaders = new Set();
    }
    public getName(): string {
        return this.worksheet.name;
    }
    public writeJsonFile(filepath: string, format: boolean = true): void {
        let directory: string = path.dirname(filepath);
        let extention: string = path.extname(filepath).toLowerCase();

        if(extention.trim() !== ".json"){
            throw new Error(`Invalid json filepath: '${filepath}'`);
        }
        if(!existsSync(directory)){
            mkdirSync(directory, {
                recursive: true
            });
        }

        let jsonData: Record<string, any>[] = this.toJson();

        writeFileSync(
            filepath,
            format ? JSON.stringify(jsonData, null, 4) : JSON.stringify(jsonData),
            "utf-8"
        );
    }
    public static async readFile(config: ExcelParserReadFileConfig): Promise<ExcelParser> {
        const workbook: Workbook = new Workbook();
        await workbook.xlsx.readFile(config.excelFilePath);

        const worksheet: Worksheet = workbook.getWorksheet(config.worksheet);

        if(!worksheet) {
            throw new Error(`Worksheet '${config.worksheet}' not found in '${config.excelFilePath}'`);
        }
        return new ExcelParser({ ...config, worksheet });
    }
}

export interface WorksheetNameFilterOptions {
    mode: FilterMode;
    filters: string[];
}

export interface WriteJsonConfig {
    filepath: string;
    seperateWorksheetFiles?: boolean;
    format?: boolean;
}

export interface WorksheetData {
    name: string;
    headers: HeaderData[];
    rows: Record<string, any>[];
}

export interface ForeachWorksheetCallback {
    (worksheet: ExcelParser): void;
}

export interface ExcelDocumentConfig {
    excelFilePath: string;
    worksheetOptions?: ExcelParserOptions;
    throwErrors?: boolean;
    filterOptions?: WorksheetNameFilterOptions;
}

export default class ExcelDocument {
    private filepath: string;
    private worksheetOptions: ExcelParserOptions;
    private workbook: Workbook;
    private allowError: boolean;
    private filterOptions: WorksheetNameFilterOptions;

    constructor(config: ExcelDocumentConfig){
        this.filepath = config.excelFilePath;
        this.worksheetOptions = config.worksheetOptions ?? {};
        this.workbook = null;
        this.allowError = config.throwErrors ?? true;
        this.filterOptions = config.filterOptions ?? null;
    }
    private allowWorksheet(name: string): boolean {
        let filters: string[] = this.filterOptions?.filters ?? [];
        if(filters.length > 0){
            let mode: FilterMode = this.filterOptions?.mode ?? "includes";
            if(mode === 'includes') {
                return filters.includes(name);
            }
            if(mode === "excludes") {
                return !filters.includes(name);
            }
        }
        return true;
    }
    private async getWorkbook(): Promise<Workbook> {
        if(this.workbook === null) {
            try {
                this.workbook = new Workbook();
                await this.workbook.xlsx.readFile(this.filepath);   
            } catch (error) {
                if(this.allowError){
                    throw error;
                }
                return null;
            }
        }
        return this.workbook;
    }
    private resolveSepareteFilePath(filepath: string, worksheet: string): string {
        let extention: string = path.extname(filepath);
        let firstFilepath: string = filepath.substring(0, filepath.lastIndexOf(extention));
        let newFilePath: string = `${firstFilepath}_${worksheet}${extention}`;
        return path.resolve(newFilePath);
    }
    public async findWorksheet(name: string, options?: ExcelParserOptions): Promise<ExcelParser> {
        let workbook: Workbook = await this.getWorkbook();

        if(workbook !== null){
            let worksheet: Worksheet = workbook.getWorksheet(name);
            if(worksheet && this.allowWorksheet(worksheet.name)) {
                return new ExcelParser({
                    worksheet, ...(options ?? this.worksheetOptions)
                });
            }
        }
        return null;
    }
    public async foreach(callback: ForeachWorksheetCallback): Promise<void> {
        let workbook: Workbook = await this.getWorkbook();
        workbook.eachSheet((worksheet) => {
            if(this.allowWorksheet(worksheet.name)){
                callback(new ExcelParser({
                    worksheet, ...this.worksheetOptions
                }));
            }
        });
    }
    public async parseAsJson(): Promise<WorksheetData[]> {
        const result: WorksheetData[] = [];

        await this.foreach(function(parser) {
            let rows: Record<string, any>[] = parser.toJson();
            result.push({
                name: parser.getName(),
                headers: Array.from(parser.getHeaders().values()),
                rows
            });
        })
        return result;
    }
    public async writeJsonFile(config: WriteJsonConfig): Promise<void> {
        if(config.seperateWorksheetFiles) {
            this.foreach((parser) => {
                try {
                    parser.writeJsonFile(
                        this.resolveSepareteFilePath(config.filepath, parser.getName()),
                        config.format
                    );   
                } catch (error) {
                    if(this.allowError){
                        return error;
                    }
                    return;
                }
            })
        } else {
            let directory: string = path.dirname(config.filepath);
            let extention: string = path.extname(config.filepath).toLowerCase();

            if(extention.trim() !== ".json") {
                if(this.allowError){
                    throw new Error(`Invalid json filepath: '${config.filepath}'`);
                } else {
                    return;
                }
            }
            if(!existsSync(directory)){
                mkdirSync(directory, {
                    recursive: true
                });
            }

            let jsonData: Record<string, any>[] = await this.parseAsJson();
            let format: boolean = config.format ?? true;

            writeFileSync(
                config.filepath,
                format ? JSON.stringify(jsonData, null, 4) : JSON.stringify(jsonData),
                "utf-8"
            );
        }
    }
}


async function run(): Promise<void> {
    let excel: ExcelDocument = new ExcelDocument({
        excelFilePath: path.resolve(__dirname, "../assets/dummy.xlsx"),
        // filterOptions: {
        //     mode: 'excludes',
        //     filters: ["Sheet1"]
        // }
    });
    // writeFileSync(
    //     path.resolve(__dirname, "../output/data.json"),
    //     JSON.stringify(await excel.parseAsJson(), null, 4),
    //     "utf-8"
    // );
    excel.writeJsonFile({
        filepath: path.resolve(__dirname, "../output/data.json")
    })
}

run();

// async function run(): Promise<void> {
//     let excelWorksheet: ExcelParser = await ExcelParser.readFile({
//         excelFilePath: path.resolve(__dirname, "../assets/dummy.xlsx"),
//         worksheet: "Sheet1",
//         // filterOption: {
//         //     mode: "excludes",
//         //     filters: [
//         //         {
//         //             header: "Department"
//         //         }
//         //     ]
//         // }
//     });
//     // excelWorksheet.writeJsonFile(
//     //     path.resolve(__dirname, "../output/data.json")
//     // );

//     // excelWorksheet.foreach(function(row, controller) {
//     //     let department: string = row.getValueAsString("Department", "");

//     //     console.log(row.getAll());

//     //     if(department.toUpperCase() === "HR") {
//     //         controller.abort();
//     //     }
//     // });
// }

// run();