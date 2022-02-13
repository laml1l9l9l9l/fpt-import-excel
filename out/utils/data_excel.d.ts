export declare class DataExcel {
    private nameFile;
    static dataFileOriginal: object;
    static fileOriginal: object;
    constructor(nameFile: string);
    importExcelFile(nameFile?: string): void;
    exportExcelFile(dataNew?: object, fileOriginal?: object): void;
    exportPrettyExcelFile(nameFile: string, dataNew?: object, fileOriginal?: object): void;
    exportCompareExcelFile(nameFile: string, dataNew: object | undefined, fileOriginal: object | undefined, nameCompareFile: string, dataCompare: object): void;
    pushData2DB(dataNew?: object, fileOriginal?: object): void;
}
