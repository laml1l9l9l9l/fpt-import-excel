"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const data_excel_1 = require("./utils/data_excel");
console.log('Started Code');
let nameFileRDD1029, nameFileTaxRef;
function accuracyExcelFile(nameFile) {
    const excel = new data_excel_1.DataExcel(nameFile);
    excel.importExcelFile();
    excel.exportExcelFile();
}
function accuracyCompareExcelFiles(nameFile, nameFileCompare) {
    const excel = new data_excel_1.DataExcel(nameFile), excelCompare = new data_excel_1.DataExcel(nameFileCompare);
    let dataFileCompare, dataFileAccuracy, fileOriginalAccuracy;
    excel.importExcelFile(nameFile);
    dataFileAccuracy = JSON.parse(JSON.stringify(data_excel_1.DataExcel.dataFileOriginal));
    fileOriginalAccuracy = JSON.parse(JSON.stringify(data_excel_1.DataExcel.fileOriginal));
    excelCompare.importExcelFile(nameFileCompare);
    dataFileCompare = JSON.parse(JSON.stringify(data_excel_1.DataExcel.dataFileOriginal));
    excel.exportCompareExcelFile(nameFile, dataFileAccuracy, fileOriginalAccuracy, nameFileCompare, dataFileCompare);
}
function prettyExcelFiles(nameFile) {
    const excel = new data_excel_1.DataExcel(nameFile);
    let dataFileCompare, dataFileAccuracy, fileOriginalAccuracy;
    excel.importExcelFile(nameFile);
    dataFileAccuracy = JSON.parse(JSON.stringify(data_excel_1.DataExcel.dataFileOriginal));
    fileOriginalAccuracy = JSON.parse(JSON.stringify(data_excel_1.DataExcel.fileOriginal));
    dataFileCompare = JSON.parse(JSON.stringify(data_excel_1.DataExcel.dataFileOriginal));
    excel.exportPrettyExcelFile(nameFile, dataFileAccuracy, fileOriginalAccuracy);
}
function pushDBCloud(nameFile) {
    const excel = new data_excel_1.DataExcel(nameFile);
    excel.importExcelFile();
    excel.pushData2DB();
}
try {
    nameFileTaxRef = '/Tax\ ref_VN\ customers.xlsx';
    nameFileRDD1029 = `/RDD1029\ Customer\ Data\ Vietnam\ LnS.xlsb`;
    // ----- Accuracy excel -----
    // accuracyExcelFile(nameFileTaxRef)
    // accuracyExcelFile(nameFileRDD1029) // | Don't use
    // accuracyCompareExcelFiles(nameFileRDD1029, nameFileTaxRef)
    // accuracyCompareExcelFiles(nameFileRDD1029, nameFileTaxRef)
    prettyExcelFiles(nameFileTaxRef);
    // ----- Push DB Cloud -----
    // pushDBCloud(nameFileTaxRef)
}
catch (e) {
    console.error('Err read file', e);
}
//# sourceMappingURL=index.js.map