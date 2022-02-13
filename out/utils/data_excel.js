'use strict';
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.DataExcel = void 0;
const axios_1 = __importDefault(require("axios"));
const console_1 = require("console");
const fs_1 = require("fs");
const node_xlsx_1 = __importDefault(require("node-xlsx"));
const config_1 = __importDefault(require("../config"));
function getSheetExcelFile(nameFile) {
    let sheet;
    switch (nameFile) {
        case '/Tax\ ref_VN\ customers.xlsx':
            sheet = 0;
            break;
        case '/RDD1029\ Customer\ Data\ Vietnam\ LnS.xlsb':
            sheet = 1;
            break;
        default:
            sheet = 0;
            break;
    }
    return sheet;
}
function getIndexArrRowExcelFile(nameFile) {
    let iArr;
    switch (nameFile) {
        case '/Tax\ ref_VN\ customers.xlsx':
            iArr = 7;
            break;
        case '/RDD1029\ Customer\ Data\ Vietnam\ LnS.xlsb':
            iArr = 6;
            break;
        default:
            iArr = 7;
            break;
    }
    return iArr;
}
function callApiInvCloudTVan(dataNew, nameFile) {
    return __awaiter(this, void 0, void 0, function* () {
        let iArr = getIndexArrRowExcelFile(nameFile);
        let newData = JSON.parse(JSON.stringify(dataNew)), taxCode, regexTaxCode = /\d+(\-?)\d+/g, token = '';
        for (const i in newData) {
            const row = newData[i];
            taxCode = row[iArr] || null;
            if (row.length === 0 || !taxCode || !taxCode.match(regexTaxCode)) {
                continue;
            }
            taxCode = cleanTaxCode(taxCode);
            if (!taxCode) {
                continue;
            }
            let params = {
                username: config_1.default.usrInvCloudTVan,
                password: config_1.default.passInvCloudTVan
            };
            try {
                const resApi = yield axios_1.default.post(`${config_1.default.apiInvCloud}c_signin`, params);
                token = resApi.data;
            }
            catch (e) {
                (0, console_1.error)('** Error call login cloud TVan api - tax code', taxCode);
            }
            params = {
                info: {
                    taxc: taxCode
                }
            };
            try {
                yield axios_1.default.post(`${config_1.default.apiInvCloudTVan}temp-data-taxc`, params, { headers: { Authorization: "Bearer " + token } });
            }
            catch (e) {
                (0, console_1.error)('** Error call temp-data-taxc cloud TVan api - tax code', taxCode);
            }
        }
    });
}
function callApiAccuracy(iTaxCodeArr, dataExcel) {
    return __awaiter(this, void 0, void 0, function* () {
        const VALID = 'valid', INVALID = 'invalid', TYPE_TAX_CODE = 1, ADD_COLUMN_ACCURACY = 3;
        let newData = JSON.parse(JSON.stringify(dataExcel)), taxCode, regexTaxCode = /^\d+(\-?)\d+/g, regexValidStandardTaxCode = /^\d{10}(\-\d{3})?$/g;
        for (const i in newData) {
            const row = newData[i];
            taxCode = row[iTaxCodeArr] || null;
            // Re-scan file - 202110032022
            let valAccuracy = row[iTaxCodeArr + ADD_COLUMN_ACCURACY];
            (0, console_1.log)('Row', i, 'Tax code', taxCode);
            if (valAccuracy) {
                continue;
            }
            // End - 202110032022
            if (row.length === 0 || !taxCode || !taxCode.match(regexTaxCode)) {
                continue;
            }
            if (!taxCode.match(regexValidStandardTaxCode)) {
                row.push(INVALID);
                continue;
            }
            // Show tax code - 202110061728
            // console.log('Tax code accuracy', taxCode)
            // continue
            // End - 202110061728
            const paramsApiTVan = {
                auth: {
                    username: config_1.default.usrTVan,
                    password: config_1.default.passTVan
                },
                params: {
                    loaiGiayTo: TYPE_TAX_CODE,
                    soGiayTo: taxCode
                }
            };
            try {
                const jsonRes = yield axios_1.default.get(`${config_1.default.apiTVan}nnt/ttinnnt`, paramsApiTVan), infoTaxPayRes = jsonRes.data && jsonRes.data.nntDoc || {};
                if (infoTaxPayRes
                    && Object.keys(infoTaxPayRes).length !== 0
                    && Object.getPrototypeOf(infoTaxPayRes) === Object.prototype) {
                    row.push(VALID);
                    (0, console_1.log)('Call successfully accuracy api - tax code', taxCode);
                }
                else {
                    row.push(INVALID);
                    (0, console_1.error)('** Error call accuracy api - tax code', taxCode);
                }
            }
            catch (e) {
                (0, console_1.error)('** Error call accuracy api - tax code', taxCode);
                if ((e && e.response && e.response.data && e.response.data.message) === 'Không tìm thấy kết quả tìm kiếm') {
                    row.push(INVALID);
                }
            }
        }
        return newData;
    });
}
function editExcel(dataNew, nameFile, fileOriginal) {
    return __awaiter(this, void 0, void 0, function* () {
        let excelData = [], iArr, sheet, infoReNameFile = reNewNameFile(nameFile, '_accuracy_fpt'), newNameFile = infoReNameFile['newName'], extFile = infoReNameFile.extFile, dirExportFile = `${config_1.default.dirExportFile}${newNameFile}`, newDataExcelFile = JSON.parse(JSON.stringify(fileOriginal)), arrData = [], arrSpliceData = JSON.parse(JSON.stringify(dataNew)), arrPromiseAll = [], rowsPerScan = 1000;
        iArr = getIndexArrRowExcelFile(nameFile);
        sheet = getSheetExcelFile(nameFile);
        while (arrSpliceData.length) {
            arrData.push(arrSpliceData.splice(0, rowsPerScan));
        }
        arrPromiseAll = yield Promise.all(arrData.map((aArr) => __awaiter(this, void 0, void 0, function* () {
            let excelNewData = yield callApiAccuracy(iArr, aArr);
            return excelNewData;
        })));
        for (let i = 0; i < arrPromiseAll.length; i++) {
            excelData = excelData.concat(arrPromiseAll[i]);
        }
        // excelData = await callApiAccuracy(iArr, dataNew)
        newDataExcelFile[sheet].data = excelData;
        newDataExcelFile = node_xlsx_1.default.build(newDataExcelFile);
        dirExportFile = `${dirExportFile}${extFile}`;
        writeFileExcel(newDataExcelFile, dirExportFile);
    });
}
function compareExcel(dataNew, nameFile, fileOriginal, dataCompare, nameCompareFile) {
    const VALID = 'valid', INVALID = 'invalid', ADD_COLUMN_ACCURACY = 3;
    let excelData = [], iArr, sheet, iArrCompareFile, infoReNameFile = reNewNameFile(nameFile, '_compare_accuracy_fpt'), newNameFile = infoReNameFile['newName'], extFile = infoReNameFile.extFile, dirExportFile = `${config_1.default.dirExportFile}${newNameFile}`, newDataExcelFile = JSON.parse(JSON.stringify(fileOriginal)), arrData = [], arrSpliceData = JSON.parse(JSON.stringify(dataNew)), arrCompare = [], arrResultCompare = [], rowsPerScan = 1000;
    iArr = getIndexArrRowExcelFile(nameFile);
    sheet = getSheetExcelFile(nameFile);
    iArrCompareFile = getIndexArrRowExcelFile(nameCompareFile);
    while (arrSpliceData.length) {
        arrData.push(arrSpliceData.splice(0, rowsPerScan));
    }
    for (const obj of dataCompare) {
        let valAccuracy = obj[iArrCompareFile + ADD_COLUMN_ACCURACY] || null;
        const newObj = {
            taxCode: obj[iArrCompareFile],
            accuracy: valAccuracy
        };
        arrCompare.push(newObj);
    }
    arrResultCompare = arrData.map((aArr) => {
        if (!aArr && aArr.length == 0) {
            return aArr;
        }
        for (const row of aArr) {
            if (!row || !Array.isArray(row) || row.length === 0) {
                return aArr;
            }
            let taxCode = row[iArr], regexTaxCode = /^\d+(\-?)\d+/g, regexValidStandardTaxCode = /^\d{10}(\-\d{3})?$/g;
            if (!taxCode || !taxCode.match(regexTaxCode)) {
                return aArr;
            }
            if (!taxCode.match(regexValidStandardTaxCode)) {
                row.push(INVALID);
                return aArr;
            }
            let arrFilterAccuracy, objFindAccuracy, valAccuracyCompare;
            arrFilterAccuracy = arrCompare.filter((element) => element.taxCode === taxCode);
            (0, console_1.log)('Starting check', taxCode);
            if (arrFilterAccuracy && arrFilterAccuracy.length > 0) {
                objFindAccuracy = arrFilterAccuracy.find((element = {}) => (element.accuracy === VALID || element.accuracy === INVALID));
                valAccuracyCompare = objFindAccuracy && objFindAccuracy.accuracy || null;
                if (valAccuracyCompare) {
                    row.push(valAccuracyCompare);
                }
            }
        }
        return aArr;
    });
    for (let i = 0; i < arrResultCompare.length; i++) {
        excelData = excelData.concat(arrResultCompare[i]);
    }
    newDataExcelFile[sheet].data = excelData;
    newDataExcelFile = node_xlsx_1.default.build(newDataExcelFile);
    dirExportFile = `${dirExportFile}${extFile}`;
    writeFileExcel(newDataExcelFile, dirExportFile);
}
function prettyFileExcel(dataNew, nameFile, fileOriginal) {
    const VALID = 'valid', INVALID = 'invalid', ADD_COLUMN_ACCURACY = 3;
    let excelData = [], iArr, sheet, infoReNameFile = reNewNameFile(nameFile, '_compare_accuracy_fpt'), newNameFile = infoReNameFile['newName'], extFile = infoReNameFile.extFile, dirExportFile = `${config_1.default.dirExportFile}${newNameFile}`, newDataExcelFile = JSON.parse(JSON.stringify(fileOriginal)), arrData = [], arrSpliceData = JSON.parse(JSON.stringify(dataNew)), arrCompare = [], arrResultCompare = [], rowsPerScan = 1000;
    iArr = getIndexArrRowExcelFile(nameFile);
    sheet = getSheetExcelFile(nameFile);
    while (arrSpliceData.length) {
        arrData.push(arrSpliceData.splice(0, rowsPerScan));
    }
    for (const obj of arrSpliceData) {
        let valAccuracy = obj[iArr + ADD_COLUMN_ACCURACY] || null;
        const newObj = {
            taxCode: obj[iArr],
            accuracy: valAccuracy
        };
        arrCompare.push(newObj);
    }
    arrResultCompare = arrData.map((aArr) => {
        if (!aArr && aArr.length == 0) {
            return aArr;
        }
        for (const row of aArr) {
            if (!row || !Array.isArray(row) || row.length === 0) {
                return aArr;
            }
            let taxCode = row[iArr], regexTaxCode = /^\d+(\-?)\d+/g, regexValidStandardTaxCode = /^\d{10}(\-\d{3})?$/g;
            if (!taxCode || !taxCode.match(regexTaxCode)) {
                return aArr;
            }
            if (!taxCode.match(regexValidStandardTaxCode)) {
                row.push(INVALID);
                return aArr;
            }
            let arrFilterAccuracy, objFindAccuracy, valAccuracyCompare;
            arrFilterAccuracy = arrCompare.filter((element) => element.taxCode === taxCode);
            (0, console_1.log)('Starting check', taxCode);
            if (arrFilterAccuracy && arrFilterAccuracy.length > 0) {
                objFindAccuracy = arrFilterAccuracy.find((element = {}) => (element.accuracy === VALID || element.accuracy === INVALID));
                valAccuracyCompare = objFindAccuracy && objFindAccuracy.accuracy || null;
                if (valAccuracyCompare) {
                    row.push(valAccuracyCompare);
                }
            }
        }
        return aArr;
    });
    for (let i = 0; i < arrResultCompare.length; i++) {
        excelData = excelData.concat(arrResultCompare[i]);
    }
    newDataExcelFile[sheet].data = excelData;
    newDataExcelFile = node_xlsx_1.default.build(newDataExcelFile);
    dirExportFile = `${dirExportFile}${extFile}`;
    writeFileExcel(newDataExcelFile, dirExportFile);
}
function reNewNameFile(oldName, addStr) {
    const splitOldName = oldName.split('.'), maxIndex = splitOldName.length - 1;
    let extFile = `.${splitOldName[maxIndex]}`, newName = `${oldName.replace(extFile, '')}${addStr}`, objRes = {
        newName,
        extFile
    };
    return objRes;
}
function addCharacterStr(start, delCount, newSubStr, oldStr) {
    return oldStr.slice(0, start) + newSubStr + oldStr.slice(start + Math.abs(delCount));
}
;
function cleanTaxCode(taxCode) {
    let newTaxCode = taxCode, regexValidStandardTaxCode = /^\d{10}(\-\d{3})?$/g;
    const TAX_CODE_FAILD = '';
    if (taxCode.match(regexValidStandardTaxCode)) {
        return newTaxCode;
    }
    if (taxCode.length < 10 || taxCode.length > 14) {
        newTaxCode = TAX_CODE_FAILD;
    }
    else {
        let regexDashes = /\-/g;
        let arrMatch = taxCode.match(regexDashes) || [];
        if (taxCode.length === 10 && arrMatch.length > 0) {
            newTaxCode = TAX_CODE_FAILD;
        }
        else if (taxCode.length === 14) {
            if (arrMatch.length !== 1) {
                newTaxCode = TAX_CODE_FAILD;
            }
        }
        else {
            if (arrMatch.length !== 1) {
                if (arrMatch.length === 0 && taxCode.length === 13) {
                    newTaxCode = addCharacterStr(10, 0, '-', taxCode);
                }
                else {
                    newTaxCode = TAX_CODE_FAILD;
                }
            }
            else {
                let arrSplitTC = taxCode.split('-'), maxLen = arrSplitTC.length - 1;
                arrSplitTC[maxLen] = arrSplitTC[maxLen].padStart(3, '0');
                newTaxCode = arrSplitTC.join('-');
            }
        }
    }
    return newTaxCode;
}
function writeFileExcel(excelData, newNameFile) {
    try {
        (0, fs_1.writeFileSync)(newNameFile, excelData);
    }
    catch (err) {
        (0, console_1.error)('** Write failed excel file', err);
    }
}
class DataExcel {
    constructor(nameFile) {
        this.nameFile = nameFile;
    }
    importExcelFile(nameFile = this.nameFile) {
        try {
            let excelFile, excelData, objData, sheet_name_list, sheet, dirImportFile = `${config_1.default.dirImportFile}${nameFile}`;
            // ---- node-xlsx - 202110041006 ----
            excelFile = (0, fs_1.readFileSync)(dirImportFile);
            excelData = node_xlsx_1.default.parse(excelFile);
            sheet = getSheetExcelFile(nameFile);
            // ---- node-xlsx - 202110041006 ----
            // ---- xlsx - 202110041031 ----
            // excelFile = _xlsx.readFile(dirImportFile)
            // sheet = getSheetExcelFile(nameFile)
            // sheet_name_list = [excelFile.SheetNames[sheet]]
            // // excelData =_xlsx.utils.sheet_to_json(excelFile.Sheets[sheet_name_list[sheet]])
            // sheet_name_list.forEach(function(y: any) {
            //   let worksheet = excelFile.Sheets[y]
            //   , headers: any = {}
            //   , data: any = []
            //   , z: any
            //   for(z in worksheet) {
            //     if(z[0] === '!') continue;
            //     //parse out the column, row, and value
            //     var tt = 0;
            //     for (var i = 0; i < z.length; i++) {
            //         if (!isNaN(z[i])) {
            //             tt = i;
            //             break;
            //         }
            //     };
            //     var col = z.substring(0,tt);
            //     var row = parseInt(z.substring(tt));
            //     var value = worksheet[z].v;
            //     //store header names
            //     if(row == 1 && value) {
            //         headers[col] = value;
            //         continue;
            //     }
            //     if(!data[row]) data[row]={};
            //     data[row][headers[col]] = value;
            //   }
            //   //drop those first two rows which are empty
            //   // data.shift();
            //   // data.shift();
            //   // console.log(data);
            //   if (data && data.length > 0) { excelData = data }
            // })
            // ---- xlsx - 202110041031 ----
            DataExcel.fileOriginal = excelData;
            objData = excelData[sheet].data;
            DataExcel.dataFileOriginal = objData;
            (0, console_1.log)('Import successfully excel file');
        }
        catch (err) {
            (0, console_1.error)('** Import failed excel file', err);
        }
    }
    exportExcelFile(dataNew = DataExcel.dataFileOriginal, fileOriginal = DataExcel.fileOriginal) {
        editExcel(dataNew, this.nameFile, fileOriginal)
            .then(() => {
            (0, console_1.log)('Export successfully excel file');
        })
            .catch((err) => {
            (0, console_1.error)('** Export failed excel file', err);
        });
    }
    exportPrettyExcelFile(nameFile, dataNew = DataExcel.dataFileOriginal, fileOriginal = DataExcel.fileOriginal) {
        prettyFileExcel(dataNew, nameFile, fileOriginal);
        (0, console_1.log)('Export successfully excel file');
    }
    exportCompareExcelFile(nameFile, dataNew = DataExcel.dataFileOriginal, fileOriginal = DataExcel.fileOriginal, nameCompareFile, dataCompare) {
        compareExcel(dataNew, nameFile, fileOriginal, dataCompare, nameCompareFile);
        (0, console_1.log)('Export successfully excel file');
    }
    pushData2DB(dataNew = DataExcel.dataFileOriginal, fileOriginal = DataExcel.fileOriginal) {
        callApiInvCloudTVan(dataNew, this.nameFile)
            .then(() => {
            (0, console_1.log)('Push data successfully to database');
        })
            .catch((err) => {
            (0, console_1.error)('** Push data failed to database', err);
        });
    }
}
exports.DataExcel = DataExcel;
DataExcel.dataFileOriginal = {};
DataExcel.fileOriginal = {};
//# sourceMappingURL=data_excel.js.map