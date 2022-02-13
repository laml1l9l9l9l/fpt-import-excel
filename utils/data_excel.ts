'use strict'
import axios from "axios"
import { log, error } from "console"
import { readFileSync, writeFileSync } from "fs"
import xlsx from "node-xlsx"
import _xlsx from "xlsx"
import config from "../config"

function getSheetExcelFile(nameFile: string): number {
  let sheet: number
  switch (nameFile) {
    case '/Tax\ ref_VN\ customers.xlsx':
      sheet = 0
      break
    case '/RDD1029\ Customer\ Data\ Vietnam\ LnS.xlsb':
      sheet = 1
      break
    default:
      sheet = 0
      break
  }
  return sheet
}

function getIndexArrRowExcelFile(nameFile: string): number {
  let iArr: number
  switch (nameFile) {
    case '/Tax\ ref_VN\ customers.xlsx':
      iArr = 7
      break
    case '/RDD1029\ Customer\ Data\ Vietnam\ LnS.xlsb':
      iArr = 6
      break
    default:
      iArr = 7
      break
  }
  return iArr
}

async function callApiInvCloudTVan (dataNew: object, nameFile: string): Promise<void> {
  let iArr = getIndexArrRowExcelFile(nameFile)
  let newData: { [index: string]: any } = JSON.parse(JSON.stringify(dataNew))
    , taxCode
    , regexTaxCode = /\d+(\-?)\d+/g
    , token: string =  ''
  for(const i in newData) {
    const row = newData[i]
    taxCode = row[iArr] || null
    if(row.length === 0 || !taxCode || !taxCode.match(regexTaxCode)) { continue }
    taxCode = cleanTaxCode(taxCode)
    if(!taxCode) {continue}
    let params: object = {
      username: config.usrInvCloudTVan,
      password: config.passInvCloudTVan
    }
    try {
      const resApi = await axios.post(`${config.apiInvCloud}c_signin`, params)
      token = resApi.data
    } catch (e: any) {
      error('** Error call login cloud TVan api - tax code', taxCode)
    }
    params = {
      info: {
        taxc: taxCode
      }
    }
    try {
      await axios.post(`${config.apiInvCloudTVan}temp-data-taxc`, params, { headers: { Authorization: "Bearer " + token } })
    } catch (e: any) {
      error('** Error call temp-data-taxc cloud TVan api - tax code', taxCode)
    }
  }
}

async function callApiAccuracy(iTaxCodeArr: number, dataExcel: object): Promise<object> {
  const VALID = 'valid'
    , INVALID = 'invalid'
    , TYPE_TAX_CODE = 1
    , ADD_COLUMN_ACCURACY = 3
  let newData: { [index: string]: any } = JSON.parse(JSON.stringify(dataExcel))
    , taxCode
    , regexTaxCode = /^\d+(\-?)\d+/g
    , regexValidStandardTaxCode = /^\d{10}(\-\d{3})?$/g
  for(const i in newData) {
    const row = newData[i]
    taxCode = row[iTaxCodeArr] || null
    // Re-scan file - 202110032022
    let valAccuracy = row[iTaxCodeArr + ADD_COLUMN_ACCURACY]
    log('Row', i, 'Tax code', taxCode)
    if (valAccuracy) {
      continue
    }
    // End - 202110032022
    if(row.length === 0 || !taxCode || !taxCode.match(regexTaxCode)) { continue }
    if(!taxCode.match(regexValidStandardTaxCode)) {
      row.push(INVALID)
      continue
    }
    // Show tax code - 202110061728
    // console.log('Tax code accuracy', taxCode)
    // continue
    // End - 202110061728
    const paramsApiTVan = {
      auth: {
        username: config.usrTVan,
        password: config.passTVan
      },
      params: {
        loaiGiayTo: TYPE_TAX_CODE,
        soGiayTo: taxCode
      }
    }
    try {
      const jsonRes = await axios.get(`${config.apiTVan}nnt/ttinnnt`, paramsApiTVan),
        infoTaxPayRes = jsonRes.data && jsonRes.data.nntDoc || {}
        if (
          infoTaxPayRes
          && Object.keys(infoTaxPayRes).length !== 0
          && Object.getPrototypeOf(infoTaxPayRes) === Object.prototype
        ) {
          row.push(VALID)
          log('Call successfully accuracy api - tax code', taxCode)
        } else {
          row.push(INVALID)
          error('** Error call accuracy api - tax code', taxCode)
        }
    } catch (e: any) {
      error('** Error call accuracy api - tax code', taxCode)
      if ((e && e.response && e.response.data && e.response.data.message) === 'Không tìm thấy kết quả tìm kiếm') {
        row.push(INVALID)
      }
    }
  }
  return newData
}

async function editExcel(dataNew: object, nameFile: string, fileOriginal: object): Promise<void> {
  let excelData: any = []
    ,iArr: number
    ,sheet: number
    ,infoReNameFile: any = reNewNameFile(nameFile, '_accuracy_fpt')
    ,newNameFile = infoReNameFile['newName']
    ,extFile = infoReNameFile.extFile
    ,dirExportFile = `${config.dirExportFile}${newNameFile}`
    ,newDataExcelFile: any = JSON.parse(JSON.stringify(fileOriginal))
    ,arrData: any = []
    ,arrSpliceData: any = JSON.parse(JSON.stringify(dataNew))
    ,arrPromiseAll: any = []
    ,rowsPerScan: number = 1000
  iArr = getIndexArrRowExcelFile(nameFile)
  sheet = getSheetExcelFile(nameFile)
  while (arrSpliceData.length) {
    arrData.push(arrSpliceData.splice(0, rowsPerScan))
  }
  arrPromiseAll = await Promise.all(arrData.map(async (aArr: any) => {
    let excelNewData = await callApiAccuracy(iArr, aArr)
    return excelNewData
  }))
  for (let i = 0; i < arrPromiseAll.length; i++) {
    excelData = excelData.concat(arrPromiseAll[i])
  }
  // excelData = await callApiAccuracy(iArr, dataNew)
  newDataExcelFile[sheet].data = excelData
  newDataExcelFile = xlsx.build(newDataExcelFile)
  dirExportFile = `${dirExportFile}${extFile}`
  writeFileExcel(newDataExcelFile, dirExportFile)
}


function compareExcel(dataNew: object, nameFile: string, fileOriginal: object, dataCompare: any, nameCompareFile: string): void {
  const VALID = 'valid'
    , INVALID = 'invalid'
    , ADD_COLUMN_ACCURACY = 3
  let excelData: any = []
    , iArr: number
    , sheet: number
    , iArrCompareFile: number
    , infoReNameFile: any = reNewNameFile(nameFile, '_compare_accuracy_fpt')
    , newNameFile = infoReNameFile['newName']
    , extFile = infoReNameFile.extFile
    , dirExportFile = `${config.dirExportFile}${newNameFile}`
    , newDataExcelFile: any = JSON.parse(JSON.stringify(fileOriginal))
    , arrData: any = []
    , arrSpliceData: any = JSON.parse(JSON.stringify(dataNew))
    , arrCompare: any = []
    , arrResultCompare: any = []
    , rowsPerScan: number = 1000
  iArr = getIndexArrRowExcelFile(nameFile)
  sheet = getSheetExcelFile(nameFile)
  iArrCompareFile = getIndexArrRowExcelFile(nameCompareFile)
  while (arrSpliceData.length) {
    arrData.push(arrSpliceData.splice(0, rowsPerScan))
  }
  for (const obj of dataCompare) {
    let valAccuracy = obj[iArrCompareFile + ADD_COLUMN_ACCURACY] || null
    const newObj = {
      taxCode: obj[iArrCompareFile],
      accuracy: valAccuracy
    }
    arrCompare.push(newObj)
  }
  arrResultCompare = arrData.map((aArr: any) => {
    if(!aArr && aArr.length == 0) { return aArr }
    for (const row of aArr) {
      if( !row || !Array.isArray(row) || row.length === 0) { return aArr }
      let taxCode = row[iArr]
        , regexTaxCode = /^\d+(\-?)\d+/g
        , regexValidStandardTaxCode = /^\d{10}(\-\d{3})?$/g
      if(!taxCode || !taxCode.match(regexTaxCode)) {
        return aArr
      }
      if(!taxCode.match(regexValidStandardTaxCode)) {
        row.push(INVALID)
        return aArr
      }
      let arrFilterAccuracy
        , objFindAccuracy
        , valAccuracyCompare
        arrFilterAccuracy = arrCompare.filter((element: any) => element.taxCode === taxCode)
      log('Starting check', taxCode)
      if (arrFilterAccuracy && arrFilterAccuracy.length > 0) {
        objFindAccuracy = arrFilterAccuracy.find((element: any = {}) => (element.accuracy === VALID || element.accuracy === INVALID))
        valAccuracyCompare = objFindAccuracy && objFindAccuracy.accuracy || null
        if (valAccuracyCompare) { row.push(valAccuracyCompare) }
      }
    }
    return aArr
  })
  for (let i = 0; i < arrResultCompare.length; i++) {
    excelData = excelData.concat(arrResultCompare[i])
  }
  newDataExcelFile[sheet].data = excelData
  newDataExcelFile = xlsx.build(newDataExcelFile)
  dirExportFile = `${dirExportFile}${extFile}`
  writeFileExcel(newDataExcelFile, dirExportFile)
}

function prettyFileExcel(dataNew: object, nameFile: string, fileOriginal: object): void {
  const VALID = 'valid'
    , INVALID = 'invalid'
    , ADD_COLUMN_ACCURACY = 3
  let excelData: any = []
    , iArr: number
    , sheet: number
    , infoReNameFile: any = reNewNameFile(nameFile, '_compare_accuracy_fpt')
    , newNameFile = infoReNameFile['newName']
    , extFile = infoReNameFile.extFile
    , dirExportFile = `${config.dirExportFile}${newNameFile}`
    , newDataExcelFile: any = JSON.parse(JSON.stringify(fileOriginal))
    , arrData: any = []
    , arrSpliceData: any = JSON.parse(JSON.stringify(dataNew))
    , arrCompare: any = []
    , arrResultCompare: any = []
    , rowsPerScan: number = 1000
  iArr = getIndexArrRowExcelFile(nameFile)
  sheet = getSheetExcelFile(nameFile)
  while (arrSpliceData.length) {
    arrData.push(arrSpliceData.splice(0, rowsPerScan))
  }
  for (const obj of arrSpliceData) {
    let valAccuracy = obj[iArr + ADD_COLUMN_ACCURACY] || null
    const newObj = {
      taxCode: obj[iArr],
      accuracy: valAccuracy
    }
    arrCompare.push(newObj)
  }
  arrResultCompare = arrData.map((aArr: any) => {
    if(!aArr && aArr.length == 0) { return aArr }
    for (const row of aArr) {
      if( !row || !Array.isArray(row) || row.length === 0) { return aArr }
      let taxCode = row[iArr]
        , regexTaxCode = /^\d+(\-?)\d+/g
        , regexValidStandardTaxCode = /^\d{10}(\-\d{3})?$/g
      if(!taxCode || !taxCode.match(regexTaxCode)) {
        return aArr
      }
      if(!taxCode.match(regexValidStandardTaxCode)) {
        row.push(INVALID)
        return aArr
      }
      let arrFilterAccuracy
        , objFindAccuracy
        , valAccuracyCompare
        arrFilterAccuracy = arrCompare.filter((element: any) => element.taxCode === taxCode)
      log('Starting check', taxCode)
      if (arrFilterAccuracy && arrFilterAccuracy.length > 0) {
        objFindAccuracy = arrFilterAccuracy.find((element: any = {}) => (element.accuracy === VALID || element.accuracy === INVALID))
        valAccuracyCompare = objFindAccuracy && objFindAccuracy.accuracy || null
        if (valAccuracyCompare) { row.push(valAccuracyCompare) }
      }
    }
    return aArr
  })
  for (let i = 0; i < arrResultCompare.length; i++) {
    excelData = excelData.concat(arrResultCompare[i])
  }
  newDataExcelFile[sheet].data = excelData
  newDataExcelFile = xlsx.build(newDataExcelFile)
  dirExportFile = `${dirExportFile}${extFile}`
  writeFileExcel(newDataExcelFile, dirExportFile)
}

function reNewNameFile(oldName: string, addStr: string): object {
  const splitOldName: string[] = oldName.split('.'),
    maxIndex: number = splitOldName.length - 1
  let extFile: string = `.${splitOldName[maxIndex]}`,
    newName: string = `${oldName.replace(extFile,'')}${addStr}`,
    objRes: object = {
      newName,
      extFile
    }
  return objRes
}

function addCharacterStr (start: number, delCount: number, newSubStr: string, oldStr: string): string {
  return oldStr.slice(0, start) + newSubStr + oldStr.slice(start + Math.abs(delCount))
};

function cleanTaxCode(taxCode: string): string {
  let newTaxCode = taxCode,
    regexValidStandardTaxCode = /^\d{10}(\-\d{3})?$/g
  const TAX_CODE_FAILD: string = ''
  if(taxCode.match(regexValidStandardTaxCode)) { return newTaxCode }
  if(taxCode.length < 10 || taxCode.length > 14) {
    newTaxCode = TAX_CODE_FAILD
  } else {
    let regexDashes = /\-/g
    let arrMatch = taxCode.match(regexDashes) || []
    if(taxCode.length === 10 && arrMatch.length > 0){
      newTaxCode = TAX_CODE_FAILD
    } else if(taxCode.length === 14) {
      if(arrMatch.length !== 1) { newTaxCode = TAX_CODE_FAILD }
    } else{
      if(arrMatch.length !== 1) {
        if(arrMatch.length === 0 && taxCode.length === 13) {
          newTaxCode = addCharacterStr(10, 0, '-', taxCode)
        } else {
          newTaxCode = TAX_CODE_FAILD
        }
      } else {
        let arrSplitTC: string[] = taxCode.split('-'),
          maxLen: number = arrSplitTC.length - 1
        arrSplitTC[maxLen] = arrSplitTC[maxLen].padStart(3, '0')
        newTaxCode = arrSplitTC.join('-')
      }
    }
  }
  return newTaxCode
}

function writeFileExcel(excelData: string, newNameFile: string): void {
  try {
    writeFileSync(newNameFile, excelData)
  } catch (err) {
    error('** Write failed excel file', err)
  }
}

export class DataExcel {
  private nameFile: string
  static dataFileOriginal: object = {}
  static fileOriginal: object = {}
  
  constructor(nameFile: string) {
    this.nameFile = nameFile
  }

  importExcelFile(nameFile: string = this.nameFile): void {
    try {
      let excelFile: any
        , excelData: any
        , objData
        , sheet_name_list
        , sheet: number
        , dirImportFile = `${config.dirImportFile}${nameFile}`
      // ---- node-xlsx - 202110041006 ----
      excelFile = readFileSync(dirImportFile)
      excelData = xlsx.parse(excelFile)
      sheet = getSheetExcelFile(nameFile)
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
      DataExcel.fileOriginal = excelData
      objData = excelData[sheet].data
      DataExcel.dataFileOriginal = objData
      log('Import successfully excel file')
    } catch (err) {
      error('** Import failed excel file', err)
    }
  }

  exportExcelFile(dataNew: object = DataExcel.dataFileOriginal, fileOriginal: object = DataExcel.fileOriginal): void {
    editExcel(dataNew, this.nameFile, fileOriginal)
      .then(() => {
        log('Export successfully excel file')
      })
      .catch((err) => {
        error('** Export failed excel file', err)
      })
  }

  exportPrettyExcelFile(nameFile: string, dataNew: object = DataExcel.dataFileOriginal, fileOriginal: object = DataExcel.fileOriginal): void {
    prettyFileExcel(dataNew, nameFile, fileOriginal)
    log('Export successfully excel file')
  }

  exportCompareExcelFile(nameFile: string, dataNew: object = DataExcel.dataFileOriginal, fileOriginal: object = DataExcel.fileOriginal, nameCompareFile: string, dataCompare: object): void {
    compareExcel(dataNew, nameFile, fileOriginal, dataCompare, nameCompareFile)
    log('Export successfully excel file')
  }

  pushData2DB(dataNew: object = DataExcel.dataFileOriginal, fileOriginal: object = DataExcel.fileOriginal): void {
    callApiInvCloudTVan(dataNew, this.nameFile)
      .then(() => {
        log('Push data successfully to database')
      })
      .catch((err) => {
        error('** Push data failed to database', err)
      })
  }
}