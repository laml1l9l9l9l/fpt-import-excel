import config from "./config"
import { DataExcel } from "./utils/data_excel"

console.log('Started Code')
let nameFileRDD1029
  , nameFileTaxRef

function accuracyExcelFile(nameFile: string) {
  const excel = new DataExcel(nameFile)
  excel.importExcelFile()
  excel.exportExcelFile()
}

function accuracyCompareExcelFiles(nameFile: string, nameFileCompare: string) {
  const excel = new DataExcel(nameFile)
    , excelCompare = new DataExcel(nameFileCompare)
  let dataFileCompare: any
    , dataFileAccuracy: any
    , fileOriginalAccuracy: any
  excel.importExcelFile(nameFile)
  dataFileAccuracy = JSON.parse(JSON.stringify(DataExcel.dataFileOriginal))
  fileOriginalAccuracy = JSON.parse(JSON.stringify(DataExcel.fileOriginal))
  excelCompare.importExcelFile(nameFileCompare)
  dataFileCompare = JSON.parse(JSON.stringify(DataExcel.dataFileOriginal))
  excel.exportCompareExcelFile(nameFile, dataFileAccuracy, fileOriginalAccuracy, nameFileCompare, dataFileCompare)
}

function prettyExcelFiles(nameFile: string) {
  const excel = new DataExcel(nameFile)
  let dataFileCompare: any
    , dataFileAccuracy: any
    , fileOriginalAccuracy: any
  excel.importExcelFile(nameFile)
  dataFileAccuracy = JSON.parse(JSON.stringify(DataExcel.dataFileOriginal))
  fileOriginalAccuracy = JSON.parse(JSON.stringify(DataExcel.fileOriginal))
  dataFileCompare = JSON.parse(JSON.stringify(DataExcel.dataFileOriginal))
  excel.exportPrettyExcelFile(nameFile, dataFileAccuracy, fileOriginalAccuracy)
}

function pushDBCloud(nameFile: string) {
  const excel = new DataExcel(nameFile)
  excel.importExcelFile()
  excel.pushData2DB()
}

try {
  nameFileTaxRef = '/Tax\ ref_VN\ customers.xlsx'
  nameFileRDD1029 = `/RDD1029\ Customer\ Data\ Vietnam\ LnS.xlsb`
  // ----- Accuracy excel -----
  // accuracyExcelFile(nameFileTaxRef)
  // accuracyExcelFile(nameFileRDD1029) // | Don't use
  // accuracyCompareExcelFiles(nameFileRDD1029, nameFileTaxRef)
  // accuracyCompareExcelFiles(nameFileRDD1029, nameFileTaxRef)
  prettyExcelFiles(nameFileTaxRef)
  // ----- Push DB Cloud -----
  // pushDBCloud(nameFileTaxRef)
} catch (e) {
  console.error('Err read file', e)
}