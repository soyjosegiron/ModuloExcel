const XLSX = require("xlsx");


function ExcelToJSON (file){
    const workbook = XLSX.utils.sheet_to_json(file)
}

