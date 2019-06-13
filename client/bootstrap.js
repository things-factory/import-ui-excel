import { UPDATE_EXTENSION } from '@things-factory/import-base'
import { store } from '@things-factory/shell'
import * as XLSX from 'xlsx'

function importXlsx(params) {
  excelToJson(params)
}

function importXls(params) {
  excelToJson(params)
}

function excelToJson(params) {
  let workbook = XLSX.read(params, {
    type: 'binary'
  })

  const firstSheetName = workbook.SheetNames[0]
  const worksheet = workbook.Sheets[firstSheetName]
  console.log(
    XLSX.utils.sheet_to_json(worksheet, {
      raw: true
    })
  )
}

export default function bootstrap() {
  store.dispatch({
    type: UPDATE_EXTENSION,
    extensions: {
      xlsx: {
        import: importXlsx
      },
      xls: {
        import: importXls
      }
    }
  })
}
