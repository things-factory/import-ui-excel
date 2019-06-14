import { UPDATE_EXTENSION } from '@things-factory/import-base'
import { store } from '@things-factory/shell'
import * as XLSX from 'xlsx'

function excelToJson(params) {
  let workbook = XLSX.read(params, {
    type: 'binary'
  })

  const firstSheetName = workbook.SheetNames[0]
  const worksheet = workbook.Sheets[firstSheetName]
  return XLSX.utils.sheet_to_json(worksheet, { raw: true })
}

export default function bootstrap() {
  store.dispatch({
    type: UPDATE_EXTENSION,
    extensions: {
      xlsx: {
        import: excelToJson
      },
      xls: {
        import: excelToJson
      }
    }
  })
}
