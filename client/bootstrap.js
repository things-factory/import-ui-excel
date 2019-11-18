import { UPDATE_EXTENSION } from '@things-factory/import-base'
import { store } from '@things-factory/shell'
import * as XLSX from '!xlsx'
import Excel from '!exceljs'

/**
 * Convert Excel to Object
 * @param {ArrayBufferTypes} params - Array Buffer of the excel file.
 * @return {object} - Return the first sheet into Object, other sheet is used as a supporting data.
 */
async function excelToObj(params) {
  let workbook = new Excel.Workbook()
  await workbook.xlsx.load(params)

  let ws = workbook.getWorksheet(1)
  ws._rows[0]._cells.map((cell, index) => {
    ws._columns[index]._key = cell.name
  })

  ////Fetch supporting data and place it under 'extraData' object.
  let extraData = {}
  if (workbook._worksheets.length > 2) {
    for (let index = 2; index < workbook._worksheets.length; index++) {
      let header = []
      let worksheetName = workbook._worksheets[index].name

      workbook.getWorksheet(worksheetName).eachRow((x, index) => {
        let row = x.values.filter(val => val)
        let obj = {}
        if (index === 1) {
          header = row
          extraData[worksheetName] = []
        } else {
          header.map((key, i) => {
            obj[key] = row[i]
          })
          extraData[worksheetName].push(obj)
        }
      })
    }
  }

  ////Fetch all data and place it in records array, 'extraData' object is used to map the id for list type.
  let records = []
  for (let rowcount = 1; rowcount < ws._rows.length; rowcount++) {
    let objRow = {}

    for (let columncount = 0; columncount < ws._rows[rowcount]._cells.length; columncount++) {
      let currentCell = ws._rows[rowcount]._cells[columncount]
      let columnType = ''
      if (currentCell) {
        let currentColumnCode = currentCell._address.match(/[a-z]+|[^a-z]+/gi)[0]
        columnType = ws.dataValidations.model[currentColumnCode + '2:' + currentColumnCode + ws._rows.length.toString()]
          ? ws.dataValidations.model[currentColumnCode + '2:' + currentColumnCode + ws._rows.length.toString()].type
          : undefined
      }

      let cellVal = currentCell ? currentCell.value : ''

      let arrColumnKeys = ws._columns[columncount].key.split('.')
      if (arrColumnKeys.length > 1) {
        arrColumnKeys.reduce((prev, e, index, arr) => {
          if (arr.length - 1 > index) {
            prev[e] = {}
          } else {
            if (extraData[ws._columns[columncount]._key]) {
              let extraDataItem = extraData[ws._columns[columncount]._key].filter(itm => itm.name === cellVal)
              prev['id'] = extraDataItem.length > 0 ? extraDataItem[0].id.toString() : null
            }
            switch (columnType) {
              case 'decimal':
                prev[e] = cellVal ? cellVal : 0
                break
              default:
                prev[e] = cellVal ? cellVal.toString() : null
                break
            }
          }
          return prev[e]
        }, objRow)
      } else {
        if (
          (ws._rows[0]._cells[columncount].value === 'id' && cellVal !== '') ||
          ws._rows[0]._cells[columncount].value !== 'id'
        ) {
          switch (columnType) {
            case 'decimal':
              objRow[ws._columns[columncount].key] = cellVal ? cellVal : 0
              break
            default:
              objRow[ws._columns[columncount].key] = cellVal ? cellVal.toString() : null
              break
          }
        }
      }
    }
    records.push(objRow)
  }

  return records
}

function excelToJson(params) {
  let workbook = XLSX.read(params, {
    type: 'binary'
  })

  const firstSheetName = workbook.SheetNames[0]
  const worksheet = workbook.Sheets[firstSheetName]
  let records = XLSX.utils.sheet_to_json(worksheet, { raw: true })
  return records
}

export default function bootstrap() {
  store.dispatch({
    type: UPDATE_EXTENSION,
    extensions: {
      xlsx: {
        import: excelToObj
      },
      xls: {
        import: excelToJson
      }
    }
  })
}
