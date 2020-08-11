const fs = require('fs')
const moment = require('moment-timezone')
const XLSX = require('xlsx')
moment.locale('ru')

class BOParser {

  constructor () {
    // TODO: тут наверное надо воткнуть логгер и его же добавить в вызовы функций
  }

  parse (file) {
    const config = this.getScheme()
    const workbook = XLSX.readFile(file)

    const result = Object.keys(config).reduce((out, current) => {
      let info = config[current]
      let worksheet = workbook.Sheets[info.name]
      this.findCell(worksheet, info)
      let balance = this.findDateColumns(worksheet, info)
      this.parseBalance(worksheet, info, balance)
      return { ...out, [current]: balance }
    }, {})

    // Оставил для дебага
    // fs.writeFileSync('result.json', JSON.stringify(result, null, 2))

    return result
  }

  getScheme () {
    return {
      f1: {
        name: 'Balance',
        start: null,
        startPattern: /Код строки/,
        amountCallback: str => parseInt(str.trim().replace(/ /g, '')) || "",
        columns: []
      },
      f2: {
        name: 'Financial Result',
        start: null,
        startPattern: /Код строки/,
        amountCallback: str => {
          if (str.trim().match(/^\([0-9 ]{1,}\)$/)) {
            str = '-' + str.replace(/\(|\)/g, '')
          }
          return parseInt(str.trim().replace(/ /g, '')) || ""
        },
        columns: []
      }
    }
  }

  findCell (ws, info) {
    const { startPattern } = info
    const range = XLSX.utils.decode_range(ws['!ref'])
    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        let address = { c: C, r: R }
        let cell = ws[XLSX.utils.encode_cell(address)]
        if (cell.v.match(startPattern)) {
          info.start = address
          return
        }
      }
    }
  }

  findDateColumns (ws, info) {
    const { start } = info
    const range = XLSX.utils.decode_range(ws['!ref'])
    const result = []
    for (let C = start.c + 1; C <= range.e.c; C++) {
      let address = { c: C, r: start.r }
      let cell = ws[XLSX.utils.encode_cell(address)]
      if (cell.v) {
        info.columns.push(C)
        result.push({ date: this.parseDate(cell.v), assets: [] })
      }
    }
    return result
  }

  parseBalance (ws, info, result) {
    const { start, amountCallback, columns } = info
    const range = XLSX.utils.decode_range(ws['!ref'])
    for (let R = start.r; R <= range.e.r; R++) {
      let codeCell = ws[XLSX.utils.encode_cell({ c: start.c, r: R })]
      if (codeCell.v.trim().match(/\d{4}/)) {
        for (let i = 0; i < columns.length; i++) {
          let amountCell = ws[XLSX.utils.encode_cell({ c: columns[i], r: R })]
          result[i].assets.push({
            code: parseInt(codeCell.v),
            amount: amountCallback ? amountCallback(amountCell.v) : amountCell.v
          })
        }
      }
    }
  }

  parseDate (str) {
    str = str.replace(/На|За|г\./g, '').trim()
    if (str.match(/^\d{4}$/)) {
      str = `31 декабря ${str}`
    }
    return moment.tz(str, "DD MMMM YYYY", 'UTC').toDate()
  }

}

const parser = new BOParser()
parser.parse('test/1651000010.xlsx')
parser.parse('test/1831176365.xlsx')
parser.parse('test/7714420892.xlsx')
parser.parse('test/7810633420.xlsx')
