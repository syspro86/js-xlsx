'use strict'

var jsxlsx = require('../lib/index')

var wb = new jsxlsx.Workbook()
wb.loadFile('sample.xlsx').then((excel) => {
    console.log(excel)
})
