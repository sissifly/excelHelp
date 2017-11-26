const XLSX = require('xlsx')
const { readDataFromDir } = require('./util')

const rs = readDataFromDir('./samples/')
console.log(
)

XLSX.writeFile({
  SheetNames: ['mySheet'],
  Sheets: {
    'mySheet': {
        '!ref': 'A1:ZZ1000', // 必须要有这个范围才能输出，否则导出的 excel 会是一个空表
        ...rs,
    }
  }
}, 'output.xlsx')
