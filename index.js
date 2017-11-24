const XLSX = require('xlsx')
const { readDataFromDir } = require('./util')

const rs = readDataFromDir('./data/three/part1/')
console.log(
)

XLSX.writeFile({
  SheetNames: ['mySheet'],
  Sheets: {
    'mySheet': {
        '!ref': 'A1:E1000', // 必须要有这个范围才能输出，否则导出的 excel 会是一个空表
        ...rs,
    }
  }
}, './output/three/part1.xlsx')
