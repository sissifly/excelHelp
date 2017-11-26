const fs = require('fs')
const XLSX = require('xlsx')
const R = require('ramda')

// =========================================================
// 读取数据范围
const range = { start: ['A', 4], end: ['M', Infinity]}
// 文件路径存储列
const pathCol = 'N'
// 写入output.xlsx的起始行
const startRow = 4
// =========================================================

function inRange(pos) {
  const row = getRow(pos)
  const col = getCol(pos)

  return row >= range.start[1]
    && row <= range.end[1]
    && col >= range.start[0]
    && col <= range.end[0]
}

function getRow(pos) { return Number(pos.replace(/^[A-Z]+/, '')) }
function getCol(pos) { return pos.replace(/\d+$/, '') }

function getSheetData(sheet, fileName) {
  return R.pipe(
    R.pickBy((val, key) => (/^[A-Z]+\d+$/.test(key) && inRange(key))),  // 去除范围外或特殊属性数据
    (sheetDatas) => {
      const rowNos = R.pipe(
        R.keys,
        R.map(getRow),
      )(sheetDatas)
      const minRow = R.apply(Math.min)(rowNos)
      const maxRow = R.apply(Math.max)(rowNos)

      // 将文件路径插入到sheet指定位置中
      const pathCell = R.pipe(
        R.map(row => pathCol + row),
        R.map(pos => ({ [pos]: { v: fileName } })),
        R.mergeAll,
      )(R.range(minRow, maxRow + 1))

      return R.merge(sheetDatas, pathCell)
    },
    R.toPairs, // 转换格式为序对[A1, { v: xxx }]
  )(sheet)
}

exports.readDataFromDir = function(dir) {
  const data = R.pipe(
    R.filter(R.test(/\.xlsx$/)),                                  // 读取目录，过滤掉非xlsx文件
    R.map(R.concat(dir)),                                         // 拼接目录与文件名 ./samples/xxx.xlsx
    R.map(path => [path, XLSX.readFile(path)]),                   // 转换文件为 workbook
    R.map(([path, wb]) => [path, wb.Sheets[wb.SheetNames[1]]]),   // 建立path与workbook的关系，后续读取数据需要使用path
    R.map(([path, sheet]) => getSheetData(sheet, path)),          // 读取每个sheet中符合要求求的数据
  )(fs.readdirSync(dir))

  // 将多个文件数据合并，每个文件数据开始行的游标
  let rowCursor = 1
  // 根据游标更新位置中的行
  function updateRow(pos, minRow) {
    return getCol(pos) + (getRow(pos) - minRow + rowCursor)
  }
  // 根据起始行计算偏移
  function offsetRow(pos, offsetRow) {
    return getCol(pos) + (getRow(pos) + offsetRow - 1)
  }

  return R.pipe(
    R.map((sheetDatas) => {
      const rowNos = R.pipe(
        R.map(R.head),
        R.map(R.replace(/^[A-Z]+/, '')),
        R.map(Number),
      )(sheetDatas)
      const minRow = R.apply(Math.min)(rowNos)
      const maxRow = R.apply(Math.max)(rowNos)

      const rs = R.map(([pos, cell]) => [updateRow(pos, minRow), cell])(sheetDatas)

      // 更新游标，拼接下一个sheet数据行号开始位置
      rowCursor += (maxRow - minRow + 1)

      return rs
    }),
    // 合并多个sheet数据
    R.unnest,
    // 根据起始行位移
    R.map(([pos, cell]) => [offsetRow(pos, startRow), cell]),
    R.fromPairs,
    (v) => (console.log(111, v), v),
  )(data)
}
