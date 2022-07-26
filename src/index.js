const path = require('path');
const fs = require('fs');
const _ = require('lodash');
const exceljs = require('exceljs');

const { common: zhObj } = require('./data/zh');

const workBook = new exceljs.Workbook();
const readExcel = (src, sheet, col1, col2) => {
  const transMapArr = [];
  return new Promise((resolve, reject) => {
    workBook.xlsx.readFile(src).then(() => {
      // 获取第2张表
      const worksheet = workBook.getWorksheet(sheet);
      worksheet.eachRow((row, rowNumber) => {
        const arr = row.values
        transMapArr.push({
          key: arr[col1],
          value: arr[col2],
        })
        if (rowNumber + 1 === worksheet.columnCount) {
          resolve(transMapArr);
        }
      })
    }, err => {
      console.error(err);
      reject([]);
    })
  })
}

const writeTxt = async (src, data) => {
  const ws = fs.createWriteStream(src, {
    encoding: 'utf8',
  })
  ws.write(data);
}

const startDiff = async (srcPath, outPath) => {
  const inputSrc = path.join(srcPath, '系统中英翻译表.xlsx');
  const outputTxtSrc = path.join(outPath, '系统定义但产品缺少的字段.txt');
  const outputENSrc = path.join(outPath, '翻译后的EN字段.txt');

  // 以下是替换为英文的操作
  const transMapArr = await readExcel(inputSrc, 2, 5, 6);
  const zhArr = Object.entries(zhObj);
  const cloneZhObj = _.cloneDeep(zhObj);
  for (let i = 0; i < transMapArr.length; i++) {
    const { key, value } = transMapArr[i];
    for (let j = 0; j < zhArr.length; j++) {
      const zhKey = zhArr[j][0];
      const zhValue = zhArr[j][1];
      if (key === zhValue) {
        cloneZhObj[zhKey] = value;
        continue;
      }
    }
  }
  const arr1 = Object.values(zhObj);
  const arr2 = transMapArr.map(item => item.key);
  // A表是我这里有产品没有的
  const ASheet = _.difference(arr1, arr2);
  // B表是产品有我这里没有的
  const BSheet = _.difference(arr2, arr1);

  writeTxt(outputTxtSrc, ASheet.join('\n'));
  writeTxt(outputENSrc, JSON.stringify(cloneZhObj));
  return {
    zhObj: cloneZhObj,
    ASheet,
    BSheet,
  };
}

const src = path.resolve(__dirname, 'data');
const out = path.resolve(__dirname, 'out');
startDiff(src, out).then(data => {
  console.log(data);
}, err => {
  console.log(err);
})
