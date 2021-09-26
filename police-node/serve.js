var express = require('express')
// const mysql = require('mysql')
const nodeExcel = require('excel-export')
var xlsx = require('node-xlsx')
const fs = require('fs')
const chrome = require('child_process')
const path = require('path')
var app = express()
app.get('/', function(req, res) {
  res.send('welcome！')
})

app.listen(3001, function() {
  console.log('app is listen at port 3001...')
})

const folderName = process.env.HOME || process.env.USERPROFILE + '\\Desktop\\police'
//! !用于结束后清空文件夹
// var files = fs.readdirSync(folderName)
// files.forEach((file) => {
//   fs.unlink(folderName + '\\' + file, (err) => {
//     if (err) throw err
//   })
// })
// drop table test;CREATE TABLE test(phone_id VARCHAR(20));
// 封装一个函数
var areaArr = ['浔阳','濂溪','柴桑','瑞昌','共青城','庐山','修水','武宁','永修','德安','都昌','湖口','彭泽']
var areaRegx = ['/浔阳/g','/濂溪/g','/柴桑/g','/瑞昌/g','/共青城/g','/庐山/g','/修水/g','/武宁/g','/永修/g','/德安/g','/都昌/g','/湖口/g','/彭泽/g']
var getMonth = (index, res) => {
  return new Promise((resolve, reject) => {
    try {
      if (!fs.existsSync(folderName)) {
        fs.mkdirSync(folderName)
      }
    } catch (err) {
      console.error(err)
    }

    fs.readdir(folderName, 'utf-8', (err, data) => {
      // var contentName = ''
      if (err) throw err
      console.log(data)
      var highRisk = /[\u5168\u56fd\u6d89\u75ab]{1}[\u4e00-\u9fa5]+[\u6f2b\u5165\u6c5f\u897f]{1}[\u4e00-\u9fa5]+[0-9]*/
      var highRisk = /[\u5168\u56fd\u6d89\u75ab]{1}[\u4e00-\u9fa5]+[\u6f2b\u5165\u6c5f\u897f]{1}[\u4e00-\u9fa5]+[0-9]*/
        for (const key of data) {
          // if (key === '涉疫重点关注手机号码.xlsx' || key === '涉疫重点关注手机号码.xls') {
          if (highRisk.test(key)) {
            var excelName = ''
            var sheetName = ''
            excelName = key
              var oldName = folderName + '\\' + key
              console.log(oldName)
              try {
              // Truncate Table tel
              // 表数据
              var tableData = xlsx.parse(oldName)
              // 循环读取表数据
              // 用户表数据
              console.log(tableData);
              for (var val in tableData) {
                  var userTableData = [];
                  var userTableData0 = [];var userTableData1 = [];var userTableData2 = [];var userTableData3 = [];var userTableData4 = [];
                  var userTableData5 = [];var userTableData6 = [];var userTableData7 = [];var userTableData8 = [];var userTableData9 = [];
                  var userTableData10 = [];var userTableData11 = [];var userTableData12 = [];var userTableData13 = [];var userTableData14 = [];
                  var userTableData15 = [];
                  var data0 = [];var data1 = [];var data2 = [];var data3 = [];var data4 = [];var data5 = [];var data6 = [];var data7 = [];
                  var data8 = [];var data9 = [];var data10 = [];var data11 = [];var data12 = [];var data13 = [];var data14 = [];var data15 = [];
                  var tiele = [];
                  var tempArr = null;
                  // 下标数据
                  var itemData = tableData[val]
                  sheetName = itemData.name
                  for (var index in itemData.data) {
                  // 0为表头数据
                  tempArr = itemData.data[2]
                  title = itemData.data[2]
                    if (index === 0 || index ===1) {
                      continue
                    }
                    var regx = /九江/g;
                    var str = itemData.data[index][2] + ''
                    var str2 = itemData.data[index][3] + ''
                    var phoneNums = str.match(regx)
                    if (regx.test(str)) {
                      for(let i = 0 ; i < areaArr.length; i++){
                        let rule = new RegExp(areaArr[i],'g')
                        if(str2.match(rule) && i==0){
                          userTableData0.push(itemData.data[index])
                        }
                        else if(str2.match(rule) && i==1){
                          userTableData1.push(itemData.data[index])
                        }
                        else if(str2.match(rule) && i==2){
                          userTableData2.push(itemData.data[index])
                        }
                        else if(str2.match(rule) && i==3){
                          userTableData3.push(itemData.data[index])
                        }
                        else if(str2.match(rule) && i==4){
                          userTableData4.push(itemData.data[index])
                        }
                        else if(str2.match(rule) && i==5){
                          userTableData5.push(itemData.data[index])
                        }
                        else if(str2.match(rule) && i==6){
                          userTableData6.push(itemData.data[index])
                        }
                        else if(str2.match(rule) && i==7){
                          userTableData7.push(itemData.data[index])
                        }
                        else if(str2.match(rule) && i==8){
                          userTableData8.push(itemData.data[index])
                        }
                        else if(str2.match(rule) && i==9){
                          userTableData9.push(itemData.data[index])
                        }
                        else if(str2.match(rule) && i==10){
                          userTableData10.push(itemData.data[index])
                        }
                        else if(str2.match(rule) && i==11){
                          userTableData11.push(itemData.data[index])
                        }
                        else if(str2.match(rule) && i==12){
                          userTableData12.push(itemData.data[index])
                        }
                        else if(str2 == "位置仅到地市" && i==12){   //i的值只为了循环得到一次userTableData13
                          userTableData13.push(itemData.data[index])
                        }
                        // else if(str2.match(rule) && i==14){
                        //   userTableData14.push(itemData.data[index])
                        // }
                        // else if(str2.match(rule) && i==15){
                        //   userTableData15.push(itemData.data[index])
                        // }
                      }

                      // userTableData.push(itemData.data[index])
                    }
                  }
                  console.log('走访表数据提取：', ['userTableData'])
                  //写入excel表
                  const conf = {}
                  conf.cols = []
                  conf.rows = []
                  for (const item of tempArr) {
                    const tits = {}
                    // 添加内容
                    tits.caption = item
                    // 添加对应类型，这类型对应数据库中的类型，入number，data但一般导出的都是转换为string类型的
                    tits.type = 'string'
                    // 将每一个表头加入cols中
                    conf.cols.push(tits)
                  }
                  conf.rows = userTableData
                  //由于各列数据长度不同，可以设置一下列宽
                  // const options = {'!cols': [{ wch: 10 }, { wch: 5 }, { wch: 15 }, { wch: 20 } ]};
                  //生成表格
                  if(userTableData0.length > 0){
                    data0.push(title)
                    data0.push(...userTableData0)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data0 }]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + areaArr[0] + userTableData0.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  if(userTableData1.length > 0){
                    data1.push(title)
                    data1.push(...userTableData1)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data1}]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + areaArr[1] + userTableData1.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  if(userTableData2.length > 0){
                    data2.push(title)
                    data2.push(...userTableData2)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data2}]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + areaArr[2] + userTableData2.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  if(userTableData3.length > 0){
                    data3.push(title)
                    data3.push(...userTableData3)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data3}]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + areaArr[3] + userTableData3.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  if(userTableData4.length > 0){
                    data4.push(title)
                    data4.push(...userTableData4)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data4}]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + areaArr[4] + userTableData4.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  if(userTableData5.length > 0){
                    data5.push(title)
                    data5.push(...userTableData5)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data5}]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + areaArr[5] + userTableData5.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  if(userTableData6.length > 0){
                    data6.push(title)
                    data6.push(...userTableData6)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data6}]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + areaArr[6] + userTableData6.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  if(userTableData7.length > 0){
                    data7.push(title)
                    data7.push(...userTableData7)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data7}]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + areaArr[7] + userTableData7.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  if(userTableData8.length > 0){
                    data8.push(title)
                    data8.push(...userTableData8)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data8}]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + areaArr[8] + userTableData8.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  if(userTableData9.length > 0){
                    data9.push(title)
                    data9.push(...userTableData9)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data9}]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + areaArr[9] + userTableData9.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  if(userTableData10.length > 0){
                    data10.push(title)
                    data10.push(...userTableData10)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data10}]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + areaArr[10] + userTableData10.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  if(userTableData11.length > 0){
                    data11.push(title)
                    data11.push(...userTableData11)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data11}]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + areaArr[11] + userTableData11.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  if(userTableData12.length > 0){
                    data12.push(title)
                    data12.push(...userTableData12)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data12}]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + areaArr[12] + userTableData12.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  if(userTableData13.length > 0){
                    data13.push(title)
                    data13.push(...userTableData13)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data13}]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + 'others' + userTableData13.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  if(userTableData14.length > 0){
                    data14.push(title)
                    data14.push(...userTableData14)
                    // console.log(...userTableData);
                    let buffer = xlsx.build([{name:'sheet1',data:data14}]);
                    // `sheet九江数据：+${userTableData.length}`
                    let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                    let filePath = excelName.match(excelNameRegx) +  "-"+ sheetName + "-" + 'areaArr[14]' + userTableData14.length +"条数据" + '.xlsx';
                    let finalPath = path.resolve(__dirname,filePath)
                    fs.writeFileSync(finalPath,buffer,{'flag':'w'});
                  }
                  // 将所有数据写入nodeExcel中
                  // console.log(conf);
                  // const result = nodeExcel.execute([conf])
                  // // 设置响应头，在Content-Type中加入编码格式为utf-8即可实现文件内容支持中文
                  // res.setHeader('Content-Type', 'application/vnd.openxmlformats;charset=utf-8')
                  // // 设置下载文件命名，中文文件名可以通过编码转化写入到header中。
                  // res.setHeader('Content-Disposition', 'attachment; filename=' + encodeURI(`${oldName}筛选后的文件`) + '.xlsx')
                  // // 将文件内容传入
                  // res.end(result, 'binary')
                }
              // console.log('-------------end-------------')
              } catch (e) {
              // 输出日志
                console.log('excel读取异常,error=%s', e.stack)
              }
          } 
          else if (key === '通话.xlsx' || key === '通话.xls') {
            var promise2 = new Promise((resolve, reject) => {
              var excelName = ''
              var sheetName = ''
              excelName = key
              var oldName = folderName + '\\' + key
              console.log(oldName)
              try {
                  // 表数据
                  var tableData = xlsx.parse(oldName)
                  // 循环读取表数据
                  // 用户表数据
                  console.log(tableData);
                  for (var val in tableData) {
                      var userTableData = [];
                      var userTableData0 = []; var userTableData1 = []; var userTableData2 = []; var userTableData3 = []; var userTableData4 = [];
                      var userTableData5 = []; var userTableData6 = []; var userTableData7 = []; var userTableData8 = []; var userTableData9 = [];
                      var userTableData10 = []; var userTableData11 = []; var userTableData12 = []; var userTableData13 = []; var userTableData14 = [];
                      var userTableData15 = [];
                      var data0 = []; var data1 = []; var data2 = []; var data3 = []; var data4 = []; var data5 = []; var data6 = []; var data7 = [];
                      var data8 = []; var data9 = []; var data10 = []; var data11 = []; var data12 = []; var data13 = []; var data14 = []; var data15 = [];
                      var tiele = [];
                      var tempArr = null;
                      // 下标数据
                      var itemData = tableData[val]
                      sheetName = itemData.name
                      for (var index in itemData.data) {
                          // 0为表头数据
                          tempArr = itemData.data[2]
                          title = itemData.data[2]
                          if (index === 0 || index === 1) {
                              continue
                          }
                          var regx = /九江/g;
                          var str = itemData.data[index][2] + ''
                          var str2 = itemData.data[index][3] + ''
                          var phoneNums = str.match(regx)
                          if (regx.test(str)) {
                              for (let i = 0; i < areaArr.length; i++) {
                                  let rule = new RegExp(areaArr[i], 'g')
                                  if (str2.match(rule) && i == 0) {
                                      userTableData0.push(itemData.data[index])
                                  }
                                  else if (str2.match(rule) && i == 1) {
                                      userTableData1.push(itemData.data[index])
                                  }
                                  else if (str2.match(rule) && i == 2) {
                                      userTableData2.push(itemData.data[index])
                                  }
                                  else if (str2.match(rule) && i == 3) {
                                      userTableData3.push(itemData.data[index])
                                  }
                                  else if (str2.match(rule) && i == 4) {
                                      userTableData4.push(itemData.data[index])
                                  }
                                  else if (str2.match(rule) && i == 5) {
                                      userTableData5.push(itemData.data[index])
                                  }
                                  else if (str2.match(rule) && i == 6) {
                                      userTableData6.push(itemData.data[index])
                                  }
                                  else if (str2.match(rule) && i == 7) {
                                      userTableData7.push(itemData.data[index])
                                  }
                                  else if (str2.match(rule) && i == 8) {
                                      userTableData8.push(itemData.data[index])
                                  }
                                  else if (str2.match(rule) && i == 9) {
                                      userTableData9.push(itemData.data[index])
                                  }
                                  else if (str2.match(rule) && i == 10) {
                                      userTableData10.push(itemData.data[index])
                                  }
                                  else if (str2.match(rule) && i == 11) {
                                      userTableData11.push(itemData.data[index])
                                  }
                                  else if (str2.match(rule) && i == 12) {
                                      userTableData12.push(itemData.data[index])
                                  }
                                  else if (str2 == "位置仅到地市" && i == 12) {   //i的值只为了循环得到一次userTableData13
                                      userTableData13.push(itemData.data[index])
                                  }
                              }
                          }
                      }
                      console.log('走访表数据提取：', ['userTableData'])
                      //写入excel表
                      const conf = {}
                      conf.cols = []
                      conf.rows = []
                      for (const item of tempArr) {
                          const tits = {}
                          // 添加内容
                          tits.caption = item
                          // 添加对应类型，这类型对应数据库中的类型，入number，data但一般导出的都是转换为string类型的
                          tits.type = 'string'
                          // 将每一个表头加入cols中
                          conf.cols.push(tits)
                      }
                      conf.rows = userTableData
                      //由于各列数据长度不同，可以设置一下列宽
                      // const options = {'!cols': [{ wch: 10 }, { wch: 5 }, { wch: 15 }, { wch: 20 } ]};
                      //生成表格
                      if (userTableData0.length > 0) {
                          data0.push(title)
                          data0.push(...userTableData0)
                          let buffer = xlsx.build([{ name: 'sheet1', data: data0 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + areaArr[0] + userTableData0.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                      if (userTableData1.length > 0) {
                          data1.push(title)
                          data1.push(...userTableData1)
                          let buffer = xlsx.build([{ name: 'sheet1', data: data1 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + areaArr[1] + userTableData1.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                      if (userTableData2.length > 0) {
                          data2.push(title)
                          data2.push(...userTableData2)
                          let buffer = xlsx.build([{ name: 'sheet1', data: data2 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + areaArr[2] + userTableData2.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                      if (userTableData3.length > 0) {
                          data3.push(title)
                          data3.push(...userTableData3)
                          let buffer = xlsx.build([{ name: 'sheet1', data: data3 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + areaArr[3] + userTableData3.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                      if (userTableData4.length > 0) {
                          data4.push(title)
                          data4.push(...userTableData4)
                          let buffer = xlsx.build([{ name: 'sheet1', data: data4 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + areaArr[4] + userTableData4.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                      if (userTableData5.length > 0) {
                          data5.push(title)
                          data5.push(...userTableData5)
                          let buffer = xlsx.build([{ name: 'sheet1', data: data5 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + areaArr[5] + userTableData5.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                      if (userTableData6.length > 0) {
                          data6.push(title)
                          data6.push(...userTableData6)
                          let buffer = xlsx.build([{ name: 'sheet1', data: data6 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + areaArr[6] + userTableData6.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                      if (userTableData7.length > 0) {
                          data7.push(title)
                          data7.push(...userTableData7)
                          let buffer = xlsx.build([{ name: 'sheet1', data: data7 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + areaArr[7] + userTableData7.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                      if (userTableData8.length > 0) {
                          data8.push(title)
                          data8.push(...userTableData8)
                          let buffer = xlsx.build([{ name: 'sheet1', data: data8 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + areaArr[8] + userTableData8.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                      if (userTableData9.length > 0) {
                          data9.push(title)
                          data9.push(...userTableData9)
                          let buffer = xlsx.build([{ name: 'sheet1', data: data9 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + areaArr[9] + userTableData9.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                      if (userTableData10.length > 0) {
                          data10.push(title)
                          data10.push(...userTableData10)
                          let buffer = xlsx.build([{ name: 'sheet1', data: data10 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + areaArr[10] + userTableData10.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                      if (userTableData11.length > 0) {
                          data11.push(title)
                          data11.push(...userTableData11)
                          let buffer = xlsx.build([{ name: 'sheet1', data: data11 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + areaArr[11] + userTableData11.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                      if (userTableData12.length > 0) {
                          data12.push(title)
                          data12.push(...userTableData12)
                          let buffer = xlsx.build([{ name: 'sheet1', data: data12 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + areaArr[12] + userTableData12.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                      if (userTableData13.length > 0) {
                          data13.push(title)
                          data13.push(...userTableData13)
                          // console.log(...userTableData);
                          let buffer = xlsx.build([{ name: 'sheet1', data: data13 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + 'others' + userTableData13.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                      if (userTableData14.length > 0) {
                          data14.push(title)
                          data14.push(...userTableData14)
                          // console.log(...userTableData);
                          let buffer = xlsx.build([{ name: 'sheet1', data: data14 }]);
                          let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
                          let filePath = excelName.match(excelNameRegx) + "-" + sheetName + "-" + 'areaArr[14]' + userTableData14.length + "条数据" + '.xlsx';
                          let finalPath = path.resolve(__dirname, filePath)
                          fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
                      }
                  }
              } catch (e) {
                  console.log('excel读取异常,error=%s', e.stack)
              }
            })
          }
        }
    })    
    resolve(res)
  })
}
// 后台导出走访表接口
{app.get('/exportExcel1', function(req, res) {
  const conf = {}
  conf.cols = []
  const thisres = res;
  (() => {
    getMonth(1, thisres)
  })()
})
app.get('/exportExcel2', function(req, res) {
  const conf = {}
  conf.cols = []
  const thisres = res;
  (() => {
    getMonth(2, thisres)
  })()
})
app.get('/exportExcel3', function(req, res) {
  const conf = {}
  conf.cols = []
  const thisres = res;
  (() => {
    getMonth(3, thisres)
  })()
})
app.get('/exportExcel4', function(req, res) {
  const conf = {}
  conf.cols = []
  const thisres = res;
  (() => {
    getMonth(4, thisres)
  })()
})
app.get('/exportExcel5', function(req, res) {
  const conf = {}
  conf.cols = []
  const thisres = res;
  (() => {
    getMonth(5, thisres)
  })()
})
app.get('/exportExcel6', function(req, res) {
  const conf = {}
  conf.cols = []
  const thisres = res;
  (() => {
    getMonth(6, thisres)
  })()
})
app.get('/exportExcel7', function(req, res) {
  const conf = {}
  conf.cols = []
  const thisres = res;
  (() => {
    getMonth(7, thisres)
  })()
})
app.get('/exportExcel8', function(req, res) {
  const conf = {}
  conf.cols = []
  const thisres = res;
  (() => {
    getMonth(8, thisres)
  })()
})
app.get('/exportExcel9', function(req, res) {
  const conf = {}
  conf.cols = []
  const thisres = res;
  (() => {
    getMonth(9, thisres)
  })()
})
app.get('/exportExcel10', function(req, res) {
  const conf = {}
  conf.cols = []
  const thisres = res;
  (() => {
    getMonth(10, thisres)
  })()
})
app.get('/exportExcel11', function(req, res) {
  const conf = {}
  conf.cols = []
  const thisres = res;
  (() => {
    getMonth(11, thisres)
  })()
})
app.get('/exportExcel12', function(req, res) {
  const conf = {}
  conf.cols = []
  const thisres = res;
  (() => {
    getMonth(12, thisres)
  })()
})}

getMonth();
// chrome.exec('start http://localhost:3001/exportExcel1')
// for (let i = 1; i <= 12; i++) {
//   const url = 'start http://localhost:3001/exportExcel' + i
//   chrome.exec(url)
// }
// chrome.exec('start http://localhost:3000/exportExcel3')

// const first = new Promise((resolve, reject) => {
//   setTimeout(resolve, 500, '第一个')
// })
// const second = new Promise((resolve, reject) => {
//   setTimeout(resolve, 100, '第二个')
// })

// Promise.race([first, second]).then(result => {
//   console.log(result) // 第二个
// })

// ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec']
// ['浔阳区','濂溪区','柴桑区','瑞昌市','共青城市','庐山市','修水县','武宁县','永修县','德安县','都昌县','湖口县','彭泽县']
// ['浔阳','濂溪','柴桑','瑞昌','共青城','庐山','修水','武宁','永修','德安','都昌','湖口','彭泽']
// ['/浔阳/g','/濂溪/g','/柴桑/g','/瑞昌/g','/共青城/g','/庐山/g','/修水/g','/武宁/g','/永修/g','/德安/g','/都昌/g','/湖口/g','/彭泽/g']
// var regx = /九江/g
// var regx0 = /浔阳/g
// var regx1 = /濂溪/g
// var regx2 = /柴桑/g
// var regx3 = /瑞昌/g
// var regx4 = /共青城/g
// var regx5 = /庐山/g
// var regx6 = /修水/g
// var regx7 = /武宁/g
// var regx8 = /永修/g
// var regx9 = /德安/g
// var regx10 = /都昌/g
// var regx11 = /湖口/g
// var regx12 = /彭泽/g
