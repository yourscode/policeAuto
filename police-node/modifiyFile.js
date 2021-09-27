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
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + areaArr[0] + userTableData0.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
        if (userTableData1.length > 0) {
            data1.push(title)
            data1.push(...userTableData1)
            let buffer = xlsx.build([{ name: 'sheet1', data: data1 }]);
            let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + areaArr[1] + userTableData1.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
        if (userTableData2.length > 0) {
            data2.push(title)
            data2.push(...userTableData2)
            let buffer = xlsx.build([{ name: 'sheet1', data: data2 }]);
            let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + areaArr[2] + userTableData2.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
        if (userTableData3.length > 0) {
            data3.push(title)
            data3.push(...userTableData3)
            let buffer = xlsx.build([{ name: 'sheet1', data: data3 }]);
            let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + areaArr[3] + userTableData3.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
        if (userTableData4.length > 0) {
            data4.push(title)
            data4.push(...userTableData4)
            let buffer = xlsx.build([{ name: 'sheet1', data: data4 }]);
            let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + areaArr[4] + userTableData4.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
        if (userTableData5.length > 0) {
            data5.push(title)
            data5.push(...userTableData5)
            let buffer = xlsx.build([{ name: 'sheet1', data: data5 }]);
            let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + areaArr[5] + userTableData5.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
        if (userTableData6.length > 0) {
            data6.push(title)
            data6.push(...userTableData6)
            let buffer = xlsx.build([{ name: 'sheet1', data: data6 }]);
            let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + areaArr[6] + userTableData6.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
        if (userTableData7.length > 0) {
            data7.push(title)
            data7.push(...userTableData7)
            let buffer = xlsx.build([{ name: 'sheet1', data: data7 }]);
            let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + areaArr[7] + userTableData7.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
        if (userTableData8.length > 0) {
            data8.push(title)
            data8.push(...userTableData8)
            let buffer = xlsx.build([{ name: 'sheet1', data: data8 }]);
            let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + areaArr[8] + userTableData8.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
        if (userTableData9.length > 0) {
            data9.push(title)
            data9.push(...userTableData9)
            let buffer = xlsx.build([{ name: 'sheet1', data: data9 }]);
            let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + areaArr[9] + userTableData9.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
        if (userTableData10.length > 0) {
            data10.push(title)
            data10.push(...userTableData10)
            let buffer = xlsx.build([{ name: 'sheet1', data: data10 }]);
            let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + areaArr[10] + userTableData10.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
        if (userTableData11.length > 0) {
            data11.push(title)
            data11.push(...userTableData11)
            let buffer = xlsx.build([{ name: 'sheet1', data: data11 }]);
            let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + areaArr[11] + userTableData11.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
        if (userTableData12.length > 0) {
            data12.push(title)
            data12.push(...userTableData12)
            let buffer = xlsx.build([{ name: 'sheet1', data: data12 }]);
            let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + areaArr[12] + userTableData12.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
        if (userTableData13.length > 0) {
            data13.push(title)
            data13.push(...userTableData13)
            // console.log(...userTableData);
            let buffer = xlsx.build([{ name: 'sheet1', data: data13 }]);
            let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + 'others' + userTableData13.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
        if (userTableData14.length > 0) {
            data14.push(title)
            data14.push(...userTableData14)
            // console.log(...userTableData);
            let buffer = xlsx.build([{ name: 'sheet1', data: data14 }]);
            let excelNameRegx = /[\u4e00-\u9fa5]+[0-9]*/
            let filePath =  excelName.substring(0,excelName.length-4) + "-" + sheetName + "-" + 'areaArr[14]' + userTableData14.length + "条数据" + '.xlsx';
            let finalPath = path.resolve(__dirname, filePath)
            fs.writeFileSync(finalPath, buffer, { 'flag': 'w' });
        }
    }
} catch (e) {
    console.log('excel读取异常,error=%s', e.stack)
}