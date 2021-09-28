var express = require('express')
// const mysql = require('mysql')
const nodeExcel = require('excel-export')
var xlsx = require('node-xlsx')
const fs = require('fs')
const chrome = require('child_process')
const path = require('path')
const e = require('express')
const multiparty = require('multiparty')
const getMonth = require("./screen.js")
var app = express()

app.listen(3001, function() {
  console.log('app is listen at port 3001...')
})
var areaArr = ['浔阳','濂溪','柴桑','瑞昌','共青城','庐山','修水','武宁','永修','德安','都昌','湖口','彭泽','八里湖']
var areaRegx = ['/浔阳/g','/濂溪/g','/柴桑/g','/瑞昌/g','/共青城/g','/庐山/g','/修水/g','/武宁/g','/永修/g','/德安/g','/都昌/g','/湖口/g','/彭泽/g','/八里湖/g']
//设置跨域访问
app.all("*",function(req,res,next){
  //设置允许跨域的域名，*代表允许任意域名跨域
  res.header("Access-Control-Allow-Origin","*");
  //允许的header类型
  res.header("Access-Control-Allow-Headers","content-type");
  //跨域允许的请求方式 
  res.header("Access-Control-Allow-Methods","DELETE,PUT,POST,GET,OPTIONS");
  if (req.method.toLowerCase() == 'options')
      res.send(200);  //让options尝试请求快速结束
  else
      next();
})
const folderName = process.env.HOME || process.env.USERPROFILE + '\\Desktop\\police'
//! !用于结束后清空文件夹
// var files = fs.readdirSync(folderName)
// files.forEach((file) => {
//   fs.unlink(folderName + '\\' + file, (err) => {
//     if (err) throw err
//   })
// })
app.get('/', function(req, res) {
  var params = req.query.data;
  console.log(params);
  getMonth.getMonth(params).then((resp)=>{

    console.log(resp,"this is backdata");
    res.status(200).json({
      data : {
          message : resp
     }
  });
  }).catch(err=>{
    res.status(500).json({
      data : {
          message : err
     }
  });
  })
})
//文件导入
app.post("/importExcel",  function (req, res) {
  /* 生成multiparty对象，并配置上传目标路径 */
  let form = new multiparty.Form();
  // 设置编码
  form.encoding = 'utf-8';
  // 设置文件存储路径，以当前编辑的文件为相对路径
//   form.uploadDir = './images';
  form.uploadDir = folderName
  // 设置文件大小限制
  // form.maxFilesSize = 1 * 1024 * 1024;
  form.parse(req, function (err, fields, files) {
    try {
      let inputFile = files.file[0];
      let newPath = form.uploadDir + "/" + inputFile.originalFilename;
      // 同步重命名文件名 fs.renameSync(oldPath, newPath)
　　　 //oldPath  不得作更改，使用默认上传路径就好
      fs.renameSync(inputFile.path, newPath);
      res.status(200).json({
        data : {
            message : '成功'
       }
    });
    } catch (err) {
      console.log(err);
      res.status(404).json({
        data : {
            message : err.message
       }
    });
    };
  })
});

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

// getMonth();
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
