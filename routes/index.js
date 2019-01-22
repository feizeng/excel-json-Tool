var express = require('express');
var router = express.Router();
var formidable = require('formidable');
var fs = require('fs');
var os = require("os")
var path = require('path');
var xl = require('xlsx');
var jsonfile = require('jsonfile');
var _ = require('lodash');
var str = [];
var xlsx = require('node-xlsx');
/* GET home page. */
router.get('/', function(req, res, next) {
    res.render('form', { title: 'excel And json', layout: 'index' });
});

router.post('/excel', function(req, res, next) {
    let form = new formidable.IncomingForm();
    let uploaddir = 'Desktop/';
    let targetFile = path.join(os.homedir(), uploaddir);
    form.uploadDir = targetFile;
    form.parse(req, function(err, fields, files) {
        if (err) throw err;
        let oldpath = files.excel.path;
        let newpath = path.join(path.dirname(oldpath), "NEW_"+files.excel.name);
        if (files.excel.type.indexOf('sheet') !== -1) {
            fs.rename(oldpath, newpath, err => {
                if (err) throw err;
                let workbook = xl.readFile(newpath);
                const sheetNames = workbook.SheetNames;
                for (let num = 0; num < sheetNames.length; num++) {
                    let worksheet = workbook.Sheets[sheetNames[num]];
                    let dataa = xl.utils.sheet_to_json(worksheet);
                    let json_files = path.join(
                        os.homedir(),
                        'Desktop/excel' + num + '.json'
                    );
                    let data = creatJson(dataa);
                    fs.open(json_files, 'w', function(err) {
                        if (err) throw err;
                        jsonfile.writeFile(json_files, data, function(err) {
                            if (err) console.error(err);
                            fs.unlink(newpath,function(){
                                console.log("删除成功")
                            });
                        });
                    });
                }
            });
        }
        if (files.excel.type.indexOf('json') !== -1) {
            str = [];
            str.push(['KEYS', 'VALS']);
            jsonfile.readFile(oldpath, function(err, obj) {
                tableToExcel(obj, '');
                let num = 0;
                let excel_files = path.join(os.homedir(),
                    'Desktop/excel'+num+'.xlsx'
                );
                var buffer = xlsx.build([
                    {
                        name: 'sheet1',
                        data: str
                    }
                ]);
                fs.writeFileSync(excel_files, buffer, { flag: 'w' });
                fs.unlink(oldpath,function(){
                    console.log("删除成功")
                });
            });
        }
    });

    res.render('excel', {
        title: 'excel And json',
        result: '处理成功',
        layout: 'index'
    });
});

function tableToExcel(jsonData, keystr) {
    for (let item in jsonData) {
        if (typeof jsonData[item] === 'object') {
            if (keystr == '') {
                tableToExcel(jsonData[item], item);
            } else {
                tableToExcel(jsonData[item], keystr + '.' + item);
            }
        } else {
            let temparr = [];
            if(keystr == ''){
                temparr.push(`${item}`);
            }else{
                temparr.push(`${keystr}.${item}`);
            }
            temparr.push(jsonData[item]);
            str.push(temparr);
        }
    }
}
function creatJson(Arr) {
    let newjson = {};
    let lastjson = {};
    let alljson = {};
    for (let i = 0; i < Arr.length; i++) {
        for (let j = i; j < Arr.length; j++) {
            if (Arr[j + 1] == undefined) {
                newjson[Arr[i]['KEYS']] = Arr[i]['VALS'];
            } else {
                if (Arr[i]['KEYS'] === Arr[j + 1]['KEYS']) {
                    delete newjson[Arr[i]['KEYS']];
                    break;
                } else {
                    newjson[Arr[i]['KEYS']] = Arr[i]['VALS'];
                }
            }
        }
    }
    for (let item in newjson) {
        let vals = newjson[item];
        let itmArr = item.split('.');
        getjson(itmArr, vals, alljson);
        _.defaultsDeep(lastjson, alljson);
    }
    return lastjson;
}
function getjson(itemarr, val, argObj) {
    if (itemarr.length == 1) {
        argObj[itemarr[0]] = val;
    } else {
        let key = itemarr.shift();
        argObj[key] = {}; //{arr1:{}}
        getjson(itemarr, val, argObj[key]);
    }
}

module.exports = router;
