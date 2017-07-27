var express = require('express');
var router = express.Router();
var xlsx = require('xlsx');
var fs = require('fs-extra');
var path = require('path');
var multer = require('multer');
var excelJs = require('exceljs');

var storge = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, './public/files')
    },
    filename: function (req, file, cb) {
        var fileformat = (file.originalname).split('.');
        cb(null, file.fieldname + '-' + Date.now() + '.' + fileformat[fileformat.length - 1]);
    }
});
var upload = multer({storage: storge});

/* GET home page. */
router.get('/', function (req, res, next) {
    res.render('index', {title: 'Express'});
});

router.post('/upload/excel', upload.single('excel'), function (req, res, next) {
    console.log(req.body);
    if (req.file) {
        console.log(req.file);
        var fileName = req.file.filename.split('.')[0];
        var fileExecPath = path.join(__dirname, '../' + req.file.path);
        var workbook = xlsx.readFile(fileExecPath);
        var sheetName = workbook.SheetNames;
        var data = sheetName[0];
        sheetName.forEach(function (item) {
            if (item == req.body.tableName) {
                data = item;
            }
        });
        var workSheet = workbook.Sheets[data];
        var jsonTxt = xlsx.utils.sheet_to_json(workSheet, {raw: true, defval: ""});
        var fileNameJson = path.join(__dirname, '../public/json/' + fileName + '-' + data + '.json');
        fs.writeJson(fileNameJson, jsonTxt, {spaces: "\t"}, function (err) {
            if (err) {
                console.log(err);
            }
            var filePath = fileNameJson;
            var stats = fs.statSync(filePath);
            if (stats.isFile()) {
                res.set({
                    'Content-Type': 'application/octet-stream',
                    'Content-Disposition': 'attachment; filename=' + fileName + '-' + data + '.json',
                    'Content-Length': stats.size
                });
                fs.createReadStream(filePath).pipe(res);
                fs.remove(fileExecPath, function (err) {
                    if (err)console.log(err);
                });
                fs.remove(fileNameJson, function (err) {
                    if (err) {
                        console.log(err)
                    }
                })
            } else {
                res.end(404);
            }
        })
    }
});

router.post('/upload/json', upload.single('json'), function (req, res, next) {
    if (req.file) {
        var fileName = req.file.filename.split('.')[0];
        var data = JSON.parse(fs.readFileSync(req.file.path));
        var columns = Object.keys(data[0]);

        var workbook = new excelJs.stream.xlsx.WorkbookWriter({
            filename: path.join(__dirname, '../public/xlsx/' + fileName + '.xlsx')
        });
        var worksheet = workbook.addWorksheet('Sheet');
        worksheet.properties.rowCount = data.length;
        var items = [];
        columns.forEach(function (item) {
            items.push({header: item, key: item})
        });
        worksheet.columns = items;
        for (var i in data) {
            worksheet.addRow(data[i]).commit();
        }
        workbook.commit();
        var filePath = path.join(__dirname, '../public/xlsx/' + fileName + '.xlsx');
        var stats = fs.statSync(filePath);
        if (stats.isFile()) {
            res.set({
                'Content-Type': 'application/vnd.ms-excel',
                'Content-Disposition': 'attachment; filename=' + fileName + '.xlsx'
            });
            fs.createReadStream(filePath).pipe(res);
            fs.remove(path.join(__dirname, '../' + req.file.path), function (err) {
                if (err) console.log(err);
            });
            fs.remove(filePath, function (err) {
                if (err) {
                    console.log(err)
                }
            })
        } else {
            res.end(404);
        }
    }
});

module.exports = router;
