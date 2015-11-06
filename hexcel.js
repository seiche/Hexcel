#!/usr/bin/env node
var fs = require("fs");
var Excel = require("exceljs");

var filename = process.argv[2];
var buffer = fs.readFileSync(filename);
var out_file = filename.split(".");
if(out_file.length > 1){
  out_file.pop();
}
out_file = out_file.join("") + ".xlsx";

var workbook = new Excel.Workbook();
var sheet = workbook.addWorksheet("My Sheet");
sheet.columns = [
  { header: "Offset", width: 16 },
  { header: "0x00", width: 16 },
  { header: "0x04", width: 16 },
  { header: "0x08", width: 16 },
  { header: "0x0C", width: 16 },
  { header: "Ascii", width: 24 },
];
var row = sheet.getRow(1);
row.eachCell({}, function(cell, colNumber) {
  cell.font = {
    name: "Courier New",
    size: 10,
    bold: true
  };
  cell.alignment = { horizontal: "center" };
  cell.fill = {
    type: "pattern",
    pattern:"fill",
    bgColor : {argb:"FF9FC5E8"}
  }
});

for(var i = 0; i < buffer.length - 0x10; i += 0x10){
  var num = i.toString(16);
  while(num.length < 6){
    num = "0" + num;
  }
  num = "0x" + num;

  var str_buf = new Buffer(16);
  buffer.copy(str_buf, 0, i, i + 0x10);
  for(var k = 0; k < str_buf.length; k++){
    if(str_buf[k] < 33 || str_buf[k] > 126){
      str_buf[k] = 46;
    }
  }

  var row = sheet.addRow([
    num,
    buffer.toString("hex", i, i + 0x04),
    buffer.toString("hex", i + 0x04, i + 0x08),
    buffer.toString("hex", i + 0x08, i + 0x0C),
    buffer.toString("hex", i + 0x0C, i + 0x10),
    str_buf.toString("ascii"),
  ]);

  row.eachCell({}, function(cell, colNumber) {
    cell.font = {
      name: "Courier New",
      size: 10,
    };
    if(colNumber == 1){
      cell.alignment = { horizontal: "right" };
    }else{
      cell.alignment = { horizontal: "center" };
    }
  });

}

workbook.xlsx.writeFile(out_file).then(function(){
  console.log(out_file + " written");
});
