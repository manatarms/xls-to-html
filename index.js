#!/usr/bin/env node
if (typeof require !== "undefined") XLSX = require("xlsx");
const fs = require("fs");
const handlebars = require("handlebars");

let workbook = XLSX.readFile("./assets/HDFC_CC_work_master_052318.xlsx");
let sheets = workbook.SheetNames;
let headers = [], preheaders = [];

sheets.forEach(sheet => {
  Object.keys(workbook.Sheets[sheet]).forEach(cell => {
    if (cell > "B5" && cell < "C1") {
      headers.push(workbook.Sheets[sheet][cell].w);
    }
  });
});

console.log(headers.length);
let htmlSource = fs.readFileSync(
  __dirname + "/assets/ING_April18_CC_EE_MoneyBack_Mailer_MoneyBack_3C_CG.htm",
  "utf-8"
);

let htmlTemplate = handlebars.compile(htmlSource);

let data = {
  SUBJECT_LINE: headers[0]
};

let outputHtml = htmlTemplate(data);

let outputDir = `${__dirname}/output/`;

if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir);
}

fs.writeFile(outputDir + "SP1.html", outputHtml, err => {
  if (err) throw err;
  console.log("The file has been saved!");
});
