const ADODB = require('node-adodb');
var Excel = require('exceljs');
const connection = ADODB.open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\\mdb\\BR206.mdb;');
var workbook = new Excel.Workbook();
let switchName = [];
 
async function query() {
  try {
    const users = await connection.query('SELECT * FROM ModuleTemplate_Switch');
 
    JSON.stringify(users, null, 2);
    users.map((value) => {
      if (switchName.indexOf(value.SwitchName.split('_')[0]) === -1) switchName.push(value.SwitchName.split('_')[0])
    });
  } catch (error) {
    console.error(error);
  }
  console.log(switchName.map((value) => value));
}

query();

workbook.xlsx.readFile('.\\mdb\\template.xlsx')
  .then(function() {
      var worksheet = workbook.getWorksheet(1);
      var row = worksheet.getRow(6);
      row.getCell(2).value = 5; // A5's value set to 5
      row.commit();
      //var lastColumn = workbook.getWorksheet(1).lastColumn;
      //console.log(lastColumn.number);
      return workbook.xlsx.writeFile('.\\mdb\\new.xlsx');
  })