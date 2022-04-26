// Import required libraries
const express = require('express');
const cors = require('cors')
const ADODB = require('node-adodb');
const Excel = require('exceljs');
const excelHeader = require('./header.json');

// Initialize variables required for working with data
const connection = ADODB.open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\\mdb\\BR206.mdb;');
const workbook = new Excel.Workbook();
const app = express();
const port = process.env.PORT || 3200;

// Function needs to run express server
/*app.listen(port, () => {
  console.log(`App is listening at http://localhost:${port}`);
});*/

// Resolve any CORS issue that may be encountered
app.use(cors());

//app.get('/', (req, res) => {
  CreateDBObject()
  async function CreateDBObject() {
    try {
      let database = [];
      let switchName = [];
      let completeData = [];
      let counter;
      // Select all required data from key DB tables
      database.push(await connection.query('SELECT ModuleTableID, TableName FROM ModuleTemplate_Table ORDER BY TableName'))
      database.push(await connection.query('SELECT ModuleTableID, ModuleID, ConnectorCode, XCode, ModuleID, ModuleNumber, CoordX, CoordY FROM ModuleTemplate_Modules ORDER BY ModuleTableID, ModuleID'))
      database.push(await connection.query('SELECT ModuleTableID, ModuleID, SwitchName FROM ModuleTemplate_Switch'))
      database.push(await connection.query('SELECT ModuleTableID, ModuleID, NumberPins, Pushback FROM ModuleTemplate_Pins'))

      // Push into array only unique names before first delimiter "_"
      database[2].map((value) => {
        value.SwitchName = value.SwitchName.split('_')[0]
        if (switchName.indexOf(value.SwitchName) === -1) switchName.push(value.SwitchName)
      });

      const CalculateSwitches = (switchesArray, equipmentID, moduleID) => {
        switchAmount = []
        switchesArray.map(element => {
          counter = 0
          database[2].map(value => {
            if (value.ModuleTableID == equipmentID && value.ModuleID == moduleID && value.SwitchName == element) {
              counter = counter + 1
            }
          })
          switchesData = {
            switchName: element,
            switchAmount: counter
          }
          switchAmount.push(switchesData)
        })
        return switchAmount
      }

      database[0].map(table_value => {
        database[1].map(module_value => {
          if (table_value.ModuleTableID == module_value.ModuleTableID) {
            database[3].map(pin_value => {
              if (module_value.ModuleID == pin_value.ModuleID) {
                numberPins = pin_value.NumberPins
                numberPushback = pin_value.Pushback
              }
            });
            objectToPush = {
              tableName: table_value.TableName,
              moduleData: {
                xcode: module_value.XCode,
                connectorCode: module_value.ConnectorCode.split('<')[0],
                moduleNumber: module_value.ModuleNumber.split('&')[0],
                checkDimensions: module_value.ModuleNumber.split('&')[1],
                checkStability: module_value.ModuleNumber.split('&')[2],
                checkSplitDimensions: module_value.ModuleNumber.split('&')[3],
                checkTightness: module_value.ModuleNumber.split('&')[4],
                coordX: module_value.CoordX,
                coordY: module_value.CoordY,
                numberPins: numberPins,
                pushback: numberPushback,
                switches: CalculateSwitches(switchName, module_value.ModuleTableID, module_value.ModuleID)
              }
            }
            completeData.push(objectToPush)
          }
        })
      });
      //console.log(completeData[0].moduleData)
      ExtractDataToExcel(completeData, switchName)
      //res.send(completeData[0]);
    } catch (error) {
      console.error(error);
    }
  }
//});

const ExtractDataToExcel = (completeData, switchName) => {
  const sheet = workbook.addWorksheet('Pasport');
  let worksheet = workbook.getWorksheet(1);
  let startDataRow = 6
  let startHeaderSwitches = 11
  let NumCounter = 1

  let row = worksheet.getRow(5)


  row.getCell(1).value = excelHeader.passportHeader.positionColumn
  row.getCell(2).value = excelHeader.passportHeader.incommingInspectionCheck
  row.getCell(3).value = excelHeader.passportHeader.checkingDate
  row.getCell(4).value = excelHeader.passportHeader.inspectorNamePP
  row.getCell(5).value = excelHeader.passportHeader.signaturePP
  row.getCell(6).value = excelHeader.passportHeader.xcode
  row.getCell(7).value = excelHeader.passportHeader.moduleName
  row.getCell(8).value = excelHeader.passportHeader.index
  row.getCell(9).value = 'Electric pins'
  row.getCell(10).value = 'Pushback function'

  row.getCell(11 + completeData[0].moduleData.switches.length).value = excelHeader.passportHeader.checkDimensions
  row.getCell(12 + completeData[0].moduleData.switches.length).value = excelHeader.passportHeader.checkStability
  row.getCell(13 + completeData[0].moduleData.switches.length).value = excelHeader.passportHeader.checkSplitDimensions
  row.getCell(14 + completeData[0].moduleData.switches.length).value = excelHeader.passportHeader.checkTightness
  row.getCell(15 + completeData[0].moduleData.switches.length).value = excelHeader.passportHeader.inspectorNameQM
  row.getCell(16 + completeData[0].moduleData.switches.length).value = excelHeader.passportHeader.signatureQM

  for (let i = 11; i < 15; i++) {
    row.getCell(i + completeData[0].moduleData.switches.length).font = { name: 'Arial', size: 8 }
    row.getCell(i + completeData[0].moduleData.switches.length).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
  }

  for (let i = 15; i <17; i++) { 
    row.getCell(i + completeData[0].moduleData.switches.length).font = { name: 'Times New Roman', size: 10, bold: true }
    row.getCell(i + completeData[0].moduleData.switches.length).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
  }

  for (let i = 11; i < 17; i++) {
    row.getCell(i + completeData[0].moduleData.switches.length).alignment = { vertical: 'middle', horizontal: 'center', textRotation: 90, wrapText: true }
    row.getCell(i + completeData[0].moduleData.switches.length).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
  }

  for (let i = 1; i < 9; i++) {
    row.getCell(i).alignment = { vertical: 'middle', horizontal: 'center', textRotation: 90, wrapText: true }
    row.getCell(i).font = { name: 'Times New Roman', size: 10, bold: true }
    row.getCell(i).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
  }

  for (let i = 9; i < 11; i++) {
    row.getCell(i).alignment = { textRotation: 90 }
    row.getCell(i).font = { name: 'Arial', size: 9 }
    row.getCell(i).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
  }

  row.getCell(2).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'BFBFBF'} }
  
  worksheet.getColumn(1).width = 4
  worksheet.getColumn(2).width = 8.14
  worksheet.getColumn(3).width = 8.85
  worksheet.getColumn(4).width = 15
  worksheet.getColumn(5).width = 10
  worksheet.getColumn(6).width = 17
  worksheet.getColumn(7).width = 20
  worksheet.getColumn(8).width = 20
  worksheet.getColumn(9).width = 3.15
  worksheet.getColumn(10).width = 3.15
  
  row.height = 147
  row.commit()

  switchName.map((value) => {
    let row = worksheet.getRow(5)
    worksheet.getColumn(startHeaderSwitches).width = 3.15

    row.getCell(startHeaderSwitches).alignment = { textRotation: 90 }
    row.getCell(startHeaderSwitches).font = { name: 'Arial', size: 9 }
    row.getCell(startHeaderSwitches).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
    row.getCell(startHeaderSwitches).value = value
    row.commit()
    startHeaderSwitches = startHeaderSwitches + 1
  })
  completeData.map((value) => {
    let row = worksheet.getRow(startDataRow)
    startHeaderSwitches = 11

    row.getCell(3).font = { name: 'Arial', size: 9 }
    row.getCell(3).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
    
    for (let i = 6; i < 8; i++) {
      row.getCell(i).font = { name: 'Arial', size: 9 }
      row.getCell(i).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
    }

    row.getCell(1).value = NumCounter
    row.getCell(3).value = new Date()
    row.getCell(4).value = "Перцович В.В."
    row.getCell(6).value = value.moduleData.xcode
    row.getCell(7).value = value.moduleData.connectorCode
    row.getCell(8).value = value.moduleData.moduleNumber
    row.getCell(9).value = value.moduleData.numberPins === 0 ? '' : value.moduleData.numberPins
    row.getCell(10).value = value.moduleData.pushback === 0 ? '' : '+'

    for (let i = 1; i < 11; i++) {
      row.getCell(i).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
    }
    
    value.moduleData.switches.map((element) => {
      row.getCell(startHeaderSwitches).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
      element.switchAmount === 0 ? '' : row.getCell(startHeaderSwitches).value = element.switchAmount
      startHeaderSwitches = startHeaderSwitches + 1
    })

    row.getCell(11 + value.moduleData.switches.length).value = value.moduleData.checkDimensions
    row.getCell(12 + value.moduleData.switches.length).value = value.moduleData.checkStability
    row.getCell(13 + value.moduleData.switches.length).value = value.moduleData.checkSplitDimensions
    row.getCell(14 + value.moduleData.switches.length).value = value.moduleData.checkTightness

    for (let i = 11; i < 17; i++) {
      row.getCell(i + value.moduleData.switches.length).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
    }
    row.commit()

    startDataRow = startDataRow + 1
    NumCounter = NumCounter + 1
  })
  return workbook.xlsx.writeFile('.\\mdb\\new.xlsx');
};