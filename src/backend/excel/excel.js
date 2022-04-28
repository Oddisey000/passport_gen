const Excel = require('exceljs');
const workbook = new Excel.Workbook();

const passportHeader = require('./passport_header.json');

function convertDate(date) {
  var yyyy = date.getFullYear().toString();
  var mm = (date.getMonth()+1).toString();
  var dd  = date.getDate().toString();
  var mmChars = mm.split('');
  var ddChars = dd.split('');
  return (ddChars[1]?dd:"0"+ddChars[0]) + '-' + (mmChars[1]?mm:"0"+mmChars[0]) + '-' + yyyy;
};

const TestEquipmentPassport = (completeData, switchName) => {
  workbook.removeWorksheet(1)
  let sheet = workbook.addWorksheet('Pasport');
  let worksheet = workbook.getWorksheet(1);
  let startDataRow = 6
  let startHeaderSwitches = 11
  let NumCounter = 1

  let row = worksheet.getRow(1)
    row.getCell(1).value = passportHeader.passportHeader.documentRevision
    row.getCell(1).font = { name: 'Times New Roman', size: 9 }
    row.getCell(1).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    row.getCell(1).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
  
    row.getCell(5).value = passportHeader.passportHeader.documentTitle
    row.getCell(5).font = { name: 'Times New Roman', size: 16, bold: true }
    row.getCell(5).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    row.getCell(5).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }

    row.getCell(9 + completeData[0].moduleData.switches.length + 7).value = passportHeader.passportHeader.documentLogo
    row.getCell(9 + completeData[0].moduleData.switches.length + 7).font = { name: 'Arial Black', size: 26, bold: true, color: { argb: '0000FF' } }
    row.getCell(9 + completeData[0].moduleData.switches.length + 7).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    row.getCell(9 + completeData[0].moduleData.switches.length + 7).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
  row.commit()

  row = worksheet.getRow(2)
    row.getCell(5).value = passportHeader.passportHeader.documentSubTitle
    row.getCell(5).font = { name: 'Times New Roman', size: 14, bold: true }
    row.getCell(5).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    row.getCell(5).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
  row.commit()

  row = worksheet.getRow(3)
    row.getCell(5).value = passportHeader.passportHeader.documentReminder
    row.getCell(5).font = { name: 'Times New Roman', size: 9 }
    row.getCell(5).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    row.getCell(5).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }

    row.getCell(9 + completeData[0].moduleData.switches.length + 7).value = passportHeader.passportHeader.documentLogoBelow
    row.getCell(9 + completeData[0].moduleData.switches.length + 7).font = { name: 'Times New Roman', size: 11, bold: true }
    row.getCell(9 + completeData[0].moduleData.switches.length + 7).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
    row.getCell(9 + completeData[0].moduleData.switches.length + 8).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
  row.commit()

  row = worksheet.getRow(4)
    row.getCell(1).value = passportHeader.passportHeader.positionColumn
    row.getCell(1).alignment = { vertical: 'middle', horizontal: 'center', textRotation: 90, wrapText: true }
    row.getCell(1).font = { name: 'Times New Roman', size: 10, bold: true }
    row.getCell(1).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }

    row.getCell(2).value = passportHeader.passportHeader.documentGroupOne
    row.getCell(2).font = { name: 'Times New Roman', size: 11, bold: true }
    row.getCell(2).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    row.getCell(2).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }

    row.getCell(6).value = passportHeader.passportHeader.documentGroupTwo
    row.getCell(6).font = { name: 'Times New Roman', size: 11, bold: true }
    row.getCell(6).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    row.getCell(6).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }

    row.getCell(9).value = passportHeader.passportHeader.documentGroupThree
    row.getCell(9).font = { name: 'Times New Roman', size: 11, bold: true }
    row.getCell(9).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    row.getCell(9).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }

    row.getCell(9 + completeData[0].moduleData.switches.length + 7).value = passportHeader.passportHeader.documentControllGroup
    row.getCell(9 + completeData[0].moduleData.switches.length + 7).font = { name: 'Times New Roman', size: 11, bold: true }
    row.getCell(9 + completeData[0].moduleData.switches.length + 7).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    row.getCell(9 + completeData[0].moduleData.switches.length + 7).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
  row.commit()

  row = worksheet.getRow(5)
  row.getCell(2).value = passportHeader.passportHeader.incommingInspectionCheck
  row.getCell(3).value = passportHeader.passportHeader.checkingDate
  row.getCell(4).value = passportHeader.passportHeader.inspectorNamePP
  row.getCell(5).value = passportHeader.passportHeader.signaturePP
  row.getCell(6).value = passportHeader.passportHeader.xcode
  row.getCell(7).value = passportHeader.passportHeader.moduleName
  row.getCell(8).value = passportHeader.passportHeader.index
  row.getCell(9).value = 'Electric pins'
  row.getCell(10).value = 'Pushback function'

  row.getCell(11 + completeData[0].moduleData.switches.length).value = passportHeader.passportHeader.checkDimensions
  row.getCell(12 + completeData[0].moduleData.switches.length).value = passportHeader.passportHeader.checkStability
  row.getCell(13 + completeData[0].moduleData.switches.length).value = passportHeader.passportHeader.checkSplitDimensions
  row.getCell(14 + completeData[0].moduleData.switches.length).value = passportHeader.passportHeader.checkTightness
  row.getCell(15 + completeData[0].moduleData.switches.length).value = passportHeader.passportHeader.colorDetection
  row.getCell(16 + completeData[0].moduleData.switches.length).value = passportHeader.passportHeader.inspectorNameQM
  row.getCell(17 + completeData[0].moduleData.switches.length).value = passportHeader.passportHeader.signatureQM

  for (let i = 11; i < 16; i++) {
    row.getCell(i + completeData[0].moduleData.switches.length).font = { name: 'Arial', size: 8 }
    row.getCell(i + completeData[0].moduleData.switches.length).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
  }

  for (let i = 16; i <18; i++) { 
    row.getCell(i + completeData[0].moduleData.switches.length).font = { name: 'Times New Roman', size: 10, bold: true }
    row.getCell(i + completeData[0].moduleData.switches.length).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
  }

  for (let i = 11; i < 18; i++) {
    row.getCell(i + completeData[0].moduleData.switches.length).alignment = { vertical: 'middle', horizontal: 'center', textRotation: 90, wrapText: true }
    row.getCell(i + completeData[0].moduleData.switches.length).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
  }

  for (let i = 2; i < 9; i++) {
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
  
  //StartRow, StartColumn, EndRow, EndColumn
  worksheet.mergeCells(1,5,1,9 + completeData[0].moduleData.switches.length + 6)
  worksheet.mergeCells(2,5,2,9 + completeData[0].moduleData.switches.length + 6)
  worksheet.mergeCells(3,5,3,9 + completeData[0].moduleData.switches.length + 6)
  worksheet.mergeCells(1,9 + completeData[0].moduleData.switches.length + 7,2,9 + completeData[0].moduleData.switches.length + 8)
  worksheet.mergeCells(4,1,5,1)
  worksheet.mergeCells(1,1,3,4)
  worksheet.mergeCells(4,2,4,5)
  worksheet.mergeCells(4,6,4,8)
  worksheet.mergeCells(4,9,4,9 + completeData[0].moduleData.switches.length + 6)
  worksheet.mergeCells(4,9 + completeData[0].moduleData.switches.length + 7,4,9 + completeData[0].moduleData.switches.length + 8)

  worksheet.getColumn(1).width = 4
  worksheet.getColumn(2).width = 8.14
  worksheet.getColumn(3).width = 8.85
  worksheet.getColumn(4).width = 15
  worksheet.getColumn(5).width = 10
  worksheet.getColumn(6).width = 17
  worksheet.getColumn(7).width = 20
  worksheet.getColumn(8).width = 20
  worksheet.getColumn(9).width = 3.86
  worksheet.getColumn(10).width = 3.15
  worksheet.getColumn(9 + completeData[0].moduleData.switches.length + 7).width = 9.94
  
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
    row.getCell(4).value = "Гнатюк Ю.А."
    row.getCell(6).value = value.moduleData.xcode
    row.getCell(7).value = value.moduleData.connectorCode
    row.getCell(8).value = value.moduleData.moduleNumber
    row.getCell(9).value = value.moduleData.numberPins ? value.moduleData.numberPins : 0
    row.getCell(10).value = value.moduleData.pushback > 0 ? '+' : ''

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

    for (let i = 11; i < 18; i++) {
      row.getCell(i + value.moduleData.switches.length).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
    }
    row.commit()

    startDataRow = startDataRow + 1
    NumCounter = NumCounter + 1
  })
  return workbook.xlsx.writeFile(`.\\ready\\${completeData[0].tableName + '__' + convertDate(new Date())}.xlsx`);
};

module.exports = {
  TestEquipmentPassport,
  convertDate
}