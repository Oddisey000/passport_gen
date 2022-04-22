const ADODB = require('node-adodb');
var Excel = require('exceljs');
const connection = ADODB.open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\\mdb\\BR206.mdb;');
var workbook = new Excel.Workbook();
let switchName = [];
let completeData = [];
let counter;
 
async function query() {
  try {
    // Select all required data from key DB tables
    const module_template_table = await connection.query('SELECT ModuleTableID, TableName FROM ModuleTemplate_Table ORDER BY TableName');
    JSON.stringify(module_template_table, null, 2);
    const module_template_modules = await connection.query('SELECT ModuleTableID, ModuleID, ConnectorCode, XCode, ModuleID, ModuleNumber, CoordX, CoordY FROM ModuleTemplate_Modules ORDER BY ModuleTableID, ModuleID');
    JSON.stringify(module_template_modules, null, 2);
    let module_template_switch = await connection.query('SELECT ModuleTableID, ModuleID, SwitchName FROM ModuleTemplate_Switch');
    JSON.stringify(module_template_switch, null, 2);
    const module_template_pins = await connection.query('SELECT ModuleTableID, ModuleID, NumberPins, Pushback FROM ModuleTemplate_Pins');
    JSON.stringify(module_template_pins, null, 2);

    // Push into array only unique names before first delimiter "_"
    module_template_switch.map((value) => {
      value.SwitchName = value.SwitchName.split('_')[0]
      if (switchName.indexOf(value.SwitchName) === -1) switchName.push(value.SwitchName)
    });

    const CalculateSwitches = (switchesArray, equipmentID, moduleID) => {
      switchAmount = []
      switchesArray.map(element => {
        counter = 0
        module_template_switch.map(value => {
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

    module_template_table.map(table_value => {
      module_template_modules.map(module_value => {
        if (table_value.ModuleTableID == module_value.ModuleTableID) {
          module_template_pins.map(pin_value => {
            if (module_value.ModuleID == pin_value.ModuleID) {
              numberPins = pin_value.NumberPins
              numberPushback = pin_value.Pushback
            }
          });
          objectToPush = {
            tableName: table_value.TableName,
            moduleData: {
              xcode: module_value.XCode,
              connectorCode: module_value.ConnectorCode,
              moduleNumber: module_value.ModuleNumber,
              coordX: module_value.CoordX,
              coordY: module_value.CoordY,
              numberPins: '',
              pushback: '',
              numberPins: numberPins,
              pushback: numberPushback,
              switches: CalculateSwitches(switchName, module_value.ModuleTableID, module_value.ModuleID)
            }
          }
          completeData.push(objectToPush)
        }
      })
    });
    console.log(completeData[0].moduleData)
  } catch (error) {
    console.error(error);
  }
}

query();

/** workbook.xlsx.readFile('.\\mdb\\template.xlsx')
  .then(function() {
      var worksheet = workbook.getWorksheet(1);
      var row = worksheet.getRow(6);
      row.getCell(2).value = 5; // A5's value set to 5
      row.commit();
      //var lastColumn = workbook.getWorksheet(1).lastColumn;
      //console.log(lastColumn.number);
      return workbook.xlsx.writeFile('.\\mdb\\new.xlsx');
  }) */