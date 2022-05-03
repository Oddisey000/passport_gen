const ADODB = require('node-adodb');

async function CreatePassportData(mdbDocument, equipmentID) {
  const connection = ADODB.open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\\upload\\' + mdbDocument +';')
  try {
    let completeData = [];
    let database = [];
    let switchName = [];
    let counter;
    // Select all required data from key DB tables
    database.push(await connection.query('SELECT ModuleTableID, TableName FROM ModuleTemplate_Table WHERE ModuleTableID =' + equipmentID + ' ORDER BY TableName'))
    database.push(await connection.query('SELECT ModuleTableID, ModuleID, ConnectorCode, XCode, ModuleID, ModuleNumber, CoordX, CoordY FROM ModuleTemplate_Modules WHERE ModuleTableID =' + equipmentID + ' ORDER BY ModuleTableID, ModuleID'))
    database.push(await connection.query('SELECT ModuleTableID, ModuleID, SwitchName FROM ModuleTemplate_Switch WHERE ModuleTableID =' + equipmentID))
    database.push(await connection.query('SELECT ModuleTableID, ModuleID, NumberPins, Pushback FROM ModuleTemplate_Pins WHERE ModuleTableID =' + equipmentID))

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

    const CalculatePins = (moduleID, moduleTableID) => {
      let pins
      database[3].map((value) => {
        if (moduleID == value.ModuleID && moduleTableID == value.ModuleTableID) {
          pins = value.NumberPins
        }
      })
      return pins
    }

    const CalculatePushbacks = (moduleID, moduleTableID) => {
      let pushback
      database[3].map((value) => {
        if (moduleID == value.ModuleID && moduleTableID == value.ModuleTableID) {
          pushback = value.Pushback
        }
      })
      return pushback
    }

    database[0].map(table_value => {
      database[1].map(module_value => {
        if (table_value.ModuleTableID == module_value.ModuleTableID) {
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
              numberPins: CalculatePins(module_value.ModuleID, module_value.ModuleTableID),
              pushback: CalculatePushbacks(module_value.ModuleID, module_value.ModuleTableID),
              switches: CalculateSwitches(switchName, module_value.ModuleTableID, module_value.ModuleID)
            }
          }
          completeData.push(objectToPush)
        }
      })
    });
    let obj = {
      completeData: completeData,
      switchName: switchName
    }
    return obj
  } catch (error) {
    console.error(error);
  }
}

async function TakeEquipmentInfo(mdbDocument) {
  const connection = ADODB.open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\\upload\\' + mdbDocument +';')
  try {
    let database = [];
    database.push(await connection.query('SELECT ModuleTableID, TableName FROM ModuleTemplate_Table ORDER BY TableName'))
    return database
  } catch(error) {
    console.error(error)
  }
}

module.exports = {
  CreatePassportData,
  TakeEquipmentInfo
};