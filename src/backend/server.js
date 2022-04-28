// Import required libraries
const express = require('express')
const fileUpload = require("express-fileupload");
const cors = require('cors')
const ADODB = require('node-adodb')
const path = require('path')
const excelFunctions = require('./excel/excel')

let completeData = [];


// Initialize variables required for working with data
const connection = ADODB.open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\\mdb\\BR206.mdb;');
const app = express();
const port = process.env.PORT || 3200;

// Function needs to run express server
app.listen(port, () => {
  console.log(`App is listening at http://localhost:${port}`);
});

// Resolve any CORS issue that may be encountered
app.use(cors());
app.use(fileUpload());

app.get('/', (req, res) => {
  res.send(`<form ref='uploadForm' 
    id='uploadForm' 
    action='http://localhost:${port}/mdb/' 
    method='post' 
    encType="multipart/form-data">
      <input type="file" name="sampleFile" />
      <input type='submit' value='Upload!' />
  </form>`)
})

app.post('/mdb', function(req, res) {
  let sampleFile;
  let uploadPath;

  if (!req.files || Object.keys(req.files).length === 0) {
    return res.status(400).send('No files were uploaded.');
  }

  // The name of the input field (i.e. "sampleFile") is used to retrieve the uploaded file
  sampleFile = req.files.sampleFile;
  uploadPath = __dirname + '/mdb/' + sampleFile.name;

  // Use the mv() method to place the file somewhere on your server
  sampleFile.mv(uploadPath, function(err) {
    if (err)
      return res.status(500).send(err);

    res.send('File uploaded!');
  });
});

app.get('/something', (req, res) => {
  CreateDBObject()
  async function CreateDBObject() {
    try {
      let database = [];
      let switchName = [];
      let counter;
      // Select all required data from key DB tables
      database.push(await connection.query('SELECT ModuleTableID, TableName FROM ModuleTemplate_Table WHERE ModuleTableID = 138604445 ORDER BY TableName'))
      database.push(await connection.query('SELECT ModuleTableID, ModuleID, ConnectorCode, XCode, ModuleID, ModuleNumber, CoordX, CoordY FROM ModuleTemplate_Modules WHERE ModuleTableID = 138604445 ORDER BY ModuleTableID, ModuleID'))
      database.push(await connection.query('SELECT ModuleTableID, ModuleID, SwitchName FROM ModuleTemplate_Switch WHERE ModuleTableID = 138604445'))
      database.push(await connection.query('SELECT ModuleTableID, ModuleID, NumberPins, Pushback FROM ModuleTemplate_Pins WHERE ModuleTableID = 138604445'))

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
      excelFunctions.TestEquipmentPassport(completeData, switchName)
      res.send('<a href="/download">Download</a>');

    } catch (error) {
      console.error(error);
    }
  }
});

app.get('/download', (req, res) => {
  res.sendFile(path.join(__dirname, `.\\ready\\${completeData[0].tableName + '__' + excelFunctions.convertDate(new Date())}.xlsx`))
});