// Import required libraries
const express = require('express')
const cors = require('cors')
const fs = require('fs');
const fileUpload = require("express-fileupload")
const path = require('path')
const bodyParser = require('body-parser')

const excelFunctions = require('./excel/excel')
const mdbFunctions = require('./mdb/mdb');

let excelDocument;

const getMostRecentFile = (dir) => {
  const files = orderReccentFiles(dir);
  return files.length ? files[0] : undefined;
};

const orderReccentFiles = (dir) => {
  return fs.readdirSync(dir)
      .filter(file => fs.lstatSync(path.join(dir, file)).isFile())
      .map(file => ({ file, mtime: fs.lstatSync(path.join(dir, file)).mtime }))
      .sort((a, b) => b.mtime.getTime() - a.mtime.getTime());
};

let mdbDocument = getMostRecentFile(__dirname + '/upload/') ? getMostRecentFile(__dirname + '/upload/').file : ''

// Initialize variables required for working with data
const app = express();
const port = process.env.PORT || 3200;

// Function needs to run express server
app.listen(port, () => {
  console.log(`App is listening at http://localhost:${port}`);
});

// Resolve any CORS issue that may be encountered
app.use(cors());
app.use(fileUpload());
app.use(express.static("files"))
app.use(bodyParser.json())
app.use(bodyParser.urlencoded({extended: true}))

app.post("/upload", (req, res) => {
  const newPath = __dirname + "/upload/"
  const file = req.files.file
  const fileName = file.name

  file.mv(`${newPath}${fileName}`, (err) => {
    if(err) {
      res.status(500).send({
        message: "File upload failed",
        code: 200
      })
    }
    mdbFunctions.TakeEquipmentInfo(mdbDocument)
    .then(value => {
      res.status(200).send({
        message: "Send data back to frontend",
        code: 200,
        data: value
      })
    })
  })
});

app.get('/getDocument', (req, res) => {
  if (mdbDocument) {
    const reqData = JSON.parse(req.query.senddata)
    const userName = reqData.userName
    const equipmentName = reqData.equipmentName
    const equipmentID = reqData.equipmentID
    
    const SendResponse = (res,value) => {
      excelDocument = `.\\export\\${value.completeData[0].tableName}__${excelFunctions.convertDate(new Date())}.xlsx`
      //res.send('<a href="/download">Download</a>')
      res.status(200).send({
        message: "Ready to download",
        code: 200
      })
    }
  
    mdbFunctions.CreatePassportData(mdbDocument, equipmentID)
    .then(value => {
      excelFunctions.TestEquipmentPassport(value.completeData, value.switchName, userName, equipmentName)
      SendResponse(res,value)
    })
  } else {
    res.send('No mdb file uploaded')
  }
});

app.get('/download', (req, res) => {
  res.download(path.join(__dirname, excelDocument))
});