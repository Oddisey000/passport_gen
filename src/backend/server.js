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
console.log(mdbDocument)

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

app.get('/', (req, res) => {
  mdbFunctions.TakeEquipmentInfo(mdbDocument)
    .then(value => {
      SendResponse(res,value)
    })
});

app.get('/form', (req, res) => {
  res.send(`<form ref='uploadForm' 
    id='uploadForm' 
    action='http://localhost:${port}/upload/' 
    method='post' 
    encType="multipart/form-data">
      <input type="file" name="sampleFile" />
      <input type='submit' value='Upload!' />
  </form>`)
});

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
    res.status(200).send({
      message: "File uploaded successfuly",
      code: 200
    })
  })
});

app.get('/mdb', (req, res) => {

  if (mdbDocument) {
    const SendResponse = (res,value) => {
      excelDocument = `.\\export\\${value.completeData[0].tableName}__${excelFunctions.convertDate(new Date())}.xlsx`
      res.send('<a href="/download">Download</a>')
    }
  
    mdbFunctions.CreatePassportData(mdbDocument)
    .then(value => {
      excelFunctions.TestEquipmentPassport(value.completeData, value.switchName)
      SendResponse(res,value)
    })
  } else {
    res.send('No mdb file uploaded')
  }
});

app.get('/download', (req, res) => {
  res.download(path.join(__dirname, excelDocument))
});