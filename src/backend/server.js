// Import required libraries
const express = require('express')
const cors = require('cors')
const fileUpload = require("express-fileupload");
const path = require('path')

const excelFunctions = require('./excel/excel')
const mdbFunctions = require('./mdb/mdb');

let excelDocument = '';


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

app.get('/', (req, res) => {
  res.send(`<form ref='uploadForm' 
    id='uploadForm' 
    action='http://localhost:${port}/upload/' 
    method='post' 
    encType="multipart/form-data">
      <input type="file" name="sampleFile" />
      <input type='submit' value='Upload!' />
  </form>`)
})

app.post('/upload', function(req, res) {
  let sampleFile;
  let uploadPath;

  if (!req.files || Object.keys(req.files).length === 0) {
    return res.status(400).send('No files were uploaded.');
  }

  // The name of the input field (i.e. "sampleFile") is used to retrieve the uploaded file
  sampleFile = req.files.sampleFile;
  uploadPath = __dirname + '/upload/' + sampleFile.name;

  // Use the mv() method to place the file somewhere on your server
  sampleFile.mv(uploadPath, function(err) {
    if (err)
      return res.status(500).send(err);

    res.send('File uploaded!');
  });
});

app.get('/something', (req, res) => {
  mdbFunctions.CreateDBObject()
  .then(value => {
    excelFunctions.TestEquipmentPassport(value.completeData, value.switchName)
    excelDocument = `.\\export\\${value.completeData[0].tableName}__${excelFunctions.convertDate(new Date())}.xlsx`
  })
  setTimeout(() => {
    res.send('<a href="/download">Download</a>')
  }, 1000);
});

app.get('/download', (req, res) => {
  res.download(path.join(__dirname, excelDocument))
});