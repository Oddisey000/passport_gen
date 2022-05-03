import React from "react";
import axios from "axios";
import './initial.form.component.css'

const InitialFormComponent = () => {
  const [file, setFile] = React.useState()
  const [fileName, setFileName] = React.useState()
  const [dbSelectors, setdbSelectors] = React.useState()
  let equipmentID

  const DisplayAdditionalElements = () => {
    document.getElementById("user-data").style.display = "block"
    document.getElementById("equipment-data").style.display = "block"
    document.getElementById("generate-btn").style.display = "block"
  };

  const SaveFile = (e) => {
    setFile(e.target.files[0])
    setFileName(e.target.files[0].name)
    setTimeout(() => {
      document.getElementById("btn-click").click()
    }, 500);
  };
  const UploadFile = async(e) => {
    const formData = new FormData()
    formData.append("file", file)
    formData.append("fileName", fileName)
    try {
      const res = await axios.post(
        "http://localhost:3200/upload", formData
      )
      setdbSelectors(res.data.data)
      DisplayAdditionalElements()
    } catch(error) {
      console.log(error)
    }
  };

  const HandleChange = (e) => {
    equipmentID = ''
    dbSelectors.map(value => 
      value.map(element => {
        if (e.target.value == element.TableName) {equipmentID = element.ModuleTableID}
      })
    )
  }
  const SubmitData = async(e) => {
    const userName = document.getElementById("user_name").value
    const equipmentName = document.getElementById("equipment_name").value
    if (userName !== '' && equipmentName !== '' && equipmentID !== '') {
      const dataObj = {
        userName: userName,
        equipmentName: equipmentName,
        equipmentID: equipmentID
      }
      SendRequest("http://localhost:3200/getDocument", dataObj)
    }
  }
  const SendRequest = (APIrequest, dataToSend) => {
    fetch(APIrequest + '?senddata=' + JSON.stringify(dataToSend))
    .then(res => {
      console.log(res);
      document.getElementById("generate-btn").style.display = "none"
      document.getElementById("download-btn").style.display = "block"
    })
    .catch(error => {
      //console.log(error);
    })
  };
  return (
    <div id="initial-form">
      <div className="form-style-5">
         <fieldset>
          <legend><span className="number">1</span>База даних</legend>
            <input type="file" accept=".mdb" onChange={SaveFile} />
            <button id="btn-click" style={{display: "none"}} onClick={UploadFile}></button>
          </fieldset>
        <form>
        <fieldset>
          <div id="user-data" style={{display: "none"}}>
            <legend><span className="number">2</span>Дані користувача</legend>
            <input type="text" id="user_name" required placeholder="Прізвище та ініціали  *" />
            <input type="text" id="equipment_name" required placeholder="Назва контрольного столу *" />
          </div>
          <div id="equipment-data" style={{display: "none"}}>
            <legend><span className="number">3</span>Вибір обладнання</legend>
              <select name="equipment_array" onChange={HandleChange}>
                <option></option>
                {dbSelectors ? dbSelectors.map(value => 
                  value.map(element => <option id={element.ModuleTableID} key={element.ModuleTableID}>{element.TableName}</option>)
                ) : ""}
              </select>
          </div>
          </fieldset>
          <fieldset style={{display: "none"}}>
          <legend><span className="number">4</span>Додатково згенерувати</legend>
            <div className="additional-checkboxes">
              <input type="checkbox" id="vehicle1" name="optional_one" value="AA 3211-25 дод.79"/>
              <label> AA 3211-25 дод.79</label><br/>
            </div>
            <div className="additional-checkboxes">
              <input type="checkbox" id="vehicle1" name="optional_two" value="AA 3211-25 дод.67"/>
              <label> AA 3211-25 дод.67</label><br/>
            </div>
          </fieldset>
          <div id="btn-block">
            <a style={{display: "none"}} id="download-btn" type="button" href="http://localhost:3200/download">Завантажити</a>
            <input style={{display: "none"}} type="button" id="generate-btn" value="Згенерувати" onClick={SubmitData} />
          </div>
        </form>
      </div>
    </div>
  )
}

export default InitialFormComponent;