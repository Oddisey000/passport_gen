import React from "react";
import axios from "axios";
import './initial.form.component.css'

const InitialFormComponent = () => {
  const [file, setFile] = React.useState()
  const [fileName, setFileName] = React.useState()

  const SaveFile = (e) => {
    setFile(e.target.files[0])
    setFileName(e.target.files[0].name)
    setTimeout(() => {
      document.getElementById("btn-click").click()
    }, 500);
  };
  const UploadFile = async (e) => {
    const formData = new FormData()
    formData.append("file", file)
    formData.append("fileName", fileName)
    try {
      const res = await axios.post(
        "http://localhost:3200/upload", formData
      )
      console.log(res)
    } catch(error) {
      console.log(error)
    }
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
          <legend><span className="number">2</span>Дані користувача</legend>
            <input type="text" name="user_name" required placeholder="Прізвище та ініціали  *" />
            <input type="text" name="equipment_name" required placeholder="Назва контрольного столу *" />
          <legend><span className="number">3</span>Вибір обладнання</legend>
            <select id="equipment_select" name="equipment_select_list">
              <option value="equipment_name">V297_1</option>
              <option value="equipment_name">V297_2</option>
              <option value="equipment_name">V297_3</option>
              <option value="equipment_name">X254_1</option>
              <option value="equipment_name">VACUM_V297_1</option>
              <option value="equipment_name">VACUM_X254_1</option>
            </select>
          </fieldset>
          <fieldset>
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
          <input type="submit" value="Згенерувати" />
        </form>
      </div>
    </div>
  )
}

export default InitialFormComponent;