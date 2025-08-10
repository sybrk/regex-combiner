import { useCallback, useMemo, useState } from 'react'

import './App.css'
import { idGenerator, newRegexFile, readFile, regexNodeBuilder, regexParserObj, xmlParser } from './utils/Util'
import { saveAs } from 'file-saver'
import { useDispatch, useSelector } from 'react-redux'
import { getRegexIds, importRegexes, updateRegex } from './features/regexesSlice'
import EditableTextarea from './components/EditableTextarea'
import EditableSelect from './components/EditableSelect'




function App() {


  //const [regexList, setRegexList] = useState({})
  const dispatch = useDispatch()
  const regexes = useSelector(getRegexIds)
  //console.log("coming", regexes)
  const importRegex = async () => {
   
    const files = document.getElementById("regexfiles").files
    const result = await regexParserObj(files);
    dispatch(importRegexes(result))
  }
  const handleChanges = (e) => {
    const data = e.target.dataset;
    
    dispatch(updateRegex({id: data.regexId, field: data.field, data: e.target.value}))
  }

  const removeRegex = (e) => {
    console.log("neredeyim", e.target)
    const data = e.target.dataset;
    const tmpObj = { ...regexList }
    delete tmpObj[data.fileId]["regexes"][data.regexId]
    setRegexList(tmpObj)
  }

  const createNew = () => {
    const tmpObj = { ...regexList }
    if (tmpObj["newRegexes"] == undefined) {
      tmpObj["newRegexes"] = {
        "name": "newRegexes",
        regexes: {}
      }
    }

    const newRegexId = Object.keys(tmpObj["newRegexes"]["regexes"]).length
    tmpObj["newRegexes"]["regexes"][newRegexId] = {
      "description": "",
      ignoreCase: "true",
      source: "",
      target: "",
      condition: "TargetAndSource"
    }
    setRegexList(tmpObj)
  }
  const combine = () => {
    const newFile = newRegexFile();
    const parent = newFile.querySelector("SettingsGroup");
    const regexCount = Object.keys(regexList).map(x => Object.keys(regexList[x]["regexes"]).length).reduce(((a, b) => a + b), 0);
    const regexCountNode = document.createElementNS("", "Setting");
    regexCountNode.setAttribute("Id", "RegExRulesCount");
    regexCountNode.textContent = regexCount;
    parent.appendChild(regexCountNode)
    let regexId = 0;
    Object.keys(regexList).map((file, i) => {
      Object.keys(regexList[file]["regexes"]).map((regex) => {
        const settingNode = regexNodeBuilder(regexList[file]["regexes"][regex], regexId);
        parent.appendChild(settingNode)
        regexId++
      })
    })
    const serialer = new XMLSerializer();
    const serializedFile = serialer.serializeToString(newFile);
    const fileToDownload = new File([serializedFile], "combined.sdlqasettings", {
      type: "text/xml",
    });
    saveAs(fileToDownload)
  }


  return (
    <>
      <title>Regex Combiner</title>
      <div className='mt-5 flex flex-row justify-center gap-4 items-center'>
        <input className="file-input file-input-primary" type="file" name="regexfiles" id="regexfiles" multiple />
        <button className="btn btn-primary" onClick={importRegex}>Import</button>
        <p>


          {regexes.length}
        </p>
        <button className="btn btn-success rounded-2xl" onClick={createNew}>+</button>
        <button className="btn btn-success rounded-2xl" onClick={combine}>Combine</button>
      </div>

      <div className="">
        <table className="table">
          <thead>
            <tr>
              <th>File</th>
              <th>Description</th>
              <th>IgnoreCase</th>
              <th>Source</th>
              <th>Target</th>
              <th>Condition</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {regexes
              ? regexes.map((regex, i) => {
                //console.log("this is rendered again", regex)
                return (
                  <tr className="hover:bg-base-300" key={regex} data-regex-id={regex}>
                    <td>{/* {regexes[regex].file} */}</td>
                    <td>
                      <EditableTextarea regexId={regex} field={"description"} />

                    </td>
                    <td>
                      <EditableSelect regexId= {regex} field = {"ignoreCase"} options = {["true", "false"]} />
                    </td>
                    <td className='wrap-anywhere'>
                    <EditableTextarea regexId={regex} field={"source"} />
                    </td>
                    <td className='wrap-anywhere'>
                    <EditableTextarea regexId={regex} field={"target"} />
                    </td>
                    <td className=''>
                    <EditableSelect regexId= {regex} field = {"condition"} options = {["TargetAndSource", "TargetNotSource", "SourceNotTarget", "SourceOnly", "TargetOnly", "DifferentCount", "GroupedSourceNotTarget", "GroupedTargetAndSource"]} />
                      

                    </td>
                    <td>
                      <button
                        data-regex-id={regex}
                        onClick={removeRegex} className="btn btn-sm rounded-2xl btn-error"
                      >
                        x
                      </button>
                    </td>
                  </tr>
                );
              })


              : null}
          </tbody>
        </table>
      </div >


    </>
  )
}

export default App
