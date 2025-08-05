import { useState } from 'react'

import './App.css'
import { idGenerator, readFile, xmlParser } from './utils/Util'

function App() {

  const [regexList, setRegexList] = useState({})

  const importRegex = async () => {
    const tmpObj = { ...regexList }
    const files = document.getElementById("regexfiles").files
    for (let index = 0; index < files.length; index++) {
      const file = files[index];
      const fileRead = await readFile(file);
      const parseFile = xmlParser(fileRead.fileContent)
      const fileId = idGenerator();
      tmpObj[fileId] = {}
      tmpObj[fileId]["name"] = fileRead.fileName;
      tmpObj[fileId]["original"] = fileRead.fileContent;
      tmpObj[fileId]["regexes"] = {}
      let regexRules = Array.from(parseFile.querySelectorAll("RegExRule"));
      console.log("regexrules", regexRules)
      regexRules = regexRules.filter(x => /RegExRules\d+$/.test(x.parentElement.getAttribute('Id')))

      regexRules.map((x, i) => {
        tmpObj[fileId]["regexes"][i] = {}
        tmpObj[fileId]["regexes"][i]["description"] = x.querySelector("Description").textContent;
        tmpObj[fileId]["regexes"][i]["ignoreCase"] = x.querySelector("IgnoreCase").textContent;
        tmpObj[fileId]["regexes"][i]["source"] = x.querySelector("RegExSource").textContent;
        tmpObj[fileId]["regexes"][i]["target"] = x.querySelector("RegExTarget").textContent;
        tmpObj[fileId]["regexes"][i]["condition"] = x.querySelector("RuleCondition").textContent;
      });

    }
    setRegexList(tmpObj)
  }
  const handleChanges = (e) => {
    const data = e.target.dataset;
    const tmpObj = { ...regexList }
    tmpObj[data.fileId]["regexes"][data.regexId][data.field] = e.target.value;
    setRegexList(tmpObj)
  }

  const removeRegex = (e) => {
    const data = e.target.dataset;
    const tmpObj = { ...regexList }
    delete tmpObj[data.fileId]["regexes"][data.regexId]
    setRegexList(tmpObj)
  }

  const createNew = () => {
    const tmpObj = { ...regexList }
    if(tmpObj["newRegexes"] == undefined) {
      tmpObj["newRegexes"] = {
        "name": "newRegexes",
        regexes: {}
      }
    }
    
    const newRegexId = Object.keys(tmpObj["newRegexes"]["regexes"]).length
    tmpObj["newRegexes"]["regexes"][newRegexId] = {
      "description" : "",
      ignoreCase: "true",
      source: "",
      target: "",
      condition: "TargetAndSource"
    }
    setRegexList(tmpObj)
  }
  return (
    <>
      <div className='mt-5 flex flex-row justify-center gap-4 items-center'>
        <input className="file-input file-input-primary" type="file" name="regexfiles" id="regexfiles" multiple />
        <button className="btn btn-primary" onClick={importRegex}>Import</button>
        <p>
          {Object.keys(regexList)
              ?
              Object.keys(regexList).map(x => Object.keys(regexList[x]["regexes"]).length).reduce(((a,b)=> a + b),0)
              :
              null
            }
        </p>
        <button className="btn btn-success rounded-2xl" onClick={createNew}>+</button>
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
            {Object.keys(regexList)
              ? Object.keys(regexList).map((file, i) => {
                return (
                  <>
                    {
                      Object.keys(regexList[file]["regexes"]).map((regex, j) => {
                        return (
                          <tr className="hover:bg-base-300" key={"regex" + j} data-file-id={file} data-regex-id={regex}>
                            <td>{regexList[file]?.name}</td>
                            <td>
                              <textarea
                                className='textarea'
                                data-file-id={file} data-regex-id={regex} data-field="description"
                                onChange={handleChanges}
                                value={regexList[file]["regexes"][regex]?.description} />
                            </td>
                            <td>
                              <select data-file-id={file} data-regex-id={regex} data-field="ignoreCase"
                                onChange={handleChanges}
                                value={regexList[file]["regexes"][regex]?.ignoreCase}
                                className="select select-accent">

                                <option>true</option>
                                <option>false</option>

                              </select>
                            </td>
                            <td className='wrap-anywhere'>
                              <textarea
                                className='textarea'
                                data-file-id={file} data-regex-id={regex} data-field="source"
                                onChange={handleChanges}
                                value={regexList[file]["regexes"][regex]?.source}
                                disabled={regexList[file]["regexes"][regex]?.condition == "TargetOnly"}
                              />
                            </td>
                            <td className='wrap-anywhere'>
                              <textarea
                                className='textarea'
                                data-file-id={file} data-regex-id={regex} data-field="target"
                                onChange={handleChanges}
                                value={regexList[file]["regexes"][regex]?.target}
                                disabled={regexList[file]["regexes"][regex]?.condition == "SourceOnly"}
                                 />
                            </td>
                            <td className=''>
                              <select data-file-id={file} data-regex-id={regex} data-field="condition"
                                onChange={handleChanges}
                                value={regexList[file]["regexes"][regex]?.condition}
                                className="select select-accent">

                                <option>TargetAndSource</option>
                                <option>TargetNotSource</option>
                                <option>SourceNotTarget</option>
                                <option>SourceOnly</option>
                                <option>TargetOnly</option>
                                <option>DifferentCount</option>
                                <option>GroupedSourceNotTarget</option>
                                <option>GroupedTargetAndSource</option>

                              </select>

                            </td>
                            <td>
                            <button 
                              data-file-id={file} data-regex-id={regex}
                              onClick={removeRegex} className="btn btn-sm rounded-2xl btn-error"
                            >
                              x
                            </button>
                            </td>
                          </tr>
                        );
                      })
                    }
                  </>
                )


              })
              : null}
          </tbody>
        </table>
      </div>


    </>
  )
}

export default App
