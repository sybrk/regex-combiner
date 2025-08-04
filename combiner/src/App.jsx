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
    const tmpObj = {...regexList}
    tmpObj[data.fileId]["regexes"][data.regexId][data.field] = e.target.value;
    setRegexList(tmpObj)
  }
  return (
    <>
      <input type="file" name="regexfiles" id="regexfiles" multiple />
      <button onClick={importRegex}>Import</button>
      <table>
        <thead>
          <tr>
            <th>File</th>
            <th>Description</th>
            <th>IgnoreCase</th>
            <th>Source</th>
            <th>Target</th>
            <th>Condition</th>
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
                        <tr key={"regex" + j} data-file-id={file} data-regex-id={j}>
                          <td>{regexList[file]?.name}</td>
                          <td><textarea data-file-id={file} data-regex-id={j} data-field="description" onChange={handleChanges} type="text" name="" id="" value={regexList[file]["regexes"][regex]?.description} /></td>
                          <td>{regexList[file]["regexes"][regex]?.ignoreCase}</td>
                          <td>{regexList[file]["regexes"][regex]?.source}</td>
                          <td>{regexList[file]["regexes"][regex]?.target}</td>
                          <td>{regexList[file]["regexes"][regex]?.condition}</td>
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

    </>
  )
}

export default App
