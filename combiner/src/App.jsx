import { useState } from 'react'

import './App.css'
import { readFile, xmlParser } from './utils/Util'

function App() {

  const [regexList, setRegexList] = useState([])

  const importRegex = async() => {
    const files = document.getElementById("regexfiles").files
    for (let index = 0; index < files.length; index++) {
      const file = files[index];
      const fileRead = await readFile(file);
      const parseFile = xmlParser(fileRead.fileContent)
      let regexRules = Array.from(parseFile.querySelectorAll("RegExRule"));
      console.log("regexrules", regexRules)
      regexRules = regexRules.filter(x => /RegExRules\d+$/.test(x.parentElement.getAttribute('Id')))
      const regexes = regexRules.map(x => {
        return {
          description: x.querySelector("Description").textContent,
          ignoreCase: x.querySelector("IgnoreCase").textContent,
          source: x.querySelector("RegExSource").textContent,
          target: x.querySelector("RegExTarget").textContent,
          condition: x.querySelector("RuleCondition").textContent,
        }
      });
      setRegexList(prev => [...prev, ...regexes])
    }
  }
  return (
    <>
      <input type="file" name="regexfiles" id="regexfiles" multiple />
      <button onClick={importRegex}>Import</button>
      <table>
        <thead>
          <tr>
            <th>Description</th>
            <th>IgnoreCase</th>
            <th>Source</th>
            <th>Target</th>
            <th>Condition</th>
          </tr>
        </thead>
        <tbody>
          {regexList
            ? regexList.map((regex, i) => {
              return (
                <tr key={"regex" + i}>
                  <td>{regex?.description}</td>
                  <td>{regex?.ignoreCase}</td>
                  <td>{regex?.source}</td>
                  <td>{regex?.target}</td>
                  <td>{regex?.condition}</td>
                </tr>
              );
            })
            : null}
        </tbody>
      </table>

    </>
  )
}

export default App
