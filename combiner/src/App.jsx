import { useCallback, useMemo, useState } from 'react'

import './App.css'
import { idGenerator, newRegexFile, readFile, regexNodeBuilder, regexParserObj, xmlParser } from './utils/Util'
import { saveAs } from 'file-saver'
import { useDispatch, useSelector } from 'react-redux'
import { importRegexes, newRegex, removeRegex, selectIds, updateRegex } from './features/regexesSlice'
import EditableTextarea from './components/EditableTextarea'
import EditableSelect from './components/EditableSelect'
import RegexFile from './components/RegexFile'




function App() {


  //const [regexList, setRegexList] = useState({})
  const dispatch = useDispatch()
  const regexes = useSelector((state) => selectIds(state))
  const regexObjects = useSelector(state => state.regexes.value)
  console.log("coming", regexes)
  const importRegex = async () => {

    const files = document.getElementById("regexfiles").files
    const result = await regexParserObj(files);
    dispatch(importRegexes(result))
  }
  const handleChanges = (e) => {
    const data = e.target.dataset;

    dispatch(updateRegex({ id: data.regexId, field: data.field, data: e.target.value }))
  }


  const createNew = () => {
    dispatch(newRegex())
  }
  const combine = () => {
    const newFile = newRegexFile();
    const parent = newFile.querySelector("SettingsGroup");
    const regexCount = regexes.length
    const regexCountNode = document.createElementNS("", "Setting");
    regexCountNode.setAttribute("Id", "RegExRulesCount");
    regexCountNode.textContent = regexCount;
    parent.appendChild(regexCountNode)
    let regexId = 0;

    Object.keys(regexObjects).map((regex) => {
      const settingNode = regexNodeBuilder(regexObjects[regex], regexId);
      parent.appendChild(settingNode)
      regexId++
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


          {regexes?.length}
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
                    <td><RegexFile regexId={regex} /></td>
                    <td>
                      <EditableTextarea regexId={regex} field={"description"} />

                    </td>
                    <td>
                      <EditableSelect regexId={regex} field={"ignoreCase"} options={["true", "false"]} />
                    </td>
                    <td className='wrap-anywhere'>
                      <EditableTextarea regexId={regex} field={"source"} />
                    </td>
                    <td className='wrap-anywhere'>
                      <EditableTextarea regexId={regex} field={"target"} />
                    </td>
                    <td className=''>
                      <EditableSelect regexId={regex} field={"condition"} options={["TargetAndSource", "TargetNotSource", "SourceNotTarget", "SourceOnly", "TargetOnly", "DifferentCount", "GroupedSourceNotTarget", "GroupedTargetAndSource"]} />


                    </td>
                    <td>
                      <button
                        data-regex-id={regex}
                        onClick={() => dispatch(removeRegex({ regexId: regex }))} className="btn btn-sm rounded-2xl btn-error"
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
