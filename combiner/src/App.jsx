import { useCallback, useEffect, useMemo, useRef, useState } from 'react'

import './App.css'
import {  regexParserObj, } from './utils/Util'

import { useDispatch, useSelector } from 'react-redux'
import { combineRegexes, importRegexes, newRegex, removeRegex, selectIds } from './features/regexesSlice'
import EditableSelect from './components/EditableSelect'
import RegexFile from './components/RegexFile'
import Description from './components/Description'
import SourceRegex from './components/SourceRegex'
import TargetRegex from './components/TargetRegex'




function App() {


  
  const iframeRef = useRef(null);
  const dispatch = useDispatch()
  const regexes = useSelector((state) => selectIds(state))
 
  const [shouldCombine, setShouldCombine] = useState(false);
  
  const importRegex = async () => {

    const files = document.getElementById("regexfiles").files
    const result = await regexParserObj(files);
    dispatch(importRegexes(result))
  }
 


  const createNew = () => {
    dispatch(newRegex())
  }

  
  useEffect(() => {
    if (!shouldCombine) return;
    setShouldCombine(false); // Reset flag
  }, [shouldCombine]);

  const combine = () => {
    dispatch(combineRegexes())
    setShouldCombine(true);
  }


  return (
    <>
    <iframe
        id='blazor_regex'
        ref={iframeRef}
        src="/wwwroot/index.html" // your Blazor WASM build
        style={{ display: "none" }}
      />
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
                      <Description regexId={regex} />

                    </td>
                    <td>
                      <EditableSelect regexId={regex} field={"ignoreCase"} options={["true", "false"]} />
                    </td>
                    <td className='wrap-anywhere'>
                      <SourceRegex regexId={regex} iframeRef = {iframeRef} />
                    </td>
                    <td className='wrap-anywhere'>
                      <TargetRegex regexId={regex} iframeRef = {iframeRef}/>
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
