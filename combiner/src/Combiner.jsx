import { useRef, useState } from 'react'


import { regexParserObj, } from './utils/Util'

import { useDispatch, useSelector } from 'react-redux'
import { combineRegexes, importRegexes, newRegex, removeRegex, selectHasInvalidEntries, selectIds, selectPagedRegexes, selectTotalPages, setPage } from './features/regexesSlice'
import EditableSelect from './components/EditableSelect'
import RegexFile from './components/RegexFile'
import Description from './components/Description'
import SourceRegex from './components/SourceRegex'
import TargetRegex from './components/TargetRegex'
import ScrollToBottom from './components/ScrollToBottom'
import ScrollToTop from './components/ScrollToTop'




function Combiner() {



  const iframeRef = useRef(null);
  const dispatch = useDispatch()
  //const regexes = useSelector((state) => selectIds(state))
  const regexes = useSelector(selectPagedRegexes);
  const totalPages = useSelector(selectTotalPages);
  const page = useSelector(state => state.regexes.page);

  const [shouldCombine, setShouldCombine] = useState(false);
  const regexCount = useSelector(state => Object.keys(state.regexes.value).length)
  const hasInvalidEntries = useSelector(selectHasInvalidEntries)
  console.log("invalidoe", hasInvalidEntries)
  const batchValidate = (regexesToValidate) => {
    return new Promise(resolve => {
      let expected = Object.keys(regexesToValidate).length * 2; // source + target
      let received = 0;
      function handler(event) {
        if (event.data?.type === "regex-result" && event.data.section.includes("combine")) {
          received++;
          regexesToValidate[event.data.regexId][event.data.section == "combineSource" ? "sourceValid" : "targetValid"] = event.data.result
          if (!(event.data.result)) {
            regexesToValidate[event.data.regexId]["hasIssues"] = regexesToValidate[event.data.regexId]["hasIssues"] != true && true
          }

          if (received === expected) {
            window.removeEventListener("message", handler);
            resolve();
          }
        }
      }

      window.addEventListener("message", handler);

      // send all regexes
      Object.keys(regexesToValidate).forEach(x => {
        iframeRef.current.contentWindow.postMessage(
          { type: "regex", pattern: regexesToValidate[x]["source"], regexId: x, section: "combineSource" },
          "*"
        );
        iframeRef.current.contentWindow.postMessage(
          { type: "regex", pattern: regexesToValidate[x]["target"], regexId: x, section: "combineTarget" },
          "*"
        );
      });
    });
  };

  const importRegex = async () => {

    const files = document.getElementById("regexfiles").files
    const result = await regexParserObj(files);
    console.log("result import", result[1])
    await batchValidate(result)
    dispatch(importRegexes(result))
  }



  const createNew = () => {
    if(totalPages != 0) {
      dispatch(setPage(totalPages))
    }
    
    dispatch(newRegex())
    window.scrollTo(0, document.body.scrollHeight);
  }


 /*  useEffect(() => {
    if (!shouldCombine) return;
    setShouldCombine(false); // Reset flag
  }, [shouldCombine]); */

  const combine = () => {
    dispatch(combineRegexes())
    //setShouldCombine(true);
  }


  return (
    <>
    <div className='mt-20'>
    <iframe
        id='blazor_regex'
        ref={iframeRef}
        src="/regex-combiner/wwwroot/index.html" // your Blazor WASM build
        style={{ display: "none" }}
      />
      <title>Regex Combiner</title>

      <div className='mt-5 flex flex-row justify-center gap-4 items-center'>
        <input className="file-input file-input-primary" type="file" name="regexfiles" id="regexfiles" multiple />
        <button className="btn btn-primary" onClick={importRegex}>Import</button>
        <p>


          {regexCount}
        </p>
        <button className={"btn btn-success rounded-2xl"} onClick={createNew}>+</button>
        <button className={"btn btn-success rounded-2xl " + (!regexes.length && "btn-disabled")} onClick={combine}>Combine</button>
      </div>
      <div className='flex flex-col items-center gap-2 my-3'>
        <div className="join">
        <button className={"join-item btn " + (page === 1 && "btn-disabled")}
            onClick={() => dispatch(setPage(1))}
          >««</button>
          <button className={"join-item btn " + (page === 1 && "btn-disabled")}
            onClick={() => dispatch(setPage(page - 1))}
          >«</button>
          <button className="join-item btn">{page + (totalPages > 0 && (" / " + totalPages)) }</button>
          <button className={"join-item btn " + ((page === totalPages || totalPages == 0) && "btn-disabled")}
            onClick={() => dispatch(setPage(page + 1))}
          >»</button>
          <button className={"join-item btn " + ((page === totalPages || totalPages == 0)  && "btn-disabled")}
            onClick={() => dispatch(setPage(totalPages))}
          >»»</button>
        </div>
        <div role="alert" className={"alert alert-error " + (hasInvalidEntries ? "" : "hidden")}>
          <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6 shrink-0 stroke-current" fill="none" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M10 14l2-2m0 0l2-2m-2 2l-2-2m2 2l2 2m7-2a9 9 0 11-18 0 9 9 0 0118 0z" />
          </svg>
          <span>Error! Please fix the errors below before combining.</span>
        </div>
      </div>

      <div className="mx-2 rounded-box border border-base-content/5 bg-base-100">
        <table className="table table-zebra">
          <thead className='sticky top-0 z-50 bg-base-300 text-base-content'>
            <tr className='text-center'>
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
                  <tr className="" key={regex} data-regex-id={regex}>
                    <td><RegexFile regexId={regex} /></td>
                    <td>
                      <Description regexId={regex} />

                    </td>
                    <td>
                      <EditableSelect regexId={regex} field={"ignoreCase"} options={["true", "false"]} />
                    </td>
                    <td className='wrap-anywhere'>
                      <SourceRegex regexId={regex} iframeRef={iframeRef} />
                    </td>
                    <td className='wrap-anywhere'>
                      <TargetRegex regexId={regex} iframeRef={iframeRef} />
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
        <ScrollToBottom />
        <ScrollToTop />
      </div >

    </div>
      

    </>
  )
}

export default Combiner
