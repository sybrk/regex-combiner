import { useRef, type MouseEvent} from 'react'


import { combineRegexes , importRegexes, newRegex, removeAllRegexes, removeFilters, removeRegex, searchByDescription, selectPagedRegexes, selectTotalPages, setPage} from '../../features/regexesSlice';
import useWaitingHandler from '../../hooks/useWaitingHandler';
import { useAppDispatch, useAppSelector } from '../../hooks/reduxHooks';
import { regexParserObj, type RegexRecordCollection } from '../../utils/Util';
import WaitingModal from '../modals/WaitingModal';
import FileInput from '../fileInput/FileInput';
import RegexFile from '../RegexFile';
import Description from '../Description';
import IgnoreCase from '../IgnoreCase';
import Condition from '../Condition';
import SourceRegex from '../SourceRegex';
import TargetRegex from '../TargetRegex';
import ScrollToBottom from '../ScrollToBottom';
import ScrollToTop from '../ScrollToTop';




function Combiner() {

  const [waitingMessage, setWaitingMessage, openModal, closeModal] = useWaitingHandler()

  const iframeRef = useRef<HTMLIFrameElement>(null);
  const dispatch = useAppDispatch()
  //const regexes = useSelector((state) => selectIds(state))
  const regexes = useAppSelector(selectPagedRegexes);
  const totalPages = useAppSelector(selectTotalPages);
  const page = useAppSelector(state => state.regexes.page);


  const regexCount = useAppSelector(state => Object.keys(state.regexes.value).length)
  const hasInvalidEntries = useAppSelector(state => state.regexes.invalidIds.length > 0)
  //console.log("invalidoe", hasInvalidEntries)
  const batchValidate = (regexesToValidate: RegexRecordCollection) => {
    return new Promise(resolve => {
      let expected = Object.keys(regexesToValidate).length * 2; // source + target
      let received = 0;
      function handler(event: MessageEvent) {

        if (event.data?.type === "regex-result" && event.data.section.includes("combine")) {
          received++;
          //console.log("about to set", regexesToValidate[event.data.regexId])
          regexesToValidate[event.data.regexId][event.data.section == "combineSource" ? "sourceValid" : "targetValid"] = event.data.result
          if (!(event.data.result)) {
            regexesToValidate[event.data.regexId]["hasIssues"] = regexesToValidate[event.data.regexId]["hasIssues"] != true && true
          }

          if (received === expected) {
            window.removeEventListener("message", handler);
            resolve(null);
          }
        }
        if (event.data?.type === "regex-result" && event.data.section.includes("nopattern")) {
          received++
        }
      }

      window.addEventListener("message", handler);

      // send all regexes
      Object.keys(regexesToValidate).forEach(x => {
        if (regexesToValidate[x]["source"]) {
          iframeRef.current?.contentWindow?.postMessage(
            { type: "regex", pattern: regexesToValidate[x]["source"], regexId: x, section: "combineSource" },
            "*"
          );
        } else {
          iframeRef.current?.contentWindow?.postMessage(
            { type: "regex", pattern: "hello", regexId: x, section: "nopattern" },
            "*"
          );

        }
        if (regexesToValidate[x]["target"]) {
          iframeRef.current?.contentWindow?.postMessage(
            { type: "regex", pattern: regexesToValidate[x]["target"], regexId: x, section: "combineTarget" },
            "*"
          );
        } else {
          iframeRef.current?.contentWindow?.postMessage(
            { type: "regex", pattern: "hello", regexId: x, section: "nopattern" },
            "*"
          );
        }

      });
    });
  };

  const importRegex = async (files: File[]) => {

    //const files = document.getElementById("regexfiles").files
    
    setWaitingMessage("Importing Regexes")
    openModal()
    const result = await regexParserObj(files);

    await batchValidate(result)
    //console.log("result import", result)
    dispatch(importRegexes(result))
    closeModal()

  }



  const createNew = () => {
    if (totalPages != 0) {
      dispatch(setPage(totalPages))
    }

    dispatch(newRegex())
    window.scrollTo(0, document.body.scrollHeight);
  }


  /*  useEffect(() => {
     if (!shouldCombine) return;
     setShouldCombine(false); // Reset flag
   }, [shouldCombine]); */

  const combine = (e : MouseEvent<HTMLButtonElement>) => {
    e.preventDefault()
    openModal()
    setWaitingMessage("Please wait while regexes are validated and output file is created.")
    dispatch(combineRegexes())
    closeModal()
    //setShouldCombine(true);
  }

  const searchDescriptions = (e: React.FormEvent) => {
    e.preventDefault()
    const searchTerm = (document.getElementById("description-search") as HTMLInputElement)?.value
    dispatch(searchByDescription(searchTerm));
    (document.getElementById("description-search") as HTMLInputElement).value = ""
  }

  return (
    <>
    
    
      <WaitingModal message={waitingMessage} />
      <div className='mt-20'>
        <iframe
          id='blazor_regex'
          ref={iframeRef}
          src="/regex-combiner/wwwroot/index.html" // your Blazor WASM build
          style={{ display: "none" }}
        />
        <title>Regex Combiner</title>
        <FileInput
          description={"Drag and drop .sdlqasettings files here or click to select."}
          fileHandler={importRegex}
          isMultiple={true}
          fileType={".sdlqasettings"}
        />
        <div className='mt-5 flex flex-row justify-center gap-4 items-center'>
          {/* <input className="file-input file-input-primary" type="file" name="regexfiles" id="regexfiles" multiple />
        <button className="btn btn-primary" onClick={importRegex}>Import</button> */}
          <p>
            Regex Count: {regexCount}
          </p>
          <button className={"btn btn-success rounded-2xl"} onClick={createNew}>+</button>
          <button className={"btn btn-success rounded-2xl " + (!regexes.length && "btn-disabled")} onClick={combine}>Combine</button>
          <button className={"btn btn-error rounded-2xl " + (!regexes.length && "btn-disabled")} onClick={() => dispatch(removeAllRegexes())}>Remove all</button>
        </div>
        <div className='mt-5 flex flex-row justify-center gap-4'>
          <label className="input">
            <svg className="h-[1em] opacity-50" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
              <g
                strokeLinejoin="round"
                strokeLinecap="round"
                strokeWidth="2.5"
                fill="none"
                stroke="currentColor"
              >
                <circle cx="11" cy="11" r="8"></circle>
                <path d="m21 21-4.3-4.3"></path>
              </g>
            </svg>
            <input id='description-search' type="search" className="grow" placeholder="Search in descriptions" />
            
          </label>
          <button className={"btn btn-info rounded-2xl"} onClick={searchDescriptions}>Search</button>
          <button className={"btn btn-error rounded-2xl"} onClick={() => dispatch(removeFilters())}>Clear</button>
        </div>
        <div className='flex flex-col items-center gap-2 my-3'>
          <div className="join">
            <button className={"join-item btn " + (page === 1 && "btn-disabled")}
              onClick={() => dispatch(setPage(1))}
            >««</button>
            <button className={"join-item btn " + (page === 1 && "btn-disabled")}
              onClick={() => dispatch(setPage(page - 1))}
            >«</button>
            <button className="join-item btn">{totalPages > 0 ? ( page.toString() + " / " + totalPages.toString()) : page}</button>
            <button className={"join-item btn " + ((page === totalPages || totalPages == 0) && "btn-disabled")}
              onClick={() => dispatch(setPage(page + 1))}
            >»</button>
            <button className={"join-item btn " + ((page === totalPages || totalPages == 0) && "btn-disabled")}
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
          <table className="table">
            <thead className='sticky top-16 z-50 bg-base-300 text-base-content'>
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
                ? regexes.map((regex: string, i) => {
                  //console.log("this is rendered again", regex)
                  return (
                    <tr className="" key={regex} data-regex-id={regex}>
                      <td><RegexFile regexId={regex} /></td>
                      <td>
                        <Description regexId={regex} />

                      </td>
                      <td>
                        <IgnoreCase regexId={regex} />
                      </td>
                      <td className='wrap-anywhere'>
                        <SourceRegex regexId={regex} iframeRef={iframeRef} />
                      </td>
                      <td className='wrap-anywhere'>
                        <TargetRegex regexId={regex} iframeRef={iframeRef} />
                      </td>
                      <td className=''>
                        <Condition regexId={regex} />


                      </td>
                      <td>
                        <button
                          data-regex-id={regex}
                          onClick={() => dispatch(removeRegex(regex))} className="btn btn-sm rounded-2xl btn-error"
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
