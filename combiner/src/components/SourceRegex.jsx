import { memo, useEffect, useState } from "react";
import { useDispatch, useSelector } from "react-redux";
import { selectRegexByIdAndField, updateRegex } from "../features/regexesSlice";

const SourceRegex = memo(({ regexId, iframeRef }) => {

  //console.log(regexId, "component rendering")

  const source = useSelector((state) => selectRegexByIdAndField(state, regexId, "source"))
  const condition = useSelector((state) => selectRegexByIdAndField(state, regexId, "condition"))
  const dispatch = useDispatch()
  
  const sourceValid = useSelector((state) => selectRegexByIdAndField(state, regexId, "sourceValid"))
  

  /* useEffect(() => {
    runRegex(source)
  },[]) */
  useEffect(() => {
    function handler(event) {
      //console.log(event.data)
      if (event.data?.type === "regex-result" && event.data.regexId === regexId  && event.data.section === "source") {
        dispatch(updateRegex({ id: regexId, field: "sourceValid", data: event.data.result }))
        //console.log("validosource", sourceValid)
      }
    }
  
    window.addEventListener("message", handler);
    return () => window.removeEventListener("message", handler); // cleanup
  }, []);
    
  const runRegex = (pattern) => {
    

    iframeRef.current.contentWindow.postMessage(
      { type: "regex", pattern, regexId, section: "source" },
      "*"
    );
  };
   
  

    const onChange = (e) => {
      dispatch(updateRegex({ id: regexId, field: "source", data: e.target.value }))
      runRegex(e.target.value)
  
    }





    return (
      <>
      
      <textarea
        className={"textarea " + (sourceValid != null && !sourceValid && "textarea-error text-error")}
        value={source}
        onChange={onChange}
        disabled={condition == "TargetOnly"}
        spellCheck={false}
      />
      <p className={"text-error text-sm mt-1 " + ((sourceValid != null && !sourceValid) ? "" : "hidden")}>
        Source regex is not valid. Please fix it.
      </p>
      </>
      
    );
  });
  export default SourceRegex