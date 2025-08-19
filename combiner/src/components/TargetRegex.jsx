import { memo, useEffect, useState } from "react";
import { useDispatch, useSelector } from "react-redux";
import {  selectRegexByIdAndField, updateRegex } from "../features/regexesSlice";

const TargetRegex = memo(({ regexId, iframeRef }) => {

  console.log(regexId,  "component rendering")
    
  const target = useSelector((state) => selectRegexByIdAndField(state, regexId, "target"))
  const condition = useSelector((state) => selectRegexByIdAndField(state, regexId, "condition"))
  const dispatch = useDispatch()
  const targetValid = useSelector((state) => selectRegexByIdAndField(state, regexId, "targetValid"))
  
  /* useEffect(() => {
    runRegex(target)
  },[]) */
    
  useEffect(() => {
    function handler(event) {
      //console.log(event.data)
      if (event.data?.type === "regex-result" && event.data.regexId === regexId && event.data.section === "target") {
        dispatch(updateRegex({ id: regexId, field: "targetValid", data: event.data.result }))
      }
    }
  
    window.addEventListener("message", handler);
    return () => window.removeEventListener("message", handler); // cleanup
  }, []);
    
  const runRegex = (pattern) => {
    

    iframeRef.current.contentWindow.postMessage(
      { type: "regex", pattern, regexId, section: "target" },
      "*"
    );
  };
   
  

    const onChange = (e) => {
      dispatch(updateRegex({ id: regexId, field: "target", data: e.target.value }))
      runRegex(e.target.value)
  
    }

    
    
   
    
    return (
      <textarea
        className={"textarea " + (targetValid != null && !targetValid && "textarea-error text-error")}
        value={target}
        onChange={onChange}
        
        disabled={condition == "SourceOnly"}
      />
    );
  });
export default TargetRegex