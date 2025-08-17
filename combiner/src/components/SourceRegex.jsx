import { memo, useEffect, useState } from "react";
import { useDispatch, useSelector } from "react-redux";
import { selectRegexByIdAndField, updateRegex } from "../features/regexesSlice";

const SourceRegex = memo(({ regexId, iframeRef }) => {

  console.log(regexId, "component rendering")

  const source = useSelector((state) => selectRegexByIdAndField(state, regexId, "source"))
  const condition = useSelector((state) => selectRegexByIdAndField(state, regexId, "condition"))
  const dispatch = useDispatch()
  
  const sourceValid = useSelector((state) => selectRegexByIdAndField(state, regexId, "sourceValid"))
  

  /* useEffect(() => {
    runRegex(source)
  },[]) */
  useEffect(() => {
    function handler(event) {
      console.log(event.data)
      if (event.data?.type === "regex-result" && event.data.regexId === regexId  && event.data.section === "source") {
        dispatch(updateRegex({ id: regexId, field: "sourceValid", data: event.data.result }))
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
      <textarea
        className={"textarea " + (sourceValid != null && !sourceValid && "textarea-error text-error")}
        value={source}
        onChange={onChange}

        disabled={condition == "TargetOnly"}
      />
    );
  });
  export default SourceRegex