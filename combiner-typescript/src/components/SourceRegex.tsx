import { memo, useEffect, type ChangeEvent } from "react";
import { useAppDispatch, useAppSelector } from "../hooks/reduxHooks";
import { selectRegexByIdAndField, updateRegex } from "../features/regexesSlice";



const SourceRegex = memo((props: { regexId: string, iframeRef: React.RefObject<HTMLIFrameElement | null> }) => {

  const { regexId, iframeRef } = props
  //console.log(regexId, "component rendering")

  const source = useAppSelector((state) => selectRegexByIdAndField(state.regexes, regexId, "source")) as string
  const condition = useAppSelector((state) => selectRegexByIdAndField(state.regexes, regexId, "condition")) as string
  const dispatch = useAppDispatch()
  
  const sourceValid = useAppSelector((state) => selectRegexByIdAndField(state.regexes, regexId, "sourceValid")) as boolean | null
  

  /* useEffect(() => {
    runRegex(source)
  },[]) */
  useEffect(() => {
    function handler(event: MessageEvent) {
      //console.log(event.data)
      if (event.data?.type === "regex-result" && event.data.regexId === regexId  && event.data.section === "source") {
        dispatch(updateRegex({ id: regexId, field: "sourceValid", data: event.data.result }))
        //console.log("validosource", sourceValid)
      }
    }
  
    window.addEventListener("message", handler);
    return () => window.removeEventListener("message", handler); // cleanup
  }, []);
    
  const runRegex = (pattern: string) => {
    

    iframeRef.current?.contentWindow?.postMessage(
      { type: "regex", pattern, regexId, section: "source" },
      "*"
    );
  };
   
  

    const onChange = (e: ChangeEvent<HTMLTextAreaElement>) => {
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