import { memo, useEffect, type ChangeEvent } from "react";
import { useAppDispatch, useAppSelector } from "../hooks/reduxHooks";
import { selectRegexByIdAndField, updateRegex } from "../features/regexesSlice";


const TargetRegex = memo((props: { regexId: string, iframeRef: React.RefObject<HTMLIFrameElement | null> }) => {

  //console.log(regexId,  "component rendering")

  const { regexId, iframeRef } = props
    
  const target = useAppSelector((state) => selectRegexByIdAndField(state.regexes, regexId, "target")) as string
  const condition = useAppSelector((state) => selectRegexByIdAndField(state.regexes, regexId, "condition")) as string
  const dispatch = useAppDispatch()
  const targetValid = useAppSelector((state) => selectRegexByIdAndField(state.regexes, regexId, "targetValid")) as boolean | null
  
  /* useEffect(() => {
    runRegex(target)
  },[]) */
    
  useEffect(() => {
    function handler(event : MessageEvent) {
      //console.log(event.data)
      if (event.data?.type === "regex-result" && event.data.regexId === regexId && event.data.section === "target") {
        dispatch(updateRegex({ id: regexId, field: "targetValid", data: event.data.result }))
      }
    }
  
    window.addEventListener("message", handler);
    return () => window.removeEventListener("message", handler); // cleanup
  }, []);
    
  const runRegex = (pattern: string) => {
    

    iframeRef.current?.contentWindow?.postMessage(
      { type: "regex", pattern, regexId, section: "target" },
      "*"
    );
  };
   
  

    const onChange = (e: ChangeEvent<HTMLTextAreaElement>) => {
      dispatch(updateRegex({ id: regexId, field: "target", data: e.target.value }))
      runRegex(e.target.value)
  
    }

    
    
   
    
    return (
      <>
      <textarea
        className={"textarea " + (targetValid != null && !targetValid && "textarea-error text-error")}
        value={target}
        onChange={onChange}
        
        disabled={condition == "SourceOnly"}
        spellCheck={false}
      />
      <p className={"text-error text-sm mt-1 " + ((targetValid != null && !targetValid) ? "" : "hidden")}>
        Target regex is not valid. Please fix it.
      </p>
      </>
    );
  });
export default TargetRegex