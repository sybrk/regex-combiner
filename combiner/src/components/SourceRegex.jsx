import { memo, useEffect, useState } from "react";
import { useDispatch, useSelector } from "react-redux";
import { selectRegexByIdAndField, updateRegex } from "../features/regexesSlice";

const SourceRegex = memo(({ regexId, iframeRef }) => {

  console.log(regexId, "component rendering")

  const source = useSelector((state) => selectRegexByIdAndField(state, regexId, "source"))
  const condition = useSelector((state) => selectRegexByIdAndField(state, regexId, "condition"))
  const dispatch = useDispatch()
  const [hasError, setHaserror] = useState(false)

  
    
  const runRegex = (pattern) => {
    window.addEventListener("message", function handler(event) {
      if (event.data?.type === "regex-result") {
        console.log("regex result", event.data.result)
        setHaserror(!(event.data.result));
        window.removeEventListener("message", handler);
      }
    });

    iframeRef.current.contentWindow.postMessage(
      { type: "regex", pattern },
      "*"
    );
  };
   
  

    const onChange = (e) => {
      dispatch(updateRegex({ id: regexId, field: "source", data: e.target.value }))
      runRegex(e.target.value)
      /* const testString = "Hello World"

      try {
        const regexString = new RegExp(e.target.value, "g")
        regexString.test(testString)
        setHaserror(false)
      } catch (error) {
        console.error("regex err", error)
        setHaserror(true)
      } */
    }





    return (
      <textarea
        className={"textarea " + (hasError && "textarea-error text-error")}
        value={source}
        onChange={onChange}

        disabled={condition == "TargetOnly"}
      />
    );
  });
  export default SourceRegex