import { memo, useEffect, useState } from "react";
import { useDispatch, useSelector } from "react-redux";
import { selectRegexByIdAndField, updateRegex } from "../features/regexesSlice";

const SourceRegex = memo(({ regexId }) => {

  console.log(regexId, "component rendering")

  const source = useSelector((state) => selectRegexByIdAndField(state, regexId, "source"))
  const condition = useSelector((state) => selectRegexByIdAndField(state, regexId, "condition"))
  const dispatch = useDispatch()
  const [hasError, setHaserror] = useState(false)

  useEffect(() => {

    const checkRegex = async () => {
      const result = await window.DotNet.invokeMethodAsync(
        "RegexValidator", // Assembly name from Blazor .csproj
        "ValidateRegex",
        source
      );
      setHaserror(result);
    }
    
    if (!window.blazorStarted) {
      // Start Blazor and then call checkRegex
      window.blazorStarted = window.Blazor.start().then(() => {
        checkRegex();
      });
    } else {
      checkRegex();
    }
    
    
   
  },[source])

    const onChange = (e) => {
      dispatch(updateRegex({ id: regexId, field: "source", data: e.target.value }))
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