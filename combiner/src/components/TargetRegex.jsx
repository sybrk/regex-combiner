import { memo, useState } from "react";
import { useDispatch, useSelector } from "react-redux";
import {  selectRegexByIdAndField, updateRegex } from "../features/regexesSlice";

const TargetRegex = memo(({ regexId }) => {

  console.log(regexId,  "component rendering")
    
  const target = useSelector((state) => selectRegexByIdAndField(state, regexId, "target"))
  const condition = useSelector((state) => selectRegexByIdAndField(state, regexId, "condition"))
  const dispatch = useDispatch()
  const [hasError, setHaserror] = useState(false)
  
    
    const onChange = (e) => {
      dispatch(updateRegex({id: regexId, field:"target", data: e.target.value}))
      const testString = "Hello World"
      
      try {
        const regexString = new RegExp(e.target.value, "g")
        regexString.test(testString)
        setHaserror(false)
      } catch (error) {
        console.error("regex err", error)
        setHaserror(true)
      }
    }

    
    
   
    
    return (
      <textarea
        className={"textarea " + (hasError && "textarea-error text-error")}
        value={target}
        onChange={onChange}
        
        disabled={condition == "SourceOnly"}
      />
    );
  });
export default TargetRegex