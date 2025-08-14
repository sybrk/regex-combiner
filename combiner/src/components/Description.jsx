import { memo } from "react";
import { useDispatch, useSelector } from "react-redux";
import {  selectRegexByIdAndField, updateRegex } from "../features/regexesSlice";

const Description = memo(({ regexId }) => {

  console.log(regexId,  "component rendering")
    
  const regexField = useSelector((state) => selectRegexByIdAndField(state, regexId, "description"))
  const duplicates = useSelector((state) => state.regexes.duplicates)
  const dispatch = useDispatch()
  
    
    const onChange = (e) => {
      dispatch(updateRegex({id: regexId, field: "description", data: e.target.value}))
    }
    
   
    
    return (
      <textarea
        className={"textarea " + (((duplicates && duplicates.includes(regexField)) ||  !regexField.length) && "textarea-error text-error")}
        value={regexField}
        onChange={onChange}
        
      />
    );
  });
export default Description