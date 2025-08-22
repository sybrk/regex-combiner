import { memo } from "react";
import { useDispatch, useSelector } from "react-redux";
import {  selectRegexByIdAndField, updateRegex } from "../features/regexesSlice";

const Description = memo(({ regexId }) => {

  //console.log(regexId,  "component rendering")
    
  const descriptionField = useSelector((state) => selectRegexByIdAndField(state, regexId, "description"))
  const duplicates = useSelector((state) => state.regexes.duplicates)
  const dispatch = useDispatch()
  
    
    const onChange = (e) => {
      dispatch(updateRegex({id: regexId, field: "description", data: e.target.value}))
      if((duplicates && duplicates.includes(descriptionField)) || !descriptionField.length) {
        dispatch(updateRegex({id: regexId, field: "hasIssues", data: true}))
      }
    }
    
   
    
    return (
      <textarea
        className={"textarea " + (((duplicates && duplicates.includes(descriptionField)) ||  !descriptionField.length) && "textarea-error text-error")}
        value={descriptionField}
        onChange={onChange}
        
      />
    );
  });
export default Description