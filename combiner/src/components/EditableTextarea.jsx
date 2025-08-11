import { memo } from "react";
import { useDispatch, useSelector } from "react-redux";
import {  selectRegexByIdAndField, updateRegex } from "../features/regexesSlice";

const EditableTextarea = memo(({ regexId, field }) => {

  console.log(regexId,  "component rendering")
    
  const regexField = useSelector((state) => selectRegexByIdAndField(state, regexId, field))
  const dispatch = useDispatch()
    
    const onChange = (e) => {
      dispatch(updateRegex({id: regexId, field, data: e.target.value}))
    }
    
   
    
    return (
      <textarea
        className='textarea'
        value={regexField}
        onChange={onChange}
        
      />
    );
  });
export default EditableTextarea