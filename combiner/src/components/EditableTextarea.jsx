import { memo } from "react";
import { useDispatch, useSelector } from "react-redux";
import { selectRegexById, updateRegex } from "../features/regexesSlice";

const EditableTextarea = memo(({ regexId, field }) => {

  console.log(regexId,  "component rendering")
    
  const regex = useSelector((state) => selectRegexById(state, regexId))
  const dispatch = useDispatch()
    
    const onChange = (e) => {
      dispatch(updateRegex({id: regexId, field, data: e.target.value}))
    }
    
   
    
    return (
      <textarea
        className='textarea'
        value={regex[field]}
        onChange={onChange}
        
      />
    );
  });
export default EditableTextarea