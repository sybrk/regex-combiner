import { memo } from "react";
import { useDispatch, useSelector } from "react-redux";
import { updateRegex } from "../features/regexesSlice";

const EditableSelect = memo(({ regexId, field, options }) => {
  const regex = useSelector((state) => state.regexes.value[regexId])
  const dispatch = useDispatch()
    
    const onChange = (e) => {
      dispatch(updateRegex({id: regexId, field, data: e.target.value}))
    }
    
    return (
      <select
        value={regex[field]}
        onChange={onChange}
        
        className="select select-accent"
      >
        {options.map(option => (
          <option key={option} value={option}>{option}</option>
        ))}
      </select>
    );
  });

export default EditableSelect