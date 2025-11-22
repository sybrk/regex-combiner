import { memo } from "react";
import type { RegexRecord } from "../utils/Util";


const EditableSelect = memo((props: { regexId: string, field: keyof RegexRecord, options }) => {

  const { regexId, field, options } = props
  const regexField = useSelector((state) => selectRegexByIdAndField(state, regexId, field))
  const dispatch = useDispatch()
    
    const onChange = (e) => {
      dispatch(updateRegex({id: regexId, field, data: e.target.value}))
    }
    
    return (
      <select
        value={regexField}
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