import { memo, useCallback, useEffect, useState} from "react";

const EditableSelect = memo(({ value: initialValue, rowIndex, columnId, updateData, options }) => {
    const [value, setValue] = useState(initialValue);
    
    const onBlur = useCallback(() => {
      updateData(rowIndex, columnId, value);
    }, [updateData, rowIndex, columnId, value]);
    
    useEffect(() => {
      setValue(initialValue);
    }, [initialValue]);
    
    return (
      <select
        value={value}
        onChange={e => setValue(e.target.value)}
        onBlur={onBlur}
        className="select select-accent"
      >
        {options.map(option => (
          <option key={option} value={option}>{option}</option>
        ))}
      </select>
    );
  });

export default EditableSelect