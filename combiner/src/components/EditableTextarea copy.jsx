import { memo, useCallback, useEffect, useState } from "react";

const EditableTextarea = memo(({ value: initialValue, rowIndex, columnId, updateData }) => {
    const [value, setValue] = useState(initialValue);
    
    const onBlur = useCallback(() => {
      updateData(rowIndex, columnId, value);
    }, [updateData, rowIndex, columnId, value]);
    
    useEffect(() => {
      setValue(initialValue);
    }, [initialValue]);
    
    return (
      <textarea
        className='textarea'
        value={value}
        onChange={e => setValue(e.target.value)}
        onBlur={onBlur}
      />
    );
  });
export default EditableTextarea