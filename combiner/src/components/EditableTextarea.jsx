import { memo, useCallback } from "react";

const EditableTextarea = memo(({ value, rowIndex, columnId, updateData }) => {
   

    const onBlur = useCallback(() => {
        updateData(rowIndex, columnId, value);
    }, [updateData, rowIndex, columnId, value]);

    

    return (
        <textarea
            className='textarea'
            defaultValue={value}
            onBlur={onBlur}
        />
    );
});
export default EditableTextarea