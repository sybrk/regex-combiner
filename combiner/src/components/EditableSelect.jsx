import { memo, useCallback} from "react";

const EditableSelect = memo(({ value, rowIndex, columnId, updateData, options }) => {
    

    const onBlur = useCallback(() => {
        updateData(rowIndex, columnId, value);
    }, [updateData, rowIndex, columnId, value]);

   

    return (
        <select
            defaultValue={value}
            
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