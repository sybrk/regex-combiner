import React from "react"

const EditingWithTextarea = React.memo(({ item, onChange }) => {
    
    return (
        <>
            <textarea
                className='textarea'
                onChange={(e) => onChange({ ...item, description: e.target.value })}
                value={item.description} />
        </>
    )
})

export default EditingWithTextarea