import { memo, type ChangeEvent } from "react";
import { useAppDispatch, useAppSelector } from "../hooks/reduxHooks";
import { selectRegexByIdAndField, updateRegex } from "../features/regexesSlice";


const Description = memo((props: {regexId: string}) => {

  const {regexId} = props
  //console.log(regexId,  "component rendering")
    
  const descriptionField = useAppSelector((state) => selectRegexByIdAndField(state.regexes, regexId, "description")) as string
  const duplicates = useAppSelector((state) => state.regexes.duplicates)
  const dispatch = useAppDispatch()
  
    
    const onChange = (e: ChangeEvent<HTMLTextAreaElement>) => {
      dispatch(updateRegex({id: regexId, field: "description", data: e.target.value}))
      if((duplicates && duplicates.includes(descriptionField)) || !descriptionField.length) {
        dispatch(updateRegex({id: regexId, field: "hasIssues", data: true}))
      }
    }
    
   
    
    return (
      <>
      
      <textarea
        className={"textarea " + (((duplicates && duplicates.includes(descriptionField)) ||  !descriptionField.length) && "textarea-error text-error")}
        value={descriptionField}
        placeholder="Please enter a description."
        onChange={onChange}
        
      />
      <p className={"text-error text-sm mt-1 " + ((duplicates && duplicates.includes(descriptionField)) ? "" : "hidden")}>
        Duplicate description, please fix it and try again.
      </p>
      </>
    );
  });
export default Description