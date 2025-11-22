import { memo, type ChangeEvent } from "react";
import { useAppDispatch, useAppSelector } from "../hooks/reduxHooks";
import { selectRegexByIdAndField, updateRegex } from "../features/regexesSlice"
import { IGNORE_CASE_OPTIONS } from "../utils/Util";



const IgnoreCase = memo((props: { regexId: string }) => {

  const { regexId } = props
  const regexField = useAppSelector((state) => selectRegexByIdAndField(state.regexes, regexId, "ignoreCase")) as string
  const dispatch = useAppDispatch()



  const onChange = (e: ChangeEvent<HTMLSelectElement>) => {
    dispatch(updateRegex({ id: regexId, field: "ignoreCase", data: e.target.value }))
  }

  return (
    <select
      value={regexField}
      onChange={onChange}

      className="select select-accent"
    >
      {IGNORE_CASE_OPTIONS.map(option => (
        <option key={option} value={option}>{option}</option>
      ))}
    </select>
  );
});

export default IgnoreCase