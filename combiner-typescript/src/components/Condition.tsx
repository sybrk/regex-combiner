import { memo, type ChangeEvent } from "react";
import { useAppDispatch, useAppSelector } from "../hooks/reduxHooks";
import { selectRegexByIdAndField, updateRegex } from "../features/regexesSlice"
import { CONDITIONS } from "../utils/Util";



const Condition = memo((props: { regexId: string }) => {

  const { regexId } = props
  const regexField = useAppSelector((state) => selectRegexByIdAndField(state.regexes, regexId, "condition")) as string
  const dispatch = useAppDispatch()



  const onChange = (e: ChangeEvent<HTMLSelectElement>) => {
    dispatch(updateRegex({ id: regexId, field: "condition", data: e.target.value }))
  }

  return (
    <select
      value={regexField}
      onChange={onChange}

      className="select select-accent"
    >
      {CONDITIONS.map(option => (
        <option key={option} value={option}>{option}</option>
      ))}
    </select>
  );
});

export default Condition