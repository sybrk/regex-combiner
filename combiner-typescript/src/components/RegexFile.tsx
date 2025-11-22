import { selectRegexByIdAndField } from "../features/regexesSlice";
import { useAppSelector } from "../hooks/reduxHooks";


const RegexFile = (props: {regexId: string}) => {

  const {regexId} = props
  const regexField = useAppSelector(state => selectRegexByIdAndField(state.regexes,regexId, "file"))
    
   
    
    return (
      <p>{regexField}</p>
    );
  };
export default RegexFile