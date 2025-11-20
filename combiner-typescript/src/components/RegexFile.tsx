
import {  useSelector } from "react-redux";
import {  selectRegexByIdAndField } from "../features/regexesSlice";

const RegexFile = ({ regexId }) => {

  
  const regexField = useSelector((state) => selectRegexByIdAndField(state, regexId, "file"))
    
   
    
    return (
      <p>{regexField}</p>
    );
  };
export default RegexFile