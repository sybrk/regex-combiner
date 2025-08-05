export const xmlParser = (xmlContent) => {
  let parser = new DOMParser();
  let xmldoc = parser.parseFromString(xmlContent, "text/xml");
  return xmldoc;
}

export const readFile = async (file) => {
  let reader = new FileReader();
  // read file as text
  reader.readAsText(file);
  // await the file to be read.
  await new Promise((resolve) => (reader.onload = () => resolve()));

  // return filename and filecontent with keys.
  return {
    fileName: file.name,
    fileContent: reader.result,
  };
};
export const idGenerator = () => {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'
    .replace(/[xy]/g, function (c) {
      const r = Math.random() * 16 | 0,
        v = c == 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
}

export const newRegexFile = () => {
  const xmlString = `<?xml version="1.0" encoding="utf-8"?><SettingsBundle><SettingsGroup Id="QAVerificationSettings"><Setting Id="RegExRules">True</Setting></SettingsGroup></SettingsBundle>`;
  const newFile = xmlParser(xmlString);
  return newFile;
}

export const regexNodeBuilder = (regexObj, id) => {
  const settingNode = document.createElementNS("","Setting");
  settingNode.setAttribute("Id", id)
  const regexNode = document.createElementNS("http://schemas.datacontract.org/2004/07/Sdl.Verification.QAChecker.RegEx","RegExRule");
  regexNode.setAttribute("xmlns:i", "http://www.w3.org/2001/XMLSchema-instance")
  const descriptionNode = document.createElementNS("http://schemas.datacontract.org/2004/07/Sdl.Verification.QAChecker.RegEx", "Description");
  const caseNode = document.createElementNS("http://schemas.datacontract.org/2004/07/Sdl.Verification.QAChecker.RegEx","IgnoreCase");
  const sourceNode = document.createElementNS("http://schemas.datacontract.org/2004/07/Sdl.Verification.QAChecker.RegEx","RegExSource");
  const targetNode = document.createElementNS("http://schemas.datacontract.org/2004/07/Sdl.Verification.QAChecker.RegEx","RegExTarget");
  const conditionNode = document.createElementNS("http://schemas.datacontract.org/2004/07/Sdl.Verification.QAChecker.RegEx","RuleCondition");
  descriptionNode.textContent = regexObj.description;
  caseNode.textContent = regexObj.ignoreCase;
  sourceNode.textContent = regexObj.source;
  targetNode.textContent = regexObj.target;
  conditionNode.textContent = regexObj.condition;
  regexNode.appendChild(descriptionNode)
  regexNode.appendChild(caseNode)
  regexNode.appendChild(sourceNode)
  regexNode.appendChild(targetNode)
  regexNode.appendChild(conditionNode)
  settingNode.appendChild(regexNode)
  return settingNode;
}