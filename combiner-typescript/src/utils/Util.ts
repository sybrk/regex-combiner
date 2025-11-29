export type ReadFile = {
  fileName: string,
  fileContent: string
}

export const CONDITIONS = [
  "TargetAndSource",
  "TargetNotSource",
  "SourceNotTarget",
  "SourceOnly",
  "TargetOnly",
  "DifferentCount",
  "GroupedSourceNotTarget",
  "GroupedTargetAndSource"
] as const

export const IGNORE_CASE_OPTIONS = [
  "true",
  "false"
] as const

export type RegexRecord = {

  file: string,
  description: string,
  ignoreCase: typeof IGNORE_CASE_OPTIONS[number],
  source: string | null,
  sourceValid: boolean | null,
  target: string | null,
  targetValid: boolean | null,
  condition: typeof CONDITIONS[number],
  hasIssues: boolean | null

}

export type RegexRecordCollection = Record<string, RegexRecord>;




export const xmlParser = (xmlContent: string): Document => {
  let parser = new DOMParser();
  let xmldoc = parser.parseFromString(xmlContent, "text/xml");
  return xmldoc;
}

export const readFile = async (file: File): Promise<ReadFile> => {
  let reader = new FileReader();
  // read file as text
  reader.readAsText(file);
  // await the file to be read.
  await new Promise((resolve) => (reader.onload = () => resolve(null)));

  // return filename and filecontent with keys.
  return {
    fileName: file.name,
    fileContent: reader.result as string,
  };
};
export const idGenerator = (): string => {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'
    .replace(/[xy]/g, function (c) {
      const r = Math.random() * 16 | 0,
        v = c == 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
}

export const generateIdFiveChar = (): string => {
  const chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  let id = '';
  for (let i = 0; i < 5; i++) {
    id += chars[Math.floor(Math.random() * chars.length)];
  }
  return id;
}

export const newRegexFile = (): Document => {
  const xmlString = `<?xml version="1.0" encoding="utf-8" standalone="yes"?><SettingsBundle xmlns:ns1="http://schemas.datacontract.org/2004/07/Sdl.Verification.QAChecker.RegEx" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><SettingsGroup Id="QAVerificationSettings"></SettingsGroup></SettingsBundle>`;
  const newFile = xmlParser(xmlString);
  return newFile;
}

export const regexNodeBuilder = (regexObj: RegexRecord, id: string): Element => {
  const settingNode = document.createElementNS("", "Setting");
  settingNode.setAttribute("Id", id)
  const regexNode = document.createElementNS(null, "RegExRule");
  const descriptionNode = document.createElementNS(null, "Description");
  const caseNode = document.createElementNS(null, "IgnoreCase");
  const sourceNode = document.createElementNS(null, "RegExSource");
  const targetNode = document.createElementNS(null, "RegExTarget");
  const conditionNode = document.createElementNS(null, "RuleCondition");
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



export const regexParserObj = async (filesArr: File[]): Promise<RegexRecordCollection> => {
  const regexRecordCollection: RegexRecordCollection = {}
  let id = 1;
  for (let index = 0; index < filesArr.length; index++) {
    const file: File = filesArr[index];
    const fileRead: ReadFile = await readFile(file);
    const parseFile: Document = xmlParser(fileRead.fileContent)
    let regexRules: Element[] = Array.from(parseFile.querySelectorAll("RegExRule"));

    regexRules = regexRules.filter(x => /RegExRules\d+$/.test(x.parentElement?.getAttribute('Id') as string))
    //console.log("regexrules", regexRules, "hoo")
    regexRules.map((x) => {
      //console.log("processing", x)
      const regexId = id
      regexRecordCollection[regexId] = {
        file : fileRead.fileName,
        description: x.querySelector("Description")?.textContent || "",
        ignoreCase: (x.querySelector("IgnoreCase")?.textContent?.toLowerCase() || "false") as "true" | "false",
        source: x.querySelector("RegExSource")?.textContent ?? "",
        sourceValid: null,
        target: x.querySelector("RegExTarget")?.textContent ?? "",
        targetValid: null,
        condition: (x.querySelector("RuleCondition")?.textContent || "TargetAndSource") as "TargetAndSource" | "TargetNotSource" | "SourceNotTarget" | "SourceOnly" | "TargetOnly" | "DifferentCount" | "GroupedSourceNotTarget" | "GroupedTargetAndSource",
        hasIssues: null
      };
      id++
    });
  }
  return regexRecordCollection
}