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