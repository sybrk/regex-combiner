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