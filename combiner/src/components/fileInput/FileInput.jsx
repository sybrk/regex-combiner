import { useState } from "react";
import DownwardArrowIcon from "./DownwardArrow";


const FileInput = (props) => {
  const { isMultiple, fileHandler, description, fileType } = props;
  const [isDragging, setIsDragging] = useState(false);
  const [showDetails, setShowDetails] = useState(false);
  const [fileMessage, setFileMessage] = useState("");

  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = async (e) => {
    e.preventDefault();
    setIsDragging(false);

    const files = isMultiple ? e.dataTransfer.files : e.dataTransfer.files[0];
    await handleFiles(files);
  };

  const handleFileInput = async (e) => {
    const files = isMultiple ? e.target.files : e.target.files[0];
    await handleFiles(files);
  };

  const handleFiles = async (files) => {
    console.log("files", files);
    if(files.length) {
      setShowDetails(true);
      setFileMessage(`${files.length} file(s) selected`);
    }
    if(typeof files === "object" && files.length === undefined) {
      setShowDetails(true);
      setFileMessage(`${files.name} is selected`);
    }
    await fileHandler(files);
  };

  return (
    <>
      <div className="w-full max-w-md mx-auto">
        <div
          className={`relative border-2 border-dashed rounded-lg p-8 text-center ${
            isDragging ? "border-blue-500 bg-blue-50" : "border-gray-300"
          }`}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
        >
          <input
            id="modal_image_input"
            type="file"
            multiple={isMultiple}
            accept={fileType}
            onChange={handleFileInput}
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
          />

          <div className="pt-4">
            <div className="flex flex-col items-center"><DownwardArrowIcon /></div>
            <div className="text-gray-600">
              <p className="font-medium">{description}</p>
            </div>
            {
              showDetails &&
              <div className="mt-4 text-ellipsis">
                <p>{fileMessage}</p>
              </div>
            }
          </div>
        </div>
      </div>
    </>
  );
};

export default FileInput;