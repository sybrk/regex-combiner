import { useState, type ChangeEvent, type DragEvent } from "react";


const FileInput = (props : { isMultiple: boolean, fileHandler: Function, description: string, fileType: string }) => {
  const { isMultiple, fileHandler, description, fileType } = props;
  const [isDragging, setIsDragging] = useState(false);
  

  const handleDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = async (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    const dataTransfer = e.dataTransfer;
    const files = isMultiple ? dataTransfer.files : (dataTransfer.files[0] ? dataTransfer.files[0] : null);
    await handleFiles(files);
  };

  const handleFileInput = async (e: ChangeEvent<HTMLInputElement>) => {
    const target = e.target as HTMLInputElement;
    const files = isMultiple ? target.files : (target.files?.[0] ? [target.files[0]] : null);
    await handleFiles(files);
  };

  const handleFiles = async (files: FileList | File[] | null | File) => {
    
    await fileHandler(files);
  };

  return (
    <>
      <div className="w-full max-w-lg mx-auto bg-neutral">
        <div
          className={`relative border-2 border-dashed rounded-lg p-4 text-center ${
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

          <div className="">
            <div className="flex flex-col items-center">

            </div>
            <div className="text-neutral-content/30">
              <p className="font-medium">{description}</p>
            </div>
          </div>
        </div>
      </div>
    </>
  );
};

export default FileInput;