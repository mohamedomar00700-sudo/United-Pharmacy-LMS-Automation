import React, { useRef, useState } from 'react';
import { Upload, CheckCircle, FileText, AlertCircle } from 'lucide-react';
import { FileType } from '../types';

interface FileUploadProps {
  type: FileType;
  onFileSelect: (type: FileType, file: File) => void;
  isUploaded: boolean;
  fileName?: string;
}

const FileUpload: React.FC<FileUploadProps> = ({ type, onFileSelect, isUploaded, fileName }) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [dragActive, setDragActive] = useState(false);

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setDragActive(true);
    } else if (e.type === 'dragleave') {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      onFileSelect(type, e.dataTransfer.files[0]);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      onFileSelect(type, e.target.files[0]);
    }
  };

  return (
    <div
      className={`relative flex flex-col items-center justify-center w-full h-32 border-2 border-dashed rounded-lg cursor-pointer transition-colors duration-200
        ${isUploaded ? 'border-green-500 bg-green-50' : 'border-gray-300 hover:border-[#F4A460] hover:bg-orange-50'}
        ${dragActive ? 'border-[#F4A460] bg-orange-50' : ''}
      `}
      onDragEnter={handleDrag}
      onDragLeave={handleDrag}
      onDragOver={handleDrag}
      onDrop={handleDrop}
      onClick={() => fileInputRef.current?.click()}
    >
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx, .xls, .csv"
        className="hidden"
        onChange={handleChange}
      />
      
      {isUploaded ? (
        <div className="flex flex-col items-center text-green-700 animate-in fade-in zoom-in duration-300">
          <CheckCircle className="w-8 h-8 mb-2" />
          <p className="text-sm font-semibold">{fileName}</p>
          <p className="text-xs opacity-75">Ready to process</p>
        </div>
      ) : (
        <div className="flex flex-col items-center text-gray-500">
          <Upload className="w-8 h-8 mb-2 text-gray-400" />
          <p className="text-sm font-semibold text-center px-2">{type}</p>
          <p className="text-xs opacity-75 mt-1">Drag & drop or click to upload</p>
        </div>
      )}
    </div>
  );
};

export default FileUpload;