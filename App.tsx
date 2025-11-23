import React, { useState } from 'react';
import { Pill, Database, BarChart3, RotateCcw, Loader2 } from 'lucide-react';
import FileUpload from './components/FileUpload';
import Dashboard from './components/Dashboard';
import { FileType, LMSDataRow, MasterDataRow, FinalReportRow } from './types';
import { parseExcel, processData } from './services/excelService';

const App: React.FC = () => {
  const [files, setFiles] = useState<Record<FileType, File | null>>({
    [FileType.LMS1]: null,
    [FileType.LMS2]: null,
    [FileType.MASTER]: null,
  });

  const [finalData, setFinalData] = useState<FinalReportRow[] | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleFileSelect = (type: FileType, file: File) => {
    setFiles((prev) => ({ ...prev, [type]: file }));
    setError(null);
  };

  const handleProcess = async () => {
    if (!files[FileType.LMS1] || !files[FileType.LMS2] || !files[FileType.MASTER]) {
      setError("Please upload all 3 required files before processing.");
      return;
    }

    setIsProcessing(true);
    setError(null);

    try {
      // 1. Parse all files
      const lms1Data = await parseExcel<LMSDataRow>(files[FileType.LMS1]!);
      const lms2Data = await parseExcel<LMSDataRow>(files[FileType.LMS2]!);
      const masterData = await parseExcel<MasterDataRow>(files[FileType.MASTER]!);

      // 2. Process logic
      const result = processData(lms1Data, lms2Data, masterData);
      
      // Artificial delay for UX (so user sees the spinner)
      setTimeout(() => {
        setFinalData(result);
        setIsProcessing(false);
      }, 800);

    } catch (err: any) {
      console.error(err);
      setError("Error processing files. Please ensure the column names match the requirements exactly.");
      setIsProcessing(false);
    }
  };

  const handleReset = () => {
    setFiles({
      [FileType.LMS1]: null,
      [FileType.LMS2]: null,
      [FileType.MASTER]: null,
    });
    setFinalData(null);
    setError(null);
  };

  const allFilesUploaded = files[FileType.LMS1] && files[FileType.LMS2] && files[FileType.MASTER];

  return (
    <div className="min-h-screen bg-gray-50 text-gray-800 pb-12">
      {/* Header */}
      <header className="bg-[#F4A460] text-white shadow-lg sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 py-4 sm:px-6 lg:px-8 flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="bg-white/20 p-2 rounded-lg">
               <Pill className="w-8 h-8 text-white" />
            </div>
            <div>
              <h1 className="text-2xl font-bold tracking-tight">United Pharmacy</h1>
              <p className="text-orange-100 text-sm font-medium">LMS Automation Tool</p>
            </div>
          </div>
          <div className="hidden md:flex items-center space-x-2 text-sm bg-white/10 px-3 py-1 rounded-full">
            <Database className="w-4 h-4" />
            <span>v1.0.0 Stable</span>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8 sm:px-6 lg:px-8">
        
        {/* Intro */}
        {!finalData && (
          <div className="text-center mb-10 space-y-2">
            <h2 className="text-3xl font-bold text-gray-900">Data Processor</h2>
            <p className="text-gray-500 max-w-2xl mx-auto">
              Upload your raw LMS exports and Master Sheet below. The system will automatically merge, clean, and analyze the completion rates.
            </p>
          </div>
        )}

        {/* Upload Section */}
        {!finalData && (
          <div className="bg-white rounded-2xl shadow-xl p-8 mb-8 border border-gray-100">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
              <FileUpload 
                type={FileType.LMS1} 
                onFileSelect={handleFileSelect} 
                isUploaded={!!files[FileType.LMS1]} 
                fileName={files[FileType.LMS1]?.name} 
              />
              <FileUpload 
                type={FileType.LMS2} 
                onFileSelect={handleFileSelect} 
                isUploaded={!!files[FileType.LMS2]} 
                fileName={files[FileType.LMS2]?.name} 
              />
              <FileUpload 
                type={FileType.MASTER} 
                onFileSelect={handleFileSelect} 
                isUploaded={!!files[FileType.MASTER]} 
                fileName={files[FileType.MASTER]?.name} 
              />
            </div>

            {error && (
              <div className="bg-red-50 border-l-4 border-red-500 p-4 mb-6 rounded-r">
                <p className="text-red-700 font-medium">{error}</p>
              </div>
            )}

            <div className="flex justify-center">
              <button
                onClick={handleProcess}
                disabled={!allFilesUploaded || isProcessing}
                className={`
                  flex items-center space-x-2 px-8 py-4 rounded-xl font-bold text-lg transition-all transform hover:scale-105
                  ${allFilesUploaded && !isProcessing 
                    ? 'bg-[#F4A460] hover:bg-orange-500 text-white shadow-lg hover:shadow-orange-200' 
                    : 'bg-gray-200 text-gray-400 cursor-not-allowed'}
                `}
              >
                {isProcessing ? (
                  <>
                    <Loader2 className="w-6 h-6 animate-spin" />
                    <span>Processing Data...</span>
                  </>
                ) : (
                  <>
                    <BarChart3 className="w-6 h-6" />
                    <span>Run Automation</span>
                  </>
                )}
              </button>
            </div>
          </div>
        )}

        {/* Dashboard Section */}
        {finalData && (
          <div>
            <div className="flex justify-between items-center mb-6">
              <h2 className="text-2xl font-bold text-gray-800">Analytics Dashboard</h2>
              <button 
                onClick={handleReset}
                className="flex items-center space-x-2 text-gray-500 hover:text-red-500 transition-colors px-4 py-2 rounded-lg hover:bg-red-50"
              >
                <RotateCcw className="w-4 h-4" />
                <span>Start Over</span>
              </button>
            </div>
            <Dashboard data={finalData} />
          </div>
        )}
      </main>
    </div>
  );
};

export default App;