import React, { useState, useMemo } from 'react';
import { FinalReportRow } from '../types';
import { Copy, Check, X, MessageCircle } from 'lucide-react';

interface ReminderGeneratorProps {
  data: FinalReportRow[];
  onClose: () => void;
}

const ReminderGenerator: React.FC<ReminderGeneratorProps> = ({ data, onClose }) => {
  const [selectedSupervisor, setSelectedSupervisor] = useState<string>('');
  const [copied, setCopied] = useState(false);

  // Get supervisors who have incomplete employees
  const incompleteData = useMemo(() => 
    data.filter(d => d["Completion Rate"] < 0.999), 
  [data]);

  const supervisors = useMemo(() => {
    const s = new Set(incompleteData.map(d => d["Supervisor Name"]));
    return Array.from(s).sort();
  }, [incompleteData]);

  const generatedMessage = useMemo(() => {
    if (!selectedSupervisor) return '';
    
    const employees = incompleteData.filter(d => d["Supervisor Name"] === selectedSupervisor);
    
    let msg = `Dear ${selectedSupervisor},\n\n`;
    msg += `Please note that the following employees under your supervision have not completed their assigned training:\n\n`;
    
    employees.forEach(emp => {
      const status = emp["Completion Rate"] === 0 ? "Not Started" : `${Math.round(emp["Completion Rate"] * 100)}%`;
      msg += `- ${emp["Display Name (Pharmacist name)"]} (${emp["User/Employee ID"]}) - Status: ${status}\n`;
    });

    msg += `\nPlease ensure they complete it by the deadline.\n\nBest Regards,\nLMS Admin`;
    return msg;
  }, [selectedSupervisor, incompleteData]);

  const handleCopy = () => {
    navigator.clipboard.writeText(generatedMessage);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4 animate-in fade-in duration-200">
      <div className="bg-white rounded-xl shadow-2xl w-full max-w-2xl max-h-[90vh] flex flex-col">
        <div className="p-6 border-b border-gray-100 flex justify-between items-center bg-[#F4A460] text-white rounded-t-xl">
          <h3 className="text-xl font-bold">Reminder Message Generator</h3>
          <button onClick={onClose} className="hover:bg-white/20 p-1 rounded-full transition-colors">
            <X className="w-6 h-6" />
          </button>
        </div>

        <div className="p-6 flex-1 overflow-y-auto">
          <div className="mb-4">
            <label className="block text-sm font-medium text-gray-700 mb-2">Select Supervisor</label>
            <select 
              className="w-full border border-gray-300 rounded-lg px-4 py-2 focus:ring-2 focus:ring-[#F4A460] outline-none"
              value={selectedSupervisor}
              onChange={(e) => setSelectedSupervisor(e.target.value)}
            >
              <option value="">-- Choose Supervisor --</option>
              {supervisors.map(s => (
                <option key={s} value={s}>{s} ({incompleteData.filter(d => d["Supervisor Name"] === s).length} pending)</option>
              ))}
            </select>
          </div>

          {selectedSupervisor && (
            <div className="bg-gray-50 p-4 rounded-lg border border-gray-200">
              <div className="flex justify-between items-center mb-2">
                 <span className="text-xs font-semibold text-gray-500 uppercase">Preview Message</span>
                 <div className="flex space-x-2">
                    <button 
                      onClick={handleCopy}
                      className={`flex items-center space-x-1 px-3 py-1 rounded text-xs font-bold transition-colors ${copied ? 'bg-green-100 text-green-700' : 'bg-white border text-gray-600 hover:bg-gray-100'}`}
                    >
                      {copied ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                      <span>{copied ? 'Copied!' : 'Copy Text'}</span>
                    </button>
                 </div>
              </div>
              <textarea 
                readOnly 
                value={generatedMessage}
                className="w-full h-64 bg-white p-3 rounded border border-gray-200 text-sm font-mono resize-none focus:outline-none"
              />
            </div>
          )}
        </div>

        <div className="p-4 border-t border-gray-100 bg-gray-50 rounded-b-xl flex justify-end">
           <button onClick={onClose} className="px-6 py-2 text-gray-600 font-medium hover:bg-gray-200 rounded-lg transition-colors">
             Close
           </button>
        </div>
      </div>
    </div>
  );
};

export default ReminderGenerator;