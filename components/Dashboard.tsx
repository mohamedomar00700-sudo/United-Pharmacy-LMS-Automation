import React, { useMemo, useState } from 'react';
import { FinalReportRow, ProcessedStats } from '../types';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer, 
  PieChart, Pie, Cell 
} from 'recharts';
import { Download, Filter } from 'lucide-react';
import { generateExcelFile } from '../services/excelService';

interface DashboardProps {
  data: FinalReportRow[];
}

const Dashboard: React.FC<DashboardProps> = ({ data }) => {
  const [districtFilter, setDistrictFilter] = useState<string>('All');
  const [supervisorFilter, setSupervisorFilter] = useState<string>('All');

  // Unique values for filters
  const districts = useMemo(() => ['All', ...Array.from(new Set(data.map(d => d.District))).sort()], [data]);
  const supervisors = useMemo(() => ['All', ...Array.from(new Set(data.map(d => d["Supervisor Name"]))).sort()], [data]);

  // Filtered Data
  const filteredData = useMemo(() => {
    return data.filter(row => {
      const matchDistrict = districtFilter === 'All' || row.District === districtFilter;
      const matchSupervisor = supervisorFilter === 'All' || row["Supervisor Name"] === supervisorFilter;
      return matchDistrict && matchSupervisor;
    });
  }, [data, districtFilter, supervisorFilter]);

  // Statistics Calculation
  const stats: ProcessedStats = useMemo(() => {
    const totalPharmacists = filteredData.length;
    
    // Status Logic: 
    // Data comes as Decimal (0.0 - 1.0)
    // 1.0 (or close to it) = Completed
    // 0 < x < 1.0 = In Progress
    // 0 = Not Started
    const totalCompleted = filteredData.filter(d => d["Completion Rate"] >= 0.999).length;
    const totalInProgress = filteredData.filter(d => d["Completion Rate"] > 0 && d["Completion Rate"] < 0.999).length;
    const totalNotStarted = filteredData.filter(d => d["Completion Rate"] === 0).length;
    
    const completionPercentage = totalPharmacists > 0 ? (totalCompleted / totalPharmacists) * 100 : 0;

    // By District for Chart
    const districtCounts: Record<string, { total: number; completed: number }> = {};
    filteredData.forEach(row => {
      if (!districtCounts[row.District]) districtCounts[row.District] = { total: 0, completed: 0 };
      districtCounts[row.District].total++;
      if (row["Completion Rate"] >= 0.999) districtCounts[row.District].completed++;
    });

    const byDistrict = Object.entries(districtCounts).map(([name, val]) => ({
      name: name || 'Unknown',
      Completed: val.completed,
      Pending: val.total - val.completed
    })).sort((a, b) => b.Completed - a.Completed).slice(0, 10); // Top 10

    // Completion Distribution for Pie
    const byStatus = [
        { name: 'Completed', value: totalCompleted },
        { name: 'In Progress', value: totalInProgress },
        { name: 'Not Started', value: totalNotStarted }
    ];

    // Extended stats for the table
    return {
      totalPharmacists,
      totalCompleted,
      completionPercentage,
      byDistrict,
      bySupervisor: [], 
      byStatus,
      // Add extra stats for the summary table
      totalInProgress,
      totalNotStarted
    } as any; 
  }, [filteredData]);

  // Format percent helper
  const fmtPct = (val: number, total: number) => total > 0 ? Math.round((val / total) * 100) + '%' : '0%';

  return (
    <div className="space-y-6 animate-in slide-in-from-bottom duration-500">
      
      {/* Action Bar */}
      <div className="flex flex-col md:flex-row justify-between items-center bg-white p-4 rounded-xl shadow-sm border border-gray-100">
        <div className="flex items-center space-x-4 w-full md:w-auto mb-4 md:mb-0">
          <div className="flex items-center space-x-2">
            <Filter className="w-5 h-5 text-[#F4A460]" />
            <span className="font-semibold text-gray-700">Filters:</span>
          </div>
          
          <select 
            className="border rounded-md px-3 py-1.5 text-sm bg-gray-50 focus:ring-2 focus:ring-[#F4A460] outline-none"
            value={districtFilter}
            onChange={(e) => setDistrictFilter(e.target.value)}
          >
            {districts.map(d => <option key={d} value={d}>{d === 'All' ? 'All Districts' : d}</option>)}
          </select>

          <select 
            className="border rounded-md px-3 py-1.5 text-sm bg-gray-50 focus:ring-2 focus:ring-[#F4A460] outline-none"
            value={supervisorFilter}
            onChange={(e) => setSupervisorFilter(e.target.value)}
          >
            {supervisors.map(s => <option key={s} value={s}>{s === 'All' ? 'All Supervisors' : s}</option>)}
          </select>
        </div>

        <button 
          onClick={() => generateExcelFile(filteredData)}
          className="flex items-center space-x-2 bg-[#F4A460] hover:bg-orange-500 text-white px-6 py-2 rounded-lg font-medium transition-colors shadow-md w-full md:w-auto justify-center"
        >
          <Download className="w-5 h-5" />
          <span>Download Final Report</span>
        </button>
      </div>

      {/* NEW: Summarization Status Table (Matching Screenshot) */}
      <div className="bg-white rounded-xl shadow-sm overflow-hidden border border-gray-200">
        <div className="bg-[#F4A460] text-white px-6 py-3 font-bold text-lg">
          Summarization Status
        </div>
        <div className="overflow-x-auto">
          <table className="min-w-full text-sm">
            <thead>
              <tr className="bg-[#F4A460] text-white border-t border-orange-300">
                <th className="px-6 py-2 text-left w-1/2">Summarization Status</th>
                <th className="px-6 py-2 text-center w-1/4">Number Count</th>
                <th className="px-6 py-2 text-center w-1/4">Percentage</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {/* Total Row */}
              <tr className="hover:bg-orange-50">
                <td className="px-6 py-3 font-medium text-gray-700">Number Of Student</td>
                <td className="px-6 py-3 text-center font-bold text-[#F4A460]">{stats.totalPharmacists}</td>
                <td className="px-6 py-3 text-center text-gray-600">100%</td>
              </tr>
              {/* Completed Row */}
              <tr className="hover:bg-orange-50">
                <td className="px-6 py-3 font-medium text-gray-700">Number Of Student Who Completed</td>
                <td className="px-6 py-3 text-center font-bold text-green-600">{(stats as any).totalCompleted}</td>
                <td className="px-6 py-3 text-center text-gray-600">{fmtPct((stats as any).totalCompleted, stats.totalPharmacists)}</td>
              </tr>
              {/* In Progress Row */}
              <tr className="hover:bg-orange-50">
                <td className="px-6 py-3 font-medium text-gray-700">Number OF Student Who In Progress</td>
                <td className="px-6 py-3 text-center font-bold text-orange-500">{(stats as any).totalInProgress}</td>
                <td className="px-6 py-3 text-center text-gray-600">{fmtPct((stats as any).totalInProgress, stats.totalPharmacists)}</td>
              </tr>
              {/* Not Started Row */}
              <tr className="hover:bg-orange-50">
                <td className="px-6 py-3 font-medium text-gray-700">Number Of Students Who Did Not Started</td>
                <td className="px-6 py-3 text-center font-bold text-gray-400">{(stats as any).totalNotStarted}</td>
                <td className="px-6 py-3 text-center text-gray-600">{fmtPct((stats as any).totalNotStarted, stats.totalPharmacists)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      {/* Charts Area */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Bar Chart: Completion by District */}
        <div className="lg:col-span-2 bg-white p-6 rounded-xl shadow-sm">
          <h3 className="text-lg font-bold text-gray-800 mb-4">Completion by District (Top 10)</h3>
          <div className="h-80">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={stats.byDistrict} layout="vertical" margin={{ top: 5, right: 30, left: 20, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" horizontal={true} vertical={false} />
                <XAxis type="number" />
                <YAxis dataKey="name" type="category" width={100} tick={{fontSize: 12}} />
                <RechartsTooltip />
                <Legend />
                <Bar dataKey="Completed" stackId="a" fill="#10B981" radius={[0, 4, 4, 0]} />
                <Bar dataKey="Pending" stackId="a" fill="#E5E7EB" radius={[0, 4, 4, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Pie Chart: Overall Status */}
        <div className="bg-white p-6 rounded-xl shadow-sm">
          <h3 className="text-lg font-bold text-gray-800 mb-4">Overall Status</h3>
          <div className="h-80 flex items-center justify-center">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={stats.byStatus}
                  cx="50%"
                  cy="50%"
                  innerRadius={60}
                  outerRadius={100}
                  fill="#8884d8"
                  paddingAngle={5}
                  dataKey="value"
                >
                  <Cell fill="#10B981" /> {/* Completed */}
                  <Cell fill="#F4A460" /> {/* In Progress */}
                  <Cell fill="#E5E7EB" /> {/* Not Started */}
                </Pie>
                <RechartsTooltip />
                <Legend verticalAlign="bottom" height={36}/>
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* Preview Table */}
      <div className="bg-white rounded-xl shadow-sm overflow-hidden">
        <div className="px-6 py-4 border-b border-gray-100 bg-gray-50">
          <h3 className="font-bold text-lg text-gray-800">Detailed Report Preview</h3>
        </div>
        <div className="overflow-x-auto max-h-96">
          <table className="min-w-full text-left text-sm whitespace-nowrap">
            <thead className="bg-gray-100 text-gray-600 font-semibold sticky top-0">
              <tr>
                <th className="px-4 py-3">Employee ID</th>
                <th className="px-4 py-3">Pharmacist Name</th>
                <th className="px-4 py-3">District</th>
                <th className="px-4 py-3">Supervisor</th>
                <th className="px-4 py-3 text-center">Rate</th>
                <th className="px-4 py-3 text-center">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filteredData.slice(0, 50).map((row, idx) => {
                // Convert decimal to integer for display (0.5 -> 50)
                const displayRate = Math.round(row["Completion Rate"] * 100);
                
                return (
                  <tr key={`${row["User/Employee ID"]}-${idx}`} className="hover:bg-gray-50">
                    <td className="px-4 py-3 font-mono text-gray-500">{row["User/Employee ID"]}</td>
                    <td className="px-4 py-3 font-medium text-gray-900">{row["Display Name (Pharmacist name)"]}</td>
                    <td className="px-4 py-3 text-gray-600">{row.District}</td>
                    <td className="px-4 py-3 text-gray-600">{row["Supervisor Name"]}</td>
                    <td className="px-4 py-3 text-center font-bold">
                      <span className={`${displayRate === 100 ? 'text-green-600' : displayRate > 0 ? 'text-orange-500' : 'text-gray-400'}`}>
                        {displayRate}%
                      </span>
                    </td>
                    <td className="px-4 py-3 text-center">
                      {displayRate === 100 ? (
                        <span className="bg-green-100 text-green-700 px-2 py-1 rounded-full text-xs">Completed</span>
                      ) : displayRate > 0 ? (
                        <span className="bg-orange-100 text-orange-700 px-2 py-1 rounded-full text-xs">In Progress</span>
                      ) : (
                        <span className="bg-gray-100 text-gray-600 px-2 py-1 rounded-full text-xs">Not Started</span>
                      )}
                    </td>
                  </tr>
                );
              })}
              {filteredData.length > 50 && (
                <tr>
                  <td colSpan={6} className="px-4 py-3 text-center text-gray-500 italic bg-gray-50">
                    ... and {filteredData.length - 50} more rows (Download to see full report)
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default Dashboard;