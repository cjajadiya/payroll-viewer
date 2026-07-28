"use client";
import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';

// --- Pure Helper Functions ---
function robustNormalizeDate(val: any) {
  if (val == null || val === '') return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number' && !isNaN(val)) {
    return new Date(Date.UTC(1899, 11, 30) + val * 86400000);
  }
  const strVal = String(val).trim();
  const numVal = parseFloat(strVal);
  if (!isNaN(numVal) && isFinite(numVal) && numVal > 20000 && numVal < 60000) {
    return new Date(Date.UTC(1899, 11, 30) + numVal * 86400000);
  }
  const d = new Date(strVal);
  if (!isNaN(d.getTime())) return d;
  return null;
}

function formatName(raw: string) {
  if (!raw) return 'Unknown';
  const p = raw.split(',');
  return p.length === 2 ? p[1].trim() + ' ' + p[0].trim() : raw.trim();
}

function fmtHrs(n: number) {
  return (Math.round((n || 0) * 100) / 100).toFixed(2);
}

const namePat = /^[A-Za-zÀ-ÖØ-öø-ÿ'\-]+(?:\s*,\s*|\s+)[A-Za-zÀ-ÖØ-öø-ÿ'\- ]+/;

// --- Main Component ---
export default function PayrollDashboard() {
  const [employees, setEmployees] = useState<any[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    setIsProcessing(true);
    setError('');

    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: 'array', cellDates: false });
      const rows = XLSX.utils.sheet_to_json<any[]>(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true });
      
      let headerIdx = -1;
      let inCol = 2, outCol = 3, riC = 6, oiC = 8;

      // Locate Headers
      for (let i = 0; i < Math.min(rows.length, 20); i++) {
        const cols = (rows[i] || []).map((c: any) => String(c).toLowerCase().trim());
        if (cols.includes('name') || cols.includes('employee')) {
          headerIdx = i;
          const rIdx = cols.findIndex((c: string) => c.includes('regular'));
          const oIdx = cols.findIndex((c: string) => c.includes('overtime'));
          if (rIdx >= 0) riC = rIdx;
          if (oIdx >= 0) oiC = oIdx;
          break;
        }
      }

      const parsedStaff: Record<string, any> = {};
      let currentEmpName = "";

      // Parse Data Rows
      const startIdx = headerIdx >= 0 ? headerIdx + 1 : 1;
      for (let i = startIdx; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0) continue;

        const firstCol = String(row[0] || '').trim();
        if (!firstCol || firstCol === 'Total' || firstCol.startsWith('If you are using')) continue;

        if (namePat.test(firstCol)) {
          currentEmpName = firstCol;
          if (!parsedStaff[currentEmpName]) {
            parsedStaff[currentEmpName] = { 
              id: `emp_${Object.keys(parsedStaff).length}`, 
              name: currentEmpName, 
              totalHours: 0,
              shifts: [] 
            };
          }
        }

        if (currentEmpName && parsedStaff[currentEmpName]) {
          const ciDate = robustNormalizeDate(row[inCol]);
          const parsedReg = parseFloat(row[riC]) || 0;
          const parsedOt  = parseFloat(row[oiC]) || 0;
          const hrs = parsedReg + parsedOt;

          if (ciDate || hrs > 0) {
            parsedStaff[currentEmpName].shifts.push({ hrs });
            parsedStaff[currentEmpName].totalHours += hrs;
          }
        }
      }

      const finalEmployees = Object.values(parsedStaff)
        .filter(e => e.shifts.length > 0)
        .sort((a, b) => a.name.localeCompare(b.name));

      if (finalEmployees.length === 0) {
        throw new Error("No hours found. Verify file matches POS format.");
      }

      setEmployees(finalEmployees);
    } catch (err: any) {
      setError(err.message || 'Failed to parse file structure.');
    } finally {
      setIsProcessing(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const totalStoreHours = useMemo(() => {
    return employees.reduce((acc, emp) => acc + emp.totalHours, 0);
  }, [employees]);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans flex flex-col">
      <header className="sticky top-0 z-30 bg-white/95 backdrop-blur border-b border-slate-200 shadow-sm px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="h-10 w-10 flex items-center justify-center bg-emerald-600 text-white font-bold rounded-lg shadow-sm">CL</div>
          <div>
            <h1 className="text-base font-semibold leading-tight text-slate-900">The Chopped Leaf</h1>
            <p className="text-xs text-slate-500">Enterprise Payroll Dashboard</p>
          </div>
        </div>
        {employees.length > 0 && (
          <button onClick={() => setEmployees([])} className="text-xs font-semibold text-slate-500 hover:text-slate-800 transition">
            Start Over
          </button>
        )}
      </header>

      <main className="flex-1 max-w-5xl w-full mx-auto p-6 space-y-6">
        {employees.length === 0 && (
          <label className="flex flex-col items-center justify-center p-12 mt-12 border-2 border-dashed border-slate-300 rounded-xl bg-white hover:border-emerald-500 transition cursor-pointer focus-within:ring-2 focus-within:ring-emerald-500 shadow-sm">
            <div className="rounded-full bg-emerald-50 p-3 text-emerald-600 mb-4">
              <svg className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2">
                <path strokeLinecap="round" strokeLinejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/>
              </svg>
            </div>
            <span className="text-base font-semibold text-slate-700">Upload Time Card Data</span>
            <span className="text-sm text-slate-500 mt-1">Drag and drop your XLSX, XLS, or CSV</span>
            <input type="file" accept=".xlsx,.xls,.csv" className="sr-only" onChange={handleFileUpload} ref={fileInputRef} />
            {isProcessing && <div className="mt-4 text-sm font-medium text-emerald-600 animate-pulse">Processing...</div>}
            {error && <div className="mt-4 text-sm font-medium text-rose-600 bg-rose-50 px-4 py-2 rounded-md border border-rose-100">{error}</div>}
          </label>
        )}

        {employees.length > 0 && (
          <div className="space-y-6 animate-in fade-in duration-500">
            <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-sm flex justify-between items-center">
              <div>
                <div className="text-xs font-bold text-slate-400 uppercase tracking-wider">Total Store Hours</div>
                <div className="text-sm text-slate-500 mt-0.5">Calculated across {employees.length} employees</div>
              </div>
              <div className="text-3xl font-bold text-emerald-600 font-mono tracking-tight">{totalStoreHours.toFixed(2)}</div>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
              <table className="w-full text-left text-sm border-collapse">
                <thead className="bg-slate-50 border-b border-slate-200 text-xs font-bold text-slate-500 uppercase tracking-wider">
                  <tr>
                    <th className="px-6 py-4">Employee Name</th>
                    <th className="px-6 py-4 text-right">Total Hours</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {employees.map((emp) => (
                    <tr key={emp.id} className="hover:bg-slate-50 transition group cursor-default">
                      <td className="px-6 py-4 font-semibold text-slate-800">{formatName(emp.name)}</td>
                      <td className="px-6 py-4 text-right font-mono font-bold text-slate-700">{fmtHrs(emp.totalHours)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
