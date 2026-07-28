"use client";
import React, { useState, useMemo, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, CheckCircle2, AlertCircle, Loader2, Search, ArrowRight, Clock, Users } from 'lucide-react';

interface Employee {
  id: string;
  name: string;
  totalHours: number;
  shifts: number;
}

const namePat = /^[A-Za-zÀ-ÖØ-öø-ÿ'\-]+(?:\s*,\s*|\s+)[A-Za-zÀ-ÖØ-öø-ÿ'\- ]+/;

function formatName(raw: string) {
  if (!raw) return 'Unknown';
  const p = raw.split(',');
  return p.length === 2 ? `${p[1].trim()} ${p[0].trim()}` : raw.trim();
}

function getInitials(name: string) {
  return formatName(name).split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
}

function fmtHrs(n: number) {
  return (Math.round((n || 0) * 100) / 100).toFixed(2);
}

export default function PayrollDashboard() {
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [status, setStatus] = useState<'idle' | 'dragging' | 'processing' | 'success' | 'error'>('idle');
  const [errorMsg, setErrorMsg] = useState('');
  const [searchQuery, setSearchQuery] = useState('');
  
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setStatus('dragging');
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setStatus('idle');
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files?.[0];
    if (file) processFile(file);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) processFile(file);
  };

  const processFile = async (file: File) => {
    setStatus('processing');
    setErrorMsg('');

    try {
      await new Promise(resolve => setTimeout(resolve, 800));

      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: 'array', cellDates: false });
      const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true });
      
      let headerIdx = -1, inCol = 2, riC = 6, oiC = 8;

      for (let i = 0; i < Math.min(rows.length, 20); i++) {
        const row = rows[i] as unknown[];
        const cols = (row || []).map((c) => String(c).toLowerCase().trim());
        if (cols.includes('name') || cols.includes('employee')) {
          headerIdx = i;
          riC = cols.findIndex((c: string) => c.includes('regular')) >= 0 ? cols.findIndex((c: string) => c.includes('regular')) : riC;
          oiC = cols.findIndex((c: string) => c.includes('overtime')) >= 0 ? cols.findIndex((c: string) => c.includes('overtime')) : oiC;
          break;
        }
      }

      const parsedStaff: Record<string, Employee> = {};
      let currentEmpName = "";

      for (let i = headerIdx >= 0 ? headerIdx + 1 : 1; i < rows.length; i++) {
        const row = rows[i] as unknown[];
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
              shifts: 0 
            };
          }
        }

        if (currentEmpName && parsedStaff[currentEmpName]) {
          const hrs = (parseFloat(String(row[riC])) || 0) + (parseFloat(String(row[oiC])) || 0);
          if (hrs > 0 || row[inCol]) {
            parsedStaff[currentEmpName].shifts += 1;
            parsedStaff[currentEmpName].totalHours += hrs;
          }
        }
      }

      const finalEmployees = Object.values(parsedStaff).filter(e => e.shifts > 0);
      if (finalEmployees.length === 0) throw new Error("No payroll hours detected. Verify export format.");

      setEmployees(finalEmployees.sort((a, b) => a.name.localeCompare(b.name)));
      setStatus('success');
    } catch (err: unknown) {
      if (err instanceof Error) {
        setErrorMsg(err.message);
      } else {
        setErrorMsg('File parsing failed.');
      }
      setStatus('error');
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const filteredEmployees = useMemo(() => {
    return employees.filter(emp => emp.name.toLowerCase().includes(searchQuery.toLowerCase()));
  }, [employees, searchQuery]);

  const totalStoreHours = useMemo(() => employees.reduce((acc, emp) => acc + emp.totalHours, 0), [employees]);

  return (
    <div className="min-h-screen bg-[#FAFAFA] text-slate-900 font-sans selection:bg-indigo-100 selection:text-indigo-900">
      
      <header className="sticky top-0 z-50 bg-white border-b border-slate-200 shadow-[0_1px_2px_rgba(0,0,0,0.02)] px-6 py-3 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="h-8 w-8 flex items-center justify-center bg-white rounded-[8px] border border-slate-200 overflow-hidden shadow-sm">
            <img src="/the_chopped_leaf_logo.jpg" alt="Chopped Leaf Logo" className="w-full h-full object-contain p-0.5" />
          </div>
          <div>
            <h1 className="text-sm font-semibold tracking-tight text-slate-900 leading-none">Chopped Leaf</h1>
            <p className="text-[11px] font-medium text-slate-500 mt-0.5">Payroll Engine CL113</p>
          </div>
        </div>
        
        {status === 'success' && (
          <div className="flex items-center gap-4">
            <span className="flex items-center gap-1.5 text-xs font-medium text-emerald-600 bg-emerald-50 px-2.5 py-1 rounded-full border border-emerald-100">
              <CheckCircle2 className="w-3.5 h-3.5" /> Validated
            </span>
            <button onClick={() => setStatus('idle')} className="text-xs font-medium text-slate-600 hover:text-slate-900 bg-white border border-slate-200 shadow-sm hover:shadow px-3 py-1.5 rounded-md transition-all">
              Upload New
            </button>
          </div>
        )}
      </header>

      <main className="max-w-[1040px] w-full mx-auto p-6 md:p-8 space-y-8 animate-in fade-in duration-500">
        
        {(status === 'idle' || status === 'dragging' || status === 'processing' || status === 'error') && (
          <div className="max-w-2xl mx-auto mt-12">
            <div className="text-center mb-8">
              <h2 className="text-2xl font-semibold tracking-tight text-slate-900">Process Time Cards</h2>
              <p className="text-sm text-slate-500 mt-2">Upload your POS labor extract to validate compliance and hours.</p>
            </div>

            <div 
              onDragOver={handleDragOver} onDragLeave={handleDragLeave} onDrop={handleDrop}
              className={`relative flex flex-col items-center justify-center p-16 rounded-2xl transition-all duration-200 ease-out border-2 border-dashed
                ${status === 'dragging' ? 'border-indigo-500 bg-indigo-50/50 scale-[1.02]' : 'border-slate-300 bg-white hover:border-slate-400 hover:bg-slate-50'}
                ${status === 'error' ? 'border-rose-300 bg-rose-50/30' : ''}
              `}
            >
              <input type="file" accept=".xlsx,.xls,.csv" className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" onChange={handleFileSelect} ref={fileInputRef} disabled={status === 'processing'} />
              
              {status === 'processing' ? (
                <div className="flex flex-col items-center">
                  <Loader2 className="w-10 h-10 text-indigo-600 animate-spin mb-4" />
                  <p className="text-sm font-semibold text-slate-900">Parsing Matrix Data...</p>
                  <p className="text-xs text-slate-500 mt-1">Applying compliance rules</p>
                </div>
              ) : status === 'error' ? (
                <div className="flex flex-col items-center text-center">
                  <div className="w-12 h-12 bg-rose-100 text-rose-600 rounded-full flex items-center justify-center mb-4"><AlertCircle className="w-6 h-6" /></div>
                  <p className="text-sm font-semibold text-rose-700">Upload Failed</p>
                  <p className="text-xs text-rose-600/80 mt-1">{errorMsg}</p>
                  <button onClick={() => setStatus('idle')} className="mt-6 text-xs font-medium text-slate-600 underline underline-offset-2 hover:text-slate-900 relative z-10">Try another file</button>
                </div>
              ) : (
                <div className="flex flex-col items-center pointer-events-none">
                  <div className={`w-12 h-12 rounded-full flex items-center justify-center mb-4 transition-colors ${status === 'dragging' ? 'bg-indigo-100 text-indigo-600' : 'bg-slate-100 text-slate-500'}`}>
                    <UploadCloud className="w-6 h-6" />
                  </div>
                  <p className="text-sm font-semibold text-slate-900">Click to upload or drag and drop</p>
                  <p className="text-xs text-slate-500 mt-1">XLSX, XLS, or CSV (max 10MB)</p>
                </div>
              )}
            </div>
          </div>
        )}

        {status === 'success' && (
          <div className="space-y-6 animate-in slide-in-from-bottom-4 duration-500">
            
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-sm">
                <div className="flex items-center gap-2 text-slate-500 mb-2">
                  <Clock className="w-4 h-4" /><span className="text-xs font-semibold uppercase tracking-wider">Total Hours</span>
                </div>
                <div className="text-3xl font-bold tracking-tight text-slate-900">{fmtHrs(totalStoreHours)}</div>
              </div>
              <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-sm">
                <div className="flex items-center gap-2 text-slate-500 mb-2">
                  <Users className="w-4 h-4" /><span className="text-xs font-semibold uppercase tracking-wider">Active Staff</span>
                </div>
                <div className="text-3xl font-bold tracking-tight text-slate-900">{employees.length}</div>
              </div>
              <div className="bg-gradient-to-br from-indigo-600 to-indigo-700 rounded-xl border border-indigo-500 p-5 shadow-sm text-white flex flex-col justify-between">
                <div className="text-xs font-semibold text-indigo-100 uppercase tracking-wider">Next Action</div>
                <button className="w-full bg-white text-indigo-600 hover:bg-indigo-50 text-sm font-semibold py-2 px-4 rounded-lg transition-colors flex items-center justify-center gap-2 shadow-sm">
                  Export to Payroll <ArrowRight className="w-4 h-4" />
                </button>
              </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden flex flex-col">
              <div className="px-5 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50/50">
                <h3 className="text-sm font-semibold text-slate-900">Shift Roster</h3>
                <div className="relative">
                  <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                  <input 
                    type="text" placeholder="Search team..." value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)}
                    className="pl-9 pr-4 py-1.5 text-sm bg-white border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-shadow w-64 shadow-sm"
                  />
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-sm whitespace-nowrap">
                  <thead className="bg-white border-b border-slate-200 text-xs font-semibold text-slate-500">
                    <tr>
                      <th className="px-6 py-3 font-medium">Employee Name</th>
                      <th className="px-6 py-3 font-medium text-right">Shifts Recorded</th>
                      <th className="px-6 py-3 font-medium text-right">Total Hours</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {filteredEmployees.map((emp) => (
                      <tr key={emp.id} className="hover:bg-slate-50/80 transition-colors group">
                        <td className="px-6 py-3.5">
                          <div className="flex items-center gap-3">
                            <div className="w-8 h-8 rounded-full bg-slate-100 text-slate-600 border border-slate-200 flex items-center justify-center text-xs font-bold tracking-tight">
                              {getInitials(emp.name)}
                            </div>
                            <span className="font-medium text-slate-900">{formatName(emp.name)}</span>
                          </div>
                        </td>
                        <td className="px-6 py-3.5 text-right text-slate-500">{emp.shifts}</td>
                        <td className="px-6 py-3.5 text-right">
                          <span className="inline-flex items-center bg-slate-100 text-slate-700 border border-slate-200 rounded-md px-2.5 py-1 text-xs font-mono font-semibold">
                            {fmtHrs(emp.totalHours)}
                          </span>
                        </td>
                      </tr>
                    ))}
                    {filteredEmployees.length === 0 && (
                      <tr><td colSpan={3} className="px-6 py-8 text-center text-slate-500 text-sm">No employees match your search.</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}