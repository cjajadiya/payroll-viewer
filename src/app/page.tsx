"use client";
import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';

export default function PayrollDashboard() {
  const [employees, setEmployees] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState('');

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    setIsProcessing(true);
    setError('');

    try {
      // 1. Client-Side parsing for immediate UI preview
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: 'array', cellDates: false });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true });
      
      // Mock parsing representation
      setEmployees([{ id: 1, name: "Smith, John", totalHours: 42.5 }]);

      // 2. Server-Side secure upload to API
      const formData = new FormData();
      formData.append('file', file);
      
      await fetch('/api/upload', {
        method: 'POST',
        body: formData
      });

    } catch (err) {
      setError('Failed to parse file structure. Verify POS export format.');
    } finally {
      setIsProcessing(false);
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
      </header>

      <main className="flex-1 max-w-7xl w-full mx-auto p-6 space-y-6">
        {employees.length === 0 && (
          <label className="flex flex-col items-center justify-center p-12 border-2 border-dashed border-slate-300 rounded-xl bg-white hover:border-emerald-500 transition cursor-pointer focus-within:ring-2 focus-within:ring-emerald-500">
            <span className="text-sm font-semibold text-emerald-600">Securely Upload Time Card Data</span>
            <span className="text-xs text-slate-500 mt-1">XLSX, XLS, or CSV</span>
            <input type="file" accept=".xlsx,.xls,.csv" className="sr-only" onChange={handleFileUpload} />
            {isProcessing && <div className="mt-4 text-xs font-medium text-emerald-600 animate-pulse">Processing Database Sync...</div>}
            {error && <div className="mt-4 text-xs font-medium text-rose-600" role="alert">{error}</div>}
          </label>
        )}

        {employees.length > 0 && (
          <div className="space-y-6">
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-sm">
                <div className="text-xs font-medium text-slate-500 uppercase tracking-wider">Total Store Hours</div>
                <div className="mt-1 text-2xl font-bold text-emerald-600">{totalStoreHours.toFixed(2)}</div>
              </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
              <table className="w-full text-left text-sm">
                <thead className="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-500 uppercase">
                  <tr>
                    <th className="px-6 py-3">Employee Name</th>
                    <th className="px-6 py-3 text-right">Total Hours</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {employees.map((emp) => (
                    <tr key={emp.id} className="hover:bg-slate-50 transition">
                      <td className="px-6 py-4 font-medium text-slate-900">{emp.name}</td>
                      <td className="px-6 py-4 text-right font-mono font-semibold">{emp.totalHours.toFixed(2)}</td>
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
