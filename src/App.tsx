/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  FileSpreadsheet, 
  X, 
  CheckCircle2, 
  AlertCircle, 
  Download,
  Layers,
  FileText,
  Loader2,
  ChevronRight
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from '@/src/lib/utils';

interface ExcelFile {
  id: string;
  name: string;
  size: number;
  data: ArrayBuffer;
  sheets: string[];
}

interface MergedResult {
  sheetName: string;
  data: any[][];
}

export default function App() {
  const [files, setFiles] = useState<ExcelFile[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [mergedResults, setMergedResults] = useState<MergedResult[]>([]);
  const [error, setError] = useState<string | null>(null);

  const onFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = e.target.files;
    if (!selectedFiles) return;
    await processFiles(Array.from(selectedFiles));
  };

  const processFiles = async (fileList: File[]) => {
    setIsProcessing(true);
    setError(null);
    const newFiles: ExcelFile[] = [];

    for (const file of fileList) {
      try {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        newFiles.push({
          id: Math.random().toString(36).substring(7),
          name: file.name,
          size: file.size,
          data: buffer,
          sheets: workbook.SheetNames,
        });
      } catch (err) {
        console.error(`Error reading ${file.name}:`, err);
        setError(`Failed to read ${file.name}. Please ensure it's a valid Excel file.`);
      }
    }

    setFiles(prev => [...prev, ...newFiles]);
    setIsProcessing(false);
  };

  const removeFile = (id: string) => {
    setFiles(prev => prev.filter(f => f.id !== id));
    setMergedResults([]);
  };

  const allSheets = useMemo(() => {
    if (files.length === 0) return [];
    
    const uniqueSheets = new Set<string>();
    files.forEach(file => {
      file.sheets.forEach(sheet => uniqueSheets.add(sheet));
    });

    return Array.from(uniqueSheets);
  }, [files]);

  const mergeData = useCallback(() => {
    if (files.length === 0) return;
    setIsProcessing(true);
    
    try {
      const results: MergedResult[] = [];

      allSheets.forEach(sheetName => {
        let combinedData: any[][] = [];
        let headers: any[] = [];

        files.forEach(file => {
          if (file.sheets.includes(sheetName)) {
            const workbook = XLSX.read(file.data, { type: 'array' });
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

            if (jsonData.length > 0) {
              if (headers.length === 0) {
                headers = jsonData[0];
                combinedData.push(headers);
              }
              
              // Append data rows (skipping header and any row matching the header)
              const rows = jsonData.slice(1).filter(row => {
                // Robust header check: trim and compare up to the length of headers
                const isHeader = row.length > 0 && headers.length > 0 && 
                  headers.every((hVal, i) => {
                    const rVal = row[i];
                    return String(rVal || '').trim() === String(hVal || '').trim();
                  });
                
                // Also skip completely empty rows
                const isEmpty = row.every(val => val === null || val === undefined || val === '');
                
                return !isHeader && !isEmpty;
              });
              
              combinedData = [...combinedData, ...rows];
            }
          }
        });

        if (combinedData.length > 1) {
          results.push({ sheetName, data: combinedData });
        }
      });

      if (results.length === 0) {
        setError('No data rows found in any of the identified tabs.');
      }

      setMergedResults(results);
    } catch (err) {
      console.error('Merge error:', err);
      setError('An error occurred during merging. Please check your file formats.');
    } finally {
      setIsProcessing(false);
    }
  }, [files, allSheets]);

  const downloadMerged = () => {
    if (mergedResults.length === 0) return;

    const newWorkbook = XLSX.utils.book_new();
    mergedResults.forEach(result => {
      const worksheet = XLSX.utils.aoa_to_sheet(result.data);
      XLSX.utils.book_append_sheet(newWorkbook, worksheet, result.sheetName);
    });

    XLSX.writeFile(newWorkbook, 'merged_excel_data.xlsx');
  };

  const formatSize = (bytes: number) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  return (
    <div className="min-h-screen bg-[#f5f5f5] text-[#1a1a1a] font-sans p-6 md:p-12">
      <div className="max-w-4xl mx-auto">
        {/* Header */}
        <header className="mb-12">
          <div className="flex items-center gap-3 mb-2">
            <div className="p-2 bg-black rounded-lg">
              <Layers className="w-6 h-6 text-white" />
            </div>
            <h1 className="text-2xl font-semibold tracking-tight">Excel Tab Merger</h1>
          </div>
          <p className="text-muted-foreground">
            Consolidate data from multiple Excel files by merging identical tabs into a single workbook.
          </p>
        </header>

        <div className="grid grid-cols-1 gap-8">
          {/* Upload Section */}
          <section className="bg-white rounded-3xl p-8 shadow-sm border border-black/5">
            <div 
              className="border-2 border-dashed border-black/10 rounded-2xl p-12 flex flex-col items-center justify-center transition-colors hover:border-black/20 group cursor-pointer relative"
              onDragOver={(e) => e.preventDefault()}
              onDrop={async (e) => {
                e.preventDefault();
                await processFiles(Array.from(e.dataTransfer.files));
              }}
            >
              <input 
                type="file" 
                multiple 
                accept=".xlsx, .xls" 
                className="absolute inset-0 opacity-0 cursor-pointer" 
                onChange={onFileChange}
              />
              <div className="w-16 h-16 bg-[#f5f5f5] rounded-full flex items-center justify-center mb-4 group-hover:scale-110 transition-transform">
                <Upload className="w-8 h-8 text-black/40" />
              </div>
              <h3 className="text-lg font-medium mb-1">Upload Excel Sheets</h3>
              <p className="text-sm text-muted-foreground">Drag and drop or click to select files (.xlsx, .xls)</p>
            </div>

            {/* File List */}
            <AnimatePresence>
              {files.length > 0 && (
                <motion.div 
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="mt-8 space-y-3"
                >
                  <div className="flex justify-between items-center mb-4">
                    <h4 className="text-xs font-semibold uppercase tracking-wider text-black/40">
                      Uploaded Files ({files.length})
                    </h4>
                    {files.length > 1 && (
                      <button 
                        onClick={() => setFiles([])}
                        className="text-xs font-medium text-red-500 hover:text-red-600 transition-colors"
                      >
                        Clear All
                      </button>
                    )}
                  </div>
                  {files.map((file) => (
                    <motion.div 
                      key={file.id}
                      layout
                      initial={{ opacity: 0, scale: 0.98 }}
                      animate={{ opacity: 1, scale: 1 }}
                      exit={{ opacity: 0, scale: 0.98 }}
                      className="flex items-center justify-between p-4 bg-[#f9f9f9] rounded-xl border border-black/5"
                    >
                      <div className="flex items-center gap-3 overflow-hidden">
                        <div className="p-2 bg-green-50 rounded-lg">
                          <FileSpreadsheet className="w-5 h-5 text-green-600" />
                        </div>
                        <div className="overflow-hidden">
                          <p className="text-sm font-medium truncate">{file.name}</p>
                          <p className="text-[10px] text-black/40 font-mono uppercase">{formatSize(file.size)} • {file.sheets.length} SHEETS</p>
                        </div>
                      </div>
                      <button 
                        onClick={() => removeFile(file.id)}
                        className="p-1 hover:bg-black/5 rounded-full transition-colors"
                      >
                        <X className="w-4 h-4 text-black/40" />
                      </button>
                    </motion.div>
                  ))}
                </motion.div>
              )}
            </AnimatePresence>
          </section>

          {/* Analysis & Actions */}
          {files.length >= 1 && (
            <section className="bg-white rounded-3xl p-8 shadow-sm border border-black/5">
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
                <div>
                  <h3 className="text-lg font-medium mb-1">Tabs Identified</h3>
                  <p className="text-sm text-muted-foreground">
                    {allSheets.length > 0 
                      ? `Found ${allSheets.length} unique tabs across your files.` 
                      : "No tabs found in the uploaded files."}
                  </p>
                  
                  {allSheets.length > 0 && (
                    <div className="flex flex-wrap gap-2 mt-4">
                      {allSheets.map(sheet => (
                        <span key={sheet} className="px-3 py-1 bg-[#f5f5f5] border border-black/5 rounded-full text-[11px] font-medium">
                          {sheet}
                        </span>
                      ))}
                    </div>
                  )}
                </div>

                <div className="flex flex-col gap-3">
                  <button
                    disabled={allSheets.length === 0 || isProcessing}
                    onClick={mergeData}
                    className={cn(
                      "px-8 py-3 rounded-2xl font-medium transition-all flex items-center justify-center gap-2",
                      allSheets.length > 0 
                        ? "bg-black text-white hover:bg-black/90 active:scale-95" 
                        : "bg-black/5 text-black/20 cursor-not-allowed"
                    )}
                  >
                    {isProcessing ? (
                      <Loader2 className="w-4 h-4 animate-spin" />
                    ) : (
                      <Layers className="w-4 h-4" />
                    )}
                    Merge All Tabs
                  </button>
                </div>
              </div>

              {/* Error Message */}
              {error && (
                <div className="mt-6 p-4 bg-red-50 border border-red-100 rounded-2xl flex items-center gap-3 text-red-600">
                  <AlertCircle className="w-5 h-5 flex-shrink-0" />
                  <p className="text-sm">{error}</p>
                </div>
              )}

              {/* Results Section */}
              <AnimatePresence>
                {mergedResults.length > 0 && (
                  <motion.div 
                    initial={{ opacity: 0, height: 0 }}
                    animate={{ opacity: 1, height: 'auto' }}
                    className="mt-8 pt-8 border-t border-black/5"
                  >
                    <div className="flex items-center justify-between mb-6">
                      <div className="flex items-center gap-2">
                        <CheckCircle2 className="w-5 h-5 text-green-600" />
                        <h4 className="text-sm font-semibold">Consolidation Complete</h4>
                      </div>
                      <button 
                        onClick={downloadMerged}
                        className="flex items-center gap-2 px-6 py-2 bg-green-600 text-white rounded-xl text-sm font-medium hover:bg-green-700 transition-colors active:scale-95"
                      >
                        <Download className="w-4 h-4" />
                        Download Merged File
                      </button>
                    </div>

                    <div className="space-y-3">
                      {mergedResults.map((result, idx) => (
                        <div key={idx} className="flex items-center justify-between p-4 bg-[#f9f9f9] rounded-xl border border-black/5">
                          <div className="flex items-center gap-3">
                            <div className="p-2 bg-white rounded-lg border border-black/5">
                              <FileText className="w-4 h-4 text-black/60" />
                            </div>
                            <div>
                              <p className="text-sm font-medium">{result.sheetName}</p>
                              <p className="text-[10px] text-black/40 font-mono uppercase">{result.data.length - 1} ROWS CONSOLIDATED</p>
                            </div>
                          </div>
                          <ChevronRight className="w-4 h-4 text-black/20" />
                        </div>
                      ))}
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>
            </section>
          )}

          {/* Empty State / Instructions */}
          {files.length < 1 && !isProcessing && (
            <section className="text-center py-12 px-8 border-2 border-dashed border-black/5 rounded-3xl">
              <div className="max-w-xs mx-auto">
                <p className="text-sm text-muted-foreground leading-relaxed">
                  Upload Excel files to begin the consolidation process. We'll automatically identify and merge data from all tabs found in your files.
                </p>
              </div>
            </section>
          )}
        </div>

        {/* Footer */}
        <footer className="mt-12 text-center">
          <p className="text-[10px] uppercase tracking-widest text-black/20 font-medium">
            Secure browser-side processing • No data leaves your machine
          </p>
        </footer>
      </div>
    </div>
  );
}
