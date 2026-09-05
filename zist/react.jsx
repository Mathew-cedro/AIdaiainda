import React, { useState, useRef } from "react";
import {
  Upload,
  FileSpreadsheet,
  Download,
  CheckCircle2,
  AlertCircle,
  Users,
  FileCheck,
  RefreshCw,
  Search,
  Settings,
  Layers,
  Archive
} from "lucide-react";

export default function SF9GeneratorApp() {
  const [file, setFile] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [loadingType, setLoadingType] = useState(null); // 'front' | 'back' | 'zip' | 'preview' | null
  const [previewData, setPreviewData] = useState(null);
  const [searchTerm, setSearchTerm] = useState("");
  const [error, setError] = useState(null);
  const [successMessage, setSuccessMessage] = useState(null);
  const [showSettings, setShowSettings] = useState(false);

  // Global overrides
  const [overrides, setOverrides] = useState({
    school_year: "",
    school_head: "",
    adviser_name: "",
    admitted_in: "",
    transfer_date: "",
    admitted_to_grade: "",
    eligible_for: "",
  });

  const fileInputRef = useRef(null);

  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const processFile = async (selectedFile) => {
    if (!selectedFile) return;
    if (!selectedFile.name.match(/\.(xlsx|xlsm|xltx)$/i)) {
      setError("Please upload a valid Excel file (.xlsx)");
      return;
    }

    setFile(selectedFile);
    setError(null);
    setSuccessMessage(null);
    setLoadingType("preview");

    try {
      const formData = new FormData();
      formData.append("file", selectedFile);

      const response = await fetch("/api/preview", {
        method: "POST",
        body: formData,
      });

      const data = await response.json();
      if (!response.ok || !data.success) {
        throw new Error(data.message || data.detail || "Failed to parse Excel file.");
      }

      setPreviewData(data);
      if (data.summary?.school_year && !overrides.school_year) {
        setOverrides((prev) => ({ ...prev, school_year: data.summary.school_year }));
      }
    } catch (err) {
      setError(err.message || "An error occurred while reading the file.");
      setPreviewData(null);
    } finally {
      setLoadingType(null);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      processFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileSelect = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      processFile(e.target.files[0]);
    }
  };

  const handleExport = async (type) => {
    if (!file) {
      setError("Please upload an Excel file first.");
      return;
    }

    setLoadingType(type);
    setError(null);

    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("overrides", JSON.stringify(overrides));

      const endpoint =
        type === "front"
          ? "/api/convert/front"
          : type === "back"
          ? "/api/convert/back"
          : "/api/convert/zip";

      const response = await fetch(endpoint, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const errJson = await response.json().catch(() => ({}));
        throw new Error(errJson.detail || `Failed to generate ${type.toUpperCase()} file.`);
      }

      let filename =
        type === "front"
          ? previewData?.front_filename || "SF9_FRONT.xlsx"
          : type === "back"
          ? previewData?.back_filename || "SF9_BACK.xlsx"
          : previewData?.zip_filename || "SF9_Complete.zip";

      const disposition = response.headers.get("Content-Disposition");
      if (disposition && disposition.indexOf("filename=") !== -1) {
        const match = disposition.match(/filename="?([^"]+)"?/);
        if (match && match[1]) {
          filename = match[1];
        }
      }

      const blob = await response.blob();
      const downloadUrl = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = downloadUrl;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(downloadUrl);

      setSuccessMessage(
        `Downloaded "${filename}" with ${previewData?.count || 0} student sheets.`
      );
    } catch (err) {
      setError(err.message || "Export failed.");
    } finally {
      setLoadingType(null);
    }
  };

  const handleDownloadTemplate = async () => {
    try {
      const response = await fetch("/api/template");
      if (!response.ok) throw new Error("Template not found");
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "SF9_Data_Input_Template.xlsx";
      a.click();
      a.remove();
    } catch {
      setError("Could not download template file.");
    }
  };

  const filteredStudents =
    previewData?.students?.filter((s) => {
      if (!searchTerm) return true;
      const term = searchTerm.toLowerCase();
      return (
        (s.name && String(s.name).toLowerCase().includes(term)) ||
        (s.lrn && String(s.lrn).toLowerCase().includes(term)) ||
        (s.section && String(s.section).toLowerCase().includes(term))
      );
    }) || [];

  return (
    <div className="min-h-screen bg-slate-900 text-slate-100 flex flex-col font-sans">
      {/* Header */}
      <header className="border-b border-slate-800 bg-slate-950/80 backdrop-blur sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-indigo-500 to-sky-600 flex items-center justify-center shadow-lg shadow-indigo-500/20">
              <FileSpreadsheet className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="text-lg font-bold text-white tracking-tight flex items-center gap-2">
                SF9 Report Card Generator
                <span className="px-2 py-0.5 text-xs font-semibold bg-indigo-500/10 text-indigo-400 border border-indigo-500/30 rounded-full">
                  2026–2027 Format
                </span>
              </h1>
              <p className="text-xs text-slate-400">Polangui General Comprehensive High School (PGCHS)</p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            <button
              onClick={handleDownloadTemplate}
              className="inline-flex items-center gap-2 text-xs font-medium px-3 py-2 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-700 transition"
              title="Download input Excel template"
            >
              <Download className="w-3.5 h-3.5 text-slate-400" />
              Sample Template
            </button>
            <button
              onClick={() => setShowSettings(!showSettings)}
              className={`inline-flex items-center gap-1.5 text-xs font-medium px-3 py-2 rounded-lg border transition ${
                showSettings
                  ? "bg-indigo-600 text-white border-indigo-500"
                  : "bg-slate-800 text-slate-300 border-slate-700 hover:bg-slate-700"
              }`}
            >
              <Settings className="w-3.5 h-3.5" />
              Overrides
            </button>
          </div>
        </div>
      </header>

      {/* Main Area */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 flex-1 w-full flex flex-col gap-6">
        {/* Banner Alerts */}
        {error && (
          <div className="p-4 rounded-xl bg-rose-500/10 border border-rose-500/30 flex items-start gap-3 text-rose-300 text-sm animate-fadeIn">
            <AlertCircle className="w-5 h-5 text-rose-400 shrink-0 mt-0.5" />
            <div className="flex-1">{error}</div>
            <button onClick={() => setError(null)} className="text-rose-400 hover:text-rose-200">
              ✕
            </button>
          </div>
        )}

        {successMessage && (
          <div className="p-4 rounded-xl bg-emerald-500/10 border border-emerald-500/30 flex items-start gap-3 text-emerald-300 text-sm animate-fadeIn">
            <CheckCircle2 className="w-5 h-5 text-emerald-400 shrink-0 mt-0.5" />
            <div className="flex-1">{successMessage}</div>
            <button onClick={() => setSuccessMessage(null)} className="text-emerald-400 hover:text-emerald-200">
              ✕
            </button>
          </div>
        )}

        {/* Global Settings Drawer */}
        {showSettings && (
          <div className="p-6 rounded-2xl bg-slate-800/80 border border-slate-700 shadow-xl space-y-4">
            <div className="flex items-center justify-between pb-3 border-b border-slate-700">
              <div>
                <h2 className="text-sm font-semibold text-white">Global Information Overrides (Optional)</h2>
                <p className="text-xs text-slate-400">Values entered here will apply to every student report card</p>
              </div>
              <button onClick={() => setShowSettings(false)} className="text-xs text-slate-400 hover:text-white">
                Close
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4 text-xs">
              <div>
                <label className="block text-slate-400 mb-1">School Year (Cell D12)</label>
                <input
                  type="text"
                  placeholder="e.g. 2026-2027"
                  value={overrides.school_year}
                  onChange={(e) => setOverrides({ ...overrides, school_year: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-slate-200 focus:outline-none focus:border-indigo-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1">Teacher / Adviser Name (Cell J34)</label>
                <input
                  type="text"
                  placeholder="e.g. Maria Santos"
                  value={overrides.adviser_name}
                  onChange={(e) => setOverrides({ ...overrides, adviser_name: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-slate-200 focus:outline-none focus:border-indigo-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1">Approved School Head (Cell C34)</label>
                <input
                  type="text"
                  placeholder="e.g. Juan De La Cruz, Principal IV"
                  value={overrides.school_head}
                  onChange={(e) => setOverrides({ ...overrides, school_head: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-slate-200 focus:outline-none focus:border-indigo-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1">Admitted to Grade (Cell D32)</label>
                <input
                  type="text"
                  placeholder="e.g. Grade 8"
                  value={overrides.admitted_to_grade}
                  onChange={(e) => setOverrides({ ...overrides, admitted_to_grade: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-slate-200 focus:outline-none focus:border-indigo-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1">Eligible for Admission to (Cell M32)</label>
                <input
                  type="text"
                  placeholder="e.g. Grade 8"
                  value={overrides.eligible_for}
                  onChange={(e) => setOverrides({ ...overrides, eligible_for: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-slate-200 focus:outline-none focus:border-indigo-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1">Admitted In (C41) & Date (K41)</label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    placeholder="School Name"
                    value={overrides.admitted_in}
                    onChange={(e) => setOverrides({ ...overrides, admitted_in: e.target.value })}
                    className="w-1/2 bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-slate-200 focus:outline-none focus:border-indigo-500"
                  />
                  <input
                    type="text"
                    placeholder="YYYY-MM-DD"
                    value={overrides.transfer_date}
                    onChange={(e) => setOverrides({ ...overrides, transfer_date: e.target.value })}
                    className="w-1/2 bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-slate-200 focus:outline-none focus:border-indigo-500"
                  />
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Drag & Drop File Zone */}
        <div
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
          className={`relative border-2 border-dashed rounded-3xl p-8 sm:p-12 text-center cursor-pointer transition-all duration-200 ${
            isDragging
              ? "border-indigo-500 bg-indigo-500/10 scale-[1.01]"
              : file
              ? "border-emerald-500/50 bg-slate-800/40 hover:bg-slate-800/60"
              : "border-slate-700 hover:border-slate-500 bg-slate-800/20 hover:bg-slate-800/40"
          }`}
        >
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileSelect}
            accept=".xlsx,.xlsm,.xltx"
            className="hidden"
          />

          <div className="flex flex-col items-center justify-center space-y-4 max-w-md mx-auto">
            <div
              className={`w-16 h-16 rounded-2xl flex items-center justify-center transition ${
                file ? "bg-emerald-500/20 text-emerald-400" : "bg-indigo-500/10 text-indigo-400"
              }`}
            >
              {file ? <FileCheck className="w-8 h-8" /> : <Upload className="w-8 h-8" />}
            </div>

            <div>
              <h3 className="text-base sm:text-lg font-semibold text-white">
                {file ? file.name : "Drag & drop your student Excel file here"}
              </h3>
              <p className="text-xs sm:text-sm text-slate-400 mt-1">
                {file
                  ? `${(file.size / 1024).toFixed(1)} KB • Click or drop another file to replace`
                  : "Upload .xlsx spreadsheet containing student list, grades, and attendance"}
              </p>
            </div>

            {!file && (
              <span className="inline-flex items-center gap-2 text-xs font-medium px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white shadow-lg shadow-indigo-600/30 transition">
                Browse Excel File
              </span>
            )}
          </div>
        </div>

        {/* Action / Export Toolbar */}
        {previewData && (
          <div className="bg-slate-800/90 border border-slate-700/80 rounded-2xl p-5 flex flex-col lg:flex-row items-center justify-between gap-5 shadow-xl animate-fadeIn">
            <div className="flex items-center gap-4 w-full lg:w-auto">
              <div className="w-12 h-12 rounded-xl bg-indigo-500/20 flex items-center justify-center text-indigo-400 shrink-0">
                <Users className="w-6 h-6" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <h3 className="text-sm font-bold text-white">
                    {previewData.count} Students Loaded
                  </h3>
                  <span className="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                    Ready to Export
                  </span>
                </div>
                <p className="text-xs text-slate-400 font-mono mt-0.5">
                  1 sheet per student will be generated
                </p>
              </div>
            </div>

            {/* Separate Download Buttons */}
            <div className="flex flex-wrap items-center gap-3 w-full lg:w-auto justify-end">
              {/* Download Front */}
              <button
                onClick={() => handleExport("front")}
                disabled={loadingType !== null}
                className="flex-1 sm:flex-none inline-flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white font-semibold text-xs shadow-lg shadow-blue-500/20 transition disabled:opacity-50"
              >
                {loadingType === "front" ? (
                  <RefreshCw className="w-4 h-4 animate-spin" />
                ) : (
                  <Layers className="w-4 h-4" />
                )}
                Download FRONT (.xlsx)
              </button>

              {/* Download Back */}
              <button
                onClick={() => handleExport("back")}
                disabled={loadingType !== null}
                className="flex-1 sm:flex-none inline-flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-semibold text-xs shadow-lg shadow-emerald-500/20 transition disabled:opacity-50"
              >
                {loadingType === "back" ? (
                  <RefreshCw className="w-4 h-4 animate-spin" />
                ) : (
                  <Layers className="w-4 h-4" />
                )}
                Download BACK (.xlsx)
              </button>

              {/* Download Both ZIP */}
              <button
                onClick={() => handleExport("zip")}
                disabled={loadingType !== null}
                className="flex-1 sm:flex-none inline-flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl bg-slate-700 hover:bg-slate-600 text-white font-semibold text-xs border border-slate-600 transition disabled:opacity-50"
              >
                {loadingType === "zip" ? (
                  <RefreshCw className="w-4 h-4 animate-spin" />
                ) : (
                  <Archive className="w-4 h-4 text-sky-400" />
                )}
                Both (ZIP Bundle)
              </button>
            </div>
          </div>
        )}

        {/* Students Table */}
        {previewData && (
          <div className="bg-slate-800/50 border border-slate-700/60 rounded-2xl overflow-hidden flex flex-col flex-1 shadow-lg">
            <div className="p-4 border-b border-slate-700/80 flex flex-col sm:flex-row items-center justify-between gap-3 bg-slate-800/70">
              <div className="relative w-full sm:w-72">
                <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="text"
                  placeholder="Search student name or LRN..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="w-full pl-9 pr-3 py-1.5 bg-slate-900/90 border border-slate-700 rounded-lg text-xs text-slate-200 focus:outline-none focus:border-indigo-500"
                />
              </div>

              <div className="text-xs text-slate-400">
                Showing <span className="font-semibold text-slate-200">{filteredStudents.length}</span> of {previewData.count} students
              </div>
            </div>

            <div className="overflow-x-auto max-h-96">
              <table className="w-full text-left text-xs border-collapse">
                <thead className="bg-slate-900/90 sticky top-0 z-10 text-slate-400 font-medium uppercase text-[10px] tracking-wider border-b border-slate-700">
                  <tr>
                    <th className="px-4 py-3">#</th>
                    <th className="px-4 py-3">Student Name</th>
                    <th className="px-4 py-3">LRN</th>
                    <th className="px-4 py-3">Sex</th>
                    <th className="px-4 py-3">Grade & Sec</th>
                    <th className="px-4 py-3 text-center">Filipino</th>
                    <th className="px-4 py-3 text-center">English</th>
                    <th className="px-4 py-3 text-center">Math</th>
                    <th className="px-4 py-3 text-center">Science</th>
                    <th className="px-4 py-3 text-center">TLE (Creative Tech)</th>
                    <th className="px-4 py-3 text-center">MAPEH</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-700/50">
                  {filteredStudents.length > 0 ? (
                    filteredStudents.map((s, idx) => (
                      <tr key={idx} className="hover:bg-slate-700/30 transition">
                        <td className="px-4 py-2.5 text-slate-500 font-mono">{idx + 1}</td>
                        <td className="px-4 py-2.5 font-semibold text-slate-100">{s.name || "—"}</td>
                        <td className="px-4 py-2.5 font-mono text-slate-300">{s.lrn || "—"}</td>
                        <td className="px-4 py-2.5 text-slate-300">{s.sex || "—"}</td>
                        <td className="px-4 py-2.5 text-slate-300">
                          {s.grade ? `Gr ${s.grade}` : ""} {s.section ? `• ${s.section}` : ""}
                        </td>
                        <td className="px-4 py-2.5 text-center text-slate-300 font-mono">
                          {s.filipino_t1 ?? "—"}
                        </td>
                        <td className="px-4 py-2.5 text-center text-slate-300 font-mono">
                          {s.english_t1 ?? "—"}
                        </td>
                        <td className="px-4 py-2.5 text-center text-slate-300 font-mono">
                          {s.math_t1 ?? "—"}
                        </td>
                        <td className="px-4 py-2.5 text-center text-slate-300 font-mono">
                          {s.science_t1 ?? "—"}
                        </td>
                        <td className="px-4 py-2.5 text-center text-slate-300 font-mono">
                          {s.tle_t1 ?? "—"}
                        </td>
                        <td className="px-4 py-2.5 text-center text-slate-300 font-mono">
                          {s.mapeh_t1 ?? "—"}
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={11} className="px-4 py-8 text-center text-slate-400">
                        No students match your search filter.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-800 py-4 text-center text-xs text-slate-500">
        Polangui General Comprehensive High School • SF9 Report Card Generator (2026–2027)
      </footer>
    </div>
  );
}
