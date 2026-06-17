/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import {
  FileSpreadsheet,
  Upload,
  Download,
  Settings,
  Table,
  Plus,
  Trash2,
  Filter,
  Info,
  Database,
  FileDown,
  CheckCircle,
  AlertCircle,
  HelpCircle
} from 'lucide-react';

import { RawRecord, ColumnMapping } from './types';
import { parseUploadedExcel, generateFilteredReport } from './utils/excelProcessor';

// Brand navy used across header / buttons / table headers
const NAVY = "#1B365D";

export default function App() {
  // --- STATE ---
  const [records, setRecords] = useState<RawRecord[]>([]);
  const [selectedProvinces, setSelectedProvinces] = useState<string[]>(["พิษณุโลก", "สุโขทัย"]);
  const [activeTab, setActiveTab] = useState<'preview' | 'raw_editor'>('preview');

  // Mapping columns as specified in the Python script
  const [columnMapping, setColumnMapping] = useState<ColumnMapping>({
    province: "F",
    store: "C",
    bill: "I",
    pCode: "J",
    pName: "K",
    qty: "L",
    truck: "Q"
  });

  const [feedback, setFeedback] = useState<{ type: 'success' | 'error' | 'info'; message: string } | null>({
    type: 'info',
    message: 'ยินดีต้อนรับครับพี่! กรุณาอัปโหลดไฟล์ Excel ปลายทาง (เช่นไฟล์ 15-6.xlsx) ทางซ้ายมือเพื่อเริ่มดึงข้อมูลและสรุปยอดรวมย่อยอัตโนมัติทันทีครับ'
  });

  // For adding new row in the editor
  const [newRow, setNewRow] = useState<Omit<RawRecord, 'id'>>({
    province: 'พิษณุโลก',
    store: '',
    bill: '',
    pCode: '',
    pName: '',
    qty: 0,
    truck: '',
  });

  // Loading state
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Extracted list of all unique provinces present in the current records
  const allProvincesInDataset = useMemo(() => {
    const set = new Set<string>();
    records.forEach(r => {
      if (r.province.trim()) set.add(r.province.trim());
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'th'));
  }, [records]);

  // --- ACTIONS ---

  // Handle file uploads
  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    const file = files[0];
    setIsProcessing(true);
    setFeedback({ type: 'info', message: 'กำลังอ่านโครงสร้างไฟล์และดึงข้อมูลโปรดรอสักครู่...' });

    try {
      const provincesSet = new Set<string>();
      const result = await parseUploadedExcel(file, columnMapping, provincesSet);

      if (result.records.length === 0) {
        setFeedback({
          type: 'error',
          message: '❌ ไม่พบข้อมูลใด ๆ ในไฟล์ตามการตั้งค่าคอลัมน์นี้ครับ ลองตรวจสอบ Column Mapping ด้านซ้ายมือ'
        });
      } else {
        setRecords(result.records);
        // Identify provinces and auto-select Phitsanulok & Sukhothai if present, otherwise select the first few
        const foundProvs = result.allProvinces;
        const defaultSelected = foundProvs.filter(p => ["พิษณุโลก", "สุโขทัย"].includes(p));

        setSelectedProvinces(defaultSelected.length > 0 ? defaultSelected : foundProvs.slice(0, 3));
        setFeedback({
          type: 'success',
          message: `✔️ ดึงข้อมูลเสร็จสิ้น! โหลดข้อมูลเสร็จสำเร็จ ${result.records.length} แถว (พบจังวัดทั้งหมดในระบบ: ${foundProvs.join(', ')})`
        });
      }
    } catch (err: any) {
      setFeedback({
        type: 'error',
        message: err.message || 'เกิดข้อผิดพลาดในการแปลรายงานไฟล์ Excel'
      });
    } finally {
      setIsProcessing(false);
      // Reset input element
      if (event.target) event.target.value = '';
    }
  };

  // Handle Drag Over
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  };

  // Handle Drop
  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();

    const files = e.dataTransfer.files;
    if (!files || files.length === 0) return;

    const file = files[0];
    // Re-use logic from handleFileUpload but for a direct File object
    setIsProcessing(true);
    setFeedback({ type: 'info', message: 'กำลังอ่านโครงสร้างไฟล์จากการลากวางโปรดรอสักครู่...' });

    try {
      const provincesSet = new Set<string>();
      const result = await parseUploadedExcel(file, columnMapping, provincesSet);

      if (result.records.length === 0) {
        setFeedback({
          type: 'error',
          message: '❌ ไม่พบข้อมูลใด ๆ ในไฟล์จากการลากวางครับ ลองตรวจสอบ Column Mapping'
        });
      } else {
        setRecords(result.records);
        const foundProvs = result.allProvinces;
        const defaultSelected = foundProvs.filter(p => ["พิษณุโลก", "สุโขทัย"].includes(p));
        setSelectedProvinces(defaultSelected.length > 0 ? defaultSelected : foundProvs.slice(0, 3));
        setFeedback({
          type: 'success',
          message: `✔️ ลากวางไฟล์เสร็จสิ้น! โหลดข้อมูลสำเร็จ ${result.records.length} แถว`
        });
      }
    } catch (err: any) {
      setFeedback({
        type: 'error',
        message: err.message || 'เกิดข้อผิดพลาดในการลากวางไฟล์ Excel'
      });
    } finally {
      setIsProcessing(false);
    }
  };

  // Trigger Excel file generation mimicking python script exactly
  const handleExportExcel = async () => {
    setIsProcessing(true);
    try {
      const activeProvs = selectedProvinces.filter(p => records.some(r => r.province === p));
      if (activeProvs.length === 0) {
        setFeedback({
          type: 'error',
          message: '❌ ไม่พบกลุ่มจังหวัดที่เลือกในชุดข้อมูลปัจจุบัน กรุณาติ๊กเลือกอย่างน้อยหนึ่งจังหวัด'
        });
        setIsProcessing(false);
        return;
      }

      await generateFilteredReport(
        records,
        activeProvs,
        "รายงานสรุปการจัดส่งสินค้า แยกรายร้านค้าพร้อมยอดรวมย่อย.xlsx"
      );

      setFeedback({
        type: 'success',
        message: '🎉 ส่งออกรายงาน Excel เรียบร้อยแล้ว! ไฟล์เดียวมี 2 ชีต: "รายงานสรุปแยกจังหวัด" และ "สรุปยอดตามทะเบียนรถ" 🚚'
      });
    } catch (err: any) {
      setFeedback({
        type: 'error',
        message: err.message || 'เกิดข้อผิดพลาดในการเขียนชุดรายงาน Excel'
      });
    } finally {
      setIsProcessing(false);
    }
  };

  // Convert current sorted table data to CSV download
  const handleExportCSV = () => {
    try {
      const activeProvs = selectedProvinces.filter(p => records.some(r => r.province === p));
      const filtered = records.filter(r => activeProvs.includes(r.province));

      if (filtered.length === 0) {
        setFeedback({ type: 'error', message: '❌ ไม่มีข้อมูลที่จะส่งออกเป็น CSV ในจังหวัดที่เลือก' });
        return;
      }

      // Sort: Prov -> Store -> Bill
      filtered.sort((a, b) => {
        if (a.province !== b.province) return a.province.localeCompare(b.province, 'th');
        if (a.store !== b.store) return a.store.localeCompare(b.store, 'th');
        return a.bill.localeCompare(b.bill, 'th');
      });

      const headers = ["จังหวัด", "ร้านค้า", "เลขที่บิล", "รหัสสินค้า", "สินค้า", "จำนวน(หีบ)"];
      const csvContent = [
        "﻿" + headers.join(","), // UTF-8 BOM
        ...filtered.map(r => `"${r.province}","${r.store}","${r.bill}","${r.pCode}","${r.pName}",${r.qty}`)
      ].join("\n");

      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.setAttribute("href", url);
      link.setAttribute("download", `ข้อมูลดิบที่กรอง_${activeProvs.join("_")}.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      setFeedback({ type: 'success', message: '✔️ ดาวน์โหลดไฟล์ข้อมูลดิบ CSV เรียบร้อยแล้วพี่!' });
    } catch (err: any) {
      setFeedback({ type: 'error', message: 'เกิดข้อผิดพลาดในการสร้างไฟล์ CSV' });
    }
  };

  // Clear state completely
  const handleClearAll = () => {
    setRecords([]);
    setSelectedProvinces([]);
    setFeedback({
      type: 'info',
      message: '🧹 ล้างหน้ากระดานเรียบร้อยแล้ว เพิ่มข้อมูลโดยอัปโหลดไฟล์ Excel หรือพิมพ์ในส่วนแท็บแก้ไขข้อมูลดิบได้เลยครับ'
    });
  };

  // Toggle dynamic province checks
  const handleProvinceToggle = (provName: string) => {
    setSelectedProvinces(prev =>
      prev.includes(provName)
        ? prev.filter(p => p !== provName)
        : [...prev, provName]
    );
  };

  // Edit record value dynamically
  const handleEditRecordField = (id: string, field: keyof RawRecord, value: any) => {
    setRecords(prev => prev.map(rec => {
      if (rec.id === id) {
        let finalVal = value;
        if (field === 'qty') {
          const parsed = Number(value);
          finalVal = isNaN(parsed) ? 0 : parsed;
        }
        return { ...rec, [field]: finalVal };
      }
      return rec;
    }));
  };

  // Remove individual row
  const handleDeleteRecord = (id: string) => {
    setRecords(prev => prev.filter(rec => rec.id !== id));
    setFeedback({
      type: 'success',
      message: 'ลบแถวข้อมูลสำเร็จ'
    });
  };

  // Add manually structured row
  const handleAddNewRow = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newRow.store.trim() || !newRow.pName.trim()) {
      setFeedback({
        type: 'error',
        message: '❌ กรุณากรอกชื่อร้านค้า และ ชื่อสินค้าให้สมบูรณ์ก่อนบันทึกแถวใหม่ครับพี่'
      });
      return;
    }

    const item: RawRecord = {
      ...newRow,
      id: `manual-${Date.now()}-${Math.floor(Math.random() * 1000)}`
    };

    setRecords(prev => [...prev, item]);
    setNewRow({
      province: newRow.province || 'พิษณุโลก',
      store: '',
      bill: '',
      pCode: '',
      pName: '',
      qty: 0,
      truck: '',
    });
    setFeedback({
      type: 'success',
      message: `✔️ เพิ่มแถวยอดส่งของร้าน "${item.store}" เข้าสู่ระบบและคำนวณสรุปผลสดทันที!`
    });
  };

  // --- MEMOIZED CALCULATIONS FOR RENDER PREVIEW GRID---

  // Build sequential exact table structure similar to Python sheet formatting
  const previewRows = useMemo(() => {
    // Filter active records by provinces that exist inside current records and selected list
    const filteredDataset = records.filter(r => selectedProvinces.includes(r.province));

    // Sort: Province -> Store -> Bill
    filteredDataset.sort((a, b) => {
      if (a.province !== b.province) return a.province.localeCompare(b.province, 'th');
      if (a.store !== b.store) return a.store.localeCompare(b.store, 'th');
      return a.bill.localeCompare(b.bill, 'th');
    });

    const rows: {
      key: string;
      type: 'data' | 'store_subtotal' | 'prov_total' | 'grand_total' | 'empty';
      label?: string;
      isZebra?: boolean;
      record?: RawRecord;
      qtyValue?: number | string;
    }[] = [];

    // Find unique provinces listed in this filtered result set
    const activeProvincesInFiltered = Array.from(new Set(filteredDataset.map(r => r.province))) as string[];
    activeProvincesInFiltered.sort((a, b) => a.localeCompare(b, 'th'));

    let overallGrandTotal = 0;

    activeProvincesInFiltered.forEach((targetProv) => {
      const provRecords = filteredDataset.filter(r => r.province === targetProv);
      if (provRecords.length === 0) return;

      let idx = 0;
      let provSumAccumulated = 0;

      // Extract unique stores sequentially
      while (idx < provRecords.length) {
        const currentStore = provRecords[idx].store;
        const storeRecords: RawRecord[] = [];

        while (idx < provRecords.length && provRecords[idx].store === currentStore) {
          storeRecords.push(provRecords[idx]);
          idx++;
        }

        // Add each data record row
        let isZebra = false;
        let storeSumAccumulated = 0;
        storeRecords.forEach((item, innerIdx) => {
          rows.push({
            key: `data-${item.id}-${innerIdx}`,
            type: 'data',
            record: item,
            isZebra
          });
          storeSumAccumulated += item.qty;
          isZebra = !isZebra;
        });

        // Add Store subtotal row
        rows.push({
          key: `subtotal-store-${targetProv}-${currentStore}`,
          type: 'store_subtotal',
          label: `รวมย่อยร้าน - ${currentStore}`,
          qtyValue: storeSumAccumulated
        });

        provSumAccumulated += storeSumAccumulated;
      }

      // Add Provincial Total Row
      rows.push({
        key: `total-prov-${targetProv}`,
        type: 'prov_total',
        label: `รวมทั้งสิ้น จังหวัด${targetProv}`,
        qtyValue: provSumAccumulated
      });

      overallGrandTotal += provSumAccumulated;

      // Add space row like Python `currentRowNum += 2`
      rows.push({
        key: `space-after-${targetProv}`,
        type: 'empty'
      });
    });

    // Remove the very last empty spacer row if present, then add Grand Total Row
    if (rows.length > 0 && rows[rows.length - 1].type === 'empty') {
      rows.pop();
    }

    if (rows.length > 0) {
      rows.push({
        key: 'grand-total-row',
        type: 'grand_total',
        label: `รวมทั้งสิ้นสุทธิ (${selectedProvinces.join(" & ")})`,
        qtyValue: overallGrandTotal
      });
    }

    return {
      tableRows: rows,
      grandTotal: overallGrandTotal,
      recordCount: filteredDataset.length,
      storeCount: new Set(filteredDataset.map(r => r.store)).size,
      provinceCount: activeProvincesInFiltered.length
    };
  }, [records, selectedProvinces]);

  // Shared input styling for column-mapping boxes
  const mapInput = "w-full bg-slate-50 border border-slate-300 rounded-lg px-2.5 py-1.5 text-sm font-semibold font-mono text-center text-[#1B365D] focus:outline-none focus:ring-2 focus:ring-[#1B365D]/15 focus:border-[#1B365D] transition";

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 flex flex-col antialiased">

      {/* HEADER BANNER */}
      <header className="bg-white border-b border-slate-200 py-4 px-6 shadow-sm sticky top-0 z-40">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row items-center justify-between gap-4">

          <div className="flex items-center gap-4">
            <div className="text-white p-3 rounded-xl shadow-sm flex items-center justify-center" style={{ backgroundColor: NAVY }}>
              <FileSpreadsheet className="w-7 h-7" id="header-icon" />
            </div>
            <div>
              <h1 className="text-xl md:text-2xl font-bold tracking-tight text-slate-800">
                ระบบจัดการรายงานและสรุปยอดส่งสินค้าแยกรายร้านค้า
              </h1>
              <p className="text-sm text-slate-500">
                ประมวลผล ดึงคอลัมน์ L อัตโนมัติ พร้อมคำนวณแถวสรุปยอดรวมย่อย (Subtotal) สำหรับ พิษณุโลก &amp; สุโขทัย
              </p>
            </div>
          </div>

          {/* QUICK SYSTEM ACTIONS */}
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={handleClearAll}
              type="button"
              className="px-3 py-2 bg-slate-50 hover:bg-rose-50 hover:text-rose-600 hover:border-rose-200 text-slate-500 text-xs rounded-lg border border-slate-200 transition flex items-center gap-1.5 cursor-pointer font-medium"
              id="clear-btn"
            >
              <Trash2 className="w-3.5 h-3.5" />
              ล้างหน้าตาราง
            </button>
          </div>
        </div>
      </header>

      {/* SYSTEM FEEDBACK NOTIFICATION BAR */}
      {feedback && (
        <div className={`border-b py-3 px-4 ${
          feedback.type === 'error' ? 'bg-rose-50 border-rose-100'
            : feedback.type === 'success' ? 'bg-emerald-50 border-emerald-100'
            : 'bg-sky-50 border-sky-100'
        }`}>
          <div className="max-w-7xl mx-auto flex items-start gap-3">
            {feedback.type === 'success' && <CheckCircle className="w-5 h-5 text-emerald-500 shrink-0 mt-0.5" id="feedback-success" />}
            {feedback.type === 'error' && <AlertCircle className="w-5 h-5 text-rose-500 shrink-0 mt-0.5" id="feedback-error" />}
            {feedback.type === 'info' && <Info className="w-5 h-5 text-sky-500 shrink-0 mt-0.5" id="feedback-info" />}
            <p className={`text-xs md:text-sm ${feedback.type === 'error' ? 'text-rose-700' : feedback.type === 'success' ? 'text-emerald-700' : 'text-sky-800'}`}>
              <strong>แจ้งระบบ:</strong> {feedback.message}
            </p>
          </div>
        </div>
      )}

      {/* DASHBOARD ANALYTICS OVERVIEW */}
      <section className="py-6 px-6">
        <div className="max-w-7xl mx-auto grid grid-cols-2 md:grid-cols-4 gap-4">

          <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm">
            <span className="text-xs font-semibold text-slate-400 uppercase tracking-wider block">จำนวนรายการที่กรอง</span>
            <div className="flex items-baseline gap-2 mt-2">
              <span className="text-2xl md:text-3xl font-extrabold" style={{ color: NAVY }}>{previewRows.recordCount}</span>
              <span className="text-xs text-slate-400">รายการ</span>
            </div>
            <div className="text-[11px] text-slate-400 mt-1">จากทั้งหมด {records.length} รายการ</div>
          </div>

          <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm">
            <span className="text-xs font-semibold text-slate-400 uppercase tracking-wider block">ร้านค้าทั้งหมด</span>
            <div className="flex items-baseline gap-2 mt-2">
              <span className="text-2xl md:text-3xl font-extrabold text-violet-600">{previewRows.storeCount}</span>
              <span className="text-xs text-slate-400">ร้านค้า</span>
            </div>
            <div className="text-[11px] text-slate-400 mt-1">แยกเป็นกลุ่มจังหวัดปลายทาง</div>
          </div>

          <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm">
            <span className="text-xs font-semibold text-slate-400 uppercase tracking-wider block">ปริมาณรวบสุทธิ (หีบ)</span>
            <div className="flex items-baseline gap-2 mt-2">
              <span className="text-2xl md:text-3xl font-extrabold text-emerald-600">{previewRows.grandTotal.toLocaleString()}</span>
              <span className="text-xs text-slate-400">หีบ</span>
            </div>
            <div className="text-[11px] text-slate-400 mt-1">อ้างอิงจากคอลัมน์ L ใน Excel</div>
          </div>

          <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm">
            <span className="text-xs font-semibold text-slate-400 uppercase tracking-wider block">กลุ่มจังหวัดที่เลือกกรอง</span>
            <div className="flex items-center gap-1.5 mt-2 flex-wrap">
              {selectedProvinces.length === 0 ? (
                <span className="text-xs text-slate-400">ไม่ได้เลือก</span>
              ) : (
                selectedProvinces.map(p => (
                  <span key={p} className="bg-blue-50 text-xs px-2 py-0.5 rounded-md text-[#1B365D] border border-blue-100 font-semibold">
                    {p}
                  </span>
                ))
              )}
            </div>
            <div className="text-[11px] text-slate-400 mt-2">รวมทั้งหมด {previewRows.provinceCount} จังหวัดในตารางจำลอง</div>
          </div>

        </div>
      </section>

      {/* MAIN LAYOUT GRID */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 md:px-6 pb-6 grid grid-cols-1 lg:grid-cols-12 gap-6">

        {/* LEFT COLUMN: CONTROL & SETTINGS (4 spans) */}
        <section className="lg:col-span-4 flex flex-col gap-6">

          {/* UPLOAD & PARSER SETTINGS */}
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">

            <h2 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2 pb-3 border-b border-slate-100">
              <span className="w-6 h-6 rounded-md text-white text-xs flex items-center justify-center font-bold" style={{ backgroundColor: NAVY }}>1</span>
              <Upload className="w-4 h-4 text-[#1B365D]" />
              นำเข้าไฟล์ Excel จริง
            </h2>

            <div
              onClick={() => fileInputRef.current?.click()}
              onDragOver={handleDragOver}
              onDrop={handleDrop}
              className="border-2 border-dashed border-slate-300 hover:border-[#1B365D] hover:bg-blue-50/40 bg-slate-50 p-6 rounded-xl flex flex-col items-center justify-center text-center cursor-pointer transition-all group"
            >
              <input
                type="file"
                ref={fileInputRef}
                onChange={handleFileUpload}
                accept=".xlsx,.xls"
                className="hidden"
              />
              <FileSpreadsheet className="w-12 h-12 text-slate-400 group-hover:text-[#1B365D] group-hover:scale-105 transition-all mb-3" />
              <p className="text-sm font-semibold text-slate-700">ลากหรือคลิกเพื่ออัปโหลดไฟล์ Excel</p>
              <p className="text-xs text-slate-400 mt-1">รองรับไฟล์ตระกูล .xlsx (ไฟล์สรุปรายจังหวัด)</p>
            </div>

            <div className="mt-4 bg-blue-50/50 rounded-lg p-3 text-xs text-slate-600 border border-blue-100">
              <span className="font-semibold text-slate-700 flex items-center gap-1 mb-1">
                <Info className="w-3.5 h-3.5 text-[#1B365D]" /> คำชี้แจงโครงสร้างไฟล์:
              </span>
              โปรแกรมจะอ่านข้อมูลแผ่นงานแรก เริ่มถอดแถวข้อมูลตั้งแต่แถวที่ 2 เป็นต้นไป (โดยถือว่าแถวที่ 1 เป็นส่วนหัวรายงานของพี่)
            </div>

          </div>

          {/* COLUMN MAPPING DEFINITIONS */}
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">

            <h2 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2 pb-3 border-b border-slate-100">
              <span className="w-6 h-6 rounded-md text-white text-xs flex items-center justify-center font-bold" style={{ backgroundColor: NAVY }}>2</span>
              <Settings className="w-4 h-4 text-slate-500" />
              ตั้งค่าคอลัมน์ (Column Mapping)
            </h2>

            <p className="text-xs text-slate-500 mb-4">
              กำหนดคุณลักษณะเฉพาะคอลัมน์ของไฟล์ Excel เพื่อดึงค่าข้อมูลแต่ละหมวดให้ถูกต้อง:
            </p>

            <form className="space-y-3.5">

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-slate-500 mb-1">คอลัมน์ - จังหวัด</label>
                  <input
                    type="text"
                    value={columnMapping.province}
                    onChange={(e) => setColumnMapping(prev => ({ ...prev, province: e.target.value.toUpperCase() }))}
                    className={mapInput}
                    placeholder="F"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-500 mb-1">คอลัมน์ - ร้านค้า</label>
                  <input
                    type="text"
                    value={columnMapping.store}
                    onChange={(e) => setColumnMapping(prev => ({ ...prev, store: e.target.value.toUpperCase() }))}
                    className={mapInput}
                    placeholder="C"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-slate-500 mb-1">คอลัมน์ - เลขที่บิล</label>
                  <input
                    type="text"
                    value={columnMapping.bill}
                    onChange={(e) => setColumnMapping(prev => ({ ...prev, bill: e.target.value.toUpperCase() }))}
                    className={mapInput}
                    placeholder="I"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-500 mb-1">คอลัมน์ - รหัสสินค้า</label>
                  <input
                    type="text"
                    value={columnMapping.pCode}
                    onChange={(e) => setColumnMapping(prev => ({ ...prev, pCode: e.target.value.toUpperCase() }))}
                    className={mapInput}
                    placeholder="J"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3 border-t border-slate-100 pt-3.5">
                <div>
                  <label className="block text-xs font-medium text-slate-500 mb-1">คอลัมน์ - สินค้า</label>
                  <input
                    type="text"
                    value={columnMapping.pName}
                    onChange={(e) => setColumnMapping(prev => ({ ...prev, pName: e.target.value.toUpperCase() }))}
                    className={mapInput}
                    placeholder="K"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-emerald-600 mb-1">🎯 ยอดสินค้า (คอลัมน์ L)</label>
                  <input
                    type="text"
                    value={columnMapping.qty}
                    onChange={(e) => setColumnMapping(prev => ({ ...prev, qty: e.target.value.toUpperCase() }))}
                    className="w-full bg-emerald-50 border-2 border-emerald-200 rounded-lg px-2.5 py-1.5 text-sm text-emerald-700 font-bold font-mono text-center focus:outline-none focus:ring-2 focus:ring-emerald-300/40 focus:border-emerald-500 transition"
                    placeholder="L"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3 border-t border-slate-100 pt-3.5">
                <div>
                  <label className="block text-xs font-semibold text-sky-600 mb-1">🚚 ทะเบียนรถ (คอลัมน์ Q)</label>
                  <input
                    type="text"
                    value={columnMapping.truck}
                    onChange={(e) => setColumnMapping(prev => ({ ...prev, truck: e.target.value.toUpperCase() }))}
                    className="w-full bg-sky-50 border-2 border-sky-200 rounded-lg px-2.5 py-1.5 text-sm text-sky-700 font-bold font-mono text-center focus:outline-none focus:ring-2 focus:ring-sky-300/40 focus:border-sky-500 transition"
                    placeholder="Q"
                  />
                </div>
                <div className="flex items-end">
                  <p className="text-[11px] text-slate-400 leading-tight pb-1">
                    ใช้สร้างชีต <span className="text-sky-600 font-semibold">"สรุปยอดตามทะเบียนรถ"</span> เพิ่มในไฟล์เดียวกัน
                  </p>
                </div>
              </div>

              <div className="bg-slate-50 p-2.5 rounded-lg text-[11px] text-slate-400 text-center font-mono border border-slate-100">
                ค่าเริ่มต้น: F=จังหวัด | C=ร้าน | I=บิล | J=รหัส | K=สินค้า | L=จำนวน(หีบ) | Q=ทะเบียนรถ
              </div>

            </form>

          </div>

          {/* DYNAMIC PROVINCE FILTERING */}
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">

            <h2 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2 pb-3 border-b border-slate-100">
              <span className="w-6 h-6 rounded-md text-white text-xs flex items-center justify-center font-bold" style={{ backgroundColor: NAVY }}>3</span>
              <Filter className="w-4 h-4 text-amber-500" />
              เลือกกรองจังหวัดตรวจรับงาน
            </h2>

            <p className="text-xs text-slate-500 mb-3">
              ระบบดึงรายการจังวัดที่มีในไฟล์ปัจจุบันให้อัตโนมัติ ติ๊กเลือกกลุ่มจังหวัดที่ต้องการนำลงตาราง:
            </p>

            {allProvincesInDataset.length === 0 ? (
              <div className="py-5 text-center text-xs text-slate-400 bg-slate-50 rounded-lg border border-dashed border-slate-200">
                ไม่มีข้อมูลจังหวัดในระบบชั่วคราว
              </div>
            ) : (
              <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
                {allProvincesInDataset.map((prov) => {
                  const isChecked = selectedProvinces.includes(prov);
                  const count = records.filter(r => r.province === prov).length;
                  const isHighlight = ["พิษณุโลก", "สุโขทัย"].includes(prov);

                  return (
                    <label
                      key={prov}
                      className={`flex items-center justify-between p-2.5 rounded-lg border text-xs cursor-pointer transition-all ${isChecked
                        ? 'bg-blue-50 border-[#1B365D]/30 text-[#1B365D]'
                        : 'bg-white border-slate-200 text-slate-500 hover:border-slate-300 hover:bg-slate-50'
                        }`}
                    >
                      <div className="flex items-center gap-2.5">
                        <input
                          type="checkbox"
                          checked={isChecked}
                          onChange={() => handleProvinceToggle(prov)}
                          className="rounded border-slate-300 accent-[#1B365D] w-4 h-4 cursor-pointer"
                        />
                        <span className={`font-semibold ${isHighlight ? 'text-[#1B365D]' : ''}`}>
                          {prov}
                        </span>
                        {isHighlight && (
                          <span className="text-[10px] bg-[#1B365D] text-white px-1.5 py-0.5 rounded">
                            เป้าหมาย
                          </span>
                        )}
                      </div>
                      <span className="font-mono text-[10px] text-slate-400">({count} แถว)</span>
                    </label>
                  );
                })}
              </div>
            )}

            <div className="flex items-center justify-between gap-2 mt-4 pt-4 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setSelectedProvinces(["พิษณุโลก", "สุโขทัย"])}
                className="text-[11px] font-medium text-[#1B365D] hover:underline transition cursor-pointer"
              >
                เลือกเฉพาะเป้าหมายหลัก
              </button>
              <button
                type="button"
                onClick={() => setSelectedProvinces(allProvincesInDataset)}
                className="text-[11px] text-slate-500 hover:text-slate-800 transition cursor-pointer"
              >
                เลือกทั้งหมด
              </button>
            </div>

          </div>

        </section>

        {/* RIGHT COLUMN: MAIN PREVIEW AND EDIT SYSTEM (8 spans) */}
        <section className="lg:col-span-8 flex flex-col gap-6">

          {/* TABS CONTAINER */}
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col">

            {/* TABS HEADER AND EXPORTS */}
            <div className="bg-slate-50/70 border-b border-slate-200 px-4 pt-4 flex flex-col sm:flex-row items-center justify-between gap-4">

              {/* TAB SELECTORS */}
              <div className="flex gap-2 w-full sm:w-auto">
                <button
                  type="button"
                  onClick={() => setActiveTab('preview')}
                  className={`px-4 py-2.5 text-xs font-bold transition-all relative flex items-center gap-2 cursor-pointer ${activeTab === 'preview'
                    ? 'text-[#1B365D]'
                    : 'text-slate-400 hover:text-slate-600'
                    }`}
                >
                  <Table className="w-4 h-4" />
                  ดูตารางจำลอง Subtotal (Excel Preview)
                  {activeTab === 'preview' && (
                    <motion.div
                      layoutId="activeTabUnderline"
                      className="absolute bottom-0 left-0 right-0 h-0.5"
                      style={{ backgroundColor: NAVY }}
                    />
                  )}
                </button>
                <button
                  type="button"
                  onClick={() => setActiveTab('raw_editor')}
                  className={`px-4 py-2.5 text-xs font-bold transition-all relative flex items-center gap-2 cursor-pointer ${activeTab === 'raw_editor'
                    ? 'text-[#1B365D]'
                    : 'text-slate-400 hover:text-slate-600'
                    }`}
                >
                  <Database className="w-4 h-4" />
                  แก้ไข/กรอกข้อมูลดิบสด ({records.length})
                  {activeTab === 'raw_editor' && (
                    <motion.div
                      layoutId="activeTabUnderline"
                      className="absolute bottom-0 left-0 right-0 h-0.5"
                      style={{ backgroundColor: NAVY }}
                    />
                  )}
                </button>
              </div>

              {/* DOWNLOAD REPORT SYSTEM */}
              <div className="flex items-center gap-2 pb-4 sm:pb-0 w-full sm:w-auto justify-end">
                <button
                  type="button"
                  onClick={handleExportCSV}
                  className="px-3.5 py-2 text-slate-600 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg text-xs font-medium cursor-pointer transition flex items-center gap-1.5"
                  id="csv-download-btn"
                >
                  <FileDown className="w-3.5 h-3.5" />
                  ข้อมูลดิบ (.csv)
                </button>
                <button
                  type="button"
                  onClick={handleExportExcel}
                  disabled={isProcessing}
                  className="px-4 py-2 text-white rounded-lg text-xs font-bold cursor-pointer transition-all shadow-sm hover:brightness-110 flex items-center gap-2 disabled:opacity-40"
                  style={{ backgroundColor: NAVY }}
                  id="excel-download-btn"
                >
                  <Download className="w-3.5 h-3.5" />
                  ดาวน์โหลดรายงานสรุป Excel ✨
                </button>
              </div>

            </div>

            {/* TAB CONTENT */}
            <div className="p-3 md:p-5 min-h-[500px]">

              <AnimatePresence mode="wait">
                {activeTab === 'preview' ? (
                  <motion.div
                    key="tab-preview"
                    initial={{ opacity: 0, y: 5 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -5 }}
                    transition={{ duration: 0.15 }}
                    className="space-y-4"
                  >
                    {/* HELP TIP */}
                    <div className="flex items-start gap-2 bg-blue-50/60 border border-blue-100 p-3.5 rounded-xl text-xs text-slate-600">
                      <HelpCircle className="w-4 h-4 shrink-0 mt-0.5 text-[#1B365D]" />
                      <div>
                        ระบบจำลองตารางนี้คำนวณและแสดงผลในโครงสร้างแบบเดียวกับไฟล์ Excel จริง! แถวสีเทาคือ <strong>รวมย่อยรายร้านค้า (Subtotal)</strong> แถวสีครีมคือ <strong>รวมประจำจังหวัด</strong> และสีน้ำเงินสุดท้ายคือ <strong>ยอดรวมทั้งหมดพร้อมฟิลด์สูตร SUM แบบไดนามิก</strong>
                      </div>
                    </div>

                    {/* INTERACTIVE GRAPHICS TABLE PREVIEW */}
                    <div className="overflow-x-auto rounded-xl border border-slate-200 relative bg-white">

                      <table className="w-full text-left text-xs border-collapse">

                        {/* Table Header */}
                        <thead>
                          <tr className="text-white" style={{ backgroundColor: NAVY }}>
                            <th className="p-2.5 font-bold text-center border-r border-white/15 w-[12%]">จังหวัด</th>
                            <th className="p-2.5 font-bold border-r border-white/15 w-[23%]">ร้านค้า</th>
                            <th className="p-2.5 font-bold text-center border-r border-white/15 w-[12%]">เลขที่บิล</th>
                            <th className="p-2.5 font-bold text-center border-r border-white/15 w-[12%]">รหัสสินค้า</th>
                            <th className="p-2.5 font-bold border-r border-white/15 w-[29%]">สินค้า</th>
                            <th className="p-2.5 font-bold text-right w-[12%]">จำนวน(หีบ)</th>
                          </tr>
                        </thead>

                        {/* Table Body */}
                        <tbody>
                          {previewRows.tableRows.length === 0 ? (
                            <tr>
                              <td colSpan={6} className="text-center py-20 text-slate-400">
                                <div className="flex flex-col items-center justify-center gap-3">
                                  <Table className="w-10 h-10 text-slate-300" />
                                  <span className="text-sm">ไม่มีข้อมูลในจังหวัดที่เลือกสำหรับตารางแสดงผลช่วงนี้</span>
                                  <p className="text-xs text-slate-400 max-w-xs">ลองเช็คค่าจังหวัด หรืออัปโหลดไฟล์ Excel เพื่อเริ่มระบบ</p>
                                </div>
                              </td>
                            </tr>
                          ) : (
                            previewRows.tableRows.map((row) => {

                              if (row.type === 'data' && row.record) {
                                const { province, store, bill, pCode, pName, qty } = row.record;
                                return (
                                  <tr
                                    key={row.key}
                                    className={`border-b border-slate-100 hover:bg-blue-50/40 transition-colors ${row.isZebra ? 'bg-slate-50/70' : 'bg-white'
                                      }`}
                                  >
                                    <td className="p-2 text-center text-slate-600 font-medium border-r border-slate-100">{province}</td>
                                    <td className="p-2 truncate text-slate-700 border-r border-slate-100" title={store}>{store}</td>
                                    <td className="p-2 text-center text-slate-500 font-mono border-r border-slate-100">{bill || "-"}</td>
                                    <td className="p-2 text-center text-slate-500 font-mono border-r border-slate-100">{pCode || "-"}</td>
                                    <td className="p-2 text-slate-700 truncate border-r border-slate-100" title={pName}>{pName}</td>
                                    <td className="p-2 text-right text-emerald-600 font-semibold font-mono">{qty.toLocaleString()}</td>
                                  </tr>
                                );
                              }

                              if (row.type === 'store_subtotal') {
                                return (
                                  <tr key={row.key} className="bg-slate-100 border-y border-slate-200">
                                    <td colSpan={5} className="p-2.5 pl-4 text-left font-semibold text-slate-600">
                                      {row.label}
                                    </td>
                                    <td className="p-2.5 text-right font-bold font-mono text-[#1B365D]">
                                      {row.qtyValue?.toLocaleString()}
                                    </td>
                                  </tr>
                                );
                              }

                              if (row.type === 'prov_total') {
                                return (
                                  <tr key={row.key} className="bg-amber-50 border-y border-amber-200">
                                    <td colSpan={5} className="p-3 pl-4 text-left font-bold text-amber-800">
                                      {row.label}
                                    </td>
                                    <td className="p-3 text-right text-amber-700 font-extrabold font-mono">
                                      {row.qtyValue?.toLocaleString()}
                                    </td>
                                  </tr>
                                );
                              }

                              if (row.type === 'empty') {
                                return (
                                  <tr key={row.key} className="h-3 bg-white">
                                    <td colSpan={6} className="p-0 border-none"></td>
                                  </tr>
                                );
                              }

                              if (row.type === 'grand_total') {
                                return (
                                  <tr key={row.key} className="text-white font-extrabold" style={{ backgroundColor: NAVY }}>
                                    <td colSpan={5} className="p-3.5 pl-4 text-left">
                                      {row.label}
                                    </td>
                                    <td className="p-3.5 text-right text-emerald-300 text-sm font-extrabold font-mono">
                                      {row.qtyValue?.toLocaleString()}
                                    </td>
                                  </tr>
                                );
                              }

                              return null;
                            })
                          )}
                        </tbody>
                      </table>

                    </div>

                    {/* DATA INTEGRITY / CALCULATION DETAILS FOOTER */}
                    {previewRows.tableRows.length > 0 && (
                      <div className="flex flex-col sm:flex-row items-center justify-between text-xs text-slate-400 gap-2 px-1">
                        <div>
                          <span>รวมยอดคำนวณทั้งหมด {previewRows.recordCount} รายการจัดสรรสินค้าในระบบ</span>
                        </div>
                        <div className="bg-slate-100 px-2 py-1 rounded text-slate-500 text-[10px] font-mono">
                          รูปแบบสรุปรหัส: SUM(F{`{X}`}:F{`{Y}`})
                        </div>
                      </div>
                    )}

                  </motion.div>
                ) : (
                  <motion.div
                    key="tab-raw-editor"
                    initial={{ opacity: 0, y: 5 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -5 }}
                    transition={{ duration: 0.15 }}
                    className="space-y-6"
                  >

                    {/* ADD NEW RECORD FORM AREA */}
                    <div className="bg-slate-50 rounded-xl border border-slate-200 p-4">
                      <h3 className="text-xs font-bold text-slate-600 uppercase tracking-wider mb-3 flex items-center gap-1.5">
                        <Plus className="w-4 h-4 text-emerald-500" />
                        เพิ่มรายการส่งมอบใหม่เพื่อการทดสอบ
                      </h3>

                      <form onSubmit={handleAddNewRow} className="grid grid-cols-1 sm:grid-cols-6 gap-3">

                        {/* Prov selection */}
                        <div className="col-span-1">
                          <label className="block text-[10px] text-slate-500 font-medium mb-1">จังหวัด</label>
                          <select
                            value={newRow.province}
                            onChange={(e) => setNewRow(p => ({ ...p, province: e.target.value }))}
                            className="w-full bg-white border border-slate-300 rounded-lg px-2 py-1.5 text-xs text-[#1B365D] font-medium focus:outline-none focus:border-[#1B365D]"
                          >
                            <option value="พิษณุโลก">พิษณุโลก</option>
                            <option value="สุโขทัย">สุโขทัย</option>
                            <option value="กรุงเทพมหานคร">กรุงเทพฯ</option>
                            <option value="เชียงใหม่">เชียงใหม่</option>
                          </select>
                        </div>

                        {/* Store name */}
                        <div className="col-span-1 sm:col-span-2">
                          <label className="block text-[10px] text-slate-500 font-medium mb-1">ร้านค้าปลายทาง</label>
                          <input
                            type="text"
                            value={newRow.store}
                            onChange={(e) => setNewRow(p => ({ ...p, store: e.target.value }))}
                            className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-[#1B365D]"
                            placeholder="เช่น ร้านเจริญพานิช"
                          />
                        </div>

                        {/* Bill number */}
                        <div className="col-span-1">
                          <label className="block text-[10px] text-slate-500 font-medium mb-1">เลขที่บิล</label>
                          <input
                            type="text"
                            value={newRow.bill}
                            onChange={(e) => setNewRow(p => ({ ...p, bill: e.target.value }))}
                            className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 placeholder-slate-400 font-mono focus:outline-none focus:border-[#1B365D]"
                            placeholder="INV..."
                          />
                        </div>

                        {/* Product Code */}
                        <div className="col-span-1">
                          <label className="block text-[10px] text-slate-500 font-medium mb-1">รหัสสินค้า</label>
                          <input
                            type="text"
                            value={newRow.pCode}
                            onChange={(e) => setNewRow(p => ({ ...p, pCode: e.target.value }))}
                            className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 placeholder-slate-400 font-mono focus:outline-none focus:border-[#1B365D]"
                            placeholder="PD-..."
                          />
                        </div>

                        {/* Qty box */}
                        <div className="col-span-1">
                          <label className="block text-[10px] text-emerald-600 font-semibold mb-1">จํานวนสินค้า (หีบ)</label>
                          <input
                            type="number"
                            value={newRow.qty || ''}
                            onChange={(e) => setNewRow(p => ({ ...p, qty: Number(e.target.value) }))}
                            className="w-full bg-emerald-50 border-2 border-emerald-200 rounded-lg px-2 py-1.5 text-xs text-emerald-700 font-bold focus:outline-none focus:border-emerald-500"
                            placeholder="0"
                            min="0"
                          />
                        </div>

                        {/* Truck registration */}
                        <div className="col-span-1">
                          <label className="block text-[10px] text-sky-600 font-semibold mb-1">ทะเบียนรถ</label>
                          <input
                            type="text"
                            value={newRow.truck}
                            onChange={(e) => setNewRow(p => ({ ...p, truck: e.target.value }))}
                            className="w-full bg-sky-50 border border-sky-200 rounded-lg px-2.5 py-1.5 text-xs text-sky-700 placeholder-slate-400 font-mono focus:outline-none focus:border-sky-500"
                            placeholder="1กข-1234"
                          />
                        </div>

                        {/* Product Name */}
                        <div className="col-span-1 sm:col-span-4">
                          <label className="block text-[10px] text-slate-500 font-medium mb-1">รายละเอียดสินค้า</label>
                          <input
                            type="text"
                            value={newRow.pName}
                            onChange={(e) => setNewRow(p => ({ ...p, pName: e.target.value }))}
                            className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-[#1B365D]"
                            placeholder="ระบุชื่อแบรนด์หรือรายละเอียดแพ็กเช่น ข้าวสกัด ตราน้ำทิพย์ 10กิโล"
                          />
                        </div>

                        {/* Action Submit */}
                        <div className="col-span-1 sm:col-span-1 flex items-end">
                          <button
                            type="submit"
                            className="w-full bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg py-2 text-xs font-bold transition flex items-center justify-center gap-1 cursor-pointer shadow-sm"
                          >
                            <Plus className="w-3.5 h-3.5" />
                            เพิ่มข้อมูล
                          </button>
                        </div>

                      </form>
                    </div>

                    {/* SOURCE RAW RECORDS LIST */}
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <h3 className="text-xs font-bold text-slate-600 uppercase tracking-wider">
                          รายการแถวข้อมูลดิบในหน้าแอป ({records.length} แถวปัจจุบัน)
                        </h3>
                        <span className="text-[11px] text-slate-400 font-mono">
                          ตารางแผ่นงานข้อมูลดิบสามารถแก้ไขค่าได้ทันที
                        </span>
                      </div>

                      <div className="overflow-x-auto rounded-xl border border-slate-200 max-h-96">
                        <table className="w-full text-left text-xs">
                          <thead className="sticky top-0">
                            <tr className="bg-slate-100 border-b border-slate-200 text-slate-500 font-semibold">
                              <th className="p-2.5 w-[13%]">จังหวัด</th>
                              <th className="p-2.5 w-[12%] border-l border-slate-200">ทะเบียนรถ</th>
                              <th className="p-2.5 w-[21%] border-l border-slate-200">ร้านค้าปลายทาง</th>
                              <th className="p-2.5 w-[13%] border-l border-slate-200">เลขที่บิล</th>
                              <th className="p-2.5 w-[11%] border-l border-slate-200">รหัส</th>
                              <th className="p-2.5 w-[16%] border-l border-slate-200">สินค้า</th>
                              <th className="p-2.5 w-[10%] text-right border-l border-slate-200">จำนวน(L)</th>
                              <th className="p-2.5 w-[4%] text-center">ลบ</th>
                            </tr>
                          </thead>
                          <tbody>
                            {records.length === 0 ? (
                              <tr>
                                <td colSpan={8} className="text-center py-16 text-slate-400">
                                  ยังไม่มีข้อมูลในระบบ — อัปโหลดไฟล์ Excel หรือเพิ่มแถวด้วยฟอร์มด้านบนเพื่อเริ่มต้น
                                </td>
                              </tr>
                            ) : (
                              records.map((rec) => (
                                <tr key={rec.id} className="border-b border-slate-100 hover:bg-slate-50 transition">

                                  {/* Prov */}
                                  <td className="p-1">
                                    <input
                                      type="text"
                                      value={rec.province}
                                      onChange={(e) => handleEditRecordField(rec.id, 'province', e.target.value)}
                                      className="w-full bg-transparent border border-transparent hover:border-slate-200 focus:bg-white focus:border-[#1B365D] rounded px-1.5 py-1 text-xs text-[#1B365D] font-medium focus:outline-none"
                                    />
                                  </td>

                                  {/* Truck */}
                                  <td className="p-1 border-l border-slate-100">
                                    <input
                                      type="text"
                                      value={rec.truck}
                                      onChange={(e) => handleEditRecordField(rec.id, 'truck', e.target.value)}
                                      className="w-full bg-transparent border border-transparent hover:border-slate-200 focus:bg-white focus:border-sky-500 rounded px-1.5 py-1 text-xs text-sky-700 font-mono focus:outline-none"
                                    />
                                  </td>

                                  {/* Store */}
                                  <td className="p-1 border-l border-slate-100">
                                    <input
                                      type="text"
                                      value={rec.store}
                                      onChange={(e) => handleEditRecordField(rec.id, 'store', e.target.value)}
                                      className="w-full bg-transparent border border-transparent hover:border-slate-200 focus:bg-white focus:border-[#1B365D] rounded px-1.5 py-1 text-xs text-slate-700 focus:outline-none"
                                    />
                                  </td>

                                  {/* Bill */}
                                  <td className="p-1 border-l border-slate-100">
                                    <input
                                      type="text"
                                      value={rec.bill}
                                      onChange={(e) => handleEditRecordField(rec.id, 'bill', e.target.value)}
                                      className="w-full bg-transparent border border-transparent hover:border-slate-200 focus:bg-white focus:border-[#1B365D] rounded px-1.5 py-1 text-xs text-slate-500 font-mono focus:outline-none"
                                    />
                                  </td>

                                  {/* Code */}
                                  <td className="p-1 border-l border-slate-100">
                                    <input
                                      type="text"
                                      value={rec.pCode}
                                      onChange={(e) => handleEditRecordField(rec.id, 'pCode', e.target.value)}
                                      className="w-full bg-transparent border border-transparent hover:border-slate-200 focus:bg-white focus:border-[#1B365D] rounded px-1.5 py-1 text-xs text-slate-500 font-mono focus:outline-none"
                                    />
                                  </td>

                                  {/* Product name */}
                                  <td className="p-1 border-l border-slate-100">
                                    <input
                                      type="text"
                                      value={rec.pName}
                                      onChange={(e) => handleEditRecordField(rec.id, 'pName', e.target.value)}
                                      className="w-full bg-transparent border border-transparent hover:border-slate-200 focus:bg-white focus:border-[#1B365D] rounded px-1.5 py-1 text-xs text-slate-700 truncate focus:outline-none"
                                    />
                                  </td>

                                  {/* Qty (col L) */}
                                  <td className="p-1 border-l border-slate-100">
                                    <input
                                      type="number"
                                      value={rec.qty}
                                      onChange={(e) => handleEditRecordField(rec.id, 'qty', e.target.value)}
                                      className="w-full bg-transparent border border-transparent hover:border-emerald-300 focus:bg-emerald-50 focus:border-emerald-500 rounded px-1 py-1 text-xs text-emerald-700 font-bold font-mono text-right focus:outline-none"
                                    />
                                  </td>

                                  {/* Delete row */}
                                  <td className="p-1 text-center">
                                    <button
                                      type="button"
                                      onClick={() => handleDeleteRecord(rec.id)}
                                      className="text-slate-400 hover:text-rose-500 p-1 rounded transition cursor-pointer"
                                      title="ลบแถวนี้"
                                    >
                                      <Trash2 className="w-3.5 h-3.5" />
                                    </button>
                                  </td>

                                </tr>
                              ))
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>

                  </motion.div>
                )}
              </AnimatePresence>

            </div>

          </div>

          {/* PYTHON CONTEXT BRIEF TO BUILD CONFIDENCE */}
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-col gap-3">
            <h3 className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
              <CheckCircle className="w-4 h-4 text-emerald-500" />
              การรับประกันคุณภาพรายงานด้านความงามเลย์เอาต์ (Visual Guarantee)
            </h3>
            <p className="text-xs text-slate-500 leading-relaxed">
              สคริปต์ส่งออก Excel รันแบบ client-side ผ่านเฟรมเวิร์ก <strong>ExcelJS</strong> ด้วยโค้ดที่ถอดแบบสัดส่วน, ฟอนต์สไตล์ <code className="text-[10px] bg-slate-100 text-[#1B365D] px-1 py-0.5 rounded font-mono">Cordia New</code>, และรหัสค่าสีน้ำเงิน-เทาสุภาพ (Solid Navy headers #1B365D, Light Gray Subtotal #F2F4F4, Cream Yellow State totals #FEF9E7) จากโครงสร้างโปรแกรมเดิมเพื่อรับประกันไฟล์รายงานที่ส่งมอบงานให้พี่ได้อย่างถูกต้องไม่มีค้างคาใจแน่นอนครับ!
            </p>
          </div>

        </section>

      </main>

      {/* FOOTER */}
      <footer className="bg-white border-t border-slate-200 py-4 px-6 text-center text-xs text-slate-400 mt-6">
        <div className="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-4">
          <p>© {new Date().getFullYear()} ระบบแปลงรายงานแยกกลุ่มจังหวัดพร้อม Subtotal ยอดรวมย่อย</p>
          <p className="text-[11px] text-slate-400 font-mono">ดึงยอดจำนวนคอลัมน์ L และแยกชีตตามทะเบียนรถเรียบร้อย</p>
        </div>
      </footer>

    </div>
  );
}
