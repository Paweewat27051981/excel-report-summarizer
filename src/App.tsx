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
  RefreshCw,
  Database,
  FileDown,
  CheckCircle,
  AlertCircle,
  HelpCircle
} from 'lucide-react';

import { RawRecord, ColumnMapping } from './types';
import { parseUploadedExcel, generateFilteredReport } from './utils/excelProcessor';

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
        "\uFEFF" + headers.join(","), // UTF-8 BOM
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
    const provTotalRowsAccumulator: { prov: string; sum: number }[] = [];

    activeProvincesInFiltered.forEach((targetProv) => {
      const provRecords = filteredDataset.filter(r => r.province === targetProv);
      if (provRecords.length === 0) return;

      let idx = 0;
      let provSumAccumulated = 0;
      const storeTotals: number[] = [];

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

  return (
    <div className="min-h-screen bg-slate-900 text-slate-100 flex flex-col font-sans antialiased">

      {/* HEADER BANNER */}
      <header className="bg-gradient-to-r from-slate-950 via-slate-900 to-[#12243d] border-b border-slate-800 py-5 px-6 shadow-xl sticky top-0 z-40">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row items-center justify-between gap-4">

          <div className="flex items-center gap-4">
            <div className="bg-[#1B365D] text-white p-3 rounded-xl shadow-md border border-slate-700 flex items-center justify-center">
              <FileSpreadsheet className="w-8 h-8 text-blue-200" id="header-icon" />
            </div>
            <div>
              <h1 className="text-xl md:text-2xl font-bold tracking-tight text-white flex items-center gap-2">
                ระบบจัดการรายงานและสรุปยอดส่งสินค้าแยกรายร้านค้า
              </h1>
              <p className="text-sm text-slate-400">
                ประมวลผล ดึงคอลัมน์ L อัตโนมัติ พร้อมคำนวณแถวสรุปยอดรวมย่อย (Subtotal) สำหรับ พิษณุโลก & สุโขทัย
              </p>
            </div>
          </div>

          {/* QUICK SYSTEM ACTIONS */}
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={handleClearAll}
              type="button"
              className="px-3 py-1.5 bg-slate-800 hover:bg-rose-950/40 hover:text-rose-400 text-slate-400 text-xs rounded-lg border border-slate-800 transition flex items-center gap-1 cursor-pointer font-medium"
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
        <div className="bg-slate-950 border-b border-slate-800 py-3 px-4">
          <div className="max-w-7xl mx-auto flex items-start gap-3">
            {feedback.type === 'success' && <CheckCircle className="w-5 h-5 text-emerald-400 shrink-0 mt-0.5" id="feedback-success" />}
            {feedback.type === 'error' && <AlertCircle className="w-5 h-5 text-rose-400 shrink-0 mt-0.5" id="feedback-error" />}
            {feedback.type === 'info' && <Info className="w-5 h-5 text-sky-400 shrink-0 mt-0.5" id="feedback-info" />}
            <p className={`text-xs md:text-sm ${feedback.type === 'error' ? 'text-rose-300' : feedback.type === 'success' ? 'text-emerald-300' : 'text-slate-300'}`}>
              <strong>แจ้งระบบ:</strong> {feedback.message}
            </p>
          </div>
        </div>
      )}

      {/* DASHBOARD ANALYTICS OVERVIEW */}
      <section className="bg-slate-900 border-b border-slate-800/80 py-6 px-6">
        <div className="max-w-7xl mx-auto grid grid-cols-2 md:grid-cols-4 gap-4">

          <div className="bg-slate-950 rounded-2xl p-4 border border-slate-800 shadow-md">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-widest block">จำนวนรายการที่กรอง</span>
            <div className="flex items-baseline gap-2 mt-2">
              <span className="text-2xl md:text-3xl font-extrabold text-blue-400">{previewRows.recordCount}</span>
              <span className="text-xs text-slate-400">รายการ</span>
            </div>
            <div className="text-[11px] text-slate-500 mt-1">จากทั้งหมด {records.length} รายการ</div>
          </div>

          <div className="bg-slate-950 rounded-2xl p-4 border border-slate-800 shadow-md">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-widest block">ร้านค้าทั้งหมด</span>
            <div className="flex items-baseline gap-2 mt-2">
              <span className="text-2xl md:text-3xl font-extrabold text-purple-400">{previewRows.storeCount}</span>
              <span className="text-xs text-slate-400">ร้านค้า</span>
            </div>
            <div className="text-[11px] text-slate-500 mt-1">แยกเป็นกลุ่มจังหวัดปลายทาง</div>
          </div>

          <div className="bg-slate-950 rounded-2xl p-4 border border-slate-800 shadow-md">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-widest block">ปริมาณรวบสุทธิ (หีบ)</span>
            <div className="flex items-baseline gap-2 mt-2">
              <span className="text-2xl md:text-3xl font-extrabold text-emerald-400">{previewRows.grandTotal.toLocaleString()}</span>
              <span className="text-xs text-slate-400">หีบ</span>
            </div>
            <div className="text-[11px] text-slate-500 mt-1">อ้างอิงจากคอลัมน์ L ใน Excel</div>
          </div>

          <div className="bg-slate-950 rounded-2xl p-4 border border-slate-800 shadow-md">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-widest block">กลุ่มจังหวัดที่เลือกกรอง</span>
            <div className="flex items-center gap-1.5 mt-2 flex-wrap">
              {selectedProvinces.length === 0 ? (
                <span className="text-xs text-slate-400">ไม่ได้เลือก</span>
              ) : (
                selectedProvinces.map(p => (
                  <span key={p} className="bg-slate-800 text-xs px-2 py-0.5 rounded text-yellow-400 border border-slate-700/60 font-medium">
                    {p}
                  </span>
                ))
              )}
            </div>
            <div className="text-[11px] text-slate-500 mt-2">รวมทั้งหมด {previewRows.provinceCount} จังหวัดในตารางจำลอง</div>
          </div>

        </div>
      </section>

      {/* MAIN LAYOUT GRID */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 md:p-6 grid grid-cols-1 lg:grid-cols-12 gap-6">

        {/* LEFT COLUMN: CONTROL & SETTINGS (4 spans) */}
        <section className="lg:col-span-4 flex flex-col gap-6">

          {/* UPLOAD & PARSER SETTINGS */}
          <div className="bg-slate-950 rounded-2xl border border-slate-800 p-5 shadow-lg">

            <h2 className="text-md font-bold text-white mb-4 flex items-center gap-2 pb-2 border-b border-slate-800">
              <Upload className="w-5 h-5 text-blue-400" />
              1. นำเข้าไฟล์ Excel จริง
            </h2>

            <div
              onClick={() => fileInputRef.current?.click()}
              onDragOver={handleDragOver}
              onDrop={handleDrop}
              className="border-2 border-dashed border-slate-800 hover:border-slate-700 hover:bg-slate-900/80 bg-slate-900/50 p-6 rounded-xl flex flex-col items-center justify-center text-center cursor-pointer transition-all group"
            >
              <input
                type="file"
                ref={fileInputRef}
                onChange={handleFileUpload}
                accept=".xlsx,.xls"
                className="hidden"
              />
              <FileSpreadsheet className="w-12 h-12 text-slate-400 group-hover:text-blue-400 group-hover:scale-105 transition-all mb-3" />
              <p className="text-sm font-semibold text-slate-300">ลากหรือคลิกเพื่ออัปโหลดไฟล์ Excel</p>
              <p className="text-xs text-slate-500 mt-1">รองรับไฟล์ตระกูล .xlsx (ไฟล์สรุปรายจังหวัด)</p>
            </div>

            <div className="mt-4 bg-slate-900/60 rounded-lg p-3 text-xs text-slate-400 border border-slate-850">
              <span className="font-semibold text-slate-300 flex items-center gap-1 mb-1">
                <Info className="w-3.5 h-3.5 text-blue-400" /> คำชี้แจงโครงสร้างไฟล์:
              </span>
              โปรแกรมจะอ่านข้อมูลแผ่นงานแรก เริ่มถอดแถวข้อมูลตั้งแต่แถวที่ 2 เป็นต้นไป (โดยถือว่าแถวที่ 1 เป็นส่วนหัวรายงานของพี่)
            </div>

          </div>

          {/* COLUMN MAPPING DEFINITIONS */}
          <div className="bg-slate-950 rounded-2xl border border-slate-800 p-5 shadow-lg">

            <h2 className="text-md font-bold text-white mb-4 flex items-center gap-2 pb-2 border-b border-slate-800">
              <Settings className="w-5 h-5 text-slate-400" />
              2. ตั้งค่าคอลัมน์ (Column Mapping)
            </h2>

            <p className="text-xs text-slate-400 mb-4">
              กำหนดคุณลักษณะเฉพาะคอลัมน์ของไฟล์ Excel เพื่อดึงค่าข้อมูลแต่ละหมวดให้ถูกต้อง:
            </p>

            <form className="space-y-3.5">

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-slate-400 mb-1">คอลัมน์ - จังหวัด</label>
                  <input
                    type="text"
                    value={columnMapping.province}
                    onChange={(e) => setColumnMapping(prev => ({ ...prev, province: e.target.value.toUpperCase() }))}
                    className="w-full bg-slate-900 border border-slate-800 rounded px-2.5 py-1.5 text-sm text-yellow-400 font-mono text-center focus:outline-none focus:border-blue-500"
                    placeholder="F"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-400 mb-1">คอลัมน์ - ร้านค้า</label>
                  <input
                    type="text"
                    value={columnMapping.store}
                    onChange={(e) => setColumnMapping(prev => ({ ...prev, store: e.target.value.toUpperCase() }))}
                    className="w-full bg-slate-900 border border-slate-800 rounded px-2.5 py-1.5 text-sm text-yellow-400 font-mono text-center focus:outline-none focus:border-blue-500"
                    placeholder="C"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-slate-400 mb-1">คอลัมน์ - เลขที่บิล</label>
                  <input
                    type="text"
                    value={columnMapping.bill}
                    onChange={(e) => setColumnMapping(prev => ({ ...prev, bill: e.target.value.toUpperCase() }))}
                    className="w-full bg-slate-900 border border-slate-800 rounded px-2.5 py-1.5 text-sm font-mono text-center focus:outline-none focus:border-blue-500"
                    placeholder="I"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-400 mb-1">คอลัมน์ - รหัสสินค้า</label>
                  <input
                    type="text"
                    value={columnMapping.pCode}
                    onChange={(e) => setColumnMapping(prev => ({ ...prev, pCode: e.target.value.toUpperCase() }))}
                    className="w-full bg-slate-900 border border-slate-800 rounded px-2.5 py-1.5 text-sm font-mono text-center focus:outline-none focus:border-blue-500"
                    placeholder="J"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3 border-t border-slate-850 pt-2.5">
                <div>
                  <label className="block text-xs font-medium text-slate-400 mb-1">คอลัมน์ - สินค้า</label>
                  <input
                    type="text"
                    value={columnMapping.pName}
                    onChange={(e) => setColumnMapping(prev => ({ ...prev, pName: e.target.value.toUpperCase() }))}
                    className="w-full bg-slate-900 border border-slate-800 rounded px-2.5 py-1.5 text-sm font-mono text-center focus:outline-none focus:border-blue-500"
                    placeholder="K"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-emerald-400 mb-1">🎯 ยอดสินค้า (คอลัมน์ L)</label>
                  <input
                    type="text"
                    value={columnMapping.qty}
                    onChange={(e) => setColumnMapping(prev => ({ ...prev, qty: e.target.value.toUpperCase() }))}
                    className="w-full bg-slate-900 border-2 border-emerald-900/60 rounded px-2.5 py-1.5 text-sm text-emerald-300 font-bold font-mono text-center focus:outline-none focus:border-emerald-500"
                    placeholder="L"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3 border-t border-slate-850 pt-2.5">
                <div>
                  <label className="block text-xs font-semibold text-sky-400 mb-1">🚚 ทะเบียนรถ (คอลัมน์ Q)</label>
                  <input
                    type="text"
                    value={columnMapping.truck}
                    onChange={(e) => setColumnMapping(prev => ({ ...prev, truck: e.target.value.toUpperCase() }))}
                    className="w-full bg-slate-900 border-2 border-sky-900/60 rounded px-2.5 py-1.5 text-sm text-sky-300 font-bold font-mono text-center focus:outline-none focus:border-sky-500"
                    placeholder="Q"
                  />
                </div>
                <div className="flex items-end">
                  <p className="text-[11px] text-slate-500 leading-tight pb-1">
                    ใช้สร้างชีต <span className="text-sky-400 font-semibold">"สรุปยอดตามทะเบียนรถ"</span> เพิ่มในไฟล์เดียวกัน
                  </p>
                </div>
              </div>

              <div className="bg-slate-900/50 p-2.5 rounded text-[11px] text-slate-500 text-center font-mono">
                ค่าเริ่มต้น: F=จังหวัด | C=ร้าน | I=บิล | J=รหัส | K=สินค้า | L=จำนวน(หีบ) | Q=ทะเบียนรถ
              </div>

            </form>

          </div>

          {/* DYNAMIC PROVINCE FILTERING */}
          <div className="bg-slate-950 rounded-2xl border border-slate-800 p-5 shadow-lg">

            <h2 className="text-md font-bold text-white mb-4 flex items-center gap-2 pb-2 border-b border-slate-800">
              <Filter className="w-5 h-5 text-yellow-400" />
              3. เลือกกรองจังหวัดตรวจรับงาน
            </h2>

            <p className="text-xs text-slate-400 mb-3">
              ระบบดึงรายการจังวัดที่มีในไฟล์ปัจจุบันให้อัตโนมัติ ติ๊กเลือกกลุ่มจังหวัดที่ต้องการนำลงตาราง:
            </p>

            {allProvincesInDataset.length === 0 ? (
              <div className="py-4 text-center text-xs text-slate-600 bg-slate-900/50 rounded-lg">
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
                      className={`flex items-center justify-between p-2 rounded-lg border text-xs cursor-pointer transition-all ${isChecked
                        ? 'bg-slate-900 border-slate-700 text-yellow-400'
                        : 'bg-slate-950/60 border-slate-850 text-slate-400 hover:border-slate-800'
                        }`}
                    >
                      <div className="flex items-center gap-2.5">
                        <input
                          type="checkbox"
                          checked={isChecked}
                          onChange={() => handleProvinceToggle(prov)}
                          className="rounded border-slate-700 accent-yellow-400 text-slate-950 focus:ring-0 focus:ring-offset-0 w-4 h-4 cursor-pointer"
                        />
                        <span className={`font-semibold ${isHighlight ? 'text-blue-300' : ''}`}>
                          {prov}
                        </span>
                        {isHighlight && (
                          <span className="text-[10px] bg-blue-950 text-blue-300 px-1.5 py-0.2 rounded border border-blue-900/40">
                            เป้าหมาย
                          </span>
                        )}
                      </div>
                      <span className="font-mono text-[10px] text-slate-500">({count} แถว)</span>
                    </label>
                  );
                })}
              </div>
            )}

            <div className="flex items-center justify-between gap-2 mt-4 pt-4 border-t border-slate-850">
              <button
                type="button"
                onClick={() => setSelectedProvinces(["พิษณุโลก", "สุโขทัย"])}
                className="text-[11px] text-yellow-400/80 hover:text-yellow-400 transition cursor-pointer"
              >
                เลือกเฉพาะเป้าหมายหลัก
              </button>
              <button
                type="button"
                onClick={() => setSelectedProvinces(allProvincesInDataset)}
                className="text-[11px] text-slate-400 hover:text-white transition cursor-pointer"
              >
                เลือกทั้งหมด
              </button>
            </div>

          </div>

        </section>

        {/* RIGHT COLUMN: MAIN PREVIEW AND EDIT SYSTEM (8 spans) */}
        <section className="lg:col-span-8 flex flex-col gap-6">

          {/* TABS CONTAINER */}
          <div className="bg-slate-950 rounded-2xl border border-slate-800 shadow-xl overflow-hidden flex flex-col">

            {/* TABS HEADER AND EXPORTS */}
            <div className="bg-slate-950/90 border-b border-slate-800 px-4 pt-4 flex flex-col sm:flex-row items-center justify-between gap-4">

              {/* TAB SELECTORS */}
              <div className="flex border-b border-transparent gap-2 w-full sm:w-auto">
                <button
                  type="button"
                  onClick={() => setActiveTab('preview')}
                  className={`px-4 py-2.5 text-xs font-bold transition-all relative flex items-center gap-2 cursor-pointer ${activeTab === 'preview'
                    ? 'text-yellow-400'
                    : 'text-slate-400 hover:text-slate-200'
                    }`}
                >
                  <Table className="w-4 h-4" />
                  ดูตารางจำลอง Subtotal (Excel Preview)
                  {activeTab === 'preview' && (
                    <motion.div
                      layoutId="activeTabUnderline"
                      className="absolute bottom-0 left-0 right-0 h-0.5 bg-yellow-400"
                    />
                  )}
                </button>
                <button
                  type="button"
                  onClick={() => setActiveTab('raw_editor')}
                  className={`px-4 py-2.5 text-xs font-bold transition-all relative flex items-center gap-2 cursor-pointer ${activeTab === 'raw_editor'
                    ? 'text-yellow-400'
                    : 'text-slate-400 hover:text-slate-200'
                    }`}
                >
                  <Database className="w-4 h-4" />
                  แก้ไข/กรอกข้อมูลดิบสด ({records.length})
                  {activeTab === 'raw_editor' && (
                    <motion.div
                      layoutId="activeTabUnderline"
                      className="absolute bottom-0 left-0 right-0 h-0.5 bg-yellow-400"
                    />
                  )}
                </button>
              </div>

              {/* DOWNLOAD REPORT SYSTEM */}
              <div className="flex items-center gap-2 pb-4 sm:pb-0 w-full sm:w-auto justify-end">
                <button
                  type="button"
                  onClick={handleExportCSV}
                  className="px-3.5 py-1.5 text-slate-300 bg-slate-900 border border-slate-850 hover:bg-slate-800 rounded-lg text-xs font-medium cursor-pointer transition flex items-center gap-1.5"
                  id="csv-download-btn"
                >
                  <FileDown className="w-3.5 h-3.5" />
                  ข้อมูลดิบ (.csv)
                </button>
                <button
                  type="button"
                  onClick={handleExportExcel}
                  disabled={isProcessing}
                  className="px-4 py-2 bg-[#1B365D] hover:bg-blue-800 text-white rounded-lg text-xs font-bold cursor-pointer transition-all shadow-md hover:shadow-blue-900/30 border border-blue-800 flex items-center gap-2 disabled:opacity-40"
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
                    <div className="flex items-start gap-2 bg-[#12243d]/30 border border-blue-900/40 p-3.5 rounded-xl text-xs text-blue-300">
                      <HelpCircle className="w-4 h-4 shrink-0 mt-0.5" />
                      <div>
                        ระบบจำลองตารางนี้คำนวณและแสดงผลในโครงสร้างแบบเดียวกับไฟล์ Excel จริง! แถวสีเทาคือ <strong>รวมย่อยรายร้านค้า (Subtotal)</strong> แถวสีครีมคือ <strong>รวมประจำจังหวัด</strong> และสีฟ้าสุดท้ายคือ <strong>ยอดรวมทั้งหมดพร้อมฟิลด์สูตร SUM แบบไดนามิก</strong>
                      </div>
                    </div>

                    {/* INTERACTIVE GRAPHICS TABLE PREVIEW */}
                    <div className="overflow-x-auto rounded-xl border border-slate-800 relative bg-slate-950">

                      <table className="w-full text-left text-xs border-collapse">

                        {/* Table Header */}
                        <thead>
                          <tr className="bg-[#1B365D] text-white border-b border-slate-850">
                            <th className="p-2.5 font-bold text-center border-r border-slate-800/40 w-[12%]">จังหวัด</th>
                            <th className="p-2.5 font-bold border-r border-slate-800/40 w-[23%]">ร้านค้า</th>
                            <th className="p-2.5 font-bold text-center border-r border-slate-800/40 w-[12%]">เลขที่บิล</th>
                            <th className="p-2.5 font-bold text-center border-r border-slate-800/40 w-[12%]">รหัสสินค้า</th>
                            <th className="p-2.5 font-bold border-r border-slate-800/40 w-[29%]">สินค้า</th>
                            <th className="p-2.5 font-bold text-right w-[12%]">จำนวน(หีบ)</th>
                          </tr>
                        </thead>

                        {/* Table Body */}
                        <tbody>
                          {previewRows.tableRows.length === 0 ? (
                            <tr>
                              <td colSpan={6} className="text-center py-20 text-slate-500">
                                <div className="flex flex-col items-center justify-center gap-3">
                                  <Table className="w-10 h-10 text-slate-700" />
                                  <span className="text-sm">ไม่มีข้อมูลในจังหวัดที่เลือกสำหรับตารางแสดงผลช่วงนี้</span>
                                  <p className="text-xs text-slate-600 max-w-xs">ลองเช็คค่าจังหวัด ข้อมูลแฮมเบอร์เกอร์หรืออัปโหลดไฟล์ Excel เพื่อเริ่มระบบ</p>
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
                                    className={`border-b border-slate-900 hover:bg-slate-900/50 transition-colors ${row.isZebra ? 'bg-slate-900/25' : ''
                                      }`}
                                  >
                                    <td className="p-2 text-center text-slate-300 font-medium border-r border-slate-900">{province}</td>
                                    <td className="p-2 truncate text-slate-300 border-r border-slate-900" title={store}>{store}</td>
                                    <td className="p-2 text-center text-slate-400 font-mono border-r border-slate-900">{bill || "-"}</td>
                                    <td className="p-2 text-center text-slate-400 font-mono border-r border-slate-900">{pCode || "-"}</td>
                                    <td className="p-2 text-slate-300 truncate border-r border-slate-900" title={pName}>{pName}</td>
                                    <td className="p-2 text-right text-emerald-400 font-semibold font-mono">{qty.toLocaleString()}</td>
                                  </tr>
                                );
                              }

                              if (row.type === 'store_subtotal') {
                                return (
                                  <tr key={row.key} className="bg-slate-800/70 text-slate-300 font-semibold border-y border-slate-700/60">
                                    <td colSpan={5} className="p-2.5 pl-4 text-left font-semibold text-slate-400 font-sans">
                                      {row.label}
                                    </td>
                                    <td className="p-2.5 text-right text-sky-300 font-bold font-mono border-t border-b border-slate-700">
                                      {row.qtyValue?.toLocaleString()}
                                    </td>
                                  </tr>
                                );
                              }

                              if (row.type === 'prov_total') {
                                return (
                                  <tr key={row.key} className="bg-amber-950/40 text-orange-400 font-bold border-y border-amber-900/30">
                                    <td colSpan={5} className="p-3 pl-4 text-left text-amber-300">
                                      {row.label}
                                    </td>
                                    <td className="p-3 text-right text-amber-400 font-extrabold font-mono border-t border-b border-amber-900/40">
                                      {row.qtyValue?.toLocaleString()}
                                    </td>
                                  </tr>
                                );
                              }

                              if (row.type === 'empty') {
                                return (
                                  <tr key={row.key} className="h-4 bg-slate-950">
                                    <td colSpan={6} className="p-0 border-none"></td>
                                  </tr>
                                );
                              }

                              if (row.type === 'grand_total') {
                                return (
                                  <tr key={row.key} className="bg-[#12243d]/80 text-[#93c5fd] font-extrabold border-y-2 border-blue-500/50">
                                    <td colSpan={5} className="p-3.5 pl-4 text-left text-blue-200">
                                      {row.label}
                                    </td>
                                    <td className="p-3.5 text-right text-emerald-400 text-sm font-extrabold font-mono">
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
                      <div className="flex flex-col sm:flex-row items-center justify-between text-xs text-slate-500 gap-2 px-1">
                        <div>
                          <span>รวมยอดคำนวณทั้งหมด {previewRows.recordCount} รายการจัดสรรสินค้าในระบบ</span>
                        </div>
                        <div className="bg-slate-900 px-2 py-1 rounded text-slate-400 text-[10px] font-mono">
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
                    <div className="bg-slate-900/60 rounded-xl border border-slate-800 p-4">
                      <h3 className="text-xs font-bold text-slate-300 uppercase tracking-wider mb-3 flex items-center gap-1.5">
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
                            className="w-full bg-slate-950 border border-slate-800 rounded px-2 py-1.5 text-xs text-yellow-400 focus:outline-none"
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
                            className="w-full bg-slate-950 border border-slate-800 rounded px-2.5 py-1 text-xs text-white placeholder-slate-600 focus:outline-none focus:border-blue-500"
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
                            className="w-full bg-slate-950 border border-slate-800 rounded px-2.5 py-1 text-xs text-white placeholder-slate-600 font-mono focus:outline-none focus:border-blue-500"
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
                            className="w-full bg-slate-950 border border-slate-800 rounded px-2.5 py-1 text-xs text-white placeholder-slate-600 font-mono focus:outline-none focus:border-blue-500"
                            placeholder="PD-..."
                          />
                        </div>

                        {/* Qty box */}
                        <div className="col-span-1">
                          <label className="block text-[10px] text-emerald-500 font-semibold mb-1">จํานวนสินค้า (หีบ)</label>
                          <input
                            type="number"
                            value={newRow.qty || ''}
                            onChange={(e) => setNewRow(p => ({ ...p, qty: Number(e.target.value) }))}
                            className="w-full bg-slate-950 border-2 border-emerald-950 rounded px-2 text-xs text-emerald-400 font-bold focus:outline-none focus:border-emerald-500"
                            placeholder="0"
                            min="0"
                          />
                        </div>

                        {/* Truck registration */}
                        <div className="col-span-1">
                          <label className="block text-[10px] text-sky-500 font-semibold mb-1">ทะเบียนรถ</label>
                          <input
                            type="text"
                            value={newRow.truck}
                            onChange={(e) => setNewRow(p => ({ ...p, truck: e.target.value }))}
                            className="w-full bg-slate-950 border border-sky-950 rounded px-2.5 py-1 text-xs text-sky-300 placeholder-slate-600 font-mono focus:outline-none focus:border-sky-500"
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
                            className="w-full bg-slate-950 border border-slate-800 rounded px-2.5 py-1 text-xs text-slate-300 placeholder-slate-600 focus:outline-none focus:border-blue-500"
                            placeholder="ระบุชื่อแบรนด์หรือรายละเอียดแพ็กเช่น ข้าวสกัด ตราน้ำทิพย์ 10กิโล"
                          />
                        </div>

                        {/* Action Submit */}
                        <div className="col-span-1 sm:col-span-1 flex items-end">
                          <button
                            type="submit"
                            className="w-full bg-emerald-600 hover:bg-emerald-500 text-white rounded py-1.5 text-xs font-bold transition flex items-center justify-center gap-1 cursor-pointer"
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
                        <h3 className="text-xs font-bold text-slate-300 uppercase tracking-wider mb-1">
                          รายการแถวข้อมูลดิบในหน้าแอป ({records.length} แถวปัจจุบัน)
                        </h3>
                        <span className="text-[11px] text-slate-500 font-mono">
                          ตารางแผ่นงานข้อมูลดิบสามารถแก้ไขค่าได้ทันที
                        </span>
                      </div>

                      <div className="overflow-x-auto rounded-xl border border-slate-800 max-h-96">
                        <table className="w-full text-left text-xs">
                          <thead>
                            <tr className="bg-slate-900 border-b border-slate-800 text-slate-400 font-semibold font-mono">
                              <th className="p-2 w-[13%]">จังหวัด</th>
                              <th className="p-2 w-[12%] border-l border-slate-850">ทะเบียนรถ</th>
                              <th className="p-2 w-[21%] border-l border-slate-850">ร้านค้าปลายทาง</th>
                              <th className="p-2 w-[13%] border-l border-slate-850">เลขที่บิล</th>
                              <th className="p-2 w-[11%] border-l border-slate-850">รหัส</th>
                              <th className="p-2 w-[16%] border-l border-slate-850">สินค้า</th>
                              <th className="p-2 w-[10%] text-right border-l border-slate-850">จำนวน(L)</th>
                              <th className="p-2 w-[4%] text-center">ลบ</th>
                            </tr>
                          </thead>
                          <tbody>
                            {records.length === 0 ? (
                              <tr>
                                <td colSpan={8} className="text-center py-16 text-slate-500">
                                  ไม่มีข้อมูลแฝงในชีตเลย ติ๊ก "โหลดข้อมูลจำลอง" เมนูด้านบนเพื่อเริ่มต้น
                                </td>
                              </tr>
                            ) : (
                              records.map((rec) => (
                                <tr key={rec.id} className="border-b border-slate-900 hover:bg-slate-900/40 transition">

                                  {/* Prov */}
                                  <td className="p-1">
                                    <input
                                      type="text"
                                      value={rec.province}
                                      onChange={(e) => handleEditRecordField(rec.id, 'province', e.target.value)}
                                      className="w-full bg-transparent border border-transparent hover:border-slate-800 focus:bg-slate-950 focus:border-blue-500 rounded px-1.5 py-0.5 text-xs text-yellow-500 font-medium focus:outline-none"
                                    />
                                  </td>

                                  {/* Truck */}
                                  <td className="p-1 border-l border-slate-900">
                                    <input
                                      type="text"
                                      value={rec.truck}
                                      onChange={(e) => handleEditRecordField(rec.id, 'truck', e.target.value)}
                                      className="w-full bg-transparent border border-transparent hover:border-slate-800 focus:bg-slate-950 focus:border-sky-500 rounded px-1.5 py-0.5 text-xs text-sky-300 font-mono focus:outline-none"
                                    />
                                  </td>

                                  {/* Store */}
                                  <td className="p-1 border-l border-slate-900">
                                    <input
                                      type="text"
                                      value={rec.store}
                                      onChange={(e) => handleEditRecordField(rec.id, 'store', e.target.value)}
                                      className="w-full bg-transparent border border-transparent hover:border-slate-800 focus:bg-slate-950 focus:border-blue-500 rounded px-1.5 py-0.5 text-xs text-slate-300 focus:outline-none"
                                    />
                                  </td>

                                  {/* Bill */}
                                  <td className="p-1 border-l border-slate-900">
                                    <input
                                      type="text"
                                      value={rec.bill}
                                      onChange={(e) => handleEditRecordField(rec.id, 'bill', e.target.value)}
                                      className="w-full bg-transparent border border-transparent hover:border-slate-800 focus:bg-slate-950 focus:border-blue-500 rounded px-1.5 py-0.5 text-xs text-slate-400 font-mono focus:outline-none animate-none"
                                    />
                                  </td>

                                  {/* Code */}
                                  <td className="p-1 border-l border-slate-900">
                                    <input
                                      type="text"
                                      value={rec.pCode}
                                      onChange={(e) => handleEditRecordField(rec.id, 'pCode', e.target.value)}
                                      className="w-full bg-transparent border border-transparent hover:border-slate-800 focus:bg-slate-950 focus:border-blue-500 rounded px-1.5 py-0.5 text-xs text-slate-400 font-mono focus:outline-none"
                                    />
                                  </td>

                                  {/* Product name */}
                                  <td className="p-1 border-l border-slate-900">
                                    <input
                                      type="text"
                                      value={rec.pName}
                                      onChange={(e) => handleEditRecordField(rec.id, 'pName', e.target.value)}
                                      className="w-full bg-transparent border border-transparent hover:border-slate-800 focus:bg-slate-950 focus:border-blue-500 rounded px-1.5 py-0.5 text-xs text-slate-300 truncate focus:outline-none"
                                    />
                                  </td>

                                  {/* Qty (col L) */}
                                  <td className="p-1 border-l border-slate-900">
                                    <input
                                      type="number"
                                      value={rec.qty}
                                      onChange={(e) => handleEditRecordField(rec.id, 'qty', e.target.value)}
                                      className="w-full bg-transparent border border-transparent hover:border-emerald-600 focus:bg-emerald-950 focus:border-emerald-500 rounded px-1 text-xs text-emerald-400 font-bold font-mono text-right focus:outline-none"
                                    />
                                  </td>

                                  {/* Delete row */}
                                  <td className="p-1 text-center">
                                    <button
                                      type="button"
                                      onClick={() => handleDeleteRecord(rec.id)}
                                      className="text-slate-600 hover:text-rose-400 p-1 rounded transition cursor-pointer"
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
          <div className="bg-slate-950 rounded-2xl border border-slate-800 p-5 shadow-lg flex flex-col gap-3">
            <h3 className="text-xs font-bold text-white uppercase tracking-wider flex items-center gap-1.5">
              <CheckCircle className="w-4 h-4 text-emerald-400" />
              การรับประกันคุณภาพรายงานด้านความงามเลย์เอาต์ (Visual Guarantee)
            </h3>
            <p className="text-xs text-slate-400 leading-relaxed">
              สคริปต์ส่งออก Excel รันแบบ client-side ผ่านเฟรมเวิร์ก <strong>ExcelJS</strong> ด้วยโค้ดที่ถอดแบบสัดส่วน, ฟอนต์สไตล์ <code className="text-[10px] bg-slate-900 text-yellow-300 px-1 py-0.5 rounded font-mono">Cordia New</code>, และและรหัสค่าสีน้ำเงิน-เทาสุภาพ (Solid Navy headers #1B365D, Light Gray Subtotal #F2F4F4, Cream Yellow State totals #FEF9E7) จากโครงสร้างโปรแกรมเดิมเพื่อรับประกันไฟล์รายงานที่ส่งมอบงานให้พี่ได้อย่างถูกต้องไม่มีค้างคาใจแน่นอนครับ!
            </p>
          </div>

        </section>

      </main>

      {/* FOOTER */}
      <footer className="bg-slate-950 border-t border-slate-850 py-4 px-6 text-center text-xs text-slate-500 mt-12">
        <div className="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-4">
          <p>© {new Date().getFullYear()} ระบบแปลงรายงานแยกกลุ่มจังหวัดพร้อม Subtotal ยอดรวมย่อย</p>
          <p className="text-[11px] text-slate-600 font-mono">ขยายสปริงบีแอลเพื่อดึงยอดจำนวนคอลัมน์ L เรียบร้อย</p>
        </div>
      </footer>

    </div>
  );
}
