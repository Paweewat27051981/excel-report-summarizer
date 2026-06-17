/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { RawRecord, ColumnMapping } from '../types';

/**
 * Converts a column letter (e.g. "A", "C", "AA") to a 0-based index.
 */
export function columnLetterToIndex(letter: string): number {
  const temp = letter.toUpperCase().trim();
  let index = 0;
  for (let i = 0; i < temp.length; i++) {
    const charCode = temp.charCodeAt(i);
    if (charCode < 65 || charCode > 90) continue; // safety check
    index = index * 26 + (charCode - 64);
  }
  return index - 1; // 0-based
}

/**
 * Parse an uploaded Excel file using SheetJS (xlsx)
 * and extract a list of raw records based on the column mapping and selected provinces.
 */
export async function parseUploadedExcel(
  file: File,
  mapping: ColumnMapping,
  availableProvincesSet: Set<string>
): Promise<{ records: RawRecord[]; allProvinces: string[] }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Read options: convert sheet to 2D array of rows
        const rawRows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });

        const pIndex = columnLetterToIndex(mapping.province);
        const sIndex = columnLetterToIndex(mapping.store);
        const bIndex = columnLetterToIndex(mapping.bill);
        const pCodeIndex = columnLetterToIndex(mapping.pCode);
        const pNameIndex = columnLetterToIndex(mapping.pName);
        const qtyIndex = columnLetterToIndex(mapping.qty);
        const truckIndex = columnLetterToIndex(mapping.truck);

        const records: RawRecord[] = [];
        const provincesFound = new Set<string>();

        // Skip row 1 (index 0 is header row, similar to row_idx from 2 in Python)
        for (let i = 1; i < rawRows.length; i++) {
          const row = rawRows[i];
          if (!row || row.length === 0) continue;

          const provVal = row[pIndex];
          const prov = provVal !== undefined ? String(provVal).trim() : "";

          if (prov) {
            provincesFound.add(prov);
          }

          const storeVal = row[sIndex];
          const store = storeVal !== undefined ? String(storeVal).trim() : "";

          const billVal = row[bIndex];
          const bill = billVal !== undefined ? String(billVal).trim() : "";

          const pCodeVal = row[pCodeIndex];
          const pCode = pCodeVal !== undefined ? String(pCodeVal).trim() : "";

          const pNameVal = row[pNameIndex];
          const pName = pNameVal !== undefined ? String(pNameVal).trim() : "";

          const rawQty = row[qtyIndex];
          let qty = 0;
          if (rawQty !== undefined && rawQty !== null) {
            qty = Number(rawQty);
            if (isNaN(qty)) qty = 0;
          }

          const truckVal = row[truckIndex];
          const truck = truckVal !== undefined ? String(truckVal).trim() : "";

          // Generate a safe unique ID
          const id = `row-${i}-${Date.now()}`;

          records.push({
            id,
            province: prov,
            store,
            bill,
            pCode,
            pName,
            qty,
            truck
          });
        }

        resolve({
          records,
          allProvinces: Array.from(provincesFound)
        });
      } catch (err) {
        reject(new Error("ไม่สามารถอ่านไฟล์ Excel สเปกนี้ได้ กรุณาตรวจสอบและอัปโหลดไฟล์ที่ถูกต้องครับ"));
      }
    };
    reader.onerror = () => {
      reject(new Error("เกิดข้อผิดพลาดในการโหลดไฟล์"));
    };
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Generate a beautifully stylized Excel file according to the Python script guidelines.
 * Saves the file directly to the client's device using ExcelJS.
 */
export async function generateFilteredReport(
  rawRecords: RawRecord[],
  selectedProvinces: string[],
  outputFileName: string = "รายงานสรุปการจัดส่งสินค้า แยกรายร้านค้าพร้อมยอดรวมย่อย.xlsx"
): Promise<void> {
  // 1. Filter and sort similar to raw_data.sort(key=lambda x: (x[0], x[1], x[2]))
  // Index values in raw_data are: 0=prov, 1=store, 2=bill, 3=p_code, 4=p_name, 5=qty
  const filteredData = rawRecords.filter((rec) => selectedProvinces.includes(rec.province));

  // Sort: Province -> Store -> Bill
  filteredData.sort((a, b) => {
    if (a.province !== b.province) {
      return a.province.localeCompare(b.province, 'th');
    }
    if (a.store !== b.store) {
      return a.store.localeCompare(b.store, 'th');
    }
    return a.bill.localeCompare(b.bill, 'th');
  });

  if (filteredData.length === 0) {
    throw new Error("ไม่มีข้อมูลให้สกัดสำหรับจังหวัดที่เลือกครับพี่");
  }

  // 2. Initialize ExcelJS workbook
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("รายงานสรุปแยกจังหวัด", {
    views: [{ showGridLines: true }]
  });

  // Fonts definition (incorporating "Cordia New" style with fallback to TH Sarabun / Arial)
  const fontTitle: Partial<ExcelJS.Font> = { name: "Cordia New", size: 18, bold: true, color: { argb: "FF1B365D" } };
  const fontHeader: Partial<ExcelJS.Font> = { name: "Cordia New", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
  const fontBody: Partial<ExcelJS.Font> = { name: "Cordia New", size: 13 };
  const fontSubtotal: Partial<ExcelJS.Font> = { name: "Cordia New", size: 13, bold: true, color: { argb: "FF2C3E50" } };
  const fontProvTotal: Partial<ExcelJS.Font> = { name: "Cordia New", size: 14, bold: true, color: { argb: "FFD35400" } };
  const fontTotal: Partial<ExcelJS.Font> = { name: "Cordia New", size: 14, bold: true, color: { argb: "FF1B365D" } };

  // Fills
  const fillHeader: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1B365D" } };
  const fillSubtotal: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F4F4" } };
  const fillProvTotal: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF9E7" } };
  const fillTotal: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE6EEF8" } };
  const fillZebra: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF9FAFC" } };

  // Borders
  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: "thin", color: { argb: "FFCCCCCC" } },
    left: { style: "thin", color: { argb: "FFCCCCCC" } },
    bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
    right: { style: "thin", color: { argb: "FFCCCCCC" } }
  };

  const subtotalBorder: Partial<ExcelJS.Borders> = {
    top: { style: "thin", color: { argb: "FFAAAAAA" } },
    bottom: { style: "thin", color: { argb: "FFAAAAAA" } }
  };

  const totalBorder: Partial<ExcelJS.Borders> = {
    top: { style: "thin", color: { argb: "FF1B365D" } },
    bottom: { style: "double", color: { argb: "FF1B365D" } }
  };

  // Sheet title - row 1
  ws.getCell("A1").value = `รายงานสรุปการจัดส่งสินค้า แยกรายร้านค้าพร้อมยอดรวมย่อย (${selectedProvinces.join(" & ")})`;
  ws.getCell("A1").font = fontTitle;
  ws.getRow(1).height = 30;

  // Header table - row 3
  const headers = ["จังหวัด", "ร้านค้า", "เลขที่บิล", "รหัสสินค้า", "สินค้า", "จำนวน(หีบ)"];
  const headerRow = ws.getRow(3);
  headerRow.height = 25;

  headers.forEach((header, colIdx) => {
    const cell = headerRow.getCell(colIdx + 1);
    cell.value = header;
    cell.font = fontHeader;
    cell.fill = fillHeader;
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = thinBorder;
  });

  let currentRowNum = 4;
  const provTotalRows: number[] = [];

  // Loop provinces sequentially
  for (const targetProv of selectedProvinces) {
    const provData = filteredData.filter((x) => x.province === targetProv);
    if (provData.length === 0) continue;

    let idx = 0;
    const totalProvItems = provData.length;
    const storeSubtotalRows: number[] = [];

    // Loop store inside province
    while (idx < totalProvItems) {
      const currentStore = provData[idx].store;
      const startRow = currentRowNum;
      let zebraFlag = false;

      while (idx < totalProvItems && provData[idx].store === currentStore) {
        const item = provData[idx];
        const itemRow = ws.getRow(currentRowNum);
        itemRow.height = 20;

        const rowValues = [item.province, item.store, item.bill, item.pCode, item.pName, item.qty];

        rowValues.forEach((val, colIdx) => {
          const cell = itemRow.getCell(colIdx + 1);
          cell.value = val;
          cell.font = fontBody;
          cell.border = thinBorder;

          // Alignments according to spec
          if (colIdx === 0 || colIdx === 1 || colIdx === 4) {
            cell.alignment = { horizontal: "left", vertical: "middle" };
          } else if (colIdx === 2 || colIdx === 3) {
            cell.alignment = { horizontal: "center", vertical: "middle" };
          } else if (colIdx === 5) {
            cell.alignment = { horizontal: "right", vertical: "middle" };
            cell.numFmt = "#,##0";
          }

          if (zebraFlag) {
            cell.fill = fillZebra;
          }
        });

        currentRowNum++;
        zebraFlag = !zebraFlag;
        idx++;
      }

      const endRow = currentRowNum - 1;

      // Subtotal row for store - Merge columns 1-5 (A to E)
      ws.mergeCells(currentRowNum, 1, currentRowNum, 5);
      const subtotalLabelCell = ws.getCell(`A${currentRowNum}`);
      subtotalLabelCell.value = `รวมย่อยร้าน - ${currentStore}`;
      subtotalLabelCell.font = fontSubtotal;
      subtotalLabelCell.alignment = { horizontal: "left", vertical: "middle" };

      // Apply style to columns 1-5 of this subtotal row
      for (let c = 1; c <= 5; c++) {
        const cell = ws.getCell(currentRowNum, c);
        cell.fill = fillSubtotal;
        cell.border = subtotalBorder;
        if (c > 1) {
          cell.font = fontSubtotal; // set font for merged area cells to ensure style correctness
        }
      }

      // Column 6 (F): Sum of quantity using formula =SUM(F{startRow}:F{endRow})
      const subCell = ws.getCell(`F${currentRowNum}`);
      subCell.value = { formula: `SUM(F${startRow}:F${endRow})` };
      subCell.font = fontSubtotal;
      subCell.alignment = { horizontal: "right", vertical: "middle" };
      subCell.numFmt = "#,##0";
      subCell.fill = fillSubtotal;
      subCell.border = subtotalBorder;

      storeSubtotalRows.push(currentRowNum);
      ws.getRow(currentRowNum).height = 20;
      currentRowNum++;
    }

    // Provincial Total Row - Merge columns 1-5 (A to E)
    ws.mergeCells(currentRowNum, 1, currentRowNum, 5);
    const provLabelCell = ws.getCell(`A${currentRowNum}`);
    provLabelCell.value = `รวมทั้งสิ้น จังหวัด${targetProv}`;
    provLabelCell.font = fontProvTotal;
    provLabelCell.alignment = { horizontal: "left", vertical: "middle" };

    // Fill style to columns 1-5 of this prov total row
    for (let c = 1; c <= 5; c++) {
      const cell = ws.getCell(currentRowNum, c);
      cell.fill = fillProvTotal;
      cell.border = subtotalBorder;
      if (c > 1) {
        cell.font = fontProvTotal;
      }
    }

    // Column 6 Provincial formula: =F10+F15+F22 (summing the individual store rows)
    const provFormula = storeSubtotalRows.length > 0
      ? storeSubtotalRows.map(r => `F${r}`).join("+")
      : "0";

    const provCell = ws.getCell(`F${currentRowNum}`);
    provCell.value = { formula: provFormula };
    provCell.font = fontProvTotal;
    provCell.alignment = { horizontal: "right", vertical: "middle" };
    provCell.numFmt = "#,##0";
    provCell.fill = fillProvTotal;
    provCell.border = subtotalBorder;

    provTotalRows.push(currentRowNum);
    ws.getRow(currentRowNum).height = 22;
    currentRowNum += 2; // skip a row (creates blank row, active row index jumps)
  }

  // Grand Total Row back down one (remove trailing spacing gap row increment)
  currentRowNum--;

  // Merge columns 1-5
  ws.mergeCells(currentRowNum, 1, currentRowNum, 5);
  const grandLabelCell = ws.getCell(`A${currentRowNum}`);
  grandLabelCell.value = `รวมทั้งสิ้นสุทธิ (${selectedProvinces.join(" & ")})`;
  grandLabelCell.font = fontTotal;
  grandLabelCell.alignment = { horizontal: "left", vertical: "middle" };

  for (let c = 1; c <= 5; c++) {
    const cell = ws.getCell(currentRowNum, c);
    cell.fill = fillTotal;
    cell.border = totalBorder;
    if (c > 1) {
      cell.font = fontTotal;
    }
  }

  const grandFormula = provTotalRows.length > 0
    ? provTotalRows.map(r => `F${r}`).join("+")
    : "0";

  const totalCell = ws.getCell(`F${currentRowNum}`);
  totalCell.value = { formula: grandFormula };
  totalCell.font = fontTotal;
  totalCell.alignment = { horizontal: "right", vertical: "middle" };
  totalCell.numFmt = "#,##0";
  totalCell.fill = fillTotal;
  totalCell.border = totalBorder;
  ws.getRow(currentRowNum).height = 26;

  // Set widths explicitly
  const widths = [14, 45, 16, 16, 45, 14];
  widths.forEach((w, idx) => {
    ws.getColumn(idx + 1).width = w;
  });

  // ============================================================
  // SHEET 2: "สรุปยอดตามทะเบียนรถ" (group by truck registration)
  // Mirrors the Python script: เน้นทะเบียนรถเป็นหลัก, รวมยอดต่อคันรถ
  // ============================================================
  const wsTruck = wb.addWorksheet("สรุปยอดตามทะเบียนรถ", {
    views: [{ showGridLines: true }]
  });

  // Sort: Truck -> Province -> Store -> Bill (records with no truck go last as "(ไม่ระบุทะเบียนรถ)")
  const truckData = [...filteredData].sort((a, b) => {
    const ta = a.truck.trim();
    const tb = b.truck.trim();
    if (ta !== tb) {
      if (!ta) return 1;   // empty truck sinks to bottom
      if (!tb) return -1;
      return ta.localeCompare(tb, 'th');
    }
    if (a.province !== b.province) return a.province.localeCompare(b.province, 'th');
    if (a.store !== b.store) return a.store.localeCompare(b.store, 'th');
    return a.bill.localeCompare(b.bill, 'th');
  });

  // Truck-specific styling (light blue subtotal #EAF2F8, navy underline)
  const fillTruckSub: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAF2F8" } };
  const truckSubBorder: Partial<ExcelJS.Borders> = {
    top: { style: "thin", color: { argb: "FFAAAAAA" } },
    bottom: { style: "thin", color: { argb: "FF1B365D" } }
  };

  // Title - row 1
  wsTruck.getCell("A1").value = `รายงานสรุปการจัดส่งสินค้า เน้นทะเบียนรถเป็นหลัก (${selectedProvinces.join(" & ")})`;
  wsTruck.getCell("A1").font = fontTitle;
  wsTruck.getRow(1).height = 30;

  // Header table - row 3 (truck first)
  const truckHeaders = ["ทะเบียนรถ", "จังหวัด", "ร้านค้า", "เลขที่บิล", "รหัสสินค้า", "สินค้า", "จำนวน(หีบ)"];
  const truckHeaderRow = wsTruck.getRow(3);
  truckHeaderRow.height = 25;
  truckHeaders.forEach((header, colIdx) => {
    const cell = truckHeaderRow.getCell(colIdx + 1);
    cell.value = header;
    cell.font = fontHeader;
    cell.fill = fillHeader;
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = thinBorder;
  });

  let truckRowNum = 4;
  const truckTotalRows: number[] = [];
  let tIdx = 0;
  const truckTotalLen = truckData.length;

  // Loop trucks sequentially
  while (tIdx < truckTotalLen) {
    const currentTruck = truckData[tIdx].truck.trim();
    const truckLabel = currentTruck || "(ไม่ระบุทะเบียนรถ)";
    const startRow = truckRowNum;
    let zebraFlag = false;

    while (tIdx < truckTotalLen && truckData[tIdx].truck.trim() === currentTruck) {
      const item = truckData[tIdx];
      const itemRow = wsTruck.getRow(truckRowNum);
      itemRow.height = 20;

      // Column order: truck, province, store, bill, pCode, pName, qty
      const rowValues = [truckLabel, item.province, item.store, item.bill, item.pCode, item.pName, item.qty];

      rowValues.forEach((val, colIdx) => {
        const cell = itemRow.getCell(colIdx + 1);
        cell.value = val;
        cell.font = fontBody;
        cell.border = thinBorder;

        // cols 1,2,3,6 left | 4,5 center | 7 right
        if (colIdx === 0 || colIdx === 1 || colIdx === 2 || colIdx === 5) {
          cell.alignment = { horizontal: "left", vertical: "middle" };
        } else if (colIdx === 3 || colIdx === 4) {
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else if (colIdx === 6) {
          cell.alignment = { horizontal: "right", vertical: "middle" };
          cell.numFmt = "#,##0";
        }

        if (zebraFlag) {
          cell.fill = fillZebra;
        }
      });

      truckRowNum++;
      zebraFlag = !zebraFlag;
      tIdx++;
    }

    const endRow = truckRowNum - 1;

    // Truck subtotal row - merge columns 1-6 (A to F)
    wsTruck.mergeCells(truckRowNum, 1, truckRowNum, 6);
    const subLabelCell = wsTruck.getCell(`A${truckRowNum}`);
    subLabelCell.value = `รวมยอดรถทะเบียน - ${truckLabel}`;
    subLabelCell.font = fontSubtotal;
    subLabelCell.alignment = { horizontal: "left", vertical: "middle" };

    for (let c = 1; c <= 6; c++) {
      const cell = wsTruck.getCell(truckRowNum, c);
      cell.fill = fillTruckSub;
      cell.border = truckSubBorder;
      if (c > 1) {
        cell.font = fontSubtotal;
      }
    }

    // Column 7 (G): sum of qty for this truck
    const subCell = wsTruck.getCell(`G${truckRowNum}`);
    subCell.value = { formula: `SUM(G${startRow}:G${endRow})` };
    subCell.font = fontSubtotal;
    subCell.alignment = { horizontal: "right", vertical: "middle" };
    subCell.numFmt = "#,##0";
    subCell.fill = fillTruckSub;
    subCell.border = truckSubBorder;

    truckTotalRows.push(truckRowNum);
    wsTruck.getRow(truckRowNum).height = 20;
    truckRowNum += 2; // blank spacer row between trucks
  }

  // Grand Total Row (sum of all truck subtotals)
  truckRowNum--;
  wsTruck.mergeCells(truckRowNum, 1, truckRowNum, 6);
  const truckGrandLabel = wsTruck.getCell(`A${truckRowNum}`);
  truckGrandLabel.value = `รวมทั้งสิ้นสุทธิ (${selectedProvinces.join(" & ")})`;
  truckGrandLabel.font = fontTotal;
  truckGrandLabel.alignment = { horizontal: "left", vertical: "middle" };

  for (let c = 1; c <= 6; c++) {
    const cell = wsTruck.getCell(truckRowNum, c);
    cell.fill = fillTotal;
    cell.border = totalBorder;
    if (c > 1) {
      cell.font = fontTotal;
    }
  }

  const truckGrandFormula = truckTotalRows.length > 0
    ? truckTotalRows.map(r => `G${r}`).join("+")
    : "0";

  const truckTotalCell = wsTruck.getCell(`G${truckRowNum}`);
  truckTotalCell.value = { formula: truckGrandFormula };
  truckTotalCell.font = fontTotal;
  truckTotalCell.alignment = { horizontal: "right", vertical: "middle" };
  truckTotalCell.numFmt = "#,##0";
  truckTotalCell.fill = fillTotal;
  truckTotalCell.border = totalBorder;
  wsTruck.getRow(truckRowNum).height = 26;

  // Column widths (truck first)
  const truckWidths = [18, 14, 45, 16, 16, 45, 14];
  truckWidths.forEach((w, idx) => {
    wsTruck.getColumn(idx + 1).width = w;
  });

  // Write and Save
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = outputFileName;
  // Anchor must be in the DOM for Chrome to honor the `download` attribute,
  // otherwise it falls back to the blob UUID as the filename (no .xlsx).
  anchor.style.display = "none";
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  // Delay revoke so the download isn't interrupted before it starts.
  setTimeout(() => window.URL.revokeObjectURL(url), 1000);
}
