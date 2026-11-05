/**
 * G080 (สหพัฒน์) — ดึงไฟล์ "ข้อมูลการเบิก" จากระบบ trg080.sslog.co.th โดยตรง
 *
 * ลอกกลไกจากโปรเจกต์ logistics-delivery-kpi-management (backend/src/g080/fetch.ts):
 *   trg080.sslog.co.th เป็นแค่ frameset → เซิร์ฟเวอร์จริงคือ http://83.118.71.78:8101/trg080 (HTTP)
 *   login.php (username/password) → คุกกี้ PHPSESSID → issue.php?getid=<user> + report=รายงาน → ได้ .xls (BIFF)
 *   ช่อง date_from = "วันเบิก" (ไม่ใช่วันที่หน้าบิล — ไฟล์วันเบิก 4 พ.ย. มีบิลลงวันที่ 30 ต.ค.–4 พ.ย.)
 *
 * ไฟล์นี้รันฝั่ง Node เท่านั้น (เรียกจาก middleware ใน vite.config.ts) — เบราว์เซอร์เรียกตรงไม่ได้
 * เพราะเป็น HTTP + ข้ามโดเมน และรหัสผ่านต้องไม่หลุดไปถึงหน้าเว็บ
 */
import * as XLSX from 'xlsx';

/** ไฟล์ที่มีแต่หัวตาราง (วันที่ไม่มีงานเบิก) ≈ 5.6KB */
const EMPTY_XLS_MAX = 6000;
/** คอลัมน์รหัสสินค้า (J) — ใช้แยกแถวข้อมูลจริงออกจากหัวรายงาน/แถว "ยอดรวม" */
const PCODE_COL = 9;
/** คอลัมน์จำนวน(หีบ) (L) — ใช้รวมยอดต่อวันไว้ตรวจทาน */
const QTY_COL = 11;

// ข้อความไทยในไฟล์เป็น TIS-620 — อ่านเป็นไบต์ดิบ (codepage 1252) แล้วถอดเป็น windows-874 เอง
const td = new TextDecoder('windows-874');
// ถ้า SheetJS ถอด cp1252 จริง ไบต์ 0x80–0x9F จะกลายเป็น U+20xx (… “ ” – —) → แปลงกลับเป็นไบต์เดิมก่อน
// (ตอนนี้ไลบรารีคืนไบต์ตรงตัว ตรวจไฟล์จริง 2 วันพบ 0 ตัว แต่กันไว้เผื่ออัปเดตไลบรารี — Gemini P1)
const CP1252_HIGH: Record<number, number> = {
  0x20ac: 0x80, 0x201a: 0x82, 0x0192: 0x83, 0x201e: 0x84, 0x2026: 0x85, 0x2020: 0x86, 0x2021: 0x87,
  0x02c6: 0x88, 0x2030: 0x89, 0x0160: 0x8a, 0x2039: 0x8b, 0x0152: 0x8c, 0x017d: 0x8e, 0x2018: 0x91,
  0x2019: 0x92, 0x201c: 0x93, 0x201d: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97, 0x02dc: 0x98,
  0x2122: 0x99, 0x0161: 0x9a, 0x203a: 0x9b, 0x0153: 0x9c, 0x017e: 0x9e, 0x0178: 0x9f,
};
const toByte = (ch: string) => { const k = ch.charCodeAt(0); return k <= 0xff ? k : (CP1252_HIGH[k] ?? 0x3f); };
const cp874 = (v: unknown): unknown =>
  typeof v === 'string' ? td.decode(Uint8Array.from(v, toByte)) : v;

export type G080Config = { baseUrl: string; user: string; pass: string };

export type G080DaySummary = { date: string; rows: number; qty: number; empty: boolean };

export type G080Result = {
  header: string[];
  rows: unknown[][];        // แถวข้อมูล (ถอดไทยแล้ว) ตำแหน่งคอลัมน์เดิมของไฟล์ A..R
  days: G080DaySummary[];
  totalRows: number;
  totalQty: number;
  fromIso: string;
  toIso: string;
};

const isIso = (s: string) => /^\d{4}-\d{2}-\d{2}$/.test(s);
/** yyyy-mm-dd → dd/MM/yyyy (G080 ใช้ ค.ศ.) */
const toDdmmyyyy = (iso: string) => { const [y, m, d] = iso.split('-'); return `${d}/${m}/${y}`; };

function dateRange(fromIso: string, toIso: string): string[] {
  if (!isIso(fromIso) || !isIso(toIso)) throw new Error('รูปแบบวันที่ไม่ถูกต้อง (ต้องเป็น YYYY-MM-DD)');
  const d = new Date(fromIso + 'T00:00:00Z');
  const end = new Date(toIso + 'T00:00:00Z');
  if (isNaN(d.getTime()) || isNaN(end.getTime())) throw new Error('วันที่ไม่ถูกต้อง');
  if (d > end) throw new Error('วันที่เริ่มต้องไม่เกินวันที่สิ้นสุด');
  // JS เลื่อนวันที่ไม่มีจริง (เช่น 2026-02-31 → 3 มี.ค.) เงียบๆ — ต้องตรงกับที่ส่งมาเป๊ะ (Gemini P3)
  if (d.toISOString().slice(0, 10) !== fromIso || end.toISOString().slice(0, 10) !== toIso) {
    throw new Error('วันที่ไม่มีอยู่จริงในปฏิทิน');
  }
  const out: string[] = [];
  while (d <= end) { out.push(d.toISOString().slice(0, 10)); d.setUTCDate(d.getUTCDate() + 1); }
  return out;
}

async function login(cfg: G080Config): Promise<string> {
  const r = await fetch(`${cfg.baseUrl}/login.php`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ username: cfg.user, password: cfg.pass }),
    redirect: 'manual',
    signal: AbortSignal.timeout(10000),   // login 10 + ดึง 1 วัน 45 < 60 วิ (Vercel maxDuration)
  });
  const m = (r.headers.get('set-cookie') || '').match(/PHPSESSID=[^;]+/);
  if (!m) throw new Error('เข้าระบบ G080 ไม่สำเร็จ (ชื่อผู้ใช้/รหัสผ่านอาจผิด หรือระบบเปลี่ยน)');
  return m[0];
}

/** ดึง .xls ของวันเบิก 1 วัน — คืน null เมื่อวันนั้นไม่มีงานจริง */
async function fetchDay(cfg: G080Config, cookie: string, iso: string): Promise<Buffer | null> {
  const r = await fetch(`${cfg.baseUrl}/issue.php?getid=${encodeURIComponent(cfg.user)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', Cookie: cookie },
    body: new URLSearchParams({ date_from: toDdmmyyyy(iso), runn: '', inv: '', report: '  รายงาน  ' }),
    // 45 วิ — ต้องน้อยกว่า maxDuration 60 วิ ของ Vercel (vercel.json) จะได้ตอบ error เป็นข้อความไทย ไม่ใช่ 504 เปล่าๆ
    signal: AbortSignal.timeout(45000),
  });
  if (!r.ok) throw new Error(`G080 ตอบ HTTP ${r.status}`);
  const buf = Buffer.from(await r.arrayBuffer());
  // OLE2/BIFF magic D0 CF 11 E0 — ไม่ใช่ = ได้หน้า HTML/login กลับมา (session หมด)
  const isXls = buf.length >= 8 && buf[0] === 0xd0 && buf[1] === 0xcf && buf[2] === 0x11 && buf[3] === 0xe0;
  if (!isXls) throw new Error('G080 ไม่คืนไฟล์ Excel (session อาจหมดอายุ)');
  return buf.length <= EMPTY_XLS_MAX ? null : buf;
}

/**
 * ดึงงานเบิกช่วงวัน (รวมทุกวันเป็นชุดเดียว)
 * ⚠️ วันใดดึงพลาด = ยกเลิกทั้งช่วง — ห้ามถือเป็น "วันว่าง" ไม่งั้นยอดขาดหายเงียบๆ (บทเรียน Codex P1 จากโปรเจกต์ KPI)
 */
export async function fetchG080Range(cfg: G080Config, fromIso: string, toIso: string, maxDays = 31): Promise<G080Result> {
  if (!cfg.user || !cfg.pass) {
    throw new Error('ยังไม่ได้ตั้งค่า G080_USER / G080_PASS (เครื่องนี้: ไฟล์ .env.local · ออนไลน์: Vercel Environment Variables)');
  }
  const days = dateRange(fromIso, toIso);
  if (days.length > maxDays) throw new Error(`ช่วงวันยาวเกินไป (${days.length} วัน) — ดึงได้ไม่เกิน ${maxDays} วันต่อครั้ง`);

  const cookie = await login(cfg);
  let header: string[] | null = null;
  const rows: unknown[][] = [];
  const summary: G080DaySummary[] = [];

  for (const iso of days) {
    let buf: Buffer | null;
    try {
      buf = await fetchDay(cfg, cookie, iso);
    } catch (e: any) {
      throw new Error(`ดึงวันเบิก ${iso} ไม่สำเร็จ (${e?.message || e}) — ยกเลิกทั้งช่วงกันข้อมูลขาด ลองใหม่อีกครั้ง`);
    }
    if (!buf) { summary.push({ date: iso, rows: 0, qty: 0, empty: true }); continue; }

    const wb = XLSX.read(buf, { type: 'buffer', codepage: 1252 });
    const raw: unknown[][] = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
    const hr = raw.slice(0, 6).findIndex(r => r.some(c => String(cp874(String(c))).includes('รหัสสินค้า')));
    // มีข้อมูลแต่หาหัวตารางไม่เจอ = รูปแบบไฟล์เปลี่ยน → หยุด ไม่ข้ามเงียบๆ
    if (hr < 0) throw new Error(`ไฟล์วันเบิก ${iso} หาหัวตาราง "รหัสสินค้า" ไม่เจอ — ระบบ G080 อาจเปลี่ยนรูปแบบ`);
    if (!header) header = raw[hr].map(c => String(cp874(c)));

    let n = 0, qty = 0;
    let fileTotal: number | null = null;   // ตัวเลขจากแถว "ยอดรวม" ท้ายไฟล์ เช่น "12565 หีบ"
    const dayRows: unknown[][] = [];
    for (let i = hr + 1; i < raw.length; i++) {
      if (String(cp874(String(raw[i][PCODE_COL + 1] ?? ''))).trim() === 'ยอดรวม') {
        const m = String(cp874(String(raw[i][QTY_COL] ?? ''))).replace(/,/g, '').match(/-?\d+(\.\d+)?/);
        if (m) fileTotal = Number(m[0]);
        continue;
      }
      // แถวข้อมูลจริงต้องมีรหัสสินค้าเป็นตัวเลข — ตัดแถวว่าง/หัวรายงาน
      if (!/^\d{3,6}$/.test(String(raw[i][PCODE_COL] ?? '').trim())) continue;
      dayRows.push(raw[i].map(cp874));
      n++;
      qty += Number(raw[i][QTY_COL]) || 0;
    }
    // ไฟล์ใหญ่กว่าไฟล์เปล่าแต่อ่านแถวไม่ได้เลย = รูปแบบเปลี่ยน → หยุด ห้ามนับเป็น 0 เงียบๆ (Gemini P2)
    if (n === 0) throw new Error(`ไฟล์วันเบิก ${iso} มีข้อมูลแต่อ่านแถวสินค้าไม่ได้ — ระบบ G080 อาจเปลี่ยนรูปแบบ`);
    // ยอดหีบที่อ่านได้ต้องเท่ากับ "ยอดรวม" ที่ G080 พิมพ์ท้ายไฟล์ — ไม่เท่า = มีแถวหลุด
    if (fileTotal !== null && Math.abs(fileTotal - qty) > 0.001) {
      throw new Error(`วันเบิก ${iso} อ่านได้ ${qty} หีบ แต่ G080 แจ้งยอดรวม ${fileTotal} หีบ — ข้อมูลไม่ครบ ยกเลิกทั้งช่วง`);
    }
    rows.push(...dayRows);
    summary.push({ date: iso, rows: n, qty, empty: false });
  }

  return {
    header: header || [],
    rows,
    days: summary,
    totalRows: rows.length,
    totalQty: summary.reduce((a, d) => a + d.qty, 0),
    fromIso,
    toIso,
  };
}
