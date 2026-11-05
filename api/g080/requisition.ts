/**
 * Vercel Serverless Function — GET /api/g080/requisition?from=YYYY-MM-DD&to=YYYY-MM-DD
 * ฉบับออนไลน์ของตัวกลางใน vite.config.ts (ให้ทีมงานเครื่องอื่นใช้ปุ่ม "ดึงจาก G080" ได้)
 *
 * ค่าที่ต้องตั้งใน Vercel → Settings → Environment Variables:
 *   G080_USER / G080_PASS  = บัญชี G080 (ห้ามใส่ในโค้ด)
 *   APP_ACCESS_CODE        = รหัสทีม — เว็บนี้เปิดสาธารณะ ไม่มีรหัสนี้ใครก็ดึงข้อมูลลูกค้าได้
 * ⚠️ ดึงได้ครั้งละไม่เกิน 3 วัน (คำตอบ Vercel จำกัด ~4.5MB · 1 วัน ≈ 0.3–0.6MB) — หน้าเว็บวนดึงทีละวันเอง
 */
import type { IncomingMessage, ServerResponse } from 'http';
import { timingSafeEqual } from 'crypto';
import { fetchG080Range } from '../../server/g080';

const send = (res: ServerResponse, status: number, body: unknown) => {
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');   // ข้อมูลลูกค้า ห้าม CDN เก็บ
  res.end(JSON.stringify(body));
};

/** เทียบรหัสแบบเวลาคงที่ กันเดารหัสจากเวลาตอบ */
const sameCode = (a: string, b: string) => {
  const x = Buffer.from(a), y = Buffer.from(b);
  return x.length === y.length && timingSafeEqual(x, y);
};

export default async function handler(req: IncomingMessage, res: ServerResponse) {
  if (req.method !== 'GET') return send(res, 405, { error: 'Method not allowed' });

  const code = process.env.APP_ACCESS_CODE || '';
  if (!code) return send(res, 503, { error: 'ผู้ดูแลยังไม่ได้ตั้งรหัสทีม (APP_ACCESS_CODE) ใน Vercel' });
  const given = String(req.headers['x-access-code'] || '');
  if (!given || !sameCode(given, code)) return send(res, 401, { error: 'รหัสทีมไม่ถูกต้อง', needCode: true });

  const url = new URL(req.url || '/', 'http://localhost');
  const from = url.searchParams.get('from') || '';
  const to = url.searchParams.get('to') || from;
  try {
    const cfg = {
      baseUrl: (process.env.G080_BASE_URL || 'http://83.118.71.78:8101/trg080').replace(/\/+$/, ''),
      user: process.env.G080_USER || '',
      pass: process.env.G080_PASS || '',
    };
    send(res, 200, await fetchG080Range(cfg, from, to, 3));
  } catch (e: any) {
    send(res, 502, { error: String(e?.message || e) });
  }
}
