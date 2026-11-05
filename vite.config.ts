import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import type { IncomingMessage, ServerResponse } from 'http';
import {defineConfig, loadEnv, type Plugin} from 'vite';
import { fetchG080Range } from './server/g080';

/**
 * GET /api/g080/requisition?from=YYYY-MM-DD&to=YYYY-MM-DD
 * ตัวกลางดึงงานเบิกจาก G080 (ทำงานเฉพาะตอน `npm run dev` / `npm run preview` บนเครื่องนี้)
 * รหัส G080 อ่านจาก .env.local (G080_USER / G080_PASS) — ไม่ส่งไปหน้าเว็บ
 */
function g080Api(env: Record<string, string>): Plugin {
  const cfg = {
    baseUrl: (env.G080_BASE_URL || 'http://83.118.71.78:8101/trg080').replace(/\/+$/, ''),
    user: env.G080_USER || '',
    pass: env.G080_PASS || '',
  };
  const send = (res: ServerResponse, status: number, body: unknown) => {
    res.statusCode = status;
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.end(JSON.stringify(body));
  };
  const handler = async (req: IncomingMessage, res: ServerResponse, next: () => void) => {
    const url = new URL(req.url || '/', 'http://localhost');
    if (url.pathname !== '/api/g080/requisition') return next();
    // dev server เปิด --host=0.0.0.0 → กันเครื่องอื่นในวงแลนดึงข้อมูลลูกค้าผ่านเครื่องนี้
    const ip = req.socket.remoteAddress || '';
    if (!['127.0.0.1', '::1', '::ffff:127.0.0.1'].includes(ip)) {
      return send(res, 403, { error: 'ดึงข้อมูล G080 ได้เฉพาะจากเครื่องที่รันโปรแกรมเท่านั้น' });
    }
    if (req.method !== 'GET') return send(res, 405, { error: 'Method not allowed' });
    const from = url.searchParams.get('from') || '';
    const to = url.searchParams.get('to') || from;
    try {
      send(res, 200, await fetchG080Range(cfg, from, to));
    } catch (e: any) {
      send(res, 502, { error: String(e?.message || e) });
    }
  };
  return {
    name: 'g080-api',
    configureServer(server) { server.middlewares.use(handler); },
    configurePreviewServer(server) { server.middlewares.use(handler); },
  };
}

export default defineConfig(({ mode }) => {
  // prefix '' = อ่านทุกคีย์ใน .env* (ใช้ฝั่ง Node เท่านั้น ไม่ฝังลง bundle)
  const env = loadEnv(mode, process.cwd(), '');
  return {
    plugins: [react(), tailwindcss(), g080Api(env)],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modify — file watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
      // Disable file watching when DISABLE_HMR is true to save CPU during agent edits.
      watch: process.env.DISABLE_HMR === 'true' ? null : {},
    },
  };
});
