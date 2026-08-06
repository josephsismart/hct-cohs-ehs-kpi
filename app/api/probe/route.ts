import { NextResponse } from 'next/server';
import { fetchSheet, fetchReport } from '@/lib/smartsheet';

export const dynamic = 'force-dynamic';

export async function GET(request: Request) {
  const token = process.env.SMARTSHEET_TOKEN;
  if (!token) return NextResponse.json({ error: 'no token' }, { status: 500 });
  const u = new URL(request.url);
  const action = u.searchParams.get('action') || 'list_reports';
  const id = u.searchParams.get('id') || '';

  if (action === 'list_reports') {
    const r = await fetch('https://api.smartsheet.com/2.0/reports?pageSize=100', {
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' }
    });
    const d = await r.json();
    return NextResponse.json((d.data || []).map((x: any) => ({ id: x.id, name: x.name })));
  }

  if (action === 'report' && id) {
    const rows = await fetchReport(id, token);
    return NextResponse.json({ id, count: rows.length, cols: rows.length > 0 ? Object.keys(rows[0]) : [], sample: rows.slice(0, 3) });
  }

  if (action === 'sheet' && id) {
    const rows = await fetchSheet(id, token);
    return NextResponse.json({ id, count: rows.length, cols: rows.length > 0 ? Object.keys(rows[0]) : [], sample: rows.slice(0, 3) });
  }

  return NextResponse.json({ error: 'use ?action=list_reports or ?action=report&id=X or ?action=sheet&id=X' });
}
