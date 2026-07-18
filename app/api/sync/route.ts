import { NextResponse } from 'next/server';
import { SYNC_SOURCES, fetchSheet, fetchReport, processSource, KpiRow } from '@/lib/smartsheet';

export const dynamic = 'force-dynamic';
export const maxDuration = 60;

export async function GET() {
  const token = process.env.SMARTSHEET_TOKEN;
  if (!token) return NextResponse.json({ error: 'SMARTSHEET_TOKEN not set' }, { status: 500 });

  const results: Record<string, { rows: KpiRow[]; error?: string }> = {};
  const errors: string[] = [];

  await Promise.allSettled(
    SYNC_SOURCES.map(async (src) => {
      try {
        const raw = src.sheetId
          ? await fetchSheet(src.sheetId, token)
          : await fetchReport(src.reportId!, token);
        results[src.key] = { rows: processSource(src, raw) };
      } catch (e: any) {
        results[src.key] = { rows: [], error: e.message };
        errors.push(src.key + ': ' + e.message);
      }
    })
  );

  const campusSet = new Set<string>();
  const monthSet = new Set<string>();
  Object.values(results).forEach(({ rows }) => {
    rows.forEach(r => {
      if (r.campus) campusSet.add(r.campus);
      if (r.month) monthSet.add(r.month);
    });
  });

  return NextResponse.json({
    syncedAt: new Date().toISOString(),
    sources: results,
    campuses: [...campusSet].sort(),
    months: [...monthSet],
    errors,
  });
}
