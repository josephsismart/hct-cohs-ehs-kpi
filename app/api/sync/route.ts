import { NextResponse } from 'next/server';
import { SYNC_SOURCES, fetchSheet, fetchReport, processSource, normalizeMonth, KpiRow } from '@/lib/smartsheet';

export const dynamic = 'force-dynamic';
export const maxDuration = 60;

const WASTE_COLS = ['Total Waste','General Waste','Food Waste','Paper Waste','Paper Cup/Carton','PET Bottle','Single Use Plastic'];

export async function GET() {
  const token = process.env.SMARTSHEET_TOKEN;
  if (!token) return NextResponse.json({ error: 'SMARTSHEET_TOKEN not set' }, { status: 500 });

  const results: Record<string, { rows: KpiRow[]; error?: string }> = {};
  const errors: string[] = [];
  let wasteData: Record<string, any>[] = [];

  await Promise.allSettled(
    SYNC_SOURCES.map(async (src) => {
      try {
        const raw = src.sheetId
          ? await fetchSheet(src.sheetId, token)
          : await fetchReport(src.reportId!, token);
        const processed = processSource(src, raw);
        results[src.key] = { rows: processed };
        // For waste segregation, also pass raw columnar data for the table
        if (src.key === 'v2_waste_segregation') {
          wasteData = raw.map(r => {
            const campus = String(r[src.campusCol] || '').trim();
            const month = normalizeMonth(r[src.monthCol || '']);
            const row: Record<string, any> = { campus, month };
            WASTE_COLS.forEach(c => { row[c] = parseFloat(r[c]) || 0; });
            return row;
          }).filter(r => r.campus);
        }
      } catch (e: any) {
        results[src.key] = { rows: [], error: e.message };
        errors.push(`${src.key}: ${e.message}`);
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
    wasteData,
  });
}
