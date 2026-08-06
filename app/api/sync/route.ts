import { NextResponse } from 'next/server';
import { SYNC_SOURCES, fetchSheet, fetchReport, processSource, KpiRow } from '@/lib/smartsheet';

export const dynamic = 'force-dynamic';
export const maxDuration = 60;

export async function GET(request: Request) {
  const token = process.env.SMARTSHEET_TOKEN;
  if (!token) return NextResponse.json({ error: 'SMARTSHEET_TOKEN not set' }, { status: 500 });

  // Debug: list all reports or fetch a specific report's columns
  const url = new URL(request.url);
  const listReports = url.searchParams.get('list_reports');
  const probeReport = url.searchParams.get('probe_report');
  const probeSheet = url.searchParams.get('probe_sheet');
  if (listReports) {
    const res = await fetch('https://api.smartsheet.com/2.0/reports?pageSize=100', { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' } });
    const data = await res.json();
    return NextResponse.json({ reports: (data.data || []).map((r: any) => ({ id: r.id, name: r.name, accessLevel: r.accessLevel })) });
  }
  if (probeReport) {
    const rows = await fetchReport(probeReport, token);
    return NextResponse.json({ reportId: probeReport, rowCount: rows.length, columns: rows.length > 0 ? Object.keys(rows[0]) : [], sample: rows.slice(0, 3) });
  }
  if (probeSheet) {
    const rows = await fetchSheet(probeSheet, token);
    return NextResponse.json({ sheetId: probeSheet, rowCount: rows.length, columns: rows.length > 0 ? Object.keys(rows[0]) : [], sample: rows.slice(0, 3) });
  }

  const results: Record<string, { rows: KpiRow[]; error?: string }> = {};
  const errors: string[] = [];

  const debug: Record<string, any> = {};
  await Promise.allSettled(
    SYNC_SOURCES.map(async (src) => {
      try {
        const raw = src.sheetId
          ? await fetchSheet(src.sheetId, token)
          : await fetchReport(src.reportId!, token);
        const processed = processSource(src, raw);
        // Debug: always capture raw info for column verification
        debug[src.key] = { rawCount: raw.length, processedCount: processed.length, columns: raw.length > 0 ? Object.keys(raw[0]) : [], config: { campusCol: src.campusCol, monthCol: src.monthCol, plannedCol: src.plannedCol, actualCol: src.actualCol, valueCol: src.valueCol } };
        if (raw.length > 0 && (processed.length === 0 || processed.length !== raw.length)) {
          debug[src.key].sample = raw[0];
          // Check for column mismatch
          const cols = Object.keys(raw[0]);
          const missing: string[] = [];
          if (src.campusCol && !cols.includes(src.campusCol)) missing.push(`campusCol: '${src.campusCol}'`);
          if (src.monthCol && !cols.includes(src.monthCol)) missing.push(`monthCol: '${src.monthCol}'`);
          if (src.plannedCol && !cols.includes(src.plannedCol)) missing.push(`plannedCol: '${src.plannedCol}'`);
          if (src.actualCol && !cols.includes(src.actualCol)) missing.push(`actualCol: '${src.actualCol}'`);
          if (src.valueCol && !cols.includes(src.valueCol)) missing.push(`valueCol: '${src.valueCol}'`);
          if (missing.length > 0) debug[src.key].missingColumns = missing;
        }
        results[src.key] = { rows: processed };
      } catch (e: any) {
        results[src.key] = { rows: [], error: e.message };
        errors.push(`${src.key}: ${e.message}`);
      }
    })
  );

  // Extract unique campuses and months — skip isolated sources (pie charts, committee, mgmt review)
  const campusSet = new Set<string>();
  const monthSet = new Set<string>();
  const isolatedKeys = new Set(SYNC_SOURCES.filter(s => s.isolateFromCampusSet).map(s => s.key));
  Object.entries(results).forEach(([key, { rows }]) => {
    rows.forEach(r => {
      if (r.campus && !isolatedKeys.has(key)) campusSet.add(r.campus);
      if (r.month) monthSet.add(r.month);
    });
  });

  // Fetch waste segregation raw data for the waste table
  let wasteData: Record<string, any>[] = [];
  try {
    const wasteRaw = await fetchSheet('6383128615538564', token);
    const WASTE_COLS = ['Total Waste (General/Recyclable)','General Waste','Food Waste','Paper Waste','Paper Cup/Carton','PET Bottle','Single Use Plastic','Aluminum','Tissue','Scrap Metal','E-Waste','Hazardous','Medical'];
    wasteData = wasteRaw.map(r => {
      const campus = String(r['Campus Code'] || '').trim();
      if (!campus) return null;
      const month = (() => { const v = r['Reporting Month']; if (!v) return null; const s = String(v).trim(); const MNAMES = ['January','February','March','April','May','June','July','August','September','October','November','December']; const idx = MNAMES.findIndex(m => m.toLowerCase() === s.toLowerCase()); if (idx >= 0) return MNAMES[idx]; const abbr = s.substring(0,3).toLowerCase(); const MABBR: Record<string,number> = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11}; if (MABBR[abbr] !== undefined) return MNAMES[MABBR[abbr]]; return null; })();
      const row: Record<string, any> = { campus, month };
      WASTE_COLS.forEach(col => { row[col] = parseFloat(r[col]) || 0; });
      // Alias new column name for backward compatibility
      if (row['Total Waste (General/Recyclable)'] !== undefined) row['Total Waste'] = row['Total Waste (General/Recyclable)'];
      return row;
    }).filter(Boolean) as Record<string, any>[];
  } catch (e: any) { errors.push(`wasteData: ${e.message}`); }

  return NextResponse.json({
    syncedAt: new Date().toISOString(),
    sources: results,
    campuses: [...campusSet].sort(),
    months: [...monthSet],
    errors,
    debug,
    wasteData,
  });
        }import { NextResponse } from 'next/server';
import { SYNC_SOURCES, fetchSheet, fetchReport, processSource, KpiRow } from '@/lib/smartsheet';

export const dynamic = 'force-dynamic';
export const maxDuration = 60;

export async function GET() {
  const token = process.env.SMARTSHEET_TOKEN;
  if (!token) return NextResponse.json({ error: 'SMARTSHEET_TOKEN not set' }, { status: 500 });

  const results: Record<string, { rows: KpiRow[]; error?: string }> = {};
  const errors: string[] = [];

  const debug: Record<string, any> = {};
  await Promise.allSettled(
    SYNC_SOURCES.map(async (src) => {
      try {
        const raw = src.sheetId
          ? await fetchSheet(src.sheetId, token)
          : await fetchReport(src.reportId!, token);
        const processed = processSource(src, raw);
        // Debug: always capture raw info for column verification
        debug[src.key] = { rawCount: raw.length, processedCount: processed.length, columns: raw.length > 0 ? Object.keys(raw[0]) : [], config: { campusCol: src.campusCol, monthCol: src.monthCol, plannedCol: src.plannedCol, actualCol: src.actualCol, valueCol: src.valueCol } };
        if (raw.length > 0 && (processed.length === 0 || processed.length !== raw.length)) {
          debug[src.key].sample = raw[0];
          // Check for column mismatch
          const cols = Object.keys(raw[0]);
          const missing: string[] = [];
          if (src.campusCol && !cols.includes(src.campusCol)) missing.push(`campusCol: '${src.campusCol}'`);
          if (src.monthCol && !cols.includes(src.monthCol)) missing.push(`monthCol: '${src.monthCol}'`);
          if (src.plannedCol && !cols.includes(src.plannedCol)) missing.push(`plannedCol: '${src.plannedCol}'`);
          if (src.actualCol && !cols.includes(src.actualCol)) missing.push(`actualCol: '${src.actualCol}'`);
          if (src.valueCol && !cols.includes(src.valueCol)) missing.push(`valueCol: '${src.valueCol}'`);
          if (missing.length > 0) debug[src.key].missingColumns = missing;
        }
        results[src.key] = { rows: processed };
      } catch (e: any) {
        results[src.key] = { rows: [], error: e.message };
        errors.push(`${src.key}: ${e.message}`);
      }
    })
  );

  // Extract unique campuses and months — skip isolated sources (pie charts, committee, mgmt review)
  const campusSet = new Set<string>();
  const monthSet = new Set<string>();
  const isolatedKeys = new Set(SYNC_SOURCES.filter(s => s.isolateFromCampusSet).map(s => s.key));
  Object.entries(results).forEach(([key, { rows }]) => {
    rows.forEach(r => {
      if (r.campus && !isolatedKeys.has(key)) campusSet.add(r.campus);
      if (r.month) monthSet.add(r.month);
    });
  });

  // Fetch waste segregation raw data for the waste table
  let wasteData: Record<string, any>[] = [];
  try {
    const wasteRaw = await fetchSheet('6383128615538564', token);
    const WASTE_COLS = ['Total Waste (General/Recyclable)','General Waste','Food Waste','Paper Waste','Paper Cup/Carton','PET Bottle','Single Use Plastic','Aluminum','Tissue','Scrap Metal','E-Waste','Hazardous','Medical'];
    wasteData = wasteRaw.map(r => {
      const campus = String(r['Campus Code'] || '').trim();
      if (!campus) return null;
      const month = (() => { const v = r['Reporting Month']; if (!v) return null; const s = String(v).trim(); const MNAMES = ['January','February','March','April','May','June','July','August','September','October','November','December']; const idx = MNAMES.findIndex(m => m.toLowerCase() === s.toLowerCase()); if (idx >= 0) return MNAMES[idx]; const abbr = s.substring(0,3).toLowerCase(); const MABBR: Record<string,number> = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11}; if (MABBR[abbr] !== undefined) return MNAMES[MABBR[abbr]]; return null; })();
      const row: Record<string, any> = { campus, month };
      WASTE_COLS.forEach(col => { row[col] = parseFloat(r[col]) || 0; });
      // Alias new column name for backward compatibility
      if (row['Total Waste (General/Recyclable)'] !== undefined) row['Total Waste'] = row['Total Waste (General/Recyclable)'];
      return row;
    }).filter(Boolean) as Record<string, any>[];
  } catch (e: any) { errors.push(`wasteData: ${e.message}`); }

  return NextResponse.json({
    syncedAt: new Date().toISOString(),
    sources: results,
    campuses: [...campusSet].sort(),
    months: [...monthSet],
    errors,
    debug,
    wasteData,
  });
}
