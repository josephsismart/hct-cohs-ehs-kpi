// Smartsheet API client — mirrors SyncService.gs SYNC_SOURCES config
// UPDATED: synced with GAS SyncService.gs.FIXED.txt (2026-07-19)

export interface SyncSource {
  key: string;
  reportId?: string;
  sheetId?: string;
  tab: string;
  campusCol: string;
  monthCol?: string;
  plannedCol?: string;
  actualCol?: string;
  valueCol?: string;
  hasMonth: boolean;h
  isolateFromCampusSet?: boolean;
  yesNoCount?: boolean;
}

export const SYNC_SOURCES: SyncSource[] = [
  // Original 7 KPIs — matched to GAS SyncService.gs
  { key: 'drills', sheetId: '5053158949605252', tab: 'raw_drills', campusCol: 'Campus Code', monthCol: 'Reporting Month', plannedCol: 'Planned Drill? (Yes/No)', actualCol: 'Are there any submission?', hasMonth: true, yesNoCount: true },
  { key: 'ehs', sheetId: '4947401822392196', tab: 'raw_ehs', monthCol: 'Primary', campusCol: 'Campus Code', plannedCol: 'No. of EHS Inspections Planned', actualCol: 'No. of EHS Inspections Completed', hasMonth: true },
  { key: 'findings', sheetId: '4947401822392196', tab: 'raw_findings', monthCol: 'Primary', campusCol: 'Campus Code', plannedCol: 'No. of Findings in Reporting Month', actualCol: 'No. of Findings Closed', hasMonth: true },
  { key: 'notification', reportId: '1199821531598724', tab: 'raw_notification', monthCol: 'Reporting Month', campusCol: 'Campus Code', plannedCol: 'Total Incident', actualCol: 'Notification Submitted on Time', hasMonth: true },
  { key: 'risk', reportId: '8565044722749316', tab: 'raw_risk', campusCol: 'Campus Code', monthCol: 'Reporting Month', plannedCol: 'Total Assessments Register', actualCol: 'RA Validated and Signed Off', hasMonth: true },
  { key: 'training', sheetId: '8549734774951812', tab: 'raw_training', campusCol: 'Campus Code', valueCol: 'Total Hours', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'incidents', sheetId: '7165378768621444', tab: 'raw_incidents', campusCol: 'Campus Code', monthCol: 'Date Reported', valueCol: 'Total Incident', hasMonth: true },

  // Pie chart — Incidents by Type
  { key: 'v2_incident_types', reportId: '20949779828612', tab: 'raw_v2_incident_types', campusCol: 'Incident Type', valueCol: 'Total Incident', monthCol: 'Reporting Month', hasMonth: true, isolateFromCampusSet: true },

  // V2 KPIs — matched to GAS SyncService.gs
  { key: 'v2_hs_committee', sheetId: '435993944477572', tab: 'raw_v2_hs_committee', campusCol: 'Committee', plannedCol: 'Meeting Planned', actualCol: 'Meeting Conducted', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_findings_on_time', sheetId: '4947401822392196', tab: 'raw_v2_findings_on_time', campusCol: 'Campus Code', plannedCol: 'No. of Findings in Reporting Month', actualCol: 'No. of Findings Due', monthCol: 'Primary', hasMonth: true },
  { key: 'v2_risk_closed', sheetId: '7323092115214212', tab: 'raw_v2_risk_closed', campusCol: 'Campus', plannedCol: 'Total Risk Assessments Registered', actualCol: 'Risk Assessment Closed', monthCol: 'Primary', hasMonth: true },
  { key: 'v2_risk_validated', sheetId: '7323092115214212', tab: 'raw_v2_risk_validated', campusCol: 'Campus', plannedCol: 'Total Assessments Register', actualCol: 'RA Validated and Signed Off', monthCol: 'Primary', hasMonth: true },
  { key: 'v2_ehs_inspection', sheetId: '4947401822392196', tab: 'raw_v2_ehs_inspection', campusCol: 'Campus Code', plannedCol: 'No. of EHS Inspections Planned', actualCol: 'No. of EHS Inspections Completed', monthCol: 'Primary', hasMonth: true },
  { key: 'v2_hs_kpi_report', reportId: '4811266391494532', tab: 'raw_v2_hs_kpi_report', campusCol: 'Campuses', valueCol: 'Submitted', monthCol: 'Primary', hasMonth: true },
  { key: 'v2_external_compliance', sheetId: '4198632256393092', tab: 'raw_v2_external_compliance', campusCol: 'Campus Code', plannedCol: 'Applicable Compliance', actualCol: 'Actual Compliance', monthCol: 'Primary', hasMonth: true },
  { key: 'v2_safe_working', sheetId: '1693592581001092', tab: 'raw_v2_safe_working', campusCol: 'Campus', plannedCol: 'No. of SOPs Verified', actualCol: 'No. of SOPs Implemented', monthCol: 'Primary', hasMonth: true },
  { key: 'v2_permit_to_work', sheetId: '5899016251330436', tab: 'raw_v2_permit_to_work', campusCol: 'Campus Code', plannedCol: 'No. of PTWs Issued', actualCol: 'Total Work Registered', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_hazard_id', sheetId: '7323092115214212', tab: 'raw_v2_hazard_id', campusCol: 'Campus Code', plannedCol: 'Total Controls Identified', actualCol: 'Implemented Controls', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_onsite_induction', sheetId: '5899016251330436', tab: 'raw_v2_onsite_induction', campusCol: 'Campus Code', plannedCol: 'No. of New Contractors (Individuals)', actualCol: 'Contractors Inducted in the Reporting Month', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_investigation_on_time', reportId: '6831846506581892', tab: 'raw_v2_investigation_on_time', campusCol: 'Campus Code', plannedCol: 'Total Incident Investigated', actualCol: 'Investigation Completed on Time', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_planned_training', reportId: '5332685084905348', tab: 'raw_v2_planned_training', campusCol: 'Campus', plannedCol: 'Planned Training', actualCol: 'Training Conducted', monthCol: 'Primary', hasMonth: true },
  { key: 'v2_drills', sheetId: '5053158949605252', tab: 'raw_v2_drills', campusCol: 'Campus Code', monthCol: 'Reporting Month', plannedCol: 'Planned Drill? (Yes/No)', actualCol: 'Are there any submission?', hasMonth: true, yesNoCount: true },
  { key: 'v2_waste_segregation', sheetId: '8150747345538948', tab: 'raw_v2_waste_segregation', campusCol: 'Campus Code', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_mgmt_review_actions', sheetId: '3267265049874308', tab: 'raw_v2_mgmt_review_actions', campusCol: 'Campus Group', monthCol: 'Reporting Month', hasMonth: true },
];

const MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const MONTH_ABBR: Record<string, number> = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };

export function normalizeMonth(v: any): string | null {
  if (!v && v !== 0) return null;
  const s = String(v).trim();
  if (!s) return null;
  const idx = MONTH_NAMES.findIndex(m => m.toLowerCase() === s.toLowerCase());
  if (idx >= 0) return MONTH_NAMES[idx];
  const abbr = s.substring(0, 3).toLowerCase();
  if (MONTH_ABBR[abbr] !== undefined) return MONTH_NAMES[MONTH_ABBR[abbr]];
  const d = new Date(s);
  if (!isNaN(d.getTime())) return MONTH_NAMES[d.getMonth()];
  return null;
}

export async function fetchSheet(sheetId: string, token: string): Promise<Record<string, any>[]> {
  const url = `https://api.smartsheet.com/2.0/sheets/${sheetId}?pageSize=500`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' } });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const data = await res.json();
  if (!data.rows?.length) return [];
  const colMap: Record<number, string> = {};
  (data.columns || []).forEach((col: any) => { if (col.id) colMap[col.id] = col.title; });
  return data.rows.map((row: any) => {
    const rec: Record<string, any> = {};
    (row.cells || []).forEach((cell: any) => {
      const title = colMap[cell.columnId];
      if (title) rec[title] = cell.displayValue ?? cell.value ?? '';
    });
    return rec;
  }).filter((r: any) => Object.keys(r).length > 0);
}

export async function fetchReport(reportId: string, token: string): Promise<Record<string, any>[]> {
  const url = `https://api.smartsheet.com/2.0/reports/${reportId}?pageSize=500&level=1`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' } });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const data = await res.json();
  if (!data.rows?.length) return [];
  const colMap: Record<number, string> = {};
  (data.columns || []).forEach((col: any) => {
    if (col.id) colMap[col.id] = col.title;
    if (col.virtualId) colMap[col.virtualId] = col.title;
  });
  return data.rows.map((row: any) => {
    const rec: Record<string, any> = {};
    (row.cells || []).forEach((cell: any) => {
      const colId = cell.virtualColumnId || cell.columnId;
      const title = colMap[colId];
      if (title) rec[title] = cell.displayValue ?? cell.value ?? '';
    });
    return rec;
  }).filter((r: any) => Object.keys(r).length > 0);
}

export interface KpiRow {
  campus: string;
  month: string | null;
  planned: number;
  actual: number;
  value: number;
}

export function processSource(src: SyncSource, records: Record<string, any>[]): KpiRow[] {
  return records.map(r => {
    const campus = String(r[src.campusCol] || '').trim();
    if (!campus) return null;
// Try configured monthCol first, then fallback columns (matches GAS SheetService logic)
    let month: string | null = null;
    if (src.hasMonth && src.monthCol) {
      month = normalizeMonth(r[src.monthCol]);
      if (!month) month = normalizeMonth(r['Reporting Month']);
      if (!month) month = normalizeMonth(r['Date Reported']);
      if (!month) month = normalizeMonth(r['Primary']);
    }

    let planned = 0, actual = 0, value = 0;

    if (src.yesNoCount) {
      // Yes/No columns: count "Yes" as 1, anything else as 0
      const pVal = String(r[src.plannedCol || ''] || '').trim().toLowerCase();
      const aVal = String(r[src.actualCol || ''] || '').trim().toLowerCase();
      planned = pVal === 'yes' ? 1 : 0;
      actual = aVal === 'yes' ? 1 : 0;
    } else {
      planned = parseFloat(r[src.plannedCol || '']) || 0;
      actual = parseFloat(r[src.actualCol || '']) || 0;
    }
    value = parseFloat(r[src.valueCol || '']) || 0;

    return { campus, month, planned, actual, value } as KpiRow;
  }).filter(Boolean) as KpiRow[];
}
