// Smartsheet API client — mirrors SyncService.gs SYNC_SOURCES config

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
  hasMonth: boolean;
  isolateFromCampusSet?: boolean;
  yesNoCount?: boolean;
}

export const SYNC_SOURCES: SyncSource[] = [
  { key: 'drills', sheetId: '5053158949605252', tab: 'raw_drills', campusCol: 'Campus Code', monthCol: 'Reporting Month', plannedCol: 'Planned Drill? (Yes/No)', actualCol: 'Are there any submission?', hasMonth: true, yesNoCount: true },
  { key: 'ehs', sheetId: '4947401822392196', tab: 'raw_ehs', monthCol: 'Primary', campusCol: 'Campus Code', plannedCol: 'No. of EHS Inspections Planned', actualCol: 'No. of EHS Inspections Completed', hasMonth: true },
  { key: 'findings', sheetId: '4947401822392196', tab: 'raw_findings', monthCol: 'Primary', campusCol: 'Campus Code', plannedCol: 'No. of Findings in Reporting Month', actualCol: 'No. of Findings Closed', hasMonth: true },
  { key: 'notification', reportId: '1199821531598724', tab: 'raw_notification', monthCol: 'Primary', campusCol: 'Campus', plannedCol: 'Total Incident', actualCol: 'Notification Submitted on Time', hasMonth: true },
  { key: 'risk', reportId: '8565044722749316', tab: 'raw_risk', campusCol: 'Campus', plannedCol: 'Total Assessments Register', actualCol: 'RA Validated and Signed Off', hasMonth: false },
  { key: 'training', reportId: '4366423609528196', tab: 'raw_training', campusCol: 'Primary', valueCol: 'Total Training Hrs', hasMonth: false },
  { key: 'incidents', reportId: '6831846506581892', tab: 'raw_incidents', monthCol: 'Primary', campusCol: 'Campus', valueCol: 'Total Incident', hasMonth: true },
  { key: 'v2_hs_committee', reportId: '3061649186443140', tab: 'raw_v2_hs_committee', campusCol: 'Committee', valueCol: 'No. of Committee Meeting', monthCol: 'Primary', hasMonth: true },
  { key: 'v2_findings_on_time', reportId: '6263081167441796', tab: 'raw_v2_findings_on_time', campusCol: 'Campus Code', plannedCol: 'No. of Findings in Reporting Month', actualCol: 'No. of Findings Due', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_risk_closed', reportId: '352656412331908', tab: 'raw_v2_risk_closed', campusCol: 'Campus', plannedCol: 'Total Risk Assessments Registered', actualCol: 'Risk Assessment Closed', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_risk_validated', reportId: '8565044722749316', tab: 'raw_v2_risk_validated', campusCol: 'Campus', plannedCol: 'Total Assessments Register', actualCol: 'RA Validated and Signed Off', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_ehs_inspection', reportId: '4046605446500228', tab: 'raw_v2_ehs_inspection', campusCol: 'Campus Code', plannedCol: 'No. of EHS Inspections Planned', actualCol: 'No. of EHS Inspections Completed', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_hs_kpi_report', reportId: '4811266391494532', tab: 'raw_v2_hs_kpi_report', campusCol: 'Campuses', valueCol: 'Submitted', monthCol: 'Primary', hasMonth: false },
  { key: 'v2_external_compliance', reportId: '1199958970552196', tab: 'raw_v2_external_compliance', campusCol: 'Campus Code', plannedCol: 'Applicable Compliance', actualCol: 'Actual Compliance', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_safe_working', reportId: '385077644054404', tab: 'raw_v2_safe_working', campusCol: 'Campus', plannedCol: 'No. of SOPs Verified', actualCol: 'No. of SOPs Implemented', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_permit_to_work', reportId: '6636146455957380', tab: 'raw_v2_permit_to_work', campusCol: 'Campus', plannedCol: 'No. of PTWs Issued', actualCol: 'Total Work Registered', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_hazard_id', reportId: '2222036767166340', tab: 'raw_v2_hazard_id', campusCol: 'Campus', plannedCol: 'Total Control Sampled', actualCol: 'Implemented Controls', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_onsite_induction', reportId: '6619039466999684', tab: 'raw_v2_onsite_induction', campusCol: 'Campus', plannedCol: 'No. of New Contractors (Individuals)', actualCol: 'Contractors Inducted in the Reporting Month', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_investigation_on_time', reportId: '6831846506581892', tab: 'raw_v2_investigation_on_time', campusCol: 'Campus', plannedCol: 'Total Incident', actualCol: 'Investigation Completed on Time', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_training_hours', reportId: '3729812953714564', tab: 'raw_v2_training_hours', campusCol: 'Campus', valueCol: 'No. of Trainings Hours', monthCol: 'Primary', hasMonth: false },
  { key: 'v2_planned_training', reportId: '5332685084905348', tab: 'raw_v2_planned_training', campusCol: 'Campus', plannedCol: 'Planned Training', actualCol: 'Training Conducted', monthCol: 'Primary', hasMonth: false },
  { key: 'v2_drills', sheetId: '5053158949605252', tab: 'raw_v2_drills', campusCol: 'Campus Code', monthCol: 'Reporting Month', plannedCol: 'Planned Drill? (Yes/No)', actualCol: 'Are there any submission?', hasMonth: true, yesNoCount: true },
  { key: 'v2_waste_segregation', sheetId: '4947401822392196', tab: 'raw_v2_waste_segregation', campusCol: 'Campus Code', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'v2_mgmt_review_actions', reportId: '1234567890', tab: 'raw_v2_mgmt_review_actions', campusCol: 'Campus', monthCol: 'Reporting Month', hasMonth: true },
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
    const month = (src.hasMonth && src.monthCol) ? normalizeMonth(r[src.monthCol]) : null;
    const planned = parseFloat(r[src.plannedCol || '']) || 0;
    const actual = parseFloat(r[src.actualCol || '']) || 0;
    const value = parseFloat(r[src.valueCol || '']) || 0;
    return { campus, month, planned, actual, value } as KpiRow;
  }).filter(Boolean) as KpiRow[];
}
