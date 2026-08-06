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
  hasMonth: boolean;
  isolateFromCampusSet?: boolean;
  yesNoCount?: boolean;
}


export const SYNC_SOURCES: SyncSource[] = [
  // Original 7 KPIs — matched to GAS SyncService.gs
  { key: 'drills', sheetId: '7139786694283140', tab: 'raw_drills', campusCol: 'Campus Code', monthCol: 'Reporting Month', plannedCol: 'Planned Drill? (Yes/No)', actualCol: 'Are there any submission?', hasMonth: true, yesNoCount: true },
  { key: 'ehs', sheetId: '1510149721116548', tab: 'raw_ehs', monthCol: 'Primary', campusCol: 'Campus Code', plannedCol: 'No. of EHS Inspections Planned', actualCol: 'No. of EHS Inspections Completed', hasMonth: true },
  { key: 'findings', sheetId: '1510149721116548', tab: 'raw_findings', monthCol: 'Primary', campusCol: 'Campus Code', plannedCol: 'No. of Findings Due', actualCol: 'No. of Findings Closed', hasMonth: true },
  { key: 'notification', reportId: '8527961731846020', tab: 'raw_notification', monthCol: 'Reporting Month', campusCol: 'Campus Code', plannedCol: 'Total Incident', actualCol: 'Notification Submitted on Time', hasMonth: true },
  { key: 'risk', reportId: '5427282301636484', tab: 'raw_risk', campusCol: 'Campus Code', monthCol: 'Reporting Month', plannedCol: 'Total Assessments Register', actualCol: 'RA Validated and Signed Off', hasMonth: true },
  { key: 'training', sheetId: '4456464805482372', tab: 'raw_training', campusCol: 'Campus Code', valueCol: 'Total Hours', monthCol: 'Reporting Month', hasMonth: true },
  { key: 'incidents', sheetId: '5977763159691140', tab: 'raw_incidents', campusCol: 'Campus Code', monthCol: 'Reporting Month', valueCol: 'Total Incident', hasMonth: true },


  // Pie chart — Incidents by Campus (no Incident Type column in new workspace)
  { key: 'v2_incident_types', reportId: '223991699558276', tab: 'raw_v2_incident_types', campusCol: 'Incident Type', valueCol: 'Total Incident', hasMonth: false, isolateFromCampusSet: true },


  // V2 KPIs — matched to GAS SyncService.gs
  { key: 'v2_hs_committee', sheetId: '5093607634587524', tab: 'raw_v2_hs_committee', campusCol: 'Committee', plannedCol: 'Are there any submission?', actualCol: 'Was a meeting held?', monthCol: 'Reporting Month', hasMonth: true, isolateFromCampusSet: true, yesNoCount: true },
  { key: 'v2_findings_on_time', sheetId: '1510149721116548', tab: 'raw_v2_findings_on_time', campusCol: 'Campus Code', plannedCol: 'No. of Findings Closed', actualCol: 'No. of Findings Due', monthCol: 'Primary', hasMonth: true },
  { key: 'v2_risk_closed', sheetId: '7524088825204612', tab: 'raw_v2_risk_closed', campusCol: 'Campus Code', plannedCol: 'Total Risk Assessments Registered', actualCol: 'Risk Assessment Closed', monthCol: 'Primary', hasMonth: true },
  { key: 'v2_risk_validated', sheetId: '7524088825204612', tab: 'raw_v2_risk_validated', campusCol: 'Campus Code', plannedCol: 'Total Risk Assessments Registered', actualCol: 'Risk Assessment and Validation', monthCol: 'Primary', hasMonth: true },
