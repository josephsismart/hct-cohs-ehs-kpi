'use client';
import { useState, useMemo, useCallback, useEffect } from 'react';
import Highcharts from 'highcharts';
import HighchartsReact from 'highcharts-react-official';

const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const QUARTERS: Record<string, string[]> = {
  Q1: ['January','February','March'], Q2: ['April','May','June'],
  Q3: ['July','August','September'], Q4: ['October','November','December'],
};
const QUARTER_LABELS: Record<string, string> = {
  Q1: 'Q1 (Jan–Mar)', Q2: 'Q2 (Apr–Jun)',
  Q3: 'Q3 (Jul–Sep)', Q4: 'Q4 (Oct–Dec)',
};

const SUMMARY_CARDS = [
  { key: 'incidents', label: 'Total Incidents', unit: 'count', color: '#dc3545' },
  { key: 'training', label: 'Training Hours', unit: 'val', color: '#0d6efd' },
  { key: 'ehs', label: 'EHS Inspection Rate', unit: 'pct', color: '#198754' },
  { key: 'drills', label: 'Drills Completion', unit: 'pct', color: '#20c997' },
  { key: 'findings', label: 'Findings Closed', unit: 'pct', color: '#ffc107' },
  { key: 'notification', label: 'Notification Rate', unit: 'pct', color: '#fd7e14' },
  { key: 'risk', label: 'Risk Assessment', unit: 'pct', color: '#6f42c1' },
];

const SMARTSHEET_LINKS: Record<string, string> = {
  v2_onsite_induction: 'https://app.smartsheet.com/reports/488MxwHw83MF8fwxqJx95PgqwjGrjgMXFR7hXGq1',
  v2_permit_to_work: 'https://app.smartsheet.com/reports/MgpHGX276R7R7pxw4jx8X5JcCX3PQRHHG7p62hc1',
  v2_hazard_id: 'https://app.smartsheet.com/reports/JCXVCQFPM5F6vWgjcVj3Mm66Mp7VMH97G2MvcMC1',
  v2_risk_closed: 'https://app.smartsheet.com/reports/MWxppxgGMMq85WhfCpg57x3w92Qqp6pHg79Gwx71',
  v2_risk_validated: 'https://app.smartsheet.com/reports/FvCmH32V3fp6pMPgp7gX2wv53chVCrW57qqGQ5H1',
  v2_safe_working: 'https://app.smartsheet.com/reports/39xG9RjhqVjR6fPpV4Qgh4Gr2w7X9rGM8xccWqc1',
  v2_findings_on_time: 'https://app.smartsheet.com/reports/wC59JHJM3x57Q6vFCRFmVpgc93gXr8GxxQ2gwXq1',
  v2_ehs_inspection: 'https://app.smartsheet.com/reports/FmV6pG8cXJg5cfhxgppGWwwJM94Qv4QQWfPh2j61',
  v2_hs_committee: 'https://app.smartsheet.com/reports/X828rgMpqgWpj2MRFRj7wRw9CG465PGfxQMGrg91',
  v2_planned_training: 'https://app.smartsheet.com/reports/6wXHp4RpPR378F6F2r7Q9FpvjmMqhFmcgGwXcmp1',
  incidents: 'https://app.smartsheet.com/reports/Gcg9cW5qG3FGXR3xcwcrHHwxMR8w9QxXp9JxHVC1',
  v2_incident_types: 'https://app.smartsheet.com/reports/pJJP7hJ3WghvVfC9w7GPx73mMwVCvRq4X3fC2Cx1',
  training: 'https://app.smartsheet.com/reports/v5VMcRR6j97qWvFjJm9rjr9WPVmpGcfg3jfvg561',
  ehs_rate: 'https://app.smartsheet.com/reports/FmV6pG8cXJg5cfhxgppGWwwJM94Qv4QQWfPh2j61',
  findings_rate: 'https://app.smartsheet.com/reports/wC59JHJM3x57Q6vFCRFmVpgc93gXr8GxxQ2gwXq1',
};

const KPI_CHARTS = [
  { key: 'v2_onsite_induction', label: 'Contractor Activity', plannedLabel: 'No. of New Contractors (Individuals)', actualLabel: 'Contractors Inducted in the Reporting Month — Met/Exceeded', belowLabel: 'Contractors Inducted in the Reporting Month — Below Target', type: 'planned_actual_below' },
  { key: 'v2_permit_to_work', label: 'Permit to Work', plannedLabel: 'No. of PTWs Issued', actualLabel: 'Total Work Registered', type: 'planned_actual' },
  { key: 'v2_hazard_id', label: 'Implemented Control Measures', plannedLabel: 'Total Control Sampled', actualLabel: 'Implemented Controls — Met/Exceeded', belowLabel: 'Implemented Controls — Below Target', type: 'planned_actual_below' },
  { key: 'v2_risk_closed', label: 'Risk Assessment Closed', plannedLabel: 'Total Risk Assessments Registered', actualLabel: 'Risk Assessment Closed — Met/Exceeded', belowLabel: 'Risk Assessment Closed — Below Target', type: 'planned_actual_below' },
  { key: 'v2_risk_validated', label: 'Risk Assessment Validated & Signed Off', plannedLabel: 'Total Assessments Register', actualLabel: 'RA Validated and Signed Off — Met/Exceeded', belowLabel: 'RA Validated and Signed Off — Below Target', type: 'planned_actual_below' },
  { key: 'v2_safe_working', label: 'Safe Working Procedure', plannedLabel: 'No. of SOPs Verified', actualLabel: 'No. of SOPs Implemented — Met/Exceeded', belowLabel: 'No. of SOPs Implemented — Below Target', type: 'planned_actual_below' },
  { key: 'v2_findings_on_time', label: 'Findings Closed On Time', plannedLabel: 'No. of Findings in Reporting Month', actualLabel: 'No. of Findings Closed — Met/Exceeded', belowLabel: 'No. of Findings Closed — Below Target', type: 'planned_actual_below' },
  { key: 'v2_ehs_inspection', label: 'Scheduled EHS Inspection', plannedLabel: 'No. of EHS Inspections Planned', actualLabel: 'No. of EHS Inspections Completed — Met/Exceeded', belowLabel: 'No. of EHS Inspections Completed — Below Target', type: 'planned_actual_below' },
  { key: 'v2_hs_committee', label: 'EHS Committee Meeting', valueLabel: 'No. of Committee Meeting', type: 'value' },
  { key: 'v2_planned_training', label: 'Planned Training Report', plannedLabel: 'Planned Training', actualLabel: 'Training Conducted — Met/Exceeded', belowLabel: 'Training Conducted — Below Target', type: 'planned_actual_below' },
  { key: 'v2_hs_kpi_report', label: 'HS KPI Report', valueLabel: 'Submitted', type: 'value' },
  { key: 'v2_external_compliance', label: 'External Authority Compliance', plannedLabel: 'Applicable Compliance', actualLabel: 'Actual Compliance — Met/Exceeded', belowLabel: 'Actual Compliance — Below Target', type: 'planned_actual_below' },
  { key: 'v2_investigation_on_time', label: 'Investigation Completed on Time', plannedLabel: 'Total Incident', actualLabel: 'Investigation Completed on Time', type: 'planned_actual' },
  { key: 'notification', label: 'Notification on Time', plannedLabel: 'Total Incident', actualLabel: 'Notification Submitted on Time', type: 'planned_actual' },
];

const EXTRA_CHARTS = [
  { key: 'incidents', label: 'Total Incidents', subtitle: 'Incident count by campus — lower is better', valueLabel: 'Incidents', type: 'value' },
  { key: 'v2_incident_types', label: 'Incidents by Type', subtitle: 'Count of incidents per category', type: 'pie' },
  { key: 'training', label: 'Total Training Hours by Campus', subtitle: 'Actual training hours per campus', valueLabel: 'Hours', type: 'value_hours' },
  { key: 'ehs_rate', label: 'EHS Inspection Rate', subtitle: 'Scheduled vs Completed EHS Inspections', sourceKey: 'ehs', type: 'rate_pct' },
  { key: 'findings_rate', label: 'Findings Closed Rate', subtitle: 'Total Findings vs Findings Closed', sourceKey: 'findings', type: 'rate_pct' },
];

interface KpiRow { campus: string; month: string | null; planned: number; actual: number; value: number; }
interface SyncData {
  syncedAt: string;
  sources: Record<string, { rows: KpiRow[]; error?: string }>;
  campuses: string[];
  months: string[];
  errors: string[];
  wasteData?: Record<string, any>[];
}

function filterRows(rows: KpiRow[], campus: string, month: string, quarter: string): KpiRow[] {
  let filtered = rows;
  if (campus !== 'ALL') filtered = filtered.filter(r => r.campus === campus);
  if (month !== 'ALL') {
    filtered = filtered.filter(r => r.month === month);
  } else if (quarter !== 'ALL' && QUARTERS[quarter]) {
    filtered = filtered.filter(r => r.month && QUARTERS[quarter].includes(r.month));
  }
  return filtered;
}

function aggregate(rows: KpiRow[]) {
  return rows.reduce((acc, r) => ({
    planned: acc.planned + r.planned, actual: acc.actual + r.actual, value: acc.value + r.value,
  }), { planned: 0, actual: 0, value: 0 });
}

function aggregateByCampus(rows: KpiRow[]) {
  const map: Record<string, { planned: number; actual: number; value: number }> = {};
  rows.forEach(r => {
    if (!map[r.campus]) map[r.campus] = { planned: 0, actual: 0, value: 0 };
    map[r.campus].planned += r.planned;
    map[r.campus].actual += r.actual;
    map[r.campus].value += r.value;
  });
  return map;
}

function KpiBarChart({ chartDef, rows }: { chartDef: typeof KPI_CHARTS[0]; rows: KpiRow[] }) {
  const byCampus = aggregateByCampus(rows);
  const campuses = Object.keys(byCampus).sort();
  if (campuses.length === 0) return <div className="no-data">No data available</div>;

  const series: Highcharts.SeriesOptionsType[] = [];
  if (chartDef.type === 'value') {
    series.push({
      type: 'column', name: chartDef.valueLabel || 'Value',
      data: campuses.map(c => byCampus[c].value || byCampus[c].actual || byCampus[c].planned),
      color: '#4A90D9',
    });
  } else if (chartDef.type === 'planned_actual_below') {
    series.push(
      { type: 'column', name: chartDef.plannedLabel || 'Planned', data: campuses.map(c => byCampus[c].planned), color: 'rgba(74,144,217,0.4)' },
      { type: 'column', name: chartDef.actualLabel || 'Met/Exceeded', data: campuses.map(c => Math.max(0, byCampus[c].actual)), color: '#1D9E75' },
      { type: 'column', name: chartDef.belowLabel || 'Below Target', data: campuses.map(c => Math.max(0, byCampus[c].planned - byCampus[c].actual)), color: '#EA352E' },
    );
  } else {
    series.push(
      { type: 'column', name: chartDef.plannedLabel || 'Planned / Target', data: campuses.map(c => byCampus[c].planned), color: 'rgba(74,144,217,0.4)' },
      { type: 'column', name: chartDef.actualLabel || 'Actual', data: campuses.map(c => byCampus[c].actual), color: '#0a3d62' },
    );
  }

  const options: Highcharts.Options = {
    chart: { type: 'column', height: 280, style: { fontFamily: "'Segoe UI', Arial, sans-serif" } },
    title: { text: undefined },
    xAxis: { categories: campuses, labels: { style: { fontSize: '10px' } } },
    yAxis: { title: { text: null }, gridLineColor: '#f0f0f0' },
    legend: { align: 'center', verticalAlign: 'bottom', itemStyle: { fontSize: '10px' } },
    plotOptions: {
      column: { borderRadius: 2, groupPadding: 0.15, pointPadding: 0.05, dataLabels: { enabled: true, style: { fontSize: '9px', fontWeight: 'normal' } } },
    },
    series,
    credits: { enabled: false },
    tooltip: { shared: true },
  };

  return <HighchartsReact highcharts={Highcharts} options={options} />;
}

const PIE_COLORS = ['#1A1F71', '#F59E0B', '#1D9E75', '#4A90D9', '#7C3AED', '#FFA500', '#EA352E'];

function KpiPieChart({ rows }: { rows: KpiRow[] }) {
  const byCampus = aggregateByCampus(rows);
  const entries = Object.entries(byCampus).sort((a, b) => b[1].value - a[1].value);
  if (entries.length === 0) return <div className="no-data">No data available</div>;
  const options: Highcharts.Options = {
    chart: { type: 'pie', height: 500, style: { fontFamily: "'Segoe UI', Arial, sans-serif" } },
    title: { text: undefined },
    plotOptions: { pie: { dataLabels: { enabled: true, format: '{point.y}', distance: -30, style: { fontSize: '14px', fontWeight: 'bold', color: 'white', textOutline: 'none' } }, showInLegend: true } },
    legend: { align: 'right', verticalAlign: 'middle', layout: 'vertical', itemStyle: { fontSize: '11px' } },
    series: [{ type: 'pie', name: 'Incidents', data: entries.map(([name, v], i) => ({ name, y: v.value, color: PIE_COLORS[i % PIE_COLORS.length] })) }],
    credits: { enabled: false },
  };
  return <HighchartsReact highcharts={Highcharts} options={options} />;
}

function KpiRateChart({ rows }: { rows: KpiRow[] }) {
  const byCampus = aggregateByCampus(rows);
  const campuses = Object.keys(byCampus).sort();
  if (campuses.length === 0) return <div className="no-data">No data available</div>;
  const data = campuses.map(c => {
    const { planned, actual } = byCampus[c];
    const pct = planned > 0 ? Math.min(Math.round(actual / planned * 100), 100) : 0;
    return { campus: c, pct, met: pct >= 100 };
  });
  const options: Highcharts.Options = {
    chart: { type: 'column', height: 280, style: { fontFamily: "'Segoe UI', Arial, sans-serif" } },
    title: { text: undefined },
    xAxis: { categories: campuses, labels: { style: { fontSize: '10px' } } },
    yAxis: { title: { text: null }, max: 100, labels: { format: '{value}%' }, gridLineColor: '#f0f0f0' },
    legend: { align: 'center', verticalAlign: 'bottom', itemStyle: { fontSize: '10px' } },
    plotOptions: { column: { borderRadius: 2, groupPadding: 0.15, pointPadding: 0.05, dataLabels: { enabled: true, format: '{y}%', style: { fontSize: '9px', fontWeight: 'normal' } } } },
    series: [
      { type: 'column', name: 'Met Target', data: data.map(d => ({ y: d.met ? d.pct : 0, color: '#1D9E75' })), showInLegend: true, color: '#1D9E75' },
      { type: 'column', name: 'Below Target', data: data.map(d => ({ y: !d.met && d.pct > 0 ? d.pct : 0, color: '#EA352E' })), showInLegend: true, color: '#EA352E' },
    ],
    credits: { enabled: false },
    tooltip: { shared: true },
  };
  return <HighchartsReact highcharts={Highcharts} options={options} />;
}

function KpiValueHoursChart({ rows }: { rows: KpiRow[] }) {
  const byCampus = aggregateByCampus(rows);
  const campuses = Object.keys(byCampus).sort();
  if (campuses.length === 0) return <div className="no-data">No data available</div>;
  const options: Highcharts.Options = {
    chart: { type: 'column', height: 280, style: { fontFamily: "'Segoe UI', Arial, sans-serif" } },
    title: { text: undefined },
    xAxis: { categories: campuses, labels: { style: { fontSize: '10px' } } },
    yAxis: { title: { text: null }, gridLineColor: '#f0f0f0' },
    legend: { enabled: false },
    plotOptions: { column: { borderRadius: 2, dataLabels: { enabled: true, format: '{y}h', style: { fontSize: '9px', fontWeight: 'normal' } } } },
    series: [{ type: 'column', name: 'Hours', data: campuses.map(c => byCampus[c].value || byCampus[c].actual), color: '#4A90D9' }],
    credits: { enabled: false },
  };
  return <HighchartsReact highcharts={Highcharts} options={options} />;
}

function SyncBadge({ syncedAt }: { syncedAt: string }) {
  const mins = Math.round((Date.now() - new Date(syncedAt).getTime()) / 60000);
  let label: string;
  let cls: string;
  if (mins < 5) { label = 'Just now'; cls = 'sync-fresh'; }
  else if (mins < 60) { label = `${mins} min ago`; cls = 'sync-fresh'; }
  else if (mins < 1440) { label = `${Math.round(mins / 60)} hr ago`; cls = 'sync-stale'; }
  else { label = `${Math.round(mins / 1440)} days ago`; cls = 'sync-old'; }
  return <span className={`sync-badge ${cls}`}>Synced {label}</span>;
}

export default function Dashboard() {
  const [data, setData] = useState<SyncData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [campus, setCampus] = useState('ALL');
  const [month, setMonth] = useState('ALL');
  const [quarter, setQuarter] = useState('ALL');
  const [year, setYear] = useState('ALL');
  const [darkMode, setDarkMode] = useState(false);
  const [showReport, setShowReport] = useState(false);
  const [reportName, setReportName] = useState('');
  const [pptRegion, setPptRegion] = useState('Abu Dhabi');
  const [pptLoading, setPptLoading] = useState(false);
  const [wordLoading, setWordLoading] = useState(false);
  const [pdfLoading, setPdfLoading] = useState(false);

  const getReportParams = useCallback(() => {
    const m = month !== 'ALL' ? month : MONTHS[new Date().getMonth() - 1] || 'December';
    const y = year !== 'ALL' ? year : String(new Date().getFullYear());
    return { month: m, year: y };
  }, [month, year]);

  const downloadFile = useCallback(async (endpoint: string, ext: string, setLoading: (v: boolean) => void) => {
    setLoading(true);
    try {
      const { month: m, year: y } = getReportParams();
      const url = `/api/${endpoint}?region=${encodeURIComponent(pptRegion)}&month=${encodeURIComponent(m)}&year=${y}`;
      const res = await fetch(url);
      if (!res.ok) { const err = await res.json(); throw new Error(err.error || `HTTP ${res.status}`); }
      const blob = await res.blob();
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `HCT_KPI_${pptRegion.replace(/ /g, '_')}_${m}_${y}.${ext}`;
      a.click();
      URL.revokeObjectURL(a.href);
    } catch (e: any) { alert(`${ext.toUpperCase()} generation failed: ` + e.message); }
    finally { setLoading(false); }
  }, [pptRegion, getReportParams]);

  const downloadPpt = useCallback(() => downloadFile('generate-ppt', 'pptx', setPptLoading), [downloadFile]);
  const downloadWord = useCallback(() => downloadFile('generate-word', 'docx', setWordLoading), [downloadFile]);
  const downloadPdf = useCallback(() => downloadFile('generate-pdf', 'pdf', setPdfLoading), [downloadFile]);

  const doSync = useCallback(async () => {
    setLoading(true); setError('');
    try {
      const res = await fetch('/api/sync', { cache: 'no-store' });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      setData(await res.json());
    } catch (e: any) { setError(e.message); }
    finally { setLoading(false); }
  }, []);

  // Auto-sync on first load
  useEffect(() => {
    doSync();
    // Auto-poll every 60 seconds for near real-time updates
    const interval = setInterval(() => {
      fetch('/api/sync', { cache: 'no-store' })
        .then(res => res.ok ? res.json() : null)
        .then(d => { if (d) setData(d); })
        .catch(() => {});
    }, 60000);
    return () => clearInterval(interval);
  }, [doSync]);

  // Dark mode toggle
  useEffect(() => {
    document.documentElement.classList.toggle('dark', darkMode);
  }, [darkMode]);

  // Generate report name from filters
  useEffect(() => {
    const parts = ['HCT_COHS_EHS_KPI'];
    if (year !== 'ALL') parts.push(year);
    if (quarter !== 'ALL') parts.push(quarter);
    if (month !== 'ALL') parts.push(month);
    if (campus !== 'ALL') parts.push(campus);
    setReportName(parts.join('_'));
  }, [year, quarter, month, campus]);

  const getRows = useCallback((key: string): KpiRow[] => {
    if (!data?.sources[key]) return [];
    const rows = data.sources[key].rows;
    const allNullMonth = rows.every(r => !r.month);
    if (allNullMonth && (month !== 'ALL' || quarter !== 'ALL'))
      return filterRows(rows, campus, 'ALL', 'ALL');
    return filterRows(rows, campus, month, quarter);
  }, [data, campus, month, quarter]);

  const summaryValues = useMemo(() => {
    if (!data) return null;
    return SUMMARY_CARDS.map(card => {
      const rows = getRows(card.key);
      const agg = aggregate(rows);
      if (card.unit === 'pct') {
        const pct = agg.planned > 0 ? Math.min(Math.round(agg.actual / agg.planned * 100), 100) : 0;
        return { ...card, display: `${pct}%`, sub: `${agg.actual} / ${agg.planned}` };
      }
      if (card.unit === 'val') {
        const val = agg.value || agg.actual || agg.planned;
        return { ...card, display: card.key === 'training' ? `${val} hrs` : String(val), sub: '' };
      }
      const val = agg.value || agg.actual || agg.planned;
      return { ...card, display: String(val), sub: `${data.campuses.length} campuses` };
    });
  }, [data, getRows]);

  return (
    <div className={`dashboard${darkMode ? " dark" : ""}`}>
      {/* REPORT MODAL */}
      {showReport && (
        <div className="modal-overlay" onClick={() => setShowReport(false)}>
          <div className="modal-content" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h3><i className="fa fa-file-export" style={{ marginRight: 8 }}></i>Export Report</h3>
              <button className="modal-close" onClick={() => setShowReport(false)}><i className="fa fa-times"></i></button>
            </div>
            <div className="modal-body">
              <label>Report Name</label>
              <input type="text" value={reportName} onChange={e => setReportName(e.target.value)} placeholder="Enter report name..." />
              <label>Choose Format</label>
              <div className="report-buttons">
                <button className="report-btn ppt" onClick={() => { window.open('/api/generate-ppt?name=' + encodeURIComponent(reportName)); setShowReport(false); }}>
                  <i className="fa fa-file-powerpoint"></i>
                  <span>PowerPoint</span>
                </button>
                <button className="report-btn word" onClick={() => { window.open('/api/generate-word?name=' + encodeURIComponent(reportName)); setShowReport(false); }}>
                  <i className="fa fa-file-word"></i>
                  <span>Word</span>
                </button>
                <button className="report-btn pdf" onClick={() => { window.open('/api/generate-pdf?name=' + encodeURIComponent(reportName)); setShowReport(false); }}>
                  <i className="fa fa-file-pdf"></i>
                  <span>PDF</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* HEADER */}
      <div className="brand-header">
        <div className="brand-logo brand-ohs">
          <img src="/ohs-logo.png" alt="OHS" onError={(e) => (e.currentTarget.style.display = 'none')} />
        </div>
        <div className="brand-title">HCT - COHS EHS KPI REPORTING</div>
        <div className="brand-logo brand-hct">
          <img src="/hct-logo.png" alt="HCT" onError={(e) => (e.currentTarget.style.display = 'none')} />
        </div>
      </div>

      {/* FILTER BAR */}
      <div className="filter-bar">
        <span className="filter-label">Health &amp; Safety Performance Report</span>
        <div className="filter-controls">
          <label>Year
            <select value={year} onChange={e => setYear(e.target.value)}>
              <option value="ALL">All Years</option>
              <option value="2026">2026</option>
            </select>
          </label>
          <label>Quarter
            <select value={quarter} onChange={e => { setQuarter(e.target.value); if (e.target.value !== 'ALL') setMonth('ALL'); }}>
              <option value="ALL">All Quarters</option>
              {Object.keys(QUARTERS).map(q => <option key={q} value={q}>{QUARTER_LABELS[q]}</option>)}
            </select>
          </label>
          <label>Month
            <select value={month} onChange={e => { setMonth(e.target.value); if (e.target.value !== 'ALL') setQuarter('ALL'); }}>
              <option value="ALL">All Months</option>
              {MONTHS.map(m => <option key={m} value={m}>{m}</option>)}
            </select>
          </label>
          <label>Campus
            <select value={campus} onChange={e => setCampus(e.target.value)}>
              <option value="ALL">All Campuses</option>
              {data?.campuses.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </label>
          <div className="btn-group">
            <button className="btn-refresh" onClick={() => { setCampus('ALL'); setMonth('ALL'); setQuarter('ALL'); setYear('ALL'); }}><i className="fa fa-redo"></i> Refresh</button>
            <button className="btn-sync" onClick={doSync} disabled={loading}>
              <i className="fa fa-sync"></i> {loading ? 'Syncing...' : 'Sync Now'}
            </button>
            <button className="btn-theme" onClick={() => setDarkMode(!darkMode)}>
              <i className={darkMode ? 'fa fa-sun' : 'fa fa-moon'}></i> {darkMode ? 'Light' : 'Dark'}
            </button>
            <button className="btn-report" onClick={() => setShowReport(true)}>
              <i className="fa fa-file-export"></i> Export Report
            </button>
          </div>
          {data && <SyncBadge syncedAt={data.syncedAt} />}
        </div>
      </div>

      {/* CONTENT */}
      <div className="content">
        {error && <div className="alert-error">{error}</div>}

        {!data && !loading && (
          <div className="empty-state">
            <h3>No data loaded</h3>
            <p>Click &quot;Sync Now&quot; to fetch data from Smartsheet</p>
            <button className="btn-sync btn-lg" onClick={doSync}>Sync Now</button>
          </div>
        )}

        {loading && (
          <div className="loading-overlay">
            <div className="spinner" />
            <p>Syncing data from Smartsheet...</p>
          </div>
        )}

        {data && (
          <>
            <h3 className="section-title">OVERALL KPI SUMMARY</h3>
            <div className="kpi-summary-grid">
              {summaryValues?.map(card => (
                <div key={card.key} className="kpi-summary-card">
                  <div className="kpi-card-label">{card.label.toUpperCase()}</div>
                  <div className="kpi-card-value" style={{ color: card.color }}>{card.display}</div>
                  {card.sub && <div className="kpi-card-sub">{card.sub}</div>}
                </div>
              ))}
            </div>
            <div className="legend-row">
              <span className="legend-dot" style={{ background: '#4A90D9' }} /> Planned / Target
              <span className="legend-dot" style={{ background: '#1D9E75' }} /> Actual {'—'} Met or Exceeded
              <span className="legend-dot" style={{ background: '#EA352E' }} /> Actual {'—'} Below Target
              <span className="legend-dot" style={{ background: '#F59E0B' }} /> No target set
            </div>

            <h3 className="section-title">KPI CHARTS</h3>
            <div className="charts-grid">
              {KPI_CHARTS.map(chartDef => {
                const rows = getRows(chartDef.key);
                const ssLink = SMARTSHEET_LINKS[chartDef.key];
                return (
                  <div key={chartDef.key} className="chart-card">
                    <div className="chart-card-header">
                      <span>{chartDef.label}</span>
                      {ssLink ? <a className="btn-smartsheet" href={ssLink} target="_blank" rel="noopener noreferrer">View in Smartsheet</a> : null}
                    </div>
                    <div className="chart-card-body">
                      <KpiBarChart chartDef={chartDef} rows={rows} />
                    </div>
                  </div>
                );
              })}
            </div>

            <h3 className="section-title">ADDITIONAL REPORTS</h3>
            <div className="charts-grid">
              {EXTRA_CHARTS.map(chartDef => {
                const sourceKey = (chartDef as any).sourceKey || chartDef.key;
                const rows = getRows(sourceKey);
                const ssLink = SMARTSHEET_LINKS[chartDef.key];
                return (
                  <div key={chartDef.key} className="chart-card">
                    <div className="chart-card-header">
                      <div>
                        <span>{chartDef.label}</span>
                        {chartDef.subtitle && <div style={{ fontSize: '10px', color: '#999', fontWeight: 400 }}>{chartDef.subtitle}</div>}
                      </div>
                      {ssLink ? <a className="btn-smartsheet" href={ssLink} target="_blank" rel="noopener noreferrer">View in Smartsheet</a> : null}
                    </div>
                    <div className="chart-card-body">
                      {chartDef.type === 'pie' ? <KpiPieChart rows={rows} /> :
                       chartDef.type === 'rate_pct' ? <KpiRateChart rows={rows} /> :
                       chartDef.type === 'value_hours' ? <KpiValueHoursChart rows={rows} /> :
                       <KpiBarChart chartDef={chartDef as any} rows={rows} />}
                    </div>
                  </div>
                );
              })}
            </div>

            {/* WASTE SEGREGATION */}
            <h3 className="section-title">WASTE SEGREGATION</h3>
            {(() => {
              const WASTE_COLS = ['Total Waste','General Waste','Food Waste','Paper Waste','Paper Cup/Carton','PET Bottle','Single Use Plastic'];
              let wd = data.wasteData || [];
              if (campus !== 'ALL') wd = wd.filter((r: any) => r.campus === campus);
              if (month !== 'ALL') wd = wd.filter((r: any) => r.month === month);
              else if (quarter !== 'ALL' && QUARTERS[quarter]) wd = wd.filter((r: any) => r.month && QUARTERS[quarter].includes(r.month));
              // Aggregate by campus
              const map: Record<string, Record<string, number>> = {};
              wd.forEach((r: any) => {
                if (!map[r.campus]) map[r.campus] = {};
                WASTE_COLS.forEach(col => { map[r.campus][col] = (map[r.campus][col] || 0) + (r[col] || 0); });
              });
              // Compute Total Waste as sum of subcategories (API returns 0 for formula columns)
              const SUB_COLS = ['General Waste','Food Waste','Paper Waste','Paper Cup/Carton','PET Bottle','Single Use Plastic'];
              Object.keys(map).forEach(c => { map[c]['Total Waste'] = SUB_COLS.reduce((s, col) => s + (map[c][col] || 0), 0); });
              const campuses = Object.keys(map).sort();
              if (campuses.length === 0) return <div className="placeholder-section"><h4>WASTE SEGREGATION</h4><p>No data available</p></div>;
              // Compute totals row
              const totals: Record<string, number> = {};
              WASTE_COLS.forEach(col => { totals[col] = campuses.reduce((s, c) => s + (map[c][col] || 0), 0); });
              return (
                <div className="chart-card" style={{ overflowX: 'auto' }}>
                  <div className="chart-card-header">
                    <span>Waste Segregation by Campus (kg)</span>
                    <a className="btn-smartsheet" href="https://app.smartsheet.com/reports/FwPfmrvwgpxRMmvp8VX6m59WvpvQ5RGvFCXF8Wx1" target="_blank" rel="noopener noreferrer">View in Smartsheet</a>
                  </div>
                  <table className="waste-table">
                    <thead>
                      <tr><th>Campus</th>{WASTE_COLS.map(c => <th key={c}>{c}</th>)}</tr>
                    </thead>
                    <tbody>
                      {campuses.map(c => (
                        <tr key={c}><td>{c}</td>{WASTE_COLS.map(col => <td key={col}>{(map[c][col] || 0).toLocaleString(undefined, {maximumFractionDigits: 1})}</td>)}</tr>
                      ))}
                      <tr className="totals-row"><td><strong>Total</strong></td>{WASTE_COLS.map(col => <td key={col}><strong>{totals[col].toLocaleString(undefined, {maximumFractionDigits: 1})}</strong></td>)}</tr>
                    </tbody>
                  </table>
                </div>
              );
            })()}

            {/* EXECUTIVE KPI SUMMARY */}
            <h3 className="section-title">EXECUTIVE KPI SUMMARY — BY CAMPUS</h3>
            {(() => {
              const EXEC_KPIS = [
                { key: 'drills', label: 'Drills Completion', type: 'pct' },
                { key: 'ehs', label: 'EHS Inspection', type: 'pct' },
                { key: 'findings', label: 'Findings Closed', type: 'pct' },
                { key: 'notification', label: 'Incident Notification', type: 'pct' },
                { key: 'risk', label: 'Risk Assessment', type: 'pct' },
                { key: 'training', label: 'Training Hours', type: 'val' },
                { key: 'incidents', label: 'Total Incidents', type: 'count' },
              ];
              // Get all campuses from the data
              const campusSet = new Set<string>();
              EXEC_KPIS.forEach(kpi => {
                const rows = getRows(kpi.key);
                rows.forEach(r => { if (r.campus) campusSet.add(r.campus); });
              });
              const allCampuses = [...campusSet].sort();
              if (allCampuses.length === 0) return <div className="placeholder-section"><h4>EXECUTIVE KPI SUMMARY</h4><p>No data available</p></div>;

              // Compute per-campus values
              const campusData: Record<string, Record<string, { planned: number; actual: number; value: number }>> = {};
              allCampuses.forEach(c => { campusData[c] = {}; });
              EXEC_KPIS.forEach(kpi => {
                const rows = getRows(kpi.key);
                allCampuses.forEach(c => {
                  const cr = rows.filter(r => r.campus === c);
                  const agg = cr.reduce((a, r) => ({ planned: a.planned + r.planned, actual: a.actual + r.actual, value: a.value + r.value }), { planned: 0, actual: 0, value: 0 });
                  campusData[c][kpi.key] = agg;
                });
              });

              // Color function
              const pctColor = (pct: number) => {
                if (pct >= 100) return '#c6efce';
                if (pct >= 75) return '#ffeb9c';
                if (pct >= 50) return '#fdd';
                return '#ffc7ce';
              };

              // Total row
              const totals: Record<string, { planned: number; actual: number; value: number }> = {};
              EXEC_KPIS.forEach(kpi => {
                totals[kpi.key] = allCampuses.reduce((a, c) => {
                  const d = campusData[c][kpi.key];
                  return { planned: a.planned + d.planned, actual: a.actual + d.actual, value: a.value + d.value };
                }, { planned: 0, actual: 0, value: 0 });
              });

              const renderCell = (kpi: typeof EXEC_KPIS[0], d: { planned: number; actual: number; value: number }) => {
                if (kpi.type === 'val') return <td style={{ background: '#dce6f1' }}>{d.value ? d.value.toLocaleString() + ' hrs' : '—'}</td>;
                if (kpi.type === 'count') return <td style={{ background: d.value > 0 ? '#ffc7ce' : '#c6efce' }}>{d.value || '—'}</td>;
                if (!d.planned) return <td style={{ background: '#f0f0f0' }}>{'—'}</td>;
                const pct = Math.min(100, Math.round((d.actual / d.planned) * 100));
                return <td style={{ background: pctColor(pct) }}>{pct}%</td>;
              };

              return (
                <div className="chart-card" style={{ overflowX: 'auto' }}>
                  <table className="waste-table">
                    <thead>
                      <tr><th>Campus</th>{EXEC_KPIS.map(k => <th key={k.key}>{k.label}</th>)}</tr>
                    </thead>
                    <tbody>
                      {allCampuses.map(c => (
                        <tr key={c}><td>{c}</td>{EXEC_KPIS.map(kpi => renderCell(kpi, campusData[c][kpi.key]))}</tr>
                      ))}
                      <tr className="totals-row">
                        <td><strong>TOTAL / AVG</strong></td>
                        {EXEC_KPIS.map(kpi => {
                          const d = totals[kpi.key];
                          if (kpi.type === 'val') return <td key={kpi.key} style={{ background: '#dce6f1' }}><strong>{d.value ? d.value.toLocaleString() + ' hrs' : '—'}</strong></td>;
                          if (kpi.type === 'count') return <td key={kpi.key} style={{ background: '#ffc7ce' }}><strong>{d.value}</strong></td>;
                          const pct = d.planned ? Math.min(100, Math.round((d.actual / d.planned) * 100)) : 0;
                          return <td key={kpi.key} style={{ background: pctColor(pct) }}><strong>{pct}%</strong></td>;
                        })}
                      </tr>
                    </tbody>
                  </table>
                </div>
              );
            })()}

            {/* 6-MONTH TREND ANALYSIS */}
            {(() => {
              // Respect dashboard filters: campus, month, quarter
              let availableMonths = MONTHS.filter(m => data.months.includes(m));
              if (month !== 'ALL') { availableMonths = availableMonths.filter(m => m === month); }
              else if (quarter !== 'ALL' && QUARTERS[quarter]) { availableMonths = availableMonths.filter(m => QUARTERS[quarter].includes(m)); }
              const allSrc = ['incidents','v2_ehs_inspection','training','v2_external_compliance'];
              const trendMonths = availableMonths.filter(m => allSrc.some(k => { const rows = data.sources[k]?.rows || []; const filtered = campus !== 'ALL' ? rows.filter(r => r.campus === campus) : rows; return filtered.some(r => r.month === m); }));
              if (trendMonths.length === 0) return null;
              const label = `6-MONTH TREND ANALYSIS (${trendMonths[0].toUpperCase()} \u2013 ${trendMonths[trendMonths.length-1].toUpperCase()})`;

              const TREND_CHARTS: { title: string; sourceKey: string; mode: 'sum'|'pct'; color: string; fillColor: string }[] = [
                { title: 'Total Incidents \u2014 Monthly Trend', sourceKey: 'incidents', mode: 'sum', color: '#dc3545', fillColor: 'rgba(220,53,69,0.12)' },
                { title: 'EHS Inspection Rate % \u2014 Monthly Trend', sourceKey: 'v2_ehs_inspection', mode: 'pct', color: '#198754', fillColor: 'rgba(25,135,84,0.10)' },
                { title: 'Training Hours \u2014 Monthly Trend', sourceKey: 'training', mode: 'sum', color: '#4A90D9', fillColor: 'rgba(74,144,217,0.12)' },
                { title: 'Compliance Rate % \u2014 Monthly Trend', sourceKey: 'v2_external_compliance', mode: 'pct', color: '#1A1F71', fillColor: 'rgba(26,31,113,0.10)' },
              ];

              return (<>
                <h3 className="section-title">{label}</h3>
                <div className="trend-grid">
                  {TREND_CHARTS.map(tc => {
                    const allRows = data.sources[tc.sourceKey]?.rows || [];
                    const srcRows = campus !== 'ALL' ? allRows.filter(r => r.campus === campus) : allRows;
                    const pts = trendMonths.map(m => {
                      const mRows = srcRows.filter(r => r.month === m);
                      if (tc.mode === 'sum') {
                        return Math.round(mRows.reduce((s, r) => s + (r.value || r.actual || 0), 0) * 10) / 10;
                      } else {
                        const p = mRows.reduce((s, r) => s + r.planned, 0);
                        const a = mRows.reduce((s, r) => s + r.actual, 0);
                        return p > 0 ? Math.round(a / p * 100) : 0;
                      }
                    });
                    const opts: Highcharts.Options = {
                      chart: { type: 'areaspline', height: 220, style: { fontFamily: 'var(--font)' } },
                      title: { text: tc.title, style: { fontSize: '13px', fontWeight: '700', fontStyle: 'italic' } },
                      xAxis: { categories: trendMonths, labels: { style: { fontSize: '10px' } } },
                      yAxis: { title: { text: '' }, labels: { style: { fontSize: '10px' } }, min: 0 },
                      legend: { enabled: false },
                      credits: { enabled: false },
                      plotOptions: { areaspline: { fillColor: tc.fillColor, marker: { enabled: true, radius: 4, fillColor: tc.color }, lineColor: tc.color, lineWidth: 2, dataLabels: { enabled: true, style: { fontSize: '10px', fontWeight: '600' } } } },
                      series: [{ type: 'areaspline', name: tc.title, data: pts, color: tc.color }],
                    };
                    return (
                      <div key={tc.sourceKey} className="chart-card">
                        <div className="chart-card-body">
                          <HighchartsReact highcharts={Highcharts} options={opts} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </>);
            })()}

            {data.errors.length > 0 && (
              <div className="sync-errors">
                <h4>Sync Errors ({data.errors.length})</h4>
                <ul>{data.errors.map((e, i) => <li key={i}>{e}</li>)}</ul>
              </div>
            )}

            {/* FOOTER */}
            <div className="dashboard-footer">
              Data sourced from Smartsheet {'â¢'} Last synced: {data.syncedAt ? new Date(data.syncedAt).toLocaleString() : 'N/A'} {'â¢'} Click Sync Now to force reload
            </div>
          </>
        )}
      </div>
    </div>
  );
}
