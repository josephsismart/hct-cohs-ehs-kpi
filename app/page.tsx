'use client';
import { useState, useMemo, useCallback } from 'react';
import Highcharts from 'highcharts';
import HighchartsReact from 'highcharts-react-official';

const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const QUARTERS: Record<string, string[]> = {
  Q1: ['January','February','March'], Q2: ['April','May','June'],
  Q3: ['July','August','September'], Q4: ['October','November','December'],
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

const KPI_CHARTS = [
  { key: 'v2_ehs_inspection', label: 'Authority Compliance Rate', plannedLabel: 'Applicable Compliance', actualLabel: 'Actual Compliance', type: 'planned_actual' },
  { key: 'v2_hs_kpi_report', label: 'HS KPI Report', valueLabel: 'Submitted', type: 'value' },
  { key: 'v2_permit_to_work', label: 'Permit to Work', plannedLabel: 'No. of PTWs Issued', actualLabel: 'Total Work Registered', type: 'planned_actual' },
  { key: 'v2_hs_committee', label: 'EHS Committee Meeting', valueLabel: 'No. of Committee Meeting', type: 'value' },
  { key: 'v2_findings_on_time', label: 'Findings Closed On Time', plannedLabel: 'Findings in Reporting Month', actualLabel: 'Findings Due Met/Exceeded', belowLabel: 'Findings Due Below Target', type: 'planned_actual_below' },
  { key: 'v2_risk_closed', label: 'Risk Assessment Closed', plannedLabel: 'Total Risk Assessments', actualLabel: 'Risk Assessment Closed', type: 'planned_actual' },
  { key: 'v2_risk_validated', label: 'Risk Assessment Validated & Signed Off', plannedLabel: 'Total Assessments Register', actualLabel: 'RA Validated and Signed Off', type: 'planned_actual' },
  { key: 'v2_external_compliance', label: 'External Compliance', plannedLabel: 'Applicable Compliance', actualLabel: 'Compliance Met/Exceeded', belowLabel: 'Compliance Below Target', type: 'planned_actual_below' },
  { key: 'v2_safe_working', label: 'Safe Working Procedure', plannedLabel: 'No. of SOPs Verified', actualLabel: 'No. of SOPs Implemented', type: 'planned_actual' },
  { key: 'v2_hazard_id', label: 'Hazard Identification', plannedLabel: 'Total Control Sampled', actualLabel: 'Implemented Controls', type: 'planned_actual' },
  { key: 'v2_onsite_induction', label: 'Onsite Induction', plannedLabel: 'New Contractors', actualLabel: 'Contractors Inducted', type: 'planned_actual' },
  { key: 'v2_contractor_activity', label: 'Contractor Activity', plannedLabel: 'Registered Contractors', actualLabel: 'Active Contractors', type: 'planned_actual' },
  { key: 'v2_investigation_on_time', label: 'Investigation Completed on Time', plannedLabel: 'Total Incident', actualLabel: 'Investigation Completed on Time', type: 'planned_actual' },
  { key: 'notification', label: 'Notification on Time', plannedLabel: 'Total Incident', actualLabel: 'Notification Submitted on Time', type: 'planned_actual' },
];

interface KpiRow { campus: string; month: string | null; planned: number; actual: number; value: number; }
interface SyncData {
  syncedAt: string;
  sources: Record<string, { rows: KpiRow[]; error?: string }>;
  campuses: string[];
  months: string[];
  errors: string[];
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
      color: '#2196F3',
    });
  } else if (chartDef.type === 'planned_actual_below') {
    series.push(
      { type: 'column', name: chartDef.plannedLabel || 'Planned', data: campuses.map(c => byCampus[c].planned), color: 'rgba(13,110,253,0.4)' },
      { type: 'column', name: chartDef.actualLabel || 'Met/Exceeded', data: campuses.map(c => Math.max(0, byCampus[c].actual)), color: '#198754' },
      { type: 'column', name: chartDef.belowLabel || 'Below Target', data: campuses.map(c => Math.max(0, byCampus[c].planned - byCampus[c].actual)), color: '#dc3545' },
    );
  } else {
    series.push(
      { type: 'column', name: chartDef.plannedLabel || 'Planned / Target', data: campuses.map(c => byCampus[c].planned), color: 'rgba(13,110,253,0.4)' },
      { type: 'column', name: chartDef.actualLabel || 'Actual', data: campuses.map(c => byCampus[c].actual), color: '#0a3d62' },
    );
  }

  const options: Highcharts.Options = {
    chart: { type: 'column', height: 280, style: { fontFamily: 'Inter, sans-serif' } },
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

export default function Dashboard() {
  const [data, setData] = useState<SyncData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [campus, setCampus] = useState('ALL');
  const [month, setMonth] = useState('ALL');
  const [quarter, setQuarter] = useState('ALL');
  const [year, setYear] = useState('ALL');

  const doSync = useCallback(async () => {
    setLoading(true); setError('');
    try {
      const res = await fetch('/api/sync');
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      setData(await res.json());
    } catch (e: any) { setError(e.message); }
    finally { setLoading(false); }
  }, []);

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
        const pct = agg.planned > 0 ? Math.round(agg.actual / agg.planned * 100) : 0;
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

  const syncLabel = data ? `Synced ${Math.round((Date.now() - new Date(data.syncedAt).getTime()) / 3600000)} hr ago` : '';

  return (
    <div className="dashboard">
      <div className="brand-header">
        <div className="brand-logo brand-ohs">
          <img src="/ohs-logo.png" alt="OHS" onError={(e) => (e.currentTarget.style.display = 'none')} />
        </div>
        <div className="brand-title">HCT - COHS EHS KPI REPORTING</div>
        <div className="brand-logo brand-hct">
          <img src="/hct-logo.png" alt="HCT" onError={(e) => (e.currentTarget.style.display = 'none')} />
        </div>
      </div>

      <div className="filter-bar">
        <span className="filter-label">Health &amp; Safety Performance Report</span>
        <div className="filter-controls">
          <label>Year
            <select value={year} onChange={e => setYear(e.target.value)}>
              <option value="ALL">All Years</option>
              <option value="2026">2026</option><option value="2025">2025</option>
            </select>
          </label>
          <label>Quarter
            <select value={quarter} onChange={e => { setQuarter(e.target.value); if (e.target.value !== 'ALL') setMonth('ALL'); }}>
              <option value="ALL">All Quarters</option>
              {Object.keys(QUARTERS).map(q => <option key={q} value={q}>{q}</option>)}
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
          <button className="btn-refresh" onClick={() => { setCampus('ALL'); setMonth('ALL'); setQuarter('ALL'); setYear('ALL'); }}>Refresh</button>
          <button className="btn-sync" onClick={doSync} disabled={loading}>
            {loading ? 'Syncing...' : 'Sync Now'}
          </button>
          {data && <span className="sync-badge">{syncLabel}</span>}
        </div>
      </div>

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
              <span className="legend-dot" style={{ background: 'rgba(13,110,253,0.4)' }} /> Planned / Target
              <span className="legend-dot" style={{ background: '#198754' }} /> Actual \u2014 Met or Exceeded
              <span className="legend-dot" style={{ background: '#dc3545' }} /> Actual \u2014 Below Target
              <span className="legend-dot" style={{ background: '#ffc107' }} /> No target set
            </div>

            <h3 className="section-title">KPI CHARTS</h3>
            <div className="charts-grid">
              {KPI_CHARTS.map(chartDef => {
                const rows = getRows(chartDef.key);
                return (
                  <div key={chartDef.key} className="chart-card">
                    <div className="chart-card-header">
                      <span>{chartDef.label}</span>
                      <a className="btn-smartsheet" href="#" onClick={e => e.preventDefault()}>View in Smartsheet</a>
                    </div>
                    <div className="chart-card-body">
                      <KpiBarChart chartDef={chartDef} rows={rows} />
                    </div>
                  </div>
                );
              })}
            </div>

            {data.errors.length > 0 && (
              <div className="sync-errors">
                <h4>Sync Errors ({data.errors.length})</h4>
                <ul>{data.errors.map((e, i) => <li key={i}>{e}</li>)}</ul>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
'use client';
import { useState, useEffect, useMemo, useCallback } from 'react';
import { Chart as ChartJS, CategoryScale, LinearScale, BarElement, ArcElement, Title, Tooltip, Legend } from 'chart.js';
import { Bar, Doughnut } from 'react-chartjs-2';

ChartJS.register(CategoryScale, LinearScale, BarElement, ArcElement, Title, Tooltip, Legend);

const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const QUARTERS: Record<string, string[]> = {
  Q1: ['January','February','March'], Q2: ['April','May','June'],
  Q3: ['July','August','September'], Q4: ['October','November','December'],
};

const KPI_CONFIG = [
  { key: 'drills', label: 'Emergency Drills', icon: 'fa-fire-extinguisher', color: '#dc3545', type: 'planned_actual' },
  { key: 'ehs', label: 'EHS Inspections', icon: 'fa-clipboard-check', color: '#0d6efd', type: 'planned_actual' },
  { key: 'findings', label: 'Findings Closed', icon: 'fa-magnifying-glass', color: '#198754', type: 'planned_actual' },
  { key: 'notification', label: 'Notification on Time', icon: 'fa-bell', color: '#ffc107', type: 'planned_actual' },
  { key: 'risk', label: 'Risk Assessment', icon: 'fa-shield-halved', color: '#6f42c1', type: 'planned_actual' },
  { key: 'training', label: 'Training Hours', icon: 'fa-graduation-cap', color: '#0dcaf0', type: 'value' },
  { key: 'incidents', label: 'Total Incidents', icon: 'fa-triangle-exclamation', color: '#fd7e14', type: 'value' },
];

const V2_KPI_CONFIG = [
  { key: 'v2_hs_committee', label: 'HS Committee Meetings', icon: 'fa-users', color: '#20c997', type: 'value' },
  { key: 'v2_findings_on_time', label: 'Findings On Time (V2)', icon: 'fa-clock', color: '#6610f2', type: 'planned_actual' },
  { key: 'v2_risk_closed', label: 'Risk Closed (V2)', icon: 'fa-lock', color: '#d63384', type: 'planned_actual' },
  { key: 'v2_risk_validated', label: 'Risk Validated (V2)', icon: 'fa-check-double', color: '#6f42c1', type: 'planned_actual' },
  { key: 'v2_ehs_inspection', label: 'EHS Inspection (V2)', icon: 'fa-search', color: '#0d6efd', type: 'planned_actual' },
  { key: 'v2_hs_kpi_report', label: 'HS KPI Report', icon: 'fa-file-lines', color: '#198754', type: 'value' },
  { key: 'v2_external_compliance', label: 'External Compliance', icon: 'fa-gavel', color: '#dc3545', type: 'planned_actual' },
  { key: 'v2_safe_working', label: 'Safe Working (SOPs)', icon: 'fa-hard-hat', color: '#ffc107', type: 'planned_actual' },
  { key: 'v2_permit_to_work', label: 'Permit to Work', icon: 'fa-id-card', color: '#0dcaf0', type: 'planned_actual' },
  { key: 'v2_hazard_id', label: 'Hazard Identification', icon: 'fa-biohazard', color: '#fd7e14', type: 'planned_actual' },
  { key: 'v2_onsite_induction', label: 'Onsite Induction', icon: 'fa-person-walking', color: '#20c997', type: 'planned_actual' },
  { key: 'v2_investigation_on_time', label: 'Investigation On Time', icon: 'fa-microscope', color: '#6610f2', type: 'planned_actual' },
  { key: 'v2_training_hours', label: 'Training Hours (V2)', icon: 'fa-chalkboard-teacher', color: '#d63384', type: 'value' },
  { key: 'v2_planned_training', label: 'Planned Training (V2)', icon: 'fa-calendar-check', color: '#6f42c1', type: 'planned_actual' },
  { key: 'v2_drills', label: 'Drills (V2)', icon: 'fa-fire', color: '#dc3545', type: 'planned_actual' },
  { key: 'v2_waste_segregation', label: 'Waste Segregation', icon: 'fa-recycle', color: '#198754', type: 'value' },
];

interface KpiRow { campus: string; month: string | null; planned: number; actual: number; value: number; }
interface SyncData { syncedAt: string; sources: Record<string, { rows: KpiRow[]; error?: string }>; campuses: string[]; months: string[]; errors: string[]; }

function filterRows(rows: KpiRow[], campus: string, month: string, quarter: string, year: string): KpiRow[] {
  let filtered = rows;
  if (campus !== 'ALL') filtered = filtered.filter(r => r.campus === campus);
  if (month !== 'ALL') filtered = filtered.filter(r => r.month === month);
  else if (quarter !== 'ALL' && QUARTERS[quarter]) filtered = filtered.filter(r => r.month && QUARTERS[quarter].includes(r.month));
  return filtered;
}

function aggregate(rows: KpiRow[]) {
  return rows.reduce((acc, r) => ({ planned: acc.planned + r.planned, actual: acc.actual + r.actual, value: acc.value + r.value }), { planned: 0, actual: 0, value: 0 });
}

function aggregateByCampus(rows: KpiRow[]) {
  const map: Record<string, { planned: number; actual: number; value: number }> = {};
  rows.forEach(r => { if (!map[r.campus]) map[r.campus] = { planned: 0, actual: 0, value: 0 }; map[r.campus].planned += r.planned; map[r.campus].actual += r.actual; map[r.campus].value += r.value; });
  return map;
}

export default function Dashboard() {
  const [data, setData] = useState<SyncData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [campus, setCampus] = useState('ALL');
  const [month, setMonth] = useState('ALL');
  const [quarter, setQuarter] = useState('ALL');
  const [year, setYear] = useState('ALL');
  const [tab, setTab] = useState<'v1' | 'v2'>('v1');

  const doSync = useCallback(async () => {
    setLoading(true); setError('');
    try { const res = await fetch('/api/sync'); if (!res.ok) throw new Error('HTTP ' + res.status); setData(await res.json()); }
    catch (e: any) { setError(e.message); } finally { setLoading(false); }
  }, []);

  const kpiConfigs = tab === 'v1' ? KPI_CONFIG : V2_KPI_CONFIG;

  const getRows = useCallback((key: string): KpiRow[] => {
    if (!data?.sources[key]) return [];
    const rows = data.sources[key].rows;
    const allNullMonth = rows.every(r => !r.month);
    if (allNullMonth && (month !== 'ALL' || quarter !== 'ALL')) return filterRows(rows, campus, 'ALL', 'ALL', year);
    return filterRows(rows, campus, month, quarter, year);
  }, [data, campus, month, quarter, year]);

  const chartData = useMemo(() => {
    if (!data) return null;
    const configs = tab === 'v1' ? KPI_CONFIG : V2_KPI_CONFIG;
    const barKpi = configs.find(c => c.type === 'planned_actual');
    if (!barKpi) return null;
    const rows = getRows(barKpi.key);
    const byCampus = aggregateByCampus(rows);
    const labels = Object.keys(byCampus).sort();
    return { labels, datasets: [
      { label: 'Planned', data: labels.map(c => byCampus[c].planned), backgroundColor: 'rgba(13,110,253,0.7)' },
      { label: 'Actual', data: labels.map(c => byCampus[c].actual), backgroundColor: 'rgba(25,135,84,0.7)' },
    ], kpiLabel: barKpi.label };
  }, [data, tab, getRows]);

  return (
    <div className="min-vh-100">
      <nav className="navbar navbar-dark bg-dark px-4">
        <span className="navbar-brand fw-bold"><i className="fas fa-chart-line me-2" />HCT EHS Dashboard</span>
        <div className="d-flex align-items-center gap-3">
          {data && <span className="badge bg-success sync-badge"><i className="fas fa-check me-1" />Synced {new Date(data.syncedAt).toLocaleTimeString()}</span>}
          <button className="btn btn-outline-light btn-sm" onClick={doSync} disabled={loading}>
            <i className={'fas fa-sync-alt me-1 ' + (loading ? 'fa-spin' : '')} />{loading ? 'Syncing...' : 'Sync Now'}
          </button>
        </div>
      </nav>
      <div className="container-fluid p-4">
        {error && <div className="alert alert-danger"><i className="fas fa-exclamation-triangle me-2" />{error}</div>}
        {!data && !loading && (
          <div className="text-center py-5">
            <i className="fas fa-database fa-3x text-muted mb-3 d-block" />
            <h4>No data loaded</h4>
            <p className="text-muted">Click Sync Now to fetch data from Smartsheet</p>
            <button className="btn btn-primary btn-lg" onClick={doSync}><i className="fas fa-sync-alt me-2" />Sync Now</button>
          </div>
        )}
        {data && (<>
          <div className="filter-bar p-3 mb-4 shadow-sm">
            <div className="row g-3 align-items-end">
              <div className="col-auto"><label className="form-label fw-semibold"><i className="fas fa-filter me-1" />Filters</label></div>
              <div className="col"><label className="form-label small">Campus</label>
                <select className="form-select form-select-sm" value={campus} onChange={e => setCampus(e.target.value)}>
                  <option value="ALL">All Campuses</option>{data.campuses.map(c => <option key={c} value={c}>{c}</option>)}
                </select></div>
              <div className="col"><label className="form-label small">Month</label>
                <select className="form-select form-select-sm" value={month} onChange={e => { setMonth(e.target.value); if (e.target.value !== 'ALL') setQuarter('ALL'); }}>
                  <option value="ALL">All Months</option>{MONTHS.map(m => <option key={m} value={m}>{m}</option>)}
                </select></div>
              <div className="col"><label className="form-label small">Quarter</label>
                <select className="form-select form-select-sm" value={quarter} onChange={e => { setQuarter(e.target.value); if (e.target.value !== 'ALL') setMonth('ALL'); }}>
                  <option value="ALL">All Quarters</option>{Object.keys(QUARTERS).map(q => <option key={q} value={q}>{q}</option>)}
                </select></div>
              <div className="col-auto"><button className="btn btn-outline-secondary btn-sm" onClick={() => { setCampus('ALL'); setMonth('ALL'); setQuarter('ALL'); setYear('ALL'); }}><i className="fas fa-times me-1" />Clear</button></div>
            </div>
          </div>
          <ul className="nav nav-tabs mb-4">
            <li className="nav-item"><button className={'nav-link ' + (tab === 'v1' ? 'active' : '')} onClick={() => setTab('v1')}><i className="fas fa-chart-bar me-1" />Original KPIs</button></li>
            <li className="nav-item"><button className={'nav-link ' + (tab === 'v2' ? 'active' : '')} onClick={() => setTab('v2')}><i className="fas fa-chart-pie me-1" />V2 KPIs</button></li>
          </ul>
          <div className="row g-3 mb-4">
            {kpiConfigs.map(kpi => {
              const rows = getRows(kpi.key); const agg = aggregate(rows); const srcError = data.sources[kpi.key]?.error;
              return (<div key={kpi.key} className="col-6 col-md-4 col-lg-3"><div className="card kpi-card h-100" style={{ borderLeftColor: kpi.color }}><div className="card-body">
                <div className="d-flex justify-content-between align-items-start"><span className="label text-uppercase">{kpi.label}</span><i className={'fas ' + kpi.icon} style={{ color: kpi.color, fontSize: '1.2rem' }} /></div>
                {srcError ? <div className="text-danger small mt-2"><i className="fas fa-exclamation-circle me-1" />{srcError}</div>
                : kpi.type === 'planned_actual' ? (<><div className="value mt-2" style={{ color: kpi.color }}>{agg.actual}</div>
                  <div className="small text-muted">of {agg.planned} planned ({agg.planned > 0 ? Math.round(agg.actual / agg.planned * 100) : 0}%)</div>
                  <div className="progress mt-2" style={{ height: '6px' }}><div className="progress-bar" style={{ width: (agg.planned > 0 ? Math.min(100, agg.actual / agg.planned * 100) : 0) + '%', backgroundColor: kpi.color }} /></div></>)
                : <div className="value mt-2" style={{ color: kpi.color }}>{agg.value || agg.actual || agg.planned}</div>}
                <div className="small text-muted mt-1">{rows.length} records</div>
              </div></div></div>);
            })}
          </div>
          <div className="row g-4 mb-4">
            <div className="col-lg-8"><div className="card shadow-sm"><div className="card-header bg-white fw-semibold"><i className="fas fa-chart-bar me-2" />{chartData?.kpiLabel || 'KPI'} — Campus Breakdown</div>
              <div className="card-body chart-container">{chartData && <Bar data={chartData} options={{ responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'top' } }, scales: { y: { beginAtZero: true } } }} />}</div></div></div>
            <div className="col-lg-4"><div className="card shadow-sm"><div className="card-header bg-white fw-semibold"><i className="fas fa-chart-pie me-2" />Overview</div>
              <div className="card-body chart-container d-flex align-items-center justify-content-center">{chartData && <Doughnut data={{ labels: ['Actual', 'Gap'], datasets: [{ data: [chartData.datasets[1].data.reduce((a, b) => a + b, 0), Math.max(0, chartData.datasets[0].data.reduce((a, b) => a + b, 0) - chartData.datasets[1].data.reduce((a, b) => a + b, 0))], backgroundColor: ['rgba(25,135,84,0.8)', 'rgba(220,53,69,0.3)'] }] }} options={{ responsive: true, plugins: { legend: { position: 'bottom' } } }} />}</div></div></div>
          </div>
          <div className="card shadow-sm mb-4"><div className="card-header bg-white fw-semibold"><i className="fas fa-table me-2" />Detailed Data</div>
            <div className="card-body p-0"><div className="table-responsive"><table className="table table-hover table-striped mb-0"><thead className="table-dark"><tr><th>KPI</th><th>Campus</th><th>Month</th><th className="text-end">Planned</th><th className="text-end">Actual</th><th className="text-end">Value</th></tr></thead>
              <tbody>{kpiConfigs.flatMap(kpi => { const rows = getRows(kpi.key); return rows.slice(0, 50).map((r, i) => (<tr key={kpi.key + '-' + i}>{i === 0 && <td rowSpan={Math.min(rows.length, 50)} className="fw-semibold" style={{ verticalAlign: 'top' }}>{kpi.label}</td>}<td>{r.campus}</td><td>{r.month || '\u2014'}</td><td className="text-end">{r.planned || '\u2014'}</td><td className="text-end">{r.actual || '\u2014'}</td><td className="text-end">{r.value || '\u2014'}</td></tr>)); })}</tbody></table></div></div></div>
          {data.errors.length > 0 && (<div className="card border-warning mb-4"><div className="card-header bg-warning text-dark fw-semibold"><i className="fas fa-exclamation-triangle me-2" />Sync Errors ({data.errors.length})</div>
            <div className="card-body"><ul className="mb-0 small">{data.errors.map((e, i) => <li key={i}>{e}</li>)}</ul></div></div>)}
        </>)}
      </div>
    </div>
  );
}
