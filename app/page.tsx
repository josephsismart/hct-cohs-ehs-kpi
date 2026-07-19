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
  { key: 'v2_findings_on_time', label: 'Findings Closed On Time', plannedLabel: 'No. of Findings in Reporting Month', actualLabel: 'No. of Findings Due — Met/Exceeded', belowLabel: 'No. of Findings Due — Below Target', type: 'planned_actual_below' },
  { key: 'v2_risk_closed', label: 'Risk Assessment Closed', plannedLabel: 'Total Risk Assessments', actualLabel: 'Risk Assessment Closed', type: 'planned_actual' },
  { key: 'v2_risk_validated', label: 'Risk Assessment Validated & Signed Off', plannedLabel: 'Total Assessments Register', actualLabel: 'RA Validated and Signed Off', type: 'planned_actual' },
  { key: 'v2_external_compliance', label: 'External Compliance', plannedLabel: 'Applicable Compliance', actualLabel: 'Actual Compliance — Met/Exceeded', belowLabel: 'Actual Compliance — Below Target', type: 'planned_actual_below' },
  { key: 'v2_safe_working', label: 'Safe Working Procedure', plannedLabel: 'No. of SOPs Verified', actualLabel: 'No. of SOPs Implemented', type: 'planned_actual' },
  { key: 'v2_hazard_id', label: 'Hazard Identification', plannedLabel: 'Total Control Sampled', actualLabel: 'Implemented Controls', type: 'planned_actual' },
  { key: 'v2_onsite_induction', label: 'Onsite Induction', plannedLabel: 'New Contractors', actualLabel: 'Contractors Inducted', type: 'planned_actual' },
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
              <span className="legend-dot" style={{ background: 'rgba(13,110,253,0.4)' }} /> Planned / Target
              <span className="legend-dot" style={{ background: '#198754' }} /> Actual — Met or Exceeded
              <span className="legend-dot" style={{ background: '#dc3545' }} /> Actual — Below Target
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
