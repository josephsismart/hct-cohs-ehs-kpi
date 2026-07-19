'use client';
import { useState, useMemo, useCallback } from 'react';
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
  v2_permit_to_work: 'https://app.smartsheet.com/reports/MgpHGX276R7R7pxw4jx8X5KccX3PQRHHG7p62hc1',
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
    chart: { type: 'pie', height: 320, style: { fontFamily: "'Segoe UI', Arial, sans-serif" } },
    title: { text: undefined },
    plotOptions: { pie: { dataLabels: { enabled: true, format: '<b>{point.name}</b>: {point.y}', style: { fontSize: '10px' } }, showInLegend: true } },
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
    tooltip: { shared: true },
  };
  return <HighchartsReact highcharts={Highcharts} options={options} />;
}

function KpiLineChart({ rows, chartDef }: { rows: KpiRow[]; chartDef: ChartDef }) {
  const byCampus = aggregateByCampus(rows);
  const campuses = Object.keys(byCampus).sort();
  if (campuses.length === 0) return <div className="no-data">No data available</div>;
  const options: Highcharts.Options = {
    chart: { type: 'line', height: 280, style: { fontFamily: "'Segoe UI', Arial, sans-serif" } },
    title: { text: undefined },
    xAxis: { categories: campuses, labels: { style: { fontSize: '10px' } } },
    yAxis: { title: { text: null }, gridLineColor: '#f0f0f0' },
    legend: { align: 'center', verticalAlign: 'bottom', itemStyle: { fontSize: '10px' } },
    series: [
      { type: 'line', name: chartDef.plannedLabel || 'Planned / Target', data: campuses.map(c => byCampus[c].planned), color: 'rgba(74,144,217,0.4)', dashStyle: 'Dash' },
      { type: 'line', name: chartDef.actualLabel || 'Actual', data: campuses.map(c => byCampus[c].actual), color: '#0a3d62' },
    ],
    credits: { enabled: false },
    tooltip: { shared: true },
  };
  return <HighchartsReact highcharts={Highcharts} options={options} />;
}

function KpiStackedBarChart({ rows }: { rows: KpiRow[] }) {
  const byCampus = aggregateByCampus(rows);
  const campuses = Object.keys(byCampus).sort();
  if (campuses.length === 0) return <div className="no-data">No data available</div>;
  const options: Highcharts.Options = {
    chart: { type: 'bar', height: 280, style: { fontFamily: "'Segoe UI', Arial, sans-serif" } },
    title: { text: undefined },
    xAxis: { categories: campuses, labels: { style: { fontSize: '10px' } } },
    yAxis: { title: { text: null }, stackLabels: { enabled: true }, gridLineColor: '#f0f0f0' },
    legend: { align: 'center', verticalAlign: 'bottom', itemStyle: { fontSize: '10px' } },
    plotOptions: {
      series: { stacking: 'normal', dataLabels: { enabled: true, format: '{y}', style: { fontSize: '9px', fontWeight: 'normal' } } },
    },
    series: [
      { type: 'bar', name: 'Closed', data: campuses.map(c => byCampus[c].actual), color: '#1D9E75' },
      { type: 'bar', name: 'Open', data: campuses.map(c => Math.max(0, byCampus[c].planned - byCampus[c].actual)), color: '#EA352E' },
    ],
    credits: { enabled: false },
    tooltip: { shared: true },
  };
  return <HighchartsReact highcharts={Highcharts} options={options} />;
}

// ----------------------------------------------------------
// Layout & Dashboard
// ----------------------------------------------------------

function KpiCard({ chartDef, rows, smartsheetLink }: { chartDef: ChartDef; rows: KpiRow[]; smartsheetLink?: string }) {
  const ChartComponent = chartDef.chartType === 'line' ? KpiLineChart
    : chartDef.chartType === 'bar' ? KpiBarChart
    : chartDef.chartType === 'rate' ? KpiRateChart
    : chartDef.chartType === 'value_hours' ? KpiValueHoursChart
    : chartDef.chartType === 'pie' ? KpiPieChart
    : chartDef.chartType === 'stacked_bar' ? KpiStackedBarChart
    : KpiBarChart;
  return (
    <div className="kpi-card">
      <div className="kpi-card-header">
        <h3 className="kpi-card-title">{chartDef.title}</h3>
        <div className="kpi-card-actions">
          {smartsheetLink && (
            <a href={smartsheetLink} target="_blank" rel="noopener noreferrer" className="btn-report">
              View in Smartsheet
            </a>
          )}
        </div>
      </div>
      <div className="kpi-card-body">
        <ChartComponent rows={rows} chartDef={chartDef} />
      </div>
    </div>
  );
}

function ExtraCard({ chartDef, rows, smartsheetLink }: { chartDef: ChartDef; rows: KpiRow[]; smartsheetLink?: string }) {
  return <KpiCard chartDef={chartDef} rows={rows} smartsheetLink={smartsheetLink} />;
}

function MonthFilter({ months, selected, onChange }: { months: string[]; selected: string; onChange: (m: string) => void }) {
  return (
    <div className="month-filter">
      <label>Month:</label>
      <select value={selected} onChange={e => onChange(e.target.value)}>
        <option value="">All Months</option>
        {months.map(m => <option key={m} value={m}>{m}</option>)}
      </select>
    </div>
  );
}

export default function Dashboard() {
  const [data, setData] = useState<Record<string, KpiRow[]>>({});
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [selectedMonth, setSelectedMonth] = useState('');
  const [availableMonths, setAvailableMonths] = useState<string[]>([]);

  useEffect(() => {
    fetch('/api/sync')
      .then(r => r.json())
      .then((res: any) => {
        if (res.error) { setError(res.error); setLoading(false); return; }
        const parsed: Record<string, KpiRow[]> = {};
        const allMonths = new Set<string>();
        for (const [key, rows] of Object.entries(res.data || {})) {
          parsed[key] = (rows as any[]).map((r: any) => ({
            campus: r.campus || '',
            month: r.month || null,
            planned: r.planned || 0,
            actual: r.actual || 0,
            value: r.value || 0,
          }));
          parsed[key].forEach(r => { if (r.month) allMonths.add(r.month); });
        }
        const MONTH_ORDER = ['January','February','March','April','May','June','July','August','September','October','November','December'];
        const sorted = [...allMonths].sort((a, b) => MONTH_ORDER.indexOf(a) - MONTH_ORDER.indexOf(b));
        setAvailableMonths(sorted);
        setData(parsed);
        setLoading(false);
      })
      .catch(e => { setError(e.message); setLoading(false); });
  }, []);

  const filtered = useMemo(() => {
    if (!selectedMonth) return data;
    const out: Record<string, KpiRow[]> = {};
    for (const [key, rows] of Object.entries(data)) {
      out[key] = rows.filter(r => !r.month || r.month === selectedMonth);
    }
    return out;
  }, [data, selectedMonth]);

  if (loading) return (
    <div className="loading-container">
      <div className="spinner" />
      <p>Loading dashboard data...</p>
    </div>
  );
  if (error) return (
    <div className="error-container">
      <p>Error: {error}</p>
      <button className="btn-report" onClick={() => window.location.reload()}>Retry</button>
    </div>
  );

  return (
    <main className="dashboard">
      <header className="dashboard-header">
        <div>
          <h1>HCT-COHS EHS KPI Dashboard</h1>
          <p className="subtitle">Environmental Health &amp; Safety Key Performance Indicators</p>
        </div>
        <MonthFilter months={availableMonths} selected={selectedMonth} onChange={setSelectedMonth} />
      </header>

      <section className="chart-grid">
        {KPI_CHARTS.map(ch => (
          <KpiCard
            key={ch.key}
            chartDef={ch}
            rows={filtered[ch.key] || []}
            smartsheetLink={SMARTSHEET_LINKS[ch.key]}
          />
        ))}
      </section>

      <section className="extra-section">
        <h2 className="section-title">Additional Reports</h2>
        <div className="chart-grid">
          {EXTRA_CHARTS.map(ch => (
            <ExtraCard
              key={ch.key}
              chartDef={ch}
              rows={filtered[ch.key] || []}
              smartsheetLink={SMARTSHEET_LINKS[ch.key]}
            />
          ))}
        </div>
      </section>
    </main>
  );
}
ows} />
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
            <div className="placeholder-section">
              <h4>WASTE SEGREGATION</h4>
              <p>Data table coming soon</p>
            </div>

            {/* EXECUTIVE KPI SUMMARY */}
            <div className="placeholder-section">
              <h4>EXECUTIVE KPI SUMMARY</h4>
              <p>Summary table coming soon</p>
            </div>

            {/* 6-MONTH TREND ANALYSIS */}
            <h3 className="section-title">6-MONTH TREND ANALYSIS</h3>
            <div className="trend-grid">
              {['Inspection Trends', 'Findings Trends', 'Risk Assessment Trends', 'Compliance Trends'].map(t => (
                <div key={t} className="placeholder-section">
                  <h4>{t}</h4>
                  <p>Coming soon</p>
                </div>
              ))}
            </div>

            {data.errors.length > 0 && (
              <div className="sync-errors">
                <h4>Sync Errors ({data.errors.length})</h4>
                <ul>{data.errors.map((e, i) => <li key={i}>{e}</li>)}</ul>
              </div>
            )}

            {/* FOOTER */}
            <div className="dashboard-footer">
              Data sourced from Smartsheet {'•'} Cached 5 min {'•'} Click Sync Now to force reload
            </div>
          </>
        )}
      </div>
    </div>
  );
}
