import React, { useState } from 'react';
import { quotationsAPI, APIError } from '../api/client';
import { formatINR, formatNumber, formatWeight, formatLength } from '../utils/format';

export function CalculationResultsView({ quotationId, snapshot, onNavigate, onEditQuotation, showToast }) {
  const [exporting, setExporting] = useState(false);

  if (!snapshot) {
    return (
      <div className="results-container">
        <div className="alert alert-error">
          <p>No calculation snapshot available. Please run calculation from the quotation editor.</p>
          <button type="button" className="btn btn-outline btn-sm" onClick={() => onNavigate('dashboard')}>
            Back to Dashboard
          </button>
        </div>
      </div>
    );
  }

  const handleExport = async () => {
    try {
      setExporting(true);
      showToast?.('Generating commercial Excel quotation workbook...', 'info');
      const filename = await quotationsAPI.exportExcel(quotationId);
      showToast?.(`Downloaded ${filename} successfully`, 'success');
    } catch (err) {
      console.error('Export error:', err);
      const msg = err instanceof APIError ? err.message : 'Excel export failed.';
      showToast?.(msg, 'error');
    } finally {
      setExporting(false);
    }
  };

  const lines = snapshot.results_snapshot?.lines || [];
  const weighted = snapshot.results_snapshot?.weighted || null;

  return (
    <div className="results-container">
      {/* Top Bar */}
      <div className="results-top-bar">
        <div className="results-header-text">
          <div className="badge-row">
            <span className={`badge-mode mode-${(snapshot.calculation_mode || 'reference').toLowerCase()}`}>
              Mode: {snapshot.calculation_mode}
            </span>
            <span className="version-tag">Snapshot v{snapshot.version}</span>
            <span className="status-badge status-calculated">Authoritative Engine</span>
          </div>
          <h1 className="results-title">Commercial Calculation Results</h1>
          <p className="results-subtitle">
            Calculated via authoritative Python calculation engine with exact commercial parity.
          </p>
        </div>

        <div className="results-actions">
          <button
            type="button"
            className="btn btn-outline"
            onClick={() => onNavigate('dashboard')}
          >
            Dashboard
          </button>

          <button
            type="button"
            className="btn btn-outline"
            onClick={() => onEditQuotation(quotationId)}
          >
            Edit Quotation
          </button>

          <button
            type="button"
            className="btn btn-primary btn-export-highlight"
            onClick={handleExport}
            disabled={exporting}
          >
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="icon">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
              <polyline points="7 10 12 15 17 10" />
              <line x1="12" y1="15" x2="12" y2="3" />
            </svg>
            {exporting ? 'Exporting...' : 'Export Commercial Excel (.xlsx)'}
          </button>
        </div>
      </div>

      {/* Main KPI Summary Cards */}
      <div className="results-kpi-grid">
        <div className="kpi-card highlight-card">
          <span className="kpi-label">Grand Total (Incl. GST)</span>
          <span className="kpi-value grand-total-value">{formatINR(snapshot.grand_total, 2)}</span>
          <span className="kpi-sub">Final commercial landed billing</span>
        </div>

        <div className="kpi-card">
          <span className="kpi-label">Subtotal (Excl. GST)</span>
          <span className="kpi-value">{formatINR(snapshot.subtotal_excl_gst, 2)}</span>
          <span className="kpi-sub">Commercial basic value</span>
        </div>

        <div className="kpi-card">
          <span className="kpi-label">GST Tax Amount</span>
          <span className="kpi-value">{formatINR(snapshot.gst_amount, 2)}</span>
          <span className="kpi-sub">GST @ 18%</span>
        </div>

        <div className="kpi-card">
          <span className="kpi-label">Total Volume</span>
          <span className="kpi-value">{formatLength(snapshot.total_metres)}</span>
          <span className="kpi-sub">{formatWeight(snapshot.total_mt)} total mass</span>
        </div>

        <div className="kpi-card">
          <span className="kpi-label">Logistics Trailers</span>
          <span className="kpi-value">{formatNumber(snapshot.total_trips, 2)} Trips</span>
          <span className="kpi-sub">Pooled trailer allocation</span>
        </div>
      </div>

      {/* Line-by-Line Breakdown Table */}
      <div className="results-table-card">
        <div className="table-card-header">
          <h2 className="section-title">Line-by-Line Commercial Breakdown ({lines.length} items)</h2>
        </div>

        <div className="table-wrapper">
          <table className="data-table">
            <thead>
              <tr>
                <th>#</th>
                <th>OD / ID (mm)</th>
                <th>Mass (kg/m)</th>
                <th>Metres</th>
                <th>Weight (MT)</th>
                <th>Trips</th>
                <th>Bare / m</th>
                <th>Ext / m</th>
                <th>Int / m</th>
                <th>Freight / m</th>
                <th>Rate / m</th>
                <th>Rate / MT</th>
                <th className="text-right">Line Value</th>
              </tr>
            </thead>
            <tbody>
              {lines.map((item, idx) => (
                <tr key={idx}>
                  <td>
                    <span className="line-number-pill">#{idx + 1}</span>
                  </td>
                  <td>
                    <div className="dimension-cell">
                      <span>OD: {formatNumber(item.od_mm, 1)} mm</span>
                      <span className="sub-dim">ID: {formatNumber(item.inner_d_mm, 1)} mm</span>
                    </div>
                  </td>
                  <td>{formatNumber(item.mass_kg_per_m, 2)}</td>
                  <td>{formatLength(item.metres)}</td>
                  <td>{formatWeight(item.weight_mt)}</td>
                  <td>{formatNumber(item.freight_trips, 2)}</td>
                  <td>₹{formatNumber(item.bare_per_m, 2)}</td>
                  <td>₹{formatNumber(item.external_per_m, 2)}</td>
                  <td>₹{formatNumber(item.internal_per_m, 2)}</td>
                  <td>₹{formatNumber(item.freight_per_m, 2)}</td>
                  <td className="font-semibold">₹{formatNumber(item.total_per_m, 2)}</td>
                  <td className="font-semibold">₹{formatNumber(item.total_per_mt, 2)}</td>
                  <td className="text-right font-bold line-val-cell">
                    {formatINR(item.line_value, 2)}
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr className="table-total-row">
                <td colSpan={3}>
                  <strong>Total Quotation Summary</strong>
                </td>
                <td>
                  <strong>{formatLength(snapshot.total_metres)}</strong>
                </td>
                <td>
                  <strong>{formatWeight(snapshot.total_mt)}</strong>
                </td>
                <td>
                  <strong>{formatNumber(snapshot.total_trips, 2)}</strong>
                </td>
                <td colSpan={6} className="text-right">
                  <div className="footer-summary-group">
                    <span className="subtotal-line">
                      Subtotal: <strong>{formatINR(snapshot.subtotal_excl_gst, 2)}</strong>
                    </span>
                    <span className="grandtotal-line">
                      Grand Total: <strong>{formatINR(snapshot.grand_total, 2)}</strong>
                    </span>
                  </div>
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      {/* Weighted Commercial Summary (if present) */}
      {weighted && (
        <div className="weighted-summary-card">
          <h3 className="section-title">Weighted Cost Sheet Summary (₹ / MT)</h3>
          <div className="weighted-metrics-grid">
            <div className="weighted-item">
              <span className="weighted-label">Weighted Bare Pipe</span>
              <span className="weighted-val">₹{formatNumber(weighted.bare_pipe_cost_per_mt, 2)} / MT</span>
            </div>
            <div className="weighted-item">
              <span className="weighted-label">Weighted External Coating</span>
              <span className="weighted-val">₹{formatNumber(weighted.external_cost_per_mt, 2)} / MT</span>
            </div>
            <div className="weighted-item">
              <span className="weighted-label">Weighted Internal Coating</span>
              <span className="weighted-val">₹{formatNumber(weighted.internal_cost_per_mt, 2)} / MT</span>
            </div>
            <div className="weighted-item">
              <span className="weighted-label">Weighted Freight Cost</span>
              <span className="weighted-val">₹{formatNumber(weighted.freight_per_mt, 2)} / MT</span>
            </div>
            <div className="weighted-item">
              <span className="weighted-label">Weighted Financial Cost</span>
              <span className="weighted-val">₹{formatNumber(weighted.financial_cost_per_mt, 2)} / MT</span>
            </div>
            <div className="weighted-item highlight-weighted">
              <span className="weighted-label">Effective Blended Selling Price</span>
              <span className="weighted-val">₹{formatNumber(weighted.total_per_mt, 2)} / MT</span>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
