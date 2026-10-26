import React, { useState, useEffect } from 'react';
import { quotationsAPI, APIError } from '../api/client';
import { formatNumber } from '../utils/format';

const DEFAULT_LINE = {
  line_number: 1,
  size_mm: 508,
  thickness_mm: 6.4,
  basis: 'ID',
  quantity: 1000,
  unit: 'Metres',
  pipe_length_m: 12.0,
  pipes_per_trailer: 4.0,
  has_external: true,
  has_internal: true,
  steel_rate: 61000.0,
  grade_extra: 0.0,
  yield_pct: 2.0,
  conversion: 5000.0,
  testing: 0.0,
  margin: 0.0,
  ext_rate: 500.0,
  int_rate: 275.0,
  handling: 100.0,
  financial_pct: 1.5,
  gst_pct: 18.0,
  manual_fc: 0.0,
  manual_selling_price: 0.0,
  override_reason: '',
};

export function QuotationEditorView({ quotationId, onNavigate, onViewResults, showToast }) {
  const isEditing = Boolean(quotationId);

  // Form State
  const [enquiryNo, setEnquiryNo] = useState('');
  const [customerName, setCustomerName] = useState('');
  const [projectName, setProjectName] = useState('');
  const [deliveryLocation, setDeliveryLocation] = useState('');
  const [distanceKm, setDistanceKm] = useState(0);
  const [tripRate, setTripRate] = useState(98000);
  const [calculationMode, setCalculationMode] = useState('REFERENCE');
  const [lines, setLines] = useState([{ ...DEFAULT_LINE }]);
  const [expandedLineIndex, setExpandedLineIndex] = useState(null);

  // Status & Metadata
  const [status, setStatus] = useState('DRAFT');
  const [version, setVersion] = useState(1);
  const [loading, setLoading] = useState(isEditing);
  const [saving, setSaving] = useState(false);
  const [calculating, setCalculating] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [formErrors, setFormErrors] = useState({});

  // Fetch existing quotation if editing
  useEffect(() => {
    if (!quotationId) {
      // Pre-fill enquiry number for new quote with timestamp suffix
      const now = new Date();
      const code = `ENQ-${now.getFullYear().toString().slice(-2)}${(now.getMonth() + 1).toString().padStart(2, '0')}-${Math.floor(1000 + Math.random() * 9000)}`;
      setEnquiryNo(code);
      return;
    }

    const loadQuotation = async () => {
      try {
        setLoading(true);
        const q = await quotationsAPI.get(quotationId);
        setEnquiryNo(q.enquiry_no || '');
        setCustomerName(q.customer_name || '');
        setProjectName(q.project_name || '');
        setDeliveryLocation(q.delivery_location || '');
        setDistanceKm(q.distance_km || 0);
        setTripRate(q.trip_rate || 98000);
        setCalculationMode(q.calculation_mode || 'REFERENCE');
        setStatus(q.status || 'DRAFT');
        setVersion(q.version || 1);

        if (q.lines && q.lines.length > 0) {
          // Normalize line numbers
          const normalizedLines = q.lines.map((l, idx) => ({
            ...DEFAULT_LINE,
            ...l,
            line_number: idx + 1,
          }));
          setLines(normalizedLines);
        }
      } catch (err) {
        console.error('Failed to load quotation:', err);
        const msg = err instanceof APIError ? err.message : 'Failed to load quotation from server.';
        showToast?.(msg, 'error');
        onNavigate('dashboard');
      } finally {
        setLoading(false);
      }
    };

    loadQuotation();
  }, [quotationId]);

  // Line item manipulation
  const handleAddLine = () => {
    const nextLineNum = lines.length + 1;
    const newLine = {
      ...DEFAULT_LINE,
      line_number: nextLineNum,
    };
    setLines([...lines, newLine]);
    setExpandedLineIndex(lines.length); // auto expand the new line
  };

  const handleDuplicateLine = (index) => {
    const source = lines[index];
    const duplicated = {
      ...source,
      line_number: lines.length + 1,
    };
    setLines([...lines, duplicated]);
    showToast?.(`Line #${index + 1} duplicated`, 'info');
  };

  const handleRemoveLine = (index) => {
    if (lines.length <= 1) {
      showToast?.('Quotation must contain at least one line item', 'warning');
      return;
    }
    const updated = lines.filter((_, idx) => idx !== index).map((l, idx) => ({
      ...l,
      line_number: idx + 1,
    }));
    setLines(updated);
    if (expandedLineIndex === index) {
      setExpandedLineIndex(null);
    } else if (expandedLineIndex > index) {
      setExpandedLineIndex(expandedLineIndex - 1);
    }
  };

  const handleLineChange = (index, field, value) => {
    setLines((prev) => {
      const copy = [...prev];
      copy[index] = { ...copy[index], [field]: value };
      return copy;
    });
  };

  const validateForm = () => {
    const errors = {};
    if (!enquiryNo.trim()) errors.enquiryNo = 'Enquiry number is required';
    if (!customerName.trim()) errors.customerName = 'Customer name is required';
    if (lines.length === 0) errors.lines = 'At least one line item is required';

    lines.forEach((l, idx) => {
      if (!l.size_mm || Number(l.size_mm) <= 0) {
        errors[`line_${idx}_size`] = `Line #${idx + 1}: Size (mm) must be positive`;
      }
      if (!l.thickness_mm || Number(l.thickness_mm) <= 0) {
        errors[`line_${idx}_thickness`] = `Line #${idx + 1}: Wall thickness must be positive`;
      }
      if (!l.quantity || Number(l.quantity) <= 0) {
        errors[`line_${idx}_qty`] = `Line #${idx + 1}: Quantity must be positive`;
      }
    });

    setFormErrors(errors);
    return Object.keys(errors).length === 0;
  };

  const preparePayload = () => {
    return {
      enquiry_no: enquiryNo.trim(),
      customer_name: customerName.trim(),
      project_name: projectName.trim() || null,
      delivery_location: deliveryLocation.trim() || null,
      distance_km: Number(distanceKm) || 0,
      trip_rate: Number(tripRate) || 0,
      calculation_mode: calculationMode,
      lines: lines.map((l, idx) => ({
        line_number: idx + 1,
        size_mm: Number(l.size_mm),
        thickness_mm: Number(l.thickness_mm),
        basis: l.basis.toUpperCase(),
        quantity: Number(l.quantity),
        unit: l.unit === 'MT' ? 'MT' : l.unit === 'Nos' ? 'Nos' : 'Metres',
        pipe_length_m: Number(l.pipe_length_m) || 12.0,
        pipes_per_trailer: Number(l.pipes_per_trailer) || 4.0,
        has_external: Boolean(l.has_external),
        has_internal: Boolean(l.has_internal),
        steel_rate: Number(l.steel_rate) || 0,
        grade_extra: Number(l.grade_extra) || 0,
        yield_pct: Number(l.yield_pct) || 0,
        conversion: Number(l.conversion) || 0,
        testing: Number(l.testing) || 0,
        margin: Number(l.margin) || 0,
        ext_rate: Number(l.ext_rate) || 0,
        int_rate: Number(l.int_rate) || 0,
        handling: Number(l.handling) || 0,
        financial_pct: Number(l.financial_pct) || 0,
        gst_pct: Number(l.gst_pct) || 0,
        manual_fc: Number(l.manual_fc) || 0,
        manual_selling_price: Number(l.manual_selling_price) || 0,
        override_reason: l.override_reason ? String(l.override_reason).slice(0, 255) : '',
      })),
    };
  };

  const handleSave = async (silent = false) => {
    if (!validateForm()) {
      showToast?.('Please resolve validation errors before saving', 'warning');
      return null;
    }

    try {
      setSaving(true);
      const payload = preparePayload();
      let savedQuote;

      if (isEditing) {
        savedQuote = await quotationsAPI.update(quotationId, payload);
        if (!silent) showToast?.(`Quotation ${savedQuote.enquiry_no} saved (v${savedQuote.version})`, 'success');
      } else {
        savedQuote = await quotationsAPI.create(payload);
        if (!silent) showToast?.(`Quotation ${savedQuote.enquiry_no} created successfully`, 'success');
      }

      setVersion(savedQuote.version);
      setStatus(savedQuote.status);
      return savedQuote;
    } catch (err) {
      console.error('Save error:', err);
      const msg = err instanceof APIError ? err.message : 'Failed to save quotation.';
      showToast?.(msg, 'error');
      return null;
    } finally {
      setSaving(false);
    }
  };

  const handleCalculate = async () => {
    const saved = await handleSave(true);
    if (!saved) return;

    try {
      setCalculating(true);
      showToast?.('Performing server-side calculation...', 'info');
      const snapshot = await quotationsAPI.calculate(saved.id);
      showToast?.(`Calculation completed successfully for ${saved.enquiry_no}`, 'success');
      onViewResults(saved.id, snapshot);
    } catch (err) {
      console.error('Calculation error:', err);
      const msg = err instanceof APIError ? err.message : 'Calculation failed.';
      showToast?.(msg, 'error');
    } finally {
      setCalculating(false);
    }
  };

  const handleExport = async () => {
    if (!quotationId) {
      showToast?.('Please save the quotation before exporting', 'warning');
      return;
    }
    try {
      setExporting(true);
      showToast?.('Generating commercial Excel workbook...', 'info');
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

  if (loading) {
    return (
      <div className="editor-loading">
        <div className="spinner" />
        <p>Loading quotation details...</p>
      </div>
    );
  }

  return (
    <div className="quotation-editor-container">
      {/* Top Action Header */}
      <div className="editor-top-bar">
        <div className="editor-title-group">
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={() => onNavigate('dashboard')}
          >
            &larr; Back to Dashboard
          </button>
          <div className="title-row">
            <h1 className="editor-title">
              {isEditing ? `Edit Quotation ${enquiryNo}` : 'New Commercial Quotation'}
            </h1>
            <span className={`status-badge status-${status.toLowerCase()}`}>
              {status}
            </span>
            {isEditing && <span className="version-tag">Version {version}</span>}
          </div>
        </div>

        <div className="editor-actions">
          {isEditing && (
            <button
              type="button"
              className="btn btn-outline"
              onClick={handleExport}
              disabled={exporting || saving || calculating}
              title="Download Excel workbook"
            >
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="icon">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                <polyline points="7 10 12 15 17 10" />
                <line x1="12" y1="15" x2="12" y2="3" />
              </svg>
              {exporting ? 'Exporting...' : 'Export Excel'}
            </button>
          )}

          <button
            type="button"
            className="btn btn-outline"
            onClick={() => handleSave(false)}
            disabled={saving || calculating}
          >
            {saving ? 'Saving...' : 'Save Draft'}
          </button>

          <button
            type="button"
            className="btn btn-primary btn-calculate-primary"
            onClick={handleCalculate}
            disabled={calculating || saving}
          >
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="icon">
              <rect x="4" y="2" width="16" height="20" rx="2" />
              <line x1="8" y1="6" x2="16" y2="6" />
              <line x1="16" y1="14" x2="16" y2="18" />
              <path d="M16 10h.01" />
              <path d="M12 10h.01" />
              <path d="M8 10h.01" />
              <path d="M12 14h.01" />
              <path d="M8 14h.01" />
              <path d="M12 18h.01" />
              <path d="M8 18h.01" />
            </svg>
            {calculating ? 'Calculating...' : 'Calculate Quotation'}
          </button>
        </div>
      </div>

      {/* Validation Errors Notice */}
      {Object.keys(formErrors).length > 0 && (
        <div className="alert alert-error">
          <strong>Please correct the following errors:</strong>
          <ul>
            {Object.entries(formErrors).map(([key, msg]) => (
              <li key={key}>{msg}</li>
            ))}
          </ul>
        </div>
      )}

      {/* Header Fields Section */}
      <section className="form-card">
        <h2 className="section-title">Quotation Header & Logistics</h2>
        <div className="form-grid-4">
          <div className="form-group">
            <label htmlFor="enquiryNo">Enquiry Number *</label>
            <input
              id="enquiryNo"
              type="text"
              className={`input ${formErrors.enquiryNo ? 'input-error' : ''}`}
              value={enquiryNo}
              onChange={(e) => setEnquiryNo(e.target.value)}
              placeholder="e.g. ENQ-26-27-001"
              required
            />
          </div>

          <div className="form-group">
            <label htmlFor="customerName">Customer Name *</label>
            <input
              id="customerName"
              type="text"
              className={`input ${formErrors.customerName ? 'input-error' : ''}`}
              value={customerName}
              onChange={(e) => setCustomerName(e.target.value)}
              placeholder="e.g. Larsen & Toubro Hydrocarbon"
              required
            />
          </div>

          <div className="form-group">
            <label htmlFor="projectName">Project Name</label>
            <input
              id="projectName"
              type="text"
              className="input"
              value={projectName}
              onChange={(e) => setProjectName(e.target.value)}
              placeholder="e.g. Hazira Pipeline Package 2"
            />
          </div>

          <div className="form-group">
            <label htmlFor="deliveryLocation">Delivery Location</label>
            <input
              id="deliveryLocation"
              type="text"
              className="input"
              value={deliveryLocation}
              onChange={(e) => setDeliveryLocation(e.target.value)}
              placeholder="e.g. Hazira, Gujarat"
            />
          </div>

          <div className="form-group">
            <label htmlFor="distanceKm">Distance (km)</label>
            <input
              id="distanceKm"
              type="number"
              min="0"
              step="1"
              className="input"
              value={distanceKm}
              onChange={(e) => setDistanceKm(e.target.value)}
            />
          </div>

          <div className="form-group">
            <label htmlFor="tripRate">Trip Rate (₹ / Trailer)</label>
            <input
              id="tripRate"
              type="number"
              min="0"
              step="500"
              className="input"
              value={tripRate}
              onChange={(e) => setTripRate(e.target.value)}
            />
          </div>

          <div className="form-group">
            <label htmlFor="calculationMode">Calculation Mode</label>
            <select
              id="calculationMode"
              className="input select-input"
              value={calculationMode}
              onChange={(e) => setCalculationMode(e.target.value)}
            >
              <option value="REFERENCE">REFERENCE (Golden Policy — Exact Commercial Parity)</option>
              <option value="NORMAL">NORMAL (Per-Line Ceiling Trailers)</option>
              <option value="LEGACY">LEGACY (Legacy Desktop Emulation)</option>
            </select>
          </div>
        </div>
      </section>

      {/* Quotation Line Items Section */}
      <section className="form-card lines-section">
        <div className="section-header-row">
          <div>
            <h2 className="section-title">Quotation Lines ({lines.length})</h2>
            <p className="section-subtitle">Define pipe specifications, quantities, coatings, and commercial rate adjustments.</p>
          </div>
          <button
            type="button"
            className="btn btn-secondary btn-sm"
            onClick={handleAddLine}
          >
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="icon">
              <line x1="12" y1="5" x2="12" y2="19" />
              <line x1="5" y1="12" x2="19" y2="12" />
            </svg>
            Add Line Item
          </button>
        </div>

        <div className="lines-list">
          {lines.map((line, idx) => {
            const isExpanded = expandedLineIndex === idx;

            return (
              <div key={idx} className="line-item-card">
                {/* Line Summary Row */}
                <div className="line-item-header">
                  <div className="line-badge-title">
                    <span className="line-number-pill">Line #{line.line_number}</span>
                    <span className="line-spec-summary">
                      {line.size_mm} mm ({line.basis}) &times; {line.thickness_mm} mm WT &bull; {formatNumber(line.quantity, 0)} {line.unit}
                    </span>
                  </div>

                  <div className="line-header-actions">
                    <button
                      type="button"
                      className="btn btn-ghost btn-xs"
                      onClick={() => setExpandedLineIndex(isExpanded ? null : idx)}
                    >
                      {isExpanded ? 'Hide Commercial Rates ▲' : 'Commercial Rates & Overrides ▼'}
                    </button>

                    <button
                      type="button"
                      className="btn btn-ghost btn-xs"
                      onClick={() => handleDuplicateLine(idx)}
                      title="Duplicate this line"
                    >
                      Duplicate
                    </button>

                    <button
                      type="button"
                      className="btn btn-ghost btn-xs text-danger"
                      onClick={() => handleRemoveLine(idx)}
                      disabled={lines.length <= 1}
                      title="Remove this line"
                    >
                      Remove
                    </button>
                  </div>
                </div>

                {/* Primary Specifications Grid */}
                <div className="line-grid-primary">
                  <div className="form-group">
                    <label>Nominal Size (mm) *</label>
                    <input
                      type="number"
                      step="0.1"
                      min="1"
                      className="input"
                      value={line.size_mm}
                      onChange={(e) => handleLineChange(idx, 'size_mm', e.target.value)}
                      required
                    />
                  </div>

                  <div className="form-group">
                    <label>Wall Thickness (mm) *</label>
                    <input
                      type="number"
                      step="0.01"
                      min="0.1"
                      className="input"
                      value={line.thickness_mm}
                      onChange={(e) => handleLineChange(idx, 'thickness_mm', e.target.value)}
                      required
                    />
                  </div>

                  <div className="form-group">
                    <label>Sizing Basis</label>
                    <select
                      className="input select-input"
                      value={line.basis}
                      onChange={(e) => handleLineChange(idx, 'basis', e.target.value)}
                    >
                      <option value="ID">Inner Diameter (ID)</option>
                      <option value="OD">Outer Diameter (OD)</option>
                    </select>
                  </div>

                  <div className="form-group">
                    <label>Quantity *</label>
                    <input
                      type="number"
                      step="1"
                      min="1"
                      className="input"
                      value={line.quantity}
                      onChange={(e) => handleLineChange(idx, 'quantity', e.target.value)}
                      required
                    />
                  </div>

                  <div className="form-group">
                    <label>Unit</label>
                    <select
                      className="input select-input"
                      value={line.unit}
                      onChange={(e) => handleLineChange(idx, 'unit', e.target.value)}
                    >
                      <option value="Metres">Metres</option>
                      <option value="MT">MT</option>
                      <option value="Nos">Nos</option>
                    </select>
                  </div>

                  <div className="form-group">
                    <label>Pipe Length (m)</label>
                    <input
                      type="number"
                      step="0.5"
                      min="6"
                      max="13"
                      className="input"
                      value={line.pipe_length_m}
                      onChange={(e) => handleLineChange(idx, 'pipe_length_m', e.target.value)}
                    />
                  </div>

                  <div className="form-group">
                    <label>Pipes / Trailer</label>
                    <input
                      type="number"
                      step="1"
                      min="1"
                      className="input"
                      value={line.pipes_per_trailer}
                      onChange={(e) => handleLineChange(idx, 'pipes_per_trailer', e.target.value)}
                    />
                  </div>

                  <div className="form-group-checkboxes">
                    <label className="checkbox-label">
                      <input
                        type="checkbox"
                        checked={line.has_external}
                        onChange={(e) => handleLineChange(idx, 'has_external', e.target.checked)}
                      />
                      <span>External Coating (3LPE)</span>
                    </label>

                    <label className="checkbox-label">
                      <input
                        type="checkbox"
                        checked={line.has_internal}
                        onChange={(e) => handleLineChange(idx, 'has_internal', e.target.checked)}
                      />
                      <span>Internal Coating (Epoxy)</span>
                    </label>
                  </div>
                </div>

                {/* Collapsible Commercial Rates & Cost Sheet Modifiers */}
                {isExpanded && (
                  <div className="line-commercial-drawer">
                    <h4 className="drawer-title">Commercial Rates & Surcharges (Line #{line.line_number})</h4>
                    <div className="drawer-grid">
                      <div className="form-group">
                        <label>Base Steel Rate (₹/MT)</label>
                        <input
                          type="number"
                          step="100"
                          className="input"
                          value={line.steel_rate}
                          onChange={(e) => handleLineChange(idx, 'steel_rate', e.target.value)}
                        />
                      </div>

                      <div className="form-group">
                        <label>Grade Extra (₹/MT)</label>
                        <input
                          type="number"
                          step="50"
                          className="input"
                          value={line.grade_extra}
                          onChange={(e) => handleLineChange(idx, 'grade_extra', e.target.value)}
                        />
                      </div>

                      <div className="form-group">
                        <label>Yield Burning (%)</label>
                        <input
                          type="number"
                          step="0.1"
                          className="input"
                          value={line.yield_pct}
                          onChange={(e) => handleLineChange(idx, 'yield_pct', e.target.value)}
                        />
                      </div>

                      <div className="form-group">
                        <label>Conversion Cost (₹/MT)</label>
                        <input
                          type="number"
                          step="100"
                          className="input"
                          value={line.conversion}
                          onChange={(e) => handleLineChange(idx, 'conversion', e.target.value)}
                        />
                      </div>

                      <div className="form-group">
                        <label>Testing & NDT (₹/MT)</label>
                        <input
                          type="number"
                          step="50"
                          className="input"
                          value={line.testing}
                          onChange={(e) => handleLineChange(idx, 'testing', e.target.value)}
                        />
                      </div>

                      <div className="form-group">
                        <label>Margin (₹/MT)</label>
                        <input
                          type="number"
                          step="100"
                          className="input"
                          value={line.margin}
                          onChange={(e) => handleLineChange(idx, 'margin', e.target.value)}
                        />
                      </div>

                      <div className="form-group">
                        <label>Ext Coating Rate (₹/m&sup2;)</label>
                        <input
                          type="number"
                          step="10"
                          className="input"
                          value={line.ext_rate}
                          onChange={(e) => handleLineChange(idx, 'ext_rate', e.target.value)}
                        />
                      </div>

                      <div className="form-group">
                        <label>Int Coating Rate (₹/m&sup2;)</label>
                        <input
                          type="number"
                          step="10"
                          className="input"
                          value={line.int_rate}
                          onChange={(e) => handleLineChange(idx, 'int_rate', e.target.value)}
                        />
                      </div>

                      <div className="form-group">
                        <label>Handling & Port (₹/MT)</label>
                        <input
                          type="number"
                          step="10"
                          className="input"
                          value={line.handling}
                          onChange={(e) => handleLineChange(idx, 'handling', e.target.value)}
                        />
                      </div>

                      <div className="form-group">
                        <label>Financial Cost (%)</label>
                        <input
                          type="number"
                          step="0.1"
                          className="input"
                          value={line.financial_pct}
                          onChange={(e) => handleLineChange(idx, 'financial_pct', e.target.value)}
                        />
                      </div>

                      <div className="form-group">
                        <label>GST (%)</label>
                        <input
                          type="number"
                          step="1"
                          className="input"
                          value={line.gst_pct}
                          onChange={(e) => handleLineChange(idx, 'gst_pct', e.target.value)}
                        />
                      </div>

                      <div className="form-group">
                        <label>Manual FC Override (₹)</label>
                        <input
                          type="number"
                          step="100"
                          className="input"
                          value={line.manual_fc}
                          onChange={(e) => handleLineChange(idx, 'manual_fc', e.target.value)}
                        />
                      </div>

                      <div className="form-group">
                        <label>Manual Price Override (₹/unit)</label>
                        <input
                          type="number"
                          step="100"
                          className="input"
                          value={line.manual_selling_price}
                          onChange={(e) => handleLineChange(idx, 'manual_selling_price', e.target.value)}
                        />
                      </div>

                      <div className="form-group form-group-span-2">
                        <label>Override Reason</label>
                        <input
                          type="text"
                          className="input"
                          value={line.override_reason}
                          onChange={(e) => handleLineChange(idx, 'override_reason', e.target.value)}
                          placeholder="e.g. Competitive benchmark against PSL / Welspun"
                        />
                      </div>
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </section>
    </div>
  );
}
