import React, { useState, useEffect } from 'react';
import { quotationsAPI, logisticsAPI, APIError } from '../api/client';
import { formatNumber } from '../utils/format';

const PRODUCT_STANDARDS = ['IS-3589-2001', 'API 5L', 'ASTM A139', 'IS 1239 Part-1', 'IS 1161'];
const PIPE_FAMILIES = ['HSAW', 'HFW'];
const DIAMETER_BASES = ['ID', 'OD'];
const CUSTOM_PIPE_SIZE_OPTIONS = ['NO', 'YES'];

const STANDARD_SIZES = {
  'HSAW-ID': [1226, 1300, 1350, 1400],
  'HSAW-OD': [457, 508, 610, 762, 914, 1016, 1219, 1422, 1524, 1829, 2134, 2540],
  'HFW-OD': [168.3, 219.1, 273.1, 323.9, 355.6, 406.4, 457.2, 508],
  'HFW-ID': [168.3, 219.1, 273.1, 323.9, 355.6, 406.4, 457.2, 508],
};

const STANDARD_THICKNESSES = {
  HSAW: [6.4, 7.1, 7.9, 8.4, 9.5, 11.1, 12.7, 15.9, 19.1, 25.4],
  HFW: [3.5, 4.0, 4.5, 5.0, 5.6, 6.4, 7.1, 8.0, 8.8, 10.0, 12.5, 14.2, 16.0],
};

const STEEL_GRADES = {
  HSAW: [
    'Fe 410',
    'IS 3589 Fe 330',
    'IS 3589 Fe 450',
    'API 5L B',
    'X42 PSL1',
    'X52 PSL1',
    'X60 PSL1',
    'X65 PSL1',
    'X70 PSL1',
    'X52 PSL2',
    'X65 PSL2',
    'X70 PSL2',
    'X80M PSL2',
  ],
  HFW: [
    'Fe 410',
    'IS 3589 Fe 330',
    'IS 3589 Fe 450',
    'IS 1239 Part-1',
    'IS 1161 YST210',
    'IS 1161 YST240',
    'IS 1161 YST310',
    'API 5L A25',
    'API 5L B',
    'API 5L X42',
    'API 5L X52',
    'API 5L X60',
    'API 5L X65',
    'API 5L X70',
    'API 5L X80M PSL2',
    'API 5CT H40',
    'API 5CT J55',
    'API 5CT K55',
  ],
};

const EXTERNAL_COATINGS = ['Bare', '3LPE', '3LPP', 'FBE', 'DFBE', 'PU', 'Zinc Rich Epoxy'];
const INTERNAL_COATINGS = ['None', 'Liquid Epoxy', 'Food Grade Epoxy', 'Solvent-Free Epoxy', 'PU', 'Zinc Rich Epoxy', 'Cement Lining'];
const QUANTITY_UNITS = ['Metres', 'MT', 'Nos'];
const PRICE_UNITS = ['₹/m', '₹/MT'];

const DEFAULT_LINE = {
  line_number: 1,
  product_standard: 'IS-3589-2001',
  pipe_family: 'HSAW',
  basis: 'ID',
  custom_pipe_size: 'NO',
  size_mm: 1226.0,
  thickness_mm: 8.4,
  steel_grade: 'Fe 410',
  external_coating: 'Bare',
  internal_coating: 'None',
  quantity: 2000,
  unit: 'Metres',
  pipe_length_m: 12.0,
  pipes_per_trailer: 4.0,
  has_external: false,
  has_internal: false,
  steel_rate: 61000.0,
  grade_extra: 0.0,
  yield_pct: 2.0,
  conversion: 4000.0,
  testing: 0.0,
  margin: 0.0,
  ext_rate: 500.0,
  int_rate: 250.0,
  handling: 100.0,
  financial_pct: 1.0,
  gst_pct: 18.0,
  manual_fc: 0.0,
  manual_selling_price: 0.0,
  price_unit: '₹/m',
  override_reason: '',
};

export function QuotationEditorView({ quotationId, onNavigate, onViewResults, showToast }) {
  const isEditing = Boolean(quotationId);

  // Form State
  const [enquiryNo, setEnquiryNo] = useState('');
  const [customerName, setCustomerName] = useState('');
  const [projectName, setProjectName] = useState('');
  const [deliveryLocation, setDeliveryLocation] = useState('');
  const [deliveryPin, setDeliveryPin] = useState('825322');
  const [distanceKm, setDistanceKm] = useState(782);
  const [tripRate, setTripRate] = useState(98000);
  const [logisticsInfo, setLogisticsInfo] = useState(null);
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

  // Resolve logistics from pincode automatically
  const resolveLogistics = async (pin, loc, mode) => {
    try {
      const res = await logisticsAPI.resolve(pin, loc, mode);
      if (res && res.valid) {
        setLogisticsInfo(res);
        setDistanceKm(res.distance_km);
        setTripRate(res.trip_rate);
        if (!deliveryLocation && res.formatted_address) {
          setDeliveryLocation(res.formatted_address);
        }
      }
    } catch (e) {
      console.warn('Logistics resolution failed, using defaults:', e);
    }
  };

  // Fetch existing quotation if editing, or initialize new
  useEffect(() => {
    if (!quotationId) {
      // Pre-fill enquiry number for new quote with timestamp suffix
      const now = new Date();
      const code = `MKTGENQ26-${Math.floor(10000 + Math.random() * 90000)}`;
      setEnquiryNo(code);
      setDeliveryPin('825322');
      setDeliveryLocation('Bagodar, Jharkhand');
      resolveLogistics('825322', 'Bagodar, Jharkhand', 'REFERENCE');
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
        setDistanceKm(q.distance_km || 782);
        setTripRate(q.trip_rate || 98000);
        setCalculationMode(q.calculation_mode || 'REFERENCE');
        setStatus(q.status || 'DRAFT');
        setVersion(q.version || 1);

        // Extract PIN from delivery location if available
        const pinMatch = (q.delivery_location || '').match(/\b([1-9][0-9]{5})\b/);
        const pin = pinMatch ? pinMatch[1] : '825322';
        setDeliveryPin(pin);
        resolveLogistics(pin, q.delivery_location, q.calculation_mode || 'REFERENCE');

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

  const handleLineMultipleChange = (index, updates) => {
    setLines((prev) => {
      const copy = [...prev];
      copy[index] = { ...copy[index], ...updates };
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
      lines: lines.map((l, idx) => {
        const ext = l.external_coating || 'Bare';
        const intCoating = l.internal_coating || 'None';
        const hasExternal = !['BARE', 'NONE', ''].includes(String(ext).toUpperCase().trim());
        const hasInternal = !['NONE', 'BARE', ''].includes(String(intCoating).toUpperCase().trim());
        return {
          line_number: idx + 1,
          size_mm: Number(l.size_mm),
          thickness_mm: Number(l.thickness_mm),
          basis: (l.basis || 'ID').toUpperCase(),
          quantity: Number(l.quantity),
          unit: l.unit === 'MT' ? 'MT' : l.unit === 'Nos' ? 'Nos' : 'Metres',
          pipe_length_m: Number(l.pipe_length_m) || 12.0,
          pipes_per_trailer: Number(l.pipes_per_trailer) || 4.0,
          has_external: hasExternal,
          has_internal: hasInternal,
          steel_rate: Number(l.steel_rate) || 61000.0,
          grade_extra: Number(l.grade_extra) || 0,
          yield_pct: Number(l.yield_pct) || 2.0,
          conversion: Number(l.conversion) || 4000.0,
          testing: Number(l.testing) || 0,
          margin: Number(l.margin) || 0,
          ext_rate: Number(l.ext_rate) || 500.0,
          int_rate: Number(l.int_rate) || 250.0,
          handling: Number(l.handling) || 100.0,
          financial_pct: Number(l.financial_pct) || 1.0,
          gst_pct: Number(l.gst_pct) || 18.0,
          manual_fc: Number(l.manual_fc) || 0.0,
          manual_selling_price: Number(l.manual_selling_price) || 0.0,
          price_unit: l.price_unit || '₹/m',
          override_reason: l.override_reason ? String(l.override_reason).slice(0, 255) : '',
          product_standard: l.product_standard || 'IS-3589-2001',
          pipe_family: l.pipe_family || 'HSAW',
          steel_grade: l.steel_grade || 'Fe 410',
          external_coating: ext,
          internal_coating: intCoating,
          custom_pipe_size: l.custom_pipe_size || 'NO',
        };
      }),
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
            <label htmlFor="deliveryPin">Delivery Pincode (6 digits) *</label>
            <input
              id="deliveryPin"
              type="text"
              maxLength="6"
              className="input"
              value={deliveryPin}
              onChange={(e) => {
                const val = e.target.value.replace(/\D/g, '');
                setDeliveryPin(val);
                if (val.length === 6 && !val.startsWith('0')) {
                  resolveLogistics(val, deliveryLocation, calculationMode);
                }
              }}
              placeholder="e.g. 825322"
              required
            />
          </div>

          <div className="form-group">
            <label htmlFor="calculationMode">Calculation Mode</label>
            <select
              id="calculationMode"
              className="input select-input"
              value={calculationMode}
              onChange={(e) => {
                setCalculationMode(e.target.value);
                resolveLogistics(deliveryPin, deliveryLocation, e.target.value);
              }}
            >
              <option value="REFERENCE">REFERENCE (Golden Policy — Exact Commercial Parity)</option>
              <option value="NORMAL">NORMAL (Dynamic Route Freight Tariff)</option>
              <option value="LEGACY">LEGACY (Legacy Desktop Emulation)</option>
            </select>
          </div>
        </div>

        {logisticsInfo && (
          <div className="logistics-badge-banner" style={{ marginTop: '1rem', padding: '0.75rem 1rem', background: 'rgba(59, 130, 246, 0.1)', border: '1px solid rgba(59, 130, 246, 0.3)', borderRadius: '6px', fontSize: '0.875rem' }}>
            📍 <strong>Logistics (Auto-Calculated from PIN {deliveryPin})</strong>: {logisticsInfo.formatted_address || 'India'} • <strong>Road Distance</strong>: {distanceKm} km • <strong>Freight Rate</strong>: ₹{Number(tripRate).toLocaleString()} / trailer
          </div>
        )}

        <details className="advanced-logistics-details" style={{ marginTop: '0.75rem', fontSize: '0.8rem', color: 'var(--text-muted)' }}>
          <summary style={{ cursor: 'pointer', padding: '0.25rem 0' }}>⚙️ Advanced Logistics Overrides (Optional)</summary>
          <div className="form-grid-2" style={{ marginTop: '0.5rem' }}>
            <div className="form-group">
              <label htmlFor="distanceKm">Distance (km)</label>
              <input id="distanceKm" type="number" min="1" step="10" className="input" value={distanceKm} onChange={(e) => setDistanceKm(Number(e.target.value))} />
            </div>
            <div className="form-group">
              <label htmlFor="tripRate">Trip Rate (₹ / Trailer)</label>
              <input id="tripRate" type="number" min="1000" step="1000" className="input" value={tripRate} onChange={(e) => setTripRate(Number(e.target.value))} />
            </div>
          </div>
        </details>
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
                      {Number(line.manual_selling_price) > 0 ? (
                        <span style={{ marginLeft: '0.5rem', color: '#10b981', fontWeight: 600 }}>
                          • 🏷️ ₹{formatNumber(line.manual_selling_price, 2)} (Price Override)
                        </span>
                      ) : (
                        <span style={{ marginLeft: '0.5rem', color: '#60a5fa' }}>
                          • ⚙️ Auto-Formula
                        </span>
                      )}
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

                {/* Primary Specifications Grid - All 13 Excel Options */}
                <div className="line-grid-primary">
                  {/* Row 1: Standards, Family, Sizing Basis & Custom Sizing */}
                  <div className="form-group">
                    <label>Product Standard</label>
                    <select
                      className="input select-input"
                      value={line.product_standard || 'IS-3589-2001'}
                      onChange={(e) => handleLineChange(idx, 'product_standard', e.target.value)}
                    >
                      {PRODUCT_STANDARDS.map((std) => (
                        <option key={std} value={std}>{std}</option>
                      ))}
                    </select>
                  </div>

                  <div className="form-group">
                    <label>Pipe Family</label>
                    <select
                      className="input select-input"
                      value={line.pipe_family || 'HSAW'}
                      onChange={(e) => {
                        const fam = e.target.value;
                        const b = line.basis || 'ID';
                        const key = `${fam}-${b}`;
                        const sizes = STANDARD_SIZES[key] || [1226];
                        const thks = STANDARD_THICKNESSES[fam] || [8.4];
                        const grades = STEEL_GRADES[fam] || ['Fe 410'];
                        handleLineMultipleChange(idx, {
                          pipe_family: fam,
                          size_mm: sizes[0],
                          thickness_mm: thks[0],
                          steel_grade: grades[0],
                        });
                      }}
                    >
                      {PIPE_FAMILIES.map((fam) => (
                        <option key={fam} value={fam}>{fam}</option>
                      ))}
                    </select>
                  </div>

                  <div className="form-group">
                    <label>Diameter Basis</label>
                    <select
                      className="input select-input"
                      value={line.basis || 'ID'}
                      onChange={(e) => {
                        const b = e.target.value;
                        const fam = line.pipe_family || 'HSAW';
                        const key = `${fam}-${b}`;
                        const sizes = STANDARD_SIZES[key] || [1226];
                        handleLineMultipleChange(idx, {
                          basis: b,
                          size_mm: sizes[0],
                        });
                      }}
                    >
                      <option value="ID">ID (Inner Diameter)</option>
                      <option value="OD">OD (Outer Diameter)</option>
                    </select>
                  </div>

                  <div className="form-group">
                    <label>Custom Pipe Size</label>
                    <select
                      className="input select-input"
                      value={line.custom_pipe_size || 'NO'}
                      onChange={(e) => handleLineChange(idx, 'custom_pipe_size', e.target.value)}
                    >
                      <option value="NO">NO (Standard Master Sizes)</option>
                      <option value="YES">YES (Custom Dimensions)</option>
                    </select>
                  </div>

                  {/* Row 2: Dimensions & Material Grade */}
                  <div className="form-group">
                    <label>Size (mm) *</label>
                    {line.custom_pipe_size === 'YES' ? (
                      <input
                        type="number"
                        step="1"
                        min="1"
                        max="5000"
                        className="input"
                        value={line.size_mm}
                        onChange={(e) => handleLineChange(idx, 'size_mm', Number(e.target.value))}
                        required
                      />
                    ) : (
                      <select
                        className="input select-input"
                        value={line.size_mm}
                        onChange={(e) => handleLineChange(idx, 'size_mm', Number(e.target.value))}
                      >
                        {(() => {
                          const key = `${line.pipe_family || 'HSAW'}-${line.basis || 'ID'}`;
                          const sizes = STANDARD_SIZES[key] || [1226];
                          const curr = Number(line.size_mm);
                          const allSizes = sizes.includes(curr) ? sizes : [curr, ...sizes];
                          return allSizes.map((s) => (
                            <option key={s} value={s}>{s} mm</option>
                          ));
                        })()}
                      </select>
                    )}
                  </div>

                  <div className="form-group">
                    <label>Wall Thickness (mm) *</label>
                    {line.custom_pipe_size === 'YES' ? (
                      <input
                        type="number"
                        step="0.1"
                        min="0.1"
                        max="100"
                        className="input"
                        value={line.thickness_mm}
                        onChange={(e) => handleLineChange(idx, 'thickness_mm', Number(e.target.value))}
                        required
                      />
                    ) : (
                      <select
                        className="input select-input"
                        value={line.thickness_mm}
                        onChange={(e) => handleLineChange(idx, 'thickness_mm', Number(e.target.value))}
                      >
                        {(() => {
                          const thks = STANDARD_THICKNESSES[line.pipe_family || 'HSAW'] || [8.4];
                          const curr = Number(line.thickness_mm);
                          const allThks = thks.includes(curr) ? thks : [curr, ...thks];
                          return allThks.map((t) => (
                            <option key={t} value={t}>{t} mm</option>
                          ));
                        })()}
                      </select>
                    )}
                  </div>

                  <div className="form-group">
                    <label>Steel Grade</label>
                    <select
                      className="input select-input"
                      value={line.steel_grade || 'Fe 410'}
                      onChange={(e) => handleLineChange(idx, 'steel_grade', e.target.value)}
                    >
                      {(STEEL_GRADES[line.pipe_family || 'HSAW'] || STEEL_GRADES.HSAW).map((g) => (
                        <option key={g} value={g}>{g}</option>
                      ))}
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
                      value={line.pipe_length_m || 12.0}
                      onChange={(e) => handleLineChange(idx, 'pipe_length_m', Number(e.target.value))}
                    />
                  </div>

                  {/* Row 3: Coatings & Quantities */}
                  <div className="form-group">
                    <label>External Coating</label>
                    <select
                      className="input select-input"
                      value={line.external_coating || 'Bare'}
                      onChange={(e) => {
                        const val = e.target.value;
                        const hasExt = !['BARE', 'NONE', ''].includes(val.toUpperCase().trim());
                        handleLineMultipleChange(idx, {
                          external_coating: val,
                          has_external: hasExt,
                        });
                      }}
                    >
                      {EXTERNAL_COATINGS.map((c) => (
                        <option key={c} value={c}>{c}</option>
                      ))}
                    </select>
                  </div>

                  <div className="form-group">
                    <label>Internal Coating</label>
                    <select
                      className="input select-input"
                      value={line.internal_coating || 'None'}
                      onChange={(e) => {
                        const val = e.target.value;
                        const hasInt = !['NONE', 'BARE', ''].includes(val.toUpperCase().trim());
                        handleLineMultipleChange(idx, {
                          internal_coating: val,
                          has_internal: hasInt,
                        });
                      }}
                    >
                      {INTERNAL_COATINGS.map((c) => (
                        <option key={c} value={c}>{c}</option>
                      ))}
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
                      onChange={(e) => handleLineChange(idx, 'quantity', Number(e.target.value))}
                      required
                    />
                  </div>

                  <div className="form-group">
                    <label>Quantity Unit</label>
                    <select
                      className="input select-input"
                      value={line.unit || 'Metres'}
                      onChange={(e) => handleLineChange(idx, 'unit', e.target.value)}
                    >
                      {QUANTITY_UNITS.map((u) => (
                        <option key={u} value={u}>{u}</option>
                      ))}
                    </select>
                  </div>
                </div>

                {/* Dedicated Primary Price & Commercial Section (Matching Excel QUOTE sheet rows 31-33) */}
                <div className="line-price-entry-card" style={{ marginTop: '1rem', padding: '1rem 1.25rem', background: 'rgba(255, 255, 255, 0.02)', border: '1px solid var(--border)', borderRadius: '8px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.75rem' }}>
                    <h4 style={{ fontSize: '0.9rem', fontWeight: 600, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                      <span>💰 Commercial Price Entry & Freight Logistics</span>
                      <span style={{ fontSize: '0.75rem', fontWeight: 400, color: 'var(--text-muted)' }}>(Line #{line.line_number})</span>
                    </h4>
                  </div>

                  <div className="line-grid-primary">
                    <div className="form-group">
                      <label>Manual Selling Price (Optional)</label>
                      <input
                        type="number"
                        step="100"
                        min="0"
                        className="input"
                        value={line.manual_selling_price || 0}
                        onChange={(e) => handleLineChange(idx, 'manual_selling_price', Number(e.target.value))}
                        placeholder="0.00 (Leave 0 for Auto-Formula)"
                      />
                      <span style={{ fontSize: '0.75rem', color: Number(line.manual_selling_price) > 0 ? '#10b981' : 'var(--text-muted)' }}>
                        {Number(line.manual_selling_price) > 0
                          ? `✓ Unit price locked at ₹${Number(line.manual_selling_price).toLocaleString()} (${line.price_unit || '₹/m'})`
                          : 'Leave 0 to calculate unit price from cost formula'}
                      </span>
                    </div>

                    <div className="form-group">
                      <label>Price Unit</label>
                      <select
                        className="input select-input"
                        value={line.price_unit || '₹/m'}
                        onChange={(e) => handleLineChange(idx, 'price_unit', e.target.value)}
                      >
                        {PRICE_UNITS.map((pu) => (
                          <option key={pu} value={pu}>{pu}</option>
                        ))}
                      </select>
                    </div>

                    <div className="form-group">
                      <label>Steel Base Rate (₹/MT)</label>
                      <input
                        type="number"
                        step="500"
                        min="0"
                        className="input"
                        value={line.steel_rate}
                        onChange={(e) => handleLineChange(idx, 'steel_rate', Number(e.target.value))}
                      />
                    </div>

                    <div className="form-group">
                      <label>Pipes / Trailer</label>
                      <input
                        type="number"
                        step="1"
                        min="1"
                        max="20"
                        className="input"
                        value={line.pipes_per_trailer || 4}
                        onChange={(e) => handleLineChange(idx, 'pipes_per_trailer', Number(e.target.value))}
                      />
                    </div>
                  </div>

                  {Number(line.manual_selling_price) > 0 && (
                    <div className="form-group" style={{ marginTop: '0.75rem' }}>
                      <label>Price Override Reason *</label>
                      <input
                        type="text"
                        className="input"
                        value={line.override_reason || ''}
                        onChange={(e) => handleLineChange(idx, 'override_reason', e.target.value)}
                        placeholder="e.g. Approved client tender rate / Special project pricing"
                        required
                      />
                    </div>
                  )}
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
