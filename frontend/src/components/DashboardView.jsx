import React, { useState, useEffect, useMemo } from 'react';
import { quotationsAPI, APIError } from '../api/client';
import { formatNumber } from '../utils/format';

export function DashboardView({ onNavigate, onSelectQuotation, onViewResults, showToast }) {
  const [quotations, setQuotations] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [actionLoadingId, setActionLoadingId] = useState(null);

  const fetchQuotations = async () => {
    try {
      setLoading(true);
      setError(null);
      const data = await quotationsAPI.list();
      setQuotations(data || []);
    } catch (err) {
      console.error('Failed to load quotations:', err);
      const msg = err instanceof APIError ? err.message : 'Could not fetch quotations from server.';
      setError(msg);
      showToast?.(msg, 'error');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchQuotations();
  }, []);

  const handleExport = async (e, id, enquiryNo) => {
    e.stopPropagation();
    try {
      setActionLoadingId(id);
      showToast?.(`Generating Excel export for ${enquiryNo}...`, 'info');
      const filename = await quotationsAPI.exportExcel(id);
      showToast?.(`Downloaded ${filename} successfully`, 'success');
    } catch (err) {
      console.error('Export error:', err);
      const msg = err instanceof APIError ? err.message : 'Excel export failed.';
      showToast?.(msg, 'error');
    } finally {
      setActionLoadingId(null);
    }
  };

  const handleCalculateDirect = async (e, id, enquiryNo) => {
    e.stopPropagation();
    try {
      setActionLoadingId(id);
      showToast?.(`Calculating quotation ${enquiryNo}...`, 'info');
      const snapshot = await quotationsAPI.calculate(id);
      showToast?.(`Quotation ${enquiryNo} calculated successfully!`, 'success');
      // Refresh quote list and open results
      await fetchQuotations();
      onViewResults(id, snapshot);
    } catch (err) {
      console.error('Calculation error:', err);
      const msg = err instanceof APIError ? err.message : 'Calculation failed.';
      showToast?.(msg, 'error');
    } finally {
      setActionLoadingId(null);
    }
  };

  const handleViewResults = async (e, id) => {
    e.stopPropagation();
    try {
      setActionLoadingId(id);
      const snapshot = await quotationsAPI.getSnapshot(id);
      onViewResults(id, snapshot);
    } catch (err) {
      console.error('Snapshot retrieval error:', err);
      const msg = err instanceof APIError ? err.message : 'Could not fetch calculation snapshot.';
      showToast?.(msg, 'error');
    } finally {
      setActionLoadingId(null);
    }
  };

  const filteredQuotes = useMemo(() => {
    if (!searchQuery.trim()) return quotations;
    const q = searchQuery.toLowerCase();
    return quotations.filter(
      (item) =>
        (item.enquiry_no && item.enquiry_no.toLowerCase().includes(q)) ||
        (item.customer_name && item.customer_name.toLowerCase().includes(q)) ||
        (item.project_name && item.project_name.toLowerCase().includes(q)) ||
        (item.delivery_location && item.delivery_location.toLowerCase().includes(q))
    );
  }, [quotations, searchQuery]);

  const stats = useMemo(() => {
    const total = quotations.length;
    const calculated = quotations.filter((q) => q.status === 'CALCULATED').length;
    const drafts = quotations.filter((q) => q.status === 'DRAFT').length;
    return { total, calculated, drafts };
  }, [quotations]);

  return (
    <div className="dashboard-container">
      {/* Top Header & Stats */}
      <div className="dashboard-header">
        <div>
          <h1 className="dashboard-title">Commercial Quotations</h1>
          <p className="dashboard-subtitle">Manage, calculate, and export industrial pipe quotations</p>
        </div>
        <div className="header-actions">
          <button
            type="button"
            className="btn btn-outline btn-refresh"
            onClick={fetchQuotations}
            disabled={loading}
            title="Refresh quotation list"
          >
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className={`icon ${loading ? 'spin' : ''}`}>
              <polyline points="23 4 23 10 17 10" />
              <polyline points="1 20 1 14 7 14" />
              <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15" />
            </svg>
            Refresh
          </button>
          <button
            type="button"
            className="btn btn-primary"
            onClick={() => onNavigate('new-quote')}
          >
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="icon">
              <line x1="12" y1="5" x2="12" y2="19" />
              <line x1="5" y1="12" x2="19" y2="12" />
            </svg>
            New Quotation
          </button>
        </div>
      </div>

      {/* KPI Stats Grid */}
      <div className="stats-grid">
        <div className="stat-card">
          <div className="stat-icon stat-icon-blue">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
              <polyline points="14 2 14 8 20 8" />
              <line x1="16" y1="13" x2="8" y2="13" />
              <line x1="16" y1="17" x2="8" y2="17" />
              <polyline points="10 9 9 9 8 9" />
            </svg>
          </div>
          <div className="stat-info">
            <span className="stat-label">Total Quotations</span>
            <span className="stat-value">{stats.total}</span>
          </div>
        </div>

        <div className="stat-card">
          <div className="stat-icon stat-icon-emerald">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <polyline points="20 6 9 17 4 12" />
            </svg>
          </div>
          <div className="stat-info">
            <span className="stat-label">Calculated & Verified</span>
            <span className="stat-value">{stats.calculated}</span>
          </div>
        </div>

        <div className="stat-card">
          <div className="stat-icon stat-icon-amber">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="12" cy="12" r="10" />
              <line x1="12" y1="8" x2="12" y2="12" />
              <line x1="12" y1="16" x2="12.01" y2="16" />
            </svg>
          </div>
          <div className="stat-info">
            <span className="stat-label">Draft Status</span>
            <span className="stat-value">{stats.drafts}</span>
          </div>
        </div>
      </div>

      {/* Filter and Search Bar */}
      <div className="table-controls">
        <div className="search-box">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="search-icon">
            <circle cx="11" cy="11" r="8" />
            <line x1="21" y1="21" x2="16.65" y2="16.65" />
          </svg>
          <input
            type="text"
            className="input search-input"
            placeholder="Search by Enquiry #, Customer, Project or Location..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
          />
          {searchQuery && (
            <button type="button" className="clear-search-btn" onClick={() => setSearchQuery('')}>
              &times;
            </button>
          )}
        </div>
      </div>

      {/* Error state */}
      {error && (
        <div className="alert alert-error">
          <div className="alert-content">
            <strong>Unable to load quotations:</strong> {error}
          </div>
          <button type="button" className="btn btn-sm btn-outline" onClick={fetchQuotations}>
            Try Again
          </button>
        </div>
      )}

      {/* Quotations Data Table */}
      <div className="table-wrapper">
        {loading ? (
          <div className="loading-state">
            <div className="spinner" />
            <p>Loading quotations from server...</p>
          </div>
        ) : filteredQuotes.length === 0 ? (
          <div className="empty-state">
            <div className="empty-icon">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                <rect x="2" y="3" width="20" height="14" rx="2" />
                <line x1="8" y1="21" x2="16" y2="21" />
                <line x1="12" y1="17" x2="12" y2="21" />
              </svg>
            </div>
            <h3>{searchQuery ? 'No quotations match your search' : 'No quotations found'}</h3>
            <p>
              {searchQuery
                ? 'Try adjusting your search criteria or clearing filters.'
                : 'Create your first commercial quotation to run cost calculations and generate Excel exports.'}
            </p>
            {searchQuery ? (
              <button type="button" className="btn btn-outline" onClick={() => setSearchQuery('')}>
                Clear Filter
              </button>
            ) : (
              <button type="button" className="btn btn-primary" onClick={() => onNavigate('new-quote')}>
                + Create Quotation
              </button>
            )}
          </div>
        ) : (
          <table className="data-table">
            <thead>
              <tr>
                <th>Enquiry #</th>
                <th>Customer / Project</th>
                <th>Delivery Location</th>
                <th>Distance / Trip</th>
                <th>Mode</th>
                <th>Lines</th>
                <th>Ver</th>
                <th>Status</th>
                <th className="text-right">Actions</th>
              </tr>
            </thead>
            <tbody>
              {filteredQuotes.map((quote) => {
                const isCalculated = quote.status === 'CALCULATED';
                const isActing = actionLoadingId === quote.id;

                return (
                  <tr
                    key={quote.id}
                    className="clickable-row"
                    onClick={() => onSelectQuotation(quote.id)}
                    title="Click to view and edit quotation"
                  >
                    <td>
                      <span className="badge-enquiry">{quote.enquiry_no}</span>
                    </td>
                    <td>
                      <div className="customer-cell">
                        <span className="customer-name">{quote.customer_name}</span>
                        {quote.project_name && (
                          <span className="project-sub">{quote.project_name}</span>
                        )}
                      </div>
                    </td>
                    <td>
                      <span className="location-text">{quote.delivery_location || '—'}</span>
                    </td>
                    <td>
                      <div className="logistics-cell">
                        <span>{formatNumber(quote.distance_km, 0)} km</span>
                        <span className="trip-sub">₹{formatNumber(quote.trip_rate, 0)}/trip</span>
                      </div>
                    </td>
                    <td>
                      <span className={`badge-mode mode-${quote.calculation_mode.toLowerCase()}`}>
                        {quote.calculation_mode}
                      </span>
                    </td>
                    <td>
                      <span className="line-count-badge">
                        {quote.lines ? quote.lines.length : 0} {quote.lines && quote.lines.length === 1 ? 'line' : 'lines'}
                      </span>
                    </td>
                    <td>
                      <span className="version-tag">v{quote.version}</span>
                    </td>
                    <td>
                      <span className={`status-badge status-${quote.status.toLowerCase()}`}>
                        {quote.status}
                      </span>
                    </td>
                    <td className="text-right actions-cell" onClick={(e) => e.stopPropagation()}>
                      <div className="btn-group">
                        {isCalculated ? (
                          <button
                            type="button"
                            className="btn btn-outline btn-xs btn-view-results"
                            onClick={(e) => handleViewResults(e, quote.id)}
                            disabled={isActing}
                            title="View authoritative calculation snapshot"
                          >
                            Results
                          </button>
                        ) : (
                          <button
                            type="button"
                            className="btn btn-secondary btn-xs btn-calculate"
                            onClick={(e) => handleCalculateDirect(e, quote.id, quote.enquiry_no)}
                            disabled={isActing}
                            title="Calculate quotation commercial totals"
                          >
                            {isActing ? '...' : 'Calculate'}
                          </button>
                        )}

                        <button
                          type="button"
                          className="btn btn-outline btn-xs"
                          onClick={() => onSelectQuotation(quote.id)}
                          disabled={isActing}
                          title="Edit quotation fields and lines"
                        >
                          Edit
                        </button>

                        <button
                          type="button"
                          className="btn btn-outline btn-xs btn-export"
                          onClick={(e) => handleExport(e, quote.id, quote.enquiry_no)}
                          disabled={isActing}
                          title="Export commercial quotation to Excel (.xlsx)"
                        >
                          {isActing ? '...' : 'Excel'}
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}
