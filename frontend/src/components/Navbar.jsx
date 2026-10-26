import React from 'react';
import { useAuth } from '../context/AuthContext';

export function Navbar({ currentView, onNavigate }) {
  const { user, logout } = useAuth();

  return (
    <header className="navbar">
      <div className="navbar-container">
        <div className="navbar-brand" onClick={() => onNavigate('dashboard')} role="button" tabIndex={0}>
          <div className="brand-logo-badge">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="brand-icon">
              <path d="M4 14.899A7 7 0 1 1 15.71 8h1.79a4.5 4.5 0 0 1 2.5 8.242" />
              <path d="M12 12v9" />
              <path d="m8 17 4 4 4-4" />
            </svg>
          </div>
          <div className="brand-titles">
            <span className="brand-name">MEGA PIPES</span>
            <span className="brand-sub">Quotation System</span>
          </div>
        </div>

        {user && (
          <nav className="navbar-nav">
            <button
              type="button"
              className={`nav-link ${currentView === 'dashboard' ? 'active' : ''}`}
              onClick={() => onNavigate('dashboard')}
            >
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="nav-icon">
                <rect x="3" y="3" width="7" height="7" />
                <rect x="14" y="3" width="7" height="7" />
                <rect x="14" y="14" width="7" height="7" />
                <rect x="3" y="14" width="7" height="7" />
              </svg>
              Dashboard
            </button>

            <button
              type="button"
              className="btn btn-primary btn-sm btn-new-quote"
              onClick={() => onNavigate('new-quote')}
            >
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="nav-icon">
                <line x1="12" y1="5" x2="12" y2="19" />
                <line x1="5" y1="12" x2="19" y2="12" />
              </svg>
              New Quotation
            </button>

            <div className="user-profile-menu">
              <div className="user-info">
                <span className="user-name">{user.full_name || user.email.split('@')[0]}</span>
                <span className="user-role-badge">Pilot</span>
              </div>
              <button
                type="button"
                className="btn btn-outline btn-sm btn-logout"
                onClick={logout}
                title="Sign out of Mega Pipes"
              >
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="nav-icon">
                  <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" />
                  <polyline points="16 17 21 12 16 7" />
                  <line x1="21" y1="12" x2="9" y2="12" />
                </svg>
                Logout
              </button>
            </div>
          </nav>
        )}
      </div>
    </header>
  );
}
