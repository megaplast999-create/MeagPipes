import React, { useState, useCallback } from 'react';
import { AuthProvider, useAuth } from './context/AuthContext';
import { Navbar } from './components/Navbar';
import { LoginView } from './components/LoginView';
import { DashboardView } from './components/DashboardView';
import { QuotationEditorView } from './components/QuotationEditorView';
import { CalculationResultsView } from './components/CalculationResultsView';
import { NotificationToast } from './components/NotificationToast';
import './App.css';

function MainContent() {
  const { user, loading } = useAuth();
  const [currentView, setCurrentView] = useState('dashboard');
  const [activeQuotationId, setActiveQuotationId] = useState(null);
  const [activeSnapshot, setActiveSnapshot] = useState(null);
  const [toast, setToast] = useState(null);

  const showToast = useCallback((message, type = 'info') => {
    setToast({ message, type });
    setTimeout(() => {
      setToast((current) => (current && current.message === message ? null : current));
    }, 5000);
  }, []);

  const dismissToast = useCallback(() => {
    setToast(null);
  }, []);

  const handleNavigate = (view) => {
    if (view === 'new-quote') {
      setActiveQuotationId(null);
      setActiveSnapshot(null);
      setCurrentView('new-quote');
    } else {
      setCurrentView(view);
    }
  };

  const handleSelectQuotation = (id) => {
    setActiveQuotationId(id);
    setActiveSnapshot(null);
    setCurrentView('edit-quote');
  };

  const handleViewResults = (id, snapshot) => {
    setActiveQuotationId(id);
    setActiveSnapshot(snapshot);
    setCurrentView('results');
  };

  if (loading) {
    return (
      <div className="app-loading-screen">
        <div className="spinner" />
        <p>Connecting to Mega Pipes quotation engine...</p>
      </div>
    );
  }

  if (!user) {
    return (
      <>
        <LoginView />
        <NotificationToast toast={toast} onDismiss={dismissToast} />
      </>
    );
  }

  return (
    <div className="app-layout">
      <Navbar currentView={currentView} onNavigate={handleNavigate} />
      <main className="main-content">
        {currentView === 'dashboard' && (
          <DashboardView
            onNavigate={handleNavigate}
            onSelectQuotation={handleSelectQuotation}
            onViewResults={handleViewResults}
            showToast={showToast}
          />
        )}
        {(currentView === 'new-quote' || currentView === 'edit-quote') && (
          <QuotationEditorView
            quotationId={activeQuotationId}
            onNavigate={handleNavigate}
            onViewResults={handleViewResults}
            showToast={showToast}
          />
        )}
        {currentView === 'results' && (
          <CalculationResultsView
            quotationId={activeQuotationId}
            snapshot={activeSnapshot}
            onNavigate={handleNavigate}
            onEditQuotation={handleSelectQuotation}
            showToast={showToast}
          />
        )}
      </main>
      <footer className="app-footer">
        <div className="footer-container">
          <span>&copy; {new Date().getFullYear()} Mega Pipes Quotation System &bull; Enterprise Web Pilot</span>
          <span className="footer-links">
            <span className="environment-badge">Live Render Production Backend</span>
          </span>
        </div>
      </footer>
      <NotificationToast toast={toast} onDismiss={dismissToast} />
    </div>
  );
}

export default function App() {
  return (
    <AuthProvider>
      <MainContent />
    </AuthProvider>
  );
}
