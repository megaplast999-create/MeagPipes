import React, { createContext, useContext, useState, useEffect, useCallback } from 'react';
import { authAPI, APIError } from '../api/client';

const AuthContext = createContext(null);

export function AuthProvider({ children }) {
  const [user, setUser] = useState(null);
  const [isLoading, setIsLoading] = useState(true);
  const [authError, setAuthError] = useState(null);

  // Restore session on application mount
  const checkSession = useCallback(async () => {
    setIsLoading(true);
    setAuthError(null);
    try {
      const res = await authAPI.getMe();
      if (res && res.authenticated && res.user) {
        setUser(res.user);
      } else {
        setUser(null);
      }
    } catch {
      setUser(null);
    } finally {
      setIsLoading(false);
    }
  }, []);

  useEffect(() => {
    checkSession();
  }, [checkSession]);

  const login = async (email, password) => {
    setAuthError(null);
    try {
      const loggedUser = await authAPI.login(email, password);
      setUser(loggedUser);
      return loggedUser;
    } catch (err) {
      const message = err instanceof APIError ? err.message : 'Login failed. Please check credentials.';
      setAuthError(message);
      throw err;
    }
  };

  const logout = async () => {
    try {
      await authAPI.logout();
    } catch {
      // Clean state even if logout request fails
    } finally {
      setUser(null);
      setAuthError(null);
    }
  };

  const value = {
    user,
    isAuthenticated: Boolean(user),
    isLoading,
    authError,
    setAuthError,
    login,
    logout,
    checkSession,
  };

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
}

export function useAuth() {
  const context = useContext(AuthContext);
  if (!context) {
    throw new Error('useAuth must be used within an AuthProvider');
  }
  return context;
}
