/**
 * Mega Pipes Quotation System - Production API Client
 * Handles authenticated session cookies, CSRF tokens, error parsing, and Excel exports.
 */

// Helper to extract a cookie value by name in the browser
export function getCookie(name) {
  if (typeof document === 'undefined') return null;
  const match = document.cookie.match(new RegExp('(^|;\\s*)(' + name + ')=([^;]*)'));
  return match ? decodeURIComponent(match[3]) : null;
}

// API error class with structured status code and backend details
export class APIError extends Error {
  constructor(message, status, details = null) {
    super(message);
    this.name = 'APIError';
    this.status = status;
    this.details = details;
  }
}

/**
 * Core fetch wrapper that attaches credentials and CSRF tokens automatically.
 */
export async function apiRequest(endpoint, options = {}) {
  const url = endpoint.startsWith('/') ? endpoint : `/${endpoint}`;
  
  const headers = {
    Accept: 'application/json',
    ...(options.headers || {}),
  };

  // If payload is JSON and body is an object, serialize it
  if (options.body && typeof options.body === 'object' && !(options.body instanceof FormData) && !(options.body instanceof Blob)) {
    headers['Content-Type'] = 'application/json';
    options.body = JSON.stringify(options.body);
  }

  // Attach CSRF token on mutating HTTP methods
  const method = (options.method || 'GET').toUpperCase();
  if (['POST', 'PUT', 'PATCH', 'DELETE'].includes(method)) {
    const csrfToken = getCookie('megapipes_csrf');
    if (csrfToken) {
      headers['X-CSRF-Token'] = csrfToken;
    }
  }

  const config = {
    ...options,
    method,
    headers,
    credentials: 'include', // Crucial for session cookies across Vercel rewrite
  };

  let response;
  try {
    response = await fetch(url, config);
  } catch (err) {
    throw new APIError('Network error connecting to backend. Please check connection.', 0, err);
  }

  // Handle binary blob responses (e.g. Excel export)
  if (options.responseType === 'blob') {
    if (!response.ok) {
      let errorMsg = `Server error (${response.status})`;
      try {
        const errJson = await response.json();
        errorMsg = errJson.detail || errorMsg;
      } catch {
        // Ignored
      }
      throw new APIError(errorMsg, response.status);
    }
    const blob = await response.blob();
    const disposition = response.headers.get('Content-Disposition') || '';
    let filename = 'Mega_Pipes_Quotation.xlsx';
    const filenameMatch = disposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
    if (filenameMatch && filenameMatch[1]) {
      filename = filenameMatch[1].replace(/['"]/g, '').trim();
    }
    return { blob, filename };
  }

  // Handle standard JSON responses
  let data = null;
  const contentType = response.headers.get('Content-Type') || '';
  if (contentType.includes('application/json')) {
    try {
      data = await response.json();
    } catch {
      data = null;
    }
  } else {
    data = await response.text();
  }

  if (!response.ok) {
    let detailMessage = `Request failed with status ${response.status}`;
    if (data && typeof data === 'object' && data.detail) {
      if (Array.isArray(data.detail)) {
        detailMessage = data.detail.map((d) => d.msg || JSON.stringify(d)).join(', ');
      } else {
        detailMessage = String(data.detail);
      }
    } else if (typeof data === 'string' && data) {
      detailMessage = data;
    }
    throw new APIError(detailMessage, response.status, data);
  }

  return data;
}

// Authentication API methods
export const authAPI = {
  login: (email, password) =>
    apiRequest('/api/auth/login', {
      method: 'POST',
      body: { email, password },
    }),

  getMe: () =>
    apiRequest('/api/auth/me', {
      method: 'GET',
    }),

  logout: () =>
    apiRequest('/api/auth/logout', {
      method: 'POST',
    }),
};

// Quotations API methods
export const quotationsAPI = {
  list: () =>
    apiRequest('/api/quotations', {
      method: 'GET',
    }),

  get: (id) =>
    apiRequest(`/api/quotations/${id}`, {
      method: 'GET',
    }),

  create: (quotationData) =>
    apiRequest('/api/quotations', {
      method: 'POST',
      body: quotationData,
    }),

  update: (id, quotationData) =>
    apiRequest(`/api/quotations/${id}`, {
      method: 'PUT',
      body: quotationData,
    }),

  calculate: (id) =>
    apiRequest(`/api/quotations/${id}/calculate`, {
      method: 'POST',
    }),

  getSnapshot: (id) =>
    apiRequest(`/api/quotations/${id}/snapshot`, {
      method: 'GET',
    }),

  exportExcel: async (id) => {
    const { blob, filename } = await apiRequest(`/api/quotations/${id}/export`, {
      method: 'GET',
      responseType: 'blob',
    });

    // Trigger browser download
    const downloadUrl = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => window.URL.revokeObjectURL(downloadUrl), 5000);
    return filename;
  },
};

export const logisticsAPI = {
  lookupPincode: (pin, mode = 'REFERENCE') =>
    apiRequest(`/api/logistics/pincode/${encodeURIComponent(pin)}?mode=${encodeURIComponent(mode)}`, {
      method: 'GET',
    }),

  resolve: (pincode, deliveryLocation, mode = 'REFERENCE') =>
    apiRequest('/api/logistics/resolve', {
      method: 'POST',
      body: {
        pincode,
        delivery_location: deliveryLocation,
        calculation_mode: mode,
      },
    }),
};

