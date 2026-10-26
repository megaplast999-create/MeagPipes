/**
 * Financial & Commercial Number Formatting Utilities for Mega Pipes Quotation System
 */

export function formatINR(val, decimals = 2) {
  if (val === null || val === undefined || isNaN(val)) return '₹0.00';
  const num = Number(val);
  return '₹' + num.toLocaleString('en-IN', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

export function formatNumber(val, decimals = 2) {
  if (val === null || val === undefined || isNaN(val)) return '0';
  const num = Number(val);
  return num.toLocaleString('en-IN', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

export function formatWeight(mt) {
  if (mt === null || mt === undefined || isNaN(mt)) return '0.00 MT';
  return `${formatNumber(mt, 3)} MT`;
}

export function formatLength(m) {
  if (m === null || m === undefined || isNaN(m)) return '0.00 m';
  return `${formatNumber(m, 2)} m`;
}
