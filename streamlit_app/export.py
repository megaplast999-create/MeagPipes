"""Excel export service and download trigger for Streamlit application."""

from __future__ import annotations

import streamlit as st
from typing import Tuple, Optional
try:
    from .api_client import api_export_excel, APIError
except (ImportError, ValueError):
    from api_client import api_export_excel, APIError



def download_quotation_excel(quotation_id: int) -> Optional[Tuple[bytes, str]]:
    """Retrieves and returns the binary Excel workbook from the authoritative backend endpoint."""
    try:
        content, filename = api_export_excel(quotation_id)
        if not content:
            st.error("Received empty export file from server.")
            return None
        return content, filename
    except APIError as err:
        st.error(f"Failed to export quotation #{quotation_id}: {err.message}")
        return None
    except Exception as err:
        st.error(f"Unexpected export error: {err}")
        return None


def render_export_button(quotation_id: int, label: str = "Download Excel (.xlsx)", key: Optional[str] = None) -> None:
    """Renders a self-contained Streamlit download button for an existing quotation."""
    res = download_quotation_excel(quotation_id)
    if res:
        content, filename = res
        st.download_button(
            label=label,
            data=content,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=key or f"export_btn_{quotation_id}",
            use_container_width=True,
        )
