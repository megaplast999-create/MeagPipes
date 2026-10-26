"""Logistics and Pincode calculation API endpoints for Mega Pipes."""

from __future__ import annotations

from typing import Optional
from fastapi import APIRouter, HTTPException, Query, status
from pydantic import BaseModel, Field

from app.services.logistics import resolve_delivery_logistics, is_valid_pin, extract_pin_from_text

router = APIRouter(prefix="/api/logistics", tags=["Logistics"])


class PincodeLookupResponse(BaseModel):
    valid: bool
    pin: str
    office_name: str
    district: str
    state: str
    distance_km: float
    trip_rate: float
    rate_per_km: float
    source: str
    plant_id: str
    plant_name: str
    formatted_address: str


class ResolveLogisticsRequest(BaseModel):
    pincode: Optional[str] = Field(None, description="6-digit Indian PIN code")
    delivery_location: Optional[str] = Field(None, description="Full delivery address or city")
    calculation_mode: str = Field(default="REFERENCE", description="REFERENCE, NORMAL, or LEGACY")


@router.get("/pincode/{pin}", response_model=PincodeLookupResponse)
async def lookup_pincode(
    pin: str,
    mode: str = Query(default="REFERENCE", description="Calculation mode: REFERENCE, NORMAL, or LEGACY"),
) -> PincodeLookupResponse:
    """Looks up delivery PIN code, extracts location details, and calculates road distance and trip rate."""
    clean_pin = pin.strip()
    if not is_valid_pin(clean_pin):
        extracted = extract_pin_from_text(clean_pin)
        if not extracted:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"'{clean_pin}' is not a valid 6-digit Indian PIN code.",
            )
        clean_pin = extracted

    result = resolve_delivery_logistics(clean_pin, mode=mode)
    return PincodeLookupResponse(**result)


@router.post("/resolve", response_model=PincodeLookupResponse)
async def resolve_logistics(req: ResolveLogisticsRequest) -> PincodeLookupResponse:
    """Resolves distance and trip rate from provided pincode or address string."""
    pin_or_addr = req.pincode or req.delivery_location
    if not pin_or_addr:
        pin_or_addr = "825322"  # Bagodar default

    result = resolve_delivery_logistics(pin_or_addr, mode=req.calculation_mode)
    return PincodeLookupResponse(**result)
