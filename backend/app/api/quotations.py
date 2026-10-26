"""Quotation management, calculation, snapshotting, and Excel export API."""

from __future__ import annotations

import io
import logging
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, Depends, HTTPException, Query, Request, Response, status
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field, field_validator
from sqlalchemy.orm import Session

from app.api.auth import get_current_user
from app.core.config import settings
from app.core.rate_limit import RateLimiter, get_client_ip
from app.db.models import (
    AuditEvent,
    CalculationSnapshot,
    Quotation,
    QuotationLine,
    QuotationStatus,
    User,
)
from app.db.session import get_db
from app.services.calculation import CalculationService
from app.services.export_excel import ExcelExportService
from app.services.logistics import resolve_delivery_logistics

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/quotations", tags=["Quotations"])


# --- Pydantic Request / Response Models ---

class QuotationLineInput(BaseModel):
    line_number: int = Field(..., ge=1, le=100)
    size_mm: float = Field(..., gt=0, le=5000, description="Nominal diameter in mm")
    thickness_mm: float = Field(..., gt=0, le=100, description="Wall thickness in mm")
    basis: str = Field(default="ID", description="Sizing basis: OD or ID")
    quantity: float = Field(..., gt=0, description="Quantity")
    unit: str = Field(default="Metres", description="Unit: Metres, MT, or Nos")
    pipe_length_m: float = Field(default=12.0, ge=6.0, le=13.0)
    pipes_per_trailer: float = Field(default=4.0, ge=1.0, le=20.0)
    has_external: bool = True
    has_internal: bool = True
    steel_rate: float = Field(default=61000.0, ge=0)
    grade_extra: float = Field(default=0.0, ge=0)
    yield_pct: float = Field(default=2.0, ge=0, le=20)
    conversion: float = Field(default=5000.0, ge=0)
    testing: float = Field(default=0.0, ge=0)
    margin: float = Field(default=0.0, ge=0)
    ext_rate: float = Field(default=500.0, ge=0)
    int_rate: float = Field(default=275.0, ge=0)
    handling: float = Field(default=100.0, ge=0)
    financial_pct: float = Field(default=1.5, ge=0, le=20)
    gst_pct: float = Field(default=18.0, ge=0, le=100)
    manual_fc: float = Field(default=0.0, ge=0)
    manual_selling_price: float = Field(default=0.0, ge=0)
    override_reason: str = Field(default="", max_length=255)

    @field_validator("basis")
    @classmethod
    def validate_basis(cls, v: str) -> str:
        v_upper = v.upper().strip()
        if v_upper not in ("OD", "ID"):
            raise ValueError("basis must be 'OD' or 'ID'")
        return v_upper

    @field_validator("unit")
    @classmethod
    def validate_unit(cls, v: str) -> str:
        v_norm = v.capitalize().strip()
        if v_norm not in ("Metres", "Mt", "Nos"):
            raise ValueError("unit must be 'Metres', 'MT', or 'Nos'")
        return "MT" if v_norm == "Mt" else v_norm


class QuotationCreateRequest(BaseModel):
    enquiry_no: str = Field(..., min_length=1, max_length=100)
    customer_name: str = Field(..., min_length=1, max_length=255)
    project_name: Optional[str] = Field(None, max_length=255)
    delivery_location: Optional[str] = Field(None, max_length=255)
    distance_km: float = Field(default=0.0, ge=0)
    trip_rate: float = Field(default=98000.0, ge=0)
    calculation_mode: str = Field(default="REFERENCE")
    lines: List[QuotationLineInput] = Field(default_factory=list)


class QuotationUpdateRequest(BaseModel):
    enquiry_no: Optional[str] = Field(None, min_length=1, max_length=100)
    customer_name: Optional[str] = Field(None, min_length=1, max_length=255)
    project_name: Optional[str] = Field(None, max_length=255)
    delivery_location: Optional[str] = Field(None, max_length=255)
    distance_km: Optional[float] = Field(None, ge=0)
    trip_rate: Optional[float] = Field(None, ge=0)
    calculation_mode: Optional[str] = None
    lines: Optional[List[QuotationLineInput]] = None


class QuotationLineResponse(BaseModel):
    id: int
    line_number: int
    size_mm: float
    thickness_mm: float
    basis: str
    quantity: float
    unit: str
    pipe_length_m: float
    pipes_per_trailer: float
    has_external: bool
    has_internal: bool
    steel_rate: float
    grade_extra: float
    yield_pct: float
    conversion: float
    testing: float
    margin: float
    ext_rate: float
    int_rate: float
    handling: float
    financial_pct: float
    gst_pct: float
    manual_fc: float
    manual_selling_price: float
    override_reason: str

    model_config = {"from_attributes": True}


class QuotationResponse(BaseModel):
    id: int
    enquiry_no: str
    customer_name: str
    project_name: Optional[str]
    delivery_location: Optional[str]
    distance_km: float
    trip_rate: float
    status: str
    calculation_mode: str
    version: int
    lines: List[QuotationLineResponse] = []

    model_config = {"from_attributes": True}


class CalculationSnapshotResponse(BaseModel):
    id: int
    quotation_id: int
    version: int
    calculation_mode: str
    subtotal_excl_gst: float
    gst_amount: float
    grand_total: float
    total_metres: float
    total_mt: float
    total_trips: float
    results_snapshot: Dict[str, Any]

    model_config = {"from_attributes": True}


# Rate limiters
calc_limiter = RateLimiter(
    action="calculate",
    max_requests=settings.RATE_LIMIT_CALCULATE_PER_MINUTE,
    window_seconds=60,
)

export_limiter = RateLimiter(
    action="export",
    max_requests=settings.RATE_LIMIT_EXPORT_PER_MINUTE,
    window_seconds=60,
)


def get_authorized_quotation(quotation_id: int, user: User, db: Session) -> Quotation:
    """Retrieves a quotation and verifies authorized ownership."""
    quote = db.query(Quotation).filter(Quotation.id == quotation_id).first()
    if not quote:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Quotation not found",
        )
    if quote.user_id != user.id and not user.is_superuser:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Not authorized to access this quotation",
        )
    return quote


# --- Endpoints ---

@router.post("", response_model=QuotationResponse, status_code=status.HTTP_201_CREATED)
async def create_quotation(
    req: QuotationCreateRequest,
    request: Request,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
) -> Any:
    """Creates a new quotation draft belonging to the authenticated user."""
    dist_km = req.distance_km
    t_rate = req.trip_rate
    if dist_km <= 0 or t_rate <= 0:
        logistics = resolve_delivery_logistics(req.delivery_location, mode=req.calculation_mode)
        if dist_km <= 0:
            dist_km = logistics["distance_km"]
        if t_rate <= 0:
            t_rate = logistics["trip_rate"]

    quote = Quotation(
        enquiry_no=req.enquiry_no.strip(),
        customer_name=req.customer_name.strip(),
        project_name=req.project_name.strip() if req.project_name else None,
        delivery_location=req.delivery_location.strip() if req.delivery_location else None,
        distance_km=dist_km,
        trip_rate=t_rate,
        calculation_mode=req.calculation_mode,
        status=QuotationStatus.DRAFT,
        user_id=current_user.id,
    )
    db.add(quote)
    db.flush()

    for idx, line_in in enumerate(req.lines, start=1):
        line = QuotationLine(
            quotation_id=quote.id,
            line_number=idx,
            size_mm=line_in.size_mm,
            thickness_mm=line_in.thickness_mm,
            basis=line_in.basis,
            quantity=line_in.quantity,
            unit=line_in.unit,
            pipe_length_m=line_in.pipe_length_m,
            pipes_per_trailer=line_in.pipes_per_trailer,
            has_external=line_in.has_external,
            has_internal=line_in.has_internal,
            steel_rate=line_in.steel_rate,
            grade_extra=line_in.grade_extra,
            yield_pct=line_in.yield_pct,
            conversion=line_in.conversion,
            testing=line_in.testing,
            margin=line_in.margin,
            ext_rate=line_in.ext_rate,
            int_rate=line_in.int_rate,
            handling=line_in.handling,
            financial_pct=line_in.financial_pct,
            gst_pct=line_in.gst_pct,
            manual_fc=line_in.manual_fc,
            manual_selling_price=line_in.manual_selling_price,
            override_reason=line_in.override_reason,
        )
        db.add(line)

    audit = AuditEvent(
        event_type="QUOTE_CREATED",
        quotation_id=quote.id,
        user_id=current_user.id,
        client_ip=get_client_ip(request),
        details={"enquiry_no": quote.enquiry_no},
    )
    db.add(audit)
    db.commit()
    db.refresh(quote)
    return quote


@router.get("", response_model=List[QuotationResponse])
async def list_quotations(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
) -> Any:
    """Lists all quotations owned by the authenticated user."""
    quotes = db.query(Quotation).filter(Quotation.user_id == current_user.id).order_by(Quotation.id.desc()).all()
    return quotes


@router.get("/{id}", response_model=QuotationResponse)
async def get_quotation(
    id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
) -> Any:
    """Retrieves an authorized quotation with lines."""
    quote = get_authorized_quotation(id, current_user, db)
    return quote


@router.put("/{id}", response_model=QuotationResponse)
async def update_quotation(
    id: int,
    req: QuotationUpdateRequest,
    request: Request,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
) -> Any:
    """Updates an authorized draft quotation."""
    quote = get_authorized_quotation(id, current_user, db)

    if req.enquiry_no is not None:
        quote.enquiry_no = req.enquiry_no.strip()
    if req.customer_name is not None:
        quote.customer_name = req.customer_name.strip()
    if req.project_name is not None:
        quote.project_name = req.project_name.strip()
    if req.delivery_location is not None:
        quote.delivery_location = req.delivery_location.strip()
    if req.distance_km is not None:
        quote.distance_km = req.distance_km
    if req.trip_rate is not None:
        quote.trip_rate = req.trip_rate
    if req.calculation_mode is not None:
        quote.calculation_mode = req.calculation_mode

    if quote.distance_km <= 0 or quote.trip_rate <= 0:
        logistics = resolve_delivery_logistics(quote.delivery_location, mode=quote.calculation_mode)
        if quote.distance_km <= 0:
            quote.distance_km = logistics["distance_km"]
        if quote.trip_rate <= 0:
            quote.trip_rate = logistics["trip_rate"]

    # If lines provided, replace existing lines
    if req.lines is not None:
        db.query(QuotationLine).filter(QuotationLine.quotation_id == quote.id).delete()
        for idx, line_in in enumerate(req.lines, start=1):
            line = QuotationLine(
                quotation_id=quote.id,
                line_number=idx,
                size_mm=line_in.size_mm,
                thickness_mm=line_in.thickness_mm,
                basis=line_in.basis,
                quantity=line_in.quantity,
                unit=line_in.unit,
                pipe_length_m=line_in.pipe_length_m,
                pipes_per_trailer=line_in.pipes_per_trailer,
                has_external=line_in.has_external,
                has_internal=line_in.has_internal,
                steel_rate=line_in.steel_rate,
                grade_extra=line_in.grade_extra,
                yield_pct=line_in.yield_pct,
                conversion=line_in.conversion,
                testing=line_in.testing,
                margin=line_in.margin,
                ext_rate=line_in.ext_rate,
                int_rate=line_in.int_rate,
                handling=line_in.handling,
                financial_pct=line_in.financial_pct,
                gst_pct=line_in.gst_pct,
                manual_fc=line_in.manual_fc,
                manual_selling_price=line_in.manual_selling_price,
                override_reason=line_in.override_reason,
            )
            db.add(line)

    # Any modification resets status to DRAFT until recalculation
    quote.status = QuotationStatus.DRAFT

    audit = AuditEvent(
        event_type="QUOTE_UPDATED",
        quotation_id=quote.id,
        user_id=current_user.id,
        client_ip=get_client_ip(request),
        details={"enquiry_no": quote.enquiry_no},
    )
    db.add(audit)
    db.commit()
    db.refresh(quote)
    return quote


@router.post("/{id}/calculate", response_model=CalculationSnapshotResponse)
async def calculate_quotation(
    id: int,
    request: Request,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
    _limiter: None = Depends(calc_limiter),
) -> Any:
    """Calculates quotation server-side and creates an immutable snapshot of results."""
    quote = get_authorized_quotation(id, current_user, db)

    lines = db.query(QuotationLine).filter(QuotationLine.quotation_id == quote.id).order_by(QuotationLine.line_number).all()
    if not lines:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Quotation contains no line items to calculate",
        )

    # Ensure positive trip rate from logistics
    effective_trip_rate = quote.trip_rate
    if effective_trip_rate <= 0:
        logistics = resolve_delivery_logistics(quote.delivery_location, mode=quote.calculation_mode)
        effective_trip_rate = logistics["trip_rate"]
        quote.trip_rate = effective_trip_rate
        if quote.distance_km <= 0:
            quote.distance_km = logistics["distance_km"]

    # Convert line models to dict for calculation engine
    engine_inputs = []
    for l in lines:
        engine_inputs.append({
            "size_mm": l.size_mm,
            "thickness_mm": l.thickness_mm,
            "basis": l.basis,
            "quantity": l.quantity,
            "unit": l.unit,
            "pipe_length_m": l.pipe_length_m,
            "pipes_per_trailer": l.pipes_per_trailer,
            "has_external": l.has_external,
            "has_internal": l.has_internal,
            "steel_rate": l.steel_rate,
            "grade_extra": l.grade_extra,
            "yield_pct": l.yield_pct,
            "conversion": l.conversion,
            "testing": l.testing,
            "margin": l.margin,
            "ext_rate": l.ext_rate,
            "int_rate": l.int_rate,
            "handling": l.handling,
            "financial_pct": l.financial_pct,
            "gst_pct": l.gst_pct,
            "manual_fc": l.manual_fc,
            "manual_selling_price": l.manual_selling_price,
            "override_reason": l.override_reason,
            "trip_rate": effective_trip_rate,
        })

    # Execute server-side calculation via thin engine adapter
    try:
        calc_res = CalculationService.calculate_quotation(
            engine_inputs,
            trip_rate=effective_trip_rate,
            mode=quote.calculation_mode,
        )
    except Exception as err:
        logger.error(f"Calculation failed for quotation {quote.id}: {err}", exc_info=True)
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Calculation engine failed: {str(err)}",
        )

    # Increment version and persist immutable calculation snapshot
    new_version = quote.version + 1
    quote.version = new_version
    quote.status = QuotationStatus.CALCULATED

    snapshot = CalculationSnapshot(
        quotation_id=quote.id,
        version=new_version,
        calculation_mode=quote.calculation_mode,
        inputs_snapshot={"lines": engine_inputs, "trip_rate": quote.trip_rate},
        results_snapshot=calc_res,
        subtotal_excl_gst=calc_res["subtotal_excl_gst"],
        gst_amount=calc_res["gst_amount"],
        grand_total=calc_res["grand_total"],
        total_metres=calc_res.get("total_metres", 0.0),
        total_mt=calc_res.get("total_mt", 0.0),
        total_trips=calc_res.get("total_trips", 0.0),
        calculated_by_id=current_user.id,
    )
    db.add(snapshot)

    audit = AuditEvent(
        event_type="QUOTE_CALCULATED",
        quotation_id=quote.id,
        user_id=current_user.id,
        client_ip=get_client_ip(request),
        details={"version": new_version, "grand_total": calc_res["grand_total"]},
    )
    db.add(audit)
    db.commit()
    db.refresh(snapshot)

    return snapshot


@router.get("/{id}/snapshot", response_model=CalculationSnapshotResponse)
async def get_latest_snapshot(
    id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
) -> Any:
    """Retrieves the latest immutable calculation snapshot for an authorized quotation."""
    quote = get_authorized_quotation(id, current_user, db)
    snapshot = (
        db.query(CalculationSnapshot)
        .filter(CalculationSnapshot.quotation_id == quote.id)
        .order_by(CalculationSnapshot.version.desc())
        .first()
    )
    if not snapshot:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="No calculation snapshot exists for this quotation. Calculate it first.",
        )
    return snapshot


@router.get("/{id}/export")
async def export_quotation(
    id: int,
    request: Request,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
    _limiter: None = Depends(export_limiter),
) -> Any:
    """Generates and downloads a password-free, fully editable Excel quotation (.xlsx)."""
    quote = get_authorized_quotation(id, current_user, db)

    snapshot = (
        db.query(CalculationSnapshot)
        .filter(CalculationSnapshot.quotation_id == quote.id)
        .order_by(CalculationSnapshot.version.desc())
        .first()
    )
    if not snapshot:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Quotation must be calculated before export",
        )

    quote_dict = {
        "enquiry_no": quote.enquiry_no,
        "customer_name": quote.customer_name,
        "project_name": quote.project_name,
        "delivery_location": quote.delivery_location,
        "distance_km": quote.distance_km,
        "trip_rate": quote.trip_rate,
    }

    workbook_bytes = ExcelExportService.generate_quotation_workbook(
        quotation_data=quote_dict,
        calculation_results=snapshot.results_snapshot,
    )

    filename = f"Quotation_{quote.enquiry_no.replace('/', '_')}_v{snapshot.version}.xlsx"

    audit = AuditEvent(
        event_type="EXPORT_GENERATED",
        quotation_id=quote.id,
        user_id=current_user.id,
        client_ip=get_client_ip(request),
        details={"filename": filename},
    )
    db.add(audit)
    db.commit()

    return StreamingResponse(
        workbook_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
