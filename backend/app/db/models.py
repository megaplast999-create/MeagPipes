"""SQLAlchemy 2.0 database models for users, quotations, snapshots, and audit events."""

from __future__ import annotations

import enum
from datetime import datetime
from typing import List, Optional

from sqlalchemy import (
    Boolean,
    Column,
    DateTime,
    Enum,
    Float,
    ForeignKey,
    Index,
    Integer,
    JSON,
    String,
    Text,
    func,
)
from sqlalchemy.orm import relationship

from app.db.session import Base


class QuotationStatus(str, enum.Enum):
    """Quotation lifecycle states."""
    DRAFT = "DRAFT"
    CALCULATING = "CALCULATING"
    CALCULATED = "CALCULATED"
    EXPORT_READY = "EXPORT_READY"
    FAILED = "FAILED"


class User(Base):
    """User account model for pilot authorization."""
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    email = Column(String(255), unique=True, index=True, nullable=False)
    hashed_password = Column(String(255), nullable=False)
    full_name = Column(String(255), nullable=True)
    is_active = Column(Boolean, default=True, nullable=False)
    is_superuser = Column(Boolean, default=False, nullable=False)
    created_at = Column(DateTime(timezone=True), server_default=func.now(), nullable=False)
    updated_at = Column(DateTime(timezone=True), server_default=func.now(), onupdate=func.now(), nullable=False)

    quotations = relationship("Quotation", back_populates="owner", cascade="all, delete-orphan")
    audit_events = relationship("AuditEvent", back_populates="user")


class Quotation(Base):
    """Header record for commercial pipe quotations."""
    __tablename__ = "quotations"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    enquiry_no = Column(String(100), index=True, nullable=False)
    customer_name = Column(String(255), nullable=False)
    project_name = Column(String(255), nullable=True)
    delivery_location = Column(String(255), nullable=True)
    distance_km = Column(Float, default=0.0, nullable=False)
    trip_rate = Column(Float, default=98000.0, nullable=False)
    status = Column(Enum(QuotationStatus), default=QuotationStatus.DRAFT, nullable=False, index=True)
    calculation_mode = Column(String(50), default="REFERENCE", nullable=False)
    version = Column(Integer, default=1, nullable=False)
    user_id = Column(Integer, ForeignKey("users.id", ondelete="CASCADE"), nullable=False, index=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now(), nullable=False)
    updated_at = Column(DateTime(timezone=True), server_default=func.now(), onupdate=func.now(), nullable=False)

    owner = relationship("User", back_populates="quotations")
    lines = relationship("QuotationLine", back_populates="quotation", cascade="all, delete-orphan", order_by="QuotationLine.line_number")
    snapshots = relationship("CalculationSnapshot", back_populates="quotation", cascade="all, delete-orphan", order_by="CalculationSnapshot.version.desc()")
    audit_events = relationship("AuditEvent", back_populates="quotation")

    __table_args__ = (
        Index("ix_quotations_user_id_status", "user_id", "status"),
    )


class QuotationLine(Base):
    """Individual line item representing pipe specification, pricing inputs, and override values."""
    __tablename__ = "quotation_lines"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    quotation_id = Column(Integer, ForeignKey("quotations.id", ondelete="CASCADE"), nullable=False, index=True)
    line_number = Column(Integer, nullable=False)
    size_mm = Column(Float, nullable=False)
    thickness_mm = Column(Float, nullable=False)
    basis = Column(String(10), default="ID", nullable=False)  # "ID" or "OD"
    quantity = Column(Float, nullable=False)
    unit = Column(String(20), default="Metres", nullable=False)  # "Metres", "MT", "Nos"
    pipe_length_m = Column(Float, default=12.0, nullable=False)
    pipes_per_trailer = Column(Float, default=4.0, nullable=False)
    has_external = Column(Boolean, default=True, nullable=False)
    has_internal = Column(Boolean, default=True, nullable=False)
    steel_rate = Column(Float, default=61000.0, nullable=False)
    grade_extra = Column(Float, default=0.0, nullable=False)
    yield_pct = Column(Float, default=2.0, nullable=False)
    conversion = Column(Float, default=5000.0, nullable=False)
    testing = Column(Float, default=0.0, nullable=False)
    margin = Column(Float, default=0.0, nullable=False)
    ext_rate = Column(Float, default=500.0, nullable=False)
    int_rate = Column(Float, default=275.0, nullable=False)
    handling = Column(Float, default=100.0, nullable=False)
    financial_pct = Column(Float, default=1.5, nullable=False)
    gst_pct = Column(Float, default=18.0, nullable=False)
    manual_fc = Column(Float, default=0.0, nullable=False)
    manual_selling_price = Column(Float, default=0.0, nullable=False)
    override_reason = Column(String(255), default="", nullable=False)
    created_at = Column(DateTime(timezone=True), server_default=func.now(), nullable=False)
    updated_at = Column(DateTime(timezone=True), server_default=func.now(), onupdate=func.now(), nullable=False)

    quotation = relationship("Quotation", back_populates="lines")

    __table_args__ = (
        Index("ix_quotation_lines_quote_line", "quotation_id", "line_number", unique=True),
    )


class CalculationSnapshot(Base):
    """Immutable audit snapshot of calculation results, preserving commercial fidelity across versions."""
    __tablename__ = "calculation_snapshots"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    quotation_id = Column(Integer, ForeignKey("quotations.id", ondelete="CASCADE"), nullable=False, index=True)
    version = Column(Integer, nullable=False)
    calculation_mode = Column(String(50), nullable=False)
    inputs_snapshot = Column(JSON, nullable=False)
    results_snapshot = Column(JSON, nullable=False)
    subtotal_excl_gst = Column(Float, nullable=False)
    gst_amount = Column(Float, nullable=False)
    grand_total = Column(Float, nullable=False)
    total_metres = Column(Float, default=0.0, nullable=False)
    total_mt = Column(Float, default=0.0, nullable=False)
    total_trips = Column(Float, default=0.0, nullable=False)
    calculated_by_id = Column(Integer, ForeignKey("users.id", ondelete="SET NULL"), nullable=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now(), nullable=False)

    quotation = relationship("Quotation", back_populates="snapshots")
    calculated_by = relationship("User")

    __table_args__ = (
        Index("ix_snapshots_quote_version", "quotation_id", "version", unique=True),
    )


class AuditEvent(Base):
    """Audit log tracking critical domain operations without leaking passwords or secrets."""
    __tablename__ = "audit_events"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    event_type = Column(String(100), index=True, nullable=False)
    quotation_id = Column(Integer, ForeignKey("quotations.id", ondelete="SET NULL"), nullable=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id", ondelete="SET NULL"), nullable=True, index=True)
    client_ip = Column(String(50), nullable=True)
    details = Column(JSON, nullable=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now(), nullable=False)

    quotation = relationship("Quotation", back_populates="audit_events")
    user = relationship("User", back_populates="audit_events")
