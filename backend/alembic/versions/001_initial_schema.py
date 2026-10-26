"""initial_schema

Revision ID: 001_initial_schema
Revises: 
Create Date: 2026-10-26 12:00:00.000000

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision: str = "001_initial_schema"
down_revision: Union[str, None] = None
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    # 1. Users Table
    op.create_table(
        "users",
        sa.Column("id", sa.Integer(), autoincrement=True, nullable=False),
        sa.Column("email", sa.String(length=255), nullable=False),
        sa.Column("hashed_password", sa.String(length=255), nullable=False),
        sa.Column("full_name", sa.String(length=255), nullable=True),
        sa.Column("is_active", sa.Boolean(), server_default=sa.text("true"), nullable=False),
        sa.Column("is_superuser", sa.Boolean(), server_default=sa.text("false"), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_users_id", "users", ["id"])
    op.create_index("ix_users_email", "users", ["email"], unique=True)

    # 2. Quotations Table
    op.create_table(
        "quotations",
        sa.Column("id", sa.Integer(), autoincrement=True, nullable=False),
        sa.Column("enquiry_no", sa.String(length=100), nullable=False),
        sa.Column("customer_name", sa.String(length=255), nullable=False),
        sa.Column("project_name", sa.String(length=255), nullable=True),
        sa.Column("delivery_location", sa.String(length=255), nullable=True),
        sa.Column("distance_km", sa.Float(), server_default=sa.text("0.0"), nullable=False),
        sa.Column("trip_rate", sa.Float(), server_default=sa.text("98000.0"), nullable=False),
        sa.Column("status", sa.String(length=50), server_default="DRAFT", nullable=False),
        sa.Column("calculation_mode", sa.String(length=50), server_default="REFERENCE", nullable=False),
        sa.Column("version", sa.Integer(), server_default=sa.text("1"), nullable=False),
        sa.Column("user_id", sa.Integer(), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.ForeignKeyConstraint(["user_id"], ["users.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_quotations_id", "quotations", ["id"])
    op.create_index("ix_quotations_enquiry_no", "quotations", ["enquiry_no"])
    op.create_index("ix_quotations_user_id", "quotations", ["user_id"])
    op.create_index("ix_quotations_status", "quotations", ["status"])
    op.create_index("ix_quotations_user_id_status", "quotations", ["user_id", "status"])

    # 3. Quotation Lines Table
    op.create_table(
        "quotation_lines",
        sa.Column("id", sa.Integer(), autoincrement=True, nullable=False),
        sa.Column("quotation_id", sa.Integer(), nullable=False),
        sa.Column("line_number", sa.Integer(), nullable=False),
        sa.Column("size_mm", sa.Float(), nullable=False),
        sa.Column("thickness_mm", sa.Float(), nullable=False),
        sa.Column("basis", sa.String(length=10), server_default="ID", nullable=False),
        sa.Column("quantity", sa.Float(), nullable=False),
        sa.Column("unit", sa.String(length=20), server_default="Metres", nullable=False),
        sa.Column("pipe_length_m", sa.Float(), server_default=sa.text("12.0"), nullable=False),
        sa.Column("pipes_per_trailer", sa.Float(), server_default=sa.text("4.0"), nullable=False),
        sa.Column("has_external", sa.Boolean(), server_default=sa.text("true"), nullable=False),
        sa.Column("has_internal", sa.Boolean(), server_default=sa.text("true"), nullable=False),
        sa.Column("steel_rate", sa.Float(), server_default=sa.text("61000.0"), nullable=False),
        sa.Column("grade_extra", sa.Float(), server_default=sa.text("0.0"), nullable=False),
        sa.Column("yield_pct", sa.Float(), server_default=sa.text("2.0"), nullable=False),
        sa.Column("conversion", sa.Float(), server_default=sa.text("5000.0"), nullable=False),
        sa.Column("testing", sa.Float(), server_default=sa.text("0.0"), nullable=False),
        sa.Column("margin", sa.Float(), server_default=sa.text("0.0"), nullable=False),
        sa.Column("ext_rate", sa.Float(), server_default=sa.text("500.0"), nullable=False),
        sa.Column("int_rate", sa.Float(), server_default=sa.text("275.0"), nullable=False),
        sa.Column("handling", sa.Float(), server_default=sa.text("100.0"), nullable=False),
        sa.Column("financial_pct", sa.Float(), server_default=sa.text("1.5"), nullable=False),
        sa.Column("gst_pct", sa.Float(), server_default=sa.text("18.0"), nullable=False),
        sa.Column("manual_fc", sa.Float(), server_default=sa.text("0.0"), nullable=False),
        sa.Column("manual_selling_price", sa.Float(), server_default=sa.text("0.0"), nullable=False),
        sa.Column("override_reason", sa.String(length=255), server_default="", nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.ForeignKeyConstraint(["quotation_id"], ["quotations.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_quotation_lines_id", "quotation_lines", ["id"])
    op.create_index("ix_quotation_lines_quotation_id", "quotation_lines", ["quotation_id"])
    op.create_index("ix_quotation_lines_quote_line", "quotation_lines", ["quotation_id", "line_number"], unique=True)

    # 4. Calculation Snapshots Table
    op.create_table(
        "calculation_snapshots",
        sa.Column("id", sa.Integer(), autoincrement=True, nullable=False),
        sa.Column("quotation_id", sa.Integer(), nullable=False),
        sa.Column("version", sa.Integer(), nullable=False),
        sa.Column("calculation_mode", sa.String(length=50), nullable=False),
        sa.Column("inputs_snapshot", sa.JSON(), nullable=False),
        sa.Column("results_snapshot", sa.JSON(), nullable=False),
        sa.Column("subtotal_excl_gst", sa.Float(), nullable=False),
        sa.Column("gst_amount", sa.Float(), nullable=False),
        sa.Column("grand_total", sa.Float(), nullable=False),
        sa.Column("total_metres", sa.Float(), server_default=sa.text("0.0"), nullable=False),
        sa.Column("total_mt", sa.Float(), server_default=sa.text("0.0"), nullable=False),
        sa.Column("total_trips", sa.Float(), server_default=sa.text("0.0"), nullable=False),
        sa.Column("calculated_by_id", sa.Integer(), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.ForeignKeyConstraint(["calculated_by_id"], ["users.id"], ondelete="SET NULL"),
        sa.ForeignKeyConstraint(["quotation_id"], ["quotations.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_calculation_snapshots_id", "calculation_snapshots", ["id"])
    op.create_index("ix_calculation_snapshots_quotation_id", "calculation_snapshots", ["quotation_id"])
    op.create_index("ix_snapshots_quote_version", "calculation_snapshots", ["quotation_id", "version"], unique=True)

    # 5. Audit Events Table
    op.create_table(
        "audit_events",
        sa.Column("id", sa.Integer(), autoincrement=True, nullable=False),
        sa.Column("event_type", sa.String(length=100), nullable=False),
        sa.Column("quotation_id", sa.Integer(), nullable=True),
        sa.Column("user_id", sa.Integer(), nullable=True),
        sa.Column("client_ip", sa.String(length=50), nullable=True),
        sa.Column("details", sa.JSON(), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.ForeignKeyConstraint(["quotation_id"], ["quotations.id"], ondelete="SET NULL"),
        sa.ForeignKeyConstraint(["user_id"], ["users.id"], ondelete="SET NULL"),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_audit_events_id", "audit_events", ["id"])
    op.create_index("ix_audit_events_event_type", "audit_events", ["event_type"])
    op.create_index("ix_audit_events_quotation_id", "audit_events", ["quotation_id"])
    op.create_index("ix_audit_events_user_id", "audit_events", ["user_id"])


def downgrade() -> None:
    op.drop_table("audit_events")
    op.drop_table("calculation_snapshots")
    op.drop_table("quotation_lines")
    op.drop_table("quotations")
    op.drop_table("users")
