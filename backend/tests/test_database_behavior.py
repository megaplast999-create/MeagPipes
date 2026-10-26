"""PostgreSQL schema compatibility, constraints, cascades, and transaction behavior tests."""

import json
import pytest
from sqlalchemy import create_engine, select
from sqlalchemy.dialects import postgresql
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session
from sqlalchemy.schema import CreateTable, CreateIndex

from app.db.models import (
    AuditEvent,
    Base,
    CalculationSnapshot,
    Quotation,
    QuotationLine,
    QuotationStatus,
    User,
)
from app.services.calculation import CalculationService


class TestDatabaseBehaviorAndPostgresCompatibility:
    """Validates PostgreSQL DDL compilation, constraints, cascades, and immutability."""

    def test_postgresql_ddl_compilation(self):
        """Verifies all SQLAlchemy models compile cleanly against PostgreSQL dialect."""
        pg_dialect = postgresql.dialect()

        for model in (User, Quotation, QuotationLine, CalculationSnapshot, AuditEvent):
            create_stmt = str(CreateTable(model.__table__).compile(dialect=pg_dialect))
            assert "CREATE TABLE" in create_stmt
            assert model.__tablename__ in create_stmt

        # Verify PostgreSQL-specific type compilation
        snap_ddl = str(CreateTable(CalculationSnapshot.__table__).compile(dialect=pg_dialect))
        assert "JSON" in snap_ddl
        assert "TIMESTAMP WITH TIME ZONE" in snap_ddl

    def test_unique_email_constraint(self, db_session: Session, test_user: User):
        """Ensures duplicate user emails are rejected with IntegrityError."""
        dup_user = User(
            email=test_user.email,
            hashed_password="some_hashed_password",
            full_name="Duplicate User",
            is_active=True,
        )
        db_session.add(dup_user)
        with pytest.raises(IntegrityError):
            db_session.commit()
        db_session.rollback()

    def test_quotation_line_unique_constraint(self, db_session: Session, test_user: User):
        """Ensures (quotation_id, line_number) unique constraint is strictly enforced."""
        quote = Quotation(
            enquiry_no="ENQ-UNIQUE-01",
            customer_name="Unique Customer",
            user_id=test_user.id,
            status=QuotationStatus.DRAFT,
        )
        db_session.add(quote)
        db_session.flush()

        line1 = QuotationLine(
            quotation_id=quote.id,
            line_number=1,
            size_mm=1226.0,
            thickness_mm=8.4,
            basis="ID",
            quantity=100.0,
        )
        line2 = QuotationLine(
            quotation_id=quote.id,
            line_number=1,  # Duplicate line number on same quote
            size_mm=1300.0,
            thickness_mm=8.4,
            basis="ID",
            quantity=200.0,
        )
        db_session.add(line1)
        db_session.add(line2)
        with pytest.raises(IntegrityError):
            db_session.commit()
        db_session.rollback()

    def test_foreign_key_cascade_deletion(self, db_session: Session):
        """Verifies CASCADE deletion: deleting quotation deletes all its lines and snapshots."""
        user = User(
            email="cascade_test@megapipes.internal",
            hashed_password="hash",
            full_name="Cascade User",
        )
        db_session.add(user)
        db_session.flush()

        quote = Quotation(
            enquiry_no="ENQ-CASCADE-01",
            customer_name="Cascade Customer",
            user_id=user.id,
            status=QuotationStatus.DRAFT,
        )
        db_session.add(quote)
        db_session.flush()

        line = QuotationLine(
            quotation_id=quote.id,
            line_number=1,
            size_mm=1226.0,
            thickness_mm=8.4,
            basis="ID",
            quantity=100.0,
        )
        snap = CalculationSnapshot(
            quotation_id=quote.id,
            version=1,
            calculation_mode="REFERENCE",
            inputs_snapshot={"test": 1},
            results_snapshot={"test": 2},
            subtotal_excl_gst=1000.0,
            gst_amount=180.0,
            grand_total=1180.0,
            calculated_by_id=user.id,
        )
        db_session.add(line)
        db_session.add(snap)
        db_session.commit()

        # Delete quotation
        db_session.delete(quote)
        db_session.commit()

        # Verify child lines and snapshots were deleted
        remaining_lines = db_session.query(QuotationLine).filter(QuotationLine.quotation_id == quote.id).all()
        assert len(remaining_lines) == 0

        remaining_snaps = db_session.query(CalculationSnapshot).filter(CalculationSnapshot.quotation_id == quote.id).all()
        assert len(remaining_snaps) == 0

    def test_snapshot_immutability(self, db_session: Session, test_user: User):
        """Verifies calculation snapshots cannot have conflicting duplicates for the same version."""
        quote = Quotation(
            enquiry_no="ENQ-IMMUTABLE-01",
            customer_name="Immutable Customer",
            user_id=test_user.id,
            status=QuotationStatus.DRAFT,
        )
        db_session.add(quote)
        db_session.flush()

        snap1 = CalculationSnapshot(
            quotation_id=quote.id,
            version=1,
            calculation_mode="REFERENCE",
            inputs_snapshot={"line_count": 1},
            results_snapshot={"subtotal": 100000.0},
            subtotal_excl_gst=100000.0,
            gst_amount=18000.0,
            grand_total=118000.0,
            calculated_by_id=test_user.id,
        )
        db_session.add(snap1)
        db_session.commit()

        # Attempt to insert another snapshot with the same quotation_id and version -> fails unique index
        snap2 = CalculationSnapshot(
            quotation_id=quote.id,
            version=1,  # Duplicate version
            calculation_mode="REFERENCE",
            inputs_snapshot={"line_count": 1, "tampered": True},
            results_snapshot={"subtotal": 99999.0},
            subtotal_excl_gst=99999.0,
            gst_amount=17999.0,
            grand_total=117998.0,
            calculated_by_id=test_user.id,
        )
        db_session.add(snap2)
        with pytest.raises(IntegrityError):
            db_session.commit()
        db_session.rollback()

    def test_transaction_rollback_cleanliness(self, db_session: Session, test_user: User):
        """Verifies transactional rollback leaves no orphan data in database."""
        initial_count = db_session.query(Quotation).count()

        try:
            quote = Quotation(
                enquiry_no="ENQ-ROLLBACK-01",
                customer_name="Rollback Cust",
                user_id=test_user.id,
            )
            db_session.add(quote)
            db_session.flush()

            # Trigger deliberate database integrity violation (duplicate email)
            dup_user = User(
                email=test_user.email,  # Duplicate email
                hashed_password="some_password_hash",
            )
            db_session.add(dup_user)
            db_session.commit()
        except IntegrityError:
            db_session.rollback()

        final_count = db_session.query(Quotation).count()
        assert final_count == initial_count
