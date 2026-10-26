"""Test fixtures, database isolation, and mock Redis setup."""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Generator

import fakeredis
import pytest
from fastapi.testclient import TestClient
from sqlalchemy import create_engine
from sqlalchemy.orm import Session, sessionmaker

# Ensure backend root is on sys.path
BACKEND_DIR = Path(__file__).resolve().parent.parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.insert(0, str(BACKEND_DIR))

# Locate Quotation System root
ROOT_DIR = BACKEND_DIR.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

TEST_DB_FILE = BACKEND_DIR / "test_backend_isolated.db"
TEST_DATABASE_URL = f"sqlite:///{TEST_DB_FILE}"

os.environ["DATABASE_URL"] = TEST_DATABASE_URL
os.environ["ENVIRONMENT"] = "test"

from app.core.config import settings
settings.DATABASE_URL = TEST_DATABASE_URL
settings.ENVIRONMENT = "test"

from app.core.rate_limit import set_redis_client
from app.core.security import create_session_token, hash_password
from app.db.models import Base, Quotation, QuotationLine, User
from app.db.session import get_db
from app.main import app

test_engine = create_engine(
    TEST_DATABASE_URL,
    connect_args={"check_same_thread": False},
)
TestingSessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=test_engine)


@pytest.fixture(scope="session", autouse=True)
def setup_test_db():
    """Initializes isolated database tables for test session and cleans up after."""
    if TEST_DB_FILE.exists():
        try:
            TEST_DB_FILE.unlink()
        except Exception:
            pass
    Base.metadata.create_all(bind=test_engine)
    yield
    Base.metadata.drop_all(bind=test_engine)
    if TEST_DB_FILE.exists():
        try:
            TEST_DB_FILE.unlink()
        except Exception:
            pass


@pytest.fixture
def fake_redis():
    """Provides an isolated in-memory Redis mock."""
    redis_instance = fakeredis.FakeRedis(decode_responses=True)
    set_redis_client(redis_instance)
    yield redis_instance
    redis_instance.flushall()
    set_redis_client(None)


@pytest.fixture
def db_session() -> Generator[Session, None, None]:
    """Provides a transactional database session per test."""
    session = TestingSessionLocal()
    try:
        yield session
    finally:
        session.close()


@pytest.fixture
def client(db_session: Session, fake_redis) -> Generator[TestClient, None, None]:
    """TestClient overriding get_db dependency."""
    def override_get_db():
        try:
            yield db_session
        finally:
            pass

    app.dependency_overrides[get_db] = override_get_db
    with TestClient(app) as test_client:
        yield test_client
    app.dependency_overrides.clear()


@pytest.fixture
def test_user(db_session: Session) -> User:
    """Creates a standard test pilot user."""
    user = db_session.query(User).filter(User.email == "pilot@megapipes.internal").first()
    if not user:
        user = User(
            email="pilot@megapipes.internal",
            hashed_password=hash_password("PilotPassword123!"),
            full_name="Mega Pipes Test User",
            is_active=True,
            is_superuser=False,
        )
        db_session.add(user)
        db_session.commit()
        db_session.refresh(user)
    else:
        user.hashed_password = hash_password("PilotPassword123!")
        user.is_superuser = False
        db_session.commit()
        db_session.refresh(user)
    return user


@pytest.fixture
def test_other_user(db_session: Session) -> User:
    """Creates a second distinct user to verify cross-user authorization barriers."""
    user = db_session.query(User).filter(User.email == "other@megapipes.internal").first()
    if not user:
        user = User(
            email="other@megapipes.internal",
            hashed_password=hash_password("OtherPassword123!"),
            full_name="Unauthorized Other User",
            is_active=True,
            is_superuser=False,
        )
        db_session.add(user)
        db_session.commit()
        db_session.refresh(user)
    else:
        user.is_superuser = False
        db_session.commit()
        db_session.refresh(user)
    return user


@pytest.fixture
def authenticated_client(client: TestClient, test_user: User) -> TestClient:
    """TestClient pre-authenticated with valid session cookie and CSRF header."""
    token, csrf = create_session_token(
        user_id=test_user.id,
        email=test_user.email,
        idle_minutes=30,
        absolute_hours=8,
    )
    client.cookies.set(settings.COOKIE_NAME, token)
    client.cookies.set(settings.CSRF_COOKIE_NAME, csrf)
    client.headers.update({"X-CSRF-Token": csrf})
    return client
