import os, json, datetime
from sqlalchemy import Column, Integer, String, Numeric, Boolean, DateTime, ForeignKey, Text, TypeDecorator
from sqlalchemy.orm import relationship
from .database import Base

# PostgreSQL ARRAY 대체: SQLite 호환 JSON 직렬화 타입
class TextList(TypeDecorator):
    impl = Text
    cache_ok = True

    def process_bind_param(self, value, dialect):
        if value is None:
            return None
        if dialect.name == "postgresql":
            return value  # PostgreSQL은 네이티브 ARRAY 사용
        return json.dumps(value, ensure_ascii=False)

    def process_result_value(self, value, dialect):
        if value is None:
            return None
        if dialect.name == "postgresql":
            return value
        if isinstance(value, list):
            return value
        return json.loads(value)


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String(100), unique=True, nullable=False, index=True)
    hashed_password = Column(String(255), nullable=False)
    name = Column(String(100), default="")
    role = Column(String(20), default="viewer")         # 'admin' | 'viewer'
    allowed_teams = Column(TextList, nullable=True)     # NULL = 전체 열람
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

    snapshots = relationship("Snapshot", back_populates="uploader")


class Snapshot(Base):
    __tablename__ = "snapshots"

    id = Column(Integer, primary_key=True, index=True)
    week_label = Column(String(100), default="")
    base_date = Column(String(50), default="")
    uploaded_by = Column(Integer, ForeignKey("users.id"), nullable=True)
    uploaded_at = Column(DateTime, default=datetime.datetime.utcnow)
    is_active = Column(Boolean, default=True)

    uploader = relationship("User", back_populates="snapshots")
    records = relationship(
        "SalesRecord", back_populates="snapshot",
        cascade="all, delete-orphan"
    )


class SalesRecord(Base):
    __tablename__ = "sales_records"

    id = Column(Integer, primary_key=True, index=True)
    snapshot_id = Column(
        Integer, ForeignKey("snapshots.id", ondelete="CASCADE"), nullable=False, index=True
    )
    team = Column(String(50))
    channel = Column(String(100))
    brand = Column(String(10))
    code = Column(String(20))
    month = Column(Integer)           # 1~12

    y2024 = Column(Numeric(15, 3), default=0)
    y2025b = Column(Numeric(15, 3), default=0)
    y2025 = Column(Numeric(15, 3), default=0)
    plan = Column(Numeric(15, 3), default=0)
    actual = Column(Numeric(15, 3), default=0)

    fw1 = Column(Numeric(15, 3), nullable=True)
    fw2 = Column(Numeric(15, 3), nullable=True)
    fw3 = Column(Numeric(15, 3), nullable=True)
    fw4 = Column(Numeric(15, 3), nullable=True)
    fw5 = Column(Numeric(15, 3), nullable=True)

    snapshot = relationship("Snapshot", back_populates="records")


class Team(Base):
    __tablename__ = "teams"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), unique=True, nullable=False)
    display_order = Column(Integer, default=0)
    is_active = Column(Boolean, default=True)


class AppConfig(Base):
    __tablename__ = "app_config"

    key = Column(String(100), primary_key=True)
    value = Column(Text, nullable=True)
