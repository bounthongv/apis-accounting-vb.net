from sqlalchemy import Column, Integer, String, DateTime, Boolean, DECIMAL, Date, ForeignKey, Text
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import relationship
from sqlalchemy.sql import func
import uuid
from ..database import Base

class User(Base):
    __tablename__ = "users"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    username = Column(String(50), unique=True, nullable=False)
    email = Column(String(100), unique=True, nullable=False)
    hashed_password = Column(String, nullable=False)
    first_name = Column(String(50))
    last_name = Column(String(50))
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())

    # Relationships
    user_roles = relationship("UserRole", back_populates="user")
    journal_entries = relationship("JournalEntry", back_populates="created_by_user")


class Role(Base):
    __tablename__ = "roles"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    name = Column(String(50), unique=True, nullable=False)
    description = Column(Text)
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    # Relationships
    user_roles = relationship("UserRole", back_populates="role")
    role_permissions = relationship("RolePermission", back_populates="role")


class UserRole(Base):
    __tablename__ = "user_roles"

    user_id = Column(UUID(as_uuid=True), ForeignKey("users.id"), primary_key=True)
    role_id = Column(UUID(as_uuid=True), ForeignKey("roles.id"), primary_key=True)
    
    # Relationships
    user = relationship("User", back_populates="user_roles")
    role = relationship("Role", back_populates="user_roles")


class Permission(Base):
    __tablename__ = "permissions"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    name = Column(String(50), unique=True, nullable=False)
    description = Column(Text)
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    # Relationships
    role_permissions = relationship("RolePermission", back_populates="permission")


class RolePermission(Base):
    __tablename__ = "role_permissions"

    role_id = Column(UUID(as_uuid=True), ForeignKey("roles.id"), primary_key=True)
    permission_id = Column(UUID(as_uuid=True), ForeignKey("permissions.id"), primary_key=True)
    
    # Relationships
    role = relationship("Role", back_populates="role_permissions")
    permission = relationship("Permission", back_populates="role_permissions")


class AccountType(Base):
    __tablename__ = "account_types"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    name = Column(String(50), nullable=False)
    code = Column(String(10), unique=True, nullable=False)  # ASSET, LIABILITY, EQUITY, REVENUE, EXPENSE
    description = Column(Text)
    normal_balance = Column(String(1))  # 'D' for Debit, 'C' for Credit
    created_at = Column(DateTime(timezone=True), server_default=func.now())


class AccountGroup(Base):
    __tablename__ = "account_groups"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    name = Column(String(100), nullable=False)
    code = Column(String(20), unique=True, nullable=False)
    parent_id = Column(UUID(as_uuid=True), ForeignKey("account_groups.id"))
    level = Column(Integer, default=0)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    # Self-referencing relationship
    parent = relationship("AccountGroup", remote_side=[id])
    children = relationship("AccountGroup")


class ChartOfAccounts(Base):
    __tablename__ = "chart_of_accounts"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    account_code = Column(String(20), unique=True, nullable=False)
    account_name = Column(String(200), nullable=False)
    account_type_id = Column(UUID(as_uuid=True), ForeignKey("account_types.id"))
    account_group_id = Column(UUID(as_uuid=True), ForeignKey("account_groups.id"))
    parent_account_id = Column(UUID(as_uuid=True), ForeignKey("chart_of_accounts.id"))
    level = Column(Integer, default=0)
    is_active = Column(Boolean, default=True)
    is_system_account = Column(Boolean, default=False)
    opening_balance = Column(DECIMAL(18, 2), default=0)
    current_balance = Column(DECIMAL(18, 2), default=0)
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())

    # Relationships
    account_type = relationship("AccountType")
    account_group = relationship("AccountGroup")
    parent_account = relationship("ChartOfAccounts", remote_side=[id])
    children_accounts = relationship("ChartOfAccounts")
    journal_entry_lines = relationship("JournalEntryLine", back_populates="account")


class PartyType(Base):
    __tablename__ = "party_types"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    name = Column(String(50), nullable=False)
    description = Column(Text)
    created_at = Column(DateTime(timezone=True), server_default=func.now())


class Party(Base):
    __tablename__ = "parties"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    party_code = Column(String(20), unique=True, nullable=False)
    party_name = Column(String(200), nullable=False)
    party_type_id = Column(UUID(as_uuid=True), ForeignKey("party_types.id"))
    address = Column(Text)
    phone = Column(String(20))
    email = Column(String(100))
    tax_id = Column(String(50))
    is_active = Column(Boolean, default=True)
    credit_limit = Column(DECIMAL(18, 2), default=0)
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())

    # Relationships
    party_type = relationship("PartyType")
    journal_entry_lines = relationship("JournalEntryLine", back_populates="party")
    party_details = relationship("PartyDetail", back_populates="party")


class PartyDetail(Base):
    __tablename__ = "party_details"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    party_id = Column(UUID(as_uuid=True), ForeignKey("parties.id"), nullable=False)
    is_customer = Column(Boolean, default=False)
    is_supplier = Column(Boolean, default=False)
    payment_terms = Column(Integer, default=0)  # Days
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())

    # Relationships
    party = relationship("Party", back_populates="party_details")


class TransactionType(Base):
    __tablename__ = "transaction_types"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    name = Column(String(50), nullable=False)
    code = Column(String(10), unique=True, nullable=False)
    description = Column(Text)
    created_at = Column(DateTime(timezone=True), server_default=func.now())


class JournalEntry(Base):
    __tablename__ = "journal_entries"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    entry_number = Column(String(50), unique=True, nullable=False)
    entry_date = Column(Date, nullable=False)
    reference_number = Column(String(50))
    description = Column(Text)
    transaction_type_id = Column(UUID(as_uuid=True), ForeignKey("transaction_types.id"))
    created_by = Column(UUID(as_uuid=True), ForeignKey("users.id"))
    posted_at = Column(DateTime(timezone=True))
    is_posted = Column(Boolean, default=False)
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())

    # Relationships
    transaction_type = relationship("TransactionType")
    created_by_user = relationship("User", back_populates="journal_entries")
    journal_entry_lines = relationship("JournalEntryLine", back_populates="journal_entry")


class JournalEntryLine(Base):
    __tablename__ = "journal_entry_lines"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    journal_entry_id = Column(UUID(as_uuid=True), ForeignKey("journal_entries.id"), nullable=False)
    account_id = Column(UUID(as_uuid=True), ForeignKey("chart_of_accounts.id"), nullable=False)
    party_id = Column(UUID(as_uuid=True), ForeignKey("parties.id"))
    debit_amount = Column(DECIMAL(18, 2), default=0)
    credit_amount = Column(DECIMAL(18, 2), default=0)
    description = Column(Text)
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    # Relationships
    journal_entry = relationship("JournalEntry", back_populates="journal_entry_lines")
    account = relationship("ChartOfAccounts", back_populates="journal_entry_lines")
    party = relationship("Party", back_populates="journal_entry_lines")


class AssetCategory(Base):
    __tablename__ = "asset_categories"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    name = Column(String(100), nullable=False)
    code = Column(String(20), unique=True, nullable=False)
    description = Column(Text)
    depreciation_method = Column(String(20))  # StraightLine, DecliningBalance, etc.
    useful_life = Column(Integer)  # in months
    salvage_value = Column(DECIMAL(18, 2), default=0)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now())


class Asset(Base):
    __tablename__ = "assets"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    asset_code = Column(String(20), unique=True, nullable=False)
    asset_name = Column(String(200), nullable=False)
    category_id = Column(UUID(as_uuid=True), ForeignKey("asset_categories.id"))
    description = Column(Text)
    purchase_date = Column(Date)
    purchase_price = Column(DECIMAL(18, 2), nullable=False)
    current_value = Column(DECIMAL(18, 2), nullable=False)
    accumulated_depreciation = Column(DECIMAL(18, 2), default=0)
    net_book_value = Column(DECIMAL(18, 2), nullable=False)
    depreciation_start_date = Column(Date)
    useful_life = Column(Integer)  # in months
    salvage_value = Column(DECIMAL(18, 2), default=0)
    status = Column(String(20), default='Active')  # Active, Disposed, UnderMaintenance
    location = Column(Text)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())

    # Relationships
    category = relationship("AssetCategory")
    asset_depreciations = relationship("AssetDepreciation", back_populates="asset")


class AssetDepreciation(Base):
    __tablename__ = "asset_depreciations"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    asset_id = Column(UUID(as_uuid=True), ForeignKey("assets.id"), nullable=False)
    depreciation_date = Column(Date, nullable=False)
    depreciation_amount = Column(DECIMAL(18, 2), nullable=False)
    accumulated_depreciation = Column(DECIMAL(18, 2), nullable=False)
    net_book_value = Column(DECIMAL(18, 2), nullable=False)
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    # Relationships
    asset = relationship("Asset", back_populates="asset_depreciations")


class Currency(Base):
    __tablename__ = "currencies"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    code = Column(String(3), unique=True, nullable=False)  # ISO currency code
    name = Column(String(50), nullable=False)
    symbol = Column(String(5))
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now())


class ExchangeRate(Base):
    __tablename__ = "exchange_rates"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    from_currency_id = Column(UUID(as_uuid=True), ForeignKey("currencies.id"))
    to_currency_id = Column(UUID(as_uuid=True), ForeignKey("currencies.id"))
    rate = Column(DECIMAL(18, 6), nullable=False)
    effective_date = Column(Date, nullable=False)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    # Relationships
    from_currency = relationship("Currency", foreign_keys=[from_currency_id])
    to_currency = relationship("Currency", foreign_keys=[to_currency_id])


class FinancialYear(Base):
    __tablename__ = "financial_years"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    year_name = Column(String(20), nullable=False)
    start_date = Column(Date, nullable=False)
    end_date = Column(Date, nullable=False)
    is_closed = Column(Boolean, default=False)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now())


class AccountingPeriod(Base):
    __tablename__ = "accounting_periods"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    period_name = Column(String(20), nullable=False)
    start_date = Column(Date, nullable=False)
    end_date = Column(Date, nullable=False)
    financial_year_id = Column(UUID(as_uuid=True), ForeignKey("financial_years.id"))
    is_closed = Column(Boolean, default=False)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    # Relationships
    financial_year = relationship("FinancialYear")