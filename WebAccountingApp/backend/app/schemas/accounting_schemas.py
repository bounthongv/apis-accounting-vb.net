from pydantic import BaseModel
from typing import Optional, List
from datetime import datetime, date
from uuid import UUID
from decimal import Decimal

# User Schemas
class UserBase(BaseModel):
    username: str
    email: str
    first_name: Optional[str] = None
    last_name: Optional[str] = None
    is_active: Optional[bool] = True

class UserCreate(UserBase):
    password: str

class UserUpdate(UserBase):
    password: Optional[str] = None

class User(UserBase):
    id: UUID
    created_at: datetime
    updated_at: Optional[datetime] = None

    class Config:
        from_attributes = True

# Role Schemas
class RoleBase(BaseModel):
    name: str
    description: Optional[str] = None

class RoleCreate(RoleBase):
    pass

class Role(RoleBase):
    id: UUID
    created_at: datetime

    class Config:
        from_attributes = True

# Permission Schemas
class PermissionBase(BaseModel):
    name: str
    description: Optional[str] = None

class PermissionCreate(PermissionBase):
    pass

class Permission(PermissionBase):
    id: UUID
    created_at: datetime

    class Config:
        from_attributes = True

# Account Type Schemas
class AccountTypeBase(BaseModel):
    name: str
    code: str
    description: Optional[str] = None
    normal_balance: str  # 'D' for Debit, 'C' for Credit

class AccountTypeCreate(AccountTypeBase):
    pass

class AccountType(AccountTypeBase):
    id: UUID
    created_at: datetime

    class Config:
        from_attributes = True

# Account Group Schemas
class AccountGroupBase(BaseModel):
    name: str
    code: str
    parent_id: Optional[UUID] = None
    level: Optional[int] = 0
    is_active: Optional[bool] = True

class AccountGroupCreate(AccountGroupBase):
    pass

class AccountGroupUpdate(AccountGroupBase):
    pass

class AccountGroup(AccountGroupBase):
    id: UUID
    created_at: datetime

    class Config:
        from_attributes = True

# Chart of Accounts Schemas
class ChartOfAccountsBase(BaseModel):
    account_code: str
    account_name: str
    account_type_id: UUID
    account_group_id: Optional[UUID] = None
    parent_account_id: Optional[UUID] = None
    level: Optional[int] = 0
    is_active: Optional[bool] = True
    is_system_account: Optional[bool] = False
    opening_balance: Optional[Decimal] = 0
    current_balance: Optional[Decimal] = 0

class ChartOfAccountsCreate(ChartOfAccountsBase):
    pass

class ChartOfAccountsUpdate(ChartOfAccountsBase):
    pass

class ChartOfAccounts(ChartOfAccountsBase):
    id: UUID
    created_at: datetime
    updated_at: Optional[datetime] = None

    class Config:
        from_attributes = True

# Party Type Schemas
class PartyTypeBase(BaseModel):
    name: str
    description: Optional[str] = None

class PartyTypeCreate(PartyTypeBase):
    pass

class PartyType(PartyTypeBase):
    id: UUID
    created_at: datetime

    class Config:
        from_attributes = True

# Party Schemas
class PartyBase(BaseModel):
    party_code: str
    party_name: str
    party_type_id: UUID
    address: Optional[str] = None
    phone: Optional[str] = None
    email: Optional[str] = None
    tax_id: Optional[str] = None
    is_active: Optional[bool] = True
    credit_limit: Optional[Decimal] = 0

class PartyCreate(PartyBase):
    pass

class PartyUpdate(PartyBase):
    pass

class Party(PartyBase):
    id: UUID
    created_at: datetime
    updated_at: Optional[datetime] = None

    class Config:
        from_attributes = True

# Party Detail Schemas
class PartyDetailBase(BaseModel):
    party_id: UUID
    is_customer: Optional[bool] = False
    is_supplier: Optional[bool] = False
    payment_terms: Optional[int] = 0

class PartyDetailCreate(PartyDetailBase):
    pass

class PartyDetailUpdate(PartyDetailBase):
    pass

class PartyDetail(PartyDetailBase):
    id: UUID
    created_at: datetime
    updated_at: Optional[datetime] = None

    class Config:
        from_attributes = True

# Transaction Type Schemas
class TransactionTypeBase(BaseModel):
    name: str
    code: str
    description: Optional[str] = None

class TransactionTypeCreate(TransactionTypeBase):
    pass

class TransactionType(TransactionTypeBase):
    id: UUID
    created_at: datetime

    class Config:
        from_attributes = True

# Journal Entry Schemas
class JournalEntryBase(BaseModel):
    entry_date: date
    reference_number: Optional[str] = None
    description: Optional[str] = None
    transaction_type_id: UUID
    created_by: UUID
    is_posted: Optional[bool] = False

class JournalEntryCreate(JournalEntryBase):
    pass

class JournalEntryUpdate(JournalEntryBase):
    pass

class JournalEntry(JournalEntryBase):
    id: UUID
    entry_number: str
    posted_at: Optional[datetime] = None
    created_at: datetime
    updated_at: Optional[datetime] = None

    class Config:
        from_attributes = True

# Journal Entry Line Schemas
class JournalEntryLineBase(BaseModel):
    journal_entry_id: UUID
    account_id: UUID
    party_id: Optional[UUID] = None
    debit_amount: Optional[Decimal] = 0
    credit_amount: Optional[Decimal] = 0
    description: Optional[str] = None

class JournalEntryLineCreate(JournalEntryLineBase):
    pass

class JournalEntryLineUpdate(JournalEntryLineBase):
    pass

class JournalEntryLine(JournalEntryLineBase):
    id: UUID
    created_at: datetime

    class Config:
        from_attributes = True

# Asset Category Schemas
class AssetCategoryBase(BaseModel):
    name: str
    code: str
    description: Optional[str] = None
    depreciation_method: Optional[str] = None
    useful_life: Optional[int] = None
    salvage_value: Optional[Decimal] = 0
    is_active: Optional[bool] = True

class AssetCategoryCreate(AssetCategoryBase):
    pass

class AssetCategoryUpdate(AssetCategoryBase):
    pass

class AssetCategory(AssetCategoryBase):
    id: UUID
    created_at: datetime

    class Config:
        from_attributes = True

# Asset Schemas
class AssetBase(BaseModel):
    asset_code: str
    asset_name: str
    category_id: UUID
    description: Optional[str] = None
    purchase_date: Optional[date] = None
    purchase_price: Decimal
    current_value: Decimal
    depreciation_start_date: Optional[date] = None
    useful_life: Optional[int] = None
    salvage_value: Optional[Decimal] = 0
    status: Optional[str] = 'Active'
    location: Optional[str] = None
    is_active: Optional[bool] = True

class AssetCreate(AssetBase):
    pass

class AssetUpdate(AssetBase):
    pass

class Asset(AssetBase):
    id: UUID
    accumulated_depreciation: Decimal
    net_book_value: Decimal
    created_at: datetime
    updated_at: Optional[datetime] = None

    class Config:
        from_attributes = True

# Asset Depreciation Schemas
class AssetDepreciationBase(BaseModel):
    asset_id: UUID
    depreciation_date: date
    depreciation_amount: Decimal
    accumulated_depreciation: Decimal
    net_book_value: Decimal

class AssetDepreciationCreate(AssetDepreciationBase):
    pass

class AssetDepreciationUpdate(AssetDepreciationBase):
    pass

class AssetDepreciation(AssetDepreciationBase):
    id: UUID
    created_at: datetime

    class Config:
        from_attributes = True

# Exchange Rate Schemas
class ExchangeRateBase(BaseModel):
    from_currency_id: UUID
    to_currency_id: UUID
    rate: Decimal
    effective_date: date
    is_active: Optional[bool] = True

class ExchangeRateCreate(ExchangeRateBase):
    pass

class ExchangeRateUpdate(ExchangeRateBase):
    pass

class ExchangeRate(ExchangeRateBase):
    id: UUID
    created_at: datetime

    class Config:
        from_attributes = True

# Financial Year Schemas
class FinancialYearBase(BaseModel):
    year_name: str
    start_date: date
    end_date: date
    is_closed: Optional[bool] = False
    is_active: Optional[bool] = True

class FinancialYearCreate(FinancialYearBase):
    pass

class FinancialYearUpdate(FinancialYearBase):
    pass

class FinancialYear(FinancialYearBase):
    id: UUID
    created_at: datetime

    class Config:
        from_attributes = True

# Accounting Period Schemas
class AccountingPeriodBase(BaseModel):
    period_name: str
    start_date: date
    end_date: date
    financial_year_id: UUID
    is_closed: Optional[bool] = False
    is_active: Optional[bool] = True

class AccountingPeriodCreate(AccountingPeriodBase):
    pass

class AccountingPeriodUpdate(AccountingPeriodBase):
    pass

class AccountingPeriod(AccountingPeriodBase):
    id: UUID
    created_at: datetime

    class Config:
        from_attributes = True