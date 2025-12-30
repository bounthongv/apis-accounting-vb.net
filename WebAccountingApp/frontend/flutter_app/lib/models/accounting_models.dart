class ChartOfAccount {
  final String id;
  final String accountCode;
  final String accountName;
  final String accountTypeId;
  final String? accountGroupId;
  final String? parentAccountId;
  final int level;
  final bool isActive;
  final bool isSystemAccount;
  final double openingBalance;
  final double currentBalance;
  final DateTime createdAt;
  final DateTime? updatedAt;

  ChartOfAccount({
    required this.id,
    required this.accountCode,
    required this.accountName,
    required this.accountTypeId,
    this.accountGroupId,
    this.parentAccountId,
    required this.level,
    required this.isActive,
    required this.isSystemAccount,
    required this.openingBalance,
    required this.currentBalance,
    required this.createdAt,
    this.updatedAt,
  });

  factory ChartOfAccount.fromJson(Map<String, dynamic> json) {
    return ChartOfAccount(
      id: json['id'],
      accountCode: json['account_code'],
      accountName: json['account_name'],
      accountTypeId: json['account_type_id'],
      accountGroupId: json['account_group_id'],
      parentAccountId: json['parent_account_id'],
      level: json['level'],
      isActive: json['is_active'],
      isSystemAccount: json['is_system_account'],
      openingBalance: double.parse(json['opening_balance'].toString()),
      currentBalance: double.parse(json['current_balance'].toString()),
      createdAt: DateTime.parse(json['created_at']),
      updatedAt: json['updated_at'] != null ? DateTime.parse(json['updated_at']) : null,
    );
  }
}

class JournalEntry {
  final String id;
  final String entryNumber;
  final DateTime entryDate;
  final String? referenceNumber;
  final String? description;
  final String? transactionTypeId;
  final String? createdById;
  final DateTime? postedAt;
  final bool isPosted;
  final DateTime createdAt;
  final DateTime? updatedAt;

  JournalEntry({
    required this.id,
    required this.entryNumber,
    required this.entryDate,
    this.referenceNumber,
    this.description,
    this.transactionTypeId,
    this.createdById,
    this.postedAt,
    required this.isPosted,
    required this.createdAt,
    this.updatedAt,
  });

  factory JournalEntry.fromJson(Map<String, dynamic> json) {
    return JournalEntry(
      id: json['id'],
      entryNumber: json['entry_number'],
      entryDate: DateTime.parse(json['entry_date'].toString()),
      referenceNumber: json['reference_number'],
      description: json['description'],
      transactionTypeId: json['transaction_type_id'],
      createdById: json['created_by'],
      postedAt: json['posted_at'] != null ? DateTime.parse(json['posted_at']) : null,
      isPosted: json['is_posted'],
      createdAt: DateTime.parse(json['created_at']),
      updatedAt: json['updated_at'] != null ? DateTime.parse(json['updated_at']) : null,
    );
  }
}

class Party {
  final String id;
  final String partyCode;
  final String partyName;
  final String? partyTypeId;
  final String? address;
  final String? phone;
  final String? email;
  final String? taxId;
  final bool isActive;
  final double creditLimit;
  final DateTime createdAt;
  final DateTime? updatedAt;

  Party({
    required this.id,
    required this.partyCode,
    required this.partyName,
    this.partyTypeId,
    this.address,
    this.phone,
    this.email,
    this.taxId,
    required this.isActive,
    required this.creditLimit,
    required this.createdAt,
    this.updatedAt,
  });

  factory Party.fromJson(Map<String, dynamic> json) {
    return Party(
      id: json['id'],
      partyCode: json['party_code'],
      partyName: json['party_name'],
      partyTypeId: json['party_type_id'],
      address: json['address'],
      phone: json['phone'],
      email: json['email'],
      taxId: json['tax_id'],
      isActive: json['is_active'],
      creditLimit: double.parse(json['credit_limit'].toString()),
      createdAt: DateTime.parse(json['created_at']),
      updatedAt: json['updated_at'] != null ? DateTime.parse(json['updated_at']) : null,
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'party_code': partyCode,
      'party_name': partyName,
      'party_type_id': partyTypeId,
      'address': address,
      'phone': phone,
      'email': email,
      'tax_id': taxId,
      'is_active': isActive,
      'credit_limit': creditLimit,
      'created_at': createdAt.toIso8601String(),
      'updated_at': updatedAt?.toIso8601String(),
    };
  }
}

class PartyDetail {
  final String id;
  final String partyId;
  final bool isCustomer;
  final bool isSupplier;
  final int paymentTerms;
  final DateTime createdAt;
  final DateTime? updatedAt;

  PartyDetail({
    required this.id,
    required this.partyId,
    required this.isCustomer,
    required this.isSupplier,
    required this.paymentTerms,
    required this.createdAt,
    this.updatedAt,
  });

  factory PartyDetail.fromJson(Map<String, dynamic> json) {
    return PartyDetail(
      id: json['id'],
      partyId: json['party_id'],
      isCustomer: json['is_customer'],
      isSupplier: json['is_supplier'],
      paymentTerms: json['payment_terms'],
      createdAt: DateTime.parse(json['created_at']),
      updatedAt: json['updated_at'] != null ? DateTime.parse(json['updated_at']) : null,
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'party_id': partyId,
      'is_customer': isCustomer,
      'is_supplier': isSupplier,
      'payment_terms': paymentTerms,
      'created_at': createdAt.toIso8601String(),
      'updated_at': updatedAt?.toIso8601String(),
    };
  }
}

class Asset {
  final String id;
  final String assetCode;
  final String assetName;
  final String? categoryId;
  final String? description;
  final DateTime? purchaseDate;
  final double purchasePrice;
  final double currentValue;
  final double accumulatedDepreciation;
  final double netBookValue;
  final DateTime? depreciationStartDate;
  final int? usefulLife;
  final double salvageValue;
  final String status;
  final String? location;
  final bool isActive;
  final DateTime createdAt;
  final DateTime? updatedAt;

  Asset({
    required this.id,
    required this.assetCode,
    required this.assetName,
    this.categoryId,
    this.description,
    this.purchaseDate,
    required this.purchasePrice,
    required this.currentValue,
    required this.accumulatedDepreciation,
    required this.netBookValue,
    this.depreciationStartDate,
    this.usefulLife,
    required this.salvageValue,
    required this.status,
    this.location,
    required this.isActive,
    required this.createdAt,
    this.updatedAt,
  });

  factory Asset.fromJson(Map<String, dynamic> json) {
    return Asset(
      id: json['id'],
      assetCode: json['asset_code'],
      assetName: json['asset_name'],
      categoryId: json['category_id'],
      description: json['description'],
      purchaseDate: json['purchase_date'] != null ? DateTime.parse(json['purchase_date']) : null,
      purchasePrice: double.parse(json['purchase_price'].toString()),
      currentValue: double.parse(json['current_value'].toString()),
      accumulatedDepreciation: double.parse(json['accumulated_depreciation'].toString()),
      netBookValue: double.parse(json['net_book_value'].toString()),
      depreciationStartDate: json['depreciation_start_date'] != null ? DateTime.parse(json['depreciation_start_date']) : null,
      usefulLife: json['useful_life'],
      salvageValue: double.parse(json['salvage_value'].toString()),
      status: json['status'],
      location: json['location'],
      isActive: json['is_active'],
      createdAt: DateTime.parse(json['created_at']),
      updatedAt: json['updated_at'] != null ? DateTime.parse(json['updated_at']) : null,
    );
  }
}