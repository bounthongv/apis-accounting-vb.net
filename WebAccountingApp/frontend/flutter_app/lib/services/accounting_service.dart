import 'dart:convert';
import 'package:http/http.dart' as http;
import '../models/accounting_models.dart';

class AccountingService {
  static const String baseUrl = 'http://localhost:8000'; // Update with your backend URL

  // Chart of Accounts methods
  static Future<List<ChartOfAccount>?> getChartOfAccounts(String token) async {
    try {
      final response = await http.get(
        Uri.parse('$baseUrl/chart_of_accounts'),
        headers: {
          'Authorization': 'Bearer $token',
          'Content-Type': 'application/json',
        },
      );

      if (response.statusCode == 200) {
        final List<dynamic> data = json.decode(response.body);
        return data.map((json) => ChartOfAccount.fromJson(json)).toList();
      } else {
        print('Get chart of accounts error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Get chart of accounts exception: $e');
      return null;
    }
  }

  static Future<ChartOfAccount?> createChartOfAccount(
      String token, ChartOfAccount chartOfAccount) async {
    try {
      final response = await http.post(
        Uri.parse('$baseUrl/chart_of_accounts'),
        headers: {
          'Authorization': 'Bearer $token',
          'Content-Type': 'application/json',
        },
        body: json.encode(chartOfAccount.toJson()),
      );

      if (response.statusCode == 200) {
        return ChartOfAccount.fromJson(json.decode(response.body));
      } else {
        print('Create chart of account error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Create chart of account exception: $e');
      return null;
    }
  }

  // Journal Entries methods
  static Future<List<JournalEntry>?> getJournalEntries(String token) async {
    try {
      final response = await http.get(
        Uri.parse('$baseUrl/journal_entries'),
        headers: {
          'Authorization': 'Bearer $token',
          'Content-Type': 'application/json',
        },
      );

      if (response.statusCode == 200) {
        final List<dynamic> data = json.decode(response.body);
        return data.map((json) => JournalEntry.fromJson(json)).toList();
      } else {
        print('Get journal entries error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Get journal entries exception: $e');
      return null;
    }
  }

  static Future<JournalEntry?> createJournalEntry(
      String token, JournalEntry journalEntry) async {
    try {
      final response = await http.post(
        Uri.parse('$baseUrl/journal_entries'),
        headers: {
          'Authorization': 'Bearer $token',
          'Content-Type': 'application/json',
        },
        body: json.encode(journalEntry.toJson()),
      );

      if (response.statusCode == 200) {
        return JournalEntry.fromJson(json.decode(response.body));
      } else {
        print('Create journal entry error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Create journal entry exception: $e');
      return null;
    }
  }

  // Parties methods
  static Future<List<Party>?> getParties(String token) async {
    try {
      final response = await http.get(
        Uri.parse('$baseUrl/parties'),
        headers: {
          'Authorization': 'Bearer $token',
          'Content-Type': 'application/json',
        },
      );

      if (response.statusCode == 200) {
        final List<dynamic> data = json.decode(response.body);
        return data.map((json) => Party.fromJson(json)).toList();
      } else {
        print('Get parties error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Get parties exception: $e');
      return null;
    }
  }

  static Future<Party?> createParty(String token, Party party) async {
    try {
      final response = await http.post(
        Uri.parse('$baseUrl/parties'),
        headers: {
          'Authorization': 'Bearer $token',
          'Content-Type': 'application/json',
        },
        body: json.encode(party.toJson()),
      );

      if (response.statusCode == 200) {
        return Party.fromJson(json.decode(response.body));
      } else {
        print('Create party error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Create party exception: $e');
      return null;
    }
  }

  // Assets methods
  static Future<List<Asset>?> getAssets(String token) async {
    try {
      final response = await http.get(
        Uri.parse('$baseUrl/assets'),
        headers: {
          'Authorization': 'Bearer $token',
          'Content-Type': 'application/json',
        },
      );

      if (response.statusCode == 200) {
        final List<dynamic> data = json.decode(response.body);
        return data.map((json) => Asset.fromJson(json)).toList();
      } else {
        print('Get assets error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Get assets exception: $e');
      return null;
    }
  }

  static Future<Asset?> createAsset(String token, Asset asset) async {
    try {
      final response = await http.post(
        Uri.parse('$baseUrl/assets'),
        headers: {
          'Authorization': 'Bearer $token',
          'Content-Type': 'application/json',
        },
        body: json.encode(asset.toJson()),
      );

      if (response.statusCode == 200) {
        return Asset.fromJson(json.decode(response.body));
      } else {
        print('Create asset error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Create asset exception: $e');
      return null;
    }
  }

  // Account Types methods
  static Future<List<AccountType>?> getAccountTypes(String token) async {
    try {
      final response = await http.get(
        Uri.parse('$baseUrl/account_types'),
        headers: {
          'Authorization': 'Bearer $token',
          'Content-Type': 'application/json',
        },
      );

      if (response.statusCode == 200) {
        final List<dynamic> data = json.decode(response.body);
        return data.map((json) => AccountType.fromJson(json)).toList();
      } else {
        print('Get account types error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Get account types exception: $e');
      return null;
    }
  }

  // Account Groups methods
  static Future<List<AccountGroup>?> getAccountGroups(String token) async {
    try {
      final response = await http.get(
        Uri.parse('$baseUrl/account_groups'),
        headers: {
          'Authorization': 'Bearer $token',
          'Content-Type': 'application/json',
        },
      );

      if (response.statusCode == 200) {
        final List<dynamic> data = json.decode(response.body);
        return data.map((json) => AccountGroup.fromJson(json)).toList();
      } else {
        print('Get account groups error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Get account groups exception: $e');
      return null;
    }
  }

  // Transaction Types methods
  static Future<List<TransactionType>?> getTransactionTypes(String token) async {
    try {
      final response = await http.get(
        Uri.parse('$baseUrl/transaction_types'),
        headers: {
          'Authorization': 'Bearer $token',
          'Content-Type': 'application/json',
        },
      );

      if (response.statusCode == 200) {
        final List<dynamic> data = json.decode(response.body);
        return data.map((json) => TransactionType.fromJson(json)).toList();
      } else {
        print('Get transaction types error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Get transaction types exception: $e');
      return null;
    }
  }

  // Financial Years methods
  static Future<List<FinancialYear>?> getFinancialYears(String token) async {
    try {
      final response = await http.get(
        Uri.parse('$baseUrl/financial_years'),
        headers: {
          'Authorization': 'Bearer $token',
          'Content-Type': 'application/json',
        },
      );

      if (response.statusCode == 200) {
        final List<dynamic> data = json.decode(response.body);
        return data.map((json) => FinancialYear.fromJson(json)).toList();
      } else {
        print('Get financial years error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Get financial years exception: $e');
      return null;
    }
  }

  // Currencies methods
  static Future<List<Currency>?> getCurrencies(String token) async {
    try {
      final response = await http.get(
        Uri.parse('$baseUrl/currencies'),
        headers: {
          'Authorization': 'Bearer $token',
          'Content-Type': 'application/json',
        },
      );

      if (response.statusCode == 200) {
        final List<dynamic> data = json.decode(response.body);
        return data.map((json) => Currency.fromJson(json)).toList();
      } else {
        print('Get currencies error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Get currencies exception: $e');
      return null;
    }
  }
}

class AccountType {
  final String id;
  final String name;
  final String code;
  final String? description;
  final String normalBalance;
  final DateTime createdAt;

  AccountType({
    required this.id,
    required this.name,
    required this.code,
    this.description,
    required this.normalBalance,
    required this.createdAt,
  });

  factory AccountType.fromJson(Map<String, dynamic> json) {
    return AccountType(
      id: json['id'],
      name: json['name'],
      code: json['code'],
      description: json['description'],
      normalBalance: json['normal_balance'],
      createdAt: DateTime.parse(json['created_at']),
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'name': name,
      'code': code,
      'description': description,
      'normal_balance': normalBalance,
      'created_at': createdAt.toIso8601String(),
    };
  }
}

class AccountGroup {
  final String id;
  final String name;
  final String code;
  final String? parentId;
  final int level;
  final bool isActive;
  final DateTime createdAt;

  AccountGroup({
    required this.id,
    required this.name,
    required this.code,
    this.parentId,
    required this.level,
    required this.isActive,
    required this.createdAt,
  });

  factory AccountGroup.fromJson(Map<String, dynamic> json) {
    return AccountGroup(
      id: json['id'],
      name: json['name'],
      code: json['code'],
      parentId: json['parent_id'],
      level: json['level'],
      isActive: json['is_active'],
      createdAt: DateTime.parse(json['created_at']),
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'name': name,
      'code': code,
      'parent_id': parentId,
      'level': level,
      'is_active': isActive,
      'created_at': createdAt.toIso8601String(),
    };
  }
}

class TransactionType {
  final String id;
  final String name;
  final String code;
  final String? description;
  final DateTime createdAt;

  TransactionType({
    required this.id,
    required this.name,
    required this.code,
    this.description,
    required this.createdAt,
  });

  factory TransactionType.fromJson(Map<String, dynamic> json) {
    return TransactionType(
      id: json['id'],
      name: json['name'],
      code: json['code'],
      description: json['description'],
      createdAt: DateTime.parse(json['created_at']),
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'name': name,
      'code': code,
      'description': description,
      'created_at': createdAt.toIso8601String(),
    };
  }
}

class FinancialYear {
  final String id;
  final String yearName;
  final DateTime startDate;
  final DateTime endDate;
  final bool isClosed;
  final bool isActive;
  final DateTime createdAt;

  FinancialYear({
    required this.id,
    required this.yearName,
    required this.startDate,
    required this.endDate,
    required this.isClosed,
    required this.isActive,
    required this.createdAt,
  });

  factory FinancialYear.fromJson(Map<String, dynamic> json) {
    return FinancialYear(
      id: json['id'],
      yearName: json['year_name'],
      startDate: DateTime.parse(json['start_date'].toString()),
      endDate: DateTime.parse(json['end_date'].toString()),
      isClosed: json['is_closed'],
      isActive: json['is_active'],
      createdAt: DateTime.parse(json['created_at']),
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'year_name': yearName,
      'start_date': startDate.toIso8601String().split('T')[0],
      'end_date': endDate.toIso8601String().split('T')[0],
      'is_closed': isClosed,
      'is_active': isActive,
      'created_at': createdAt.toIso8601String(),
    };
  }
}

class Currency {
  final String id;
  final String code;
  final String name;
  final String? symbol;
  final bool isActive;
  final DateTime createdAt;

  Currency({
    required this.id,
    required this.code,
    required this.name,
    this.symbol,
    required this.isActive,
    required this.createdAt,
  });

  factory Currency.fromJson(Map<String, dynamic> json) {
    return Currency(
      id: json['id'],
      code: json['code'],
      name: json['name'],
      symbol: json['symbol'],
      isActive: json['is_active'],
      createdAt: DateTime.parse(json['created_at']),
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'code': code,
      'name': name,
      'symbol': symbol,
      'is_active': isActive,
      'created_at': createdAt.toIso8601String(),
    };
  }
}