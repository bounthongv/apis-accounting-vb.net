import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../providers/auth_provider.dart';
import '../services/accounting_service.dart';
import '../models/accounting_models.dart';

class ChartOfAccountsScreen extends StatefulWidget {
  const ChartOfAccountsScreen({Key? key}) : super(key: key);

  @override
  State<ChartOfAccountsScreen> createState() => _ChartOfAccountsScreenState();
}

class _ChartOfAccountsScreenState extends State<ChartOfAccountsScreen> {
  List<ChartOfAccount> _accounts = [];
  bool _isLoading = true;

  @override
  void initState() {
    super.initState();
    _loadAccounts();
  }

  Future<void> _loadAccounts() async {
    final authProvider = Provider.of<AuthProvider>(context, listen: false);
    final accounts = await AccountingService.getChartOfAccounts(authProvider.token!);
    
    if (mounted) {
      setState(() {
        _accounts = accounts ?? [];
        _isLoading = false;
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Chart of Accounts'),
        centerTitle: true,
      ),
      body: _isLoading
          ? const Center(child: CircularProgressIndicator())
          : RefreshIndicator(
              onRefresh: _loadAccounts,
              child: ListView.builder(
                itemCount: _accounts.length,
                itemBuilder: (context, index) {
                  final account = _accounts[index];
                  return Card(
                    margin: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
                    child: ListTile(
                      title: Text(account.accountName),
                      subtitle: Text('${account.accountCode} - Balance: ${account.currentBalance.toStringAsFixed(2)}'),
                      trailing: Text(account.isActive ? 'Active' : 'Inactive'),
                    ),
                  );
                },
              ),
            ),
      floatingActionButton: FloatingActionButton(
        onPressed: () {
          // Add new account
        },
        child: const Icon(Icons.add),
      ),
    );
  }
}