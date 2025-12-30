import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../providers/auth_provider.dart';
import '../services/accounting_service.dart';
import '../models/accounting_models.dart';

class PartiesScreen extends StatefulWidget {
  const PartiesScreen({Key? key}) : super(key: key);

  @override
  State<PartiesScreen> createState() => _PartiesScreenState();
}

class _PartiesScreenState extends State<PartiesScreen> {
  List<Party> _parties = [];
  bool _isLoading = true;

  @override
  void initState() {
    super.initState();
    _loadParties();
  }

  Future<void> _loadParties() async {
    final authProvider = Provider.of<AuthProvider>(context, listen: false);
    final parties = await AccountingService.getParties(authProvider.token!);
    
    if (mounted) {
      setState(() {
        _parties = parties ?? [];
        _isLoading = false;
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Parties'),
        centerTitle: true,
      ),
      body: _isLoading
          ? const Center(child: CircularProgressIndicator())
          : RefreshIndicator(
              onRefresh: _loadParties,
              child: ListView.builder(
                itemCount: _parties.length,
                itemBuilder: (context, index) {
                  final party = _parties[index];
                  return Card(
                    margin: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
                    child: ListTile(
                      title: Text(party.partyName),
                      subtitle: Text('${party.partyCode} - ${party.email ?? 'No email'}'),
                      trailing: Text(party.isActive ? 'Active' : 'Inactive'),
                    ),
              ),
            ),
          ),
      floatingActionButton: FloatingActionButton(
        onPressed: () {
          // Add new party
          _showAddPartyDialog();
        },
        child: const Icon(Icons.add),
      ),
    );
  }

  Future<void> _showAddPartyDialog() async {
    final partyCodeController = TextEditingController();
    final partyNameController = TextEditingController();
    final emailController = TextEditingController();
    final phoneController = TextEditingController();

    await showDialog(
      context: context,
      builder: (context) {
        return AlertDialog(
          title: const Text('Add New Party'),
          content: SizedBox(
            width: double.maxFinite,
            child: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                TextField(
                  controller: partyCodeController,
                  decoration: const InputDecoration(labelText: 'Party Code'),
                ),
                TextField(
                  controller: partyNameController,
                  decoration: const InputDecoration(labelText: 'Party Name'),
                ),
                TextField(
                  controller: emailController,
                  decoration: const InputDecoration(labelText: 'Email'),
                ),
                TextField(
                  controller: phoneController,
                  decoration: const InputDecoration(labelText: 'Phone'),
                ),
              ],
            ),
          ),
          actions: [
            TextButton(
              onPressed: () => Navigator.of(context).pop(),
              child: const Text('Cancel'),
            ),
            ElevatedButton(
              onPressed: () {
                // Add party logic would go here
                Navigator.of(context).pop();
              },
              child: const Text('Add'),
            ),
          ],
        );
      },
    );
  }
}