import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../providers/auth_provider.dart';
import '../services/accounting_service.dart';
import '../models/accounting_models.dart';
import 'journal_entry_detail_screen.dart';

class JournalEntriesScreen extends StatefulWidget {
  const JournalEntriesScreen({Key? key}) : super(key: key);

  @override
  State<JournalEntriesScreen> createState() => _JournalEntriesScreenState();
}

class _JournalEntriesScreenState extends State<JournalEntriesScreen> {
  List<JournalEntry> _entries = [];
  bool _isLoading = true;

  @override
  void initState() {
    super.initState();
    _loadEntries();
  }

  Future<void> _loadEntries() async {
    final authProvider = Provider.of<AuthProvider>(context, listen: false);
    final entries = await AccountingService.getJournalEntries(authProvider.token!);

    if (mounted) {
      setState(() {
        _entries = entries ?? [];
        _isLoading = false;
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Journal Entries'),
        centerTitle: true,
      ),
      body: _isLoading
          ? const Center(child: CircularProgressIndicator())
          : RefreshIndicator(
              onRefresh: _loadEntries,
              child: ListView.builder(
                itemCount: _entries.length,
                itemBuilder: (context, index) {
                  final entry = _entries[index];
                  return Card(
                    margin: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
                    child: ListTile(
                      title: Text(entry.entryNumber),
                      subtitle: Text('${entry.entryDate} - ${entry.description ?? 'No description'}'),
                      trailing: Text(entry.isPosted ? 'Posted' : 'Draft'),
                      onTap: () {
                        // Navigate to journal entry detail
                        Navigator.push(
                          context,
                          MaterialPageRoute(
                            builder: (context) => JournalEntryDetailScreen(journalEntry: entry),
                          ),
                        );
                      },
                    ),
                  );
                },
              ),
            ),
      floatingActionButton: FloatingActionButton(
        onPressed: () {
          // Add new journal entry
          Navigator.push(
            context,
            MaterialPageRoute(
              builder: (context) => const JournalEntryDetailScreen(),
            ),
          );
        },
        child: const Icon(Icons.add),
      ),
    );
  }
}