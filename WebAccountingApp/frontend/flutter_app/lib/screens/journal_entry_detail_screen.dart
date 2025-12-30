import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../providers/auth_provider.dart';
import '../services/accounting_service.dart';
import '../models/accounting_models.dart';

class JournalEntryDetailScreen extends StatefulWidget {
  final JournalEntry? journalEntry;

  const JournalEntryDetailScreen({Key? key, this.journalEntry}) : super(key: key);

  @override
  State<JournalEntryDetailScreen> createState() => _JournalEntryDetailScreenState();
}

class _JournalEntryDetailScreenState extends State<JournalEntryDetailScreen> {
  final _formKey = GlobalKey<FormState>();
  final _entryDateController = TextEditingController();
  final _descriptionController = TextEditingController();
  final _referenceNumberController = TextEditingController();

  List<JournalEntryLine> _lines = [];
  bool _isLoading = false;

  @override
  void initState() {
    super.initState();
    if (widget.journalEntry != null) {
      _entryDateController.text = widget.journalEntry!.entryDate.toString().split(' ')[0];
      _descriptionController.text = widget.journalEntry!.description ?? '';
      _referenceNumberController.text = widget.journalEntry!.referenceNumber ?? '';
    } else {
      _entryDateController.text = DateTime.now().toString().split(' ')[0];
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text(widget.journalEntry != null ? 'Edit Journal Entry' : 'New Journal Entry'),
        centerTitle: true,
      ),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Form(
          key: _formKey,
          child: Column(
            children: [
              TextFormField(
                controller: _entryDateController,
                decoration: const InputDecoration(
                  labelText: 'Entry Date',
                  border: OutlineInputBorder(),
                ),
                readOnly: true,
                onTap: () async {
                  DateTime? pickedDate = await showDatePicker(
                    context: context,
                    initialDate: DateTime.now(),
                    firstDate: DateTime(2000),
                    lastDate: DateTime(2101),
                  );
                  if (pickedDate != null) {
                    _entryDateController.text = pickedDate.toString().split(' ')[0];
                  }
                },
              ),
              const SizedBox(height: 16),
              TextFormField(
                controller: _referenceNumberController,
                decoration: const InputDecoration(
                  labelText: 'Reference Number',
                  border: OutlineInputBorder(),
                ),
              ),
              const SizedBox(height: 16),
              TextFormField(
                controller: _descriptionController,
                decoration: const InputDecoration(
                  labelText: 'Description',
                  border: OutlineInputBorder(),
                ),
                maxLines: 3,
              ),
              const SizedBox(height: 16),
              const Text(
                'Journal Entry Lines',
                style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold),
              ),
              const SizedBox(height: 8),
              Expanded(
                child: _lines.isEmpty
                    ? const Center(
                        child: Text('No journal entry lines added yet'),
                      )
                    : ListView.builder(
                        itemCount: _lines.length,
                        itemBuilder: (context, index) {
                          final line = _lines[index];
                          return Card(
                            child: ListTile(
                              title: Text('Account: ${line.accountId}'),
                              subtitle: Text('Debit: ${line.debitAmount}, Credit: ${line.creditAmount}'),
                              trailing: IconButton(
                                icon: const Icon(Icons.delete),
                                onPressed: () {
                                  setState(() {
                                    _lines.removeAt(index);
                                  });
                                },
                              ),
                            ),
                          );
                        },
                      ),
              ),
              const SizedBox(height: 16),
              Row(
                children: [
                  Expanded(
                    child: ElevatedButton.icon(
                      onPressed: _addJournalLine,
                      icon: const Icon(Icons.add),
                      label: const Text('Add Line'),
                    ),
                  ),
                  const SizedBox(width: 16),
                  Expanded(
                    child: ElevatedButton.icon(
                      onPressed: _saveJournalEntry,
                      icon: const Icon(Icons.save),
                      label: const Text('Save'),
                    ),
                  ),
                ],
              ),
            ],
          ),
        ),
      ),
    );
  }

  void _addJournalLine() {
    // Add a new journal entry line
    showDialog(
      context: context,
      builder: (context) {
        final accountController = TextEditingController();
        final debitController = TextEditingController();
        final creditController = TextEditingController();
        final descriptionController = TextEditingController();

        return AlertDialog(
          title: const Text('Add Journal Entry Line'),
          content: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              TextField(
                controller: accountController,
                decoration: const InputDecoration(labelText: 'Account ID'),
              ),
              TextField(
                controller: debitController,
                decoration: const InputDecoration(labelText: 'Debit Amount'),
                keyboardType: TextInputType.number,
              ),
              TextField(
                controller: creditController,
                decoration: const InputDecoration(labelText: 'Credit Amount'),
                keyboardType: TextInputType.number,
              ),
              TextField(
                controller: descriptionController,
                decoration: const InputDecoration(labelText: 'Description'),
              ),
            ],
          ),
          actions: [
            TextButton(
              onPressed: () => Navigator.of(context).pop(),
              child: const Text('Cancel'),
            ),
            ElevatedButton(
              onPressed: () {
                final debit = double.tryParse(debitController.text) ?? 0.0;
                final credit = double.tryParse(creditController.text) ?? 0.0;
                
                if (debit > 0 || credit > 0) {
                  setState(() {
                    _lines.add(
                      JournalEntryLine(
                        id: '',
                        journalEntryId: '',
                        accountId: accountController.text,
                        partyId: null,
                        debitAmount: debit,
                        creditAmount: credit,
                        description: descriptionController.text,
                        createdAt: DateTime.now(),
                      ),
                    );
                  });
                  Navigator.of(context).pop();
                }
              },
              child: const Text('Add'),
            ),
          ],
        );
      },
    );
  }

  void _saveJournalEntry() async {
    if (_formKey.currentState!.validate()) {
      setState(() {
        _isLoading = true;
      });

      try {
        final authProvider = Provider.of<AuthProvider>(context, listen: false);
        
        // Create journal entry object
        final journalEntry = JournalEntry(
          id: widget.journalEntry?.id ?? '',
          entryNumber: widget.journalEntry?.entryNumber ?? 'TEMP',
          entryDate: DateTime.parse(_entryDateController.text),
          referenceNumber: _referenceNumberController.text.isEmpty ? null : _referenceNumberController.text,
          description: _descriptionController.text.isEmpty ? null : _descriptionController.text,
          transactionTypeId: null,
          createdById: null,
          postedAt: null,
          isPosted: false,
          createdAt: DateTime.now(),
          updatedAt: null,
        );

        // Save the journal entry
        final result = await AccountingService.createJournalEntry(
          authProvider.token!,
          journalEntry,
        );

        if (result != null) {
          ScaffoldMessenger.of(context).showSnackBar(
            const SnackBar(
              content: Text('Journal entry saved successfully'),
              backgroundColor: Colors.green,
            ),
          );
          Navigator.of(context).pop();
        } else {
          ScaffoldMessenger.of(context).showSnackBar(
            const SnackBar(
              content: Text('Failed to save journal entry'),
              backgroundColor: Colors.red,
            ),
          );
        }
      } catch (e) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
            content: Text('Error saving journal entry'),
            backgroundColor: Colors.red,
          ),
        );
      } finally {
        setState(() {
          _isLoading = false;
        });
      }
    }
  }
}

class JournalEntryLine {
  final String id;
  final String journalEntryId;
  final String accountId;
  final String? partyId;
  final double debitAmount;
  final double creditAmount;
  final String? description;
  final DateTime createdAt;

  JournalEntryLine({
    required this.id,
    required this.journalEntryId,
    required this.accountId,
    this.partyId,
    required this.debitAmount,
    required this.creditAmount,
    this.description,
    required this.createdAt,
  });
}