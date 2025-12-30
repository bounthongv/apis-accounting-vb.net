import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../providers/auth_provider.dart';
import '../services/accounting_service.dart';
import '../models/accounting_models.dart';

class AssetsScreen extends StatefulWidget {
  const AssetsScreen({Key? key}) : super(key: key);

  @override
  State<AssetsScreen> createState() => _AssetsScreenState();
}

class _AssetsScreenState extends State<AssetsScreen> {
  List<Asset> _assets = [];
  bool _isLoading = true;

  @override
  void initState() {
    super.initState();
    _loadAssets();
  }

  Future<void> _loadAssets() async {
    final authProvider = Provider.of<AuthProvider>(context, listen: false);
    final assets = await AccountingService.getAssets(authProvider.token!);
    
    if (mounted) {
      setState(() {
        _assets = assets ?? [];
        _isLoading = false;
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Assets'),
        centerTitle: true,
      ),
      body: _isLoading
          ? const Center(child: CircularProgressIndicator())
          : RefreshIndicator(
              onRefresh: _loadAssets,
              child: ListView.builder(
                itemCount: _assets.length,
                itemBuilder: (context, index) {
                  final asset = _assets[index];
                  return Card(
                    margin: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
                    child: ListTile(
                      title: Text(asset.assetName),
                      subtitle: Text('${asset.assetCode} - Status: ${asset.status}'),
                      trailing: Text('Value: \$${asset.currentValue.toStringAsFixed(2)}'),
                    ),
              ),
            ),
          ),
      floatingActionButton: FloatingActionButton(
        onPressed: () {
          // Add new asset
          _showAddAssetDialog();
        },
        child: const Icon(Icons.add),
      ),
    );
  }

  Future<void> _showAddAssetDialog() async {
    final assetCodeController = TextEditingController();
    final assetNameController = TextEditingController();
    final purchasePriceController = TextEditingController();
    final currentValueController = TextEditingController();

    await showDialog(
      context: context,
      builder: (context) {
        return AlertDialog(
          title: const Text('Add New Asset'),
          content: SizedBox(
            width: double.maxFinite,
            child: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                TextField(
                  controller: assetCodeController,
                  decoration: const InputDecoration(labelText: 'Asset Code'),
                ),
                TextField(
                  controller: assetNameController,
                  decoration: const InputDecoration(labelText: 'Asset Name'),
                ),
                TextField(
                  controller: purchasePriceController,
                  decoration: const InputDecoration(labelText: 'Purchase Price'),
                  keyboardType: TextInputType.number,
                ),
                TextField(
                  controller: currentValueController,
                  decoration: const InputDecoration(labelText: 'Current Value'),
                  keyboardType: TextInputType.number,
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
                // Add asset logic would go here
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