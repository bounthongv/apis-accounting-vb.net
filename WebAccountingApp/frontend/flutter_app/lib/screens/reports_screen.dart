import 'package:flutter/material.dart';

class ReportsScreen extends StatefulWidget {
  const ReportsScreen({Key? key}) : super(key: key);

  @override
  State<ReportsScreen> createState() => _ReportsScreenState();
}

class _ReportsScreenState extends State<ReportsScreen> {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Financial Reports'),
        centerTitle: true,
      ),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: GridView.count(
          crossAxisCount: 2,
          crossAxisSpacing: 16,
          mainAxisSpacing: 16,
          children: [
            _buildReportCard(
              context,
              'Balance Sheet',
              Icons.balance,
              Colors.blue,
            ),
            _buildReportCard(
              context,
              'Income Statement',
              Icons.show_chart,
              Colors.green,
            ),
            _buildReportCard(
              context,
              'Cash Flow Statement',
              Icons.money,
              Colors.orange,
            ),
            _buildReportCard(
              context,
              'Trial Balance',
              Icons.table_chart,
              Colors.purple,
            ),
            _buildReportCard(
              context,
              'General Ledger',
              Icons.book,
              Colors.teal,
            ),
            _buildReportCard(
              context,
              'Journal Report',
              Icons.receipt,
              Colors.red,
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildReportCard(
    BuildContext context,
    String title,
    IconData icon,
    Color color,
  ) {
    return Card(
      elevation: 4,
      child: InkWell(
        onTap: () {
          // Navigate to specific report
          _showReportDialog(title);
        },
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            Icon(
              icon,
              size: 50,
              color: color,
            ),
            const SizedBox(height: 10),
            Text(
              title,
              style: const TextStyle(
                fontSize: 16,
                fontWeight: FontWeight.w500,
              ),
              textAlign: TextAlign.center,
            ),
          ],
        ),
      ),
    );
  }

  void _showReportDialog(String reportTitle) {
    showDialog(
      context: context,
      builder: (context) {
        return AlertDialog(
          title: Text(reportTitle),
          content: const Text('Report generation functionality would be implemented here.'),
          actions: [
            TextButton(
              onPressed: () => Navigator.of(context).pop(),
              child: const Text('Close'),
            ),
          ],
        );
      },
    );
  }
}