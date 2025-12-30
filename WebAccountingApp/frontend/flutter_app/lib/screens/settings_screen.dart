import 'package:flutter/material.dart';

class SettingsScreen extends StatelessWidget {
  const SettingsScreen({Key? key}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Settings'),
        centerTitle: true,
      ),
      body: ListView(
        children: [
          const Padding(
            padding: EdgeInsets.all(16.0),
            child: Text(
              'Application Settings',
              style: TextStyle(fontSize: 20, fontWeight: FontWeight.bold),
            ),
          ),
          _buildSettingsCard(
            context,
            'Financial Year',
            'Manage financial year settings',
            Icons.date_range,
            () {
              // Navigate to financial year settings
            },
          ),
          _buildSettingsCard(
            context,
            'Account Types',
            'Manage account types',
            Icons.account_balance,
            () {
              // Navigate to account types
            },
          ),
          _buildSettingsCard(
            context,
            'Currencies',
            'Manage currencies and exchange rates',
            Icons.attach_money,
            () {
              // Navigate to currencies
            },
          ),
          _buildSettingsCard(
            context,
            'Users & Permissions',
            'Manage users and permissions',
            Icons.people,
            () {
              // Navigate to users and permissions
            },
          ),
          _buildSettingsCard(
            context,
            'Backup & Restore',
            'Backup and restore data',
            Icons.backup,
            () {
              // Navigate to backup and restore
            },
          ),
          _buildSettingsCard(
            context,
            'System Information',
            'View system information',
            Icons.info,
            () {
              // Show system information
            },
          ),
        ],
      ),
    );
  }

  Widget _buildSettingsCard(
    BuildContext context,
    String title,
    String subtitle,
    IconData icon,
    VoidCallback onTap,
  ) {
    return Card(
      margin: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
      child: ListTile(
        leading: Icon(icon, color: Colors.blue),
        title: Text(title),
        subtitle: Text(subtitle),
        trailing: const Icon(Icons.chevron_right),
        onTap: onTap,
      ),
    );
  }
}