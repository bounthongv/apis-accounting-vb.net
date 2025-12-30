import 'package:flutter/foundation.dart';
import '../services/auth_service.dart';

class AuthProvider with ChangeNotifier {
  final AuthService _authService = AuthService();
  bool _isAuthenticated = false;
  String? _username;

  bool get isAuthenticated => _isAuthenticated;
  String? get username => _username;

  AuthProvider() {
    _checkAuthStatus();
  }

  Future<void> _checkAuthStatus() async {
    _isAuthenticated = await _authService.isAuthenticated();
    if (_isAuthenticated) {
      final userProfile = await _authService.getUserProfile();
      _username = userProfile?['username'];
    }
    notifyListeners();
  }

  Future<bool> login(String username, String password) async {
    bool success = await _authService.login(username, password);
    if (success) {
      _isAuthenticated = true;
      _username = username;
      notifyListeners();
    }
    return success;
  }

  Future<bool> register(String username, String email, String password) async {
    bool success = await _authService.register(username, email, password);
    return success;
  }

  Future<void> logout() async {
    await _authService.logout();
    _isAuthenticated = false;
    _username = null;
    notifyListeners();
  }

  Future<bool> checkAuthStatus() async {
    _isAuthenticated = await _authService.isAuthenticated();
    if (_isAuthenticated) {
      final userProfile = await _authService.getUserProfile();
      _username = userProfile?['username'];
    }
    notifyListeners();
    return _isAuthenticated;
  }
}