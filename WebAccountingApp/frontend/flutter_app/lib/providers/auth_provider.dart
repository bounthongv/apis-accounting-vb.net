import 'package:flutter/foundation.dart';
import 'package:shared_preferences/shared_preferences.dart';
import '../utils/api_service.dart';

class AuthProvider with ChangeNotifier {
  String? _token;
  bool _isLoading = false;

  String? get token => _token;
  bool get isAuthenticated => _token != null;
  bool get isLoading => _isLoading;

  AuthProvider() {
    _loadToken();
  }

  Future<void> _loadToken() async {
    SharedPreferences prefs = await SharedPreferences.getInstance();
    _token = prefs.getString('token');
    notifyListeners();
  }

  Future<bool> login(String username, String password) async {
    _isLoading = true;
    notifyListeners();

    try {
      final response = await ApiService.login(username, password);
      if (response != null) {
        _token = response['access_token'];
        SharedPreferences prefs = await SharedPreferences.getInstance();
        await prefs.setString('token', _token!);
        notifyListeners();
        return true;
      }
    } catch (e) {
      debugPrint('Login error: $e');
    } finally {
      _isLoading = false;
      notifyListeners();
    }
    return false;
  }

  Future<void> logout() async {
    _token = null;
    SharedPreferences prefs = await SharedPreferences.getInstance();
    await prefs.remove('token');
    notifyListeners();
  }

  Future<bool> register(String username, String email, String password,
      String firstName, String lastName) async {
    _isLoading = true;
    notifyListeners();

    try {
      final response = await ApiService.register(
          username, email, password, firstName, lastName);
      if (response != null) {
        // Auto-login after registration
        return await login(username, password);
      }
    } catch (e) {
      debugPrint('Registration error: $e');
    } finally {
      _isLoading = false;
      notifyListeners();
    }
    return false;
  }
}