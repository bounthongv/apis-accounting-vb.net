import 'dart:convert';
import 'package:http/http.dart' as http;

class ApiService {
  static const String baseUrl = 'http://localhost:8000'; // Update with your backend URL

  static Future<Map<String, dynamic>?> login(
      String username, String password) async {
    try {
      final response = await http.post(
        Uri.parse('$baseUrl/auth/token'),
        headers: {'Content-Type': 'application/x-www-form-urlencoded'},
        body: {
          'username': username,
          'password': password,
        },
      );

      if (response.statusCode == 200) {
        return json.decode(response.body);
      } else {
        print('Login error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Login exception: $e');
      return null;
    }
  }

  static Future<Map<String, dynamic>?> register(String username, String email,
      String password, String firstName, String lastName) async {
    try {
      final response = await http.post(
        Uri.parse('$baseUrl/auth/register'),
        headers: {'Content-Type': 'application/json'},
        body: json.encode({
          'username': username,
          'email': email,
          'password': password,
          'first_name': firstName,
          'last_name': lastName,
          'is_active': true,
        }),
      );

      if (response.statusCode == 200) {
        return json.decode(response.body);
      } else {
        print('Registration error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Registration exception: $e');
      return null;
    }
  }

  static Future<Map<String, dynamic>?> getMe(String token) async {
    try {
      final response = await http.get(
        Uri.parse('$baseUrl/auth/me'),
        headers: {
          'Authorization': 'Bearer $token',
          'Content-Type': 'application/json',
        },
      );

      if (response.statusCode == 200) {
        return json.decode(response.body);
      } else {
        print('Get user error: ${response.statusCode} - ${response.body}');
        return null;
      }
    } catch (e) {
      print('Get user exception: $e');
      return null;
    }
  }
}