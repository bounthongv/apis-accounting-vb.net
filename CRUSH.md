# CRUSH.md

## General Guidelines
*   **Knowledge Management**:
    *   Use `byterover-store-knowledge` when learning new patterns, APIs, error solutions, or completing significant tasks.
    *   Use `byterover-retrieve-knowledge` before starting tasks, making architectural decisions, debugging, or working with unfamiliar code.

## Backend (Python)
*   **Dependencies**: `pip install -r WebAccountingApp/backend/requirements.txt`
*   **Run Application**: `uvicorn WebAccountingApp/backend/app.main:app --reload`
*   **Testing**:
    *   Run all tests: `pytest WebAccountingApp/backend` (assuming test files will be created in the `backend` directory)
    *   Run a single test file: `pytest WebAccountingApp/backend/path/to/your_test_file.py`
*   **Linting/Formatting**:
    *   Lint: `flake8 WebAccountingApp/backend`
    *   Format: `black WebAccountingApp/backend`
*   **Code Style**: Adhere to PEP 8. Use explicit type hints. Use clear and descriptive variable and function names.

## Frontend (Flutter/Dart)
*   **Dependencies**: `flutter pub get`
*   **Build**: `flutter build`
*   **Testing**:
    *   Run all tests: `flutter test`
    *   Run a single test file: `flutter test WebAccountingApp/frontend/flutter_app/test/widget_test.dart` (example path)
*   **Linting**: `flutter analyze`
*   **Code Style**: Follow effective Dart guidelines. Maintain consistent widget structure and use explicit type annotations.
