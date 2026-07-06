```markdown
# Sprint_sheet_generator Development Patterns

> Auto-generated skill from repository analysis

## Overview
This skill teaches you how to contribute to the `Sprint_sheet_generator` Python codebase, which is designed to generate sprint sheets for project management or agile workflows. The repository uses Python without a specific framework, and follows consistent conventions for file naming, imports, and exports. You'll learn how to maintain code style, structure tests, and execute common development workflows.

## Coding Conventions

### File Naming
- Use **snake_case** for all file names.
  - Example: `sprint_sheet_generator.py`, `data_loader.py`

### Import Style
- Use **relative imports** within the package.
  - Example:
    ```python
    from .utils import generate_sheet
    ```

### Export Style
- Use **named exports** (i.e., define functions/classes explicitly for import).
  - Example:
    ```python
    def generate_sheet(...):
        ...
    ```

### Commit Patterns
- Commit messages are freeform, with no enforced prefix.
- Average commit message length: ~56 characters.
- Example:
  ```
  Add function to parse sprint data from CSV
  ```

## Workflows

### Adding a New Feature
**Trigger:** When implementing new functionality.
**Command:** `/add-feature`

1. Create a new Python file using snake_case if needed.
2. Use relative imports to access existing modules.
3. Define new functions or classes with explicit names.
4. Write or update tests in a corresponding `*.test.*` file.
5. Commit your changes with a clear, descriptive message.

### Fixing a Bug
**Trigger:** When resolving a reported issue.
**Command:** `/fix-bug`

1. Locate the relevant module using snake_case file names.
2. Apply the fix, maintaining code style and import conventions.
3. Update or add tests to cover the bug scenario.
4. Commit with a message describing the fix.

### Running Tests
**Trigger:** To verify code correctness after changes.
**Command:** `/run-tests`

1. Identify test files (pattern: `*.test.*`).
2. Run tests using the preferred Python test runner (framework not specified; try `pytest` or `unittest`).
   - Example:
     ```
     pytest
     ```
3. Review test results and address any failures.

## Testing Patterns

- Test files follow the pattern: `*.test.*` (e.g., `sprint_sheet_generator.test.py`).
- The specific testing framework is not enforced; use standard Python testing tools like `unittest` or `pytest`.
- Place test functions/classes in these files to cover new and existing features.

**Example test file:**
```python
import unittest
from .sprint_sheet_generator import generate_sheet

class TestSprintSheetGenerator(unittest.TestCase):
    def test_generate_sheet(self):
        result = generate_sheet(...)
        self.assertEqual(result, expected_output)
```

## Commands
| Command       | Purpose                                    |
|---------------|--------------------------------------------|
| /add-feature  | Scaffold and implement a new feature       |
| /fix-bug      | Apply and test a bug fix                   |
| /run-tests    | Run all test files in the repository       |
```
