# Contributing to Banco Pichincha Excel Cleaner

Thank you for contributing! This guide outlines our contribution workflow and security standards.

---

## 🔒 Golden Rule: 100% Offline & Zero Telemetry

Because users trust this tool with their private financial bank statements:
- **No network imports**: Never import `socket`, `requests`, `urllib`, or external API clients.
- **No tracking**: Never add telemetry, analytics, or remote logging.
- **Pure local IO**: Files are read from memory/local disk and written to local disk.

---

## 🛠️ Development Setup

### 1. Clone and Install
```bash
git clone https://github.com/MaizaJoel/pichincha-excel-cleaner.git
cd pichincha-excel-cleaner
pip install -r requirements.txt
pip install pytest bandit
```

### 2. Run Tests & Security Scans
Before submitting any Pull Request, ensure all functional and security tests pass:

```bash
# Run Functional Unit Tests
pytest test_engine.py -v

# Run Offline Integrity AST Security Scanner
pytest test_offline_integrity.py -v

# Run Bandit Security Scanner
bandit -r engine.py app_gui.py
```

---

## 📋 Pull Request Guidelines

1. **Keep Changes Minimal & Focused**: Avoid unnecessary refactoring or adding unrelated libraries.
2. **Add Unit Tests**: Any new statement layout or edge-case parser fix must include a test in `test_engine.py` using **mock data**.
3. **NEVER Commit Real Bank Data**: Never commit real personal or business bank statements. Always use synthetic/anonymized fixtures.
4. **Follow the PR Template**: Ensure all checkboxes in the PR template are marked.

---

## 💡 Reporting Issues

- **Bug Reports**: Use the [Bug Report Template](.github/ISSUE_TEMPLATE/bug_report.md). Anonymize any sample row before posting!
- **Feature Requests**: Use the [Feature Request Template](.github/ISSUE_TEMPLATE/feature_request.md).
- **Security Vulnerabilities**: Follow instructions in [SECURITY.md](SECURITY.md).
