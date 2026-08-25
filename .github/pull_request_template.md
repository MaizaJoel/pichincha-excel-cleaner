## 📝 Description of Changes
<!-- Provide a brief description of what this PR introduces, fixes, or refactors. -->

---

## 🔒 Security & Offline Integrity Checklist
<!-- All checks must be verified before this PR can be merged -->

- [ ] **100% Offline**: This PR introduces **NO** network requests, sockets, telemetry, or remote API calls.
- [ ] **AST Verification**: `pytest test_offline_integrity.py` passes with zero violations.
- [ ] **Unit Tests**: `pytest test_engine.py` passes and new test cases have been added for new features.
- [ ] **Zero Sensitive Data**: No real personal/financial data, bank statements, or credentials are included in this PR.
- [ ] **Code Cleanliness**: No use of `eval()`, `exec()`, or unvetted external scripts.
