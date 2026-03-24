# Security Test Checklist

Run these tests after all security changes are complete.

## 1. Ghostscript Subprocess Security
- [ ] Process a file named `evil; calc.exe .pdf` — app should handle it, no calc.exe opens
- [ ] Process a file named `../../etc/passwd.pdf` — no path traversal
- [ ] Process a file named `"file with spaces".pdf` — works normally
- [ ] Confirm `-dSAFER` flag is present in Ghostscript calls (code review)
- [ ] Confirm all subprocess calls use list args, not shell strings (code review)

## 2. Auto-Updater Signature Verification
- [ ] Open an older version (e.g. v1.5.5) — should detect update and show banner
- [ ] Check release assets include `.sig` file alongside installer
- [ ] Manually corrupt the downloaded installer — app should reject it ("Update download failed")
- [ ] Manually edit the `.sig` file hash — app should reject it
- [ ] Confirm update only downloads from `https://github.com/` URLs

## 3. Temp File Cleanup (Legal Doc Leakage)
- [ ] Compress a PDF, close app normally, search `%TEMP%` for `compressed.pdf` — should be gone
- [ ] Merge PDFs, close app normally, search `%TEMP%` for `merged.pdf` — should be gone
- [ ] Process a PDF, kill app via Task Manager, check `%TEMP%` for leftovers (known gap — crash cleanup)
- [ ] Use Process Monitor (Sysinternals): filter by process name, watch file write events during processing

## 4. Explorer Shell Injection
- [ ] Double-click a completed file with spaces in the name — Explorer opens correctly
- [ ] Double-click a completed file with special characters in path — no command injection

## 5. Malicious PDF Fuzzing
- [ ] Test with EICAR test PDFs (safe malware test patterns)
- [ ] Test with malformed PDFs (truncated, corrupt headers, oversized objects)
- [ ] Test with PDFs containing JavaScript — should not execute
- [ ] Test with polyglot PDFs (PDF that is also exe/zip)
- [ ] Monitor for: crashes, unexpected processes spawning, files written outside temp/output dirs
- [ ] Run with Process Monitor to catch any unexpected file/registry/network activity

## 6. Installer & DLL Hijacking
- [ ] After install, run `icacls "$env:LOCALAPPDATA\Programs\PDF Tool"` — only your user should have write access
- [ ] Drop a fake `version.dll` in the same folder as `PDF Tool.exe` — app should NOT load it
- [ ] Drop a fake `winmm.dll` in the same folder — app should NOT load it
- [ ] Run Process Monitor during install — filter `Load Image` operations, check no DLLs loaded from user-writable paths outside the app/temp dirs
- [ ] Confirm installer uses `PrivilegesRequired=lowest` (no admin needed)
- [ ] Confirm no files installed to `C:\` root or world-writable locations

## 7. Static Analysis (Bandit)
- [x] No `shell=True` in subprocess calls
- [x] No `eval()` or `exec()`
- [x] No pickle/marshal deserialization
- [x] No hardcoded passwords or secrets (public key is not a secret)
- [x] No weak hashing (SHA-256 used)
- [x] No `os.system()` calls
- [x] URL validation on both download_url and sig_url
- [x] Specific exception types caught (not bare `except Exception`)

## 8. General
- [ ] App launches without errors on clean install
- [ ] Compress tab works with PDF and image files
- [ ] Merge tab works with multiple PDFs
- [ ] Only Compress and Merge tabs visible in shipped exe
- [ ] All tabs visible when running from source
