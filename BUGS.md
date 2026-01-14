# WordTrack Bug List

This document tracks identified bugs and issues in WordTrack. Bugs are listed by priority and status.

## Open Bugs

### High Priority

**BUG-001: Table Selection Causes General Exception on Insert**
- **Status:** Identified (not fixed)
- **Description:** When a table is selected in Word and the user attempts to insert Claude's response, a General Exception error occurs. The selection appears to work correctly until insertion is attempted.
- **Steps to Reproduce:**
  1. Select a table in a Word document
  2. Click "Get Selected Text" (appears to work)
  3. Send text to Claude and get response
  4. Click "Insert Claude's Response"
  5. General Exception error occurs
- **Expected Behavior:** Should either handle table selections gracefully or show a clear error message indicating tables are not supported
- **Actual Behavior:** General Exception error
- **Reported:** 2026-01-14
- **Notes:** May need to detect table selections and either prevent them or handle them differently

## Fixed Bugs

*(None yet - this is a new bug tracking document)*

## Known Limitations

- Office.js Word API does not provide programmatic access to Word's built-in spelling/grammar checking tools
- Track Changes must be enabled manually by users (programmatic enabling was removed due to reliability issues)
