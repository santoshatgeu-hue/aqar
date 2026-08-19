AQAR OneDrive Consolidator v3.2

Purpose
-------
Consolidate departmental AQAR submissions from a local OneDrive/SharePoint-synced folder while keeping:
1. Data-template information
2. Information explicitly required outside the Data Template
3. Supporting-document/evidence traceability
4. Department/criterion/metric hierarchy
5. Duplicate/conflict/missing-evidence review

v3.2 specifically incorporates the supplied "Information-not-in-Data-template" document.

Non-template metrics:
1.1.1, 1.3.1, 2.2.1, 2.3.1, 2.3.2, 2.3.3,
2.6.1, 2.6.2, 2.7.1, 4.1.1, 6.5.2, 6.5.3.

Usage on macOS/Ubuntu:
  chmod +x run.sh
  ./run.sh

Select the AQAR ROOT folder (not an individual department) in the GUI.
The program automatically identifies first-level department folders and treats
Criteria-* folders as criterion folders, not departments.

The program excludes output/, .git/, __pycache__, temporary Excel files and
its own generated workbook.

IMPORTANT:
This is an IQAC consolidation/review tool. It proposes mappings and flags items
for verification. It does not silently declare an AQAR value final.
