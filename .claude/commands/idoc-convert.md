Convert an SAP IDoc flat file (.txt) to an Excel documentation file (.xlsx).

Usage: /idoc-convert <path-to-idoc-file>

Steps:
1. The argument is: $ARGUMENTS
2. Verify the file exists. If it doesn't, report the error clearly in Polish.
3. Run the conversion using the project venv:
   - Working directory: d:\WORK\OneDrive\Dokumenty\apps_data\claude\code\idoc-to-excel-repo
   - Command: .\venv\Scripts\python idoc_parser.py "$ARGUMENTS"
4. Report the output .xlsx file path to the user in Polish.
5. If any error occurs during conversion, show the error message clearly.
