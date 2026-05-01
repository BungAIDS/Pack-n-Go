Generate a self-contained VBA `LogXxx` subroutine that can be dropped into any SolidWorks macro, following the exact same pattern used in `SaveAs_Export.bas` (`LogExport`).

The user has typed: $ARGUMENTS

1. If they gave a macro name, use it. If not, ask: "What do you want to call the macro?"
2. Ask: "What does each run log? Describe what happened (e.g. which file was processed, what the result was) and I'll turn that into column names for you."
3. Ask: "What folder should the log files go in?"

Ask all missing questions in a single message, then generate once you have the answers. Do not ask the user to specify column names directly — figure them out from their description of what the macro does.

Generate a complete, ready-to-paste VBA subroutine that:
- Has a clear signature with one `ByVal` parameter per custom column
- Uses the same constants pattern: `LOG_DIR`, `LOG_XLSX`, `LOG_OVERFLOW`, `HEADER_ROW = 3`, `DATA_START = 4`
- Creates the log directory if it doesn't exist (late-bound FSO)
- Opens Excel late-bound (`CreateObject("Excel.Application")`)
- Handles a locked/read-only .xlsx by falling through to a `GoTo Overflow` CSV fallback
- On new file: names the sheet, writes the bold summary row (`"Total Runs"` in A1, `=COUNTA(A4:A1048576)` formula in B1), writes bold column headers in row 3 (Date, Time, User, then the custom columns)
- On existing file: finds the last used row with `End(-4162)`, backfills any missing headers
- Writes the data row: `Format(Now,"YYYY-MM-DD")`, `Format(Now,"HH:MM:SS")`, `Environ("USERNAME")`, then the custom values
- After writing: re-asserts `=COUNTA(A4:A1048576)` in B1 (upgrades any old file that had a hardcoded number), AutoFits all used columns
- Saves (`.Save` if file existed, `.SaveAs logXlsx, 51` if new), closes workbook, quits Excel, cleans up all object references
- In the `Overflow:` section: creates the CSV with a header line if it doesn't exist, then appends a comma-delimited row

After the code block, show a one-liner example of how to call the new sub from the macro's main action sub.
