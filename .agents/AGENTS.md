# PROJECT RULES

The project workflow goal is strictly:
`PDF -> OCR -> Human Review -> SQLite -> Mapping -> Excel Export`

## Constraints
* **Do not** change the workflow.
* **Do not** replace SQLite.
* **Do not** replace openpyxl.
* **Do not** store cell mappings in source code.
* **Always** read from the Mapping Sheet dynamically.
