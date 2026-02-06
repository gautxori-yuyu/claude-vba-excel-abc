# VBA File Integrity Protocol

## Encoding
- ALL .cls / .bas / .frm files are ISO-8859-1 (Windows-1252 / ANSI). Never UTF-8. No BOM.
- PROHIBITED: U+FFFD (EF BF BD). If detected, STOP — do not attempt repair.

## Line Endings
- Per-file: some files in this repo are CRLF, others are LF. Check with `file` command before editing.
ALL files should be CRLF
- New files: use CRLF (standard VBA export format).

## Edit Method
- Use Python binary read/write (open 'rb' / 'wb'). NEVER use Read+Edit tools on these files
  (they normalize encoding to UTF-8, corrupting the file).
- Pattern-match on exact bytes. Verify with debug output (repr()) if a pattern is not found.
- Lines not explicitly edited must remain byte-identical to source.

## VBA Attribute Lines
- If removing WithEvents, also remove the corresponding Attribute line.

## Compilation Note
- VBA compiles ALL code, even after Stop statements. Dead code referencing removed
  functions causes compile errors. Remove dead code when removing functions.
