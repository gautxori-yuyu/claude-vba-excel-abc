# VBA File Integrity Protocol

## Encoding
- ALL .cls / .bas files are ISO-8859-1 (Windows-1252 / ANSI). Never UTF-8. No BOM.
- Spanish accented chars are single bytes: ó=\xf3 á=\xe1 é=\xe9 í=\xed ú=\xfa ñ=\xf1 ü=\xfc ¿=\xbf ¡=\xa1
- Uppercase: Ó=\xd3 Á=\xc1 É=\xc9 Í=\xcd Ú=\xda Ñ=\xd1
- PROHIBITED: U+FFFD (EF BF BD). If detected, STOP — do not attempt repair.

## Line Endings
- Per-file: some files in this repo are CRLF, others are LF. Check with `file` command before editing.
- NEVER normalize. Modified lines must match the file's existing style.
- New files: use CRLF (standard VBA export format).
- Detection reference (as of last check):
  - CRLF: clsOpportunity.cls, clsRibbonState.cls, modCALLBACKSRibbon.bas, clsFileState.cls,
           clsFileManager.cls, ThisWorkbook.cls, clsOpportunitiesMgr.cls
  - LF:   clsRibbon.cls, clsOpportunityState.cls, clsEventsMediatorDomain.cls,
           clsApplicationState.cls, clsEventsMgrInfrastructure.cls, clsApplication.cls,
           clsChartMgr.cls, clsExecutionContext.cls

## Edit Method
- Use Python binary read/write (open 'rb' / 'wb'). NEVER use Read+Edit tools on these files
  (they normalize encoding to UTF-8, corrupting the file).
- Pattern-match on exact bytes. Verify with debug output (repr()) if a pattern is not found.
- Lines not explicitly edited must remain byte-identical to source.

## VBA Attribute Lines
- `Attribute VarName.VB_VarHelpID = -1` must appear on the line IMMEDIATELY after any
  `Private WithEvents VarName As ...` declaration. Required for VBA export/import.
- If removing WithEvents, also remove the corresponding Attribute line.

## Compilation Note
- VBA compiles ALL code, even after Stop statements. Dead code referencing removed
  functions causes compile errors. Remove dead code when removing functions.
