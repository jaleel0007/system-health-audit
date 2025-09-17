# system-health-audit

Cross-platform system audit scripts (Windows/Linux) that collect OS, CPU, memory, disks, patching, and AV status, and export **JSON** and **Markdown** reports.

## Quickstart (Windows)

```powershell
# from repo root
# create reports dir if missing
New-Item -ItemType Directory -Path .\reports -Force | Out-Null

# run (PowerShell 7+ recommended)
pwsh -File .\scripts\audit_windows.ps1 -OutDir .\reports -Json -Markdown
# outputs like:
# reports/windows_<HOST>_<yyyyMMdd_HHmmss>.json
# reports/windows_<HOST>_<yyyyMMdd_HHmmss>.md
