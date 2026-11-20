# AAA Tools Hub – source code (Cost Code 2026, judge access)

A web app (FastAPI + Jinja, SQLite) that runs on one Windows machine with Microsoft Excel and serves the office network.

  hub/              web app: pages, accounts and access, job queue (worker.py), file naming, cost log, update chain
  toolkit/          the steps: Area Calculations (pipeline, workbook, 3D model, offline copy, Annex D slides)
  toolkit/pte/      PTE steps 01-06: unlock, register, stack divisions, name assets, load BOQs, link rates
                    (Excel itself is driven through PowerShell: excel_ops.ps1, excel_com.ps1)
  tools/            start the hub for the office network (start_hub_lan.ps1), readiness check (check_lan.ps1)
  pipeline.py       the Areas build, run by the hub's worker
  config/           model prices for the Claude API cost log

Run: setup.ps1 once, then  powershell -ExecutionPolicy Bypass -File tools\start_hub_lan.ps1
Accounts: python -m hub.manage adduser <email> "<Name>" --admin

Left out of this pack: project data, the hub database, specs and decks from real projects, and internal notes.
Names of real projects in code comments were replaced with sample names.
