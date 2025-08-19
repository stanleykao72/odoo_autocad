# PowerShell script to start GUI with MCP Server
Write-Host "Starting AutoCAD-Odoo Integration GUI with MCP Server..." -ForegroundColor Green
Write-Host "========================================================"
Write-Host ""

Write-Host "Launching GUI with auto-start MCP Server..." -ForegroundColor Yellow
python odoo.py --enable-mcp

Write-Host ""
Write-Host "GUI and MCP Server startup completed." -ForegroundColor Green
Read-Host "Press Enter to exit"