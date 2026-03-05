@echo off
:: ============================================================
:: Watch-ClaudeHealth.bat
:: Claude Desktop / Cowork -- Health Monitor (Manual Launch)
:: ============================================================
::
:: Usually started automatically by the Prevent-ClaudeIssues scheduled task.
:: Run this manually if you want to monitor in the foreground and see
:: real-time health check output.
::
:: The monitor watches Claude's logs for VirtioFS mount failures
:: and auto-runs the fix script when errors are detected.
::
:: Close this window to stop monitoring.
::

title Claude Desktop - Health Monitor
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Watch-ClaudeHealth.ps1"
