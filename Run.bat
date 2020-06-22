@echo off
PowerShell.exe -NoProfile -Command "& {Start-Process PowerShell.exe -Argument '-NoProfile -ExecutionPolicy Bypass -File %~dp0pre_migration_master.ps1 -generateReports' -Verb RunAs}"
