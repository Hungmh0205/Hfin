@echo off
taskkill /f /im HFin.exe 2>nul
dotnet build "%~dp0AccountingOcrTest.csproj"
dotnet run --project "%~dp0AccountingOcrTest.csproj"