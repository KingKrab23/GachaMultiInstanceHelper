@echo off
echo Starting Macro Automator GUI...

REM Change to the script directory
cd /d "%~dp0"

REM Build the C# project if not built already
cd csharp_gui
dotnet build

REM Start the GUI with default settings
REM To run with custom options, use one of the following examples:
REM 
REM 1. Start with a specific config file:
REM    start "" "bin\Debug\net6.0-windows\MacroAutomatorGUI.exe" --config "path\to\config.yaml"
REM
REM 2. Start a specific sequence immediately:
REM    start "" "bin\Debug\net6.0-windows\MacroAutomatorGUI.exe" --start "SequenceName" --iterations 5
REM
REM 3. Start a sequence to run forever:
REM    start "" "bin\Debug\net6.0-windows\MacroAutomatorGUI.exe" --start "SequenceName" --iterations 0
REM

start "" "bin\Debug\net6.0-windows\MacroAutomatorGUI.exe"

echo Macro Automator started!
