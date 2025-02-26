# Macro Automator GUI

A C# Windows Forms application to control the Macro Automator python script. This GUI provides a user-friendly interface for creating, managing, and executing macro sequences.

## Features

- Start/stop macro sequences with hotkeys (F5 to start, ESC to stop)
- Set iteration count or loop indefinitely 
- Real-time logging of macro execution
- Dark mode UI for reduced eye strain during long sessions

## Requirements

- .NET 6.0 or later
- Windows 10 or later
- Python 3.7+ (for the backend macro execution)

## Building the Application

1. Open the solution in Visual Studio 2022
2. Build the solution (Ctrl+Shift+B)
3. Run the application (F5)

## Integration with Python Backend

The GUI communicates with the Python macro automator script by:
1. Executing the script as a subprocess
2. Sending commands via command-line arguments
3. Reading output from the script's stdout

## Usage

1. Select a macro sequence from the list
2. Set the number of iterations (or check "Loop Forever")
3. Click "Start Sequence" or press F5
4. To stop execution, click "Stop" or press ESC

## Example Sequences

The application comes with several example sequences:
- Click Center: Clicks at the center of the active window
- Type Hello World: Types "Hello World" in the active window
- Press Enter 5 Times: Presses the Enter key 5 times with delays
