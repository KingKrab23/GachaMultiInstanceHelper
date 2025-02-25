# Windsurf Project

A GUI automation tool for managing multiple email accounts and verification codes in LDPlayer windows.

## Setup Requirements
- Python 3.x w/ PIP
- CustomTkinter
- LDPlayer emulator installed and running
- Outlook CLASSIC configured with email accounts

Optionally, run `setup_venv.bat` to set up a virtual environment.

## Setup
1. Run your LD PLayer Multi Instance, when you are ready for salted emails, run this program: python .\gui_app.py

## How to Use

The program provides a sequence of buttons that should be used in order:

1. **Rename HBR Windows to Salted Emails**
   - Renames your LDPlayer windows to match their associated email addresses
   - This helps keep track of which window corresponds to which account
   - This will update the salts and emails

2. **Type Salted Emails in LD Player Windows**
   - Automatically types the correct email address into each LDPlayer window
   - Uses the window names to determine which email goes where

3. **Enter Verification Codes in LD Player Windows**
   - Inputs verification codes into the appropriate windows
   - Requires codes to be present in verification_codes.json

4. **Scan Outlook / Force Sync**
   - **Scan Outlook for Verification Codes**: Checks Outlook for new verification codes
   - **Force Outlook Sync**: Forces Outlook to sync if emails haven't arrived. Only sometimes works.

5. **Take Screenshots of LD Player Windows**
   - Captures screenshots of all LDPlayer windows
   - Useful for looking back on rerolls.


## Data Files
- `verification_codes.json`: Stores the verification codes and timestamps for each email
- `persistent_variable.json`: Maintains persistent data between program runs

## Notes
- Ensure all LDPlayer instances are running before starting the automation
- Keep Outlook open and properly synced for verification code retrieval
