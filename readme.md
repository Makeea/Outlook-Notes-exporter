# Export Outlook Notes

Microsoft is forcing users into **New Outlook**, which:
- Removes the Notes feature entirely
- Breaks COM access to `IPM.StickyNote` items
- Offers no export, migration, or backup option
- Leaves years of personal and business data stranded

Classic Outlook still supports Notes via COM, but this will not last.  
Once you switch to New Outlook, your Notes become inaccessible unless you’ve backed them up.

## Why This Script Exists
There is no official migration tool for Outlook Notes.  
Users who rely on them are being pushed into a dead end.  
This script exists so you can **export and save every note you have** while you still can, with full metadata, in a human-readable format.

## What the Script Does
- Connects to Outlook (classic) via COM automation
- Scans all mail stores for Notes folders
- Exports each note to Markdown (`.md`) or plain text (`.txt`)
- Preserves metadata: creation date, modification date, categories, color, store name, folder path
- Creates safe filenames by removing or replacing illegal characters
- Optionally flattens folder structure or keeps store-based subfolders
- Can prefix dates to filenames for easier sorting
- Shows a live progress bar and current note title during export
- Can generate a CSV index file of all exported notes

## Requirements
- **Windows** with Outlook (classic) installed  
- PowerShell 5.1+ or PowerShell 7+  
- 64-bit PowerShell if Outlook is 64-bit

## How to Use
1. Download `export-outlook-notes.ps1`
2. Place it in your `Documents\Outlook Files` folder or another location
3. Unblock the script:
   ```powershell
   Unblock-File .\export-outlook-notes.ps1
   ```
4. Open 64-bit PowerShell
5. Navigate to the script folder:
   ```powershell
   cd "C:\Users\<YourName>\Documents\Outlook Files"
   ```
6. Run it:
   ```powershell
   .\export-outlook-notes.ps1
   ```
7. Follow prompts for:
   - **Output format** (`md` or `txt`)
   - **Prefix dates** in filenames (`Y/N`)
   - **Folder structure**: flat or per-store
   - **CSV index file**: include or skip

## Output Details
- Files saved under `notes_md` or `notes_txt` in the script directory
- Filenames example:  
  - With date prefix: `20250813-Project-Ideas.md`  
  - Without date prefix: `Project-Ideas.md`
- Folder structure matches Outlook stores unless flattened

## Limitations
- Will not work in **New Outlook** (no COM Notes support)
- Will not export from accounts where Notes are disabled or blocked
- Requires Outlook to be installed and configured locally

## Recommendation
Run this before Microsoft forces you into New Outlook.  
Once you lose COM access, the only way to recover Notes will be from an old backup or PST file.

# F U Microsoft!!!
