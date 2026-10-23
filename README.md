# DirectoryToc.ps1

Creates a formatted Word document (`.docx`) that lists a directory structure as a
table of contents. It walks a folder recursively and writes **only the
directories** (files are omitted) as an indented, hierarchical listing — like a
table of contents for your folder tree.

---

## Is it standalone?

**Yes — 100% standalone.** The script has:

- **No external dependencies** — no modules to install, nothing to download
- **No Word / Office requirement** — it writes the `.docx` file itself using
  .NET classes that ship with Windows
- **No internet access needed**
- **A single file** — just `DirectoryToc.ps1`, copy it anywhere and run it

It only uses components built into Windows itself:

| Component | Where it comes from |
|---|---|
| PowerShell 5.1 (or 7+) | Built into Windows 10/11 |
| `System.IO.Compression` | Built into .NET Framework (ships with Windows) |
| `Get-ChildItem`, `Test-Path`, ... | PowerShell built-in cmdlets |

---

## Requirements

- Windows 10 / 11 (or any Windows with PowerShell 5.1+, which is every
  supported Windows version)
- Write permission for the output folder
- That's it.

---

## Quick start (for beginners)

### Step 1 — Allow PowerShell to run scripts

Windows blocks `.ps1` scripts by default ("running scripts is disabled on this
system"). You only need to change this **once**. Open a normal PowerShell window
(Start → type `powershell` → Enter) and run:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

- You may need to confirm with `Y` (Yes).
- This setting applies to **your user account only** — no admin rights needed.
- `RemoteSigned` means: local scripts (like this one) run normally; scripts
  downloaded from the internet must be signed. This is the recommended,
  safe setting for everyday work.

If you prefer **not** to change any settings, skip this step and always run the
script with the one-off command shown in the next section.

### Step 2 — Run the script

```powershell
cd C:\path\to\this\folder
.\DirectoryToc.ps1 -Directory "C:\Projects\MyApp"
```

Result: `DirectoryStructure.docx` in your current folder.

### Alternative: run once without changing any settings

If you don't want to touch the execution policy at all, run the script through a
policy override that applies to that single run only:

```powershell
powershell -ExecutionPolicy Bypass -File .\DirectoryToc.ps1 -Directory "C:\Projects\MyApp"
```

### Alternative: allow scripts for the current terminal window only

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\DirectoryToc.ps1
```

This resets automatically when you close the window.

### If you downloaded the script from the internet

Windows may block downloaded files even with a permissive policy. Unblock it once:

```powershell
Unblock-File .\DirectoryToc.ps1
```

---

## Usage

```
.\DirectoryToc.ps1 [-Directory <path>] [-OutputFile <path>] [-Title <text>] [-Format <Default|Simple>]
```

All parameters are optional. Defaults: current directory, `.\DirectoryStructure.docx`,
title "Table of Contents", format `Default`.

| Parameter | Position | Default | Description |
|---|---|---|---|
| `-Directory` | 0 | current directory | Root folder to scan (subdirectories are listed recursively) |
| `-OutputFile` | 1 | `.\DirectoryStructure.docx` | Where to write the Word document |
| `-Format` | 2 | `Default` | `Default` = styled TOC, `Simple` = plain text-editor-like listing |
| `-Title` | named | `Table of Contents` | Document title line |

### Examples

```powershell
# Current directory, default output name
.\DirectoryToc.ps1

# Specific folder, specific output file
.\DirectoryToc.ps1 -Directory "C:\Projects\MyApp" -OutputFile "C:\temp\MyApp.docx"

# Positional shorthand: directory then output file
.\DirectoryToc.ps1 "C:\Projects\MyApp" "C:\temp\MyApp.docx"

# Custom title
.\DirectoryToc.ps1 -Directory "C:\Projects\MyApp" -Title "MyApp Folder Structure"

# Plain, text-editor-like listing (see Formats below)
.\DirectoryToc.ps1 -Directory "C:\Projects\MyApp" -Format Simple
```

---

## Output formats

### Default (styled table of contents)

- Centered title, scanned root path and generation stamp, separated by a rule
- Level 1 entries **bold 12 pt**, level 2 = 11 pt, level 3+ = 10 pt
- 0.25 inch indent per level, Calibri

### Simple (`-Format Simple`)

Looks like a plain text listing (Notepad-style):

- Consolas 10 pt throughout — no bold, no size changes, no colors
- 2 spaces of indentation per level
- No extra spacing between lines — rows sit directly beneath each other

Example of the Simple output:

```
Table of Contents
C:\Projects\MyApp
Generated 2026-10-23 21:00 - 11 directories

Alpha
  Beta
  Gamma
    Delta
Empty
R&D & (100%)
Ünïcödé Tëst
Zeta
```

---

## Behavior notes

- **Files are omitted** — only directories appear in the document.
- **Ordering** — depth-first: children directly below their parents, siblings
  sorted alphabetically (case-insensitive).
- **Unreadable folders** (permission errors) are skipped silently; the rest of
  the tree is still listed.
- **Hidden and system directories** are not included (PowerShell's default
  behavior without `-Force`).
- **Junctions/symlinks** appear once and are not followed into, so there are no
  infinite loops.
- **Very large trees are supported** — the document is streamed to disk. Tested
  with 380,000+ directories (≈ 3 MB `.docx`, about 2½ minutes). Memory use stays
  flat regardless of tree size.
- **Exit codes** — `0` = success, `1` = invalid input or write error.
- The output file is **overwritten** if it already exists. If it is open in Word,
  you get a clear error instead of a crash — close the file and re-run.

---

## Troubleshooting

| Problem | Cause / fix |
|---|---|
| *"running scripts is disabled on this system"* | See [Step 1](#step-1--allow-powershell-to-run-scripts) above, or run with `powershell -ExecutionPolicy Bypass -File .\DirectoryToc.ps1` |
| *"Directory not found"* | Check the path; quote it if it contains spaces |
| *"Cannot overwrite ... Is the file open in another program?"* | Close the `.docx` in Word and run again |
| *"Output directory not found"* | The folder part of `-OutputFile` must exist |
| Missing folders in the output | They are hidden/system, or access was denied |
| Warning *"does not have a .docx extension"* | Harmless — Word may still open it, but prefer a `.docx` extension |

---

## Files

| File | Purpose |
|---|---|
| `DirectoryToc.ps1` | The script (this is all you need) |
| `README.md` | This file |
| `_tmp/` | Scratch folder for generated samples; ignored by git |
