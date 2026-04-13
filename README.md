# Right Click to Markdown 
One-click Markdown from File Explorer

[简体中文](README.zh-CN.md) | [English](README.md)

Right-click a file in File Explorer and choose **Convert to Markdown (MarkItDown)** to create a `.md` file **in the same folder**, with the same base name. Conversion is powered by [MarkItDown](https://github.com/microsoft/markitdown). This repository adds installer scripts under `scripts/windows/` and a no-console launcher.

---

## 1. Requirements

- **OS**: Windows 10/11 (context menu is Windows-only)
- **Python**: 3.10+ (same as MarkItDown)
- A **virtual environment** is recommended (examples below use `venv`; Conda works too)

---

## 2. Install MarkItDown

**Option A: from PyPI (typical)**

In PowerShell or Command Prompt:

```bash
python -m pip install --upgrade pip
python -m pip install "markitdown[all]"
```

To install only some formats, replace `[all]` with e.g. `[pdf,docx,pptx]`.

**Option B: editable install from this repo**

If you cloned this monorepo and want to work against the local package:

```bash
cd <repository-root>
python -m pip install -e "packages/markitdown[all]"
```

**Verify** (either command succeeding is enough):

```bash
markitdown --version
```

or:

```bash
python -m markitdown --version
```

If the command is not found, ensure your Python **Scripts** folder is on the **user or system PATH**. Explorer-launched tools may see a smaller PATH than your terminal; adding that **Scripts** path to user environment variables usually fixes it.

---

## 3. Install the Explorer context menu

1. **Get the scripts**  
   Clone this repository and keep the folder `scripts/windows/` (contains `install-markitdown-context-menu.ps1`, `markitdown-convert.ps1`, `markitdown-convert-launcher.vbs`).

2. **Register the menu item**  
   Open PowerShell at the repository root and run:

   ```powershell
   Set-Location <repository-root>
   .\scripts\windows\install-markitdown-context-menu.ps1
   ```

   If execution policy blocks scripts, use:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\scripts\windows\install-markitdown-context-menu.ps1"
   ```

3. **Refresh the menu**  
   If the new verb does not appear immediately, restart **Windows Explorer** in Task Manager, or sign out and back in.

4. **Use**  
   Right-click a file → **Convert to Markdown (MarkItDown)**. On success, `basename.md` appears next to the source file.

> The installer writes **absolute paths** under `HKCU`. If you **move** this repository, run the install script again from the new location.

---

## 4. Uninstall the context menu

From the repository root:

```powershell
.\scripts\windows\install-markitdown-context-menu.ps1 -Uninstall
```

---

**Repository:** [github.com/qCanoe/right-click-to-markdown](https://github.com/qCanoe/right-click-to-markdown)

*Upstream MarkItDown: <https://github.com/microsoft/markitdown>*
