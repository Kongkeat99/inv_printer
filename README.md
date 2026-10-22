# Invoice Printer – Windows app

Search invoice numbers → prints only the INV / PL sheets. Batch print, auto-print watch folder, PDF export/combine, print log.

## For the developer – make the installer (on a Windows PC)
Prerequisites: Python 3.10+ (tick "Add to PATH") and [Inno Setup 6](https://jrsoftware.org/isdl.php) (free).

1. Double-click **build.bat**  → builds `dist\InvoicePrinter\` (test `InvoicePrinter.exe` here first)
2. Double-click **build_installer.bat** → creates `installer_output\InvoicePrinter_Setup_2.1.0.exe`

Send that one Setup file to users.

### Releasing an update
1. Change `APP_VERSION` in `invoice_printer.py` (e.g. "2.1.1")
2. Run build.bat, then build_installer.bat
3. Users run the new Setup – it upgrades in place and **keeps their settings and log**.
   (Never change the `AppId` GUID in installer.iss.)

## For users
1. Run `InvoicePrinter_Setup_x.y.z.exe` (no admin rights needed; Windows may show a SmartScreen warning because the installer is unsigned → "More info" → "Run anyway").
2. Start **Invoice Printer** from the Start Menu / desktop.
3. First launch: choose the top folder that contains your month folders. Done.

Requirements on the user's PC: Windows 10/11 and desktop **Microsoft Excel**. "Microsoft Print to PDF" should be enabled for PDF mode.

## Data location
Installed to `%LOCALAPPDATA%\Programs\InvoicePrinter`. Settings (`invoice_printer_config.json`), `print_log.csv`,
`auto_print_state.json`, and `error_log.txt` (only after a crash) live in that same folder and survive upgrades.
Uninstalling leaves them behind; delete the folder for a full clean-up.

## Dev mode
`pip install -r requirements.txt`, then `python invoice_printer.py` (or `Invoice_Printer_DEBUG.bat` for console output).

## Optional later
- Code-sign the installer with a certificate to remove the SmartScreen warning.
- Old `Invoice_Printer.bat` launchers are no longer needed for end users.

## Printer asks for a password / department code / PIN on every print?
Do this once per PC (the app has a **Printer setup…** button that opens step 2 directly):
1. Choose the printer in the app.
2. Click **Printer setup…** → *Printing Preferences* opens.
3. Find the authentication tab (Toshiba drivers: usually **Others / Management / Department Code / User Authentication**),
   enter the department code / PIN / user name + password, tick **save / remember / don't ask again**, click OK.
4. Also check *Printer Properties → Advanced → Printing Defaults* if the setting does not stick.
5. If Windows asks for a network login for the print server: *Control Panel → Credential Manager → Windows Credentials →
   Add* the server (`\\server`) with the user / password and "Remember".
The app now prints INV + PL of a file as ONE job (`print_one_job_per_file`), so even a printer that insists on a code per job
asks only once per invoice file.
If the printer has *secure/hold print* (PIN typed at the printer), enter the PIN once there and release all jobs together.

## PDF mode (tick "Print PDF files (search .pdf) instead of Excel")
- Search looks for `.pdf` files whose name contains the invoice number (all sub-folders of the root folder) and prints
  each PDF as it is. "Print PDF", "Print ALL found", Batch and "Combine → 1 PDF" (merges the PDFs) all work in this mode.
- "All sheets" is switched off in PDF mode; searching `C264416-1` finds only files with `C264416-1` in the name.
- Printing is done by SumatraPDF if present (best: exact printer, quiet, copies), else Adobe, else the Windows `printto`
  command. Put `SumatraPDF.exe` in the `tools` folder before `build_installer.bat` (see tools\README.txt).
- Config keys: `print_pdf_files`, `pdf_extensions`, `pdf_target_folder_name`, `pdf_viewer_path`.
- Auto-print (watch folder) still handles Excel files only.
