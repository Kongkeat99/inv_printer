Optional - needed only for the "Print PDF files" mode.

1. Download the portable SumatraPDF (64-bit) from https://www.sumatrapdfreader.org/download-free-pdf-viewer
2. Rename it to  SumatraPDF.exe  and put it in THIS folder.
3. Run build_installer.bat - the installer copies it next to InvoicePrinter.exe, and the app finds it automatically.

No SumatraPDF? The app falls back to Adobe Acrobat/Reader if installed, then to the Windows "printto" command of the
default PDF program (works with Adobe/Foxit, NOT with Microsoft Edge as the default PDF app).
You can also set  "pdf_viewer_path"  in invoice_printer_config.json to any SumatraPDF.exe / AcroRd32.exe location.
