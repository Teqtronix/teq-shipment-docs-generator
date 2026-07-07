# Mac Build Notes

Build the macOS app on a Mac. PyInstaller cannot create a working `.app` from Windows.

Steps:

1. Install Python 3 for macOS.
2. Open Terminal.
3. Go to this folder.
4. Run:

```bash
chmod +x build_mac.command
./build_mac.command
```

Output:

- `dist/Teq Shipment Docs Generator.app`
- `dist/Teq Shipment Docs Generator-mac.zip`
- `dist/Teq Shipment Docs Generator.pkg`

The `.pkg` installs the app into `/Applications`.

Current macOS limitation:

- Indigo PDF labels are generated directly in Python and should work.
- CI/PL Excel workbook generation should work.
- TJX factory label PDF mail merge currently requires the Windows version of Microsoft Word automation. On macOS, the app will return a clear error for that PDF merge step instead of running the Windows-only PowerShell path.
