# Build Mac Installer Without Staff Installing Xcode

Staff MacBooks do not need Xcode or Command Line Tools if you give them the finished `.pkg`.

Use GitHub Actions to build the Mac installer in the cloud:

1. Put this project in a GitHub repository.
2. Go to the repository on GitHub.
3. Open `Actions`.
4. Select `Build Mac Installer`.
5. Click `Run workflow`.
6. When the run finishes, download the artifact named `Teq-Shipment-Docs-Generator-pkg`.

The downloaded artifact contains:

- `Teq Shipment Docs Generator.pkg`

Staff only need to open that `.pkg` and install it. They do not need Python, PyInstaller, Xcode, or Command Line Tools.

If macOS blocks the app because it is not notarized, right-click the app and choose `Open`, or manage it through your normal company security policy.
