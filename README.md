# Excel Image Downloader

This project downloads product images from an Excel file and saves them using the matching `Mã hàng` value as the filename.

The script reads the first worksheet and expects these exact column headers:

- `Mã hàng`
- `Hình ảnh (url1,url2...)`

If a row contains multiple image URLs, the files are saved as `MAHANG_1`, `MAHANG_2`, and so on.

## Requirements

- Python 3.9 or newer
- Internet access to download the product images
- An Excel file in `.xlsx` format

Python dependencies:

- `openpyxl`
- `pyinstaller` for executable builds

## Install Dependencies

macOS or Linux:

```bash
python3 -m venv venv
source venv/bin/activate
pip install openpyxl pyinstaller
```

Windows PowerShell:

```powershell
py -m venv venv
.\venv\Scripts\Activate.ps1
pip install openpyxl pyinstaller
```

## Run the Script

Default behavior:

```bash
python app.py
```

This uses:

- Input file: `SP.xlsx`
- Output folder: `downloaded_images`

Custom input file:

```bash
python app.py "ds sp moi.xlsx"
```

Custom input file and output folder:

```bash
python app.py "ds sp moi.xlsx" "downloaded_images"
```

## Build an Executable with PyInstaller

PyInstaller does not cross-compile well between operating systems. Build the executable on the same platform you want to run it on:

- Build on Windows for Windows
- Build on macOS for macOS
- Build on Linux for Linux

The build output is created in the `dist` folder.

The project uses this PyInstaller command:

```bash
pyinstaller --noconfirm --clean --windowed --onedir --name "kv-image-downloader" app.py
```

## GitHub Actions Builds

GitHub Actions workflows are included for Windows, Linux, and macOS builds:

- `.github/workflows/windows-build.yml`
- `.github/workflows/linux-build.yml`
- `.github/workflows/macos-build.yml`
- `.github/workflows/release.yml`

The three platform build workflows:

- checks out the repository
- installs Python 3.11 plus `openpyxl` and `pyinstaller`
- defines a workflow version such as `v0.1`
- runs:

```bash
pyinstaller --noconfirm --clean --windowed --onedir --name "kv-image-downloader-v0.1" app.py
```

- packages a versioned release archive such as `kv-image-downloader-v0.1-windows.zip`
- uploads that archive as a workflow artifact

Update the `VERSION` value in the platform build workflow files when you want a new artifact version there.

You can run them from the GitHub Actions tab with `workflow_dispatch`, or let them run automatically on pushes to `main` and on pull requests.

## GitHub Releases

The `release.yml` workflow publishes built files as GitHub Release assets.

- Trigger it by pushing a tag such as `v0.1`
- Or run it manually from the Actions tab and provide a version like `v0.1`
- It uses that tag or manual input as the release version
- It builds Windows, Linux, and macOS archives
- It creates or updates a GitHub release and uploads:

```text
kv-image-downloader-v0.1-windows.zip
kv-image-downloader-v0.1-linux.zip
kv-image-downloader-v0.1-macos.zip
```

Example tag commands:

```bash
git tag v0.1
git push origin v0.1
```

Notes:

- `--onedir` creates a folder-based app instead of a single-file executable
- `--windowed` hides the console window on GUI platforms such as Windows and macOS
- If you want terminal logs during app launch, remove `--windowed`

### macOS

```bash
source venv/bin/activate
pyinstaller --noconfirm --clean --windowed --onedir --name "kv-image-downloader" app.py
```

macOS output:

```bash
dist/kv-image-downloader.app
```

When you run the macOS `.app`, place `SP.xlsx` next to `kv-image-downloader.app`, not inside the app bundle.

If you are building inside a restricted shell and PyInstaller reports a cache permission error, run:

```bash
PYINSTALLER_CONFIG_DIR=/tmp/pyinstaller pyinstaller --noconfirm --clean --windowed --onedir --name "kv-image-downloader" app.py
```

The bundle also includes a folder build at:

```bash
dist/kv-image-downloader
```

### Linux

```bash
source venv/bin/activate
pyinstaller --noconfirm --clean --windowed --onedir --name "kv-image-downloader" app.py
```

Linux output:

```bash
./dist/kv-image-downloader/kv-image-downloader
```

### Windows

```powershell
.\venv\Scripts\Activate.ps1
pyinstaller --noconfirm --clean --windowed --onedir --name "kv-image-downloader" app.py
```

Windows output:

```powershell
.\dist\kv-image-downloader\kv-image-downloader.exe
```

## Build Artifacts

After running PyInstaller, you will typically see:

- `build/` for temporary build files
- `dist/kv-image-downloader/` for the onedir build
- `dist/kv-image-downloader.app` on macOS
- `kv-image-downloader.spec` for the PyInstaller build configuration

## Notes

- If `app.py` is run without arguments, it looks for `SP.xlsx` in the current working directory.
- Downloaded files are saved into `downloaded_images` by default.
- The script keeps going even if some image URLs fail, then prints a summary at the end.
