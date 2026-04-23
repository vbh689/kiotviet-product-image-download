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

Python dependencies are listed in `requirements.txt`.

## Install Dependencies

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
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

If building on macOS and you get a cache permission error:
```bash
PYINSTALLER_CONFIG_DIR=/tmp/pyinstaller pyinstaller --noconfirm --clean --windowed --onedir --name "kv-image-downloader" app.py
```

## GitHub Actions Builds

`.github/workflows/release.yml` builds for Windows, Linux, and macOS on every push to `main` and on pull requests. It also publishes release assets when a tag matching `v*` is pushed or when triggered manually from the Actions tab.

To publish a release:

```bash
git tag v0.1
git push origin v0.1
```

This creates:

```text
kv-image-downloader-v0.1-windows.zip
kv-image-downloader-v0.1-linux.zip
kv-image-downloader-v0.1-macos.zip
```

## Notes

- `--onedir` creates a folder-based app instead of a single-file executable
- `--windowed` hides the console window on GUI platforms such as Windows and macOS
- If you want terminal logs during app launch, remove `--windowed`
- If `app.py` is run without arguments, it looks for `SP.xlsx` in the current working directory.
- Downloaded files are saved into `downloaded_images` by default.
- The script keeps going even if some image URLs fail, then prints a summary at the end.
