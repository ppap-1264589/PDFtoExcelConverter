# Build Instructions - PDF to Excel Converter

## Prerequisites

1. **Python 3.8+** with pip

## Step 1: Install Dependencies

```bash
cd ConverterApp
pip install -r requirements.txt
pip install pyinstaller
```

## Step 2: Build Executable

Run the build script:

```bash
build.sh
```

Or manually:

Output will be in: `ConverterApp\dist\PDFtoExcelConverter.exe\`

## Step 3: Test the Executable

1. Navigate to `ConverterApp\dist\`
2. Run `PDFtoExcelConverter.exe`
3. Browser should open automatically at `http://127.0.0.1:5000`
4. Test with a PDF file

## Step 4: Add Icon (Optional)

Create or add an icon file:
- Place `icon.ico` in `ConverterApp\static\`
- Or remove the `SetupIconFile` line from `installer.iss`

## Output Files

After building:

```
PDF-To-Excel-Converter/
├── ConverterApp/
│   └── dist/
│       ├── PDF-To-Excel-Converter.exe
│   ├── templates/
│   ├── static/
│   ├── uploads/
│   ├── logs/
│   └── ... (dependencies)
└──
```

## Troubleshooting

### PyInstaller errors
- Try: `pip install --upgrade pyinstaller`
- Check for missing hidden imports in the error message

### App doesn't start
- Run from command line to see error messages
- Check `logs\app.log` for details

## Notes

- The app runs a local web server on port 5000
- Browser opens automatically when the app starts
- Console window shows server logs
- Close the application to stop the server
