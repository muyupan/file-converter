# file-converter
This file converter converts TSX files to XLSX format and Markdown files to DOCX format. 
It features drag-and-drop support and allows users to select input files from folders. 
Users can choose where to save the converted files.

# Installation

## Quick Install (macOS users)
1. Go to the [Releases](../../releases) page
2. Download `File.Converter.app.zip`
3. Unzip the file
4. Move `File Converter.app` to your Applications folder
5. Right-click the app and select "Open" (required only first time)

## Build from Source (Developers)
1. Clone the repository:
```bash
git clone https://github.com/YOUR_USERNAME/file-converter.git
cd file-converter
```

2. Create and activate virtual environment:
```bash
python3 -m venv venv
source venv/bin/activate  # On macOS
```

3. Install requirements:
```bash
pip install -r requirements.txt
```

4. Build the app:
```bash
pyinstaller setup.spec
```

5. Find the built app in the `dist` folder
