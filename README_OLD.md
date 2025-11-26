# RPA Milk Run - Technical Documentation

## Overview

Automated RPA solution for processing and managing Milk Run logistics operations between DHL and STELLANTIS. The system automates demand downloads, spreadsheet processing, and data consolidation for supplier management across multiple routes (SP, SUL, CKD, FIASA).

## Tech Stack

### Core Technologies
- **Language**: Python 3.x
- **UI Framework**: Tkinter (multi-threaded with queue-based updates)
- **Browser Automation**: Playwright (Chromium-based)
- **Excel Processing**: xlwings, OpenPyXL, pyxlsb
- **Data Processing**: Pandas, NumPy
- **Build Tool**: PyInstaller

### Dependencies
```python
playwright>=1.40.0
pandas>=2.0.0
openpyxl>=3.1.0
pyxlsb>=1.0.10
xlwings>=0.30.0
```

## Architecture

### Core Components

1.  **App.py** - Main application with GUI and orchestration logic
2.  **Tasks.py** - Business logic for demand processing and file operations
3.  **config.json** - Application configuration and selectors
4.  **credencial.json** - Authentication credentials (not in version control)

### Key Features

-   Multi-threaded GUI application with real-time progress tracking
-   Automated web scraping for demand data downloads
-   Excel file processing and data consolidation
-   Supplier database management
-   Custom styling with DHL/STELLANTIS branding

## Project Structure

```
RPA Milk Run/├── App.py                      # Main application entry point├── Tasks.py                    # Business logic module├── config.json                 # Configuration settings├── credencial.json            # Credentials (gitignored)├── 1 - MATRIZ/                # Master PFEP files├── Bases/                     # Database files│   └── Forncedores_Responsavel.json├── Demanda/                   # Downloaded demand files├── Planilhas_Recebidos2/      # Received spreadsheets└── Resultados/                # Output files
```

## Configuration

### config.json Structure

-   **playwright_selectors**: Web element selectors for automation
-   **paths**: Directory and file path configurations
-   **business_logic**: Business rules and mappings
    -   Style configurations
    -   SAP exclusion lists
    -   Carrier mappings
    -   Filter criteria

### Key Functions

#### App.py

-   `load_credentials()`: Loads authentication from credencial.json
-   `get_playwright_browser_path()`: Resolves Chromium path for bundled/dev environments
-   `run_automation()`: Main automation workflow
-   `update_gui()`: Queue-based GUI updates for thread safety
-   `App class`: Tkinter application with DHL/STELLANTIS themed interface

#### Tasks.py

-   `download_Demanda()`: Automated demand file download
-   `Processar_Demandas()`: Process and consolidate demand files
-   `Copiar_planejamentos_para_cargolift_Arquivos()`: Copy planning data to Cargolift files

## Installation & Setup

### Development Environment

```powershell
# Install dependenciespip install playwright pandas openpyxl pyxlsb# Install Playwright browsersplaywright install chromium
```

### Building Executable

To generate the executable, run:

```powershell
pyinstaller --noconfirm --onefile --windowed --noconsole --name "RPA Milk Run" --icon "C:/Users/perna/Desktop/STALLANTIS/RPA Pagamentos/MilRun.ico" --add-data "C:UserspernaAppDataLocalms-playwrightchromium-1187chrome-win;ms-playwrightchromium-1187chrome-win" App.py
```

This will create a standalone executable in the `dist/` folder.

## Usage

1.  **Setup Credentials**: Create `credencial.json` with login credentials
2.  **Configure Paths**: Adjust paths in `config.json` if needed
3.  **Run Application**: Execute `App.py` or the built executable
4.  **Click "Processar"**: Starts the automated workflow
5.  **Monitor Progress**: Real-time logs and progress bar

## Data Flow

1.  **Download Phase**: Fetch demand files from ELOG system
2.  **Processing Phase**: Parse and consolidate demand data
3.  **Mapping Phase**: Apply supplier and carrier mappings
4.  **Output Phase**: Generate consolidated Excel reports

## Security Notes

-   Credentials stored in `credencial.json` (excluded from version control)
-   Automated browser runs in non-headless mode for visibility
-   No hardcoded passwords in codebase

## Maintenance

### Adding New Suppliers

Update `config.json` → `business_logic.carrier_mappings`

### Modifying Selectors

Update `config.json` → `playwright_selectors` if web UI changes

### Updating File Patterns

Modify `config.json` → `paths.dynamic_files` for new file naming conventions

## Developer

**Vincent Pernarh** - Developer

## License

Internal use - DHL ↔ STELLANTIS Operations