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

1.  
### System Design

```
┌─────────────────────────────────────────────────────────────┐
│                    GUI Layer (Tkinter)                      │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐      │
│  │  Progress    │  │  Status Log  │  │  Controls    │      │
│  └──────────────┘  └──────────────┘  └──────────────┘      │
└─────────────────────────────────────────────────────────────┘
                            ↓ Queue Communication
┌─────────────────────────────────────────────────────────────┐
│              Worker Thread (Automation Core)                │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  Playwright Browser (Chromium)                        │   │
│  │  • Login to ELOG system                              │   │
│  │  • Navigate to reports                                │   │
│  │  • Download demand files (.txt, .xlsx)               │   │
│  └──────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│              Data Processing Pipeline                        │
│  ┌──────────────────────────────────────────────────────┐   │
│  │ 1. Parse demand files (TXT/Excel)                     │   │
│  │ 2. Normalize data (SAP codes, PNs)                    │   │
│  │ 3. Apply business rules & filters                     │   │
│  │ 4. Cross-reference with supplier DB                   │   │
│  └──────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│            Excel Automation Layer (xlwings)                 │
│  ┌──────────────────────────────────────────────────────┐   │
│  │ • Update PFEP master files                            │   │
│  │ • Update Programação FIASA                            │   │
│  │ • Update Cargolift SP/SUL programs                    │   │
│  │ • Update CKD/FPT/MOPAR programs                       │   │
│  │ • Execute VBA macros (recalculation)                  │   │
│  └──────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
```

### Core Components

#### 1. **App.py** - Main Application Controller
- Multi-threaded Tkinter GUI
- Queue-based communication (thread-safe)
- Chromium path resolution (bundled vs. system)
- Credential management
- Visual progress tracking

#### 2. **Tasks.py** - Business Logic Engine (2382 lines)
**Primary Functions:**
- `download_Demanda()` - Web scraping & file downloads
- `Processar_Demandas()` - Data parsing & consolidation
- `Atualiza_PFEP()` - PFEP master file updates
- `Processar_programacao()` - FIASA program processing
- `progrma_cargolift()` - Cargolift SP automation
- `Corregir_peso_e_valor()` - Weight & value corrections
- `Copiar_planejamentos_para_cargolift_Arquivos()` - Data distribution
- `Copiar_e_Colar_Programacao_Sul()` - SUL route processing

#### 3. **config.json** - Configuration Hub
- Playwright selectors (DOM elements)
- File path mappings (dynamic file discovery)
- Business logic rules (SAP exclusions, carrier mappings)
- Sheet name references

#### 4. **credencial.json** - Secure Credentials
```json
{
  "url_Elog": "https://...",
  "user": "username",
  "password": "password"
}
```

### Key Features

- ✅ **Multi-threaded Architecture**: Non-blocking GUI with worker threads
- ✅ **Automated Web Scraping**: Date-based demand file downloads with retry logic
- ✅ **Intelligent File Processing**: Supports .txt, .csv, .xls, .xlsx, .xlsm, .xlsb
- ✅ **Dynamic File Discovery**: Config-based file pattern matching
- ✅ **Data Normalization**: SAP code & PN standardization
- ✅ **Excel Automation**: VBA macro execution, formula injection
- ✅ **Multi-Route Support**: SP, SUL, CKD, FIASA, FPT, MOPAR
- ✅ **Error Handling**: Comprehensive try-catch with logging
- ✅ **Custom Styling**: DHL/STELLANTIS branded interface

## Project Structure

```
RPA Milk Run/
├── App.py                             # Main application entry point
├── Tasks.py                           # Business logic module (2382 lines)
├── config.json                        # Configuration settings
├── credencial.json                    # Credentials (excluded from VCS)
├── Modelos.json                       # Template definitions (optional)
├── RPA Milk Run.spec                  # PyInstaller build spec
├── README.md                          # This file
│
├── 1 - MATRIZ/                        # Master planning files
│   ├── PFEP 2024 DHL FIASA OP_W37-REV 6.xlsm
│   ├── Programação FIASA - OFICIAL.xlsm
│   ├── Cargolift SP - PFEP.xlsm
│   ├── Cargolift SP - Suppliers.xlsm
│   └── [Dynamic files matched by config]
│
├── Bases/                             # Reference databases
│   ├── DB Fornecedores.xlsx           # Supplier master data
│   └── Forncedores_Responsavel.json   # Supplier responsibility mapping
│
├── Demanda/                           # Downloaded demand files
│   ├── 1_GMDNVA8.TXT                  # Auto-downloaded from ELOG
│   ├── 2_GMDNVA8.TXT
│   └── [Additional demand files]
│
├── Planilhas_Recebidos2/              # Received planning sheets
│   ├── Cargolift SP - Embalagem.xlsb
│   ├── FPT_Cálculo_cargolift.xlsm
│   ├── PROGRAMAÇÃO SUL.xlsm
│   ├── PROGRAMAÇÃO CKD.xlsm
│   └── [Route-specific files]
│
├── Resultados/                        # Output files
│   └── Demandas_Total.xlsx            # Consolidated demand report
│
├── build/                             # PyInstaller build artifacts
│   └── RPA Milk Run/
│       ├── Analysis-00.toc
│       ├── EXE-00.toc
│       └── [Build metadata]
│
└── dist/                              # Distribution folder
    └── RPA Milk Run.exe               # Compiled executable
```

## Installation & Setup

### Development Environment

```powershell
# Install dependencies
pip install playwright pandas openpyxl pyxlsb xlwings

# Install Playwright browsers
playwright install chromium
```

### Building Executable

To generate the executable, run:

```powershell
pyinstaller --noconfirm --onefile --windowed --noconsole --name "RPA Milk Run" --icon "C:/Users/perna/Desktop/STALLANTIS/RPA Pagamentos/MilRun.ico" --add-data "C:\Users\perna\AppData\Local\ms-playwright\chromium-1187\chrome-win;ms-playwright\chromium-1187\chrome-win" App.py
```

This will create a standalone executable in the `dist/` folder.

## Usage

### Prerequisites

1. **Create credencial.json** in the project root:
   ```json
   {
     "url_Elog": "https://your-elog-system.com",
     "user": "your_username",
     "password": "your_password"
   }
   ```

2. **Verify folder structure** matches config.json
3. **Ensure Excel files** are in correct locations

### Running the Application

#### Option 1: Development Mode
```powershell
python App.py
```

#### Option 2: Executable Mode
```powershell
.\dist\"RPA Milk Run.exe"
```

### Automated Workflow Steps

```
Progress  | Step
----------|--------------------------------------------------
5%        | Loading credentials
10%       | Launching browser
20%       | Downloading demand files from ELOG
30%       | Processing demand data
50%       | Updating PFEP master file
60%       | Processing FIASA programming
70%       | Updating Cargolift SP
80%       | Correcting weights and values
90%       | Distributing to all routes (SUL, CKD, FPT)
100%      | Process complete!
```

**Expected Runtime:** 6-12 minutes (varies by data volume)

## Configuration

### config.json Structure

```json
{
  "playwright_selectors": {
    "login_user_textbox": "User",
    "login_button": "Log In",
    "menu_item_text": "ELOG - Importar A8 Automatica"
  },
  "paths": {
    "folders": {
      "base_demanda": "Demanda",
      "base_matriz": "1 - MATRIZ"
    },
    "dynamic_files": {
      "pfep_search_terms": ["PFEP 2024 DHL", "PFEP 2025 DHL"]
    }
  },
  "business_logic": {
    "sap_exclusion_list": [800030982],
    "carrier_mappings": {
      "suppliers_carrier": {
        "800006524": "CARGOLIFT"
      }
    }
  }
}
```

## API Reference

### App.py Functions

#### `load_credentials() -> dict`
Loads authentication credentials from `credencial.json`.

#### `get_playwright_browser_path() -> str`
Resolves Chromium executable path for bundled (.exe) or development environments.

#### `run_automation(playwright, q) -> None`
Main automation orchestrator executing the full workflow.

### Tasks.py Functions

#### `download_Demanda(page, url_order, q, username, password) -> None`
Automated web scraping for demand file downloads with date-based retry logic.

#### `Processar_Demandas(q) -> None`
Parses and consolidates demand files supporting .txt, .csv, and Excel formats.

#### `Atualiza_PFEP(path_demandas, q) -> None`
Updates master PFEP file with SUMIFS formulas and executes VBA macros.

#### `Corregir_peso_e_valor(q, wb, demandas_path, pfep_source) -> None`
Corrects M3 and Kg values for missing PNs based on JSON mappings.

## Data Flow

### Complete Workflow

1. **Download** → Download demand files from ELOG (date-based search)
2. **Parse** → Process .txt and Excel files, normalize data
3. **Filter** → Apply business rules (state filters, SAP exclusions)
4. **Merge** → Cross-reference with supplier database
5. **Update PFEP** → Inject formulas and execute macros
6. **Update FIASA** → Filter and copy data to programming sheets
7. **Update Cargolift** → Distribute to SP/SUL routes
8. **Correct Values** → Add missing PN weight/volume data
9. **Distribute** → Copy to all route programs (CKD, FPT, MOPAR)

## Troubleshooting

### Common Issues

#### "Chromium executable not found"
**Solution:** Verify PyInstaller `--add-data` path or reinstall Playwright browsers.

#### "credencial.json not found"
**Solution:** Ensure file is in same folder as App.py/executable.

#### Login fails / Timeout errors
**Solution:** Update `playwright_selectors` in config.json if ELOG UI changed.

#### Excel automation errors
**Solutions:**
- Close all Excel instances before running
- Enable macros in Excel Trust Center
- Run as Administrator if permission denied

### Debug Mode

Enable logging in `Tasks.py`:
```python
import logging
logging.basicConfig(level=logging.DEBUG, filename='rpa_debug.log')
```

## Maintenance

### Adding New Suppliers

1. Update `config.json` → `carrier_mappings`
2. Update `Bases/DB Fornecedores.xlsx`
3. Update `Bases/Forncedores_Responsavel.json`

### Modifying Web Selectors

If ELOG website changes, update `config.json` → `playwright_selectors`

### Updating File Patterns

Add new search terms to `config.json` → `dynamic_files`

## Security Notes

- Credentials stored in `credencial.json` (excluded from version control)
- Automated browser runs in non-headless mode for visibility
- No hardcoded passwords in codebase

## Developer

**Vincent Pernarh** - RPA Developer

## License

Internal use - DHL ↔ STELLANTIS Operations
