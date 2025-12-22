# Stock Reconciliation Desktop Application

A modern Electron-based desktop application for stock movement reconciliation across multiple production units.

## Features

- 🏭 **Multi-Unit Support**: Fath1, Fath2, Fath5, Larbaa, Oran, Fibre
- 📁 **Drag & Drop**: Easy file upload with automatic keyword matching
- 🔗 **Smart Matching**: Automatic file-to-atelier linking based on filename keywords
- 📊 **Results Preview**: Interactive table view of matches and discrepancies
- 📤 **CSV Export**: Export results for further analysis

## Installation

### Prerequisites

1. **Node.js** (v18 or higher)
   - Download from: https://nodejs.org/

2. **Python** (3.8 or higher)
   - Download from: https://www.python.org/
   - Make sure Python is added to PATH

3. **Python Dependencies**
   ```bash
   pip install pandas openpyxl numpy
   ```

### Setup

1. Navigate to the StockApp folder:
   ```bash
   cd StockApp
   ```

2. Install Node.js dependencies:
   ```bash
   npm install
   ```

3. Run the application:
   ```bash
   npm start
   ```

## Usage

### Step 1: Select Unit
Choose the production unit you want to process (Fath1, Fath2, Fath5, Larbaa, Oran, or Fibre).

### Step 2: Upload Files
- **Drag & Drop** your Excel files into the drop zone
- The app will automatically:
  - Identify the **Stock file** (contains "STOCK" in filename)
  - Match **Movement files** to ateliers based on keywords

### Step 3: Review Matching
- View matched files in the Ateliers list
- Manually upload files for unmatched ateliers
- Skip ateliers you don't want to process

### Step 4: Process
Click "Process Files" to run the reconciliation calculations.

### Step 5: View Results
- Select an atelier to view its results
- Toggle between Matches and Discrepancies
- Export results as CSV

## Project Structure

```
StockApp/
├── main/
│   ├── main.js          # Electron main process
│   └── preload.js       # Preload script for IPC
├── renderer/
│   ├── index.html       # Main UI
│   ├── styles/
│   │   └── main.css     # Styling (blue & white theme)
│   └── scripts/
│       └── app.js       # Frontend logic
├── backend/
│   ├── processor.py     # Python processing engine
│   └── README.md        # Backend documentation
├── package.json
└── README.md
```

## Color Theme

The application uses a professional blue and white color scheme:
- Primary: `#2563eb` (Blue)
- Background: `#f8fafc` (Light Gray)
- Cards: `#ffffff` (White)

## Development

Run in development mode with DevTools:
```bash
npm run dev
```

## Troubleshooting

### Python not found
Make sure Python is installed and added to your system PATH.

### Excel file errors
Ensure your Excel files have the expected column headers (REFERENCE, QUANTITE, LOCALISATION, etc.)

### Missing dependencies
```bash
pip install pandas openpyxl numpy
npm install
```
