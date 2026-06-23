# IT Asset Manager Pro

Cross-platform desktop app for managing office IT assets, employee device assignments, and procurement records in one place. Built with Electron.js and ExcelJS, it gives non-technical users a simple interface on top of Excel-based records.

Live docs: https://shanakarajapakshe.github.io/it-assets-management-system/

## What It Does

- Tracks laptop inventory across the organization
- Manages employee PC assignment records
- Logs newly purchased and replacement IT items
- Reads and writes `.xlsx` files directly
- Keeps backup copies of data after updates
- Runs as a desktop app on Windows, macOS, and Linux

## Why It Is Useful

- Reduces manual spreadsheet editing
- Gives office staff a simpler UI for asset updates
- Helps keep inventory, allocation, and procurement data in sync
- Fits internal IT operations and support workflows

## Tech Stack

- Electron.js
- Node.js
- ExcelJS
- HTML, CSS, Vanilla JavaScript

## Project Structure

```text
it-assets-management-system/
├── assets/
│   └── icons/
├── data/
│   ├── CSEC_Laptop_Inventory.xlsx
│   ├── Employee_PCs_Using_Details.xlsx
│   ├── Newly_Purchasing_IT_Items_Replacing.xlsx
│   └── backups/
├── src/
│   ├── main/
│   │   ├── main.js
│   │   └── excelManager.js
│   ├── preload/
│   │   └── preload.js
│   └── renderer/
│       ├── index.html
│       ├── app.js
│       └── styles/main.css
├── docs/
├── package.json
└── README.md
```

## Run Locally

```bash
git clone https://github.com/shanakarajapakshe/it-assets-management-system.git
cd it-assets-management-system
npm install
npm start
```

## Build

```bash
npm run build
```

Windows build only:

```bash
npm run build:win
```

## Data Files

The app uses the Excel files inside the `data/` folder:

- `CSEC_Laptop_Inventory.xlsx`
- `Employee_PCs_Using_Details.xlsx`
- `Newly_Purchasing_IT_Items_Replacing.xlsx`

Backups are written to `data/backups/`.

## License

See [LICENSE](./LICENSE).
