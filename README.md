# 🖥️ IT Assets Management System

A cross-platform desktop application for managing IT assets — including laptop inventories, employee PC assignments, and IT procurement tracking. Built with **Electron.js** and **ExcelJS**, it provides a clean GUI for managing and updating Excel-based asset records without needing spreadsheet expertise.

---

## 📋 Table of Contents

- [Features](#-features)
- [Tech Stack](#-tech-stack)
- [Project Structure](#-project-structure)
- [Prerequisites](#-prerequisites)
- [Installation](#-installation)
- [Usage](#-usage)
- [Data Files](#-data-files)
- [License](#-license)

---

## ✨ Features

- 📦 **Laptop Inventory Management** — Track and update the organization's laptop inventory
- 👤 **Employee PC Assignment Tracking** — View and manage which PCs are assigned to which employees
- 🛒 **IT Procurement Tracking** — Log and monitor newly purchased or replacement IT items
- 📊 **Excel Integration** — Reads and writes `.xlsx` files directly via ExcelJS
- 🖥️ **Cross-Platform** — Runs on Windows, macOS, and Linux via Electron
- 🔒 **Secure IPC** — Uses Electron's context-isolated preload scripts for safe renderer-to-main communication

---

## 🛠️ Tech Stack

| Layer | Technology |
|---|---|
| Desktop Framework | [Electron.js](https://www.electronjs.org/) |
| Excel Processing | [ExcelJS](https://github.com/exceljs/exceljs) |
| Frontend | HTML, CSS, Vanilla JavaScript |
| Runtime | Node.js |

---

## 📁 Project Structure

```
it-assets-management-system/
├── assets/
│   └── icons/              # App icons (.png, .ico, .icns)
├── data/
│   ├── CSEC_Laptop_Inventory.xlsx
│   ├── Employee_PCs_Using_Details.xlsx
│   ├── Newly_Purchasing_IT_Items_Replacing.xlsx
│   └── backups/            # Auto-generated backups
├── src/
│   ├── main/
│   │   ├── main.js         # Electron main process
│   │   └── excelManager.js # Excel read/write logic
│   ├── preload/
│   │   └── preload.js      # Context bridge (IPC exposure)
│   └── renderer/
│       ├── index.html      # Main UI
│       ├── app.js          # Renderer process logic
│       └── styles/
│           └── main.css    # Application styles
├── package.json
└── README.md
```

---

## ✅ Prerequisites

- [Node.js](https://nodejs.org/) v16 or higher
- npm (comes with Node.js)

---

## 🚀 Installation

1. **Clone the repository**

   ```bash
   git clone https://github.com/your-username/it-assets-management-system.git
   cd it-assets-management-system
   ```

2. **Install dependencies**

   ```bash
   npm install
   ```

3. **Run the application**

   ```bash
   npm start
   ```

---

## 🏗️ Build for Distribution

To package the app as a standalone executable:

```bash
npm run build
```

> Builds are output to the `dist/` folder. Electron Builder is used to generate platform-specific installers.

---

## 📂 Usage

On launch, the application loads data from the Excel files located in the `data/` directory. From the UI you can:

1. **View asset records** across all three datasets
2. **Add, edit, or remove** asset entries
3. **Save changes** back to the Excel files

> ⚠️ Do not move or rename the Excel files in the `data/` folder, as the application references them by path.

---

## 🗂️ Data Files

| File | Description |
|---|---|
| `CSEC_Laptop_Inventory.xlsx` | Full inventory of laptops owned by the organization |
| `Employee_PCs_Using_Details.xlsx` | Records of which employees are using which PCs |
| `Newly_Purchasing_IT_Items_Replacing.xlsx` | Log of new IT purchases and replacement items |

Backup copies of these files are automatically saved to `data/backups/` on each write operation.

---

## 📄 License

This project is licensed under the terms found in the [LICENSE](./LICENSE) file.

---

> Built for internal IT asset tracking. Contributions and issues welcome.
