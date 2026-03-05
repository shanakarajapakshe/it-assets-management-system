# IT Asset Manager Pro v2.0

Fully offline Electron desktop app. No database. Data stored in Excel files.

## Quick Start

```bash
cd it-asset-v2

npm start
```

## All Features Implemented

1. Dashboard - 8 KPI cards, dept bar chart, status donut, monthly purchases chart, warranty alerts
2. Employee PC Management - Full CRUD, auto Asset IDs (PC-0001), return tracking, search/filter/sort/pagination
3. Purchasing Management - Full CRUD, auto Item IDs (IT-0001), assign/replace wizard
4. Excel Data Handling - Atomic writes (tmp→rename), write lock, styled output, auto-filter, frozen headers
5. Asset Status Logic - Replacement links old PC → Replaced, IT item → Assigned
6. Reports - Export Excel or PDF, filter by dept/status/date, 4 quick presets
7. Search + Pagination - 12 per page, smart pager, sort by any column
8. Asset ID auto-generation - PC-0001...PC-9999 and IT-0001...IT-9999

## Data Location
- Dev: ./data/
- Production: %APPDATA%/it-asset-manager-pro/data/

## Backup System
Auto-backup before every write. Max 10 per file. View in Backups page.
