# Metro Connect Services – Monthly Performance Monitoring Dashboard

A local website with three interconnected pages: **Input Page**, **Consolidation Sheet**, and **Visualization & Filters**. The original Metro Connect Services workbook is included as the persistent data source. The application adds a **Data Management** sheet for normalized daily records and keeps the original monthly sheets in the same workbook.

## Run on Windows
1. Install Python 3.10 or newer from the official Python website, ensuring Python is added to PATH.
2. Extract the ZIP to a folder where you have write permission.
3. Double-click `run.bat`. On first run it installs Flask and openpyxl, then opens the local server.
4. In Chrome, open **http://127.0.0.1:5000**.
5. Keep the command window open while using the dashboard. Close it to stop the local server.

## Features
- Entry form grouped by service/domain and metric, with persistent server-side enable/disable settings for every domain. Disabled domains cannot receive new entries; their historical records remain visible in the Consolidation Sheet and Visualization. Includes date selection, validation, duplicate protection, and explicit edit mode.
- Add New Domain control on the Input Page. New domains are saved to `data/domains.json`, appear in the entry form, Consolidation Sheet domain filters, and Visualization filters. New domains start with a Daily metric and can begin accumulating history after data entry.
- Central editable records table with search, filters, edit, delete confirmation, and recalculated monthly cumulative values.
- Dashboard KPIs, daily and cumulative trends, service comparison, monthly comparison, parking trends, and revenue/ridership charts with separate y-axes.
- Filters for a specific date, month, year, from/to date range, multi-select service/domain, and performance metric. Charts use responsive sizing, readable labels, interactive tooltips, and filtered data previews.
- Download the current workbook using **Export Excel**.
- Original monthly tabs are retained. Imported blank cells remain blank; numeric zero remains a real value. Cumulative values are recalculated separately for each month, service, and metric.

## Data storage and backup
The workbook at `data/MetroConnectData.xlsx` is the application's persistent store. Back it up regularly. The app is designed for a single local user/session and should not be used as a shared multi-user server without migrating storage to a database and adding authentication/locking.

## Notes about source fields
The monthly sheets include service daily counts, MTC ridership/revenue/GAP, and Pallavaram/Thirumangalam parking vehicle count/revenue. Additional analysis tabs are preserved in the workbook and are not rewritten by the app. Historical source values are imported from the daily rows; cumulative rows are recalculated from daily values rather than treated as separate input records.

## Troubleshooting
- If `py` is not recognized, install Python and enable “Add Python to PATH”, then reopen the terminal.
- If the dashboard cannot save, close the workbook in Excel and retry; Excel may lock the file while it is open.
- Chart.js and Google Fonts are loaded from public CDNs, so charts and remote fonts require an internet connection. The data entry and Excel storage are local.
