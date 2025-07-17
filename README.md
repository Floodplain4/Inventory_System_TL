🧾 Lenovo Parts Inventory System (v1.4)
A lightweight desktop inventory tracking tool built to manage Lenovo parts rotation at my job. This program was developed as a personal project to streamline how I handle repairs and hardware status updates.

✨ Features
Add and manage entries using:

Work Order

Serial Number

Status (e.g., Ordered, Pending, Replaced, Complete)

Notes

Update status quickly via dropdown menu and update button

Automatically saves changes to lcd_log.csv

Manual backup/export supported via Export button

Import existing logs (e.g., lcd_log.csv or Example.csv)

Includes a sample dataset (Example.csv) for demonstration

⚠️ Note: The Refresh Log button is deprecated. It still functions if you manually modify the CSV file and want to refresh the UI.

🛠️ How to Use
Download the latest Release 1.4 (link your .zip or .exe).

Extract all files from the .zip.

Run LCD1.4.exe (you may need to click "More Info" > "Run Anyway" due to unsigned code).

On first run:

lcd_log.csv should open automatically.

If not, use the Import button to select either lcd_log.csv or Example.csv.

📁 Files Included
LCD1.4.exe – Main executable

lcd_log.csv – Current working log (auto-saves here)

Example.csv – Sample data for testing

inventory_ui.py – Source code (optional if published)

README.md – This file

💡 Notes
This was built as a personal tool — think of it as a "fancy sheet of paper" with buttons.

All example data is fictional. Serial numbers and work orders are randomly generated for demonstration.

No sensitive or proprietary data is included in this repo.

🧑‍💻 About Me

I'm a helpdesk tech learning programming on the side. This project was my attempt to make my day-to-day work more efficient and learn Python/Tkinter in the process.

License: MIT License
