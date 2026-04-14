# CityCyber POS — Local Setup Guide

## Project Structure

```
citycyber_pos/
 ├── app.py               ← Flask backend (the only Excel writer)
 ├── requirements.txt     ← Python dependencies
 ├── data.xlsx            ← Your CityCyber workbook (copy here)
 ├── templates/
 │    └── index.html      ← POS dashboard (runs in browser)
 ├── static/              ← (empty; for future CSS/JS files)
 └── logs/
      └── pos.log         ← Auto-created on first run
```

---

## 1. Prerequisites

- Python 3.9 or newer
- pip (bundled with Python)
- Your `CityCyber_POS_FIXED.xlsx` file renamed to `data.xlsx` and placed in this folder

---

## 2. Install Dependencies

Open a terminal in the `citycyber_pos/` folder and run:

```bash
pip install -r requirements.txt
```

Or with pip3:

```bash
pip3 install -r requirements.txt
```

---

## 3. Start the Server

```bash
python app.py
```

You should see:

```
INFO  CityCyber POS starting — Excel: /path/to/citycyber_pos/data.xlsx
 * Running on http://0.0.0.0:5000
```

---

## 4. Open the Dashboard

Open your browser and go to:

```
http://localhost:5000
```

The POS interface loads instantly. Use it from any device on your local network:

```
http://<your-machine-ip>:5000
```

---

## 5. How to Use

1. **Filter** by category (Print, Scan, Photo, etc.) using the pills
2. **Select** a service from the dropdown **or** tap a tile in the Quick Select grid
3. **Set** the quantity (default: 1)
4. **Pick** payment mode: Cash / UPI / Card
5. **Tap** "Log Transaction" — the row is written to `TRANSACTION_LOG` in Excel immediately
6. The sidebar shows today's revenue, profit and count — auto-refreshed every 60 seconds

---

## 6. Excel Sheet Details

### `MASTER` sheet (rows 4–46)
| Column | Content       |
|--------|---------------|
| B      | SERVICE NAME  |
| C      | CATEGORY      |
| D      | SELL RATE ₹   |
| E      | COST ₹        |
| F      | MARGIN ₹      |
| G      | MARGIN %      |
| H      | UNITS SOLD    |
| I      | PRIORITY      |

### `TRANSACTION_LOG` sheet (rows 4+)
| Column | Content        | Set by  |
|--------|----------------|---------|
| A      | TIMESTAMP      | Python  |
| B      | SERVICE        | Python  |
| C      | QTY            | Python  |
| D      | PAYMENT MODE   | Python  |
| E      | REVENUE ₹      | Formula |
| F      | PROFIT ₹       | Formula |
| G      | COST ₹         | Formula |
| H      | NOTES          | Python  |

**Formulas written per row (Excel recalculates live):**

```excel
E  =IFERROR(C4*XLOOKUP(B4,MASTER!$B$4:$B$46,MASTER!$D$4:$D$46,0), <computed_fallback>)
F  =IFERROR(C4*(XLOOKUP(B4,...,Price,...)-XLOOKUP(B4,...,Cost,...)), <fallback>)
G  =IFERROR(C4*XLOOKUP(B4,...,Cost,...), <fallback>)
```

If Excel can't evaluate XLOOKUP (e.g. on recalculation), the hardcoded fallback value is used.

---

## 7. Hardening Notes

| Risk                  | Mitigation                                               |
|-----------------------|----------------------------------------------------------|
| Concurrent writes     | `fcntl.LOCK_EX` file lock — one writer at a time        |
| Missing Excel file    | Clear 500 error with message pointing to correct path   |
| Unknown service name  | Rejected with 400 before any write happens              |
| Invalid quantity      | Rejected with 400 (must be integer ≥ 1)                |
| Invalid payment mode  | Rejected with 400 (only Cash / UPI / Card)             |
| Formula errors        | Each cell has `IFERROR(…, hardcoded_fallback)`         |
| Server crash          | All exceptions logged to `logs/pos.log`                |

---

## 8. Updating Prices

Open `data.xlsx` → go to `MASTER` sheet → change any **SELL RATE ₹** or **COST ₹** value → save.  
The POS backend reads the file fresh on every `/get-services` call, so prices update within seconds (press **Refresh** in the sidebar or reload the page).

---

## 9. Stopping the Server

Press `Ctrl+C` in the terminal.

---

## 10. Logs

All transactions and errors are appended to `logs/pos.log`. Example entry:

```
2026-04-02 10:35:22 [INFO] add-transaction: service=Print - A4 B&W qty=10 payment=Cash rev=50.00 profit=40.00 row=5
```
