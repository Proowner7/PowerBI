# PowerBI Training Project

A Power BI training project including a report, template, data sources, and a mid-session snapshot.

## Project Structure

```
PowerBI Training.pbix                        # Main Power BI report
Training-Basis-1-datafiles-1/
├── 1. Template/
│   └── GR Template_PBI.pbit                 # Power BI template file
├── 2. Data/
│   ├── Customers.csv                        # Customer data
│   ├── Orders.csv                           # Orders data
│   ├── Products.csv                         # Products data
│   ├── Paymentmethod.csv                    # Payment method data
│   └── DeliveryComp.xlsx                    # Delivery company data
└── 3. Snapshot/
    └── Snapshot after lunch_PBI.pbix        # Mid-session report snapshot
```

## Getting Started

1. Install [Power BI Desktop](https://powerbi.microsoft.com/desktop/) (Windows).
2. Open **`PowerBI Training.pbix`** in Power BI Desktop.
3. If prompted for data sources, point them to the files in the `2. Data/` folder.
4. To start from the template instead, open **`GR Template_PBI.pbit`** and load the data files when prompted.

## Data Sources

| File | Description |
|------|-------------|
| `Customers.csv` | Customer records (700 customers) |
| `Orders.csv` | Order transactions (~68,652 rows, Jan 2025 – Aug 2026) |
| `Products.csv` | Product catalogue (144 products across 5 product lines) |
| `Paymentmethod.csv` | Payment method reference (7 methods) |
| `DeliveryComp.xlsx` | Delivery company reference |

## Data Enrichment — Steps Taken

The following columns were added to the data files to make them more insightful for Power BI analysis.

### Orders.csv — new columns

| Column | How it is derived | Why it is useful |
|--------|-------------------|------------------|
| `ProductID` | Numeric part extracted from the `Product` field (e.g. `"86-Kodiak"` → `86`) | Enables a clean join to `Products.csv` on `Product ID` |
| `ProductName` | Text part extracted from the `Product` field (e.g. `"86-Kodiak"` → `"Kodiak"`) | Provides a clean product name without the ID prefix |
| `Year` | Year portion of `Salesdate` (format D-M-YYYY) | Enables year-over-year comparisons |
| `Month` | Zero-padded month number (e.g. `01`, `02`) | Enables month-level filtering and sorting |
| `MonthName` | Full month name (e.g. `January`) | Human-readable label for charts and slicers |
| `Quarter` | Quarter label derived from month (e.g. `Q1`, `Q2`) | Enables quarterly trend analysis |
| `SalesCategory` | Tertile split on `Sales` — **Low** (≤ 4 708), **Medium** (4 709 – 10 074), **High** (> 10 074) | Allows quick segmentation of order value without losing numeric detail |

> **Note:** `Sales` equals `Price × Amount` in every row — no discrepancies were found.

### Products.csv — new column

| Column | How it is derived | Why it is useful |
|--------|-------------------|------------------|
| `PriceCategory` | **Budget** (≤ 80), **Mid-range** (81 – 160), **Premium** (> 160) based on `Price` | Enables price-tier analysis alongside product line breakdowns |

> Price thresholds were chosen based on the natural distribution of the 144 product prices (range 52 – 260, median 158).
