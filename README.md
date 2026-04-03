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
| `Customers.csv` | Customer records |
| `Orders.csv` | Order transactions |
| `Products.csv` | Product catalogue |
| `Paymentmethod.csv` | Payment method reference |
| `DeliveryComp.xlsx` | Delivery company reference |
