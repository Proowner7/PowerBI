# Power Query Setup Guide

Complete these steps in Power BI Desktop **before** working on the report
pages. Open **Power Query Editor** (Home → Transform data).

---

## 1. Load the five source files

| Query name | Source file | Notes |
|---|---|---|
| `Orders` | `2. Data/Orders.csv` | CSV, comma-delimited |
| `Customers` | `2. Data/Customers.csv` | CSV, comma-delimited |
| `Products` | `2. Data/Products.csv` | CSV, comma-delimited |
| `Paymentmethod` | `2. Data/Paymentmethod.csv` | CSV, comma-delimited |
| `Deliverycomp` | `2. Data/DeliveryComp.xlsx` | Excel – skip the title row (see step 3) |

Use **Home → New Source → Text/CSV** or **Excel Workbook** respectively.

---

## 2. Orders table – required transformations

### 2a. Parse Salesdate to a proper Date type

The raw column arrives as text in Dutch `D-M-YYYY` format.

```m
// In the Applied Steps for Orders, add a "Changed Type" step after loading:
#"Parse Salesdate" =
    Table.TransformColumnTypes(
        #"Previous Step",
        {{"Salesdate", type date}},
        "nl-NL"          // locale that matches the DD-MM-YYYY input
    )
```

Alternatively, use the UI: right-click the **Salesdate** column header →
*Change Type → Using Locale…* → Data type: **Date**, Locale: **Dutch (Netherlands)**.

### 2b. Extract ProductID from the Product column

The `Product` column contains values like `"86-Kodiak"`. Extract the
numeric prefix as an integer so the Orders table can join to Products.

```m
#"Added ProductID" =
    Table.AddColumn(
        #"Parse Salesdate",
        "ProductID",
        each Int64.From(Text.BeforeDelimiter([Product], "-")),
        Int64.Type
    )
```

Use the UI: **Add Column → Column From Examples** and type `86` for the
first row, or use **Add Column → Extract → Text Before Delimiter** with
delimiter `-`, then change the type to **Whole Number**.

### 2c. Flag rows with a missing Customer ID

```m
#"Added HasCustomer" =
    Table.AddColumn(
        #"Added ProductID",
        "HasCustomer",
        each [Customer ID] <> null,
        type logical
    )
```

Decide whether to keep or remove NULL-Customer rows depending on the
analysis requirements.

### 2d. Verify/add Sales column

If `Sales` is already present and equals `Price × Amount`, no action is
needed. If it is missing, add it:

```m
#"Added Sales" =
    Table.AddColumn(
        #"Previous Step",
        "Sales",
        each [Price] * [Amount],
        type number
    )
```

---

## 3. DeliveryComp.xlsx – skip the title row

Excel exports sometimes include a merged title row above the header.

1. In Power Query Editor, expand the **DeliveryComp** query.
2. In the Applied Steps pane, locate the **Promoted Headers** step.
3. Ensure the first data row (after the header) begins with `DC_ID`.
   If a title row still appears, add a step:

```m
#"Removed Top Rows" =
    Table.Skip(#"Previous Step", 1)
```

Then re-promote headers:

```m
#"Promoted Headers" =
    Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true])
```

---

## 4. Customers table – data type checks

| Column | Required type |
|---|---|
| `Latitude` | Decimal Number |
| `Longitude` | Decimal Number |
| `AccM_id` | Whole Number |
| `Zipcode ID` | Whole Number |

Change types via the column header context menu → **Change Type**.

---

## 5. Products table – keep text categories

| Column | Keep as |
|---|---|
| `Producttype` | Text |
| `Productline` | Text |

No transformation needed; just confirm the types in the header icons.

---

## 6. Set up relationships (Model view)

After closing Power Query Editor and loading the data, open the **Model**
view and create the following relationships (all single-direction):

| From (Many) | To (One) | On column |
|---|---|---|
| `Orders[Customer ID]` | `Customers[Customer ID]` | Customer ID |
| `Orders[ProductID]` | `Products[Product ID]` | Product ID |
| `Orders[PaymethodID]` | `Paymentmethod[Payment Method ID]` | Payment Method ID |
| `Orders[DelcompID]` | `Deliverycomp[DC_ID]` | Delivery company ID |
| `Orders[Salesdate]` | `Calendar[Date]` | Date |

> **Note:** The Calendar table is already present in the DataModel and is
> driven by the *Start year calendar* and *End year calendar* parameters.
> Adjust those parameters to match the date range of your Orders data.

---

## 7. Add DAX measures

See `scripts/measures.dax` for all measure definitions.
Add them to the **Orders** table via **Home → New Measure**.
