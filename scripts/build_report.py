#!/usr/bin/env python3
"""build_report.py
Rewrites 'PowerBI Training.pbix' with five report pages as specified in the
visualisation plan.

Usage (from repo root):
    python scripts/build_report.py

What it does
------------
* Reads the existing PBIX (which already contains the Calendar DataModel and
  template background image).
* Removes the single placeholder page ("ReportSection").
* Adds five new pages – each with the layout and visuals defined below –
  into the Report/definition/pages/ directory inside the ZIP.
* Re-packs the PBIX, keeping the DataModel entry as ZIP_STORED (no
  compression), which is required by Power BI Desktop.

After running this script, open the PBIX in Power BI Desktop and:
  1. Connect the five data-source files (see scripts/power_query_setup.md).
  2. Add the DAX measures listed in scripts/measures.dax.
  3. Verify all relationship joins in the Model view.
"""

import hashlib
import io
import json
import os
import zipfile

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PBIX_PATH = os.path.join(_SCRIPT_DIR, "..", "PowerBI Training.pbix")

# ---------------------------------------------------------------------------
# JSON schema URIs (copied from the existing PBIX)
# ---------------------------------------------------------------------------
SCHEMA_VISUAL = (
    "https://developer.microsoft.com/json-schemas/fabric/item/report/"
    "definition/visualContainer/2.7.0/schema.json"
)
SCHEMA_PAGE = (
    "https://developer.microsoft.com/json-schemas/fabric/item/report/"
    "definition/page/2.1.0/schema.json"
)
SCHEMA_PAGES = (
    "https://developer.microsoft.com/json-schemas/fabric/item/report/"
    "definition/pagesMetadata/1.0.0/schema.json"
)

BG_IMAGE = "getresponsive-background8687161759778721.jpg"
PAGE_W, PAGE_H = 1280, 720

# Ordered list of (internal_name, display_name)
PAGES = [
    ("SalesOverview",       "Sales Overview"),
    ("CustomerAnalysis",    "Customer Analysis"),
    ("ProductPerformance",  "Product Performance"),
    ("OperationalInsights", "Operational Insights"),
    ("TrendGrowth",         "Trend & Growth"),
]

# ---------------------------------------------------------------------------
# Aggregation function codes used in the Power BI visual JSON
# ---------------------------------------------------------------------------
SUM   = 0
AVG   = 1
MIN   = 2
MAX   = 3
COUNT = 5   # CountNonNull


# ---------------------------------------------------------------------------
# Low-level projection helpers
# ---------------------------------------------------------------------------

def _vid(seed: str) -> str:
    """Return a stable 20-char hex visual ID derived from a seed string."""
    return hashlib.md5(seed.encode()).hexdigest()[:20]


def col_proj(entity, prop, agg=None, query_ref=None, native_ref=None, active=False):
    """Projection that references a column, with an optional aggregation."""
    if agg is not None:
        field = {
            "Aggregation": {
                "Expression": {
                    "Column": {
                        "Expression": {"SourceRef": {"Entity": entity}},
                        "Property": prop,
                    }
                },
                "Function": agg,
            }
        }
        qref = query_ref or f"Aggregation({entity}.{prop},{agg})"
    else:
        field = {
            "Column": {
                "Expression": {"SourceRef": {"Entity": entity}},
                "Property": prop,
            }
        }
        qref = query_ref or f"{entity}.{prop}"
    nref = native_ref or prop
    proj = {"field": field, "queryRef": qref, "nativeQueryRef": nref}
    if active:
        proj["active"] = True
    return proj


def hier_proj(entity, prop, level, native_ref=None, active=False):
    """Projection on Power BI's auto-generated Date Hierarchy (Year/Quarter/Month/Day)."""
    field = {
        "HierarchyLevel": {
            "Expression": {
                "Hierarchy": {
                    "Expression": {
                        "PropertyVariationSource": {
                            "Expression": {"SourceRef": {"Entity": entity}},
                            "Name": "Variation",
                            "Property": prop,
                        }
                    },
                    "Hierarchy": "Date Hierarchy",
                }
            },
            "Level": level,
        }
    }
    qref = f"{entity}.{prop}.Variation.Date Hierarchy.{level}"
    nref = native_ref or f"{prop} {level}"
    proj = {"field": field, "queryRef": qref, "nativeQueryRef": nref}
    if active:
        proj["active"] = True
    return proj


# ---------------------------------------------------------------------------
# Visual-builder helpers
# ---------------------------------------------------------------------------

def _make_visual(vid_val, vtype, x, y, w, h, z, query_state):
    return {
        "$schema": SCHEMA_VISUAL,
        "name": vid_val,
        "position": {
            "x": x, "y": y, "z": z,
            "height": h, "width": w,
            "tabOrder": z,
        },
        "visual": {
            "visualType": vtype,
            "query": {"queryState": query_state},
            "drillFilterOtherVisuals": True,
        },
    }


def textbox(seed, x, y, w, h, text, z=0):
    vid_val = _vid(seed)
    return {
        "$schema": SCHEMA_VISUAL,
        "name": vid_val,
        "position": {
            "x": x, "y": y, "z": z,
            "height": h, "width": w,
            "tabOrder": z,
        },
        "visual": {
            "visualType": "textbox",
            "objects": {
                "general": [
                    {
                        "properties": {
                            "paragraphs": [
                                {
                                    "textRuns": [
                                        {
                                            "value": text,
                                            "textStyle": {
                                                "fontWeight": "bold",
                                                "fontFamily": "Calibri",
                                                "fontSize": "21pt",
                                            },
                                        }
                                    ]
                                }
                            ]
                        }
                    }
                ]
            },
            "visualContainerObjects": {
                "visualHeader": [
                    {
                        "properties": {
                            "show": {
                                "expr": {"Literal": {"Value": "false"}}
                            }
                        }
                    }
                ]
            },
            "drillFilterOtherVisuals": True,
        },
    }


def card(seed, x, y, w, h, entity, prop, agg=SUM, native_ref=None, z=0):
    return _make_visual(
        _vid(seed), "card", x, y, w, h, z,
        {"Values": {"projections": [col_proj(entity, prop, agg, native_ref=native_ref)]}},
    )


def slicer(seed, x, y, w, h, proj, z=0):
    return _make_visual(
        _vid(seed), "slicer", x, y, w, h, z,
        {"Values": {"projections": [proj]}},
    )


def bar_chart(seed, x, y, w, h, cat_proj_val, val_projs, z=0, chart_type="clusteredBarChart"):
    return _make_visual(
        _vid(seed), chart_type, x, y, w, h, z,
        {
            "Category": {"projections": [cat_proj_val]},
            "Y": {"projections": val_projs},
        },
    )


def line_chart(seed, x, y, w, h, cat_projs, val_projs, series_projs=None, z=0):
    qs = {
        "Category": {"projections": cat_projs},
        "Y": {"projections": val_projs},
    }
    if series_projs:
        qs["Series"] = {"projections": series_projs}
    return _make_visual(_vid(seed), "lineChart", x, y, w, h, z, qs)


def table_visual(seed, x, y, w, h, col_projs, z=0):
    return _make_visual(
        _vid(seed), "tableEx", x, y, w, h, z,
        {"Values": {"projections": col_projs}},
    )


def map_visual(seed, x, y, w, h, lat_p, lon_p, size_p=None, z=0):
    qs = {
        "Latitude":  {"projections": [lat_p]},
        "Longitude": {"projections": [lon_p]},
    }
    if size_p:
        qs["Size"] = {"projections": [size_p]}
    return _make_visual(_vid(seed), "map", x, y, w, h, z, qs)


def donut(seed, x, y, w, h, cat_p, val_p, z=0):
    return _make_visual(
        _vid(seed), "donutChart", x, y, w, h, z,
        {
            "Category": {"projections": [cat_p]},
            "Y": {"projections": [val_p]},
        },
    )


def treemap(seed, x, y, w, h, group_projs, val_projs, z=0):
    return _make_visual(
        _vid(seed), "treemap", x, y, w, h, z,
        {
            "Group": {"projections": group_projs},
            "Values": {"projections": val_projs},
        },
    )


def scatter(seed, x, y, w, h, x_p, y_p, detail_p=None, z=0):
    qs = {"X": {"projections": [x_p]}, "Y": {"projections": [y_p]}}
    if detail_p:
        qs["Details"] = {"projections": [detail_p]}
    return _make_visual(_vid(seed), "scatterChart", x, y, w, h, z, qs)


def waterfall(seed, x, y, w, h, cat_projs, y_projs, z=0):
    return _make_visual(
        _vid(seed), "waterfallChart", x, y, w, h, z,
        {
            "Category": {"projections": cat_projs},
            "Y": {"projections": y_projs},
        },
    )


def kpi_visual(seed, x, y, w, h, val_p, trend_p=None, z=0):
    qs = {"Value": {"projections": [val_p]}}
    if trend_p:
        qs["TrendAxis"] = {"projections": [trend_p]}
    return _make_visual(_vid(seed), "kpi", x, y, w, h, z, qs)


# ---------------------------------------------------------------------------
# Page background (re-used on every page)
# ---------------------------------------------------------------------------

def _page_bg_objects():
    return {
        "outspace": [{"properties": {}}],
        "background": [
            {
                "properties": {
                    "transparency": {"expr": {"Literal": {"Value": "0D"}}},
                    "image": {
                        "image": {
                            "name": {
                                "expr": {
                                    "Literal": {
                                        "Value": "'getresponsive-background.jpg'"
                                    }
                                }
                            },
                            "url": {
                                "expr": {
                                    "ResourcePackageItem": {
                                        "PackageName": "RegisteredResources",
                                        "PackageType": 1,
                                        "ItemName": BG_IMAGE,
                                    }
                                }
                            },
                            "scaling": {
                                "expr": {"Literal": {"Value": "'Fit'"}}
                            },
                        }
                    },
                }
            }
        ],
    }


# ---------------------------------------------------------------------------
# Page 1 – Sales Overview
# ---------------------------------------------------------------------------
# Layout (1280 × 720)
#  y=10  Title textbox
#  y=63  3 × KPI cards  |  2 × slicers (Year, Quarter)
#  y=170 Line chart (Sales by Month/Year)  |  Bar chart (Sales by Product Line)

def _page1_visuals():
    p = "p1_"
    return [
        textbox(p + "title",    20,   10, 600,  45, "Sales Overview", z=0),

        # KPI cards
        card(p + "card_sales",  20,   63, 225, 100,
             "Orders", "Sales", SUM, "Total Sales", z=1000),
        card(p + "card_orders", 253,  63, 225, 100,
             "Orders", "OrderID", COUNT, "Total Orders", z=2000),
        card(p + "card_aov",    486,  63, 225, 100,
             "Orders", "Sales", AVG, "Avg Order Value", z=3000),

        # Slicers
        slicer(p + "slicer_yr",  720,  63, 256, 100,
               hier_proj("Orders", "Salesdate", "Year",
                         native_ref="Year", active=True),
               z=4000),
        slicer(p + "slicer_qtr", 984,  63, 276, 100,
               hier_proj("Orders", "Salesdate", "Quarter",
                         native_ref="Quarter", active=True),
               z=5000),

        # Sales over time
        line_chart(
            p + "line_sales", 20, 170, 760, 530,
            cat_projs=[
                hier_proj("Orders", "Salesdate", "Month",
                          native_ref="Month", active=True)
            ],
            val_projs=[
                col_proj("Orders", "Sales", SUM, native_ref="Total Sales")
            ],
            series_projs=[
                hier_proj("Orders", "Salesdate", "Year", native_ref="Year")
            ],
            z=6000,
        ),

        # Sales by Product Line
        bar_chart(
            p + "bar_prodline", 790, 170, 465, 530,
            cat_proj_val=col_proj("Products", "Productline", active=True),
            val_projs=[col_proj("Orders", "Sales", SUM, native_ref="Sales")],
            z=7000,
        ),
    ]


# ---------------------------------------------------------------------------
# Page 2 – Customer Analysis
# ---------------------------------------------------------------------------
# Layout:
#  y=10   Title  |  Account Manager slicer (right)
#  y=65   Map (customers by Lat/Lon)  |  Bar chart (Sales by Province)
#  y=465  Top-N customer table

def _page2_visuals():
    p = "p2_"
    return [
        textbox(p + "title",   20,  10, 720,  45, "Customer Analysis", z=0),

        slicer(p + "slicer_accm", 1000, 10, 255, 45,
               col_proj("Customers", "AccM_id", active=True), z=500),

        map_visual(
            p + "map", 20, 65, 610, 390,
            lat_p=col_proj("Customers", "Latitude",  native_ref="Latitude"),
            lon_p=col_proj("Customers", "Longitude", native_ref="Longitude"),
            size_p=col_proj("Orders", "Sales", SUM, native_ref="Total Sales"),
            z=1000,
        ),

        bar_chart(
            p + "bar_province", 640, 65, 620, 390,
            cat_proj_val=col_proj("Customers", "Province", active=True),
            val_projs=[col_proj("Orders", "Sales", SUM, native_ref="Sales")],
            z=2000,
        ),

        table_visual(
            p + "table_cust", 20, 465, 1240, 240,
            col_projs=[
                col_proj("Customers", "Customer ID", active=True),
                col_proj("Customers", "Name"),
                col_proj("Customers", "Province"),
                col_proj("Orders", "Sales", SUM,
                         native_ref="Total Sales"),
                col_proj("Orders", "OrderID", COUNT,
                         native_ref="Total Orders"),
            ],
            z=3000,
        ),
    ]


# ---------------------------------------------------------------------------
# Page 3 – Product Performance
# ---------------------------------------------------------------------------
# Layout:
#  y=10   Title
#  y=65   Bar chart (Top products by Sales)  |  Treemap (Productline > type)
#  y=425  Scatter (Avg Price vs Total Amount per product)

def _page3_visuals():
    p = "p3_"
    return [
        textbox(p + "title", 20, 10, 600, 45, "Product Performance", z=0),

        bar_chart(
            p + "bar_prod", 20, 65, 615, 350,
            cat_proj_val=col_proj("Products", "Productname", active=True),
            val_projs=[col_proj("Orders", "Sales", SUM, native_ref="Sales")],
            z=1000,
        ),

        treemap(
            p + "treemap", 645, 65, 615, 350,
            group_projs=[
                col_proj("Products", "Productline", active=True),
                col_proj("Products", "Producttype"),
            ],
            val_projs=[col_proj("Orders", "Sales", SUM, native_ref="Sales")],
            z=2000,
        ),

        scatter(
            p + "scatter", 20, 425, 1240, 280,
            x_p=col_proj("Products", "Price", AVG, native_ref="Avg Price"),
            y_p=col_proj("Orders", "Amount", SUM, native_ref="Total Amount"),
            detail_p=col_proj("Products", "Productname", active=True),
            z=3000,
        ),
    ]


# ---------------------------------------------------------------------------
# Page 4 – Operational Insights
# ---------------------------------------------------------------------------
# Layout:
#  y=10   Title
#  y=65   Donut (Payment Method)  |  Bar (Delivery Company)  |  Card (Avg days) + KPI
#  y=455  Payment method detail table

def _page4_visuals():
    p = "p4_"
    return [
        textbox(p + "title", 20, 10, 600, 45, "Operational Insights", z=0),

        donut(
            p + "donut_pay", 20, 65, 400, 380,
            cat_p=col_proj(
                "Paymentmethod", "Payment Method Name", active=True
            ),
            val_p=col_proj("Orders", "OrderID", COUNT, native_ref="Orders"),
            z=1000,
        ),

        bar_chart(
            p + "bar_deliv", 430, 65, 420, 380,
            cat_proj_val=col_proj("Deliverycomp", "Company", active=True),
            val_projs=[
                col_proj("Orders", "OrderID", COUNT, native_ref="Orders")
            ],
            z=2000,
        ),

        card(
            p + "card_days", 860, 65, 390, 175,
            "Deliverycomp", "Gar delivery days", AVG,
            "Avg Delivery Days", z=3000,
        ),

        kpi_visual(
            p + "kpi_orders", 860, 250, 390, 195,
            val_p=col_proj(
                "Orders", "OrderID", COUNT, native_ref="Total Orders"
            ),
            trend_p=hier_proj(
                "Orders", "Salesdate", "Month", native_ref="Month"
            ),
            z=4000,
        ),

        table_visual(
            p + "table_pay", 20, 455, 830, 250,
            col_projs=[
                col_proj("Paymentmethod", "Payment Method Name", active=True),
                col_proj("Orders", "OrderID", COUNT,
                         native_ref="Total Orders"),
                col_proj("Orders", "Sales", SUM, native_ref="Total Sales"),
                col_proj("Orders", "Sales", AVG,
                         native_ref="Avg Order Value"),
            ],
            z=5000,
        ),
    ]


# ---------------------------------------------------------------------------
# Page 5 – Trend & Growth
# ---------------------------------------------------------------------------
# Layout:
#  y=10   Title
#  y=65   Line chart: Sales by Month with Year as series (YTD comparison)
#  y=375  Waterfall chart: Month-over-month Sales delta

def _page5_visuals():
    p = "p5_"
    return [
        textbox(p + "title", 20, 10, 600, 45, "Trend & Growth", z=0),

        line_chart(
            p + "line_ytd", 20, 65, 1240, 300,
            cat_projs=[
                hier_proj("Orders", "Salesdate", "Month",
                          native_ref="Month", active=True)
            ],
            val_projs=[
                col_proj("Orders", "Sales", SUM, native_ref="Total Sales")
            ],
            series_projs=[
                hier_proj("Orders", "Salesdate", "Year", native_ref="Year")
            ],
            z=1000,
        ),

        waterfall(
            p + "waterfall", 20, 375, 1240, 330,
            cat_projs=[
                hier_proj("Orders", "Salesdate", "Month",
                          native_ref="Month", active=True)
            ],
            y_projs=[
                col_proj("Orders", "Sales", SUM, native_ref="Total Sales")
            ],
            z=2000,
        ),
    ]


# ---------------------------------------------------------------------------
# Map: page internal name → visual builder
# ---------------------------------------------------------------------------
_PAGE_BUILDERS = {
    "SalesOverview":       _page1_visuals,
    "CustomerAnalysis":    _page2_visuals,
    "ProductPerformance":  _page3_visuals,
    "OperationalInsights": _page4_visuals,
    "TrendGrowth":         _page5_visuals,
}

_PAGES_PREFIX = "Report/definition/pages/"


# ---------------------------------------------------------------------------
# PBIX rebuild
# ---------------------------------------------------------------------------

def rebuild_pbix(pbix_path: str) -> None:
    """Read *pbix_path*, replace the report pages, write it back in place."""
    # --- Read every ZIP entry -------------------------------------------------
    with zipfile.ZipFile(pbix_path, "r") as zin:
        entries = {
            info.filename: (zin.read(info.filename), info.compress_type)
            for info in zin.infolist()
        }

    # --- Drop all existing page files ----------------------------------------
    entries = {
        k: v for k, v in entries.items()
        if not k.startswith(_PAGES_PREFIX)
    }

    # --- Build pages.json -----------------------------------------------------
    page_names = [name for name, _ in PAGES]
    pages_meta = {
        "$schema": SCHEMA_PAGES,
        "pageOrder": page_names,
        "activePageName": page_names[0],
    }
    entries[_PAGES_PREFIX + "pages.json"] = (
        json.dumps(pages_meta, indent=2).encode("utf-8"),
        zipfile.ZIP_DEFLATED,
    )

    # --- Build each page ------------------------------------------------------
    bg_objects = _page_bg_objects()
    for page_name, display_name in PAGES:
        # page.json
        page_doc = {
            "$schema": SCHEMA_PAGE,
            "name": page_name,
            "displayName": display_name,
            "displayOption": "FitToPage",
            "height": PAGE_H,
            "width": PAGE_W,
            "objects": bg_objects,
        }
        page_path = f"{_PAGES_PREFIX}{page_name}/page.json"
        entries[page_path] = (
            json.dumps(page_doc, indent=2).encode("utf-8"),
            zipfile.ZIP_DEFLATED,
        )

        # visual.json for each visual on the page
        for visual_def in _PAGE_BUILDERS[page_name]():
            visual_id = visual_def["name"]
            visual_path = (
                f"{_PAGES_PREFIX}{page_name}/visuals/{visual_id}/visual.json"
            )
            entries[visual_path] = (
                json.dumps(visual_def, indent=2).encode("utf-8"),
                zipfile.ZIP_DEFLATED,
            )

    # --- Re-pack into a new ZIP in memory then write atomically ---------------
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zout:
        for filename, (data, compress_type) in entries.items():
            # Power BI Desktop requires DataModel to be stored uncompressed.
            actual_compress = (
                zipfile.ZIP_STORED
                if filename == "DataModel"
                else compress_type
            )
            zout.writestr(filename, data, compress_type=actual_compress)

    with open(pbix_path, "wb") as fh:
        fh.write(buf.getvalue())

    # --- Summary --------------------------------------------------------------
    page_summary = ", ".join(f'"{d}"' for _, d in PAGES)
    print(f"Updated: {pbix_path}")
    print(f"Pages  : {page_summary}")
    total_visuals = sum(len(_PAGE_BUILDERS[n]()) for n, _ in PAGES)
    print(f"Visuals: {total_visuals} across {len(PAGES)} pages")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    rebuild_pbix(PBIX_PATH)
