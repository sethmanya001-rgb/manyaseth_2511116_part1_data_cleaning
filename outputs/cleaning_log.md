# 🧹 Cleaning Log — Retail Orders Dataset
**Project:** Part 1 — Business Data Cleaning, Validation & Excel Reporting  
**Dataset:** raw_orders.xlsx  
**Total Records:** 932  
**Date Range:** January 2024 – December 2025  
**Analyst:** Manya | BITSoM × Masai School

---

## 1. 🔍 Issues Found

### 1.1 Text Field Issues
| Field | Issue Found | Count |
|---|---|---|
| customer_name | Leading/trailing spaces, inconsistent casing | Multiple |
| segment | Mixed case (e.g. "consumer", "CONSUMER") | Multiple |
| region | Missing values, inconsistent casing | 26 missing |
| state | Extra spaces, inconsistent casing | Multiple |
| city | Leading/trailing spaces | Multiple |
| category | Mixed case, minor spelling variations | Multiple |
| sub_category | Inconsistent formatting | Multiple |
| ship_mode | Missing values, mixed case | 22 missing |
| payment_status | Inconsistent casing | Multiple |
| order_status | Inconsistent casing | Multiple |

### 1.2 Date Issues
| Issue | Count |
|---|---|
| Mixed date formats in order_date (text + date) | Multiple |
| Mixed date formats in ship_date | Multiple |
| Ship date before order date (negative delay) | 196 |
| Same day shipping (delay = 0 days) | 90 |
| Long delivery delay (>30 days) | 187 |
| Missing order_date | 0 |
| Missing ship_date | 0 |

### 1.3 Duplicate Issues
| Type | Count |
|---|---|
| Exact duplicate rows | 40 |
| Conflicting duplicates (same order_id, different data) | 23 |
| Unique records | 869 |

### 1.4 Discount Issues
| Issue | Count |
|---|---|
| Missing / blank discount | 26 |
| Negative discount values | 16 |
| Discount above allowed range (>1) | 0 |
| Valid discount (0 to 1) | 890 |

### 1.5 Sales & Profit Issues
| Issue | Count |
|---|---|
| Sales mismatch (calculated vs original >1 difference) | 51 |
| Missing Calculated_Sales (due to Unknown discount) | 18 |
| Missing Calculated_Profit | 18 |
| Missing Profit_Margin | 18 |

### 1.6 Order & Payment Status Issues
| Status | Count |
|---|---|
| Cancelled orders | 146 |
| Returned orders | 164 |
| Failed payments | 69 |
| Pending payments | 87 |

---

## 2. 🧼 Cleaning Actions Performed

### 2.1 Text Cleaning
- Applied `=TRIM(PROPER(cell))` formula to all text fields
- Created separate `Cleaned_*` columns for every field to preserve original data
- Used Find & Replace for specific inconsistencies
- Standardized all text to Title Case

**Columns created:**
`Cleaned Customer_Name`, `Cleaned_Segment`, `Cleaned_Region`, `Cleaned_State`, `Cleaned_City`, `Cleaned_Category`, `Cleaned_Sub_category`, `Cleaned_ship_Mode`, `Cleaned_Payment_Status`, `Cleaned_order_Status`

### 2.2 Date Cleaning
- Converted all date values to consistent DD-MM-YYYY format
- Created `Cleaned_Order_Date` and `Cleaned_Ship_Date` columns
- Used `ISNUMBER()` to detect text values stored as dates
- Calculated `Shipping_Delay_Days` = `Cleaned_Ship_Date - Cleaned_Order_Date`
- Extracted `Order_Month` using `TEXT(DATE(2024,MONTH(date),1),"MMMM")`
- Extracted `Order_Year` using `YEAR()` function

**Flag columns created:**
- `Order_Date_Flag`: Valid / Invalid - Missing Date / Invalid - Not a Date
- `Ship_Date_Flag`: Valid / Invalid - Missing Date / Invalid - Ship Before Order / Warning - Same Day Ship / Warning - Long Delay

### 2.3 Duplicate Handling
- Used `COUNTIF` to detect repeated `order_id` values
- Used `COUNTIFS` across `order_id + order_date + ship_date + customer_id` to detect exact duplicates
- Created `Duplicate_Flags` column with three values:
  - **Unique**: order_id appears only once
  - **Exact Duplicate**: all key fields match exactly
  - **Conflicting Duplicate**: same order_id but different details → flagged for review

**Records removed:** 0 (all flagged, none deleted to avoid accidental data loss)

### 2.4 Missing Value Handling
- `region` (26 blanks): Filled as **"Unknown"** as per business rule
- `ship_mode` (22 blanks): Filled as **"Unknown"** as per business rule
- `discount` (26 blanks): Set to **0** only where quantity, unit_price, and sales were all valid; else marked **"Unknown"**

### 2.5 Discount Cleaning
Formula used for `Cleaned Discount`:
```
=IF(ISBLANK(discount), IF(quantity*unit_price=cost, 0, "Unknown"), IF(discount<0, 0, IF(discount>1, 0, discount)))
```

**Flag column `Flag Discount`:**
```
=IF(discount="","Missing Discount", IF(discount<0,"Invalid - Negative Discount", IF(discount>1,"Invalid - Above Range","Valid")))
```

### 2.6 Calculated Columns Created
| Column | Formula Logic |
|---|---|
| Calculated_Sales | quantity × unit_price × (1 - Cleaned Discount) |
| Calculated Profit | Calculated_Sales - cost |
| Profit Margin | Calculated Profit / Calculated_Sales (IFERROR to handle blanks) |
| Shipping_Delay_Days | Cleaned_Ship_Date - Cleaned_Order_Date |
| Order_Month | TEXT(DATE(2024,MONTH(Cleaned_Order_Date),1),"MMMM") |
| Order_Year | YEAR(Cleaned_Order_Date) |
| Data_Quality_Flags | Combined flag: Clean / Warning / Invalid / Needs Review |

### 2.7 Data Quality Flag Logic
```
=IF(Duplicate="Exact Duplicate","Warning - Exact Duplicate",
  IF(Duplicate="Conflicting Duplicate","Needs Review - Conflicting Duplicate",
    IF(Order_Date_Flag="Invalid - Missing Date","Invalid - Missing Order Date",
      IF(Ship_Date_Flag="Invalid - Ship Before Order","Invalid - Ship Before Order",
        IF(discount<0,"Invalid - Negative Discount",
          IF(discount>1,"Invalid - Above Range Discount",
            IF(Payment_Status="Failed","Invalid - Failed Payment",
              IF(Order_Status="Cancelled","Warning - Cancelled Order",
                IF(Order_Status="Returned","Needs Review - Refunded Order",
                  "Clean")))))))))
```

---

## 3. 📋 Business Rules Applied

| Rule | Decision | Reason |
|---|---|---|
| Missing region | Fill as "Unknown" | Preserve record; region cannot be assumed |
| Missing ship_mode | Fill as "Unknown" | Preserve record; mode cannot be assumed |
| Missing discount | Treat as 0 if sales fields valid | No discount likely means full price sale |
| Negative discount | Clean to 0 | Negative discount is not a valid business value |
| Discount > 1 | Flag as invalid | Discount range must be 0 to 1 (0% to 100%) |
| Cancelled orders | Exclude from sales summary | Business rule: cancelled = no revenue realized |
| Failed payments | Exclude from sales summary | Business rule: failed payment = no revenue confirmed |
| Refunded orders | Summarize separately | Need separate tracking for refund analysis |
| Ship before order | Flag as invalid | Logistically impossible; data entry error |
| Same day shipping | Flag as warning | Possible but needs verification |
| Delay > 30 days | Flag as warning | Unusual; may indicate logistics issue |
| Conflicting duplicates | Flag, do not delete | Cannot determine correct record without source verification |

---

## 4. 📌 Assumptions Made

1. Discount values should be between 0 and 1; anything outside this range is invalid
2. Where discount was missing and all other fields were valid, discount was assumed to be 0 (full price sale)
3. "Unknown" was used for missing region and ship_mode to preserve record count for analysis
4. Conflicting duplicates were retained and flagged — not deleted — to avoid accidental data loss
5. Sales mismatch of more than ₹1 was considered a significant mismatch
6. Order_Month was extracted as full month name (e.g., January) for better readability in pivots

---

## 5. 🗑️ Records Removed

| Action | Count |
|---|---|
| Records permanently deleted | **0** |
| Records flagged as Exact Duplicate | 40 |
| Records flagged as Conflicting Duplicate | 23 |
| Records flagged as Needs Review | 16 |

> **Note:** No records were permanently deleted. All flagged records are retained in `cleaned_orders.xlsx` with appropriate flags for business review.

---

## 6. 🚩 Records Flagged

| Flag Type | Count |
|---|---|
| Invalid - Ship Before Order | 196 |
| Invalid - Negative Discount | 16 |
| Invalid - Failed Payment | 69 |
| Warning - Exact Duplicate | 40 |
| Warning - Cancelled Order | 146 |
| Warning - Long Delay | 187 |
| Warning - Same Day Ship | 90 |
| Needs Review - Conflicting Duplicate | 23 |
| Needs Review - Refunded Order | 164 |
| Clean | 916 |

---

## 7. ⚠️ Limitations of the Cleaning Process

1. **196 invalid shipping records** (ship before order) cannot be corrected without access to the original source system — they can only be flagged
2. **18 records** with "Unknown" discount have no calculable sales/profit — these are excluded from financial summaries
3. **Conflicting duplicates (23 records)** require manual business verification — the correct version cannot be determined from data alone
4. **Sales mismatch (51 records)** may be due to promotional pricing, rounding differences, or system-level adjustments not captured in the dataset
5. All cleaning was performed in Excel without direct access to the source database, so root cause fixes are not possible at this stage
6. Month ordering in pivots required manual sorting as Excel treats month names as text

---

*Log maintained by: Manya | BITSoM × Masai School — Business Analytics with Generative AI*  
*Last updated: 2025*
