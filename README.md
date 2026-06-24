# SCM Material Supply Plan Automation

Python pandas 기반의 **자재 수급계획 자동화 프로젝트**입니다.
생산계획, BOM, 현재고, 미입고 발주, 자재마스터 데이터를 조합하여 자재별 필요수량, 부족수량, 과잉수량, 추천발주수량, 발주필요일, 협력사별 리스크를 자동 산출합니다.

---

## 1. Project Overview

This project automates a material supply planning report using Python and pandas.

It combines the following SCM datasets:

* Production Plan
* BOM Master
* Inventory
* Open Purchase Orders
* Material Master

The output is an Excel-based material supply planning report that helps procurement and SCM teams quickly identify:

* Required quantity by material
* Current stock and incoming quantity
* Available quantity
* Shortage and excess quantity
* Shortage and excess amount
* MOQ-based recommended order quantity
* Supplier-level shortage risk
* Order required date based on lead time
* Risk level by material

---

## 2. Business Context

In procurement and supply chain operations, buyers often need to check whether current inventory and incoming purchase orders are enough to cover production plans.

This project simulates that workflow by converting product-level production plans into material-level requirements using BOM data.

The main business logic is:

```text
Production Plan × BOM
→ Material Requirement
→ Inventory / Open PO / Material Master Mapping
→ Shortage / Excess Calculation
→ MOQ-based Order Recommendation
→ Supplier and Risk Summary
```

---

## 3. Input Files

The script is designed to run in a Kaggle Notebook environment using the `/kaggle/working` directory.

| File Name              | Description                          | Key Columns                                                        |
| ---------------------- | ------------------------------------ | ------------------------------------------------------------------ |
| `production_plan.xlsx` | Weekly production plan               | `plan_week`, `customer`, `model`, `product_code`, `qty`            |
| `bom_master.xlsx`      | BOM information by product           | `product_code`, `material_code`, `material_name`, `usage_qty`      |
| `inventory.xlsx`       | Current inventory by material        | `material_code`, `current_stock`                                   |
| `po_open.xlsx`         | Open purchase orders and ETA         | `material_code`, `po_no`, `eta_date`, `open_qty`                   |
| `material_master.xlsx` | Supplier, lead time, MOQ, unit price | `material_code`, `supplier`, `lead_time_days`, `moq`, `unit_price` |

---

## 4. Output File

The script generates the following Excel file:

```text
material_supply_plan.xlsx
```

The output Excel file contains three sheets:

| Sheet Name             | Description                                    |
| ---------------------- | ---------------------------------------------- |
| `material_supply_plan` | Detailed material-level supply planning report |
| `summary_by_risk`      | Summary by risk level                          |
| `summary_by_supplier`  | Summary by supplier                            |

---

## 5. Main Output Columns

The main output sheet includes the following columns:

| Column                     | Description                                               |
| -------------------------- | --------------------------------------------------------- |
| `plan_week`                | Production planning week                                  |
| `plan_month`               | Planning month converted from `plan_week`                 |
| `week_start_date`          | Start date of the ISO week                                |
| `required_date`            | Required date based on the planning week                  |
| `month_end_date`           | End date of the planning month                            |
| `material_code`            | Material code                                             |
| `material_name`            | Material name                                             |
| `required_qty`             | Required quantity calculated from production plan and BOM |
| `current_stock`            | Current inventory quantity                                |
| `incoming_qty`             | Incoming quantity from open purchase orders               |
| `available_qty`            | Current stock plus incoming quantity                      |
| `shortage_qty`             | Shortage quantity                                         |
| `shortage_amount`          | Shortage amount based on unit price                       |
| `excess_qty`               | Excess quantity                                           |
| `excess_amount`            | Excess amount based on unit price                         |
| `supplier`                 | Supplier name                                             |
| `lead_time_days`           | Lead time in days                                         |
| `moq`                      | Minimum order quantity                                    |
| `unit_price`               | Unit price                                                |
| `recommended_order_qty`    | MOQ-based recommended order quantity                      |
| `recommended_order_amount` | Recommended order amount                                  |
| `order_required_date`      | Order required date based on lead time                    |
| `risk_level`               | Material supply risk level                                |

---

## 6. Key Calculation Logic

### 6.1 Required Quantity

The production plan is merged with the BOM master by `product_code`.

```python
required_qty = qty * usage_qty
```

If the same material is used in multiple products, the required quantity is aggregated by planning week and material.

---

### 6.2 Incoming Quantity

Open purchase orders are reflected based on ETA.

```text
incoming_qty = sum(open_qty where eta_date <= month_end_date)
```

This allows the report to reflect only purchase orders expected to arrive within the relevant planning month.

---

### 6.3 Available Quantity

```python
available_qty = current_stock + incoming_qty
```

---

### 6.4 Shortage Quantity

```python
shortage_qty = max(required_qty - available_qty, 0)
```

Negative values are clipped to zero.

---

### 6.5 Excess Quantity

```python
excess_qty = max(available_qty - required_qty, 0)
```

Negative values are clipped to zero.

---

### 6.6 Shortage Amount

```python
shortage_amount = shortage_qty * unit_price
```

---

### 6.7 Excess Amount

```python
excess_amount = excess_qty * unit_price
```

---

### 6.8 MOQ-based Recommended Order Quantity

The recommended order quantity is calculated based on shortage quantity and MOQ.

```text
If shortage_qty <= 0:
    recommended_order_qty = 0

If moq <= 0:
    recommended_order_qty = shortage_qty

If moq > 0:
    recommended_order_qty = ceil(shortage_qty / moq) * moq
```

---

### 6.9 Order Required Date

The order required date is calculated using the material required date and lead time.

```python
order_required_date = required_date - lead_time_days
```

This helps identify when a purchase order should be placed to meet the required date.

---

## 7. Risk Level Criteria

The script assigns a risk level to each material based on shortage quantity and lead time.

| Risk Level | Criteria                                      | Meaning                                             |
| ---------- | --------------------------------------------- | --------------------------------------------------- |
| `High`     | `shortage_qty > 0` and `lead_time_days >= 45` | Shortage exists and lead time is long               |
| `Medium`   | `shortage_qty > 0` and `lead_time_days < 45`  | Shortage exists but lead time is relatively shorter |
| `Low`      | `shortage_qty == 0` and `excess_qty > 0`      | No shortage, but excess quantity exists             |
| `Normal`   | `shortage_qty == 0` and `excess_qty == 0`     | Balanced supply condition                           |

---

## 8. Validation Logic

The script includes validation checks to reduce calculation errors.

Validation includes:

* Input file existence check
* Required column check
* Numeric column conversion
* BOM missing product code warning
* Empty result validation
* Calculation consistency check

The script checks whether:

```text
shortage_qty >= 0
excess_qty >= 0
available_qty = current_stock + incoming_qty
shortage_amount = shortage_qty × unit_price
excess_amount = excess_qty × unit_price
recommended_order_amount = recommended_order_qty × unit_price
recommended_order_qty >= shortage_qty
```

If critical validation fails, the script raises an error instead of creating an unreliable output file.

---

## 9. Excel Formatting Features

The generated Excel file includes formatting features using `openpyxl`.

Applied formatting:

* Auto-adjusted column width
* Amount columns displayed with comma separators
* Date columns displayed as `yyyy-mm-dd`
* `risk_level` cells highlighted by color

Risk color rules:

| Risk Level | Cell Color |
| ---------- | ---------- |
| `High`     | Red        |
| `Medium`   | Yellow     |
| `Low`      | Green      |
| `Normal`   | White      |

---

## 10. How to Run

### 10.1 Environment

This script is designed for Kaggle Notebook.

Required libraries:

```python
pandas
numpy
openpyxl
```

### 10.2 Upload Input Files

Upload the following files to `/kaggle/working`:

```text
production_plan.xlsx
bom_master.xlsx
inventory.xlsx
po_open.xlsx
material_master.xlsx
```

### 10.3 Run the Script

Run the Python script in Kaggle Notebook.

```python
python material_supply_plan.py
```

Or paste and run the full code directly in a Kaggle Notebook cell.

### 10.4 Check the Output

The output file will be created at:

```text
/kaggle/working/material_supply_plan.xlsx
```

---

## 11. Repository Structure

Recommended repository structure:

```text
scm_material_supply_plan/
├── README.md
├── material_supply_plan.py
├── sample_data/
│   ├── production_plan.xlsx
│   ├── bom_master.xlsx
│   ├── inventory.xlsx
│   ├── po_open.xlsx
│   └── material_master.xlsx
└── output_sample/
    └── material_supply_plan.xlsx
```

---

## 12. Sample Data Notice

All data used in this project is anonymized sample data.

The sample data uses fictional values such as:

* `Customer A`
* `Supplier A`
* `PROD-A`
* `MAT-001`
* `PO-1001`

No real company data, customer data, supplier data, pricing data, production data, or confidential information is included.

---

## 13. Project Outcome

This project demonstrates how Python and pandas can be used to automate a procurement and SCM reporting workflow.

Key outcomes:

* Converted product-level production plans into material-level requirements
* Calculated material shortages and excess quantities
* Reflected current stock and incoming purchase orders
* Calculated shortage and excess amount by unit price
* Generated MOQ-based recommended order quantity
* Calculated order required date using lead time
* Created risk-level and supplier-level summary reports
* Exported the final result to a formatted Excel file

---

## 14. Portfolio Summary

This project shows practical automation capability in procurement and supply chain operations.

It can be summarized as:

> Built a Python pandas-based SCM automation tool that converts production plan and BOM data into a material-level supply planning report. The tool calculates required quantity, available quantity, shortage, excess, MOQ-based recommended order quantity, lead-time-based order required date, and supplier-level risk summary, then exports the result to a formatted Excel workbook.

---

## 15. Tech Stack

| Category        | Tools           |
| --------------- | --------------- |
| Language        | Python          |
| Data Processing | pandas, numpy   |
| Excel Export    | openpyxl        |
| Environment     | Kaggle Notebook |
| Output Format   | Excel workbook  |
