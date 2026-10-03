# Project 3.1: Executive Performance Matrix Dashboard

## Executive Summary
The **Executive Performance Matrix** is an enterprise-grade analytics dashboard designed to consolidate multi-channel sales metrics, regional profitability, and product hierarchy trends into a unified, interactive executive interface. Built on a relational star-schema data architecture, the workbook standardizes raw transactional records into real-time business insights.

---

## 1. Data Architecture & Relational Schema

The model is structured around a central **Fact Table** surrounded by normalized **Dimension Tables**:

* **`Dim_Products`**: Product master table containing `Product_ID`, `Category`, `Sub_Category`, `Unit_Cost`, and `Unit_Price`.
* **`Dim_Customers`**: Entity table containing `Customer_ID`, `Customer_Name`, `Region_Zone`, `State`, and `Postal_Code`.
* **`Fact_Sales`**: Transactional grain storing order parameters, quantities, discounts, and relational lookups:
  $$\text{Net Revenue} = \text{Gross Revenue} \times (1 - \text{Discount Pct})$$
  $$\text{Total Profit} = \text{Net Revenue} - (\text{Quantity} \times \text{Unit Cost})$$

---

## 2. Dynamic Formula Implementation

* **Relational Field Fetching (`XLOOKUP`)**:
  ```excel
  =XLOOKUP(D2, Dim_Products!A:A, Dim_Products!B:B, "Unknown")
  =XLOOKUP(C2, Dim_Customers!A:A, Dim_Customers!C:C, "Unknown")
