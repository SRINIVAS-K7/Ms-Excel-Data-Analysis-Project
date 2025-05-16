# Ms-Excel-Data-Analysis-Project

Overview : 

   This project involves performing a detailed supply chain data analysis using Microsoft Excel. It covers key performance indicators (KPIs) like order volume, delivery 
   timelines, discounts, and product movement to evaluate supply chain efficiency and uncover actionable insights.

   The dataset simulates real-world procurement and delivery operations across multiple sales channels, customers, warehouses, and products.

Dataset Description :

   Each row in the dataset represents a transaction/order, with the following attributes:

OrderNumber	:
Unique identifier for each order
Sales Channel	:
Mode of order placement (e.g., Online, Retail, Distributor)
WarehouseCode	:
Code of the warehouse handling the order
ProcuredDate	:
Date the product was procured
OrderDate	Date :
when the order was placed
ShipDate	:
Date when the order was shipped
DeliveryDate	:
Date when the product was delivered
CurrencyCode	:
Currency used in the transaction
_SalesTeamID	:
Internal ID for the sales team handling the order
_CustomerID	:
Unique ID for the customer placing the order
_StoreID	ID :
for the store (if applicable)
_ProductID	:
Product identifier
Order Quantity	:
Quantity of product ordered
Discount Applied	:
Discount (%) applied to the order
Unit Cost	:
Cost to procure one unit of the product
Unit Price	:
Selling price of one unit of the product

Objective :

   Analyze delivery performance, procurement lead time, and shipping delays.

   Identify best- and worst-performing sales channels, products, and warehouses.

   Evaluate margin trends by comparing Unit Cost vs Unit Price.

   Analyze discount trends and their impact on order quantity or delivery time.

Tools Used :

   Microsoft Excel

   Pivot Tables

   Excel Formulas (DATEDIF, IF, VLOOKUP/XLOOKUP, SUMIFS, etc.)

   Charts (Bar, Line, Column, Pie)
 
   Conditional Formatting

Dashboard Design :

Key Analyses Performed
Lead Time Analysis :
→   Calculated Procurement Lead Time, Shipping Time, and Delivery Time

Sales & Revenue KPIs :
→    Total Orders, Revenue, Profit, Discounted Sales Volume

Delivery Performance :
→    On-Time vs Late Deliveries by Warehouse and Region

Product-Level Analysis :
→    High-performing products, cost margin per product, quantity sold

Store/Customer Behavior :
→    Repeat orders, high-discount customers, customer segmentation

Profitability Analysis :
→    Revenue vs Cost breakdown per Sales Channel and Warehouse

Key Metrics Visualized :
   Total Orders

   Revenue & Profit

   Average Lead Time

   Orders by Channel

   Product Sales Volume

   On-Time Delivery Rate

Key Insights :

   Online channel has the fastest delivery but lowest profit margins due to high discounts.

   Warehouse W2 experiences the most delays, increasing overall delivery time.

   Product P108 is the most ordered but has the lowest profit margin.

   75% of late deliveries happen when the procurement-to-ship window is less than 2 days.

   Customers with larger discounts tend to place bulk orders.

Future Improvements :

   Integrate with Power BI for real-time dashboarding.

   Automate KPIs using Excel macros or Python.

   Extend analysis with geographic mapping for delivery routes.

   -------------------------------------------------------------

