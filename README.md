# Inventory-tracker-fifo-excel
Excel-based inventory management dashboard using FIFO costing, Power Pivot, and DAX measures. Tracks COGS, revenue, profit, and stock levels across multiple products with interactive slicers and pivot charts.
📦 Inventory Tracker Dashboard with FIFO Costing | Excel + Power Pivot
     
An interactive inventory management dashboard built in Microsoft Excel that calculates Cost of Goods Sold (COGS) using the First In, First Out (FIFO) method. The dashboard uses Power Pivot, DAX measures, and dynamic pivot tables to deliver real-time business insights across multiple products.
________________________________________
🔍 Project Overview
This project simulates a real-world inventory tracking system used in supply chain and retail operations. It enables businesses to:
•	Track purchase and sales transactions across multiple products
•	Accurately compute COGS using FIFO inventory valuation
•	Monitor profit, revenue, and remaining inventory in real time
•	Filter and analyze data interactively by product using slicers
________________________________________
📊 Dashboard Preview
The final dashboard displays:
KPI	Description
💰 Revenue	Total income from sales
🏷️ COGS	Cost of Goods Sold via FIFO
📈 Profit	Revenue minus COGS
📦 Quantity Remaining	Units left in inventory
🔄 Qty Sold vs Purchased	Doughnut chart comparison
📅 Monthly Trends	Clustered column chart by month
A product slicer connects all pivot tables and charts, enabling synchronized filtering across the entire dashboard.
________________________________________
🗂️ File Structure
inventory-tracker-fifo/
│
├── Inventory_Tracker_Dashboard_with_FIFO_Calculations.xlsx   # Main Excel workbook
└── README.md                                                  # Project documentation
Sheets Inside the Workbook
Sheet Name	Purpose
Purchases	Records purchase date, product, quantity, and unit price
Sales	Records sales date, product, quantity sold, and selling price
Register	Unique product list used as a dimension table
Dashboard	Final interactive KPI dashboard with charts and slicers
________________________________________
⚙️ How It Works
1. FIFO Inventory Allocation
The FIFO formula allocates sales to purchase batches in chronological order. For each purchase batch, it calculates:
•	Units Sold per Batch — How many units from that specific batch were sold
•	Units Remaining — Quantity Purchased minus Units Sold from that batch
•	COGS per Batch — Units Sold × Unit Price of that batch
•	Inventory Value — Units Remaining × Unit Price
The key formula uses MAX, MIN, and SUMIF logic to determine how many units from each purchase batch have been consumed by cumulative sales, ensuring strict FIFO ordering.
2. Data Model (Power Pivot)
The three main tables are loaded into Excel's Power Pivot data model with the following relationships:
Register (Product) ──< Purchases (Product)
Register (Product) ──< Sales (Product)
Calendar (Date)    ──< Purchases (Purchase Date)
Calendar (Date)    ──< Sales (Sale Date)
A Calendar Table is created inside Power Pivot to enable time-based filtering and month-level trend analysis.
3. DAX Measures
Measure	Formula Logic
Revenue	SUMX(Sales, Qty Sold × Selling Price)
COGS	SUM(Purchases[Cost of Goods Sold])
Profit	[Revenue] - [COGS]
Quantity Remaining	SUM(Purchases[Units Remaining])
Quantity Sold	SUM(Sales[Quantity Sold])
Quantity Purchased	SUM(Purchases[Quantity Purchased])
________________________________________
🚀 Getting Started
Prerequisites
•	Microsoft Excel 2016 or later (with Power Pivot enabled)
•	Windows OS recommended (Power Pivot is not available on Excel for Mac)
How to Enable Power Pivot (if not visible)
1.	Go to File → Options → Add-ins
2.	At the bottom, select COM Add-ins from the dropdown → Click Go
3.	Check Microsoft Power Pivot for Excel → Click OK
Using the File
1.	Download the .xlsx file from this repository
2.	Open it in Microsoft Excel
3.	If prompted, click Enable Content to allow data model features
4.	Use the product slicer on the dashboard to filter by item
5.	To add new data: 
o	Navigate to the Purchases or Sales sheet
o	Go to the last row of the table and press Tab to extend formulas automatically
o	Return to the Dashboard and Refresh All (Data → Refresh All)
________________________________________
💡 Key Concepts Demonstrated
•	FIFO Inventory Valuation — Industry-standard costing method used in supply chain, retail, and manufacturing
•	Power Pivot & Data Modeling — Building relationships between fact and dimension tables
•	DAX (Data Analysis Expressions) — Writing measures for aggregated KPIs
•	Dynamic Excel Tables — Auto-expanding ranges for scalable data entry
•	Interactive Dashboard Design — Slicers, pivot charts, and KPI cards
________________________________________
🏭 Supply Chain Relevance
This project mirrors real tools used in supply chain management:
Industry Use Case	How This Project Covers It
Warehouse Inventory Control	Tracks stock levels per product batch
Procurement Cost Tracking	Logs purchase price per batch for accurate costing
COGS Reporting	FIFO-based cost allocation used in financial reporting
Demand vs Supply Analysis	Qty Sold vs Purchased chart shows imbalances
Multi-product Management	Product slicer enables per-SKU analysis
________________________________________
📌 Tips & Notes
•	Do not use Merge & Center — The dashboard uses "Center Across Selection" to keep cells clean for referencing
•	GetPivotData is disabled to allow simple cell referencing from pivot tables
•	Month labels in the calendar table are formatted as abbreviations (Jan, Feb...) and sorted by month number — not alphabetically
•	When adding new rows, always use the Tab key from the last cell to avoid formula errors
________________________________________
🤝 Connect
If you found this project useful or have suggestions, feel free to connect:
•	💼 [www.linkedin.com/in/swati-kumari-58a04a203](#) Linkedin 
•	📧[Swatikumari.connect540@gmail.com](#) 
________________________________________
📄 License
This project is for educational and portfolio purposes. Data used is sample/dummy data.
