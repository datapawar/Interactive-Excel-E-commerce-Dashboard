<img src="https://github.com/datapawar/Interactive-Excel-E-commerce-Dashboard/blob/main/thumbnail%20image..png" width="100%" height="100%">

# Interactive Excel E‑commerce Dashboard

Interactive Excel dashboard built from an e-commerce transactions dataset (Jan–Mar 2025) to track KPIs, trends, customer behavior, delivery performance, geography, and satisfaction. Inspired by Chandoo’s dashboard tutorial: https://www.youtube.com/watch?v=l5qkg8gzY6E

## What’s inside
- **Dashboard sheet**: Executive KPIs + interactive visuals
- **Pivots sheet** (or similar): PivotTables that act as the calculation engine
- **Data sheet**: Raw transactional data + helper columns

## Business questions this dashboard answers
- What are our headline KPIs (orders, quantity, revenue, avg rating, avg delivery time)?
  * Total performance (actuals): 2,400 orders and 11,997 units sold, generating 649,019.8 in total sales amount.
  * Revenue per order is healthy: average order value is ~270.42 (649,019.8 / 2,400), which is a solid AOV benchmark for an ecommerce apparel-style catalog.
  * Customer experience is strong: average customer rating is 3.96 / 5 (≈4.0), indicating generally positive satisfaction.
  
- How fast are we delivering (distribution of Days to Deliver)?
  * Fulfilment speed looks efficient: average delivery time is 2.34 days (≈2.3 days), suggesting fast shipping performance overall.
  * Basket size is moderate: average units per order is ~5.0 (11,997 units / 2,400 orders), implying customers often buy multiple items per order rather than single-item checkouts.

- How do customers prefer to buy (Order Mode) and how does that vary by Gender?
  * The App is the #1 channel for every gender (≈35–37% share), so it’s the dominant buying preference overall and within segments.
  * Website is consistently the #2 channel for Female and Male customers (≈24% share), indicating a strong “web” cohort alongside app-first buyers.

- Which products are most popular, and what does the gender split look like?
  * The most popular products (by units sold) are T‑Shirts, Jeans, and Sneakers.
​  * Gender split varies by product type: some are female-skewed (e.g., Graphic Tees, Sundresses), while others are more balanced (e.g., T‑Shirts, Sandals) or male-skewed (Hoodies & Sweatshirts).

- Which geographies (state/county) contribute the most quantity and revenue?
  * The top 2 counties (Los Angeles + San Diego) contribute ~48% of revenue combined (33.2% + 15.2%), showing a heavy concentration in just two counties.

  ## How this dashboard can help a company
- **Leadership visibility**: KPIs + trends give quick “health checks” for revenue, volume, delivery speed, and satisfaction.
- **Channel optimization**: Order Mode insights show where revenue is coming from and which segments are strongest.
- **Product decisions**: Identifies top-selling products and segment engagement patterns.
- **Operations improvements**: Delivery-time distribution highlights bottlenecks and sets targets for faster delivery.
- **Customer experience tracking**: Rating breakdown by month helps monitor satisfaction changes over time.
- **Territory planning**: Maps highlight high-performing or underperforming regions.

## Excel components used (and why)
### 1) Data preparation (helper columns)
- **Days to Deliver** = Ship Date − Order Date (enables delivery performance analysis).  
- **Week Number** from Order Date (enables weekly trend view for “last 13 weeks”).  
- **Cleaned Gender field** using an IF-style mapping (turns inconsistent codes/blanks into readable slicer values like Male/Female/Other/Unknown).

These columns make the PivotTables simpler and the slicers/labels cleaner.

### 2) PivotTables as a calculation engine
PivotTables were used to compute:
- Count of transactions (orders)
- Sum of quantity and amount (revenue)
- Average rating and average days-to-deliver (using “Summarize Values By → Average”)
- Breakdowns for weekly trend, product popularity, delivery distribution, and satisfaction spread

This keeps calculations dynamic: refresh pivots when new data is added.

### 3) PivotCharts + formatting
Charts created from PivotTables to visualize:
- 13-week trend (combo chart with a secondary axis to handle different scales)
- Product popularity (bar chart, sorted)
- Delivery-time distribution
- Satisfaction by month and rating (clustered columns)

Formatting focused on readability (reduced clutter like legend/field buttons where appropriate).

### 4) Purchase pattern matrix (linked picture + conditional formatting)
To create a dashboard-friendly “heatmap”:
- PivotTable showing **Order Mode × Gender** with **Show Values As → % of Grand Total**
- A linked cell grid was formatted with **conditional formatting color scales**
- The grid was copied as a **linked picture** and placed on the dashboard

This approach keeps the dashboard clean while staying fully dynamic.

### 5) Geographic view (Map chart)
Map charts were created from a regular cell range (fed by pivot outputs), because map charts can’t be built directly from PivotTable ranges.

### 6) Interactivity (Slicers + Report Connections)
Slicers were added for:
- Order Mode
- Gender

They were connected across relevant PivotTables using **Report Connections**, so one click filters multiple visuals consistently.

## How to use
1. Download the Excel file in this repo.
2. Open in Excel (2016+ recommended for map charts).
3. Use slicers to filter by Order Mode and Gender.
4. Refresh PivotTables after adding/replacing data (`Data → Refresh All`).

## Credits
- Tutorial inspiration: Chandoo – “The Ultimate Excel Dashboard: Visualize Data Like a Pro”  
  https://www.youtube.com/watch?v=l5qkg8gzY6E
