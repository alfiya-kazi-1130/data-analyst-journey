# data-analyst-journey
My data analyst learning journey and projects.

## Day 1: Superstore Sales Analysis (Google Sheets)

### Business questions
1. Which regions generate the most sales and profit?
2. Which categories and sub-categories are the most and least profitable?
3. Are higher discounts linked to lower profit margins?

### Data cleaning
- Imported Superstore.csv (9,994 rows, 21 columns)
- Removed 1 duplicate row (ignoring Row ID); 9,993 rows remain
- Added a Profit Margin column (Profit / Sales)

### Findings
1. **Regions:** I found that the West region has the highest sales (725,458) and profit (108,418).
2. **Categories:** Technology earns the most profit. Furniture sells almost as much as Office Supplies but earns only about 2.5% margin. Tables (-17,725), Bookcases (-3,472) and Supplies (-1,189) lose money.
3. **Discounts:** Central has the highest average discount (about 24%) and the lowest profit margin (about 7.9%). The two may be linked.

### Tools used
Google Sheets: Remove Duplicates, Filters, Pivot Tables

## Day 2: Excel formulas on Superstore data (Google Sheets)

**Workbook:** `superstore_clean.xlsx`
- Tab `Superstore`: cleaned data (9,993 rows) with the helper columns I added: `Profit Margin` (V), `Result` (W), `Discount Band` (X), `Order Size` (Y)
- Tab `day-2`: every practice question, my answer and the formula I used (questions in column A, answers in B, formulas in C)

### Business questions
1. How many order lines lose money, and how large are those losses compared with the profit?
2. How does the discount level relate to profit?
3. How do I pull a specific detail (customer, product) from an ID?

### 1. Conditional totals, counts and averages (SUMIF, COUNTIF, AVERAGEIF)

| Question | Formula | Result |
|---|---|---|
| Total sales, Central | `=SUMIF(Superstore!M:M,"Central",Superstore!R:R)` | 501,239.89 |
| Total profit, Technology | `=SUMIF(Superstore!O:O,"Technology",Superstore!U:U)` | 145,454.95 |
| Total of all losses | `=SUMIF(Superstore!U:U,"<0",Superstore!U:U)` | -156,119.23 |
| Total of all profitable rows | `=SUMIF(Superstore!U:U,">0",Superstore!U:U)` | 442,528.31 |
| Losses as a share of profits | `=-B6/B7` | 35.28% |
| Rows in the West region | `=COUNTIF(Superstore!M2:M9994,"West")` | 3,203 |
| Loss-making rows | `=COUNTIF(Superstore!U2:U9994,"<0")` | 1,870 |
| Share of rows that lose money | `=B17/COUNT(Superstore!U2:U9994)` | 18.71% |
| Average discount, Central | `=AVERAGEIF(Superstore!M:M,"Central",Superstore!T:T)` | 24.04% |

I also tested the same functions on a small 7-row sample first (sheet rows 2-8) and checked the answers by hand before using whole columns.

### 2. Helper columns with IF and nested IF

| Column | Logic | Formula |
|---|---|---|
| Result (W) | Loss, Break-even or Profit | `=IF(U2<0,"Loss",IF(U2=0,"Break-even","Profit"))` |
| Discount Band (X) | No discount, Low (up to 20%), Medium (up to 50%), High (above 50%) | `=IF(T2=0,"No discount",IF(T2<=0.2,"Low",IF(T2<=0.5,"Medium","High")))` |
| Order Size (Y) | Large (1,000 or more), Medium (200 or more), otherwise Small | `=IF(R2>=1000,"Large",IF(R2>=200,"Medium","Small"))` |

Check on each helper column: the group counts add up to all 9,993 rows.
- Result: 1,870 Loss + 65 Break-even + 8,058 Profit
- Order Size: 468 Large + 2,110 Medium + 7,415 Small

### 3. Lookups (XLOOKUP, VLOOKUP, INDEX/MATCH)
- Looked up the customer name for ID SO-20335 with all three methods and got the same answer (Sean O'Donnell).
- Found the product name for Row ID 9 (DXL Angle-View Binders with Locking Rings by Samsill).
- Found the largest single sale with `MAX` (22,638.48), then looked up its product (Cisco TelePresence System EX90 Videoconferencing Unit) and customer (Sean Miller).
- Wrapped a lookup in `IFERROR` so an ID that does not exist shows "Not found" instead of an error.
- Noted that a lookup returns only the first match. Customer AA-10480 appears in several cities, so the city lookup shows only the first one.

### Key findings
- 1,870 of 9,993 order lines (18.71%) lose money. Together they lose 156,119, which equals about 35% of the profit earned by the profitable rows (442,528).
- A further 65 rows break even.
- Orders with a discount above 50% (856 rows) lost 76,559 in total. Orders with no discount earned 320,988.
- Central has the highest average discount (24%), which fits the low profit margin I found in Day 1.


### Tools used
Google Sheets: SUMIF, COUNTIF, AVERAGEIF, IF (nested), XLOOKUP, VLOOKUP, INDEX/MATCH, IFERROR, MAX

## Day 3: Statistics basics, charts and correlation (Google Sheets)

**Workbook:** `superstore_clean.xlsx`, tab `day-3` (questions in column A, answers in B, formulas in C)

### Business questions
1. What is a "typical" order, and how varied are profit, quantity and discount?
2. Which orders are unusual (outliers)?
3. How do sub-categories, regions, years and discounts compare visually?
4. Which numbers move together?

### 1. Mean, median and mode
| Measure | Formula | Result |
|---|---|---|
| Quantity: mean / median / mode | `=AVERAGE(...)`, `=MEDIAN(...)`, `=MODE(...)` | 3.79 / 3 / 3 |
| Profit: mean / median | | 28.66 / 8.67 |
| Discount: mean / median / mode | | 0.156 / 0.2 / 0 |

- Profit is skewed by a few very profitable orders, so the mean (28.66) is far above the median (8.67).
- The most common discount is 0 (about 48% of order lines), but the median is 0.2 and the mean is 0.156, so the data is lumpy.

### 2. Spread and outliers (Profit)
- Min -6,599.98, max 8,399.98, range 14,999.95.
- Q1 = 1.73, Q3 = 29.36, IQR = 27.63 (`=QUARTILE(range,1)` and `=QUARTILE(range,3)`).
- Standard deviation = 234.27, about eight times the mean (28.66), so profit varies a lot from order to order.
- Outlier fences: upper = Q3 + 1.5 x IQR = 70.81, lower = Q1 - 1.5 x IQR = -39.72.
- 1,277 rows are above the upper fence and 604 are below the lower fence.

### 3. Charts
- **Bar chart, profit by sub-category:** Tables (-17,725), Bookcases (-3,473) and Supplies (-1,189) lose money.
- **Column chart, sales and profit by region:** Central sells more than South (501,240 vs 391,722) but earns less profit (39,706 vs 46,749).
- **Line chart, sales by order year:** sales dipped slightly in 2012 (470,533), then grew in 2013 (608,474) and 2014 (733,947).
- **Scatter chart, average profit by discount level:** average profit per order line is 66.90 with no discount and 24.70 at 20%. It turns negative from 30% (-45.83) and reaches its lowest at 50% (-310.70).
- **Pie chart, sales by segment:** Consumer 50.6%, Corporate 30.7%, Home Office 18.7%.
- Chart choice guide: bar for comparing categories, line for change over time, scatter for the relationship between two numbers, pie for the share of a whole with few parts.

### 4. Correlation (`=CORREL(range1, range2)`)
| Pair | r | Reading |
|---|---|---|
| Sales and Profit | 0.48 | moderate positive |
| Quantity and Sales | 0.20 | weak positive |
| Discount and Profit | -0.22 | weak negative |
| Discount and Sales | -0.03 | essentially none |
| Quantity and Profit | 0.07 | essentially none |

- I built a full 4 x 4 correlation table (Sales, Quantity, Discount, Profit). Strongest pair: Sales and Profit. Weakest: Discount and Quantity (r = 0.01).

### Key findings
- A typical order line earns about 8.67 in profit, much less than the mean of 28.66.
- Average profit per order falls as the discount rises, and discounts above 20% lose money on average.
- Bigger discounts do not go with bigger sales (r = -0.03).

- ## Day 4: Cleaning messy data (Google Sheets)

**Workbook:** `messy_day4.xlsx` (tabs: original Messy, cleaning, cleaning log, scorecard)
**Data:** practice file made from a 1,225-row sample of Superstore with common data errors added on purpose, so every fix could be checked.

### Business question
Can I trust this data? What is wrong with it, and what did I change?

### 1. Profiling (measure first, change nothing)
| Check | Formula | Found |
|---|---|---|
| Data rows | `=COUNTA(range)` | 1,225 |
| Blank cells | `=COUNTBLANK(range)` | Postal Code 40, Discount 20 |
| Text instead of numbers or dates | `=SUMPRODUCT(--ISTEXT(range))` | Order Date 40, Sales 30, Profit 12 ("N/A") |
| Impossible values | `=COUNTIF(range,"<0")`, `">1"` | Quantity below 0: 8, Discount above 1: 10 |
| Labels not exactly valid | `=SUMPRODUCT(--EXACT(range,"label"))` | Region 35, Segment 20, Ship Mode 28 |

### 2. Fixing text
- Trimmed extra spaces and fixed capitalisation with `TRIM` and `PROPER`, working in helper columns, then pasting values only.
- Used a safer formula for names so correct names such as McMath are not damaged: `=IF(EXACT(F2,LOWER(F2)),PROPER(TRIM(F2)),TRIM(F2))`.
- Fixed Ship Mode typos with Find and replace (match entire cell contents).
- Counted changes with `=SUMPRODUCT(--NOT(EXACT(original,cleaned)))`: Customer Name 80, Region 35, Segment 20, Ship Mode 28.

### 3. Duplicates and blanks
- Counted duplicates before removing them: `=ROWS(range)-ROWS(UNIQUE(range))` gave 25. Removed them with Data cleanup, leaving 1,200 rows.
- Cleaned the text first, because `West` and `west ` would hide duplicates.
- Replaced the 12 "N/A" Profit cells with real blanks. A missing profit is unknown, not zero.
- Left Postal Code (40) and Discount (20) blank: unknown values, and guessing could be wrong.
- Built a missing-value report with `=BYCOL(range,LAMBDA(c,COUNTBLANK(c)))`.

### 4. Dates and numbers stored as text
- Converted 30 text Sales values (such as `$1,706.18`) with `SUBSTITUTE` and `VALUE`, and 40 text Order Dates with `DATE`, `RIGHT`, `MID` and `LEFT`.
- Total Sales changed from 309,296.47 to 307,530.23: converting text added about 4,669, and removing duplicates took out about 6,435.
- Years in the data after cleaning: 2011, 2012, 2013, 2014.

### 5. Impossible values
- Negative Quantity (8 rows): assumed a sign typo, so used `ABS()`.
- Discount above 1 (10 rows): assumed a percent typed as a whole number, so divided by 100.
- Added data validation: Quantity a whole number of 1 or more, Discount between 0 and 1.
- Both fixes are assumptions, and I recorded them in the log.

### Before and after
| Check | Before | After |
|---|---|---|
| Rows | 1,225 | 1,200 |
| Duplicates | 25 | 0 |
| Customer Name problems | 80 | 0 |
| Region / Segment / Ship Mode invalid | 35 / 20 / 28 | 0 |
| Sales stored as text | 30 | 0 |
| Order Date stored as text | 40 | 0 |
| Profit stored as "N/A" | 12 | 0 |
| Negative Quantity / Discount above 1 | 8 / 10 | 0 |
| Blank Postal Code / Discount / Profit | 40 / 20 / 12 | 40 / 20 / 12 (kept on purpose) |

### Tools used
Google Sheets: TRIM, PROPER, EXACT, SUMPRODUCT, COUNTBLANK, UNIQUE, FILTER, BYCOL, VALUE, SUBSTITUTE, DATE, Find and replace, Remove duplicates, Data validation

### Tools used
Google Sheets: AVERAGE, MEDIAN, MODE, QUARTILE, STDEV, CORREL, COUNTIF, pivot tables, column / bar / line / scatter / pie charts
