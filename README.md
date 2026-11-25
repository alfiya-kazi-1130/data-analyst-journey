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
