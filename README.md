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
