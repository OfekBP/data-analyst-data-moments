-- Q3: net sales and profit per category in Q1 vs. Q4. Quarter = ((month - 1) / 3) + 1.
SELECT
    Category,
    ROUND(SUM(CASE WHEN q = 1 THEN LineSales END), 2)             AS Sales_Q1,
    ROUND(SUM(CASE WHEN q = 4 THEN LineSales END), 2)             AS Sales_Q4,
    ROUND(SUM(CASE WHEN q = 4 THEN LineSales END)
        / SUM(CASE WHEN q = 1 THEN LineSales END) - 1, 6)         AS SalesChangePct,
    ROUND(SUM(CASE WHEN q = 1 THEN LineSales - LineCost END), 2)  AS Profit_Q1,
    ROUND(SUM(CASE WHEN q = 4 THEN LineSales - LineCost END), 2)  AS Profit_Q4
FROM (
    SELECT *, ((CAST(strftime('%m', BillDate) AS INTEGER) - 1) / 3) + 1 AS q
    FROM sales
)
WHERE q IN (1, 4)
GROUP BY Category
ORDER BY Category;
