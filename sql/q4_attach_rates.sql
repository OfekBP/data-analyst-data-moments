-- Q4: share of receipts that contain each category, per store (long format: one row per store x category).
WITH bill_cat AS (
    SELECT DISTINCT StoreName, Bill, Category FROM sales
),
store_bills AS (
    SELECT StoreName, COUNT(DISTINCT Bill) AS n FROM sales GROUP BY StoreName
)
SELECT
    c.StoreName,
    c.Category,
    ROUND(1.0 * COUNT(*) / b.n, 6) AS Attach
FROM bill_cat AS c
JOIN store_bills AS b ON b.StoreName = c.StoreName
GROUP BY c.StoreName, c.Category
ORDER BY c.StoreName, c.Category;
