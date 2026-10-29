-- Q2: mean and median receipt amount. SQLite has no MEDIAN, so take the middle row(s) by rank.
WITH bills AS (
    SELECT Bill, SUM(LineSales) AS Total
    FROM sales
    GROUP BY Bill
),
ranked AS (
    SELECT Total,
           ROW_NUMBER() OVER (ORDER BY Total) AS rn,
           COUNT(*)     OVER ()               AS n
    FROM bills
)
SELECT
    (SELECT COUNT(*) FROM bills)                                  AS Bills,
    ROUND((SELECT AVG(Total) FROM bills), 4)                      AS Mean,
    ROUND(AVG(CASE WHEN rn IN ((n + 1) / 2, (n + 2) / 2) THEN Total END), 4) AS Median
FROM ranked;
