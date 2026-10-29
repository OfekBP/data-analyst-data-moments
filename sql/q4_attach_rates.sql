-- Q4: share of receipts containing Dry Food / Wet Food / Snacks, per store.
WITH bills AS (
    SELECT
        StoreName,
        Bill,
        MAX(Category = 'Dry Food') AS Dry,
        MAX(Category = 'Wet Food') AS Wet,
        MAX(Category = 'Snacks')   AS Snacks
    FROM sales
    GROUP BY StoreName, Bill
)
SELECT
    StoreName,
    ROUND(AVG(Dry), 6)    AS Dry,
    ROUND(AVG(Wet), 6)    AS Wet,
    ROUND(AVG(Snacks), 6) AS Snacks
FROM bills
GROUP BY StoreName
ORDER BY StoreName;
