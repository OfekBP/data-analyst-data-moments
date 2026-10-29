-- Q1: store comparison. Table `sales` = cleaned lines (duplicates removed), `Bill` = StoreCode-BillNum.
SELECT
    StoreName,
    ROUND(SUM(LineSales), 2)                                   AS Sales,
    ROUND(SUM(LineSales - LineCost), 2)                        AS Profit,
    ROUND(SUM(LineSales - LineCost) / SUM(LineSales), 6)       AS Margin,
    COUNT(DISTINCT Bill)                                       AS Bills,
    ROUND(SUM(LineSales) * 1.0 / COUNT(DISTINCT Bill), 4)      AS AvgBill
FROM sales
GROUP BY StoreName
ORDER BY StoreName;
