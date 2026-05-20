-- Stored procedure to return dashboard summary for MPRs
-- Returns latest statistics per MPR header including BaseNo and MaxRev so client can mark older revisions

IF OBJECT_ID('dbo.sp_GetMPRDashboardSummary', 'P') IS NOT NULL
    DROP PROCEDURE dbo.sp_GetMPRDashboardSummary;
GO

CREATE PROCEDURE dbo.sp_GetMPRDashboardSummary
    @search NVARCHAR(4000) = NULL,
    @status NVARCHAR(100) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    WITH Headers AS (
        SELECT
            MPR_ID,
            MPR_No,
            Project_Name,
            Required_Date,
            Status,
            ISNULL(TRY_CAST(TRY_CAST(Rev AS DECIMAL(10,2)) AS INT), 0) AS Rev,
            Created_Date,
            CASE WHEN CHARINDEX('_Rev.', MPR_No) > 0 THEN LEFT(MPR_No, CHARINDEX('_Rev.', MPR_No) - 1) ELSE MPR_No END AS BaseNo
        FROM MPR_Header
    ),
    Filtered AS (
        SELECT * FROM Headers h
        WHERE (@search IS NULL OR @search = '' OR (h.MPR_No LIKE N'%' + @search + N'%' OR h.Project_Name LIKE N'%' + @search + N'%'))
          AND (@status IS NULL OR @status = '' OR h.Status = @status)
    ),
    SeriesTotals AS (
        -- use Headers CTE which already computes BaseNo
        SELECT h.BaseNo, COUNT(DISTINCT d.Detail_ID) AS TotalItems
        FROM Headers h
        INNER JOIN MPR_Details d ON d.MPR_ID = h.MPR_ID
        GROUP BY h.BaseNo
    ),
    Ordered AS (
        SELECT h.BaseNo,
               COUNT(DISTINCT pod.PO_Detail_ID) AS OrderedItems,
               COUNT(DISTINCT pod.PO_ID) AS POCount
        FROM PO_Detail pod
        INNER JOIN MPR_Details md ON pod.MPR_Detail_ID = md.Detail_ID
        INNER JOIN Headers h ON md.MPR_ID = h.MPR_ID
        GROUP BY h.BaseNo
    ),
    MaxRev AS (
        SELECT h.BaseNo AS BaseNo, MAX(h.Rev) AS MaxRev
        FROM Headers h
        GROUP BY h.BaseNo
    )
    SELECT
        f.MPR_ID,
        f.MPR_No AS [MPR No],
        f.Project_Name AS [D? án],
        f.Required_Date AS [Ngày c?n],
        f.Status AS [Tr?ng thái],
        f.Rev AS [Rev],
        CASE WHEN ISNULL(o.POCount, 0) > 0 THEN N'? ' + CAST(o.POCount AS NVARCHAR(10)) + N' PO' ELSE N'? Ch?a có PO' END AS [Tình tr?ng PO],
        CASE WHEN ISNULL(st.TotalItems,0) = 0 THEN 0 ELSE CAST(ISNULL(o.OrderedItems,0) * 100.0 / st.TotalItems AS DECIMAL(5,1)) END AS [% Item ??t hàng],
        f.Created_Date AS [Ngày t?o],
        f.BaseNo,
        ISNULL(m.MaxRev, 0) AS MaxRev
    FROM Filtered f
    LEFT JOIN Ordered o ON o.BaseNo = f.BaseNo
    LEFT JOIN SeriesTotals st ON st.BaseNo = f.BaseNo
    LEFT JOIN MaxRev m ON m.BaseNo = f.BaseNo
    ORDER BY f.Created_Date DESC;
END
GO
