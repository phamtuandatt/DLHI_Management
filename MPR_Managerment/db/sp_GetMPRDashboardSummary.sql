-- Stored procedure: MPR Dashboard Summary
-- % Item đặt hàng = ordered items / total items cho revision CỤ THỂ đang hiển thị
-- Ordered = có PO trực tiếp (PO_Detail.MPR_Detail_ID) HOẶC qua revision cũ (same Item_No + Item_Name + Material)
-- Không đếm dòng Is_Deleted=1, không đếm PO Cancelled

IF OBJECT_ID('dbo.sp_GetMPRDashboardSummary', 'P') IS NOT NULL
    DROP PROCEDURE dbo.sp_GetMPRDashboardSummary;
GO

CREATE PROCEDURE dbo.sp_GetMPRDashboardSummary
    @search NVARCHAR(4000) = NULL,
    @status NVARCHAR(100)  = NULL
AS
BEGIN
    SET NOCOUNT ON;

    -- ── 1. Tất cả header + BaseNo ──────────────────────────────────────────────
    WITH Headers AS (
        SELECT
            MPR_ID,
            MPR_No,
            Project_Name,
            Required_Date,
            Status,
            ISNULL(Notes, N'')                                   AS Notes,
            ISNULL(TRY_CAST(TRY_CAST(Rev AS DECIMAL(10,2)) AS INT), 0) AS Rev,
            Created_Date,
            CASE WHEN CHARINDEX(N'_Rev.', MPR_No) > 0
                 THEN LEFT(MPR_No, CHARINDEX(N'_Rev.', MPR_No) - 1)
                 ELSE MPR_No
            END AS BaseNo
        FROM MPR_Header
    ),

    -- ── 2. MaxRev mỗi series ───────────────────────────────────────────────────
    MaxRevPerBase AS (
        SELECT BaseNo, MAX(Rev) AS MaxRev
        FROM Headers
        GROUP BY BaseNo
    ),

    -- ── 3a. PO trực tiếp: PO_Detail.MPR_Detail_ID → Detail_ID của MPR này ─────
    OrderedDirect AS (
        SELECT
            md.MPR_ID,
            md.Detail_ID AS DetailID,
            po.PO_ID
        FROM MPR_Details md
        INNER JOIN PO_Detail pod ON pod.MPR_Detail_ID = md.Detail_ID
        INNER JOIN PO_head   po  ON po.PO_ID  = pod.PO_ID
                                 AND ISNULL(po.Status, N'') <> N'Cancelled'
        WHERE ISNULL(md.Is_Deleted, 0) = 0
    ),

    -- ── 3b. PO qua revision cũ: cur và old cùng Item_No + Item_Name + Material ─
    OrderedViaOldRev AS (
        SELECT
            cur.MPR_ID,
            cur.Detail_ID AS DetailID,
            po.PO_ID
        FROM MPR_Details cur
        INNER JOIN Headers curH ON curH.MPR_ID = cur.MPR_ID
        -- tìm tất cả revision khác trong cùng series
        INNER JOIN Headers oldH ON (oldH.MPR_No = curH.BaseNo OR oldH.MPR_No LIKE curH.BaseNo + N'_Rev.%')
                                AND oldH.MPR_ID <> curH.MPR_ID
        -- match vật tư theo Item_No + Item_Name + Material (tránh nhầm khi revision đổi thứ tự)
        INNER JOIN MPR_Details old ON old.MPR_ID  = oldH.MPR_ID
                                   AND ISNULL(old.Is_Deleted, 0) = 0
                                   AND old.Item_No  = cur.Item_No
                                   AND ISNULL(old.Item_Name, N'') = ISNULL(cur.Item_Name, N'')
                                   AND ISNULL(old.Material,  N'') = ISNULL(cur.Material,  N'')
        INNER JOIN PO_Detail pod ON pod.MPR_Detail_ID = old.Detail_ID
        INNER JOIN PO_head   po  ON po.PO_ID  = pod.PO_ID
                                 AND ISNULL(po.Status, N'') <> N'Cancelled'
        WHERE ISNULL(cur.Is_Deleted, 0) = 0
    ),

    -- ── 4. Gộp tất cả Detail_ID đã được đặt hàng ──────────────────────────────
    AllOrdered AS (
        SELECT MPR_ID, DetailID, PO_ID FROM OrderedDirect
        UNION
        SELECT MPR_ID, DetailID, PO_ID FROM OrderedViaOldRev
    ),

    -- ── 5. Thống kê per MPR_ID ─────────────────────────────────────────────────
    --    TotalItems  = số dòng KHÔNG bị xóa trong revision này
    --    OrderedItems = số dòng trong revision này đã có PO (direct hoặc cross-rev)
    --    POCount      = số PO phân biệt liên kết với revision này
    Stats AS (
        SELECT
            h.MPR_ID,
            COUNT(DISTINCT d.Detail_ID)   AS TotalItems,
            COUNT(DISTINCT ao.DetailID)   AS OrderedItems,
            COUNT(DISTINCT ao.PO_ID)      AS POCount
        FROM Headers h
        LEFT JOIN MPR_Details d  ON d.MPR_ID  = h.MPR_ID AND ISNULL(d.Is_Deleted, 0) = 0
        LEFT JOIN AllOrdered  ao ON ao.MPR_ID = h.MPR_ID
        GROUP BY h.MPR_ID
    )

    -- ── 6. Kết quả cuối ────────────────────────────────────────────────────────
    SELECT
        h.MPR_ID,
        h.MPR_No                                                           AS [MPR No],
        h.Project_Name                                                     AS [Dự án],
        h.Required_Date                                                    AS [Ngày cần],
        h.Status                                                           AS [Trạng thái],
        h.Rev                                                              AS [Rev],
        h.BaseNo,
        ISNULL(mr.MaxRev, 0)                                               AS MaxRev,
        CASE
            WHEN ISNULL(s.POCount, 0) > 0
            THEN N'✅ ' + CAST(s.POCount AS NVARCHAR(10)) + N' PO'
            ELSE N'❌ Chưa có PO'
        END                                                                AS [Tình trạng PO],
        CASE
            WHEN ISNULL(s.TotalItems, 0) = 0 THEN CAST(0 AS DECIMAL(5,1))
            ELSE CAST(ISNULL(s.OrderedItems, 0) * 100.0 / s.TotalItems AS DECIMAL(5,1))
        END                                                                AS [% Item đặt hàng],
        h.Created_Date                                                     AS [Ngày tạo],
        h.Notes                                                            AS [Ghi chu]
    FROM Headers h
    LEFT JOIN Stats          s  ON s.MPR_ID  = h.MPR_ID
    INNER JOIN MaxRevPerBase mr ON mr.BaseNo  = h.BaseNo
    WHERE
        (@search IS NULL OR @search = N'' OR h.MPR_No LIKE N'%' + @search + N'%' OR h.Project_Name LIKE N'%' + @search + N'%')
        AND (@status IS NULL OR @status = N'' OR h.Status = @status)
    ORDER BY h.Created_Date DESC;
END
GO
