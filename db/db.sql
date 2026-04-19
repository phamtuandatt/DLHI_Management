------------------------------------------------------------------------------------------------------------------------------------------------------
--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
----------------------------------------------------------------START MODULE PRODUCT------------------------------------------------------------------
--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
---------------------------------------------------------------Đã tạo bảng----------------------------------------------------------------------------
-- Unit
CREATE TABLE Units (
	id			INT PRIMARY KEY,
	name		NVARCHAR(50),
	unit_price	DECIMAL(18, 2),
)

-- Origins
CREATE TABLE Origins (
	id		INT IDENTITY(1,1) PRIMARY KEY,
	code	VARCHAR(20),
	name	NVARCHAR(MAX),
)

-- Standard
CREATE TABLE Standards (
	id		INT IDENTITY(1,1) PRIMARY KEY,
	code	VARCHAR(20),
	name	NVARCHAR(225)
)

-- Material Categories - Danh mục vật tư (Main, Fitting,...)
CREATE TABLE Material_Categories (
    cat_id		INT IDENTITY(1,1) PRIMARY KEY,
    cat_name	NVARCHAR(100) NOT NULL,
);

-- Bảng vật tư
CREATE TABLE Materials (
    material_id		INT IDENTITY(1,1) PRIMARY KEY,
    material_code	VARCHAR(50) NOT NULL UNIQUE,
    material_name	NVARCHAR(255),
    specifications	NVARCHAR(MAX),
    created_at  	DATETIME DEFAULT GETDATE(),

    cat_id			INT NOT NULL,
    unit_id			INT,
    CONSTRAINT FK_Material_Category FOREIGN KEY (cat_id) REFERENCES Material_Categories(cat_id),
);
-- Chi tiết vật tư 
CREATE TABLE Material_Detail (
	material_detail_id		INT IDENTITY(1,1) PRIMARY KEY,
	material_detail_number	VARCHAR(5),
	material_detail_name	NVARCHAR(100),
	material_detail_code	VARCHAR(50),
	item_code_existed		VARCHAR(50)

	CONSTRAINT FK_Material_Detail_Materials FOREIGN KEY (material_detail_id) REFERENCES Materials(material_id)
)

-- Bảng sản phẩm
CREATE TABLE Products (
	id					INT IDENTITY(1,1) PRIMARY KEY,
	name				NVARCHAR(50),
	des_2				NVARCHAR(50),
	code				VARCHAR(50),
	prod_material_code	VARCHAR(50), -- Standard name
	picture_link		NVARCHAR(200),
	picture				VARBINARY(MAX),
	a_thinkness			VARCHAR(10),
	b_depth				VARCHAR(10),
	c_witdth			VARCHAR(10),
	d_web				VARCHAR(10),
	e_flag				VARCHAR(10),
	f_length			VARCHAR(10),
	g_weight			VARCHAR(10),
	used_note			NVARCHAR(100),
	 
	-- Mở rộng -> Tạo khóa ngoại
	prod_origin_id			INT DEFAULT NULL,	-- Origin: DO (trong nước), P1 (Nhập miễn thuế),...
	prod_standard_id		INT DEFAULT NULL,	-- Standard: 000, 0001,... (A363, A572,...)
	prod_material_cate_id	INT DEFAULT NULL,	-- Material_Categories: Main, Fitting,...
	prod_material_id		INT DEFAULT NULL,	-- Material: Plate, Beam,...
	prod_material_detail_id	INT DEFAULT NULL,	-- Material_detail: Plate dày, plate mỏng,...
)
------------------------------------------------------------------------------------------------------------------------------------------------------
--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
----------------------------------------------------------------END MODULE PRODUCT--------------------------------------------------------------------
--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------

GO
CREATE PROCEDURE GET_ITEM_NUMBER_OF_MATERIAL_TYPE @MMATERIAL_ID INT
AS
BEGIN
    DECLARE @MAX_ITEM_NUMBER NVARCHAR(10)
    DECLARE @NEXT_ITEM_NUMBER NVARCHAR(10)

    -- Get the maximum item_number for the given type_id
    SELECT @MAX_ITEM_NUMBER = MAX(material_detail_number)
    FROM Material_Detail
    WHERE material_detail_code = @MMATERIAL_ID

    -- If no data exists for the type_id, set the next value to '001'
    IF @MAX_ITEM_NUMBER IS NULL
    BEGIN
        SET @NEXT_ITEM_NUMBER = '001'
    END
    ELSE
    BEGIN
        -- Calculate the next item number by incrementing the max item number
        SET @NEXT_ITEM_NUMBER = RIGHT('000' + CAST(CAST(@MAX_ITEM_NUMBER AS INT) + 1 AS NVARCHAR), 3)
    END

    -- Return the next item number
    SELECT @NEXT_ITEM_NUMBER AS NEXT_ITEM_NUMBER
END

--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------
CREATE PROCEDURE [dbo].[sp_InsertProductWithCheck]
    @name NVARCHAR(255),
    @des_2 NVARCHAR(MAX),
    @code NVARCHAR(100), -- Đây là ProductCode dùng để kiểm tra trùng
    @prod_material_code NVARCHAR(100),
    @a_thinkness VARCHAR(50),
    @b_depth VARCHAR(50),
    @c_witdth VARCHAR(50),
    @d_web VARCHAR(50),
    @e_flag VARCHAR(50),
    @f_length VARCHAR(50),
    @g_weight VARCHAR(50),
    @used_note NVARCHAR(MAX),
    @prod_origin_id INT,
    @prod_standard_id INT,
    @prod_material_cate_id INT,
    @prod_material_id INT,
    @prod_material_detail_id INT
AS
BEGIN
    SET NOCOUNT ON;

    -- Kiểm tra nếu Code chưa tồn tại trong bảng Products
    IF NOT EXISTS (SELECT 1 FROM [dbo].[Products] WHERE [code] = @code)
    BEGIN
        INSERT INTO [dbo].[Products] (
            [name], [des_2], [code], [prod_material_code],
            [a_thinkness], [b_depth], [c_witdth], [d_web], [e_flag], [f_length], [g_weight],
            [used_note], [prod_origin_id], [prod_standard_id], 
            [prod_material_cate_id], [prod_material_id], [prod_material_detail_id]
        )
        VALUES (
            @name, @des_2, @code, @prod_material_code,
            @a_thinkness, @b_depth, @c_witdth, @d_web, @e_flag, @f_length, @g_weight,
            @used_note, @prod_origin_id, @prod_standard_id, 
            @prod_material_cate_id, @prod_material_id, @prod_material_detail_id
        );

        -- Trả về ID vừa tạo
        SELECT SCOPE_IDENTITY() AS ResultStatus; -- Trả về ID dương là insert thành công
    END
    ELSE
    BEGIN
        -- Trả về 0 hoặc -1 để báo hiệu ở code C# rằng dữ liệu đã tồn tại
        SELECT 0 AS ResultStatus;
    END
END
GO

--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------
GO
CREATE PROCEDURE [dbo].[sp_UpdateProduct]
	@id INT, -- Cần ID để xác định dòng cần sửa
    @name NVARCHAR(255),
    @des_2 NVARCHAR(MAX),
    @code NVARCHAR(100),
    @prod_material_code NVARCHAR(100),
    @a_thinkness VARCHAR(50),
    @b_depth VARCHAR(50),
    @c_witdth VARCHAR(50),
    @d_web VARCHAR(50),
    @e_flag VARCHAR(50),
    @f_length VARCHAR(50),
    @g_weight VARCHAR(50),
    @used_note NVARCHAR(MAX),
    @prod_origin_id INT,
    @prod_standard_id INT,
    @prod_material_cate_id INT,
    @prod_material_id INT,
    @prod_material_detail_id INT
AS
BEGIN
    SET NOCOUNT ON;

    UPDATE [dbo].[Products]
    SET 
		[name] = @name,
        [des_2] = @des_2,
        [code] = @code,
        [prod_material_code] = @prod_material_code,
        [a_thinkness] = @a_thinkness,
        [b_depth] = @b_depth,
        [c_witdth] = @c_witdth,
        [d_web] = @d_web,
        [e_flag] = @e_flag,
        [f_length] = @f_length,
        [g_weight] = @g_weight,
        [used_note] = @used_note,
        [prod_origin_id] = @prod_origin_id,
        [prod_standard_id] = @prod_standard_id,
        [prod_material_cate_id] = @prod_material_cate_id,
        [prod_material_id] = @prod_material_id,
        [prod_material_detail_id] = @prod_material_detail_id
    WHERE [id] = @id;
END
GO

--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Tạo trigger để tự động cập nhật Amount khi Price hoặc Qty thay đổi
ALTER TRIGGER [dbo].[trg_PO_Detail_UpdateAmount]
ON [dbo].[PO_Detail]
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    
    UPDATE d
    SET d.Amount = d.Qty_Per_Sheet * d.Price * (1 + d.VAT/100)
    FROM dbo.PO_Detail d
    INNER JOIN inserted i ON d.PO_Detail_ID = i.PO_Detail_ID;
    
    -- Cập nhật Total_Amount trong PO_head
    UPDATE h
    SET h.Total_Amount = (
        SELECT ISNULL(SUM(Amount), 0)
        FROM dbo.PO_Detail
        WHERE PO_ID = h.PO_ID
    )
    FROM dbo.PO_head h
    WHERE h.PO_ID IN (SELECT PO_ID FROM inserted);
END

--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------
GO
CREATE PROCEDURE [dbo].[sp_InsertWarehouseExport]
    @Export_No NVARCHAR(50),
    @Export_Date DATETIME,
    @Import_ID INT,
    @Item_Name NVARCHAR(255),
    @Material NVARCHAR(255),
    @Size NVARCHAR(255),
    @UNIT NVARCHAR(50),
    @Qty_Export DECIMAL(18, 2),
    @Weight_kg DECIMAL(18, 2),
    @ID_Code NVARCHAR(100),
    @Project_Code NVARCHAR(100),
    @WorkorderNo NVARCHAR(100),
    @Export_To NVARCHAR(255),
    @Purpose NVARCHAR(MAX),
    @Notes NVARCHAR(MAX),
    @Created_By NVARCHAR(100),
    @Warehouse_ID INT
AS
BEGIN
    SET NOCOUNT ON;
    
    -- Bắt đầu giao dịch để đảm bảo an toàn dữ liệu
    BEGIN TRANSACTION;

    BEGIN TRY
        -- 1. Insert vào bảng Warehouse_Export
        INSERT INTO [dbo].[Warehouse_Export] (
            [Export_No], [Export_Date], [Import_ID], [Item_Name], [Material], [Size], 
            [UNIT], [Qty_Export], [Weight_kg], [ID_Code], [Project_Code], 
            [WorkorderNo], [Export_To], [Purpose], [Notes], [Created_By], 
            [Created_Date], [Warehouse_ID]
        )
        VALUES (
            @Export_No, @Export_Date, @Import_ID, @Item_Name, @Material, @Size, 
            @UNIT, @Qty_Export, @Weight_kg, @ID_Code, @Project_Code, 
            @WorkorderNo, @Export_To, @Purpose, @Notes, @Created_By, 
            GETDATE(), @Warehouse_ID
        );

        -- 2. Cập nhật giảm số lượng ở bảng Warehouse_Import
        -- Trừ Qty_Import dựa trên Import_ID được truyền vào
        UPDATE [dbo].[Warehouse_Import]
        SET [Qty_Import] = [Qty_Import] - @Qty_Export
        WHERE [Import_ID] = @Import_ID;

        -- Kiểm tra nếu số lượng sau khi trừ bị âm (Tùy chọn nghiệp vụ)
        IF EXISTS (SELECT 1 FROM [dbo].[Warehouse_Import] WHERE [Import_ID] = @Import_ID AND [Qty_Import] < 0)
        BEGIN
            RAISERROR('Lỗi: Số lượng xuất vượt quá số lượng tồn kho hiện tại!', 16, 1);
        END

        -- Nếu mọi thứ ổn, xác nhận thay đổi
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        -- Nếu có bất kỳ lỗi nào, hủy bỏ toàn bộ quá trình
        ROLLBACK TRANSACTION;
        
        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@ErrorMessage, 16, 1);
    END CATCH
END
GO

CREATE PROCEDURE [dbo].[sp_InsertProductFull]
    -- Tham số cho bảng Products
    @name NVARCHAR(255),
    @des_2 NVARCHAR(MAX),
    @code NVARCHAR(100),
    @prod_material_code NVARCHAR(100),
    @a_thinkness VARCHAR(50),
    @b_depth VARCHAR(50),
    @c_witdth VARCHAR(50),
    @d_web VARCHAR(50),
    @e_flag VARCHAR(50),
    @f_length VARCHAR(50),
    @g_weight VARCHAR(50),
    @used_note NVARCHAR(MAX),
    @prod_origin_id INT,
    @prod_standard_id INT,
    @prod_material_cate_id INT,
    @prod_material_id INT,
    @prod_material_detail_id INT,
    
    -- Tham số bổ sung cho bảng Material_Detail
    @mat_detail_number NVARCHAR(100),
    @mat_detail_name NVARCHAR(255)
AS
BEGIN
    SET NOCOUNT ON;
    
    BEGIN TRANSACTION;
    BEGIN TRY
        -- 1. Chèn vào bảng Products
        INSERT INTO [dbo].[Products] (
            [name], [des_2], [code], [prod_material_code],
            [a_thinkness], [b_depth], [c_witdth], [d_web], [e_flag], [f_length], [g_weight],
            [used_note], [prod_origin_id], [prod_standard_id], 
            [prod_material_cate_id], [prod_material_id], [prod_material_detail_id]
        )
        VALUES (
            @name, @des_2, @code, @prod_material_code,
            @a_thinkness, @b_depth, @c_witdth, @d_web, @e_flag, @f_length, @g_weight,
            @used_note, @prod_origin_id, @prod_standard_id, 
            @prod_material_cate_id, @prod_material_id, @prod_material_detail_id
        );

        -- Lấy ID sản phẩm vừa tạo (nếu cần dùng làm code hoặc mapping)
        DECLARE @NewProductID INT = SCOPE_IDENTITY();

        -- 2. Chèn vào bảng Material_Detail
        -- Sử dụng @code của Product làm item_code_existed
        INSERT INTO Material_Detail (
            material_detail_number, 
            material_detail_name, 
            material_detail_code, 
            item_code_existed
        ) 
        VALUES (
            @mat_detail_number, 
            @mat_detail_name, 
            @prod_material_code, -- material_detail_code lấy theo mã vật liệu
            @code                -- item_code_existed lấy theo mã sản phẩm
        );

        COMMIT TRANSACTION;

        -- Trả về ID sản phẩm để ứng dụng C# nhận biết thành công
        SELECT @NewProductID AS NewID;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        THROW;
    END CATCH
END

--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------
CREATE PROCEDURE [dbo].[sp_InsertRIRDetail_UpdateStock]
    @RIR_ID INT,
    @PO_Detail_ID INT,
    @Item_No NVARCHAR(100),
    @item_name NVARCHAR(255),
    @Material NVARCHAR(255),
    @Size NVARCHAR(255),
    @UNIT NVARCHAR(50),
    @Qty_Per_Sheet DECIMAL(18, 2),
    @MTRno NVARCHAR(100),
    @Heatno NVARCHAR(100),
    @Qty_Required DECIMAL(18, 2),
    @Qty_Received DECIMAL(18, 2),
    @Inspect_Result NVARCHAR(100),
    @ID_Code NVARCHAR(100) -- Mã định danh RIR dùng để cập nhật vào kho
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRANSACTION;
    BEGIN TRY
        -- 1. Insert vào bảng RIR_Detail
        INSERT INTO [dbo].[RIR_Detail] (
            [RIR_ID], 
            [PO_Detail_ID], 
            [Item_No], 
            [item_name], 
            [Material], 
            [Size], 
            [UNIT], 
            [Qty_Per_Sheet], 
            [MTRno], 
            [Heatno], 
            [Created_Date], 
            [Qty_Required], 
            [Qty_Received], 
            [Inspect_Result], 
            [ID_Code]
        )
        VALUES (
            @RIR_ID, 
            @PO_Detail_ID, 
            @Item_No, 
            @item_name, 
            @Material, 
            @Size, 
            @UNIT, 
            @Qty_Per_Sheet, 
            @MTRno, 
            @Heatno, 
            GETDATE(), 
            @Qty_Required, 
            @Qty_Received, 
            @Inspect_Result, 
            @ID_Code
        );

        -- 2. Cập nhật bảng Warehouse_Import
        -- Gán QC_Code = ID_Code của RIR cho những dòng có cùng PO_Detail_ID
        UPDATE [dbo].[Warehouse_Import]
        SET [QC_Code] = @ID_Code
        WHERE [PO_Detail_ID] = @PO_Detail_ID;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        
        DECLARE @Err NVARCHAR(MAX) = ERROR_MESSAGE();
        RAISERROR(@Err, 16, 1);
    END CATCH
END
GO
--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------
  -- LẤY THÔNG TIN CHO KẾ TOÁN
-- Chạy đoạn code này trong SQL Server Management Studio (SSMS)
CREATE PROCEDURE [dbo].[GetSalesData]
AS
BEGIN
    SET NOCOUNT ON;
	SELECT 
		W.InvoiceNo, 
		W.InvoiceDate, 
		W.ID_Code, 
		W.Item_Name, 
		W.Size, 
		P.ProjectCode, 
		S.Company_Name
	FROM Warehouse_Import AS W
	INNER JOIN PO_head AS P ON W.PO_ID = P.PO_ID
	INNER JOIN Suppliers AS S ON P.Supplier_ID = S.Supplier_ID;
END
GO


  --xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------
/****** Object:  StoredProcedure [dbo].[GetSalesData]    Script Date: 07/04/2026 9:05:07 PM ******/
CREATE PROCEDURE [dbo].[sp_GetDataToFillInvoce] @POID INT
AS
BEGIN
    SET NOCOUNT ON;
	SELECT 
		W.InvoiceNo, 
		W.InvoiceDate, 
		W.ID_Code, 
		W.Item_Name, 
		W.Size, 
		P.ProjectCode, 
		S.Company_Name
	FROM Warehouse_Import AS W
	INNER JOIN PO_head AS P ON W.PO_ID = P.PO_ID
	INNER JOIN Suppliers AS S ON P.Supplier_ID = S.Supplier_ID
	WHERE P.PO_ID = @POID AND W.InvoiceNo IS NULL OR W.InvoiceNo < 0
END
--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------

CREATE PROCEDURE [dbo].[sp_UpdatePOHead_MakeImport]
    @PONO INT,
    @ImportDate DATE
AS
BEGIN
    SET NOCOUNT OFF;

    -- 1. Kiểm tra sự tồn tại của PO_ID
    IF NOT EXISTS (SELECT 1 FROM [dbo].[PO_head] WHERE [PONo] = @PONO)
    BEGIN
        RAISERROR(N'Lỗi: Không tìm thấy đơn hàng PO No: %d để cập nhật!', 16, 1, @PONO);
        RETURN;
    END

    BEGIN TRY
        BEGIN TRANSACTION;

        -- 2. Thực hiện cập nhật thông tin Header
        UPDATE [dbo].[PO_head]
        SET 
            [IS_Imported] = 1,
            [ImportedDate] = @ImportDate
        WHERE [PONo] = @PONO;

        COMMIT TRANSACTION;
        SELECT N'Cập nhật PO Header thành công!' AS Result;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        DECLARE @ErrMsg NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@ErrMsg, 16, 1);
    END CATCH
END

--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------

--CREATE PROCEDURE [dbo].[sp_UpdateWarehouseQty]
--    @Import_ID INT,             -- ID của dòng cần cập nhật
--    @Qty_Subtract DECIMAL(18, 2) -- Số lượng muốn trừ đi
--AS
--BEGIN
--    SET NOCOUNT ON;

--    -- Kiểm tra xem dòng có tồn tại không
--    IF EXISTS (SELECT 1 FROM [dbo].[Warehouse_Import] WHERE [Import_ID] = @Import_ID)
--    BEGIN
--        -- Thực hiện cập nhật số lượng
--        UPDATE [dbo].[Warehouse_Import]
--        SET [Qty_Import] = [Qty_Import] - @Qty_Subtract
--        WHERE [Import_ID] = @Import_ID;

--        -- Trả về kết quả để C# nhận biết thành công
--        SELECT 1 AS ResultStatus;
--    END
--    ELSE
--    BEGIN
--        -- Trả về 0 nếu không tìm thấy ID
--        SELECT 0 AS ResultStatus;
--    END
--END
--GO

--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------

SELECT W.ID_Code, W.Item_Name AS 'Name', M.material_detail_name AS 'Description', W.Size, SUM(W.Qty_import) AS [Qty (SUM)] FROM Warehouse_Import AS W LEFT JOIN Material_Detail AS M  ON LEFT(W.ID_Code, 9) = M.item_code_existed GROUP BY  W.ID_Code, W.Item_Name, M.material_detail_name, W.Size ORDER BY W.ID_Code; 


--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------

CREATE PROCEDURE [dbo].[sp_UpdateRIRDetail_Warehouse]
    @PO_Detail_ID INT,
    @Item_No NVARCHAR(100),
    @item_name NVARCHAR(255),
    @Material NVARCHAR(255),
    @Size NVARCHAR(255),
    @UNIT NVARCHAR(50),
    @Qty_Per_Sheet DECIMAL(18, 2),
    @MTRno NVARCHAR(100),
    @Heatno NVARCHAR(100),
    @Qty_Required DECIMAL(18, 2),
    @Qty_Received DECIMAL(18, 2),
    @Inspect_Result NVARCHAR(100),
    @ID_Code NVARCHAR(100),
	@RIR_Detail_ID INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRANSACTION;
    BEGIN TRY
        -- 1. Cập nhật vào bảng RIR_Detail thay vì Insert
        UPDATE [dbo].[RIR_Detail]
        SET 
            [Item_No] = @Item_No,
            [item_name] = @item_name,
            [Material] = @Material,
            [Size] = @Size,
            [UNIT] = @UNIT,
            [Qty_Per_Sheet] = @Qty_Per_Sheet,
            [MTRno] = @MTRno,
            [Heatno] = @Heatno,
            [Qty_Required] = @Qty_Required,
            [Qty_Received] = @Qty_Received,
            [Inspect_Result] = @Inspect_Result,
            [ID_Code] = @ID_Code,
            [Updated_Date] = GETDATE() -- Nên có cột này để theo dõi lịch sử chỉnh sửa
        WHERE RIR_Detail_ID = @RIR_Detail_ID

        -- 2. Cập nhật bảng Warehouse_Import (Giữ nguyên theo yêu cầu)
        UPDATE [dbo].[Warehouse_Import]
        SET [QC_Code] = @ID_Code, 
            [QC_Status] = @Inspect_Result
        WHERE [PO_Detail_ID] = @PO_Detail_ID;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;
        
        DECLARE @Err NVARCHAR(MAX) = ERROR_MESSAGE();
        RAISERROR(@Err, 16, 1);
    END CATCH
END

--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------

SELECT InvoiceNo, InvoiceNo, ID_Code, Item_Name, Size, Project_Code, Company_Name  FROM Warehouse_Import, Suppliers, PO_head WHERE Warehouse_Import.PO_ID = PO_head.PO_ID AND PO_head.Supplier_ID = Suppliers.Supplier_ID

--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------
 --=========================XÓA MPR================================
--SELECT *FROM MPR_Header WHERE MPR_No = 'DV-FT-2506 SAM-MPR-PD-102'
--SELECT *FROM MPR_Details WHERE MPR_ID = 132

--DELETE FROM MPR_Header WHERE MPR_ID = 132

--DELETE FROM MPR_Details WHERE Detail_ID = 1161
--DELETE FROM MPR_Details WHERE Detail_ID = 1162
--DELETE FROM MPR_Details WHERE Detail_ID = 1163
--DELETE FROM MPR_Details WHERE Detail_ID = 1164

--SELECT *FROM PO_head WHERE PONo = 'DV-SAM-PC-018'
--SELECT *FROM PO_Detail WHERE PO_ID = 356

--DELETE FROM PO_head WHERE PO_ID = 356

--DELETE FROM PO_Detail WHERE PO_Detail_ID = 2446
--DELETE FROM PO_Detail WHERE PO_Detail_ID = 2447
--DELETE FROM PO_Detail WHERE PO_Detail_ID = 2448
--DELETE FROM PO_Detail WHERE PO_Detail_ID = 2449

----SELECT *FROM PO_Revise_Transactions WHERE PO_ID = 356
----DELETE FROM PO_Revise_Transactions WHERE PO_ID = 356

----SELECT *FROM PO_Payment_Schedule WHERE PO_ID = 356
----DELETE FROM PO_Payment_Schedule WHERE Schedule_ID = 71
--=========================XÓA MPR================================

--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--
------------------------------------------------------------------------------------------------------------------------------------------------------