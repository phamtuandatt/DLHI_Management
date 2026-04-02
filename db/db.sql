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
CREATE PROCEDURE [dbo].[sp_InsertProduct]
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

    -- Trả về ID vừa tạo để sử dụng ở C# nếu cần
    SELECT SCOPE_IDENTITY() AS NewID;
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

-- GROUP ID_CODE TO SHOW QTY
	--SELECT Item_Name, Material, Size, ID_Code, COUNT(Item_Name), SUM(Qty_Import)
	--FROM Warehouse_Import
	--WHERE Project_Code = '2508-DPCII'
	--GROUP BY Item_Name, Material, Size, ID_Code

 SELECT *FROM Products
SELECT *FROM Material_Detail
SELECT * FROM Warehouse_Import WHERE Project_Code = '25G3-NGR'
SELECT *FROM ProjectInfo
SELECT *FROM PO_head WHERE ProjectCode = '25G3-NGR' OR Notes = '25G3-NGR'
UPDATE PO_head SET ProjectCode = '25G3-NGR' WHERE Notes = '25G3-NGR'
SELECT *FROM MPR_Header
