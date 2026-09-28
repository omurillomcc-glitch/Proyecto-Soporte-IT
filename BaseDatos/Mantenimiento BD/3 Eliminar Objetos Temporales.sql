USE [SANTODOMINGO]
GO

/****** Object:  StoredProcedure [dbo].[ICGCOL_DELETE_TABLES]    Script Date: 27/12/2019 00:34:28 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


create PROCEDURE [dbo].[ICGCOL_DELETE_TABLES]
AS

DROP TABLE ICGCOL_TABLAS_BASURA

--CREACION DE TABLA CONTENIENTES CON TABLAS CREADAS POR ICG
CREATE TABLE [dbo].[ICGCOL_TABLAS_BASURA](
	[TableName] [nvarchar](255) NULL,
	[create_date] [datetime] NULL,
	[modify_date] [datetime] NULL,
	[SchemaName] [nvarchar](255) NULL,
	[lock_escalation_desc] [nvarchar](255) NULL,
	[RowCount] [float] NULL,
	[TotalSpaceKB] [float] NULL,
	[UsedSpaceKB] [float] NULL,
	[UnusedSpaceKB] [float] NULL
) ON [PRIMARY]


--ALMACENO LOS REGISTROS CON TABLAS BASURA
INSERT INTO ICGCOL_TABLAS_BASURA ([TableName],[create_date],[modify_date],[SchemaName],[lock_escalation_desc],[RowCount],[TotalSpaceKB],
[UsedSpaceKB],[UnusedSpaceKB] )

SELECT 
    t.NAME AS TableName,
	t.create_date,
	t.modify_date,
    s.Name AS SchemaName,
	t.lock_escalation_desc,
    p.rows AS RowCounts,
    SUM(a.total_pages) * 8 AS TotalSpaceKB, 
    SUM(a.used_pages) * 8 AS UsedSpaceKB, 
    (SUM(a.total_pages) - SUM(a.used_pages)) * 8 AS UnusedSpaceKB
FROM 
    sys.tables t
INNER JOIN      
    sys.indexes i ON t.OBJECT_ID = i.object_id
INNER JOIN 
    sys.partitions p ON i.object_id = p.OBJECT_ID AND i.index_id = p.index_id
INNER JOIN 
    sys.allocation_units a ON p.partition_id = a.container_id
LEFT OUTER JOIN 
    sys.schemas s ON t.schema_id = s.schema_id
WHERE 
    t.NAME NOT LIKE 'dt%' 
    AND t.is_ms_shipped = 0
    AND i.OBJECT_ID > 255 
	and t.name like 'TEMP_INVENTARIO_%' or 
	t.name like 'ICGMAIL_CALCULOS%' or 
	t.name like 'ATEMPIL%' or 
	t.name like 'ATEMPARTICSLIN%' or 
	t.name like 'ATABLATEMPTIPOSDOC%' or
	t.name like '%TEMPCODARTICULO%' or 
	t.name like '%{%'

GROUP BY 
    t.Name, s.Name, p.Rows,t.create_date,t.modify_date,t.lock_escalation_desc
ORDER BY 
    t.Name


SELECT * FROM ICGCOL_TABLAS_BASURA

---CURSOR DE ELIMINACION TABLAS ENCONTRADAS CON REGISTROS MENORES A 8 DIAS
DECLARE @QUERY AS NVARCHAR(255)
DECLARE @TABLE AS NVARCHAR(255)
DECLARE @StrSql AS NVARCHAR(255)

--SENTENCIA SQL A EJECUTAR
SET @QUERY='DROP TABLE [dbo].['

DECLARE CR_TABLES CURSOR LOCAL FOR

	SELECT TableName as Modified_Days FROM ICGCOL_TABLAS_BASURA 
	WHERE DATEDIFF(day,modify_date,GETDATE())>8

OPEN CR_TABLES
FETCH NEXT FROM CR_TABLES INTO @TABLE;

WHILE @@FETCH_STATUS=0
BEGIN

	SET @StrSql=@QUERY+@TABLE +']'
	--PRINT @StrSql
	EXECUTE sp_executesql @StrSql

FETCH NEXT FROM CR_TABLES INTO @TABLE;

END
CLOSE CR_TABLES
DEALLOCATE CR_TABLES
GO


