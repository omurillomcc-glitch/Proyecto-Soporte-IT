SELECT QUOTENAME(SCHEMA_NAME([schema_id])) + N'.' + QUOTENAME([name])
FROM sys.tables
WHERE [name] LIKE 'KITS_SIN_STOCK_SEG_ALM%';
GO

DECLARE @sql NVARCHAR(MAX);
DECLARE @tn NVARCHAR(256);
DECLARE @c CURSOR;

SET @c = CURSOR LOCAL FAST_FORWARD
FOR
SELECT QUOTENAME(SCHEMA_NAME([schema_id])) + N'.' + QUOTENAME([name])
FROM sys.tables
WHERE [name] LIKE 'KITS_SIN_STOCK_SEG_ALM%';

OPEN @c;

WHILE 1 = 1
BEGIN
	FETCH NEXT FROM @c INTO  @tn;
	
	IF @@ERROR <> 0 OR @@FETCH_STATUS <> 0
		BREAK;
	
	SET @sql = N'drop table ' + @tn;
	
	PRINT @sql;
	EXEC sp_executesql @sql;	
END
GO

SELECT QUOTENAME(SCHEMA_NAME([schema_id])) + N'.' + QUOTENAME([name])
FROM sys.tables
WHERE [name] LIKE 'KITS_SIN_STOCK_SEG_ALM%';
GO
