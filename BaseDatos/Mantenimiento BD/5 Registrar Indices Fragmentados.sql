USE [FRONTPLAZA]
GO

/****** Object:  StoredProcedure [dbo].[INDICES_FRAGMENTADOS]    Script Date: 27/12/2019 00:32:14 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[INDICES_FRAGMENTADOS]
@FECHA DATE, @ORIGEN VARCHAR(1)

AS


WITH INDICES(FECHA, BD, INDICETIPO, FRAGMENTACION, INDICE, TABLA) AS (SELECT @FECHA AS FECHA, DBS.name AS BASEDEDATOS, PS.INDEX_TYPE_DESC, PS.AVG_FRAGMENTATION_IN_PERCENT, IND.name AS INDICE, 
                                                                                                                                                                                 TAB.name AS TABLA
                                                                                                                                                               FROM      SYS.DM_DB_INDEX_PHYSICAL_STATS(DB_ID(), NULL, NULL, NULL, NULL) AS PS INNER JOIN
                                                                                                                                                                                 sys.databases AS DBS ON PS.DATABASE_ID = DBS.database_id INNER JOIN
                                                                                                                                                                                 sys.indexes AS IND ON PS.OBJECT_ID = IND.object_id INNER JOIN
                                                                                                                                                                                 sys.tables AS TAB ON TAB.object_id = IND.object_id
                                                                                                                                                               WHERE   (IND.name IS NOT NULL) AND (PS.INDEX_ID = IND.index_id) AND (PS.AVG_FRAGMENTATION_IN_PERCENT > 0))
    INSERT  
    INTO        TN_INDICES_FRAGMENTADOS(FECHA,QUERY, FRAGMENTACION, BD, INDICE, TABLA,ORIGEN)
                          SELECT DISTINCT 
                                            @FECHA AS FECHA, CASE WHEN FRAGMENTACION > 5 AND 
                                            FRAGMENTACION <= 30 THEN 'ALTER INDEX ' + INDICE + ' ON ' + TABLA + ' REORGANIZE' WHEN FRAGMENTACION > 30 THEN 'ALTER INDEX ' + INDICE + ' ON ' + TABLA + ' REBUILD' END AS QUERY, FRAGMENTACION, BD, 
                                            INDICE, TABLA, @ORIGEN
                          FROM     (SELECT FRAGMENTACION, INDICE, TABLA, BD
                                            FROM      INDICES AS INDICES_1
                                            WHERE   (FRAGMENTACION > 5)) AS A
                          ORDER BY FRAGMENTACION DESC
-------------------------------------------------------------------


GO


