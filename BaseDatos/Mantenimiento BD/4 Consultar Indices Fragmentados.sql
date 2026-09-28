USE [SANTODOMINGO]
GO

/****** Object:  StoredProcedure [dbo].[CONSULTA_INDICES_FRAGMENTADOS]    Script Date: 27/12/2019 00:32:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE procedure [dbo].[CONSULTA_INDICES_FRAGMENTADOS]
AS
WITH INDICES(BD, INDICETIPO, FRAGMENTACION, INDICE, TABLA) AS (SELECT DBS.name AS BASEDEDATOS, PS.INDEX_TYPE_DESC, PS.AVG_FRAGMENTATION_IN_PERCENT, IND.name AS INDICE, 
                                                                                                                                                                                 TAB.name AS TABLA
                                                                                                                                                               FROM      SYS.DM_DB_INDEX_PHYSICAL_STATS(DB_ID(), NULL, NULL, NULL, NULL) AS PS INNER JOIN
                                                                                                                                                                                 sys.databases AS DBS ON PS.DATABASE_ID = DBS.database_id INNER JOIN
                                                                                                                                                                                 sys.indexes AS IND ON PS.OBJECT_ID = IND.object_id INNER JOIN
                                                                                                                                                                                 sys.tables AS TAB ON TAB.object_id = IND.object_id
                                                                                                                                                               WHERE   (IND.name IS NOT NULL) AND (PS.INDEX_ID = IND.index_id) AND (PS.AVG_FRAGMENTATION_IN_PERCENT > 0))

                          SELECT DISTINCT 
                                            CASE WHEN FRAGMENTACION > 5 AND 
                                            FRAGMENTACION <= 30 THEN 'ALTER INDEX ' + INDICE + ' ON ' + TABLA + ' REORGANIZE' WHEN FRAGMENTACION > 30 THEN 'ALTER INDEX ' + INDICE + ' ON ' + TABLA + ' REBUILD' END AS QUERY, FRAGMENTACION, BD, 
                                            INDICE, TABLA
                          FROM     (SELECT FRAGMENTACION, INDICE, TABLA, BD
                                            FROM      INDICES AS INDICES_1
                                            WHERE   (FRAGMENTACION > 5)) AS A
                          ORDER BY FRAGMENTACION DESC




GO


