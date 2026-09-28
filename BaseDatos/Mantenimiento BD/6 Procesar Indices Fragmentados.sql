USE [SANTODOMINGO]
GO

/****** Object:  StoredProcedure [dbo].[PROCESAR_INDICES]    Script Date: 27/12/2019 00:32:20 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[PROCESAR_INDICES]
@FECHA DATE

AS

-- Declaracion de variables para el cursor

DECLARE 
	@QUERY varchar(1000),
	@FRAGMENTACION NUMERIC(15,13),
	@BD varchar(50),
	@INDICE varchar(500),
	@TABLA VARCHAR(100)


-- Declaración del cursor

DECLARE cINDICES CURSOR FOR

select QUERY, FRAGMENTACION, BD, INDICE, TABLA from TN_INDICES_FRAGMENTADOS
WHERE FECHA=@FECHA AND ORIGEN='A'
order by fragmentacion desc

-- Apertura del cursor

OPEN cINDICES

-- Lectura de la primera fila del cursor

FETCH cINDICES INTO   

	@QUERY,
	@FRAGMENTACION,
	@BD,
	@INDICE,
	@TABLA


WHILE (@@FETCH_STATUS = 0 )

BEGIN

	PRINT @QUERY

	EXECUTE(@QUERY) 

-- Lectura de la siguiente fila del cursor

FETCH cINDICES INTO    	
	
	@QUERY,
	@FRAGMENTACION,
	@BD,
	@INDICE,
	@TABLA

END

 

-- Cierre del cursor

CLOSE cINDICES

-- Liberar los recursos

DEALLOCATE cINDICES

--------------------------------------

GO


