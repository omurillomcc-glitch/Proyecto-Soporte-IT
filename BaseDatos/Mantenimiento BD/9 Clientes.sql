SELECT COUNT(*) FROM CLIENTES

SELECT * FROM CLIENTES
WHERE   CODCLIENTE  NOT IN (
SELECT DISTINCT CODCLIENTE  FROM ALBVENTACAB
UNION ALL
SELECT DISTINCT CODCLIENTE  FROM ANUL_ALBVENTACAB)

DELETE   FROM CLIENTES
WHERE   CODCLIENTE  NOT IN (
SELECT DISTINCT CODCLIENTE  FROM ALBVENTACAB
UNION ALL
SELECT DISTINCT CODCLIENTE  FROM ANUL_ALBVENTACAB)


select * from REM_TRANSACCIONES where TIPO = 12 and IDCENTRAL = -1

delete from REM_TRANSACCIONES where TIPO = 12 and IDCENTRAL = -1


SELECT c.codigo_cliente, c.nombre, f.fecha AS fecha_ultima_compra
FROM clientes c
JOIN (
    SELECT codigo_cliente, MAX(fecha) AS fecha
    FROM facturas
    GROUP BY codigo_cliente
) AS f_ultima_compra ON c.codigo_cliente = f_ultima_compra.codigo_cliente
JOIN facturas f ON f.codigo_cliente = c.codigo_cliente AND f.fecha = f_ultima_compra.fecha;

hola tengo dos tablas la tabla CLIENTES y la tabla FACTURASVENTA, necesito un informe que me muestre el código del cliente(CODCLINETE) , nombre del cliente(NOMBRECLIENTE) y la fecha de la ultima factura que compro(FECHA), el nombre del cliente esta en la tabla CLIENTES y la fecha de las facturas esta en la tabla FACTURASVENTA y estas dos tablas se relacionan por el campo CODCLIENTE
