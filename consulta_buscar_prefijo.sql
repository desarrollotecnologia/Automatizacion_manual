-- =============================================================================
-- FASE 5: vw_pbi02 no trae a VARGAS en propietario_producto.
-- Hay 5 vistas: vw_pbi01..vw_pbi05. Vamos a ver en cuál vive el cliente
-- Y el destino con prefijo /LxM/.
-- =============================================================================


-- -----------------------------------------------------------------------------
-- N) Qué propietarios distintos hay en vw_pbi02. ¿Aparece VARGAS de otra forma?
-- -----------------------------------------------------------------------------
SELECT DISTINCT propietario_producto
FROM trazabilidad_proceso.vw_pbi02
WHERE propietario_producto IS NOT NULL
ORDER BY propietario_producto
LIMIT 100;


-- -----------------------------------------------------------------------------
-- N2) Y VARGAS con algo más suelto (por si la ñ está como 'N' normal o vacía).
-- -----------------------------------------------------------------------------
SELECT DISTINCT propietario_producto
FROM trazabilidad_proceso.vw_pbi02
WHERE propietario_producto ILIKE '%VARGAS%'
   OR propietario_producto ILIKE '%YERSON%'
   OR propietario_producto ILIKE '%REYNALDO%';


-- -----------------------------------------------------------------------------
-- O) Columnas de vw_pbi03. ¿Trae codigo / nombre_propietario / destino?
-- -----------------------------------------------------------------------------
SELECT column_name, data_type
FROM information_schema.columns
WHERE table_schema = 'trazabilidad_proceso'
  AND table_name   = 'vw_pbi03'
ORDER BY ordinal_position;


-- -----------------------------------------------------------------------------
-- P) Columnas de vw_pbi04.
-- -----------------------------------------------------------------------------
SELECT column_name, data_type
FROM information_schema.columns
WHERE table_schema = 'trazabilidad_proceso'
  AND table_name   = 'vw_pbi04'
ORDER BY ordinal_position;


-- -----------------------------------------------------------------------------
-- Q) Columnas de vw_pbi05.
-- -----------------------------------------------------------------------------
SELECT column_name, data_type
FROM information_schema.columns
WHERE table_schema = 'trazabilidad_proceso'
  AND table_name   = 'vw_pbi05'
ORDER BY ordinal_position;


-- -----------------------------------------------------------------------------
-- R) Busca prefijos /LxM/, /MxM/, etc. en TODO vw_pbi03/04/05 (fila->texto).
--    El que devuelva filas es el bueno.
-- -----------------------------------------------------------------------------
SELECT 'vw_pbi03' AS fuente, (v.*)::text AS fila
FROM trazabilidad_proceso.vw_pbi03 v
WHERE (v.*)::text ~ '/(LxM|MxM|MxJ|JxV|VxS|SxD|DxL)/'
LIMIT 5;

SELECT 'vw_pbi04' AS fuente, (v.*)::text AS fila
FROM trazabilidad_proceso.vw_pbi04 v
WHERE (v.*)::text ~ '/(LxM|MxM|MxJ|JxV|VxS|SxD|DxL)/'
LIMIT 5;

SELECT 'vw_pbi05' AS fuente, (v.*)::text AS fila
FROM trazabilidad_proceso.vw_pbi05 v
WHERE (v.*)::text ~ '/(LxM|MxM|MxJ|JxV|VxS|SxD|DxL)/'
LIMIT 5;
