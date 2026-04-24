-- =============================================================================
-- CONSULTAS DE DIAGNOSTICO
-- Ejecutar de a UNA por vez en pgAdmin (seleccionar el bloque y F5).
-- Cronometra cada una (pgAdmin muestra el tiempo abajo) para saber donde se va.
-- =============================================================================


-- ========== TEST 1: solo vw_pbi01 filtrada por fecha (SIN propietario) ========
-- Si ESTO tarda >1 min = la vista es lenta por si sola (culpa del servidor).
-- Si tarda <10s = la vista responde bien; el cuello esta en el JOIN/ILIKE.
SELECT COUNT(*) AS n_filas
FROM trazabilidad_proceso.vw_pbi01 v
WHERE v.fecha_insensibilizacion >= '2026-04-13'::timestamp
  AND v.fecha_insensibilizacion <  '2026-04-14'::timestamp;


-- ========== TEST 2: vw_pbi01 filtrada por fecha + propietario ================
-- Cliente PEQUE: prueba con cualquier propietario con pocos animales.
-- Compara el tiempo con el TEST 1: si es parecido, el ILIKE no pesa.
SELECT v.codigo, v.nombre_propietario, v.nombre_parte, v.destino
FROM trazabilidad_proceso.vw_pbi01 v
WHERE v.fecha_insensibilizacion >= '2026-04-13'::timestamp
  AND v.fecha_insensibilizacion <  '2026-04-14'::timestamp
  AND v.nombre_propietario ILIKE '%VARGAS NI%O YERSON REYNALDO%'
LIMIT 200;


-- ========== TEST 3: cuantos propietarios distintos hay ese dia ===============
-- Para ver clientes chicos y probar con uno de 1-3 animales.
-- Puedes mirar la lista y elegir el de menos animales para acelerar.
SELECT v.nombre_propietario,
       COUNT(DISTINCT v.codigo) AS n_animales
FROM trazabilidad_proceso.vw_pbi01 v
WHERE v.fecha_insensibilizacion >= '2026-04-13'::timestamp
  AND v.fecha_insensibilizacion <  '2026-04-14'::timestamp
GROUP BY v.nombre_propietario
ORDER BY n_animales ASC
LIMIT 20;


-- ========== TEST 4: version SUPER reducida de la consulta principal ==========
-- Solo "base + destinos simple", sin partes (visceras/cabeza/patas/piel).
-- Si esta termina = el problema de la principal son las CTEs de partes/peso_ult.
-- Si esta tambien muere = el problema es de la replica, sin escape.
WITH params AS MATERIALIZED (
    SELECT
        '2026-04-13'::date                    AS fecha_plan,
        '%VARGAS NI%O YERSON REYNALDO%'::text AS propietario_like
),
base_animales AS MATERIALIZED (
    SELECT DISTINCT pfp.id_producto AS animal
    FROM trazabilidad_proceso.plan_faena pf
    JOIN trazabilidad_proceso.plan_faena_producto pfp
      ON pfp.id_plan_faena = pf.id
    JOIN params p ON pf.fecha_plan = p.fecha_plan
),
pbi AS MATERIALIZED (
    SELECT v.codigo, v.nombre_propietario, v.nombre_parte,
           v.destino, v.nombre_cava
    FROM trazabilidad_proceso.vw_pbi01 v
    JOIN base_animales ba ON ba.animal = v.codigo
    WHERE v.nombre_propietario ILIKE (SELECT propietario_like FROM params)
      AND v.fecha_insensibilizacion >= (SELECT fecha_plan FROM params)::timestamp
      AND v.fecha_insensibilizacion <  (SELECT fecha_plan + INTERVAL '1 day' FROM params)
)
SELECT codigo, nombre_propietario, nombre_parte, destino, nombre_cava
FROM pbi
ORDER BY codigo, nombre_parte;
