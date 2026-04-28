-- ============================================================================
-- Version LISTA PARA PGADMIN (sin placeholders). Ejecutar TODO el bloque
-- (WITH ... ORDER BY).
--
-- Literales usados:
--   fecha_plan       = 2026-04-20
--   propietario_like = %TECNOLOG%AGROPECUARIAS%
--   use_pbi02        = FALSE
--
-- Para probar otro cliente o fecha, solo edita los literales en el CTE params.
-- ============================================================================
WITH params AS MATERIALIZED (
    SELECT
        ('2026-04-20')::date AS fecha_plan,
        ('%TECNOLOG%AGROPECUARIAS%')::text AS propietario_like,
        -- Si USE_PBI02=0 en el .env, el CTE prefijo_real_pbi02 sale vacío y
        -- se usa el prefijo calculado por DOW (LxM/MxM/...) como única fuente.
        -- vw_pbi02 es pesadísima y el JOIN con ella puede costar 10+ minutos.
        (FALSE)::boolean AS use_pbi02
),

base_animales AS MATERIALIZED (
    SELECT DISTINCT
        pfp.id_plan_faena,
        pfp.id_producto AS animal
    FROM trazabilidad_proceso.plan_faena pf
    JOIN trazabilidad_proceso.plan_faena_producto pfp
      ON pfp.id_plan_faena = pf.id
    JOIN params p
      ON pf.fecha_plan = p.fecha_plan
),

-- Solo las columnas necesarias. Filtro principal: JOIN con base_animales
-- (lista corta de animales del día del plan, viene de una tabla indexada).
--
-- Nota: evitamos filtrar por v.fecha_insensibilizacion::date = fecha_plan
-- porque el cast a ::date INUTILIZA cualquier índice sobre fecha_insensibilizacion
-- (obliga a evaluar la función por cada fila de la vista). En su lugar usamos
-- un rango de timestamp [fecha_plan, fecha_plan+1) solo como hint adicional:
-- el optimizador puede usarlo si hay índice BTREE sobre fecha_insensibilizacion.
pbi_filtrado AS MATERIALIZED (
    SELECT
        v.codigo,
        v.nombre_parte,
        v.nombre_propietario,
        v.destino,
        v.nombre_cava,
        v.identificacion,
        v.emergencia,
        v.decomiso,
        v.fecha_insensibilizacion
    FROM trazabilidad_proceso.vw_pbi01 v
    JOIN base_animales ba
      ON ba.animal = v.codigo
    JOIN params p
      ON v.fecha_insensibilizacion >= p.fecha_plan::timestamp
     AND v.fecha_insensibilizacion < (p.fecha_plan + INTERVAL '1 day')
    WHERE v.nombre_propietario ILIKE (SELECT propietario_like FROM params)
),

-- Códigos únicos de animales que pasan el filtro (ID del animal y su prefijo
-- "<a>-<b>" para buscar en vw_pbi02).
codigos_filtrados AS MATERIALIZED (
    SELECT DISTINCT
        codigo AS animal,
        codigo || '-' AS codigo_prefix
    FROM pbi_filtrado
),

animales AS MATERIALIZED (
    SELECT DISTINCT
        ba.id_plan_faena,
        ba.animal
    FROM base_animales ba
    JOIN codigos_filtrados cf
      ON cf.animal = ba.animal
),

base AS MATERIALIZED (
    -- Un animal puede quedar asociado a más de un plan_faena en el mismo día.
    -- Para el reporte debe salir una sola fila por animal.
    SELECT DISTINCT ON (pr.id)
        pr.id AS animal,
        pr.sexo,
        TRIM(BOTH ' - ' FROM CONCAT_WS(' - ', NULLIF(TRIM(pr.especie), ''), NULLIF(TRIM(pr.raza), ''))) AS especie_raza,
        pr.peso_animal_pie,
        pr.fecha_insensibilizacion::date AS fecha_sacrificio,
        pr.fecha_insensibilizacion::time AS hora_sacrificio,
        -- Prefijo GLOBAL por animal (uno solo), basado en la fecha de insensibilización del animal
        CASE EXTRACT(DOW FROM pr.fecha_insensibilizacion::date)
            WHEN 1 THEN 'LxM'
            WHEN 2 THEN 'MxM'
            WHEN 3 THEN 'MxJ'
            WHEN 4 THEN 'JxV'
            WHEN 5 THEN 'VxS'
            WHEN 6 THEN 'SxD'
            WHEN 0 THEN 'DxL'
        END AS prefijo_global,
        pr.peso_media_canal_1,
        pr.peso_media_canal_2,
        CASE
            WHEN pr.peso_animal_pie > 0
            THEN ROUND(
                (COALESCE(pr.peso_media_canal_1, 0) + COALESCE(pr.peso_media_canal_2, 0)) * 100.0 / pr.peso_animal_pie,
                2
            )
        END AS rendimiento,
        pf.fecha_plan,
        pf.id AS id_plan_faena
    FROM animales a
    JOIN trazabilidad_proceso.plan_faena pf
      ON pf.id = a.id_plan_faena
    JOIN trazabilidad_proceso.producto pr
      ON pr.id = a.animal
    ORDER BY pr.id, pf.fecha_plan DESC, pf.id DESC
),

partes_base AS MATERIALIZED (
    SELECT
        pp.id AS id_parte,
        pp.id_producto AS animal,
        pp.id_producto AS id_producto_pp,
        tpp.nombre AS nombre_parte,
        pp.peso_despacho::numeric AS peso_despacho
    FROM trazabilidad_proceso.parte_producto pp
    JOIN trazabilidad_proceso.tipo_parte_producto tpp
      ON tpp.id = pp.id_tipo_parte_producto
    JOIN animales a
      ON a.animal = pp.id_producto
    WHERE tpp.nombre IN (
        'Visceras Blancas',
        'Visceras Rojas',
        'Cabeza',
        'Patas y Manos',
        'Patas y Manos Bovino',
        'Piel'
    )
),

-- IMPORTANTE (rendimiento): antes este CTE era:
--     SELECT DISTINCT ON (ppp.id_parte_producto) ... FROM proceso_parte_producto_peso ppp
--     JOIN partes_base pb ON pb.id_parte = ppp.id_parte_producto
--     ORDER BY ppp.id_parte_producto, ppp.fecha_peso DESC, ppp.hora_peso DESC
-- El plan ordenaba la tabla ENTERA proceso_parte_producto_peso (~6.8M filas)
-- antes del DISTINCT ON. EXPLAIN: cost=47M, 230M rows estimadas. La consulta
-- duraba >6 min en el replica y la cancelaban con "conflict with recovery".
--
-- Clave: proceso_parte_producto_peso tiene un indice COMPUESTO
-- (id_producto, id_parte_producto) -- se llama
-- trazabilidad_proceso_pppp_parte_producto_proceso_salida. Si el LATERAL
-- filtra SOLO por id_parte_producto, PostgreSQL no puede usar ese indice
-- (id_producto es la primera columna) y cae a Seq Scan + Sort por cada fila.
-- Con LATERAL que filtra por AMBAS columnas, usa Index Scan directo.
peso_ult AS MATERIALIZED (
    SELECT
        pb.id_parte,
        pu.peso AS peso_ult
    FROM partes_base pb
    CROSS JOIN LATERAL (
        SELECT ppp.peso
        FROM trazabilidad_proceso.proceso_parte_producto_peso ppp
        WHERE ppp.id_producto       = pb.id_producto_pp
          AND ppp.id_parte_producto = pb.id_parte
        ORDER BY ppp.fecha_peso DESC NULLS LAST,
                 ppp.hora_peso DESC NULLS LAST
        LIMIT 1
    ) pu
),

partes AS MATERIALIZED (
    SELECT
        pb.animal,
        MAX(CASE WHEN pb.nombre_parte = 'Visceras Blancas'
                 THEN COALESCE(pu.peso_ult, pb.peso_despacho) END) AS peso_kg_3,
        MAX(CASE WHEN pb.nombre_parte = 'Visceras Rojas'
                 THEN COALESCE(pu.peso_ult, pb.peso_despacho) END) AS peso_kg_5,
        MAX(CASE WHEN pb.nombre_parte = 'Cabeza'
                 THEN COALESCE(pu.peso_ult, pb.peso_despacho) END) AS peso_kg_7,
        MAX(CASE WHEN pb.nombre_parte IN ('Patas y Manos', 'Patas y Manos Bovino')
                 THEN COALESCE(pu.peso_ult, pb.peso_despacho) END) AS peso_kg_9,
        MAX(CASE WHEN pb.nombre_parte = 'Piel'
                 THEN COALESCE(pu.peso_ult, pb.peso_despacho) END) AS peso_kg_12
    FROM partes_base pb
    LEFT JOIN peso_ult pu
      ON pu.id_parte = pb.id_parte
    GROUP BY pb.animal
),

-- ============================================================================
-- DESTINOS por parte (nueva logica simple y universal):
--   destino_texto = sucursal.nombre + (si observaciones) ' (' || obs || ')'
--
-- La tabla organizaciones.sucursal ya contiene pre-creadas las sucursales con
-- prefijo del dia (ej. "02085 /LxM/", "E115 /MxM/", "P19 /JxV/"), ademas de
-- las cavas sin prefijo ("CAVA T.A") y compradores externos ("VICTOR HUGO Y
-- CIA"). No necesitamos calcular el prefijo del dia ni aplicar reglas de CAVA.
-- ============================================================================
partes_destino AS MATERIALIZED (
    SELECT
        pp.id_producto               AS animal,
        pp.id_tipo_parte_producto    AS id_tipo,
        s.nombre                     AS sucursal_nombre,
        s.nombre
            || CASE
                   WHEN NULLIF(BTRIM(COALESCE(pp.observaciones, '')), '') IS NOT NULL
                   THEN ' (' || BTRIM(pp.observaciones) || ')'
                   ELSE ''
               END AS destino_texto
    FROM trazabilidad_proceso.parte_producto pp
    JOIN animales a
      ON a.animal = pp.id_producto
    LEFT JOIN trazabilidad_proceso.parte_producto_empresa ppe
      ON ppe.id_parte_producto = pp.id
     AND ppe.id_producto       = pp.id_producto
    LEFT JOIN trazabilidad_proceso.parte_producto_empresa_local ppel
      ON ppel.id_parte_producto_empresa = ppe.id
    LEFT JOIN organizaciones.sucursal s
      ON s.id = ppel.id_local
    WHERE pp.id_tipo_parte_producto IN (4, 5, 10, 11, 12, 13, 14)
),

-- Pivot: una fila por animal con las 7 columnas de destino.
destinos AS (
    SELECT
        pd.animal,
        MAX(CASE WHEN pd.id_tipo = 4  THEN pd.destino_texto END) AS destino_mc1,
        MAX(CASE WHEN pd.id_tipo = 5  THEN pd.destino_texto END) AS destino_mc2,
        MAX(CASE WHEN pd.id_tipo = 13 THEN pd.destino_texto END) AS destino_viscera_blanca,
        MAX(CASE WHEN pd.id_tipo = 14 THEN pd.destino_texto END) AS destino_viscera_roja,
        MAX(CASE WHEN pd.id_tipo = 10 THEN pd.destino_texto END) AS destino_cabeza,
        MAX(CASE WHEN pd.id_tipo = 11 THEN pd.destino_texto END) AS destino_patas,
        MAX(CASE WHEN pd.id_tipo = 12 THEN pd.destino_texto END) AS destino_piel
    FROM partes_destino pd
    GROUP BY pd.animal
),

-- Decomiso textual real por animal desde SAI:
--   sai.inspeccion_decomiso -> sai.decomiso -> nombre real (si existe)
--   fallback: tipo_parte_producto.nombre
decomisos_sai AS MATERIALIZED (
    SELECT
        i.id_producto AS animal,
        STRING_AGG(
            DISTINCT COALESCE(NULLIF(BTRIM(d.observacion), ''), tpp.nombre),
            ', '
            ORDER BY COALESCE(NULLIF(BTRIM(d.observacion), ''), tpp.nombre)
        ) AS decomiso_texto
    FROM sai.inspeccion i
    JOIN sai.inspeccion_decomiso idc
      ON idc.id_inspeccion = i.id
    JOIN sai.decomiso d
      ON d.id = idc.id_decomiso
    JOIN trazabilidad_proceso.tipo_parte_producto tpp
      ON tpp.id = d.id_tipo_parte_producto
    JOIN animales a
      ON a.animal = i.id_producto
    GROUP BY i.id_producto
),

-- Emergencia y propietario desde vw_pbi01 (con el filtro pbi_filtrado ya recortado).
-- Decomiso prioriza SAI textual; si no existe, cae al campo legado de vw_pbi01.
info_animal AS MATERIALIZED (
    SELECT
        v.codigo AS animal,
        MAX(v.nombre_propietario) AS propietario,
        BOOL_OR(
            CASE
                WHEN v.emergencia IS NULL THEN FALSE
                WHEN v.emergencia::text ILIKE 't%' THEN TRUE
                WHEN v.emergencia::text = '1'      THEN TRUE
                WHEN v.emergencia::text ILIKE 's%' THEN TRUE
                ELSE FALSE
            END
        ) AS emergencia,
        COALESCE(
            ds.decomiso_texto,
            NULLIF(BTRIM(MAX(v.decomiso)::text), '')
        ) AS decomiso
    FROM pbi_filtrado v
    LEFT JOIN decomisos_sai ds
      ON ds.animal = v.codigo
    GROUP BY v.codigo, ds.decomiso_texto
)

SELECT
    b.animal AS "Animal",
    b.sexo AS "Sexo",
    b.especie_raza AS "Especie - Raza",
    b.peso_animal_pie AS "Peso Animal en Pie (Kg)",
    b.fecha_sacrificio AS "Fecha Sacrificio",
    b.hora_sacrificio AS "Hora Sacrificio",

    b.peso_media_canal_1 AS "Peso (Kg.)",
    COALESCE(d.destino_mc1, '') AS "Destino",

    b.peso_media_canal_2 AS "Peso (Kg.)2",
    COALESCE(d.destino_mc2, '') AS "Destino2",

    p.peso_kg_3 AS "Peso (Kg.)3",
    COALESCE(d.destino_viscera_blanca, '') AS "Destino4",

    p.peso_kg_5 AS "Peso (Kg.)5",
    COALESCE(d.destino_viscera_roja, '') AS "Destino6",

    p.peso_kg_7 AS "Peso (Kg.)7",
    COALESCE(d.destino_cabeza, '') AS "Destino8",

    p.peso_kg_9 AS "Peso (Kg.)9",
    COALESCE(d.destino_patas, '') AS "Destino10",

    p.peso_kg_12 AS "Peso (Kg.)12",
    COALESCE(d.destino_piel, '') AS "Destino11",

    CASE WHEN COALESCE(ia.emergencia, FALSE) THEN 'SI' ELSE 'NO' END AS "Sacrificio de Emergencia",
    ia.decomiso AS "Decomiso",
    b.rendimiento AS "Rendimiento (%)",
    ia.propietario AS "Propietario",
    b.fecha_plan AS "Fecha plan",
    b.id_plan_faena
FROM base b
LEFT JOIN partes p
  ON p.animal = b.animal
LEFT JOIN destinos d
  ON d.animal = b.animal
LEFT JOIN info_animal ia
  ON ia.animal = b.animal
ORDER BY b.id_plan_faena, b.animal;
