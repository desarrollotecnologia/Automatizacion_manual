-- Ejecuta este archivo directo en PostgreSQL (pgAdmin/psql).
-- Edita aquí la fecha y el propietario antes de correr.
WITH params AS (
    SELECT
        DATE '2026-04-13' AS fecha_plan,
        '%VARGAS NI%O YERSON REYNALDO%'::text AS propietario_like
),

base_animales AS (
    SELECT DISTINCT
        pfp.id_plan_faena,
        pfp.id_producto AS animal
    FROM trazabilidad_proceso.plan_faena pf
    JOIN trazabilidad_proceso.plan_faena_producto pfp
      ON pfp.id_plan_faena = pf.id
    JOIN params p
      ON pf.fecha_plan = p.fecha_plan
),

pbi_filtrado AS (
    SELECT v.*
    FROM trazabilidad_proceso.vw_pbi01 v
    JOIN base_animales ba
      ON ba.animal = v.codigo
    JOIN params p
      ON v.nombre_propietario ILIKE p.propietario_like
),

animales AS (
    SELECT DISTINCT
        ba.id_plan_faena,
        ba.animal
    FROM base_animales ba
    WHERE EXISTS (
        SELECT 1
        FROM pbi_filtrado v
        WHERE v.codigo = ba.animal
    )
),

base AS (
    SELECT DISTINCT ON (a.id_plan_faena, pr.id)
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
    ORDER BY a.id_plan_faena, pr.id, pf.fecha_plan DESC
),

partes_base AS (
    SELECT
        pp.id AS id_parte,
        pp.id_producto AS animal,
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

peso_ult AS (
    SELECT DISTINCT ON (ppp.id_parte_producto)
        ppp.id_parte_producto AS id_parte,
        ppp.peso AS peso_ult
    FROM trazabilidad_proceso.proceso_parte_producto_peso ppp
    JOIN partes_base pb
      ON pb.id_parte = ppp.id_parte_producto
    ORDER BY
        ppp.id_parte_producto,
        ppp.fecha_peso DESC NULLS LAST,
        ppp.hora_peso DESC NULLS LAST
),

partes AS (
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

destino_base AS (
    SELECT DISTINCT ON (v.codigo, v.nombre_parte)
        v.codigo AS animal,
        v.nombre_parte,
        v.nombre_propietario,
        NULLIF(BTRIM(v.destino), '') AS destino_raw,
        NULLIF(BTRIM(v.nombre_cava), '') AS nombre_cava_raw,
        v.identificacion,
        v.emergencia,
        v.decomiso,
        v.fecha_insensibilizacion
    FROM pbi_filtrado v
    ORDER BY
        v.codigo,
        v.nombre_parte,
        v.fecha_insensibilizacion DESC NULLS LAST
),

destino_normalizado AS (
    SELECT
        d.animal,
        d.nombre_parte,
        d.nombre_propietario,
        d.emergencia,
        d.decomiso,
        d.destino_raw,
        -- Igual que consulta_rendimientos.sql (nombre_cava solo acota con MEDIA CANAL; evita solo-CAVA en todo el reporte).
        (
            (
                UPPER(BTRIM(COALESCE(d.destino_raw, ''))) LIKE 'CAVA%'
                OR (
                    NULLIF(BTRIM(COALESCE(d.destino_raw, '')), '') IS NULL
                    AND UPPER(COALESCE(d.nombre_cava_raw, '')) LIKE 'CAVA%'
                    AND UPPER(COALESCE(d.nombre_parte, '')) LIKE 'MEDIA CANAL%'
                )
                OR (
                    UPPER(COALESCE(d.nombre_cava_raw, '')) LIKE 'CAVA%'
                    AND UPPER(COALESCE(d.nombre_parte, '')) LIKE 'MEDIA CANAL%'
                    AND NOT (
                        COALESCE(BTRIM(d.destino_raw), '')
                        ~ '/[[:alnum:]]{1,4}[xX][[:alnum:]]{1,4}/'
                    )
                )
            )
            AND NOT (
                COALESCE(BTRIM(d.destino_raw), '')
                ~ '/[[:alnum:]]{1,4}[xX][[:alnum:]]{1,4}/'
            )
        ) AS es_cava,
        COALESCE(
            CASE
                WHEN BTRIM(COALESCE(d.destino_raw, '')) <> ''
                     AND SPLIT_PART(BTRIM(d.destino_raw), ' ', 1) ~ '^[0-9]{3,8}$'
                THEN LPAD(SPLIT_PART(BTRIM(d.destino_raw), ' ', 1), 5, '0')
            END,
            CASE
                WHEN BTRIM(COALESCE(d.destino_raw, '')) <> ''
                     AND SPLIT_PART(BTRIM(d.destino_raw), ' ', 1) ~ '^[A-Za-z]{1,3}[0-9]{1,4}$'
                THEN UPPER(SPLIT_PART(BTRIM(d.destino_raw), ' ', 1))
            END,
            CASE
                WHEN BTRIM(COALESCE(d.destino_raw, '')) <> ''
                     AND UPPER(SPLIT_PART(BTRIM(d.destino_raw), ' ', 1)) IN ('LP', 'JN', 'PJQ')
                THEN UPPER(SPLIT_PART(BTRIM(d.destino_raw), ' ', 1))
            END,
            CASE
                WHEN BTRIM(COALESCE(d.destino_raw, '')) <> ''
                     AND UPPER(SPLIT_PART(BTRIM(d.destino_raw), ' ', 1)) IN ('DANY', 'CHXS')
                THEN UPPER(SPLIT_PART(BTRIM(d.destino_raw), ' ', 1))
            END,
            LPAD(SUBSTRING(COALESCE(d.identificacion, '') FROM '^[0-9]+-([0-9]{5})-'), 5, '0'),
            LPAD(SUBSTRING(COALESCE(d.identificacion, '') FROM '([0-9]{5})'), 5, '0'),
            d.animal::text
        ) AS codigo_5,
        COALESCE(
            CASE EXTRACT(DOW FROM d.fecha_insensibilizacion::date)
            WHEN 1 THEN 'LxM'
            WHEN 2 THEN 'MxM'
            WHEN 3 THEN 'MxJ'
            WHEN 4 THEN 'JxV'
            WHEN 5 THEN 'VxS'
            WHEN 6 THEN 'SxD'
            WHEN 0 THEN 'DxL'
            END,
            'SIN'
        ) AS prefijo
    FROM destino_base d
),

destinos AS (
    SELECT
        n.animal,
        MAX(n.nombre_propietario) AS propietario,
        MAX(n.emergencia) AS emergencia,
        MAX(n.decomiso) AS decomiso,
        CASE
            WHEN BOOL_OR(n.es_cava) THEN COALESCE(
                MAX(
                    CASE
                        WHEN n.es_cava
                             AND UPPER(BTRIM(COALESCE(n.destino_raw, ''))) LIKE 'CAVA%'
                        THEN NULLIF(BTRIM(n.destino_raw), '')
                    END
                ),
                'CAVA.'
            )
            ELSE MAX(n.codigo_5) || ' /' || COALESCE(MAX(n.prefijo), MAX(b.prefijo_global), 'SIN') || '/'
        END AS destino_global
    FROM destino_normalizado n
    LEFT JOIN base b
      ON b.animal = n.animal
    GROUP BY n.animal
)

SELECT
    b.animal AS "Animal",
    b.sexo AS "Sexo",
    b.especie_raza AS "Especie - Raza",
    b.peso_animal_pie AS "Peso Animal en Pie (Kg)",
    b.fecha_sacrificio AS "Fecha Sacrificio",
    b.hora_sacrificio AS "Hora Sacrificio",

    b.peso_media_canal_1 AS "Peso (Kg.)",
    d.destino_global AS "Destino",

    b.peso_media_canal_2 AS "Peso (Kg.)2",
    d.destino_global AS "Destino2",

    p.peso_kg_3 AS "Peso (Kg.)3",
    d.destino_global AS "Destino4",

    p.peso_kg_5 AS "Peso (Kg.)5",
    d.destino_global AS "Destino6",

    p.peso_kg_7 AS "Peso (Kg.)7",
    d.destino_global AS "Destino8",

    p.peso_kg_9 AS "Peso (Kg.)9",
    d.destino_global AS "Destino10",

    p.peso_kg_12 AS "Peso (Kg.)12",
    d.destino_global AS "Destino11",

    CASE
        WHEN d.emergencia IS NULL THEN 'NO'
        WHEN d.emergencia::text ILIKE 't%' THEN 'SI'
        WHEN d.emergencia::text = '1' THEN 'SI'
        WHEN d.emergencia::text ILIKE 'si%' THEN 'SI'
        WHEN d.emergencia::text ILIKE 's%' THEN 'SI'
        ELSE 'NO'
    END AS "Sacrificio de Emergencia",
    d.decomiso AS "Decomiso",
    b.rendimiento AS "Rendimiento (%)",
    d.propietario AS "Propietario",
    b.fecha_plan AS "Fecha plan",
    b.id_plan_faena
FROM base b
LEFT JOIN partes p
  ON p.animal = b.animal
LEFT JOIN destinos d
  ON d.animal = b.animal
ORDER BY b.id_plan_faena, b.animal;

-- Para guardar en una tabla:
-- 1) Crear tabla con el resultado:
-- CREATE TABLE trazabilidad_proceso.reporte_rendimientos_tmp AS
-- <pega la consulta de arriba>;
--
-- 2) O insertar en una tabla existente:
-- INSERT INTO trazabilidad_proceso.reporte_rendimientos_tmp(<cols...>)
-- <pega la consulta de arriba>;
