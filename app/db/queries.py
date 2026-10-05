"""
Queries SQL del flujo de Protocolos de Calidad.

Tablas:
  dbo.FCRMVH   cabecera de movimiento (remitos)
  dbo.STRMVI   ítems / detalle de stock
  dbo.STMPDH   maestro de productos
  dbo.VTMCLH   maestro de clientes

La query trae el DETALLE COMPLETO de items con CODFOR=RX0018 para los
remitos cuya cabecera está en el rango de fechas. NO filtra por
STMPDH_TIPPRO en SQL para que la tabla 1 (resumen) pueda sumar todo el
M2 del remito; la tabla 2 (detalle) se filtra por TIPPRO=SAFED en pandas.

Los campos `Protocolo` y `Protocolo Serie LF` quedan vacíos hasta
confirmar de dónde salen.
"""
from __future__ import annotations

CODFOR_FILTER = "RX0018"

SELECT_DETALLE_POR_RANGO = """
SELECT
    F.FCRMVH_NROCTA                                              AS [Código de Cliente],
    F.FCRMVH_CODFOR                                              AS [CodFor],
    F.FCRMVH_NROFOR                                              AS [NroFor],
    (RTRIM(F.FCRMVH_CODFOR) + '-' + RTRIM(CAST(F.FCRMVH_NROFOR AS VARCHAR(50)))) AS [#Comprobante],
    F.FCRMVH_FCHMOV                                              AS [Fecha],
    F.FCRMVH_TEXTOS                                              AS [Observaciones],
    F.USR_FCRMVH_NROOC                                           AS [Número OC],
    V.VTMCLH_NOMBRE                                              AS [Razón Social],
    I.STRMVI_ARTCOD                                              AS [Producto],
    P.STMPDH_DESCRP                                              AS [_Descrp],
    P.USR_STMPDH_ADHESI                                          AS [_Adhesi],
    P.USR_STMPDH_PROLIN                                          AS [_Prolin],
    P.STMPDH_TIPPRO                                              AS [TipoProducto],
    I.STRMVI_NOTROS                                              AS [Ancho],
    I.STRMVI_NATRIB                                              AS [Largo],
    I.STRMVI_STOCKS                                              AS [#m2],
    I.STRMVI_NSERIE                                              AS [Serie]
FROM dbo.FCRMVH AS F
INNER JOIN dbo.STRMVI AS I
    ON I.STRMVI_CODFOR = F.FCRMVH_CODFOR
   AND I.STRMVI_NROFOR = F.FCRMVH_NROFOR
LEFT  JOIN dbo.STMPDH AS P
    ON P.STMPDH_ARTCOD = I.STRMVI_ARTCOD
LEFT  JOIN dbo.VTMCLH AS V
    ON V.VTMCLH_NROCTA = F.FCRMVH_NROCTA
WHERE
    F.FCRMVH_CODFOR = ?                                          -- RX0018
AND I.STRMVI_CODFOR = ?                                          -- RX0018
AND F.FCRMVH_FCHMOV >= ?
AND F.FCRMVH_FCHMOV <  DATEADD(DAY, 1, ?)
"""

FILTER_BY_CLIENTE = " AND F.FCRMVH_NROCTA = ? "
FILTER_BY_NROFOR = " AND F.FCRMVH_NROFOR = ? "

ORDER_BY = " ORDER BY F.FCRMVH_NROFOR, I.STRMVI_ARTCOD "


def build_detalle_query(
    by_cliente: bool = False,
    by_nrofor: bool = False,
) -> str:
    sql = SELECT_DETALLE_POR_RANGO
    if by_cliente:
        sql += FILTER_BY_CLIENTE
    if by_nrofor:
        sql += FILTER_BY_NROFOR
    sql += ORDER_BY
    return sql


# ---------------------------------------------------------------------------
# Trazabilidad de papel (DB FSBI)
# ---------------------------------------------------------------------------
TRAZABILIDAD_FORMULARIO_CODES = ("IR", "IRP", "LF")


def build_trazabilidad_query(n_series: int) -> str:
    """Query con IN (?, ?, ...) sobre dbo.TrazabilidadPapel.

    OJO: la columna real en la BD se llama `FormularioCodgo` (typo en la BD,
    falta la 'i'). Lo aliasamos a `FormularioCodigo` para usar el nombre
    correcto en el resto del código.
    Usa TRIM y LEFT(7) defensivos por si los valores tienen padding.
    """
    if n_series <= 0:
        raise ValueError("n_series debe ser > 0")
    placeholders = ",".join(["?"] * n_series)
    codes = ",".join(["?"] * len(TRAZABILIDAD_FORMULARIO_CODES))
    return (
        "SELECT "
        "  LEFT(LTRIM(RTRIM(Serie)), 7) AS Serie, "
        "  LTRIM(RTRIM(FormularioCodgo)) AS FormularioCodigo, "
        "  LTRIM(RTRIM(CAST(FormularioNumero AS VARCHAR(50)))) AS FormularioNumero, "
        "  LTRIM(RTRIM(Producto)) AS Producto "
        "FROM dbo.TrazabilidadPapel "
        f"WHERE LTRIM(RTRIM(FormularioCodgo)) IN ({codes}) "
        f"AND LEFT(LTRIM(RTRIM(Serie)), 7) IN ({placeholders})"
    )


# Lookup por NF directamente (usado por el MailBot Fedrigoni).
# Trae todos los matches que tengan FormularioNumero == NF y código en (LF, IRP).
# El servicio aplicará la prioridad LF > IRP en pandas.
SELECT_TRAZABILIDAD_BY_NF = """
SELECT
    LTRIM(RTRIM(FormularioCodgo)) AS FormularioCodigo,
    LTRIM(RTRIM(CAST(FormularioNumero AS VARCHAR(50)))) AS FormularioNumero,
    LTRIM(RTRIM(Producto)) AS Producto
FROM dbo.TrazabilidadPapel
WHERE LTRIM(RTRIM(FormularioCodgo)) IN ('LF', 'IRP')
  AND LTRIM(RTRIM(CAST(FormularioNumero AS VARCHAR(50)))) = ?
"""


# ---------------------------------------------------------------------------
# Búsqueda de ARTCOD por producto (MailBot Fedrigoni - Formato 2)
# ---------------------------------------------------------------------------
# El PDF F2 trae el "Producto" con formato variable:
#   - "THERMAL TOP BPA FREE FSC / P7 / YG55"  (con `/`)
#   - "TINTORETTO GESSO H+O ULTRA WS FSC SH9020 WG74"  (sin `/`)
# La BD (STMPDH SAFED) almacena DESCRP, ADHESI y PROLIN en columnas separadas.
# El match es fuzzy: las palabras del CONCAT(DESCRP, ADHESI, PROLIN) deben
# aparecer en el PDF en el mismo orden (subsecuencia ordenada), permitiendo
# palabras extras en el PDF como "FSC", "BPA FREE", etc.
#
# Esta query trae TODOS los registros SAFED (orden de miles, no más); el
# matcheo fuzzy se hace en Python.
SELECT_ALL_SAFED_PRODUCTOS = """
SELECT
    LTRIM(RTRIM(STMPDH_ARTCOD))                            AS ARTCOD,
    LTRIM(RTRIM(STMPDH_DESCRP))                            AS DESCRP,
    ISNULL(LTRIM(RTRIM(USR_STMPDH_ADHESI)), '')            AS ADHESI,
    ISNULL(LTRIM(RTRIM(USR_STMPDH_PROLIN)), '')            AS PROLIN
FROM dbo.STMPDH
WHERE LTRIM(RTRIM(STMPDH_TIPPRO)) = 'SAFED'
"""
