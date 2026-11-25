"""Entregas: avisos dentro del módulo (regla 6.6). Tabla ent_avisos: un aviso por destinatario, leído cuando la persona lo abre.

Avisos automáticos (los crea router/entregas_flujo.py y la subida de la orden):
- orden_nueva: a paquetería y al responsable de cada línea (bodeguero de la bodega o coordinador).
- faltante: al vendedor de la orden y a paquetería.
- listo_facturar: al vendedor cuando la orden queda POR_FACTURAR.
- cambio_bodega: al responsable de la nueva bodega cuando el vendedor cambia una línea faltante de bodega.
- facturada: a paquetería cuando la orden queda FACTURADA.
- lista_retiro: a paquetería cuando la orden queda LISTA_PARA_RETIRO.
- traslado_recibido: al bodeguero de la bodega destino cuando el coordinador marca que llegó la mercadería trasladada.
- anulada: a paquetería y a los responsables con líneas abiertas.
- bodega_desconocida: al administrador de Entregas cuando una orden trae una bodega que no está en el catálogo.
Recordatorios: paquetería los manda desde el tablero de seguimiento al responsable de lo que falta.
"""


def destinatarios(cursor, rol, bodega_id=None):
    """ids de usuarios activos con ese rol de Entregas (y esa bodega, para el bodeguero)."""
    sql, params = "SELECT usuario_id FROM ent_usuario_rol WHERE rol = %s AND activo = 1", [rol]
    if bodega_id is not None:
        sql += " AND bodega_id = %s"
        params.append(bodega_id)
    cursor.execute(sql, params)
    return [fila["usuario_id"] if isinstance(fila, dict) else fila[0] for fila in cursor.fetchall()]


def responsables_linea(cursor, linea):
    """Quién registra la existencia de la línea: coordinadores, o los bodegueros de su bodega."""
    if linea["responsable"] == "coordinador":
        return destinatarios(cursor, "coordinador_bodega")
    return destinatarios(cursor, "bodega", linea["bodega_id"])


def avisar(cursor, usuarios, tipo, mensaje, envio_id=None, linea_id=None, creado_por=None):
    """Crea un aviso por destinatario (sin repetir y sin avisarse a sí mismo). Devuelve cuántos se crearon."""
    unicos = sorted({u for u in usuarios if u and u != creado_por})
    for u in unicos:
        cursor.execute("INSERT INTO ent_avisos (usuario_id, envio_id, linea_id, tipo, mensaje, creado_por) VALUES (%s, %s, %s, %s, %s, %s)",
                       (u, envio_id, linea_id, tipo, mensaje[:255], creado_por))
    return len(unicos)


def avisar_sin_repetir(cursor, usuarios, tipo, mensaje, **kw):
    """Igual que avisar(), pero no repite un aviso idéntico que el destinatario todavía no ha leído."""
    pendientes = []
    for u in set(usuarios):
        cursor.execute("SELECT 1 FROM ent_avisos WHERE usuario_id = %s AND tipo = %s AND mensaje = %s AND leido_en IS NULL LIMIT 1",
                       (u, tipo, mensaje[:255]))
        if not cursor.fetchone():
            pendientes.append(u)
    return avisar(cursor, pendientes, tipo, mensaje, **kw)


def contar_no_leidos(cursor, usuario_id):
    cursor.execute("SELECT COUNT(*) AS n FROM ent_avisos WHERE usuario_id = %s AND leido_en IS NULL", (usuario_id,))
    fila = cursor.fetchone()
    return fila["n"] if isinstance(fila, dict) else fila[0]
