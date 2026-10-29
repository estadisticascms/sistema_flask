"""Estructura única del menú: la usan la página de inicio, la barra superior (layout.html) y la pantalla de Accesos.
Está organizada por ÁREA de trabajo y, dentro de cada área, por TIPO (reportes, control, herramientas).
Para agregar una pantalla nueva, agréguela en el grupo que corresponda y aparecerá en los tres lugares.

Cada opción: (título, url, icono, descripción, palabras de búsqueda opcionales).
Las secciones con "acceso_rapido" repiten pantallas de otras áreas (atajos para Gerencia); no definen permisos propios.
El menú de cada usuario solo muestra las pantallas que tiene habilitadas en Accesos."""
from flask import current_app, g, session

from utils.db import get_connection

# color: "naranja" (#FE6412) o "verde" (#027266), colores corporativos; "gris" para administración
SECCIONES = [
    {
        "id": "gerencia", "titulo": "Gerencia", "menu": "Gerencia", "icono": "bi-speedometer", "color": "verde", "acceso_rapido": True,
        "descripcion": "Lo que gerencia general revisa: ventas del mes, recuperación, cartera, compras e inventario lento.",
        "grupos": [
            {"titulo": "Tableros", "items": [
                ("Resumen de ventas", "/dashboard", "bi-speedometer2", "Mes en curso: ventas, recuperación, cartera e ingresos", "gerencia indicadores kpi mes"),
                ("Resumen general", "/resumen", "bi-collection", "Ventas y recuperación", "gerencia"),
                ("Ventas por departamento (mapa)", "/ventas_departamento", "bi-map", "Mapa, segmento, familia y subfamilia", "gerencia"),
            ]},
            {"titulo": "Riesgo y oportunidades", "items": [
                ("Análisis de cartera", "/analisis_cartera", "bi-wallet2", "Crédito vs. recuperación, límites y recomendación", "gerencia mora vencido"),
                ("Seguimiento de ingresos", "/seguimiento_ingresos", "bi-box-arrow-in-down", "Lo que ingresó: cuánto se vendió y a quién", "gerencia compras"),
                ("Poco movimiento y marketing", "/productos_poco_movimiento", "bi-megaphone", "Inventario lento, precios y estrategia", "gerencia"),
            ]},
        ],
    },
    {
        "id": "ventas", "titulo": "Ventas", "menu": "Ventas", "icono": "bi-graph-up-arrow", "color": "verde",
        "descripcion": "Reportes de venta, desempeño de vendedores y control de precios.",
        "grupos": [
            {"titulo": "Reportes de ventas", "items": [
                ("Resumen de ventas", "/dashboard", "bi-speedometer2", "Mes en curso: ventas, recuperación, cartera e ingresos", "dashboard tablero"),
                ("Ventas por día", "/reporte_ventas_dia", "bi-calendar3", "Evolución diaria", "diario"),
                ("Ventas mensuales", "/reporte_ventas_mes", "bi-calendar-month", "Por tipo de vendedor", "mes"),
                ("Armado de mangueras", "/reporte_armados", "bi-tools", "Mangueras por día, venta, margen, mano de obra y soldadura", "armados mixtas bienes taller hidraulica"),
                ("Ventas por bodega", "/ventas_bodega", "bi-shop", "Unidades vendidas por bodega", "sucursal"),
            ]},
            {"titulo": "Por zona y departamento", "items": [
                ("Ventas por departamento (mapa)", "/ventas_departamento", "bi-map", "Mapa, segmento, familia y subfamilia", "zona geografia"),
                ("Departamentos por mes", "/reporteZona", "bi-geo-alt-fill", "Comparativo mensual, clientes, productos y vendedor", "zona"),
                ("Ventas por zona mensual", "/ventas_zonas_mes", "bi-diagram-3-fill", "Evolución mensual por zona", "departamento"),
            ]},
            {"titulo": "Vendedores", "items": [
                ("Totales por vendedor", "/rep_vend_totales", "bi-trophy-fill", "Ranking para compartir con vendedores", "ranking comision"),
                ("Crédito y contado por vendedor", "/rep_vend_netocreco", "bi-file-earmark-bar-graph", "Venta neta crédito / contado", "credito contado"),
                ("Ventas por vendedor (detalle)", "/reporte_ventas", "bi-bar-chart-fill", "Detalle por vendedor y mes", ""),
                ("Top clientes por vendedor", "/top_clientes_vendedor", "bi-person-lines-fill", "Mejores clientes del período", "ranking"),
                ("Top clientes por marca", "/top_clientes_marca", "bi-person-badge", "Compras por cliente, mes y vendedor", "ranking"),
                ("Listado de existencias para clientes", "/listado_existencias", "bi-card-checklist", "Por marca, segmento y familia, con aplicación; Excel y PDF", "inventario stock disponible catalogo lista precios cliente"),
            ]},
            {"titulo": "Control", "items": [
                ("Revisión de precios para comisionistas", "/revision_precios_venta", "bi-tags", "¿Se respetó el precio del cliente? Para pagar su comisión", "comision comisionista precio venta descuento factura recibo auditoria"),
                ("Ventas perdidas", "/ventas_perdidas", "bi-graph-down-arrow", "Registrar y ver oportunidades no concretadas", "faltante"),
            ]},
        ],
    },
    {
        "id": "compras", "titulo": "Compras e inventario", "menu": "Compras", "icono": "bi-box-seam", "color": "naranja",
        "descripcion": "Qué pedir, revisar lo que ingresa y controlar el inventario y los precios.",
        "grupos": [
            {"titulo": "Pedidos e ingresos", "items": [
                ("Sugerido de compras", "/sugerido_compras", "bi-cart-plus-fill", "Pedido por marca o proveedor, editable y en Excel", "pedido fob proveedor"),
                ("Próximos a agotarse", "/productos_agotarse", "bi-hourglass-split", "Cobertura de existencias", "agotado reorden"),
                ("Revisión de factura de ingreso", "/revision_ingreso", "bi-clipboard-check", "Códigos, SAC, duplicados, nuevos y FOB", "arancel importacion proveedor"),
                ("Seguimiento de ingresos", "/seguimiento_ingresos", "bi-box-arrow-in-down", "Lo que ingresó: cuánto se vendió y a quién", "orden de compra"),
            ]},
            {"titulo": "Inventario y precios", "items": [
                ("Productos sin movimiento", "/productos_sin_movimiento", "bi-moon-stars", "Con existencia y sin ventas en meses", "lento obsoleto"),
                ("Revisión de precios", "/revision_precios", "bi-percent", "Margen bajo y sin movimiento con margen alto", "margen costo"),
                ("Top marcas", "/top_marcas", "bi-award-fill", "Marcas más vendidas", "ranking"),
                ("Clasificación y fichas", "/clase_prod", "bi-tags-fill", "Segmento, familia, equivalencias y ficha incompleta", "catalogo"),
            ]},
            {"titulo": "Marketing", "items": [
                ("Poco movimiento y marketing", "/productos_poco_movimiento", "bi-megaphone", "Existencias lentas, precios y estrategia", "promocion descuento liquidacion"),
                ("Seguimiento de campañas", "/campanas_mkt", "bi-graph-up-arrow", "Ventas de los productos en campaña", "promocion"),
            ]},
        ],
    },
    {
        "id": "cobranza", "titulo": "Clientes y cobranza", "menu": "Cobranza", "icono": "bi-people-fill", "color": "verde",
        "descripcion": "Cartera de crédito, recuperación y datos de clientes y zonas.",
        "grupos": [
            {"titulo": "Cartera y recuperación", "items": [
                ("Análisis de cartera", "/analisis_cartera", "bi-wallet2", "Crédito vs. recuperación, límites y recomendación", "mora vencido limite credito"),
                ("Recibos de caja por día", "/reporte_recibos", "bi-list-check", "Recuperación diaria", "cobro pagos"),
                ("Resumen general", "/resumen", "bi-collection", "Ventas y recuperación", ""),
            ]},
            {"titulo": "Clientes y zonas", "items": [
                ("Clientes", "/clientes", "bi-people", "Datos y zona de cada cliente", ""),
                ("Zonas", "/zonas", "bi-pin-map-fill", "Listado de zonas", ""),
                ("Crear zona", "/newzonas", "bi-plus-circle-fill", "Registrar una zona nueva", ""),
            ]},
        ],
    },
    {
        "id": "finanzas", "titulo": "Finanzas", "menu": "Finanzas", "icono": "bi-bank", "color": "verde",
        "descripcion": "Emisión e impresión de cheques con su partida contable.",
        "grupos": [
            {"titulo": "Cheques", "items": [
                ("Cheques", "/cheques", "bi-journal-check", "Buscar, imprimir y anular cheques", "pago banco"),
                ("Nuevo cheque", "/cheques/nuevo", "bi-plus-circle", "Emitir un cheque con su partida", "pago"),
            ]},
            {"titulo": "Configuración", "items": [
                ("Bancos y consecutivos", "/cheques/bancos", "bi-bank", "Cuentas, números de cheque y ajuste de impresión", ""),
                ("Cuentas contables", "/cheques/cuentas", "bi-diagram-3", "Catálogo para la partida del cheque", "contabilidad"),
            ]},
        ],
    },
    {
        "id": "cargas", "titulo": "Carga diaria de datos", "menu": "Cargas", "icono": "bi-cloud-arrow-up-fill", "color": "naranja",
        "descripcion": "Suba los reportes del sistema, de preferencia en este orden.", "pasos": True,
        "grupos": [
            {"titulo": "Cada día", "items": [
                ("Ventas del día", "/cargar_ventas", "bi-cash-coin", "Facturas de crédito y contado", "subir excel"),
                ("Ventas por producto", "/cargar_ventas_producto", "bi-cart-check-fill", "Detalle de productos vendidos", "subir excel"),
                ("Transacciones de inventario", "/cargar_inventario", "bi-arrow-left-right", "Entradas, salidas y ajustes", "subir excel"),
                ("Recibos de caja", "/cargar_recibos", "bi-receipt-cutoff", "Recuperación del día", "subir excel"),
                ("Recuperación (comisiones)", "/cargar_recuperacion", "bi-cash-stack", "Pagos de crédito por vendedor", "subir excel"),
                ("Armado de mangueras", "/cargar_armados", "bi-tools", "Reporte Factura de Armados", "subir excel mangueras"),
                ("Devoluciones", "/devoluciones", "bi-arrow-return-left", "Devoluciones de clientes", "subir excel"),
            ]},
            {"titulo": "Cuando cambian", "items": [
                ("Existencias por bodega", "/cargar_existencias", "bi-boxes", "Reemplaza la existencia actual", "subir excel"),
                ("Catálogo de productos", "/cargar_catalogo", "bi-journal-bookmark-fill", "Existencias, costos y precios", "subir excel"),
                ("Clientes", "/cargar_clientes", "bi-person-vcard", "Actualiza la lista de clientes", "subir excel"),
            ]},
        ],
    },
    {
        "id": "admin", "titulo": "Administración", "menu": "Admin", "icono": "bi-shield-lock-fill", "color": "gris",
        "descripcion": "Usuarios del sistema y permisos por pantalla.",
        "grupos": [
            {"titulo": "Usuarios", "items": [
                ("Usuarios", "/usuarios", "bi-person-gear", "Crear y editar usuarios", "contraseña clave"),
                ("Accesos", "/accesos", "bi-shield-check", "Permisos de cada usuario", "permisos"),
                ("Bitácora de accesos", "/bitacora_accesos", "bi-clock-history", "Quién entró, qué módulo usó, fecha y hora", "auditoria log historial ingresos usuarios"),
            ]},
        ],
    },
]

# Formato plano (título, url, icono, descripción) por sección, para quien recorre todas las opciones
for _s in SECCIONES:
    _s["items"] = [i[:4] for _g in _s["grupos"] for i in _g["items"]]


def vista_por_url():
    """URL del menú -> nombre del permiso (lo deja @requiere_acceso en la función de la ruta)."""
    mapa = current_app.config.get("_MENU_VISTAS")
    if mapa is None:
        mapa = {}
        for regla in current_app.url_map.iter_rules():
            vista = getattr(current_app.view_functions.get(regla.endpoint), "vista_acceso", None)
            if vista:
                mapa.setdefault(regla.rule, vista)
        current_app.config["_MENU_VISTAS"] = mapa
    return mapa


def permisos_del_usuario(usuario):
    if "menu_permisos" not in g:
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT vista FROM accesos_vistas WHERE usuario = %s AND permitido = 1", (usuario,))
            g.menu_permisos = {r[0] for r in cursor.fetchall()}
        finally:
            cursor.close()
            conn.close()
    return g.menu_permisos


def contexto_menu():
    """Context processor: el menú de la persona (solo lo que puede abrir) disponible en todas las plantillas."""
    usuario = session.get("usuario")
    if not usuario:
        return {"menu_secciones": [], "menu_buscar": []}
    try:
        permitidas = permisos_del_usuario(usuario)
        mapa = vista_por_url()
    except Exception:   # sin base de datos: menú completo (cada pantalla igual valida su permiso)
        permitidas, mapa = None, {}
    secciones, buscar, vistos = [], [], set()
    # El buscador recorre primero las áreas propias: cada pantalla aparece con su área real, no con el atajo de Gerencia
    for s in sorted(SECCIONES, key=lambda x: bool(x.get("acceso_rapido"))):
        grupos = []
        for grupo in s["grupos"]:
            items = [i for i in grupo["items"] if permitidas is None or mapa.get(i[1]) is None or mapa[i[1]] in permitidas]
            if items:
                grupos.append({"titulo": grupo["titulo"], "items": [i[:4] for i in items]})
                for titulo, url, icono, desc, *extra in items:
                    if url not in vistos:
                        vistos.add(url)
                        buscar.append({"t": titulo, "u": url, "i": icono, "d": desc, "a": s["menu"], "k": (extra[0] if extra else "")})
        if grupos:
            secciones.append({**s, "grupos": grupos, "items": [i for gr in grupos for i in gr["items"]]})
    orden = {s["id"]: n for n, s in enumerate(SECCIONES)}
    secciones.sort(key=lambda x: orden[x["id"]])   # en el menú se respeta el orden original (Gerencia primero)
    return {"menu_secciones": secciones, "menu_buscar": buscar}
