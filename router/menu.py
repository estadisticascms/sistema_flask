"""Estructura única del menú: la usan la página de inicio y la barra superior (layout.html).
Para agregar una pantalla nueva, agréguela aquí y aparecerá en ambos lugares."""

# color: "naranja" (#FE6412) o "verde" (#027266), colores corporativos; "gris" para administración
SECCIONES = [
    {
        "id": "cargas", "titulo": "Carga diaria de datos", "menu": "Cargas", "icono": "bi-cloud-arrow-up-fill",
        "color": "naranja", "descripcion": "Suba los reportes del sistema, de preferencia en este orden.",
        "items": [
            ("Ventas del día", "/cargar_ventas", "bi-cash-coin", "Facturas de crédito y contado"),
            ("Ventas por producto", "/cargar_ventas_producto", "bi-cart-check-fill", "Detalle de productos vendidos"),
            ("Transacciones de inventario", "/cargar_inventario", "bi-arrow-left-right", "Entradas, salidas y ajustes"),
            ("Recibos de caja", "/cargar_recibos", "bi-receipt-cutoff", "Recuperación del día"),
            ("Devoluciones", "/devoluciones", "bi-arrow-return-left", "Devoluciones de clientes"),
            ("Existencias por bodega", "/cargar_existencias", "bi-boxes", "Reemplaza la existencia actual"),
            ("Catálogo de productos", "/cargar_catalogo", "bi-journal-bookmark-fill", "Existencias, costos y precios"),
            ("Clientes", "/cargar_clientes", "bi-person-vcard", "Actualiza la lista de clientes"),
        ],
    },
    {
        "id": "ventas", "titulo": "Ventas y vendedores", "menu": "Ventas", "icono": "bi-graph-up-arrow",
        "color": "verde", "descripcion": "Cómo van las ventas de la empresa y de cada vendedor.",
        "items": [
            ("Resumen de ventas", "/dashboard", "bi-speedometer2", "Indicadores del día y del mes"),
            ("Totales por vendedor", "/rep_vend_totales", "bi-trophy-fill", "Ranking para compartir con vendedores"),
            ("Top clientes por vendedor", "/top_clientes_vendedor", "bi-person-lines-fill", "Mejores clientes del período"),
            ("Ventas por vendedor (detalle)", "/reporte_ventas", "bi-bar-chart-fill", "Detalle por vendedor y mes"),
            ("Crédito y contado por vendedor", "/rep_vend_netocreco", "bi-file-earmark-bar-graph", "Venta neta crédito / contado"),
            ("Ventas por día", "/reporte_ventas_dia", "bi-calendar3", "Evolución diaria"),
            ("Ventas mensuales", "/reporte_ventas_mes", "bi-calendar-month", "Por tipo de vendedor"),
            ("Bienes y mixtas", "/reporte_facturas", "bi-receipt", "Facturas por tipo"),
            ("Resumen general", "/resumen", "bi-collection", "Ventas y recuperación"),
            ("Ventas por zona", "/reporteZona", "bi-geo-alt-fill", "Clientes y ventas por zona"),
            ("Ventas por zona mensual", "/ventas_zonas_mes", "bi-diagram-3-fill", "Evolución mensual por zona"),
        ],
    },
    {
        "id": "productos", "titulo": "Productos y compras", "menu": "Productos", "icono": "bi-box-seam",
        "color": "naranja", "descripcion": "Qué comprar, qué se agota y qué se vende.",
        "items": [
            ("Sugerido de compras", "/sugerido_compras", "bi-cart-plus-fill", "Pedido por marca, editable y en Excel"),
            ("Próximos a agotarse", "/productos_agotarse", "bi-hourglass-split", "Cobertura de existencias"),
            ("Productos sin movimiento", "/productos_sin_movimiento", "bi-moon-stars", "Con existencia y sin ventas en meses"),
            ("Revisión de precios", "/revision_precios", "bi-percent", "Margen bajo y sin movimiento con margen alto"),
            ("Top marcas", "/top_marcas", "bi-award-fill", "Marcas más vendidas"),
            ("Ventas por bodega", "/ventas_bodega", "bi-shop", "Unidades vendidas por bodega"),
            ("Clasificación y fichas", "/clase_prod", "bi-tags-fill", "Segmento, familia, equivalencias y ficha incompleta"),
        ],
    },
    {
        "id": "clientes", "titulo": "Clientes y cobranza", "menu": "Clientes", "icono": "bi-people-fill",
        "color": "verde", "descripcion": "Datos de clientes, zonas, recuperación y ventas perdidas.",
        "items": [
            ("Clientes", "/clientes", "bi-people", "Datos y zona de cada cliente"),
            ("Zonas", "/zonas", "bi-pin-map-fill", "Listado de zonas"),
            ("Crear zona", "/newzonas", "bi-plus-circle-fill", "Registrar una zona nueva"),
            ("Recibos de caja por día", "/reporte_recibos", "bi-list-check", "Recuperación diaria"),
            ("Ventas perdidas", "/ventas_perdidas", "bi-graph-down-arrow", "Oportunidades no concretadas"),
        ],
    },
    {
        "id": "admin", "titulo": "Administración", "menu": "Administración", "icono": "bi-shield-lock-fill",
        "color": "gris", "descripcion": "Usuarios del sistema y permisos por pantalla.",
        "items": [
            ("Usuarios", "/usuarios", "bi-person-gear", "Crear y editar usuarios"),
            ("Accesos", "/accesos", "bi-shield-check", "Permisos de cada usuario"),
        ],
    },
]


def contexto_menu():
    """Context processor: deja el menú disponible en todas las plantillas."""
    return {"menu_secciones": SECCIONES}
