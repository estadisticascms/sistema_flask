-- Facturas anuladas a mano en /anular_facturas (ya estaban cargadas y luego se anularon en el sistema).
-- Al anular: se borra la fila de ventas_historicas, las líneas de ventas_por_producto quedan con Estatus 'Anulado'
-- y se borran sus líneas de ventas_armados. Lo que había se guarda en `respaldo` (JSON) para poder restaurarla.
-- Mientras no esté restaurada, las cargas (ventas, ventas por producto, armados) no la vuelven a meter.
CREATE TABLE IF NOT EXISTS facturas_anuladas (
  id              INT AUTO_INCREMENT PRIMARY KEY,
  factura         VARCHAR(15)    NOT NULL,
  bodega          VARCHAR(100)   NOT NULL,
  fecha           DATE           NOT NULL,
  identificador   VARCHAR(200)   NULL,              -- de ventas_historicas (factura_tipo_bodega_fecha)
  tipo_venta      VARCHAR(10)    NULL,
  cliente         VARCHAR(150)   NULL,
  vendedor        VARCHAR(100)   NULL,
  total           DECIMAL(14,2)  NULL,
  motivo          VARCHAR(255)   NOT NULL,
  respaldo        LONGTEXT       NOT NULL,          -- {"ventas_historicas": [...], "ventas_por_producto": [...], "ventas_armados": [...]}
  anulado_por     VARCHAR(50)    NULL,
  anulado         TIMESTAMP      NOT NULL DEFAULT CURRENT_TIMESTAMP,
  restaurado_por  VARCHAR(50)    NULL,
  restaurado      DATETIME       NULL,
  KEY idx_factura (factura, fecha),
  KEY idx_identificador (identificador)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
