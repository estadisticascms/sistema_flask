-- Detalle del reporte "Factura de Armados" (armado de mangueras), cargado en /cargar_armados.
-- Una fila por línea de factura. El costo y la utilidad salen de ventas_historicas (factura Mixta, misma fecha).
CREATE TABLE IF NOT EXISTS ventas_armados (
  id          INT AUTO_INCREMENT PRIMARY KEY,
  nfactura    VARCHAR(15)    NOT NULL,
  fecha       DATE           NOT NULL,
  tipo_venta  VARCHAR(10)    NOT NULL,          -- CONTADO / CREDITO
  cliente     VARCHAR(150)   NULL,
  facturador  VARCHAR(100)   NULL,              -- columna "Facturó"
  vendedor    VARCHAR(100)   NULL,
  producto    VARCHAR(255)   NOT NULL,
  cantidad    DECIMAL(12,2)  NOT NULL DEFAULT 0,
  precio      DECIMAL(14,2)  NOT NULL DEFAULT 0,
  subtotal    DECIMAL(14,2)  NOT NULL DEFAULT 0,
  ganancia    DECIMAL(14,2)  NOT NULL DEFAULT 0,
  mano_obra   DECIMAL(14,2)  NOT NULL DEFAULT 0, -- columna "MO" (total de la línea)
  otros       DECIMAL(14,2)  NOT NULL DEFAULT 0, -- columna "Otros": gasto de soldadura
  estatus     VARCHAR(20)    NULL,
  cargado     TIMESTAMP      NOT NULL DEFAULT CURRENT_TIMESTAMP,
  KEY idx_fecha (fecha),
  KEY idx_factura (nfactura, fecha)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
