-- Facturas anuladas según el reporte «FACTURAS ANULADAS» del sistema (se sube en /cargar_anuladas_sistema).
-- Es solo un registro para medir cuántas facturas anula cada vendedor y por qué motivo; no toca las tablas de ventas
-- (para sacar una factura de los reportes de ventas está «Anular facturas», tabla facturas_anuladas).
-- vendedor = vendedor asignado; si viene en blanco, el facturador. El motivo se clasifica a partir de `razon` al hacer el reporte.
-- Una factura se identifica por número + fecha + total (los números se repiten entre bodegas); si el reporte la trae otra vez,
-- se actualiza con lo último que diga el sistema.
CREATE TABLE IF NOT EXISTS facturas_anuladas_sistema (
  id                 INT AUTO_INCREMENT PRIMARY KEY,
  factura            VARCHAR(30)    NOT NULL,
  fecha              DATE           NOT NULL,
  total              DECIMAL(14,2)  NOT NULL,
  cliente            VARCHAR(150)   NULL,
  facturador         VARCHAR(100)   NULL,
  vendedor_asignado  VARCHAR(100)   NULL,
  vendedor           VARCHAR(100)   NOT NULL,
  tipo               VARCHAR(10)    NULL,              -- CREDITO / CONTADO
  descripcion        VARCHAR(255)   NULL,
  razon              VARCHAR(500)   NULL,              -- tal como viene, sin el prefijo «FACTURA ANULADA:»
  anulado_por        VARCHAR(50)    NULL,
  fecha_anulado      DATETIME       NULL,
  archivo            VARCHAR(255)   NULL,
  cargado_por        VARCHAR(50)    NULL,
  cargado            TIMESTAMP      NOT NULL DEFAULT CURRENT_TIMESTAMP,
  actualizado        DATETIME       NULL,
  UNIQUE KEY uk_factura (factura, fecha, total),
  KEY idx_fecha (fecha),
  KEY idx_vendedor (vendedor, fecha)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
