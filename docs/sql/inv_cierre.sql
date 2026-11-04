-- Cierre de mes de inventario (/cierre_inventario, se carga en /cargar_cierre_inventario).
-- Se sube el "Reporte de Análisis de Inventario" del sistema cortado al último día del mes.

-- Un registro por cierre (fecha de corte = último día del mes)
CREATE TABLE IF NOT EXISTS inv_cierre (
  fecha_corte  DATE          NOT NULL PRIMARY KEY,
  tc           DECIMAL(10,4) NOT NULL DEFAULT 36.7000,   -- tipo de cambio para el costo en U$ (el reporte usa 36.70)
  origen       VARCHAR(20)   NOT NULL,                   -- 'sistema' (archivo del sistema) / 'historico' (hoja del Excel anterior)
  archivo      VARCHAR(255)  NULL,
  cargado_por  VARCHAR(50)   NULL,
  cargado      TIMESTAMP     NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- Saldo por bodega de cada cierre (lo que muestra el reporte)
CREATE TABLE IF NOT EXISTS inv_cierre_bodega (
  fecha_corte  DATE          NOT NULL,
  bodega       VARCHAR(100)  NOT NULL,                   -- nombre tal como viene en el archivo
  unidades     DECIMAL(18,4) NOT NULL DEFAULT 0,         -- suma de Cantidad Final
  costo_mn     DECIMAL(18,4) NOT NULL DEFAULT 0,         -- suma de Monto Final (C$)
  productos    INT           NULL,                       -- filas producto/bodega con existencia o saldo (solo cargas del sistema)
  PRIMARY KEY (fecha_corte, bodega),
  CONSTRAINT fk_inv_cierre_bodega FOREIGN KEY (fecha_corte) REFERENCES inv_cierre (fecha_corte) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- Detalle por producto y bodega del cierre (solo filas con cantidad o monto distinto de cero)
CREATE TABLE IF NOT EXISTS inv_cierre_detalle (
  fecha_corte  DATE          NOT NULL,
  bodega       VARCHAR(100)  NOT NULL,
  id_producto  INT           NULL,
  codigo       VARCHAR(50)   NOT NULL,
  producto     VARCHAR(255)  NULL,
  numero_parte VARCHAR(100)  NULL,
  marca        VARCHAR(150)  NULL,
  uom          VARCHAR(20)   NULL,
  unidades     DECIMAL(18,4) NOT NULL DEFAULT 0,
  costo_mn     DECIMAL(18,4) NOT NULL DEFAULT 0,
  KEY idx_corte_bodega (fecha_corte, bodega),
  KEY idx_corte_codigo (fecha_corte, codigo),
  CONSTRAINT fk_inv_cierre_detalle FOREIGN KEY (fecha_corte) REFERENCES inv_cierre (fecha_corte) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- Sección y orden de cada bodega en el reporte. clave = nombre en mayúsculas sin espacios ni signos,
-- así "BODEGA #5 MONTOYA" y "BODEGA #5 - MONTOYA" son la misma bodega.
CREATE TABLE IF NOT EXISTS inv_bodega_grupo (
  clave   VARCHAR(100) NOT NULL PRIMARY KEY,
  bodega  VARCHAR(100) NOT NULL,                         -- nombre que se muestra
  grupo   VARCHAR(20)  NOT NULL,                         -- BODEGAS / REVISAR / LIQUIDAR
  orden   INT          NOT NULL DEFAULT 999
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
