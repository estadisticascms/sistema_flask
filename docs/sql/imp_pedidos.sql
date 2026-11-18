-- Pedidos de mercadería importada (/pedidos_importacion): encabezado, líneas y stock de seguridad por producto.
-- Las fechas son opcionales: se llenan a medida que avanza el pedido. Los tiempos (días) se calculan en router/pedidos_importacion.py.
-- Catálogo, existencias, ventas y cantidad mínima (MOQ) salen de las tablas que ya tiene el sistema
-- (catalogo_productos, existencias_bodega, ventas_por_producto, cantidad_minima_compra).

CREATE TABLE IF NOT EXISTS imp_pedidos (
  id                  INT AUTO_INCREMENT PRIMARY KEY,
  orden_compra        VARCHAR(40)   NOT NULL,
  proveedor           VARCHAR(150)  NULL,
  pais                VARCHAR(60)   NULL,
  cotizacion          VARCHAR(60)   NULL,
  factura             VARCHAR(60)   NULL,
  fecha_factura       DATE          NULL,
  descripcion         VARCHAR(500)  NULL,
  transporte          VARCHAR(20)   NULL,                    -- Marítimo / Aéreo / Terrestre
  estado              VARCHAR(20)   NOT NULL DEFAULT 'Cotizando',
  f_solicitud_cot     DATE          NULL,                    -- Solicitud cotización
  f_cot_recibida      DATE          NULL,                    -- Cotización recibida
  f_pedido            DATE          NULL,
  f_embarque          DATE          NULL,
  f_llegada_estimada  DATE          NULL,
  f_llegada_real      DATE          NULL,
  f_liberacion        DATE          NULL,                    -- Liberación aduana
  f_recepcion         DATE          NULL,                    -- Recepción bodega
  observaciones       TEXT          NULL,
  creado_por          VARCHAR(50)   NULL,
  creado              TIMESTAMP     NOT NULL DEFAULT CURRENT_TIMESTAMP,
  actualizado_por     VARCHAR(50)   NULL,
  actualizado         TIMESTAMP     NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  UNIQUE KEY uq_orden_compra (orden_compra),
  KEY idx_proveedor (proveedor),
  KEY idx_estado (estado)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS imp_pedido_detalle (
  id                  INT AUTO_INCREMENT PRIMARY KEY,
  pedido_id           INT           NOT NULL,
  codigo              VARCHAR(20)   NULL,                    -- `Codigo Producto` del catálogo
  descripcion         VARCHAR(300)  NULL,
  cantidad_pedida     DECIMAL(12,2) NOT NULL DEFAULT 0,
  cantidad_recibida   DECIMAL(12,2) NULL,
  costo_unitario      DECIMAL(14,4) NULL,
  KEY idx_pedido (pedido_id),
  KEY idx_codigo (codigo),
  CONSTRAINT fk_imp_detalle_pedido FOREIGN KEY (pedido_id) REFERENCES imp_pedidos (id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- Stock de seguridad en días por producto (si no tiene, se usa el valor general de la pantalla del sugerido)
CREATE TABLE IF NOT EXISTS imp_stock_seguridad (
  codigo              VARCHAR(20)   NOT NULL PRIMARY KEY,
  dias                INT           NOT NULL,
  actualizado_por     VARCHAR(50)   NULL,
  actualizado         TIMESTAMP     NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
