-- Cantidad mínima que se puede comprar de cada producto y número de parte del proveedor (se sube en /cantidades_minimas).
-- El sugerido de compras la usa para el «Sugerido ajustado»: el sugerido se sube al siguiente múltiplo de la cantidad mínima
-- (mínimo 10: 3 → 10, 13 → 20). El pedido y el total FOB parten del sugerido ajustado.
-- parte_proveedor: código con el que el proveedor conoce el producto (ej. 11AVX260 para nuestro Nº parte 15260); sale en el
-- Excel del sugerido para armar el pedido. Una fila puede tener solo mínimo, solo parte_proveedor o ambos.
CREATE TABLE IF NOT EXISTS cantidad_minima_compra (
  codigo          VARCHAR(20)    NOT NULL PRIMARY KEY,     -- `Codigo Producto` del catálogo
  parte_proveedor VARCHAR(60)    NULL,
  minimo          DECIMAL(12,2)  NULL,                     -- > 0; NULL = sin mínimo
  archivo         VARCHAR(255)   NULL,
  actualizado_por VARCHAR(50)    NULL,
  actualizado     TIMESTAMP      NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX idx_parte_proveedor (parte_proveedor)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- Cambio del 18/11/2026 sobre la tabla ya creada (ejecutado):
-- ALTER TABLE cantidad_minima_compra ADD COLUMN parte_proveedor VARCHAR(60) NULL AFTER codigo,
--   MODIFY minimo DECIMAL(12,2) NULL, ADD INDEX idx_parte_proveedor (parte_proveedor);
