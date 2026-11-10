-- Cantidad mínima que se puede comprar de cada producto (se sube en /cantidades_minimas).
-- El sugerido de compras la usa para el «Sugerido ajustado»: el sugerido se sube al siguiente múltiplo de la cantidad mínima
-- (mínimo 10: 3 → 10, 13 → 20). El pedido y el total FOB parten del sugerido ajustado.
CREATE TABLE IF NOT EXISTS cantidad_minima_compra (
  codigo          VARCHAR(20)    NOT NULL PRIMARY KEY,     -- `Codigo Producto` del catálogo
  minimo          DECIMAL(12,2)  NOT NULL,                 -- > 0
  archivo         VARCHAR(255)   NULL,
  actualizado_por VARCHAR(50)    NULL,
  actualizado     TIMESTAMP      NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
