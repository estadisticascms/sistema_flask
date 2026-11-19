-- Índices de las tablas de ventas (creados el 19/11/2026). No cambian los resultados de los reportes, solo la velocidad.
-- Las cargas hacen DELETE + INSERT (nunca DROP TABLE), así que los índices se conservan.
-- Fecha y Codigo de ventas_por_producto son TEXT: el índice usa los primeros caracteres (10 de la fecha, 20 del código).

ALTER TABLE ventas_por_producto
  ADD INDEX idx_vpp_fecha (Fecha(10)),
  ADD INDEX idx_vpp_codigo_fecha (Codigo(20), Fecha(10)),
  ADD INDEX idx_vpp_factura (`No. Factura`);

ALTER TABLE ventas_historicas
  ADD INDEX idx_vh_fecha (Fecha),
  ADD INDEX idx_vh_factura (factura);

-- Medido antes → después (mejor de 3):
--   ventas por producto de un mes ......... 0.59 s → 0.02 s
--   historial de un producto .............. 0.59 s → 0.00 s
--   una factura (validación) .............. 1.10 s → 0.00 s
--   ventas históricas de un mes ........... 0.16 s → 0.01 s
--   una factura en históricas ............. 0.22 s → 0.00 s
--   períodos de 6 a 24 meses: igual (~1 s), con o sin índice
-- Para quitarlos: ALTER TABLE ventas_por_producto DROP INDEX idx_vpp_fecha, DROP INDEX idx_vpp_codigo_fecha, DROP INDEX idx_vpp_factura;
--                 ALTER TABLE ventas_historicas DROP INDEX idx_vh_fecha, DROP INDEX idx_vh_factura;
