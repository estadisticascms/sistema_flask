-- Bitácora de accesos: quién entra al sistema, qué módulo usa y cuándo (router/bitacora.py).
-- evento: login, fallido, salida, vista (abrió una pantalla), accion (guardó / cargó), descarga, denegado.
CREATE TABLE IF NOT EXISTS bitacora_accesos (
  id          BIGINT AUTO_INCREMENT PRIMARY KEY,
  fecha_hora  DATETIME      NOT NULL,
  usuario     VARCHAR(100)  NOT NULL,
  evento      VARCHAR(15)   NOT NULL,
  modulo      VARCHAR(150)  NULL,            -- nombre de la pantalla según el menú
  area        VARCHAR(60)   NULL,            -- sección del menú (Ventas, Compras, ...)
  ruta        VARCHAR(255)  NULL,
  metodo      VARCHAR(8)    NULL,
  ip          VARCHAR(45)   NULL,
  navegador   VARCHAR(255)  NULL,
  KEY idx_fecha (fecha_hora),
  KEY idx_usuario (usuario, fecha_hora)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
