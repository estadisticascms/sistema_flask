-- =====================================================================================================================
-- Módulo de seguimiento de entregas — migración inicial (parte 1)
-- Especificación: docs/entregas/ESPECIFICACION_MODULO_ENTREGAS.md, sección 5.
--
-- NO APLICADA. Revisar antes de ejecutar. Solo crea tablas nuevas con prefijo ent_; no modifica ninguna tabla existente.
--
-- Diferencias con la sección 5 de la especificación:
--   1. No se crea ent_usuarios ni PIN. Las personas son las de la tabla `usuarios` existente: el administrador crea el
--      usuario y su contraseña como hasta ahora. Los roles del módulo van en ent_usuario_rol, sin tocar `usuarios`; un
--      usuario puede tener varios roles (una fila por rol). Todas las columnas que en la especificación apuntaban a
--      ent_usuarios(id) apuntan a usuarios(id).
--   2. ent_envio_eventos lleva la columna `campo` (estado, existencia, traslado, preparacion, factura) para distinguir qué
--      cambió: la especificación pide guardar ahí también los cambios de cada línea, y valores como 'PENDIENTE' existen
--      en varios campos.
--   3. ent_envios.pedido_activo (columna generada) con índice único: impide en la base subir dos veces el mismo pedido
--      mientras el anterior no esté ANULADO (la especificación lo pedía validar en código; se valida en los dos lados).
--   4. Llaves foráneas e índices que faltaban: ent_envio_lineas.bodega_prepara_id, existencia_por, actualizado_por;
--      ent_envio_eventos.linea_id; índices para las colas por bodega y por estado.
--      ent_envio_lineas.subtotal: la sección 4.6 lo extrae por línea pero la tabla no lo tenía; sirve para guardar lo
--      leído tal cual y para valorizar la venta perdida.
--   5. En ent_configuracion solo se insertan los valores decididos (hora_corte, minutos_urgencia). Los marcados
--      [PENDIENTE] no se insertan: los carga el admin cuando la empresa los dé.
--
--   6. Sin ent_bodegas.bodega_traslado_id: la bodega destino del traslado no es fija, depende del tipo de producto; la escoge el
--      coordinador en cada línea al aplicar el traslado. Por eso ent_envio_lineas.bodega_prepara_id admite NULL (destino por
--      definir) hasta ese momento. (Cambio del 24/11/2026; para la base ya creada ver el bloque «Cambios» al final.)
--
-- Orden de creación: catálogos → rol de usuario → rutas → envíos → líneas, facturas y eventos.
-- Fechas y horas: DATETIME en hora de Nicaragua (UTC-6). El servidor actual está en esa zona (time_zone = SYSTEM).
-- Para deshacer (en este orden), ver el bloque comentado al final.
-- =====================================================================================================================

-- ---------------------------------------------------------------------------------------------------------------------
-- Catálogos
-- ---------------------------------------------------------------------------------------------------------------------
CREATE TABLE ent_bodegas (
  id                 INT AUTO_INCREMENT PRIMARY KEY,
  nombre_en_orden    VARCHAR(60) NOT NULL,                -- texto exacto del PDF: 'BODEGA #1', 'BODEGA #5 MONTOYA'
  nombre             VARCHAR(60) NOT NULL,
  requiere_traslado  TINYINT(1)  NOT NULL DEFAULT 0,      -- lo define paquetería; el destino lo escoge el coordinador por línea
  activo             TINYINT(1)  NOT NULL DEFAULT 1,
  UNIQUE KEY uq_ent_bodegas_nombre_en_orden (nombre_en_orden)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE ent_transportistas (
  id                    INT AUTO_INCREMENT PRIMARY KEY,
  nombre                VARCHAR(100) NOT NULL,
  tipo                  ENUM('empresa','bus','retira_cliente','reparto_propio') NOT NULL,
  direccion             VARCHAR(255),
  mapa_url              VARCHAR(255),
  telefono              VARCHAR(30),
  minutos_traslado      INT,
  hora_limite_recepcion TIME,
  activo                TINYINT(1) NOT NULL DEFAULT 1
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE ent_transportista_horarios (
  id               INT AUTO_INCREMENT PRIMARY KEY,
  transportista_id INT NOT NULL,
  destino          VARCHAR(100) NOT NULL,
  hora_salida      TIME NOT NULL,
  dias             VARCHAR(30),                           -- 'L,M,X,J,V,S'
  CONSTRAINT fk_ent_horarios_transportista FOREIGN KEY (transportista_id) REFERENCES ent_transportistas (id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE ent_configuracion (
  clave VARCHAR(50) PRIMARY KEY,
  valor VARCHAR(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

INSERT INTO ent_configuracion (clave, valor) VALUES
  ('hora_corte', '15:00'),
  ('minutos_urgencia', '60');
-- [PENDIENTE] minutos_preparacion, minutos_traslado_bodega, minutos_facturacion, minutos_empaque, margen_amarillo:
-- estimados que debe dar la empresa; no se insertan aquí.

-- ---------------------------------------------------------------------------------------------------------------------
-- Roles de cada usuario en el módulo (reemplaza a ent_usuarios; la persona es la de `usuarios`).
-- Un usuario puede tener varios roles: una fila por rol (p. ej. paqueteria + mensajero).
-- ---------------------------------------------------------------------------------------------------------------------
CREATE TABLE ent_usuario_rol (
  usuario_id INT NOT NULL,                               -- usuarios.id
  rol        ENUM('vendedor','bodega','coordinador_bodega','paqueteria','mensajero','admin') NOT NULL,
  bodega_id  INT NULL,                                   -- solo en la fila del rol bodega (bodeguero): la bodega que ve y cambia
  activo     TINYINT(1) NOT NULL DEFAULT 1,
  PRIMARY KEY (usuario_id, rol),
  KEY idx_ent_usuario_rol_rol (rol, activo),
  KEY idx_ent_usuario_rol_bodega (bodega_id),
  CONSTRAINT fk_ent_usuario_rol_usuario FOREIGN KEY (usuario_id) REFERENCES usuarios (id),
  CONSTRAINT fk_ent_usuario_rol_bodega  FOREIGN KEY (bodega_id)  REFERENCES ent_bodegas (id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

-- ---------------------------------------------------------------------------------------------------------------------
-- Rutas del mensajero
-- ---------------------------------------------------------------------------------------------------------------------
CREATE TABLE ent_rutas (
  id           INT AUTO_INCREMENT PRIMARY KEY,
  mensajero_id INT NOT NULL,                              -- usuarios.id
  fecha        DATE NOT NULL,
  estado       ENUM('ARMADA','EN_CURSO','CERRADA') NOT NULL DEFAULT 'ARMADA',
  salida       DATETIME,
  regreso      DATETIME,
  KEY idx_ent_rutas_mensajero_fecha (mensajero_id, fecha),
  CONSTRAINT fk_ent_rutas_mensajero FOREIGN KEY (mensajero_id) REFERENCES usuarios (id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE ent_ruta_paradas (
  id               INT AUTO_INCREMENT PRIMARY KEY,
  ruta_id          INT NOT NULL,
  transportista_id INT NOT NULL,
  orden            INT NOT NULL,
  guia             VARCHAR(50),
  monto_recibo     DECIMAL(10,2),
  foto_entrega_url VARCHAR(255),
  foto_recibo_url  VARCHAR(255),
  entregado_en     DATETIME,
  KEY idx_ent_paradas_ruta (ruta_id, orden),
  CONSTRAINT fk_ent_paradas_ruta          FOREIGN KEY (ruta_id)          REFERENCES ent_rutas (id),
  CONSTRAINT fk_ent_paradas_transportista FOREIGN KEY (transportista_id) REFERENCES ent_transportistas (id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

-- ---------------------------------------------------------------------------------------------------------------------
-- Envíos (una orden de venta/pedido)
-- ---------------------------------------------------------------------------------------------------------------------
CREATE TABLE ent_envios (
  id                  INT AUTO_INCREMENT PRIMARY KEY,
  pedido_num          VARCHAR(20)  NOT NULL,              -- de la orden
  cliente             VARCHAR(150) NOT NULL,
  terminos            VARCHAR(30)  NOT NULL,              -- 'CONTADO', crédito…
  fecha_pedido        DATE,
  ingresado_por_orden VARCHAR(60),                        -- campo "Ingresado Por" del PDF
  subtotal_orden      DECIMAL(12,2),
  pdf_url             VARCHAR(255),
  vendedor_id         INT NOT NULL,                       -- usuarios.id
  transportista_id    INT NOT NULL,
  destino             VARCHAR(100),
  hora_requerida      DATETIME NOT NULL,
  prioridad           ENUM('normal','urgente') NOT NULL DEFAULT 'normal',
  riesgo              ENUM('verde','amarillo','rojo'),
  riesgo_confirmado   TINYINT(1) NOT NULL DEFAULT 0,
  forma_pago          ENUM('transferencia','contra_entrega') NULL,   -- [PENDIENTE] solo contado; opcional hasta confirmar
  flete_paga          ENUM('cliente','empresa') NOT NULL DEFAULT 'cliente',
  bultos              INT,
  peso_kg             DECIMAL(8,2),
  parada_id           INT NULL,
  flete_monto         DECIMAL(10,2),
  estado              ENUM('ORDEN_ENVIADA','EN_REVISION','POR_FACTURAR','FACTURADA','LISTA_PARA_RETIRO',
                           'RETIRADA','EMPACADA','EN_RUTA','ENTREGADA','ANULADA') NOT NULL DEFAULT 'ORDEN_ENVIADA',
  motivo_anulacion    VARCHAR(255),
  notas               TEXT,
  creado_en           DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  actualizado_en      DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  -- Igual a pedido_num mientras el envío no esté anulado; NULL si está ANULADA. El índice único permite varios NULL,
  -- así un pedido se puede volver a subir solo si el anterior se anuló.
  pedido_activo       VARCHAR(20) GENERATED ALWAYS AS (IF(estado = 'ANULADA', NULL, pedido_num)) STORED,
  UNIQUE KEY uq_ent_envios_pedido_activo (pedido_activo),
  KEY idx_ent_envios_pedido (pedido_num),
  KEY idx_ent_envios_estado (estado, hora_requerida),
  KEY idx_ent_envios_vendedor (vendedor_id, estado),
  CONSTRAINT fk_ent_envios_vendedor      FOREIGN KEY (vendedor_id)      REFERENCES usuarios (id),
  CONSTRAINT fk_ent_envios_transportista FOREIGN KEY (transportista_id) REFERENCES ent_transportistas (id),
  CONSTRAINT fk_ent_envios_parada        FOREIGN KEY (parada_id)        REFERENCES ent_ruta_paradas (id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE ent_envio_lineas (
  id                   INT AUTO_INCREMENT PRIMARY KEY,
  envio_id             INT NOT NULL,
  linea                INT NOT NULL,
  bodega_id            INT NOT NULL,
  codigo               VARCHAR(50) NOT NULL,
  producto             VARCHAR(150),
  medida               VARCHAR(60),
  marca                VARCHAR(40),
  cantidad             DECIMAL(10,2) NOT NULL,
  subtotal             DECIMAL(12,2),                     -- de la orden; sirve para el control de cuadre y la venta perdida
  cant_disponible      DECIMAL(10,2),                     -- al marcar FALTANTE
  existencia           ENUM('PENDIENTE','CONFIRMADA','FALTANTE','QUITADA') NOT NULL DEFAULT 'PENDIENTE',
  existencia_por       INT,                               -- usuarios.id de quien confirmó o marcó faltante
  existencia_en        DATETIME,
  decision             ENUM('quitar','cambiar_bodega','esperar') NULL,
  traslado_estado      ENUM('NO_REQUIERE','PENDIENTE','APLICADO','EN_CAMINO','RECIBIDO') NOT NULL DEFAULT 'NO_REQUIERE',
  traslado_num         VARCHAR(30),
  traslado_aplicado_en DATETIME,
  traslado_recibido_en DATETIME,
  preparacion          ENUM('PENDIENTE','EN_PREPARACION','LISTA','RETIRADA') NOT NULL DEFAULT 'PENDIENTE',
  responsable          ENUM('bodeguero','coordinador') NOT NULL,   -- de la existencia; según requiere_traslado al crear la línea
  bodega_prepara_id    INT NULL,                          -- bodega_id; con traslado, la destino que escoge el coordinador (NULL hasta entonces)
  actualizado_por      INT,                               -- usuarios.id
  actualizado_en       DATETIME,
  UNIQUE KEY uq_ent_lineas_envio_linea (envio_id, linea),
  KEY idx_ent_lineas_existencia (bodega_id, existencia),
  KEY idx_ent_lineas_preparacion (bodega_prepara_id, preparacion),
  KEY idx_ent_lineas_traslado (traslado_estado),
  CONSTRAINT fk_ent_lineas_envio          FOREIGN KEY (envio_id)          REFERENCES ent_envios (id),
  CONSTRAINT fk_ent_lineas_bodega         FOREIGN KEY (bodega_id)         REFERENCES ent_bodegas (id),
  CONSTRAINT fk_ent_lineas_bodega_prepara FOREIGN KEY (bodega_prepara_id) REFERENCES ent_bodegas (id),
  CONSTRAINT fk_ent_lineas_existencia_por FOREIGN KEY (existencia_por)    REFERENCES usuarios (id),
  CONSTRAINT fk_ent_lineas_actualizado    FOREIGN KEY (actualizado_por)   REFERENCES usuarios (id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE ent_envio_facturas (                         -- una factura por bodega de facturación
  id             INT AUTO_INCREMENT PRIMARY KEY,
  envio_id       INT NOT NULL,
  bodega_id      INT NOT NULL,
  factura_num    VARCHAR(30) NOT NULL,
  registrada_por INT NOT NULL,                            -- usuarios.id
  registrada_en  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY uq_ent_facturas_envio_bodega (envio_id, bodega_id),
  KEY idx_ent_facturas_numero (factura_num),              -- conciliación contra la facturación del día
  CONSTRAINT fk_ent_facturas_envio   FOREIGN KEY (envio_id)       REFERENCES ent_envios (id),
  CONSTRAINT fk_ent_facturas_bodega  FOREIGN KEY (bodega_id)      REFERENCES ent_bodegas (id),
  CONSTRAINT fk_ent_facturas_usuario FOREIGN KEY (registrada_por) REFERENCES usuarios (id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE ent_envio_eventos (                          -- historial: cada cambio de estado del envío o de una línea
  id         INT AUTO_INCREMENT PRIMARY KEY,
  envio_id   INT NOT NULL,
  linea_id   INT NULL,                                    -- NULL: cambio del envío
  usuario_id INT NOT NULL,                                -- usuarios.id
  campo      ENUM('estado','existencia','traslado','preparacion','factura') NOT NULL DEFAULT 'estado',
  de_estado  VARCHAR(20),
  a_estado   VARCHAR(20) NOT NULL,
  comentario VARCHAR(255),
  creado_en  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  KEY idx_ent_eventos_envio (envio_id, creado_en),
  KEY idx_ent_eventos_linea (linea_id),
  CONSTRAINT fk_ent_eventos_envio   FOREIGN KEY (envio_id)   REFERENCES ent_envios (id),
  CONSTRAINT fk_ent_eventos_linea   FOREIGN KEY (linea_id)   REFERENCES ent_envio_lineas (id),
  CONSTRAINT fk_ent_eventos_usuario FOREIGN KEY (usuario_id) REFERENCES usuarios (id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

-- ---------------------------------------------------------------------------------------------------------------------
-- Para deshacer (solo si todavía no hay datos que conservar), en este orden:
-- DROP TABLE ent_envio_eventos, ent_envio_facturas, ent_envio_lineas, ent_envios, ent_ruta_paradas, ent_rutas,
--            ent_usuario_rol, ent_configuracion, ent_transportista_horarios, ent_transportistas, ent_bodegas;
-- ---------------------------------------------------------------------------------------------------------------------

-- =====================================================================================================================
-- Cambios sobre la base ya creada (las tablas de arriba ya los incluyen para una instalación nueva)
-- =====================================================================================================================
-- 24/11/2026 · El destino del traslado lo escoge el coordinador por línea (depende del tipo de producto):
-- ALTER TABLE ent_bodegas DROP FOREIGN KEY fk_ent_bodegas_traslado, DROP COLUMN bodega_traslado_id;
-- ALTER TABLE ent_envio_lineas MODIFY bodega_prepara_id INT NULL;

-- 24/11/2026 · Avisos dentro del módulo (recordatorios de paquetería y avisos automáticos: orden nueva, faltante, listo para
-- facturar, cambio de bodega). Un aviso por destinatario; se marca leído cuando la persona lo abre.
CREATE TABLE ent_avisos (
  id         INT AUTO_INCREMENT PRIMARY KEY,
  usuario_id INT NOT NULL,                               -- destinatario (usuarios.id)
  envio_id   INT NULL,
  linea_id   INT NULL,
  tipo       ENUM('orden_nueva','faltante','listo_facturar','cambio_bodega','recordatorio','anulada','bodega_desconocida','facturada')
             NOT NULL,
  mensaje    VARCHAR(255) NOT NULL,
  creado_por INT NULL,                                   -- usuarios.id; NULL = aviso automático del sistema
  creado_en  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  leido_en   DATETIME NULL,
  KEY idx_ent_avisos_usuario (usuario_id, leido_en, creado_en),
  KEY idx_ent_avisos_envio (envio_id),
  CONSTRAINT fk_ent_avisos_usuario FOREIGN KEY (usuario_id) REFERENCES usuarios (id),
  CONSTRAINT fk_ent_avisos_envio   FOREIGN KEY (envio_id)   REFERENCES ent_envios (id),
  CONSTRAINT fk_ent_avisos_linea   FOREIGN KEY (linea_id)   REFERENCES ent_envio_lineas (id),
  CONSTRAINT fk_ent_avisos_creador FOREIGN KEY (creado_por) REFERENCES usuarios (id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;
-- Para deshacer: DROP TABLE ent_avisos;

-- 24/11/2026 · Tipos de aviso controlados por el código (router/entregas_avisos.py) en lugar de una lista fija:
ALTER TABLE ent_avisos MODIFY tipo VARCHAR(30) NOT NULL;

-- 24/11/2026 · Bodegas de vendedor: la mercadería ya se entregó al cliente; de sus líneas solo se envía la factura con el resto del
-- paquete (no se confirman, preparan ni retiran, pero su factura se pide y sale en la etiqueta).
ALTER TABLE ent_bodegas ADD COLUMN solo_factura TINYINT(1) NOT NULL DEFAULT 0 AFTER requiere_traslado;
