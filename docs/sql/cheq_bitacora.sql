-- Bitácora de cheques (router/cheques.py): quién creó, editó, imprimió o anuló cada cheque, desde qué equipo e IP y qué cambió.
-- cambios: JSON [{"campo": "Monto", "antes": "C$ 1,000.00", "despues": "C$ 1,500.00"}, ...]
CREATE TABLE IF NOT EXISTS cheq_bitacora (
  id           BIGINT AUTO_INCREMENT PRIMARY KEY,
  cheque_id    INT           NOT NULL,
  numero       INT           NULL,             -- número del cheque en ese momento
  fecha_hora   DATETIME      NOT NULL,
  usuario      VARCHAR(50)   NOT NULL,
  accion       VARCHAR(15)   NOT NULL,         -- Creado, Editado, Impreso, Anulado
  estado_antes VARCHAR(10)   NULL,
  motivo       VARCHAR(255)  NULL,             -- obligatorio al editar un cheque ya impreso
  cambios      TEXT          NULL,
  ip           VARCHAR(45)   NULL,
  equipo       VARCHAR(100)  NULL,             -- nombre que el usuario le dio a su equipo (se recuerda en el navegador)
  equipo_red   VARCHAR(255)  NULL,             -- nombre del equipo según la red (DNS inverso de la IP)
  navegador    VARCHAR(255)  NULL,
  KEY idx_cheque (cheque_id, fecha_hora),
  KEY idx_fecha (fecha_hora)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
