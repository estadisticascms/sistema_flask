# Módulo de seguimiento de entregas — especificación

Documento para construir el módulo dentro del sistema Python existente de COMERSA (Comercial La Esperanza S.A.).
La base de datos es MySQL en la nube. Presentación del proyecto: https://claude.ai/artifact/A9PriGbQTZjp9XhrmHLFJb

## 0. Antes de empezar (para Claude Code)

1. **Revisar la aplicación existente** antes de escribir código: framework (Flask, Django, Streamlit, FastAPI…), cómo se conecta a MySQL, cómo maneja usuarios y sesiones, estructura de carpetas y estilo del código. Construir el módulo con esas mismas piezas.
2. **No modificar tablas existentes.** Crear las tablas nuevas de la sección 5 en un archivo de migración SQL aparte, revisable antes de aplicarlo.
3. Trabajar **por fases** (sección 9). Cada fase debe poder usarse sola.
4. Lo marcado como **[PENDIENTE]** no está decidido: dejarlo configurable o preguntar, nunca inventar la regla.
5. Ejemplo real de orden: `Pedido de Cliente.pdf`. Lector de referencia ya probado: `referencia/leer_orden.py`.

## 1. Contexto: cómo funciona hoy

- El vendedor factura y avisa por un grupo de WhatsApp que hay factura en caja, con los datos de envío.
- Paquetería retira la factura y, según la prioridad, la pasa a bodega.
- Bodega prepara; si no hay existencia, se avisa al vendedor para refacturar.
- Paquetería empaca, etiqueta y despacha por empresa de envíos o bus a los departamentos.
- No hay estado visible del pedido, la aplicación actual trabaja con un día de desfase y los vendedores comparten un solo usuario.

## 2. Cambio de proceso

**El vendedor ya no factura primero.** Envía la **Orden de Venta/Pedido** del sistema de ventas en PDF, que trae cliente, términos y cada producto con su bodega. En cada bodega, **el bodeguero** confirma las existencias y prepara. En las bodegas que requieren traslado, **el coordinador de bodega** consulta las existencias, aplica el traslado y manda a traer la mercadería. Recién entonces **el vendedor modifica la orden y genera las facturas**. Así los faltantes se resuelven antes de facturar y se acaba la refacturación.

- **El sistema de ventas genera una factura por bodega**, así que un envío puede tener varias facturas.
- **Algunas bodegas requieren un traslado** antes de facturar. El vendedor no puede facturar hasta que el traslado esté aplicado.

El sistema de ventas **no da acceso a su base de datos**: la única entrada es el PDF.

## 3. Roles

| Rol | Qué hace | Dispositivo |
|---|---|---|
| `vendedor` | Sube la orden, agrega transporte y hora, decide ante faltantes, modifica la orden y registra las facturas | Computadora o celular |
| `bodega` (bodeguero) | Ve solo las líneas de su bodega. En bodegas **sin** traslado confirma existencias o faltantes. **Siempre registra la preparación**, incluida la de la mercadería trasladada que llega a su bodega | Computadora o celular |
| `coordinador_bodega` | En las bodegas **con** traslado: confirma existencias o faltantes, aplica el traslado, manda a traer la mercadería y marca cuando llega | Computadora o celular |
| `paqueteria` | **Da seguimiento a todas las órdenes** y a la preparación en cada bodega; define qué bodegas aplican traslado; retira producto, empaca, imprime etiqueta, arma rutas | Computadora |
| `mensajero` | Hace la ruta; registra entrega, fotos, recibo y flete | Celular |
| `admin` | Tablero, reportes, catálogos y configuración | Computadora |

- **Cada persona tiene su propio usuario**, no compartido. Un acceso simple (nombre + PIN) es suficiente. Cada cambio guarda quién lo hizo y cuándo.
- **Responsable de la existencia de cada línea:** si su bodega tiene `requiere_traslado = 1`, el coordinador de bodega; si no, el bodeguero asignado a esa bodega. El módulo lo fija al crear la línea (`ent_envio_lineas.responsable`).
- **Responsable de la preparación:** siempre el bodeguero de `bodega_prepara_id`, que es la bodega de la línea o, si hay traslado, la bodega a la que llega (`ent_bodegas.bodega_traslado_id`).
- Cada bodeguero está asignado a una bodega (`ent_usuarios.bodega_id`) y solo ve y cambia las líneas de esa bodega.
- Las pantallas del bodeguero, el coordinador y el mensajero deben funcionar bien en celular.

## 4. Flujo

### 4.1 Estados del envío (`envios.estado`)

```
ORDEN_ENVIADA → EN_REVISION → POR_FACTURAR → FACTURADA → LISTA_PARA_RETIRO → RETIRADA → EMPACADA → EN_RUTA → ENTREGADA
                                                                                  (ANULADA desde cualquier estado)
```

| Estado | Quién lo provoca | Cuándo |
|---|---|---|
| ORDEN_ENVIADA | Vendedor | Confirma la orden leída del PDF y agrega transporte y hora. Se avisa a paquetería y al responsable de cada línea (bodeguero o coordinador) |
| EN_REVISION | Automático | Se registra la primera confirmación de existencia |
| POR_FACTURAR | Automático | Todas las líneas tienen existencia CONFIRMADA o QUITADA, ninguna FALTANTE, y ningún traslado PENDIENTE. Se avisa al vendedor que ya puede facturar |
| FACTURADA | Automático | El vendedor registró la factura de cada bodega de facturación con líneas no quitadas (ver 4.5) |
| LISTA_PARA_RETIRO | Automático | Está FACTURADA, todas las líneas no quitadas están preparadas (LISTA) y la mercadería trasladada ya llegó (RECIBIDO) |
| RETIRADA | Automático | Paquetería marcó retiradas todas las líneas no quitadas |
| EMPACADA | Paquetería | Registra bultos y peso, imprime la etiqueta |
| EN_RUTA | Mensajero | Marca "Salí" en la ruta que contiene el envío |
| ENTREGADA | Mensajero | Registra la entrega en la parada del envío |
| ANULADA | Vendedor o admin | Con motivo obligatorio |

### 4.2 Estados de cada línea

Cada línea lleva **tres estados independientes**, porque los registran personas distintas en momentos distintos:

| Campo | Valores | Quién lo registra |
|---|---|---|
| `existencia` | PENDIENTE → CONFIRMADA, FALTANTE o QUITADA | Responsable de la línea (QUITADA: vendedor) |
| `traslado_estado` | NO_REQUIERE, o PENDIENTE → APLICADO → EN_CAMINO → RECIBIDO | Coordinador de bodega |
| `preparacion` | PENDIENTE → EN_PREPARACION → LISTA → RETIRADA | Bodeguero de la bodega que prepara (RETIRADA: paquetería) |

- **Facturar depende de la existencia y del traslado aplicado**, no de la preparación: el vendedor puede facturar aunque bodega siga preparando o la mercadería trasladada aún venga en camino.
- **Retirar depende de la preparación y de que la mercadería trasladada haya llegado.**
- En FALTANTE el responsable indica la cantidad disponible; si hay parte, se registra cuánta.
- **Paquetería** marca RETIRADA al recoger, por bodega (todas las líneas listas de esa bodega a la vez).

### 4.3 Bodeguero, coordinador y paquetería

**Bodeguero (bodegas sin traslado)**
1. Ve las líneas nuevas de su bodega, con prioridad y riesgo de tiempo.
2. Confirma existencia o marca FALTANTE, por línea o toda su parte de la orden de una vez.
3. Registra la preparación: en preparación y lista.
4. También prepara la mercadería trasladada que llega a su bodega: le aparece en su cola cuando el coordinador marca RECIBIDO, y no puede marcarla en preparación antes.

**Coordinador de bodega (bodegas con traslado)**
1. Ve las líneas de todas las bodegas que requieren traslado.
2. Consulta existencias y marca CONFIRMADA o FALTANTE.
3. **Aplica el traslado** en el sistema de ventas y registra su número → APLICADO (con esto el vendedor ya puede facturar).
4. **Manda a traer la mercadería** → EN_CAMINO, y marca RECIBIDO cuando llega.
5. Al marcar RECIBIDO, la línea pasa a la cola del bodeguero de la bodega que la recibe, que la prepara.

**Paquetería: seguimiento y bodegas con traslado**
1. Tiene un **tablero de seguimiento** con todas las órdenes abiertas: por bodega, qué falta confirmar, preparar o traer, y cuánto tiempo lleva cada paso, con el riesgo de tiempo a la vista. Desde ahí puede recordar al responsable.
2. **Define qué bodegas aplican traslado** (`ent_bodegas.requiere_traslado` y `bodega_traslado_id`). El cambio aplica a las órdenes que se suban después; las líneas ya creadas conservan su responsable.
3. Solo ve el avance; no confirma existencias ni preparación en nombre de otros.

### 4.4 Faltantes

1. El responsable marca la línea FALTANTE y se notifica al vendedor y a paquetería.
2. El envío no puede pasar a POR_FACTURAR mientras haya una línea FALTANTE.
3. El vendedor decide por línea:
   - **Quitar:** la existencia pasa a QUITADA y se registra como venta perdida. El vendedor la quita de la orden en el sistema de ventas al facturar.
   - **Cambiar de bodega:** se elige otra bodega; la línea vuelve a PENDIENTE y toma el responsable de la nueva bodega.
   - **Esperar:** queda FALTANTE hasta que el responsable la marque CONFIRMADA.
4. Si queda solo una parte (por ejemplo 2 de 5), el vendedor decide si va la parte disponible (la cantidad se ajusta y el resto queda como venta perdida) o si espera.

### 4.5 Facturación por el vendedor

1. Cuando el envío pasa a POR_FACTURAR, el vendedor recibe el aviso "listo para facturar".
2. El módulo le muestra **qué ajustar en la orden**: líneas quitadas y cantidades parciales. Así la orden queda igual a lo que de verdad sale.
3. El vendedor modifica la orden y genera las facturas en el sistema de ventas. **Se genera una factura por bodega**.
4. En el módulo registra el número de factura de cada bodega. El módulo pide exactamente una factura por cada bodega de facturación con líneas no quitadas. Cuando están todas, el envío pasa a FACTURADA.
5. **No se puede registrar ninguna factura** mientras el envío no esté POR_FACTURAR, es decir, mientras falte una confirmación de existencia o un traslado por aplicar.

- **Bodega de facturación** de una línea: es su bodega, o la bodega destino del traslado si lo requiere (`ent_bodegas.bodega_traslado_id`). **[PENDIENTE]** Confirmar que después del traslado la factura sale de la bodega destino.
- **[PENDIENTE]** ¿Siempre se factura toda la orden junta, o se puede facturar una bodega mientras otra espera? Por ahora, toda junta.

### 4.6 Lectura del PDF de la orden

- El vendedor sube el PDF de la Orden de Venta/Pedido.
- El módulo extrae: `pedido`, `cliente`, `terminos` (CONTADO o crédito), `fecha`, `ingresado_por`, `subtotal` y, por línea: `linea`, `bodega`, `codigo`, `producto`, `medida`, `marca`, `cantidad`, `subtotal`.
- Referencia: `referencia/leer_orden.py` ubica las columnas por coordenada X con PyMuPDF. Con el pedido 270091 lee las 7 líneas sin errores.
- **Control obligatorio:** la suma de los subtotales de las líneas debe igualar el subtotal impreso, y los números de línea deben ser 1..N. Si no cuadra, el módulo **no acepta** la orden y muestra el error.
- El vendedor ve lo leído y confirma antes de enviar.
- El número de pedido es único: no se puede subir la misma orden dos veces (salvo que la anterior esté ANULADA).
- Las bodegas de la orden se buscan por su texto exacto (`BODEGA #1`, `BODEGA #5 MONTOYA`) en la tabla `bodegas`. Una bodega desconocida bloquea el envío y avisa al admin.
- Se guarda el PDF original (sección 6, regla 4).
- **[PENDIENTE]** Probado solo con una orden de una página. Antes de dar la fase 1 por terminada, probar con 5 a 10 órdenes reales, incluyendo una de varias páginas.

### 4.7 Datos que agrega el vendedor

- **Transporte:** del catálogo `transportistas` (empresa de envíos, bus, cliente retira o reparto propio) y destino.
- **Hora requerida de salida.**
- Prioridad (normal o urgente) y notas, opcionales.
- **[PENDIENTE]** Falta saber si la forma de pago en ventas de contado (transferencia o pago en la entrega) y quién paga el flete (cliente o empresa) vienen en la orden o los indica el vendedor. Crear los campos (`forma_pago`, `flete_paga`) y dejarlos opcionales en el formulario hasta confirmar. Por defecto el flete lo paga el cliente.

### 4.8 Alerta de tiempo

Se calcula al enviar la orden y se recalcula en cada pantalla:

```
tiempo_necesario = minutos_preparacion + minutos_traslado_bodega (si alguna línea requiere traslado)
                 + minutos_facturacion + minutos_empaque + transportista.minutos_traslado
margen = hora_requerida - ahora - tiempo_necesario
```

- **Rojo (en riesgo):** `margen < 0`, o la orden se envía después de `hora_corte` y `hora_requerida - ahora < minutos_urgencia`.
- **Amarillo (justo):** `margen < margen_amarillo`.
- **Verde:** en otro caso.

**Avisa pero no bloquea.** En rojo, el vendedor debe confirmar o cambiar la hora. Si confirma, queda `riesgo_confirmado = 1` y el envío se muestra en rojo a bodega y paquetería.

Si el transportista tiene horarios, se sugiere la próxima salida y se avisa si ya no se llega a la hora límite de recepción.

Todos los valores están en `configuracion` y los edita el admin. Valores iniciales: `hora_corte=15:00`, `minutos_urgencia=60`. **[PENDIENTE]** `minutos_preparacion`, `minutos_traslado_bodega`, `minutos_facturacion`, `minutos_empaque` y `margen_amarillo` son estimados iniciales que debe dar la empresa; más adelante se calcularán con los tiempos reales registrados.

### 4.9 Empaque

- Paquetería registra bultos y peso e imprime la etiqueta con: número de pedido, números de factura, cliente, destino, transporte, bultos y, si es pago en la entrega, **"COBRAR"** destacado.

### 4.10 Rutas del mensajero

- Paquetería agrupa envíos EMPACADOS en una **ruta** del día y la asigna a un mensajero.
- La ruta tiene **paradas**, una por transportista o terminal, con su orden. Varios envíos pueden ir en una misma parada.
- El mensajero, desde el celular:
  - ve sus paradas en orden, con dirección, enlace al mapa y hora límite;
  - marca "Salí" (la ruta pasa a EN_CURSO y sus envíos a EN_RUTA);
  - en cada parada registra la hora de entrega, la guía o encomienda, **una foto de lo entregado**, **una foto del recibo** y el monto del recibo. Si el recibo cubre varios envíos, puede detallar el monto por envío (`envios.flete_monto`);
  - al registrar la parada, sus envíos pasan a ENTREGADA.
- Paquetería puede reordenar las paradas si una hora límite está en riesgo.
- **[PENDIENTE]** Falta definir quién cobra cuando el pago es en la entrega (transporte, mensajero u otro) y cómo se registra. No construir el registro de cobros hasta definirlo.

## 5. Base de datos (MySQL)

Tablas nuevas. **La aplicación ya tiene usuarios y accesos:** reutilizarlos y agregarles el rol del módulo y la bodega asignada, en lugar de crear `ent_usuarios` (que queda abajo solo como referencia de los campos necesarios). El prefijo `ent_` evita choques con tablas existentes.

```sql
CREATE TABLE ent_usuarios (
  id INT AUTO_INCREMENT PRIMARY KEY,
  nombre VARCHAR(80) NOT NULL,
  pin_hash VARCHAR(255) NOT NULL,
  rol ENUM('vendedor','bodega','coordinador_bodega','paqueteria','mensajero','admin') NOT NULL,
  bodega_id INT NULL,                       -- solo rol bodega (bodeguero)
  activo TINYINT(1) NOT NULL DEFAULT 1
);

CREATE TABLE ent_bodegas (
  id INT AUTO_INCREMENT PRIMARY KEY,
  nombre_en_orden VARCHAR(60) NOT NULL UNIQUE,   -- texto exacto del PDF: 'BODEGA #1'
  nombre VARCHAR(60) NOT NULL,
  requiere_traslado TINYINT(1) NOT NULL DEFAULT 0,  -- lo define paquetería
  bodega_traslado_id INT NULL,              -- a qué bodega se traslada; [PENDIENTE] confirmar
  activo TINYINT(1) NOT NULL DEFAULT 1,
  FOREIGN KEY (bodega_traslado_id) REFERENCES ent_bodegas(id)
);

CREATE TABLE ent_transportistas (
  id INT AUTO_INCREMENT PRIMARY KEY,
  nombre VARCHAR(100) NOT NULL,
  tipo ENUM('empresa','bus','retira_cliente','reparto_propio') NOT NULL,
  direccion VARCHAR(255),
  mapa_url VARCHAR(255),
  telefono VARCHAR(30),
  minutos_traslado INT,
  hora_limite_recepcion TIME,
  activo TINYINT(1) NOT NULL DEFAULT 1
);

CREATE TABLE ent_transportista_horarios (
  id INT AUTO_INCREMENT PRIMARY KEY,
  transportista_id INT NOT NULL,
  destino VARCHAR(100) NOT NULL,
  hora_salida TIME NOT NULL,
  dias VARCHAR(30),                         -- 'L,M,X,J,V,S'
  FOREIGN KEY (transportista_id) REFERENCES ent_transportistas(id)
);

CREATE TABLE ent_rutas (
  id INT AUTO_INCREMENT PRIMARY KEY,
  mensajero_id INT NOT NULL,
  fecha DATE NOT NULL,
  estado ENUM('ARMADA','EN_CURSO','CERRADA') NOT NULL DEFAULT 'ARMADA',
  salida DATETIME,
  regreso DATETIME,
  FOREIGN KEY (mensajero_id) REFERENCES ent_usuarios(id)
);

CREATE TABLE ent_ruta_paradas (
  id INT AUTO_INCREMENT PRIMARY KEY,
  ruta_id INT NOT NULL,
  transportista_id INT NOT NULL,
  orden INT NOT NULL,
  guia VARCHAR(50),
  monto_recibo DECIMAL(10,2),
  foto_entrega_url VARCHAR(255),
  foto_recibo_url VARCHAR(255),
  entregado_en DATETIME,
  FOREIGN KEY (ruta_id) REFERENCES ent_rutas(id),
  FOREIGN KEY (transportista_id) REFERENCES ent_transportistas(id)
);

CREATE TABLE ent_envios (
  id INT AUTO_INCREMENT PRIMARY KEY,
  pedido_num VARCHAR(20) NOT NULL,          -- de la orden; único entre envíos no anulados (validar en código)
  cliente VARCHAR(150) NOT NULL,
  terminos VARCHAR(30) NOT NULL,            -- 'CONTADO', crédito…
  fecha_pedido DATE,
  ingresado_por_orden VARCHAR(60),          -- campo "Ingresado Por" del PDF
  subtotal_orden DECIMAL(12,2),
  pdf_url VARCHAR(255),
  vendedor_id INT NOT NULL,
  transportista_id INT NOT NULL,
  destino VARCHAR(100),
  hora_requerida DATETIME NOT NULL,
  prioridad ENUM('normal','urgente') NOT NULL DEFAULT 'normal',
  riesgo ENUM('verde','amarillo','rojo'),
  riesgo_confirmado TINYINT(1) NOT NULL DEFAULT 0,
  forma_pago ENUM('transferencia','contra_entrega') NULL,   -- [PENDIENTE] solo contado
  flete_paga ENUM('cliente','empresa') NOT NULL DEFAULT 'cliente',
  bultos INT,
  peso_kg DECIMAL(8,2),
  parada_id INT NULL,
  flete_monto DECIMAL(10,2),
  estado ENUM('ORDEN_ENVIADA','EN_REVISION','POR_FACTURAR','FACTURADA','LISTA_PARA_RETIRO',
              'RETIRADA','EMPACADA','EN_RUTA','ENTREGADA','ANULADA') NOT NULL DEFAULT 'ORDEN_ENVIADA',
  motivo_anulacion VARCHAR(255),
  notas TEXT,
  creado_en DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  actualizado_en DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX (pedido_num), INDEX (estado),
  FOREIGN KEY (vendedor_id) REFERENCES ent_usuarios(id),
  FOREIGN KEY (transportista_id) REFERENCES ent_transportistas(id),
  FOREIGN KEY (parada_id) REFERENCES ent_ruta_paradas(id)
);

CREATE TABLE ent_envio_lineas (
  id INT AUTO_INCREMENT PRIMARY KEY,
  envio_id INT NOT NULL,
  linea INT NOT NULL,
  bodega_id INT NOT NULL,
  codigo VARCHAR(50) NOT NULL,
  producto VARCHAR(150),
  medida VARCHAR(60),
  marca VARCHAR(40),
  cantidad DECIMAL(10,2) NOT NULL,
  cant_disponible DECIMAL(10,2),            -- al marcar FALTANTE
  existencia ENUM('PENDIENTE','CONFIRMADA','FALTANTE','QUITADA') NOT NULL DEFAULT 'PENDIENTE',
  existencia_por INT,                       -- coordinador de bodega
  existencia_en DATETIME,
  decision ENUM('quitar','cambiar_bodega','esperar') NULL,
  traslado_estado ENUM('NO_REQUIERE','PENDIENTE','APLICADO','EN_CAMINO','RECIBIDO')
                  NOT NULL DEFAULT 'NO_REQUIERE',
  traslado_num VARCHAR(30),
  traslado_aplicado_en DATETIME,
  traslado_recibido_en DATETIME,
  preparacion ENUM('PENDIENTE','EN_PREPARACION','LISTA','RETIRADA') NOT NULL DEFAULT 'PENDIENTE',
  responsable ENUM('bodeguero','coordinador') NOT NULL,   -- de la existencia; según requiere_traslado al crear la línea
  bodega_prepara_id INT NOT NULL,           -- bodega_id, o la destino del traslado; su bodeguero prepara
  actualizado_por INT,
  actualizado_en DATETIME,
  FOREIGN KEY (envio_id) REFERENCES ent_envios(id),
  FOREIGN KEY (bodega_id) REFERENCES ent_bodegas(id)
);

CREATE TABLE ent_envio_facturas (          -- una factura por bodega de facturación
  id INT AUTO_INCREMENT PRIMARY KEY,
  envio_id INT NOT NULL,
  bodega_id INT NOT NULL,
  factura_num VARCHAR(30) NOT NULL,
  registrada_por INT NOT NULL,
  registrada_en DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE (envio_id, bodega_id),
  FOREIGN KEY (envio_id) REFERENCES ent_envios(id),
  FOREIGN KEY (bodega_id) REFERENCES ent_bodegas(id)
);

CREATE TABLE ent_envio_eventos (           -- historial: cada cambio de estado
  id INT AUTO_INCREMENT PRIMARY KEY,
  envio_id INT NOT NULL,
  linea_id INT NULL,
  usuario_id INT NOT NULL,
  de_estado VARCHAR(20),
  a_estado VARCHAR(20) NOT NULL,
  comentario VARCHAR(255),
  creado_en DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY (envio_id) REFERENCES ent_envios(id)
);

CREATE TABLE ent_configuracion (
  clave VARCHAR(50) PRIMARY KEY,
  valor VARCHAR(50) NOT NULL
);
INSERT INTO ent_configuracion VALUES
  ('hora_corte','15:00'), ('minutos_urgencia','60');
-- [PENDIENTE] minutos_preparacion, minutos_traslado_bodega, minutos_facturacion, minutos_empaque, margen_amarillo
```

Ventas perdidas: las líneas QUITADA y la parte no disponible de una línea parcial. Si la aplicación ya tiene un registro de ventas perdidas, alimentarlo desde aquí.

## 6. Reglas técnicas

1. **Toda transición de estado pasa por una sola función** que valida el rol y el estado de origen, actualiza el estado y escribe en `ent_envio_eventos`. Nada cambia `estado` por fuera de ella.
2. Los estados automáticos del envío (EN_REVISION, POR_FACTURAR, FACTURADA, LISTA_PARA_RETIRO, RETIRADA) se recalculan desde las líneas y las facturas dentro de esa misma función. Cada cambio de `existencia`, `traslado_estado` o `preparacion` también se guarda en `ent_envio_eventos`.
3. **Permisos por rol** en el servidor, no solo ocultando botones. La existencia la cambia solo el responsable de la línea (el bodeguero de esa bodega o el coordinador); el traslado, solo el coordinador; la preparación, solo el bodeguero de `bodega_prepara_id`, y en líneas con traslado solo después de RECIBIDO; `requiere_traslado`, solo paquetería o admin. El vendedor solo ve y decide sus envíos.
4. **Fotos y PDF** se guardan en almacenamiento de archivos (carpeta del servidor o almacenamiento en la nube, según dónde esté alojada la app). En MySQL solo va la ruta o URL. Comprimir las fotos del celular antes de guardarlas.
5. Fechas y horas en hora de Nicaragua (UTC-6).
6. **Notificaciones** dentro del módulo: contador y lista de pendientes por rol, con actualización automática. Ejemplos: cada bodeguero recibe las líneas nuevas de su bodega y la mercadería trasladada que llegó para preparar; el coordinador recibe las de bodegas con traslado y los traslados por aplicar o por recibir; el vendedor recibe faltantes y órdenes listas para facturar; paquetería recibe órdenes nuevas, faltantes, pasos atrasados y las listas para retiro. Avisos por WhatsApp, más adelante y opcionales.

## 7. Pantallas

| Rol | Pantallas |
|---|---|
| Vendedor | Subir orden (PDF → vista previa → transporte y hora → alerta → enviar). Mis envíos con estado. Faltantes por decidir. Listos para facturar: ajustes a la orden y registro de una factura por bodega |
| Bodeguero | Líneas de su bodega por prioridad y riesgo; confirmar existencia o faltante; preparación, incluida la mercadería trasladada recibida; por línea o toda su parte de la orden |
| Coordinador de bodega | Líneas de bodegas con traslado por prioridad y riesgo; confirmar existencia o faltante; traslados por aplicar, en camino y por recibir |
| Paquetería | Tablero de seguimiento de órdenes y preparación por bodega; bodegas que aplican traslado; retiro por bodega; empaque y etiqueta; armar y asignar rutas |
| Mensajero | Mis rutas del día; paradas; "Salí"; registrar entrega con cámara |
| Admin | Catálogos (bodegas, transportistas, horarios, usuarios y su bodega asignada), configuración, tablero y reportes |
| Todos | Detalle del envío con su historial |

## 8. Reportes (fase 4)

- Tiempo por etapa: orden → lista (separando la espera por traslados), lista → facturada, facturada → empacada, empacada → entregada.
- Porcentaje de envíos entregados antes de la hora requerida.
- Faltantes por bodega y por producto (ventas perdidas).
- Fletes por transportista, separando los que paga el cliente de los que paga la empresa.
- Conciliación: facturas registradas en `ent_envio_facturas` contra la facturación del día que se carga a la aplicación.

## 9. Fases y criterios de aceptación

**Fase 1 — Base**
- Usuarios con rol y PIN; catálogos de bodegas y transportistas; configuración.
- Subir el PDF, leerlo, validar el cuadre y la bodega, vista previa, transporte y hora, alerta de tiempo, guardar el envío con sus líneas.
- Listo cuando: el pedido 270091 y las demás órdenes de prueba se cargan con sus líneas y bodegas correctas, una orden repetida se rechaza y la alerta se pone en rojo después de las 3 p. m. con menos de 1 hora.

**Fase 2 — Operación**
- Pantallas del bodeguero (existencias y preparación) y del coordinador de bodega (existencias y traslados), tablero de seguimiento de paquetería, faltantes con decisión del vendedor, facturación por bodega, retiro y empaque con etiqueta; historial y notificaciones.
- Listo cuando: una orden con productos de dos bodegas, una de ellas con traslado, recorre ORDEN_ENVIADA → EMPACADA con usuarios distintos; el vendedor no puede registrar facturas mientras falte una existencia o el traslado no esté aplicado; paquetería no puede retirar hasta que la mercadería trasladada esté RECIBIDA; el envío pasa a FACTURADA solo con una factura por bodega; y un faltante quitado aparece como venta perdida y no bloquea la facturación.

**Fase 3 — Entrega**
- Rutas, paradas, pantalla del mensajero, fotos y recibos.
- Listo cuando: una ruta con dos paradas y tres envíos se completa desde un celular y cada envío queda ENTREGADO con fotos y monto.

**Fase 4 — Control**
- Tablero y reportes de la sección 8.

### 9.1 Construcción por partes

Cada parte es pequeña, se revisa en la aplicación y se aprueba antes de empezar la siguiente. **Al terminar una parte, parar y esperar la revisión.**

| # | Parte | Qué se revisa al terminar | Necesita del negocio |
|---|---|---|---|
| 1 | Revisar la aplicación y escribir la migración SQL (sin aplicarla) | Informe de cómo está hecha la app y el archivo SQL de las tablas | — |
| 2 | Agregar los roles del módulo al sistema de usuarios y accesos que ya tiene la app (no crear otro login), y los permisos | Entrar con cada rol y ver solo su menú | — |
| 3 | Catálogos: bodegas (traslado y bodeguero asignado), transportistas, horarios, configuración | Cargar las bodegas reales y marcar cuáles aplican traslado | Pendiente 4 |
| 4 | Subir la orden: PDF → lectura → control de cuadre → vista previa → transporte y hora → guardar | Subir órdenes reales y comparar lo guardado con el PDF | Pendiente 5 |
| 5 | Alerta de tiempo (verde, amarillo, rojo) | Probar órdenes a distintas horas | Pendiente 3 |
| 6 | Pantallas del bodeguero y del coordinador: existencias y faltantes; tablero de seguimiento de paquetería | Confirmar una orden de dos bodegas con usuarios distintos | — |
| 7 | Decisión del vendedor ante faltantes, y facturación por bodega | Quitar un faltante y registrar las facturas | Pendientes 6 y 7 |
| 8 | Traslados: aplicado, en camino y recibido | Una orden con traslado no deja facturar hasta aplicarlo | — |
| 9 | Preparación y retiro por bodega | El retiro solo se habilita con todo preparado y lo trasladado recibido | — |
| 10 | Empaque y etiqueta | Imprimir una etiqueta real | Pendiente 1 |
| 11 | Rutas y pantalla del mensajero, con fotos | Hacer una ruta de prueba desde un celular | Pendiente 2 |
| 12 | Tablero y reportes | Revisar los reportes con datos de las pruebas | — |

Las partes 1 a 5 forman la fase 1, las 6 a 10 la fase 2, la 11 la fase 3 y la 12 la fase 4.

## 10. Pendientes del negocio

1. ¿La forma de pago (contado) y quién paga el flete vienen en la orden o los indica el vendedor?
2. ¿Quién cobra el pago en la entrega y cómo se registra?
3. Estimados de minutos de preparación, traslado, facturación y empaque, y margen para el amarillo.
4. Lista de bodegas tal como aparecen en las órdenes, cuáles requieren traslado y hacia qué bodega (lo carga paquetería), y el bodeguero de cada una; mensajeros; transportistas con terminales y horarios.
5. Órdenes PDF adicionales para probar la lectura, incluyendo una de varias páginas.
6. Después del traslado, ¿la factura sale de la bodega destino?
7. ¿Se factura siempre toda la orden junta, o una bodega puede facturarse mientras otra espera?
