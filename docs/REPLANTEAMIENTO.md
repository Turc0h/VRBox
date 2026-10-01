# Replanteamiento de VRBox

## Producto

VRBox es una herramienta de mostrador para un local de gaming. La prioridad es resolver una reserva o una venta en pocos toques, con información legible a distancia y sin perder trazabilidad del dinero o del inventario. Está pensada primero para una computadora y un operador; sincronizar varias cajas queda fuera de este corte.

## Decisión principal: una reserva puede ocupar varias estaciones

Una reserva representa una sola intención comercial: cliente, fecha, hora y conjunto de estaciones. Cada estación genera una línea hija independiente para poder controlar ocupación, precio y cobro por equipo. La cabecera agrupa esas líneas y guarda el total cotizado.

Crear el grupo se hace en una transacción inmediata de SQLite:

1. Validar fecha, cliente, hora y estaciones seleccionadas.
2. Bloquear brevemente la escritura y comprobar disponibilidad de todas las estaciones.
3. Guardar cabecera y todas las líneas, o cancelar todo si una está ocupada.
4. Guardar una copia del precio de cada estación para que una edición futura de tarifas no altere reservas existentes.

La jornada usa bloques fijos de 60 minutos, de 16:00 a 00:00. Para el puesto, varios equipos pueden reservarse para la misma persona y hora; el importe total es la suma de sus tarifas después de una promoción, si corresponde. El cobro sigue separado por estación y medio de pago, que permite conciliación exacta.

## Flujo de mostrador

1. Indicar fecha y nombre del cliente.
2. Marcar una o más tarjetas de estación, de forma individual o mediante accesos rápidos por categoría.
3. Elegir una hora disponible para todas las estaciones seleccionadas.
4. Revisar estaciones, cantidad y cotización total antes de confirmar.
5. En el cobro, asignar el juego del turno; el dato aparece en la agenda y el cierre por juego.
6. Las ofertas activas se aplican automáticamente por hora de inicio. Se admite una oferta por grupo y se usa la que maximiza el ahorro; las estaciones bonificadas son las de menor precio, y total/descuento quedan congelados con la reserva.
7. La seña se registra hasta llegar al 50% del total; el resto puede cobrarse en un paso o mediante pagos parciales. Cada cobro se asocia a un movimiento de venta con fecha y medio, para conciliar efectivo y transferencias.
8. Completar la sesión marca el equipo como usado y el juego elegido, sin volver a cobrar. La tarjeta de reserva muestra cuánto se recibió y cuánto queda.

La agenda usa las horas como filas y las estaciones como columnas. Una tarjeta comunica libre, reservada o completada; plata, oro, PS5 y VR usan grupos visuales consistentes. El Punto de venta se mantiene separado: producto en tarjeta → carrito → cantidad → medio de pago → cobrar.

## Límites técnicos elegidos

- Mantener una app nativa Rust con `eframe/egui`; evita servidor, login remoto y runtime JavaScript. La interfaz usa temas oscuros, estados claros, selección por botones y transiciones breves en lugar de depender de animaciones decorativas permanentes.
- La navegación usa etiquetas breves y el cronómetro dibuja su aro con primitivas vectoriales de egui; la UI no depende de glifos emoji instalados en Windows.
- SQLite local con WAL. Las reservas agrupadas tienen una pequeña capa de aplicación propia (`src/reservations.rs`) para que sus reglas transaccionales no dependan de la UI.
- El dinero persistido está en centavos enteros. Precios de reserva se copian al crearla; ventas capturan precio unitario vigente y cantidad.
- Excel es una interfaz de importación/exportación para el personal, no la base de datos.
- El instalador, varios puestos sincronizados, roles, anulación auditada y emisión fiscal requieren decisiones y trabajo separados antes de operar comercialmente.

## Próximas etapas de ingeniería

1. **Cerrar la experiencia diaria:** navegación de fechas sin escribir ISO, filtros de agenda, motivo de cancelación y partidas de grupos visibles también en ventas.
2. **Completar inventario:** validar plantilla y previsualizar altas/reposiciones antes de importar; impedir cambios parciales; añadir historial de compras, mermas y costo.
3. **Auditoría operativa:** apertura de caja, gastos/retiros, pagos mixtos, anulaciones con motivo, usuario responsable y reimpresión/exportación de cierres.
4. **Separar la aplicación en módulos:** `domain`, `application`, `storage`, `presentation` y `reporting`; migrar las consultas restantes fuera de `main.rs` y versionar migraciones SQLite.
5. **Durabilidad:** copia de seguridad desde la interfaz, restauración comprobada, exportes fechados y documentación de recuperación.
6. **Producción:** instalador firmado, política de retención de datos, revisión de dependencias, accesibilidad de controles y validación fiscal según el negocio.

## Hecho en este rediseño

- Selección múltiple de estaciones más accesos rápidos Plata/Oro/PS5/VR.
- Cotización acumulada por estación y selección de hora libre para todo el grupo.
- Persistencia atómica con tabla de cabecera y líneas por estación.
- Lista diaria agrupada para que una reserva grupal no parezca una serie de clientes distintos.
- Chequeo transaccional de colisiones para evitar reservas parciales o dobles entre instancias.
- Editor de promociones 2 por 1, 3 por 2 y similares, con categoría y rango de hora inclusivo; cotización, ahorro e importe neto persistidos en grupo y líneas de reserva.
- Cronómetros PS5/VR persistentes en SQLite con cuenta regresiva, arco de progreso, controles para ajustar en bloques de cinco minutos e inicio/eliminación desde una página propia.
- Cobros de seña, saldo y pagos parciales ligados al grupo o reserva individual, con efectivo/transferencia; las devoluciones por cancelación se registran como salidas con importe negativo en caja.
- Edición transaccional de grupos antes de iniciar, con comprobación de conflictos, protección contra bajar el precio bajo lo recibido e historial de cambios; cancelación confirmada y devolución asociada.
