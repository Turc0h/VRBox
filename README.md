# VRBox

Aplicación de escritorio local para agenda, reservas, ventas rápidas, stock y cierre de caja de un ciber con simuladores.

## Estado de la versión actual

- Rust nativo con `eframe/egui`, SQLite embebido y exportación XLSX. No necesita cuenta, nube ni servidor para operar.
- Agenda diaria reordenada con horas en filas y los equipos en columnas: 6 simuladores Plata, 2 simuladores Oro, 2 PS5 y 2 VR.
- Reserva grupal express: elegir varias estaciones, incluso de categorías distintas, en el mismo horario y para el mismo cliente. Accesos rápidos permiten sumar una categoría completa; la cotización suma la tarifa de cada equipo y todos los puestos se guardan juntos o ninguno. Cada grupo queda identificado y aparece agrupado en la lista diaria.
- Pagos de reserva dentro de cada turno: registrar seña hasta completar el 50%, pagos parciales o cobrar el saldo. Cada movimiento elige efectivo o transferencia y aparece en el cierre de la fecha en que se recibió. Completar una sesión ya no vuelve a cobrar el total.
- Acciones para modificar grupos no iniciados, cancelar con confirmación y devolver lo recibido por el medio elegido. La modificación conserva pagos y bloquea cambios que dejarían la reserva por debajo de lo ya cobrado; la cancelación deja historial y registra devolución.
- Promociones configurables 2 por 1, 3 por 2 y similares: nombre, categoría y horario de inicio incluido. Se aplica automáticamente a la reserva y queda registrado el total y el descuento. Se elige la oferta que más ahorra; en cada paquete se bonifica el puesto de menor precio.
- Cronómetros persistentes para PS5 y VR: cuenta regresiva con aro de progreso dibujado por la app, nombre opcional, ajuste de cinco minutos y eliminación. Se conservan al cerrar y volver a abrir el programa.
- Navegación e indicadores con etiquetas tipográficas cortas y controles dibujados por egui; no depende de soporte de emojis ni de instalar una fuente adicional.
- Turnos fijos de 60 minutos. 6 Plata, 2 Oro, 2 PS5 y 2 VR tienen tarifas editables por categoría.
- En el cobro de una sesión se puede asignar un juego único al turno; queda visible en la franja horaria de la agenda y en el reporte por juego.
- Punto de venta con tarjetas de producto, carrito, cantidades y cobro separado en efectivo o transferencia.
- Plantilla Excel exportable/importable para altas masivas y reposición; costo unitario más recargo porcentual por producto calcula el precio sugerido de venta. El recargo se suma sobre costo; no equivale al margen bruto sobre el precio de venta.
- Cierre diario con valores esperados, recuento declarado, diferencias y archivo Excel con resumen, ventas, hoja de pagos de reservas, agenda, futuro, horas ocupadas y uso completado por juego. El Excel separa ingresos y devoluciones por medio de pago de las horas utilizadas por juego.
- Registro básico de cumpleaños y eventos próximos.

## Ejecutar

Requiere Rust estable y Cargo. Desde esta carpeta:

```powershell
cargo run --release
```

En Windows, la base queda en `%LOCALAPPDATA%\VRBox\vrbox.sqlite3`; los cierres Excel se exportan a la misma carpeta. Mantener respaldos regulares de esa carpeta. La aplicación no transmite datos.

## Modelo de dinero

Los montos de pantalla se ingresan en pesos y se persisten en centavos enteros. Una reserva grupal guarda la cabecera de cliente/horario/total y una línea por estación con su precio congelado. Los cierres son por fecha operativa local configurada en la interfaz; las reservas futuras se guardan por fecha y recurso.

## Límites conocidos de este corte

- Los turnos son de una hora fija. Todavía no hay una interfaz de anulación/modificación con motivo.
- La vista de agenda actual es diaria; el campo de fecha permite trabajar con reservas futuras.
- El cierre recoge las ventas efectivamente registradas. Las reservas sin cobrar aparecen como pendientes y no se suman a efectivo/transferencia.
- Faltan usuarios/roles, auditoría inmutable, backups dentro de la app, pruebas automatizadas, adjuntos de imágenes, personalización de puestos/categorías, vigencia por fechas/días de semana para ofertas, sincronización multi-equipo e integración fiscal. Los cronómetros actualmente no tienen pausa ni alerta sonora. No usar aún como sistema fiscal ni como única copia contable.

## Criterios de producto

- Offline-first, datos del comercio bajo control del cliente, exportación legible e interoperabilidad.
- Guardar operaciones con transacciones de base de datos; no borrar ventas para corregirlas: registrar reversos auditables.
- Minimizar datos personales: nombre y contacto solo si son necesarios para reservar o avisar un evento; definir acceso, retención y borrado conforme a las obligaciones aplicables.
- Un reporte XLSX sirve para gestión interna. En Argentina, la obligación de respaldar operaciones con factura electrónica o controlador fiscal depende del caso y debe validarse con ARCA/asesoría contable antes de producción.
- La licencia del producto queda pendiente de decisión antes de distribuirlo a clientes.

## Hoja de ruta propuesta

1. Endurecer agenda y reglas de horario; selector de fecha y prueba de solapamientos.
2. Ampliar promociones: vigencia por fecha y días de semana, combinaciones con tarifas especiales y registro en los informes de cierre.
3. Auditoría de caja: turnos/cajeros, apertura, retiros, gastos, pagos mixtos, cancelaciones y reembolsos con motivo.
4. Productos con imagen, mínimos de stock, movimientos de compra/merma, costo y margen.
5. Alertas anuales de cumpleaños, eventos, ocupación por hora y exportaciones con más hojas/gráficos.
6. Copias de seguridad verificables/restaurables, privacidad/retención, roles y revisión de seguridad/dependencias.
7. Instalador para Windows y documentación para usuarios; evaluar impresión y facturación fiscal según jurisdicción.

Las decisiones y la secuencia de trabajo del rediseño están en [docs/REPLANTEAMIENTO.md](docs/REPLANTEAMIENTO.md).

## Referencias técnicas

- [The Rust Book](https://doc.rust-lang.org/book/) y [Cargo Book](https://doc.rust-lang.org/cargo/): lenguaje, dependencias y compilación reproducible.
- [egui/eframe en GitHub](https://github.com/emilk/egui): GUI portable, escrita en Rust. `eframe` tiene dependencias gráficas; es un compromiso entre simplicidad de despliegue y tamaño del ejecutable.
- [rusqlite](https://github.com/rusqlite/rusqlite): bindings para SQLite; la opción `bundled` compila una SQLite propia y reduce requisitos en el equipo de destino.
- [rust_xlsxwriter](https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/) y [calamine](https://docs.rs/calamine/latest/calamine/): escritura y lectura local de libros XLSX.
- [rfd](https://docs.rs/rfd/latest/rfd/): diálogos de archivos del sistema para guardar e importar la plantilla.
- [RustPOS](https://github.com/dividebysandwich/rustpos): referencia de flujos POS en Rust con SQLite, inventario, roles y reportes; su propio README lo clasifica como prueba de concepto y advierte que no está certificado para producción.
- [ARCA: emisión y autorización](https://www.arca.gob.ar/fe/emision-autorizacion/sujetos.asp): referencia oficial para revisar el encuadre de comprobantes; no sustituye asesoramiento fiscal.

