# Automatización total con VBA (un clic para preparar el Excel)

Sí es posible automatizarlo. Con este módulo VBA puedes crear toda la estructura del archivo y trabajar con botones.

## ¿Qué hace la macro?
La macro `SetupJMVisionWorkbook` crea automáticamente:
- Hojas: `CONFIG`, `CLIENTES`, `PRODUCTOS`, `COTIZACION`, `HISTORICO_COTIZACIONES`, `VENTAS`, `DASHBOARD`.
- Encabezados y formatos base.
- Fórmulas de búsqueda y cálculo (incluye búsqueda por palabra clave como `camaras`).
- Validaciones de datos para cliente, SKU y estado.
- Botones en `COTIZACION` para:
  - Guardar en histórico.
  - Convertir a venta.
  - Limpiar formulario.

## Archivo VBA incluido
- `vba/JMVISION_Automatizacion.bas`

## Pasos (solo clics)
1. Abre Excel y crea un libro nuevo.
2. Guarda como **Libro habilitado para macros**: `JMVISION_Cotizaciones.xlsm`.
3. Presiona `ALT + F11` para abrir el Editor VBA.
4. Menú **File > Import File...** e importa `vba/JMVISION_Automatizacion.bas`.
5. Cierra el editor.
6. Presiona `ALT + F8`, selecciona `SetupJMVisionWorkbook` y ejecuta.
7. Carga tus datos de `CLIENTES` y `PRODUCTOS`.
8. En `COTIZACION`, escribe por ejemplo `camaras` en `J4` y empieza la prueba.

## Botones listos para prueba
En la hoja `COTIZACION` verás 3 botones:
- **Guardar en histórico**
- **Convertir a venta**
- **Limpiar formulario**

## Flujo de prueba recomendado
1. Ejecuta `SetupJMVisionWorkbook`.
2. Carga 3 clientes y 10 productos.
3. En `COTIZACION`, busca producto por palabra clave (`camaras`).
4. Completa cantidad/descuento y valida totales.
5. Clic en **Guardar en histórico**.
6. Cambia estado a `Aprobada` en histórico.
7. Clic en **Convertir a venta**.

## Nota importante
- Este VBA usa funciones modernas como `BUSCARX` y `FILTRAR` (Excel Microsoft 365 recomendado).
