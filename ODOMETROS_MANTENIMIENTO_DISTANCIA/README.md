# Odómetros de mantenimiento por distancia

Módulo local para estimar el kilometraje actual desde una orden de mantenimiento.

## Datos requeridos en el Excel

- `CHASIS`
- `FECHA_MANTENIMIENTO`
- `KM_MANTENIMIENTO`

El módulo busca el ID Flespi e IMEI por chasis en Vapor. Luego suma `mileage()` por día desde `FECHA_MANTENIMIENTO` hasta hoy y genera `KM_DEBERIA_ESTAR_HOY`.

No utiliza la calculadora Flespi.

## Ejecución

1. Copia `config.example.json` como `config.json` y completa las credenciales locales.
2. Instala dependencias: `pip install -r requirements.txt`.
3. Ejecuta `INICIAR_ODOMETROS_DISTANCIA.bat`.
4. Abre `http://127.0.0.1:8061`.

El Excel resultante permite analizar y descargar los cálculos. La actualización de Flespi se realiza solamente al seleccionar y confirmar los registros aprobados.
