# Anotador diario

Versión 1.0 — 2026-09-30 12:00 ARG

Anotador personal de tareas: hojas por tema, etiquetas de fecha (Hoy, Mañana, Semana entrante, fecha propia), "hablar con" por tarea, notas, prioridad y envío del resumen diario por mail con un PDF por hoja.

- **Frontend:** `index.html` (GitHub Pages).
- **Backend:** `apps-script/Code.gs` + `apps-script/appsscript.json`, vinculado a la planilla "Anotador diario".

## Instalación del backend

1. En la planilla: Extensiones > Apps Script. Pegar `Code.gs` y `appsscript.json` (activar "Mostrar archivo de manifiesto" en Configuración del proyecto).
2. Ejecutar `setup()` y autorizar. Crea las pestañas Tareas, Hojas, Config y Envios, y genera la clave de acceso.
3. Implementar > Nueva implementación > Aplicación web. Ejecutar como: yo. Acceso: cualquier usuario.
4. Pegar la URL `/exec` en `API_DEFAULT` dentro de `index.html`.
5. La clave se ve en la planilla: menú **Anotador > Ver clave de acceso**. Se ingresa una vez por dispositivo.

Al cambiar `Code.gs`: Implementar > Administrar implementaciones > editar > Nueva versión (la URL no cambia).

## Uso

- Enter agrega la tarea. Tecla `N` o `/` enfoca el campo desde cualquier lugar.
- Las fechas son reales: lo marcado "Mañana" pasa solo a HOY al día siguiente; lo no hecho queda como Vencida, con "Pasar todas a hoy".
- "Hoy en todas" junta lo de hoy y lo vencido de todas las hojas.
- Clic en una tarea para editarla (texto, fecha, persona, nota, hoja, prioridad).
- Enviar resumen: elegís hojas, CC (las direcciones guardadas vienen sin tildar) y comentario.
- Ajustes: destinatario, lista de personas y envío automático de lunes a viernes a la hora elegida.
- Funciona sin conexión: los cambios quedan en cola y se guardan al volver.
