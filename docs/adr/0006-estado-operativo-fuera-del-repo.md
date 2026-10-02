# 0006. El estado operativo vive fuera del repositorio

## Estado
Aceptada (Fase 0, 2026-10-01)

## Contexto

El proyecto era a la vez codigo y almacen de estado (diagnostico P8):

```
chrome_profile_whatsapp_session/   sesion de WhatsApp Web (credenciales)
.wdm/ .wdm_cache/                   cache de chromedriver (~40 MB)
chromedriver.exe (.bak)             binarios (~43 MB)
logs/                               306 archivos de log
Img informe/                        capturas del informe
errorImages/                        screenshots de error
graphify-out/                       salida de herramienta auxiliar
```

Tres problemas: repositorio pesado, riesgo de filtrar una sesion de
WhatsApp, y `.gitignore` desactualizado que no cubria la mitad de eso.

## Decision

1. `.gitignore` cubre todo el estado operativo y los datos de negocio
   (`*.xlsx`, `Archivos Validacion/`). El repo versiona codigo y configuracion,
   nada mas.
2. Las rutas pasan a ser configurables por variable de entorno, con destino
   fuera del repo por defecto:

   | Variable     | Valor por defecto                    | Reemplaza a              |
   |--------------|--------------------------------------|--------------------------|
   | `STATE_DIR`  | `%LOCALAPPDATA%\GeneracionInsumos`   | perfiles, cache, imagenes|
   | `LOGS_DIR`   | `%LOCALAPPDATA%\GeneracionInsumos\logs` | `logs/`               |

   El destino por defecto se deriva de `LOCALAPPDATA`, que en Windows ya es la
   convencion para datos de aplicacion fuera del perfil de usuario.
3. Se migro el contenido existente a la nueva ubicacion (306 logs, cache de
   drivers, perfil de WhatsApp con su sesion intacta).

## Consecuencias

Se gana: repo liviano, la sesion de WhatsApp no se versiona por accidente, y
las rutas quedan en un solo lugar configurable.

Se paga: nada, si se respeta la convencion. Si `LOCALAPPDATA` no existiera, se
cae a `Path.home()`.

Queda bloqueado: `chromedriver.exe` sigue viviendo en la raiz del proyecto
porque `core/chromedriver_utils.py` lo usa como driver local validado. Se
mueve al escribir el adaptador de Fase 5.

## Nota de seguridad

El perfil de WhatsApp Web contiene la sesion iniciada. Nunca debe entrar al
repo. `.gitignore` tiene `chrome_profile_*/` y `.env`.
