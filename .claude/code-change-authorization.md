# Directriz de Autorizacion Previa para Cambios de Codigo

## Regla Principal
**PROHIBIDO** realizar cambios en el codigo sin autorizacion explicita del usuario.

## Protocolo de Cambios

1. **Proponer primero**: Antes de implementar cualquier solucion, presentar:
   - Descripcion del problema identificado
   - Solucion propuesta con justificacion tecnica
   - Impacto en el codigo existente

2. **Esperar autorizacion**: No proceder hasta recibir visto bueno del usuario.

3. **Implementar solo lo autorizado**: Ejecutar exactamente lo acordado, sin mejoras adicionales.

## Principios de Diseno (NO negociables)

### Identificacion de Objetos
- Usar `ObjPtr()` para identificar univocamente instancias de clases VBA
- NO sustituir ObjPtr por path de ficheros u otros identificadores alternativos
- ObjPtr es la referencia canonica para tracking de objetos

### Inicializacion de Objetos
- Un objeto esta completamente inicializado cuando `Class_Initialize` termina
- NO usar flags `IsFullyInitialized` u otras variables de seguimiento de inicializacion
- Si hay problemas de inicializacion, corregir el flujo de programa, no anadir flags

### Documentacion
- Registrar TODA tarea completada en la seccion DONES de REFERENCE_NOTES.md

## Vigencia
Esta directriz aplica a partir del 2026-02-07 y permanece activa hasta nueva indicacion.
