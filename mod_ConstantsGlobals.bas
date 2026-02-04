Attribute VB_Name = "mod_ConstantsGlobals"
'TODO: BUENAS PRACTICAS DE PROGRAMACION:
' - "Patrones de diseÃ±o": CuÃ¡les de esos patrones estÃ¡n a medio implementar y cuÃ¡les deberÃ­an ser mejorados o incluso
'   aÃ±adidos a la arquitectura de cÃ³digo. PatrÃ³n MVC, ...
'   aplicando principios comoÂ Separation of Concerns, Testeable, Reutilizable, Mantenible,Â Robustez ,Claridad,
'   PrevenciÃ³n de bugs etc
' - ReutilizaciÃ³n del cÃ³digo existente. Si hay Funciones que hacen lo mismo con distinto nombre o
'   con distinto algoritmo, UNIFICARLAS
' - Usar nombres para variables y funciones conforme a las buenas prÃ¡cticas de programaciÃ³n
' - eliminar duplicidades: Si hay estructuras repetitivas Que se han implementado como una funciÃ³n, en lugar de mantener
' esas estructuras repetitivas reemplazarlas por llamadas a funciones
'HACK Uso temporal de Stop para debugging - eliminar en producciÃ³n

'TODO implementar correctamente la GESTION DE ERRORES, con "resume cleanup"s que LIMPIEN EL CONTEXTO, cuando se produzca un error...

'FIXME: DocumentaciÃ³n de subrutinas, al menos Atributos @Description en funciones clave
'  Comentarios estructurados con @Scope, @Returns, @Category...
' ************** REVISAR LOS ULTIMOS CAMBIOS DE CLAUDE: backups de hoja excel y VBA a ZIP, al listar funciones en excel...

'FIXME El dropdown no se actualiza automÃ¡ticamente tras cambios en carpeta
'BUG FindAndKillFolderWatcher puede matar procesos incorrectos si hay mÃºltiples instancias

'FIXME: estabilidad del ribbon (ver 'callbacks' etc):  (a ver si me ayuda chatgpt, o claude, a reimplementar la funcion...
' o que 'PAUSE LA ACTUALIZ DEL RIBBON', y exponga una macro  RECUPERE LA APP cuando la ejecute....

'TODO: terminar de migrar la funcionalidad del FOLDER WATCHER, REVISARLO TODO BIEN.. Y ADAPTARLO AL USO DE LA APLICACION

'TODO: terminar de implementar 'SUBCLASES DE ESTADOS'; "subclases de clsApplication", para repartir mejor la gestiÃ³n de eventos....



'FIXME: revisar CONSISTENCIA de "SincronizaciÃ³n bidireccional registro <-> configuraciÃ³n", y registro de UDFS, etc
'  (unificar en un punto el almacenamiento de datos en el registro y la configuraciÃ³n de la aplicaciÃ³n)

'TODO: **** PRIORITARIO!!! **** aÃ±adir al proceso de instalaciÃ³n, la descompresiÃ³n del COM y **su instalaciÃ³n**. e incluso...
' Â¿seria posible que la DLL COM "se cargue en memoria" cuando se abre excel, y SE BORRE EL FICHERO del sistema de archivos?
' en tal caso, en vez de instalarla al instalar el complemento, se CARGARIA, sin instalar - y se escribe y borra el registro - ,
' cada vez que se abre excel (no quiero dejar huellas de que existe)


'TODO: formularios de usuario de Excel:
' - Revisar si en los formularios conviene introducir gestiÃ³n de eventos
' - InstanciaciÃ³n de Formularios: Â¿estÃ¡n Correctamente implementadas las llamadas a los formularios? Â¿ Se instancian y destruyen
'   los formularios conforme a las mejores practicas de programacion, o se ocultan indebidamente, generando un mayor consumo
'   de memoria, recursos...?.


'TODO: FUNCIONES (clases?) DE DIAGNOSTICO: implementar en los modulos de claes FUNCIONES QUE AYUDEN A DIAGNOSTICAR EL ESTADO DE LOS ESCUCHADORES
'  DE EVENTOS (que logueen si los atributos internos (private) de esos metodos estÃ¡n  asignados, etc; y que por ejemplo en el caso de
'  clsChartEvents, tengan un atributo, que tambiÃ©n se pueda loguear, que indique **a que grafico, de que hoja**, estÃ¡ asociado
'  ese escuchador de eventos: SEGUN VAYAN ACTIVANDOSE o desactivandose ESCUCHADORES, el debug.print indique respecto a que grafico
'  se produce la conexion Y DESCONEXION.).


'FIXME AÃ±adir VALIDACION de formatos de nombre de archivo, etc
'TODO Implementar cachÃ© de oportunidades para mejorar rendimiento
'TODO: LÃGICA DE NEGOCIO (clsOpportunitiesMgr, clsFileManager, clsOpportunity, ...):Â  SerÃ¡ el Ãºltimo punto.
'  Una vez corregido todo lo que quepa corregir en la estructura de clases, Implementar lÃ³gica de negocio:
' -  implementar el fileManager, llevar a Ã©l los metodos para identificar ficheros (EsFicheroOportunidad, EsValidoGenerarGrafico ,....): solo funciones que DETERMINEN "EL TIPO DE FICHERO"; CONDICIONADO A PATRONES EN EL NOMBRE DEL FICHERO, o en el tipo de fichero en Excel, o en atributos de fichero, O EN LAS HOJAS DE EXCEL QUE TENGA, Y SU CONTENIDO... pero NUNCA EN EL ESTAADO PARTICULAR DEL ENTORNO DE EJECUCION DE EXCEL (es decir, no en funcion de si hay algo seleccionado, o si una ventana estÃ¡ minimizada, etc (NO son funciones "de identificacion de estado del ribbon", sino de identificacion del "estado de datos del fichero")
'    El ribbon los usara como callbacks para determinar si activar comandos.
'    Y podrÃ­an ser utiles tambiÃ¨n para otros usos,
' -  e implementar clsOpportunity...
' -  Implementar el procesador de las especificaciones de cliente (Procesado de PDFS):
'     detecciÃ³n de curvas de pdfplumber: aunque es suficientemente bueno
'     detectando tables, lines, images y rects, no me sirve con las curvas. he
'     estado revisando los logs que te he facilitado, y buscando patrones que
'     correspondan al PDF original, y NO VEO una correlaciÃ³n correcta. He
'     conseguido detectar bien casi todas las formas, pero NO el color de
'     relleno. Por ello me planteo combinar el trabajo de pdfplumber, con el
'     de otra libreria Python que detecte adecuadamente esas "curvas"... pero
'     Â¿tal vez como imÃ¡genes raster? Â¿usando pyMUPDF y EasyOCR, e incluso
'     otras librerias que sirven de base para pdfplumber (pensando en
'     pdfminer), u otras librerias de python para manipular y extraer
'     informacion de PDFs, podrÃ­amos obtener mejor informaciÃ³n sobre esas
'     entidades? (te adelanto que he comparado las 3 librerÃ­as pdfplumber
'     pyMUPDF y EasyOCR, y para las entidades referidas, tables, lines, images
'     y rects, el mejor trabajo lo hace pdfplumber; pero tal vez ayude con el
'     resto de geometrÃ­a) Â¿PodrÃ­as aÃ±adir al script siguiente, codigo para
'     detectar entidades mediante EasyOCR y pyMUPDF ; comparar los resultados
'     con las 3 librerÃ­as, unificar los resultados que coincidan, y presentar
'     "para depuracion visual", sobre el HTML, los resultados diferenciados?:
'     el objetivo es detectar con mas perfeccion las formas, y sus posiciones;
'     y sobre todo, si tienen un relleno NEGRO, o si es BLANCO.

'     tambien quiero que me cuentes algo mas sobre los motores de pdfplumber,
'     sean los "por defecto" (layout.py + pdfminer.six), u otros.

'TODO Crear sistema de logging persistente en archivo (baja prioridad)
'TODO Implementar sistema de versionado automÃ¡tico del complemento (baja prioridad)
'TODO AÃ±adir telemetrÃ­a opcional (con consentimiento usuario) (baja prioridad)
'TODO Crear manual de usuario con capturas (nula prioridad)
'TODO: EncriptaciÃ³n RC4 del script, e incluso del codigo VBA (baja prioridad)

'3. Code Metrics de Rubberduck
'A. MÃ©tricas a revisar regularmente
'Complejidad CiclomÃ¡tica:
'
'Objetivo: < 10 por funciÃ³n
'CrÃ­tico: > 20 (refactorizar urgente)
'Funciones complejas actuales a revisar:
'
'StartFolderWatcherSchedule: Simplificar con funciones helper
'getFileNameTag: Dividir en funciones por tipo de tag
'
'LÃ­neas de cÃ³digo:
'
'Objetivo: < 50 lÃ­neas por funciÃ³n
'Funciones largas a dividir:
'
'DiagnosticoCompleto: Dividir en submÃ©todos
'Script VBS embedido: Considerar externalizar
'4. Code Inspections de Rubberduck
'A. Inspecciones importantes a resolver
'Priority: High
'
'Variable Not used
'Procedure Not used
'Parameter can be ByVal
'Implicit Public Member
'Option Explicit not specified
'
'Priority: Medium
'
'Variable type not declared
'Function return value not used
'Empty If/ElseIf block
'
'b.Refactorings Ãºtiles
'
'Extract Method: para StartFolderWatcherSchedule
'Rename: Variables con nombres poco claros (i, j, k)
'Remove Parameter: ParÃ¡metros no utilizados
'Reorder Parameters: Agrupar parÃ¡metros relacionados
'
'5. Mejores prÃ¡cticas con Rubberduck
'a.Workflow recomendado
'
'Antes de codificar: Escribir test que falle
'Codificar: Implementar funcionalidad mÃ­nima
'Refactorizar: Usar Code Inspections
'Documentar: AÃ±adir TODOs para futuras mejoras
'Validar: Ejecutar todos los tests

'@IgnoreModule MissingAnnotationArgument
'@Folder "2-Infraestructura.Configuracion"

Option Explicit

Private Const MODULE_NAME As String = "mod_ConstantsGlobals"

' constantes de compilaciÃ³n
#Const RubberduckTest = 1 'True
#Const DebugMode = 1 'True

' Constantes para organizar la configuraciÃ³n
Public Const APP_NAME As String = "ABC_ofertas maquina especial"
Public Const FOLDERWATCHERCOM_NAME As String = "FolderWatcherCOM.dll"

' Nombres de las configuraciones
Public Const CFG_BASEFOLDER As String = "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\"

Public Const CFG_SAM As Integer = 41
Public Const CFG_PATH_SAM As String = CFG_BASEFOLDER & APP_NAME & "\SAM"

' Configuracion de rutas de carpetas
Public Const CFG_SECTION_RUTAS As String = "Folders"
Public Const CFG_PATH_SECTION_RUTAS As String = CFG_BASEFOLDER & APP_NAME & "\" & CFG_SECTION_RUTAS & "\"

Public Const CFG_RUTA_OPORTUNIDADES As String = "BaseFolderOportunidades"
Public Const CFG_RUTA_OPORTUNIDADES_DEFAULT As String = "C:\abc compressors\INTRANET\OilGas\3_OFERTAS\OFERTAS\2026\41 - SER\"
Public Const CFG_RUTA_PLANTILLAS As String = "BaseFolderPlantillas"
Public Const CFG_RUTA_PLANTILLAS_DEFAULT As String = "C:\abc compressors\INTRANET\OilGas\3_OFERTAS\OFERTAS\2026\41 - SER\_plantilla\"
Public Const CFG_RUTA_OFERGAS As String = "BaseFolderOfergas"
Public Const CFG_RUTA_OFERGAS_DEFAULT As String = "C:\Program Files (x86)\Ofertas_Gas\Excel\"
Public Const CFG_RUTA_GAS_VBNET As String = "BaseFolderGasVBNet"
Public Const CFG_RUTA_GAS_VBNET_DEFAULT As String = "C:\Aire\"
Public Const CFG_RUTA_PLCALCS As String = "BaseFolderXLSCalculos"
Public Const CFG_RUTA_PLCALCS_DEFAULT As String = "C:\abc compressors\2-0-OPORTUNIDADES\_Mis notas\"
Public Const CFG_RUTA_PLCALCNACE As String = "C:\abc compressors\INTRANET\OficinaTecnica\Documentacion\Normas\NACE\Herramienta\Herramienta_para_seleccion_de_materiales_v1.4.xlsx"
Public Const CFG_RUTA_PLSELCILS As String = "C:\abc compressors\INTRANET\OilGas\3_OFERTAS\ADJUNTOS OFERTAS\Datos cilindros 2.xlsx"
Public Const CFG_RUTA_COMPRIMGS As String = "BaseFolderImgsCompresores"
Public Const CFG_RUTA_COMPRIMGS_DEFAULT As String = "C:\abc compressors\INTRANET\OilGas\1_COMUNICACION\0- MARKETING\2- FOTOS\Fotos y planos\FOTOS\"
Public Const CFG_RUTA_COMPRDRAWPID As String = "BaseFolderPlanosPIDs"
Public Const CFG_RUTA_COMPRDRAWPID_DEFAULT As String = "C:\abc compressors\INTRANET\OilGas\5_DOCUMENTACION TECNICA\ADJUNTOS OFERTAS\3-1-PLANOS\|C:\abc compressors\INTRANET\OilGas\1_COMUNICACION\0- MARKETING\2- FOTOS\Fotos y planos\PLANOS"

' Polling de cambios en carpetas
Public Const CFG_FW_HEARTBEAT As String = "Heartbeat"
Public Const POLLING_SECONDS As Integer = 8
Public Const INACTIVITY_MINUTES As Integer = 10
Public Const WARMUP_MAX_CHECKS As Long = 5       ' Checks antes de verificar heartbeat (5 == 40 segundos aprox)

' Configuracion de registro de UDFs
Public Const CFG_RUTA_UDFS As String = CFG_BASEFOLDER & APP_NAME & "\UDFsRegistradas"
Public Const DEFAULT_CATEGORY As String = "Funciones Personalizadas"
Public Const DEFAULT_NOPARAMS As String = "(sin parÃ¡metros)"
Public Const DEFAULT_NORETURNS As String = "(ninguno)"
Public Const DEFAULT_NOARGS As String = "(sin argumentos)"

' Patrones para parsing de nombres de archivo
Public Const QUOTENR_PATTERN As String = "\d{9}(?:[\-_]\d+)?"
Public Const QUOTENR_REV_PATTERN As String = "(" & QUOTENR_PATTERN & ")(?:[ \-_]*rev\.?[ \-_]*(\d+)\b)?"
Public Const CUSTOMER_PATTERN As String = "(?:.(?! \- ))+." '"((?:.(?! \- ))+?.(?:\s*[\-_]\s*(?:.(?! \- ))+.)*?)"
Public Const PROJECT_OTHERS_PATTERN As String = "(?:.(?! \- ))+." '"((?:.(?! \- ))+?.(?:\s*[\-_]\s*(?:.(?! \- ))+.)*?)"
Public Const MODEL_PATTERN As String = "(\d)\s?T?\s*E\s?(H[AGPX])\s?\-\s?(\d)\s?\-\s?[LGT]{2,3}"
Public Const FULLMODEL_PATTERN As String = MODEL_PATTERN & "(?:\-\d\x\d+T?)+(?: (?:NACE|ATEX))*"
' en el caso de la descripcion de la oportunidad, se acepta poner XXXXX como modelo, si no estÃ¡ aÃºn definido
Public Const OPPORTUNITY_MODEL_PATTERN As String = "((?:(?:" & MODEL_PATTERN & ")[ ,y]*)+|X{3,})"
Public Const FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_PATTERN As String = "^(" & QUOTENR_PATTERN & _
")\s*\-(?:[#-]-)?\s*(" & CUSTOMER_PATTERN & ")(?:\s*\-\s*(" & PROJECT_OTHERS_PATTERN & "))??"
Public Const FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN As String = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_PATTERN & _
"\s*\-\s*" & OPPORTUNITY_MODEL_PATTERN

'--------------------------------------------------------------
'@Scope: Tipos definidos por el usuario en el proyecto VBA (nivel de mÃ³dulo o global)
'@ArgumentDescriptions: (ninguno) - estructuras estÃ¡ticas, no reciben argumentos
'--------------------------------------------------------------
Public Enum ProcType
    Macro
    udf
    internalPrivate
    internalSubPublicWithParams
    eventHandler
End Enum

Public Enum ProcKind
    proc
    PropLet
    PropSet
    PropGet
    ProcSub
    ProcFunction
End Enum

Public Enum ProcContainerType
    StdModule = 1
    ClassModule = 2
    Form = 3
    ActiveXDesigner = 11
    Sheet = 100
End Enum

'@Category: UI / Ribbon
Public Enum eRibbonMode
    Ribbon_Undefined = 0                         ' no inicializado
    Ribbon_OpportunityOnly = 1                   ' solo visible si el libro es de oportunidad
    Ribbon_User = 2                              ' modo usuario (sin grupos admin/dev)
    Ribbon_Hidden = 3                            ' no se muestra
    Ribbon_Admin = 4                             ' todo visible permanentemente
End Enum

'--------------------------------------------------------------
' @Description: Estado del Ribbon especÃ­fico por libro
' @Note: Se vincula a cada clsFileXLS para gestionar el ribbon por libro
'--------------------------------------------------------------
Public Type T_RibbonBookState
    Modo As eRibbonMode                          ' Modo del ribbon para este libro
    Visible As Boolean                           ' Visibilidad del ribbon para este libro
    ' Extensible: aÃ±adir mÃ¡s atributos por libro segÃºn necesidad futura
End Type

'--------------------------------------------------------------
' @Description: Tipos de archivo soportados
'--------------------------------------------------------------
Public Enum TipoArchivo
    UnDef = 0
    Unknown = 1
    oportunidad = 2                              ' Archivos de oportunidades
    CGASING_CurvasRendimiento = 3                ' Performance curves
    CGASING_Calcs = 4                            ' CÃ¡lculos C-GAS-ING
    PlantillaBudget = 5                          ' Budget
    PlantillaOferta = 6                          ' Quotation
End Enum

'--------------------------------------------------------------
' @Description: InformaciÃ³n generica de los archivos de Excel, relacionada con mi aplicaciÃ³n
'--------------------------------------------------------------
Public Type T_InfoArchivo
    EsValido As Boolean
    TipoDetectado As TipoArchivo
    Customer As String
    OpportunityNr As String
End Type

'--------------------------------------------------------------
'@Description: Estructura de datos que encapsula toda la informaciÃ³n relevante de un bloque de cÃ³digo VBA
'@Returns: N/A - se utiliza como tipo compuesto de datos
'@Category: Parsing de Procedimientos y AnÃ¡lisis de CÃ³digo
'--------------------------------------------------------------
Public Type T_CodeBlock
    strCode As String
    procStartLine As Long
    procSignatureLine As Long
    procNumLines As Long
    'procWrongEndLines As Long
End Type

#If Win64 Then
    ' CÃ³digo para Excel 64-bit
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    ' CÃ³digo para Excel 32-bit
    Public Declare Function GetTickCount Lib "kernel32" () As Long
#End If
#If VBA7 Then
    ' Office 2010+
    ' Usar PtrSafe en declares
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    ' Office 2007-
    ' Declares antiguos sin PtrSafe
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


