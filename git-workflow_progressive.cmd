@echo off
REM ============================================================
REM Script: git-workflow-v2.cmd
REM Version: 2.0
REM ============================================================
REM Objetivo:
REM   Usar Git como:
REM     - transporte de versiones
REM     - certificador en GitHub
REM   Usar Beyond Compare como:
REM     - herramienta de decision e integracion
REM
REM Funcionalidades v2:
REM   - Auditoria de codificacion (UTF-8, BOM, ANSI)
REM   - Configuracion automatica de .gitattributes
REM   - Conversion automatica UTF-8 <-> ANSI
REM   - Menu reorganizado por categorias
REM   - Doble confirmacion en acciones peligrosas
REM
REM NOTA:
REM   Este script NO hace merges automaticos.
REM   El criterio humano manda.
REM ============================================================

SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

REM ------------------------------------------------------------
REM FIJAR CODIFICACIÓN PARA CARACTERES ESPAÑOLES
REM ------------------------------------------------------------
REM Opción A: Windows-1252 (recomendada para Windows Español)
chcp 1252 >nul

REM ------------------------------------------------------------
REM CONFIGURACION GLOBAL
REM ------------------------------------------------------------
SET SCRIPT_VERSION=2.0
SET MAIN_DIR=main-mirror
SET CLAUDE_DIR=claude-mirror-progressive
SET CLAUDE_BRANCH=claude/review-main-progressive-Cyi0r
SET ORPHAN_DIR=claude-mirror__GITHUB_TMP
SET ORPHAN_BRANCH=qwen/claude-code-review
SET GITHUB_REPO=https://github.com/gautxori-yuyu/claude-vba-excel-abc.git
SET BC=%PORTABLE_APPS%Beyond Compare 4\BCompare.exe

REM Extensiones VBA a auditar/convertir
rem SET "TARGET_EXTENSIONS=cls bas frm vbs bat cmd"
SET "TARGET_EXTENSIONS=cls bas frm"

REM Guardar directorio inicial
SET START_DIR=%CD%

REM ------------------------------------------------------------
REM NORMALIZACION DE RUTA
REM ------------------------------------------------------------
CALL :NORMALIZE_PATH
IF ERRORLEVEL 1 GOTO :END_SCRIPT

REM ------------------------------------------------------------
REM MENU PRINCIPAL
REM ------------------------------------------------------------
:MENU_PRINCIPAL
SET "OPT="
cls
echo ============================================================
echo  GIT WORKFLOW v%SCRIPT_VERSION%
echo ============================================================
echo  Carpeta BASE: %BASEDIR%
echo  Rama de trabajo: %CLAUDE_BRANCH%
echo ============================================================
echo.
echo  --- FLUJO DIARIO ---
echo  Lo que usaras habitualmente:
echo.
echo   1 - Ver estado y recomendaciones
echo       Analiza MAIN y la rama de trabajo de IA.
echo.
echo   2 - Descargar novedades de la rama de trabajo ^(IA^)
echo       Trae los cambios que hayan subido a GitHub.
echo.
echo   3 - Publicar mis cambios en GitHub
echo       Sube tus cambios para que otros ^(IA^) los vean.
echo.
echo   4 - Comparar carpetas con Beyond Compare
echo       Abre MAIN y la rama de trabajo lado a lado.
echo.
echo  --- RESOLVER PROBLEMAS ---
echo.
echo   5 - Subir mi version DE RAMA MAIN, para revision por IA
echo       ^(Normalmente, cuando rama Main local y remota son diferentes,
echo		como alternativa a usar Beyond Compare, que la IA evalúe los cambios^).
echo   6 - Re-sincronizar RAMA CLAUDE ^(== rama de trabajo; descartar cambios locales^)
echo.
echo  --- SUBMENUS ---
echo.
echo   H - Herramientas y configuracion
echo   E - Codificacion UTF-8 / ANSI
echo.
echo   0 - Salir
echo.
set /p OPT=Opcion: 

IF "%OPT%"=="1" GOTO :STATUS
IF "%OPT%"=="2" GOTO :PULL_CLAUDE
IF "%OPT%"=="3" GOTO :PUBLISH_MAIN
IF "%OPT%"=="4" GOTO :DIFF
IF "%OPT%"=="5" GOTO :UPLOAD_FOR_REVIEW
IF "%OPT%"=="6" GOTO :RESYNC_CLAUDE
IF /I "%OPT%"=="H" GOTO :MENU_HERRAMIENTAS
IF /I "%OPT%"=="E" GOTO :MENU_CODIFICACION
IF "%OPT%"=="0" GOTO :END_SCRIPT

echo.
echo Opcion invalida. Pulsa una tecla...
pause >nul
GOTO :MENU_PRINCIPAL

REM ------------------------------------------------------------
REM SUBMENU: HERRAMIENTAS Y CONFIGURACION
REM ------------------------------------------------------------
:MENU_HERRAMIENTAS
SET "HOPT="
cls
echo ============================================================
echo  HERRAMIENTAS Y CONFIGURACION
echo ============================================================
echo.
echo  --- INFORMACION ---
echo   1 - Diagnostico rapido
echo   2 - Ver ramas y seguimiento remoto
echo   3 - Ver historial de cambios ^(30 ultimos^)
echo   4 - Comprobar credenciales de Git
echo   5 - Fetch ^(traer info de GitHub sin modificar nada^)
echo.
echo  --- COMPARACION ---
echo   6 - Comparar carpetas con Beyond Compare
echo   7 - Configurar Beyond Compare como difftool
echo.
echo  --- GESTION DE RAMAS ---
echo   8 - Cambiar rama "de trabajo" de la IA ^(actualmente: %CLAUDE_BRANCH%^)
echo   9 - Gestionar ramas huerfanas
echo.
echo  --- ETIQUETAS ---
echo  10 - Crear etiqueta de version
echo  11 - Ver etiquetas existentes
echo.
echo  --- GESTIÓN DE FICHEROS ---
echo  12 - Gestionar ficheros nuevos sin seguimiento
echo  13 - Proteger fichero ^(ignorar cambios locales^)
echo  14 - Desproteger fichero
echo.
echo  --- LIMPIEZA ---
echo  15 - Vista previa de limpieza
echo  16 - Ejecutar limpieza ^(BORRA no versionados^)
echo.
echo  --- OPERACIONES PELIGROSAS ---
echo  17 - Clonar repositorios desde cero ^(BORRA existentes^)
echo  18 - Push forzado a MAIN ^(SOBREESCRIBE GitHub^)
echo.
echo   0 - Volver al menu principal
echo.
set /p HOPT=Opcion: 

IF "%HOPT%"=="1" GOTO :DIAG
IF "%HOPT%"=="2" GOTO :TRACKING
IF "%HOPT%"=="3" GOTO :LOG
IF "%HOPT%"=="4" GOTO :CREDS
IF "%HOPT%"=="5" GOTO :FETCH
IF "%HOPT%"=="6" GOTO :DIFF
IF "%HOPT%"=="7" GOTO :CFG_BC
IF "%HOPT%"=="8" GOTO :CAMBIAR_RAMA_TRABAJO
IF "%HOPT%"=="9" GOTO :MENU_RAMAS_HUERFANAS
IF "%HOPT%"=="10" GOTO :TAG_CREATE
IF "%HOPT%"=="11" GOTO :TAG_LIST
IF "%HOPT%"=="12" GOTO :MANAGE_UNTRACKED
IF "%HOPT%"=="13" GOTO :PROTECT
IF "%HOPT%"=="14" GOTO :UNPROTECT
IF "%HOPT%"=="15" GOTO :CLEAN_PREVIEW
IF "%HOPT%"=="16" GOTO :CLEAN_EXEC
IF "%HOPT%"=="17" GOTO :CLONE_REPOS
IF "%HOPT%"=="18" GOTO :FORCE_PUSH_MAIN
IF "%HOPT%"=="0" GOTO :MENU_PRINCIPAL

echo.
echo Opcion invalida. Pulsa una tecla...
pause >nul
GOTO :MENU_HERRAMIENTAS

REM ------------------------------------------------------------
REM SUBMENU: CODIFICACION UTF-8 / ANSI
REM ------------------------------------------------------------
:MENU_CODIFICACION
SET "COPT="
cls
echo ============================================================
echo  CODIFICACION UTF-8 / ANSI
echo ============================================================
echo.
echo  El editor de VBA clasico solo entiende ANSI ^(Windows-1252^).
echo  Si un fichero esta en UTF-8, los acentos se veran mal.
echo.
echo  --- DIAGNOSTICO ---
echo   1 - Auditoria de codificacion
echo       Revisa ficheros y dice cuales pueden dar problemas.
echo.
echo  --- CONVERSION MANUAL ---
echo   2 - Convertir UN fichero de UTF-8 a ANSI
echo   3 - Convertir UN fichero de ANSI a UTF-8
echo   4 - Convertir TODOS los ficheros VBA a ANSI
echo.
echo  --- CONFIGURACION AUTOMATICA ---
echo   5 - Configurar .gitattributes
echo       Git convertira automaticamente entre UTF-8 y ANSI.
echo   6 - Ver .gitattributes actual
echo   7 - Desactivar configuraciones de conversión git
echo       Deshabilitar las conversiones desde el fichero config de Git
echo.
echo   0 - Volver al menu principal
echo.
set /p COPT=Opcion: 

IF "%COPT%"=="1" GOTO :AUDIT_ENCODING
IF "%COPT%"=="2" GOTO :CONVERT_UTF8_TO_ANSI
IF "%COPT%"=="3" GOTO :CONVERT_ANSI_TO_UTF8
IF "%COPT%"=="4" GOTO :CONVERT_ALL_TO_ANSI
IF "%COPT%"=="5" GOTO :CFG_GITATTRIBUTES
IF "%COPT%"=="6" GOTO :VIEW_GITATTRIBUTES
IF "%COPT%"=="6" GOTO :DISABLE_GIT_CONV
IF "%COPT%"=="0" GOTO :MENU_PRINCIPAL

echo.
echo Opcion invalida. Pulsa una tecla...
pause >nul
GOTO :MENU_CODIFICACION

REM ============================================================
REM SECCION: INFORMACION Y DIAGNOSTICO
REM ============================================================

:DIAG
cls
echo ============================================================
echo  DIAGNOSTICO RAPIDO
echo ============================================================
echo.
echo ---- Directorio actual ----
echo %CD%
echo.
echo ---- MAIN: %BASEDIR%\%MAIN_DIR% ----
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No existe el directorio MAIN
) ELSE (
	git rev-parse --show-toplevel 2>nul
	git remote -v
	git branch -a
)
echo.
echo ---- CLAUDE: %BASEDIR%\%CLAUDE_DIR% ----
cd /d "%BASEDIR%\%CLAUDE_DIR%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No existe el directorio CLAUDE
) ELSE (
	git rev-parse --show-toplevel 2>nul
	git remote -v
	git branch -a
)
echo.
pause
GOTO :RETURN_MENU

:STATUS
cls
echo ============================================================
echo  ESTADO DE ESPEJOS ^(con analisis^)
echo ============================================================

REM Variables para almacenar estado
SET "MAIN_HAS_LOCAL_CHANGES=0"
SET "MAIN_BEHIND_REMOTE=0"
SET "MAIN_AHEAD_REMOTE=0"
SET "MAIN_DIVERGED=0"
SET "MAIN_UNTRACKED=0"
SET "MAIN_MODIFIED=0"
SET "MAIN_AHEAD_COUNT=0"
SET "MAIN_BEHIND_COUNT=0"
SET "CLAUDE_HAS_LOCAL_CHANGES=0"
SET "CLAUDE_BEHIND_REMOTE=0"
SET "CLAUDE_AHEAD_REMOTE=0"
SET "CLAUDE_DIVERGED=0"
SET "CLAUDE_UNTRACKED=0"
SET "CLAUDE_MODIFIED=0"
SET "CLAUDE_AHEAD_COUNT=0"
SET "CLAUDE_BEHIND_COUNT=0"

echo.
echo ============================================================
echo  MAIN ^(%MAIN_DIR%^)
echo ============================================================
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] Directorio no existe
	GOTO :STATUS_CLAUDE
)

REM Actualizar referencias remotas silenciosamente
echo Consultando GitHub...
git fetch origin >nul 2>&1

REM Contar ficheros modificados y sin seguimiento
SET "MAIN_MODIFIED=0"
SET "MAIN_UNTRACKED=0"
FOR /F %%A IN ('git status --porcelain 2^>nul ^| findstr /B "^ M" ^| find /c /v ""') DO SET "MAIN_MODIFIED=%%A"
FOR /F %%A IN ('git status --porcelain 2^>nul ^| findstr /B "??" ^| find /c /v ""') DO SET "MAIN_UNTRACKED=%%A"
FOR /F %%A IN ('git status --porcelain 2^>nul ^| find /c /v ""') DO SET "LOCAL_CHANGES=%%A"
IF %LOCAL_CHANGES% GTR 0 SET "MAIN_HAS_LOCAL_CHANGES=1"

REM Analizar relacion con remoto
SET "MAIN_AHEAD_COUNT=0"
SET "MAIN_BEHIND_COUNT=0"
FOR /F "tokens=1,2" %%A IN ('git rev-list --count --left-right HEAD...origin/main 2^>nul') DO (
	SET "MAIN_AHEAD_COUNT=%%A"
	SET "MAIN_BEHIND_COUNT=%%B"
)
IF %MAIN_AHEAD_COUNT% GTR 0 IF %MAIN_BEHIND_COUNT% GTR 0 SET "MAIN_DIVERGED=1"
IF %MAIN_AHEAD_COUNT% GTR 0 IF %MAIN_BEHIND_COUNT% EQU 0 SET "MAIN_AHEAD_REMOTE=1"
IF %MAIN_AHEAD_COUNT% EQU 0 IF %MAIN_BEHIND_COUNT% GTR 0 SET "MAIN_BEHIND_REMOTE=1"

echo.
echo --- ESTADO RESUMIDO ---
IF "%MAIN_DIVERGED%"=="1" (
	echo ^[!!^] DIVERGENCIA: Tu rama MAIN tiene %MAIN_AHEAD_COUNT% commit^(s^) que GitHub no tiene
	echo                   Y GitHub tiene %MAIN_BEHIND_COUNT% commit^(s^) que tu no tienes
	echo                   ^(Hay conflicto - necesita revision manual^)
) ELSE IF "%MAIN_AHEAD_REMOTE%"=="1" (
	echo ^[i^] Tienes %MAIN_AHEAD_COUNT% commit^(s^) listos para SUBIR a GitHub
	echo     ^(Tu rama MAIN esta MAS AVANZADA que GitHub^)
) ELSE IF "%MAIN_BEHIND_REMOTE%"=="1" (
	echo ^[i^] GitHub tiene %MAIN_BEHIND_COUNT% commit^(s^) que tu NO tienes
	echo     ^(Tu rama MAIN esta ATRASADA respecto a GitHub^)
) ELSE (
	echo ^[OK^] Sincronizado con GitHub
)

IF %MAIN_MODIFIED% GTR 0 (
	echo.
	echo ^[!^] Tienes %MAIN_MODIFIED% fichero^(s^) MODIFICADO^(s^) sin guardar
	echo     ^(Cambios que has hecho pero NO has "guardado" con commit^)
)
IF %MAIN_UNTRACKED% GTR 0 (
	echo.
	echo ^[!^] Tienes %MAIN_UNTRACKED% fichero^(s^) NUEVO^(s^) sin seguimiento
	echo     ^(Ficheros que Git NO esta vigilando todavia^)
)

:STATUS_CLAUDE
echo.
echo ============================================================
echo  CLAUDE ^(%CLAUDE_DIR%^)
echo ============================================================
cd /d "%BASEDIR%\%CLAUDE_DIR%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] Directorio no existe
	GOTO :STATUS_ANALYSIS
)

REM Actualizar referencias remotas silenciosamente
echo Consultando GitHub...
git fetch origin >nul 2>&1

REM Contar ficheros modificados y sin seguimiento
SET "CLAUDE_MODIFIED=0"
SET "CLAUDE_UNTRACKED=0"
FOR /F %%A IN ('git status --porcelain 2^>nul ^| findstr /B "^ M" ^| find /c /v ""') DO SET "CLAUDE_MODIFIED=%%A"
FOR /F %%A IN ('git status --porcelain 2^>nul ^| findstr /B "??" ^| find /c /v ""') DO SET "CLAUDE_UNTRACKED=%%A"
FOR /F %%A IN ('git status --porcelain 2^>nul ^| find /c /v ""') DO SET "LOCAL_CHANGES=%%A"
IF %LOCAL_CHANGES% GTR 0 SET "CLAUDE_HAS_LOCAL_CHANGES=1"

REM Analizar relacion con remoto (rama Claude)
SET "CLAUDE_AHEAD_COUNT=0"
SET "CLAUDE_BEHIND_COUNT=0"
FOR /F "tokens=1,2" %%A IN ('git rev-list --count --left-right HEAD...origin/%CLAUDE_BRANCH% 2^>nul') DO (
	SET "CLAUDE_AHEAD_COUNT=%%A"
	SET "CLAUDE_BEHIND_COUNT=%%B"
)
IF %CLAUDE_AHEAD_COUNT% GTR 0 IF %CLAUDE_BEHIND_COUNT% GTR 0 SET "CLAUDE_DIVERGED=1"
IF %CLAUDE_AHEAD_COUNT% GTR 0 IF %CLAUDE_BEHIND_COUNT% EQU 0 SET "CLAUDE_AHEAD_REMOTE=1"
IF %CLAUDE_AHEAD_COUNT% EQU 0 IF %CLAUDE_BEHIND_COUNT% GTR 0 SET "CLAUDE_BEHIND_REMOTE=1"

REM Detectar diferencias reales contra remoto (tras filtros y conversiones)
SET "CLAUDE_EFFECTIVE_DIFF=0"
FOR /F %%A IN ('git diff --stat origin/%CLAUDE_BRANCH% -- ^| find /c /v ""') DO SET "CLAUDE_EFFECTIVE_DIFF=%%A"


REM Si hay cambios locales pero no diferencias reales, son intrascendentes
IF %LOCAL_CHANGES% GTR 0 IF %CLAUDE_EFFECTIVE_DIFF% EQU 0 (
    SET "CLAUDE_ONLY_ENCODING_CHANGES=1"
)

echo.
echo --- ESTADO RESUMIDO ---

GOTO :STATUS_ANALYSIS
REM Caso 1: Todo sincronizado REALMENTE
IF %CLAUDE_EFFECTIVE_DIFF% EQU 0 IF %CLAUDE_BEHIND_COUNT% EQU 0 IF %CLAUDE_AHEAD_COUNT% EQU 0 (
    IF %CLAUDE_HAS_LOCAL_CHANGES% EQU 0 (
        echo ^[OK^] Sincronizado con GitHub
        SET "CLAUDE_SYNC_STATE=OK"
	) ELSE (
	    REM Ojo a esta instrucción que he insertado yo 
		SET "CLAUDE_HAS_LOCAL_CHANGES=0"
	)
)ELSE IF %CLAUDE_ONLY_ENCODING_CHANGES% EQU 1 (
    echo ^[OK^] Cambios locales solo de codificacion ^(sin impacto funcional^)
    echo      No es necesaria ninguna accion de sincronizacion
    SET "CLAUDE_SYNC_STATE=ENCODING_ONLY"
)

IF "%CLAUDE_DIVERGED%"=="1" (
	echo ^[!!^] DIVERGENCIA: Situacion inusual en CLAUDE
	echo                   Recomendacion: Re-sincronizar CLAUDE ^(opcion 7^)
) ELSE IF "%CLAUDE_BEHIND_REMOTE%"=="1" (
	echo ^[i^] Claude ha subido %CLAUDE_BEHIND_COUNT% cambio^(s^) nuevo^(s^) a GitHub
	echo     ^(Puedes DESCARGAR sus cambios para revisarlos^)
) ELSE IF "%CLAUDE_AHEAD_REMOTE%"=="1" (
	echo ^[!^] INUSUAL: Tu carpeta CLAUDE tiene cambios que GitHub no tiene
	echo     ^(No deberias modificar CLAUDE - es solo para leer^)
) ELSE (
	echo ^[OK^] Sincronizado con GitHub
)

IF %CLAUDE_MODIFIED% GTR 0 (
	echo.
	echo ^[!^] INUSUAL: Hay %CLAUDE_MODIFIED% fichero^(s^) modificado^(s^) en CLAUDE
	echo     ^(No deberias modificar CLAUDE - considera re-sincronizar^)
)
IF %CLAUDE_UNTRACKED% GTR 0 (
	echo.
	echo ^[!^] Hay %CLAUDE_UNTRACKED% fichero^(s^) sin seguimiento en CLAUDE
)

:STATUS_ANALYSIS
echo.
echo ============================================================
echo  RECOMENDACIONES
echo ============================================================
echo.

SET "HAS_RECOMMENDATIONS=0"

REM --- Recomendaciones para MAIN ---
IF "%MAIN_DIVERGED%"=="1" (
	SET "HAS_RECOMMENDATIONS=1"
	echo MAIN - CONFLICTO:
	echo   Tu rama MAIN y GitHub tienen cambios DIFERENTES.
	echo   Accion: Debes decidir cual version prevalece ^(opcion D^).
	echo.
)

IF "%MAIN_BEHIND_REMOTE%"=="1" (
	SET "HAS_RECOMMENDATIONS=1"
	echo MAIN - DESACTUALIZADO:
	echo   Alguien ^(quizas tu desde otro sitio^) ha subido cambios a GitHub.
	echo   Normalmente no deberia pasar si solo tu modificas MAIN.
	echo   Accion: Revisa que hay en GitHub antes de continuar ^(opcion D^).
	echo.
)

IF "%MAIN_AHEAD_REMOTE%"=="1" (
	SET "HAS_RECOMMENDATIONS=1"
	echo MAIN - PENDIENTE DE PUBLICAR:
	echo   Tienes %MAIN_AHEAD_COUNT% commit^(s^) guardados localmente en rama MAIN que NO estan en GitHub.
	echo   Accion: Publicar en GitHub ^(opcion P^) para que Claude pueda verlos.
	echo.
)

IF "%MAIN_HAS_LOCAL_CHANGES%"=="1" (
	SET "HAS_RECOMMENDATIONS=1"
	echo MAIN - CAMBIOS SIN GUARDAR:
	echo   Has modificado ficheros en rama MAIN pero NO has hecho commit.
	echo   Accion: Publicar en GitHub ^(opcion P^) - esto hara commit y push.
	echo.
)

REM --- Recomendaciones para CLAUDE ---
IF "%CLAUDE_DIVERGED%"=="1" (
	SET "HAS_RECOMMENDATIONS=1"
	echo CLAUDE - PROBLEMA:
	echo   La carpeta CLAUDE esta en un estado inconsistente.
	echo   Accion: Re-sincronizar CLAUDE ^(opcion R^)
	echo.
)

IF "%CLAUDE_BEHIND_REMOTE%"=="1" (
	SET "HAS_RECOMMENDATIONS=1"
	echo CLAUDE - HAY NOVEDADES:
	echo   Claude ha subido %CLAUDE_BEHIND_COUNT% cambio^(s^) nuevo^(s^) a GitHub.
	echo   Accion: Descargar novedades de Claude ^(opcion A^)
	echo.
)

rem Ojo al siguiente cambio realizado por mí 
IF %CLAUDE_EFFECTIVE_DIFF% EQU 0 (
	echo   No se han detectado diferencias reales entre la carpeta local y el repositorio remoto.
	echo   Los siguientes mensajes de advertencia podrían ser imprecisos.
	echo.
) ELSE (
	echo   Accion: Re-sincronizar CLAUDE ^(opcion R^)
	echo.
)

IF "%CLAUDE_AHEAD_REMOTE%"=="1" (
	SET "HAS_RECOMMENDATIONS=1"
	echo CLAUDE - SITUACION INUSUAL:
	echo   Tu carpeta CLAUDE tiene cambios que no deberian estar ahi.
	echo   Accion: Re-sincronizar CLAUDE ^(opcion R^)
	echo.
)
IF "%CLAUDE_HAS_LOCAL_CHANGES%"=="1" (
	SET "HAS_RECOMMENDATIONS=1"
	echo CLAUDE - MODIFICACIONES NO DESEADAS:
	echo   Hay ficheros modificados en CLAUDE. Esta carpeta deberia ser solo lectura.
	echo   Accion: Re-sincronizar CLAUDE ^(opcion R^)
	echo.
)
rem Ojo al siguiente cambio realizado por mí 
IF "%CLAUDE_SYNC_STATE%"=="OK" SET "%HAS_RECOMMENDATIONS%"=="0"
IF "%HAS_RECOMMENDATIONS%"=="0" (
	echo ^[OK^] Todo esta sincronizado correctamente.
	echo      No hay acciones pendientes.
)

echo.
echo ============================================================
echo  ACCIONES RAPIDAS
echo ============================================================
echo.
echo  P - Publicar MAIN en GitHub ^(subir tus cambios^)
echo  A - Actualizar CLAUDE ^(descargar cambios de Claude^)
echo  R - Re-sincronizar CLAUDE ^(descartar cambios locales^)
echo  D - Ver diferencias con Beyond Compare
echo  Enter - Volver al menu
echo.
SET "STATUS_ACTION="
set /p STATUS_ACTION=Accion: 

IF /I "%STATUS_ACTION%"=="P" GOTO :PUBLISH_MAIN
IF /I "%STATUS_ACTION%"=="A" GOTO :PULL_CLAUDE
IF /I "%STATUS_ACTION%"=="R" GOTO :RESYNC_CLAUDE
IF /I "%STATUS_ACTION%"=="D" GOTO :DIFF

GOTO :RETURN_MENU

:TRACKING
cls
echo ============================================================
echo  RAMAS Y SEGUIMIENTO REMOTO
echo ============================================================
echo.

cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No se puede acceder al directorio MAIN
	pause
	GOTO :RETURN_MENU
)

echo Actualizando informacion de ramas remotas...
git fetch --all >nul 2>&1

echo.
echo ---- RAMAS LOCALES ----
git branch -vv
echo.

echo ---- TODAS LAS RAMAS EN GITHUB ----
git branch -r
echo.

echo ---- CONFIGURACION ACTUAL DEL SCRIPT ----
echo   Rama principal ^(MAIN^): main
echo   Rama de trabajo: %CLAUDE_BRANCH%
echo   Carpeta de trabajo: %CLAUDE_DIR%
echo   Rama huerfana: %ORPHAN_BRANCH%
echo   Carpeta huerfana: %ORPHAN_DIR%
echo.
pause
GOTO :RETURN_MENU

:LOG
cls
echo ============================================================
echo  HISTORIAL ^(ultimos 30 commits^)
echo ============================================================
echo.
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul && (
	git log --oneline --decorate --graph --all --max-count=30
)
echo.
pause
GOTO :RETURN_MENU

:CREDS
cls
echo ============================================================
echo  CONFIGURACION DE USUARIO Y CREDENCIALES
echo ============================================================
echo.
echo ---- Identidad Git ^(global^) ----
echo user.name:  
git config --global --get user.name
echo user.email: 
git config --global --get user.email
echo.
echo ---- Identidad Git ^(local, si existe^) ----
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul && (
	echo user.name:  
	git config --local --get user.name 2>nul || echo ^[no configurado^]
	echo user.email: 
	git config --local --get user.email 2>nul || echo ^[no configurado^]
)
echo.
echo ---- Credential helper ----
git config --global --get credential.helper
echo.
echo ---- URL del remoto ^(define metodo de auth^) ----
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul && git remote -v
echo.
echo NOTA: 
echo   https://... = Git usara token o helper
echo   git@github.com:... = Git usara SSH
echo.
pause
GOTO :RETURN_MENU

REM ============================================================
REM SECCION: CODIFICACION Y GITATTRIBUTES
REM ============================================================

:AUDIT_ENCODING
SET "AUDIT_OPT="
cls
echo ============================================================
echo  AUDITORIA DE CODIFICACION
echo ============================================================
echo.
echo Esta operacion NO modifica ningun fichero.
echo Analiza: %TARGET_EXTENSIONS%
echo.

REM Crear script VBS temporal para detectar codificacion
CALL :CREATE_ENCODING_DETECTOR

SET AUDIT_DIR=
echo Directorios disponibles:
echo   1 - MAIN  ^(%BASEDIR%\%MAIN_DIR%^)
echo   2 - CLAUDE ^(%BASEDIR%\%CLAUDE_DIR%^)
echo   3 - Ambos
echo.
set /p AUDIT_OPT=Selecciona: 

IF "%AUDIT_OPT%"=="1" (
	SET AUDIT_DIR=%BASEDIR%\%MAIN_DIR%
	CALL :DO_AUDIT "%BASEDIR%\%MAIN_DIR%"
)
IF "%AUDIT_OPT%"=="2" (
	SET AUDIT_DIR=%BASEDIR%\%CLAUDE_DIR%
	CALL :DO_AUDIT "%BASEDIR%\%CLAUDE_DIR%"
)
IF "%AUDIT_OPT%"=="3" (
	echo.
	echo === AUDITANDO MAIN ===
	CALL :DO_AUDIT "%BASEDIR%\%MAIN_DIR%"
	echo.
	echo === AUDITANDO CLAUDE ===
	CALL :DO_AUDIT "%BASEDIR%\%CLAUDE_DIR%"
)

echo.
echo ============================================================
echo  FIN DE AUDITORIA
echo ============================================================
pause
GOTO :RETURN_MENU

:DO_AUDIT

REM Parametro: %1 = directorio a auditar
SET "AUDIT_DIR=%~1"

echo.
echo ============================================================
echo  AUDITORIA COMPLETA DE CODIFICACION Y CONFIGURACION GIT
echo ============================================================
echo Directorio: !AUDIT_DIR!
echo ============================================================

cd /d "!AUDIT_DIR!" 2>nul
IF ERRORLEVEL 1 (
    echo ^[ERROR^] No se puede acceder al directorio
    EXIT /B 1
)

REM --- PARTE 1: AUDITORIA DE ARCHIVOS ---
echo.
echo ^[1/3^] AUDITORIA DE ARCHIVOS ^(.bas, .cls, .frm^)
echo ---------------------------------------------

SET FILES_CHECKED=0
SET FILES_UTF8_BOM=0
SET FILES_UTF8_NOBOM=0
SET FILES_ANSI=0
SET FILES_ASCII=0
SET FILES_OTHER=0
SET "FIRST_ENCODING="
SET "ENCODING_MISMATCH=0"

FOR %%E IN (%TARGET_EXTENSIONS%) DO (
    FOR /F "delims=" %%F IN ('dir /b *.%%E 2^>nul') DO (
        SET /A FILES_CHECKED+=1
        
        CALL :CHECK_FILE_ENCODING "%%F"
        
        REM Mostrar resultado
        IF "!ENCODING_RESULT!"=="UTF8-BOM" (
            SET /A FILES_UTF8_BOM+=1
            echo ^[^!^] UTF-8 BOM    : %%F ^(problematico para VBA^)
        ) ELSE IF "!ENCODING_RESULT!"=="UTF8-NOBOM" (
            SET /A FILES_UTF8_NOBOM+=1
            echo ^[^!^] UTF-8 sin BOM: %%F ^(puede dar problemas^)
        ) ELSE IF "!ENCODING_RESULT!"=="ANSI" (
            SET /A FILES_ANSI+=1
            echo ^[OK^] ANSI        : %%F ^(compatible VBA^)
        ) ELSE IF "!ENCODING_RESULT!"=="ASCII" (
            SET /A FILES_ASCII+=1
            echo ^[OK^] ASCII       : %%F ^(compatible^)
        ) ELSE (
            SET /A FILES_OTHER+=1
            echo ^[?^] !ENCODING_RESULT! : %%F
        )
        
        REM Verificar homogeneidad
        IF NOT DEFINED FIRST_ENCODING (
            IF NOT "!ENCODING_RESULT!"=="ASCII" (
                SET "FIRST_ENCODING=!ENCODING_RESULT!"
            )
        ) ELSE (
            IF NOT "!ENCODING_RESULT!"=="ASCII" (
                IF NOT "!ENCODING_RESULT!"=="!FIRST_ENCODING!" (
                    SET "ENCODING_MISMATCH=1"
                )
            )
        )
    )
)

echo.
echo RESUMEN ARCHIVOS:
echo   Total analizados: %FILES_CHECKED%
echo   UTF-8 con BOM:    %FILES_UTF8_BOM% ^(PROBLEMATICO^)
echo   UTF-8 sin BOM:    %FILES_UTF8_NOBOM% ^(POSIBLE PROBLEMA^)
echo   ANSI ^(8859-1/1252^): %FILES_ANSI% ^(OK^)
echo   ASCII:            %FILES_ASCII% ^(OK^)
echo   Otros:            %FILES_OTHER%

IF "%ENCODING_MISMATCH%"=="1" (
    echo.
    echo ^[ATENCION^] Hay mezcla de codificaciones!
    echo            Esto causara errores con working-tree-encoding.
)

REM --- PARTE 2: AUDITORIA CONFIGURACION GIT ---
echo.
echo ^[2/3^] AUDITORIA CONFIGURACION GIT
echo ----------------------------------

echo 1. .gitattributes:
IF EXIST ".gitattributes" (
    echo   ^[EXISTE^] Contenido:
    type ".gitattributes" | findstr /I "working-tree-encoding encoding"
    IF ERRORLEVEL 1 echo   ^[INFO^] No hay reglas de codificacion en .gitattributes
) ELSE (
    echo   ^[NO EXISTE^]
)

echo.
echo 2. Configuracion Git local:
echo   - workingTreeEncoding: !git config --local core.workingTreeEncoding!
echo   - autocrlf: !git config --local core.autocrlf!
echo   - safecrlf: !git config --local core.safecrlf!

echo.
echo 3. Filtros activos:
git config --local --get-regexp "^filter\." 2>nul
IF ERRORLEVEL 1 echo   ^[NO HAY FILTROS^]

REM --- PARTE 3: VERIFICACION COMPATIBILIDAD ---
echo.
echo ^[3/3^] VERIFICACION DE COMPATIBILIDAD
echo -------------------------------------

REM Verificar si .gitattributes fuerza ISO-8859-1
SET "HAS_WORKING_TREE_ENCODING=0"
IF EXIST ".gitattributes" (
    findstr /I "working-tree-encoding.*8859-1" ".gitattributes" >nul 2>&1
    IF NOT ERRORLEVEL 1 SET "HAS_WORKING_TREE_ENCODING=1"
    
    findstr /I "working-tree-encoding.*1252" ".gitattributes" >nul 2>&1
    IF NOT ERRORLEVEL 1 SET "HAS_WORKING_TREE_ENCODING=1"
)

REM Verificar config Git
git config --local core.workingTreeEncoding | findstr /I "8859-1\|1252" >nul 2>&1
IF NOT ERRORLEVEL 1 SET "HAS_WORKING_TREE_ENCODING=1"

IF "%HAS_WORKING_TREE_ENCODING%"=="1" (
    echo ^[CONFIGURADO^] Git convertira automaticamente entre UTF-8 e ISO-8859-1
    
    REM Verificar compatibilidad con archivos actuales
    IF %FILES_UTF8_BOM% GTR 0 (
        echo ^[ERROR^] Hay %FILES_UTF8_BOM% archivo^(s^) con UTF-8 BOM
        echo         Git no puede convertir UTF-8 BOM a ISO-8859-1 correctamente
        echo         Elimina el BOM primero: opcion 2 en menu Codificacion
    )
    
    IF %FILES_ANSI% GTR 0 (
        echo ^[OK^] %FILES_ANSI% archivo^(s^) ya en ANSI - compatible
    )
) ELSE (
    echo ^[NO CONFIGURADO^] No hay conversion automatica Git
    echo                   Los archivos se guardaran como estan
)

echo.
echo ============================================================
echo  FIN DE AUDITORIA
echo ============================================================

EXIT /B 0

:CHECK_FILE_ENCODING
REM Parametro: %1 = ruta completa del fichero
REM Retorna: ENCODING_RESULT = "UTF8-BOM", "UTF8-NOBOM", "ANSI", "ASCII", o "UNKNOWN"
SET "FILEPATH=%~1"
SET "FILENAME=%~nx1"

REM Verificar que el fichero existe
IF NOT EXIST "%FILEPATH%" EXIT /B 0

REM Usar file -i para detectar codificacion
SET "ENCODING_RESULT=UNKNOWN"
SET "FILE_OUTPUT="

REM Obtener salida de file -i
FOR /F "usebackq delims=" %%R IN (`file -i "!FILEPATH!" 2^>nul`) DO SET "FILE_OUTPUT=%%R"

IF NOT "!FILE_OUTPUT!"=="" (
    REM Extraer charset de la salida (ej: "text/plain; charset=utf-8")
    SET "CHARSET="
    FOR /F "tokens=2 delims=;" %%C IN ("!FILE_OUTPUT!") DO SET "CHARSET=%%C"
    
    REM Limpiar y normalizar
    SET "CHARSET=!CHARSET:charset=!"
    SET "CHARSET=!CHARSET: =!"
    
    REM Determinar tipo de codificacion
    IF /I "!CHARSET!"=="utf-8" (
        REM Verificar si tiene BOM usando PowerShell
        SET "HAS_BOM=0"
        FOR /F %%B IN ('powershell -Command "& {if ((Get-Content -Path '!FILEPATH!' -Encoding Byte -TotalCount 3)[0] -eq 239 -and (Get-Content -Path '!FILEPATH!' -Encoding Byte -TotalCount 3)[1] -eq 187 -and (Get-Content -Path '!FILEPATH!' -Encoding Byte -TotalCount 3)[2] -eq 191) {Write-Output '1'} else {Write-Output '0'}}" 2^>nul') DO SET "HAS_BOM=%%B"
        IF !HAS_BOM!==1 (
            SET "ENCODING_RESULT=UTF8-BOM"
	        SET /A FILES_UTF8_BOM+=1
	        echo ^[UTF-8 BOM^]  ATENCION: %FILENAME% tiene BOM - problematico para VBA clasico
        ) ELSE (
            SET "ENCODING_RESULT=UTF8-NOBOM"
	        SET /A FILES_UTF8_NOBOM+=1
	        echo ^[UTF-8^]      %FILENAME% - UTF-8 sin BOM
        )
    ) ELSE IF /I "!CHARSET!"=="iso-8859-1" (
        SET "ENCODING_RESULT=ANSI"
	    SET /A FILES_ANSI+=1
	    echo ^[ISO-8859-1^] %FILENAME% - ANSI/Latin-1
    ) ELSE IF /I "!CHARSET!"=="windows-1252" (
        SET "ENCODING_RESULT=ANSI"
	    SET /A FILES_ANSI+=1
	    echo ^[Windows-1252^] %FILENAME% - ANSI Windows
    ) ELSE IF /I "!CHARSET!"=="us-ascii" (
        SET "ENCODING_RESULT=ASCII"
	    SET /A FILES_ANSI+=1
	    echo ^[ASCII^]      %FILENAME% - ASCII puro
    ) ELSE (
        SET "ENCODING_RESULT=!CHARSET!"
	    SET /A FILES_SUSPICIOUS+=1
	    echo ^[NO ESPERADO-SOSPECHOSO^] %FILENAME% - !ENCODING_RESULT!
    )
)

REM Si file -i falla, usar metodo alternativo con PowerShell
IF "!ENCODING_RESULT!"=="UNKNOWN" (
    FOR /F %%R IN ('powershell -Command "& {try {$bytes = [System.IO.File]::ReadAllBytes('!FILEPATH!'); if ($bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191) {Write-Output 'UTF8-BOM'} else {$content = [System.IO.File]::ReadAllText('!FILEPATH!', [System.Text.Encoding]::UTF8); if ($? -and -not [string]::IsNullOrEmpty($content)) {Write-Output 'UTF8-NOBOM'} else {Write-Output 'ANSI'}}} catch {Write-Output 'ANSI'}}" 2^>nul') DO SET "ENCODING_RESULT=%%R"
)

EXIT /B 0

:CFG_GITATTRIBUTES
SET "CFG_OPT="
cls
echo ============================================================
echo  CONFIGURAR .gitattributes PARA VBA
echo ============================================================
echo.
echo Esta configuracion hace que Git:
echo   - Almacene los ficheros en el repositorio como quiera ^(UTF-8^)
echo   - Los convierta automaticamente a ANSI ^(windows-1252^) al hacer
echo     checkout, pull o clone en tu directorio de trabajo local
echo   - Los convierta a UTF-8 al hacer commit/push
echo.
echo Esto evita "colisiones de cambios" por diferencias de codificacion.
echo.

SET TARGET_DIR=
echo Aplicar configuracion en:
echo   1 - MAIN  ^(%BASEDIR%\%MAIN_DIR%^)
echo   2 - CLAUDE ^(%BASEDIR%\%CLAUDE_DIR%^)
echo   3 - Ambos
echo.
set /p CFG_OPT=Selecciona: 

IF "%CFG_OPT%"=="1" CALL :WRITE_GITATTRIBUTES "%BASEDIR%\%MAIN_DIR%"
IF "%CFG_OPT%"=="2" CALL :WRITE_GITATTRIBUTES "%BASEDIR%\%CLAUDE_DIR%"
IF "%CFG_OPT%"=="3" (
	CALL :WRITE_GITATTRIBUTES "%BASEDIR%\%MAIN_DIR%"
	CALL :WRITE_GITATTRIBUTES "%BASEDIR%\%CLAUDE_DIR%"
)

echo.
pause
GOTO :RETURN_MENU

:WRITE_GITATTRIBUTES
REM Parametro: %1 = directorio donde crear .gitattributes
SET TARGET=%~1
echo.
echo Configurando: %TARGET%

cd /d "%TARGET%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No se puede acceder al directorio
	EXIT /B 1
)

REM Hacer backup si existe
IF EXIST ".gitattributes" (
	echo Creando backup: .gitattributes.bak
	copy /Y ".gitattributes" ".gitattributes.bak" >nul
)

REM Crear .gitattributes con configuracion VBA
(
echo # ============================================================
echo # .gitattributes - Configuracion para proyecto VBA
echo # ============================================================
echo # Generado automaticamente
echo # Fecha: %DATE% %TIME%
echo # ============================================================
echo.
echo # ------------------------------------------------------------
echo # FICHEROS VBA - Conversion automatica de codificacion
echo # ------------------------------------------------------------
echo # working-tree-encoding=windows-1252:
echo #   - En el repositorio: Git guarda como UTF-8
echo #   - En tu directorio local: Git convierte a ANSI/Windows-1252
echo #   - Al hacer commit: Git convierte de ANSI a UTF-8
echo # ------------------------------------------------------------
echo.
echo *.bas text working-tree-encoding=windows-1252 eol=crlf
echo *.cls text working-tree-encoding=windows-1252 eol=crlf
echo *.frm text working-tree-encoding=windows-1252 eol=crlf
echo *.frx binary
echo.
echo # ------------------------------------------------------------
echo # FICHEROS VBScript
echo # ------------------------------------------------------------
echo *.vbs text working-tree-encoding=windows-1252 eol=crlf
echo.
echo # ------------------------------------------------------------
echo # FICHEROS BATCH/CMD - Tambien ANSI para compatibilidad
echo # ------------------------------------------------------------
echo *.cmd text working-tree-encoding=windows-1252 eol=crlf
echo *.bat text working-tree-encoding=windows-1252 eol=crlf
echo.
echo # ------------------------------------------------------------
echo # FICHEROS DE TEXTO GENERALES
echo # ------------------------------------------------------------
echo *.txt text eol=crlf
echo *.md text eol=crlf
echo *.json text eol=lf
echo *.xml text eol=crlf
echo.
echo # ------------------------------------------------------------
echo # FICHEROS QUE NO DEBEN MODIFICARSE
echo # ------------------------------------------------------------
echo *.xlsx binary
echo *.xlsm binary
echo *.xls binary
echo *.accdb binary
echo *.mdb binary
echo *.png binary
echo *.jpg binary
echo *.gif binary
echo *.ico binary
echo *.zip binary
) > ".gitattributes"

echo ^[OK^] .gitattributes creado

REM Preguntar si aplicar la configuracion a ficheros existentes
echo.
echo Para que los cambios afecten a ficheros existentes,
echo es necesario "refrescar" el repositorio.
echo.
choice /M "Aplicar ahora a ficheros existentes"
IF ERRORLEVEL 2 (
	echo.
	echo Puedes aplicarlo manualmente mas tarde con:
	echo   git rm --cached -r .
	echo   git reset --hard
	EXIT /B 0
)

echo.
echo Aplicando configuracion...
git rm --cached -r . >nul 2>&1
git reset --hard >nul 2>&1
echo ^[OK^] Configuracion aplicada

EXIT /B 0

:VIEW_GITATTRIBUTES
cls
echo ============================================================
echo  VER .gitattributes ACTUAL
echo ============================================================
echo.
echo ---- MAIN ----
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul && (
	IF EXIST ".gitattributes" (
		type ".gitattributes"
	) ELSE (
		echo ^[No existe .gitattributes^]
	)
)
echo.
echo ---- CLAUDE ----
cd /d "%BASEDIR%\%CLAUDE_DIR%" 2>nul && (
	IF EXIST ".gitattributes" (
		type ".gitattributes"
	) ELSE (
		echo ^[No existe .gitattributes^]
	)
)
echo.
pause
GOTO :RETURN_MENU

:DISABLE_GIT_CONV
cls
echo ============================================================
echo  DESACTIVAR CONVERSIONES EN CONFIGURACION GIT
echo ============================================================
echo.
echo ---- MAIN ----
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul && (
	git config --local --unset core.autocrlf
	git config --local --unset core.safecrlf
	git config --local --unset core.checkRoundtripEncoding
	REM DESACTIVAR CONVERSION DE FINALES DE LINEA
	REM git config --local core.autocrlf false
	REM git config --local core.safecrlf false
)
echo.
echo ---- CLAUDE ----
cd /d "%BASEDIR%\%CLAUDE_DIR%" 2>nul && (
	git config --local --unset core.autocrlf
	git config --local --unset core.safecrlf
	git config --local --unset core.checkRoundtripEncoding
	REM DESACTIVAR CONVERSION DE FINALES DE LINEA
	REM git config --local core.autocrlf false
	REM git config --local core.safecrlf false
)

echo.
pause
GOTO :RETURN_MENU

:CONVERT_UTF8_TO_ANSI
SET "CONV_DIR_OPT="
cls
echo ============================================================
echo  CONVERTIR FICHERO UTF-8 A ANSI ^(Windows-1252^)
echo ============================================================
echo.
echo Esta operacion convierte un fichero de UTF-8 a ANSI.
echo El fichero original se respalda con extension .utf8.bak
echo.
echo Directorios:
echo   1 - MAIN  ^(%BASEDIR%\%MAIN_DIR%^)
echo   2 - CLAUDE ^(%BASEDIR%\%CLAUDE_DIR%^)
echo.
set /p CONV_DIR_OPT=Selecciona directorio: 

IF "%CONV_DIR_OPT%"=="1" SET "CONV_DIR=%BASEDIR%\%MAIN_DIR%"
IF "%CONV_DIR_OPT%"=="2" SET "CONV_DIR=%BASEDIR%\%CLAUDE_DIR%"

IF NOT DEFINED CONV_DIR (
	echo Opcion invalida
	pause
	GOTO :RETURN_MENU
)

cd /d "%CONV_DIR%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No se puede acceder al directorio
	pause
	GOTO :RETURN_MENU
)

echo.
echo Directorio actual: %CD%
echo.
set /p CONV_FILE=Fichero a convertir (ruta relativa): 

IF NOT EXIST "%CONV_FILE%" (
	echo ^[ERROR^] El fichero no existe: %CONV_FILE%
	pause
	GOTO :RETURN_MENU
)

echo.
echo Creando backup: %CONV_FILE%.utf8.bak
copy /Y "%CONV_FILE%" "%CONV_FILE%.utf8.bak" >nul

echo Convirtiendo a ANSI ^(Windows-1252^)...
CALL :CONVERT_UTF8_TO_ANSI_CMD "%CD%\%CONV_FILE%"

IF NOT "%CONV_RESULT%"=="OK" (
	echo ^[ERROR^] Fallo la conversion
	echo Restaurando backup...
	copy /Y "%CONV_FILE%.utf8.bak" "%CONV_FILE%" >nul
) ELSE (
	echo ^[OK^] Conversion completada
	echo.
	echo Verificando resultado...
	CALL :CHECK_FILE_ENCODING "%CD%\%CONV_FILE%"
)

echo.
pause
GOTO :RETURN_MENU

:CONVERT_UTF8_TO_ANSI_CMD
REM Parametro: %1 = archivo a convertir
SET "FILEPATH=%~1"

REM Usar PowerShell para conversion UTF-8 -> Windows-1252
powershell -Command "& {
    try {
        $content = [System.IO.File]::ReadAllText('%FILEPATH%', [System.Text.Encoding]::UTF8)
        [System.IO.File]::WriteAllText('%FILEPATH%', $content, [System.Text.Encoding]::GetEncoding(1252))
        Write-Output 'OK'
    } catch {
        Write-Output 'ERROR: $($_.Exception.Message)'
    }
}" > "%TEMP%\conv_result.txt"

SET "CONV_RESULT=ERROR"
FOR /F "delims=" %%R IN ('type "%TEMP%\conv_result.txt" 2^>nul') DO SET "CONV_RESULT=%%R"
DEL "%TEMP%\conv_result.txt" 2>nul

EXIT /B 0

:CONVERT_ANSI_TO_UTF8
SET "CONV_DIR_OPT="
cls
echo ============================================================
echo  CONVERTIR FICHERO ANSI A UTF-8 ^(sin BOM^)
echo ============================================================
echo.
echo Esta operacion convierte un fichero de ANSI a UTF-8.
echo El fichero original se respalda con extension .ansi.bak
echo.
echo Directorios:
echo   1 - MAIN  ^(%BASEDIR%\%MAIN_DIR%^)
echo   2 - CLAUDE ^(%BASEDIR%\%CLAUDE_DIR%^)
echo.
set /p CONV_DIR_OPT=Selecciona directorio: 

IF "%CONV_DIR_OPT%"=="1" SET "CONV_DIR=%BASEDIR%\%MAIN_DIR%"
IF "%CONV_DIR_OPT%"=="2" SET "CONV_DIR=%BASEDIR%\%CLAUDE_DIR%"

IF NOT DEFINED CONV_DIR (
	echo Opcion invalida
	pause
	GOTO :RETURN_MENU
)

cd /d "%CONV_DIR%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No se puede acceder al directorio
	pause
	GOTO :RETURN_MENU
)

echo.
echo Directorio actual: %CD%
echo.
set /p CONV_FILE=Fichero a convertir (ruta relativa): 

IF NOT EXIST "%CONV_FILE%" (
	echo ^[ERROR^] El fichero no existe: %CONV_FILE%
	pause
	GOTO :RETURN_MENU
)

echo.
echo Creando backup: %CONV_FILE%.ansi.bak
copy /Y "%CONV_FILE%" "%CONV_FILE%.ansi.bak" >nul

echo Convirtiendo a UTF-8 ^(sin BOM^)...
CALL :CONVERT_ANSI_TO_UTF8_CMD "%CD%\%CONV_FILE%"

IF NOT "%CONV_RESULT%"=="OK" (
	echo ^[ERROR^] Fallo la conversion
	echo Restaurando backup...
	copy /Y "%CONV_FILE%.ansi.bak" "%CONV_FILE%" >nul
) ELSE (
	echo ^[OK^] Conversion completada
	echo.
	echo Verificando resultado...
	CALL :CHECK_FILE_ENCODING "%CD%\%CONV_FILE%"
)

echo.
pause
GOTO :RETURN_MENU

:CONVERT_ANSI_TO_UTF8_CMD
REM Parametro: %1 = archivo a convertir
SET "FILEPATH=%~1"

REM Usar PowerShell para conversion Windows-1252 -> UTF-8 sin BOM
powershell -Command "& {
    try {
        $content = [System.IO.File]::ReadAllText('%FILEPATH%', [System.Text.Encoding]::GetEncoding(1252))
        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText('%FILEPATH%', $content, $utf8NoBom)
        Write-Output 'OK'
    } catch {
        Write-Output 'ERROR: $($_.Exception.Message)'
    }
}" > "%TEMP%\conv_result.txt"

SET "CONV_RESULT=ERROR"
FOR /F "delims=" %%R IN ('type "%TEMP%\conv_result.txt" 2^>nul') DO SET "CONV_RESULT=%%R"
DEL "%TEMP%\conv_result.txt" 2>nul

EXIT /B 0

:CONVERT_ALL_TO_ANSI
SET "CONV_DIR_OPT="
cls
echo ============================================================
echo  CONVERTIR TODOS LOS FICHEROS VBA A ANSI
echo ============================================================
echo.
echo ^[ATENCION^] Esta operacion:
echo   - Busca todos los ficheros .bas, .cls, .frm, .vbs
echo   - Convierte los que estan en UTF-8 a ANSI ^(Windows-1252^)
echo   - Crea backup de cada fichero convertido ^(.utf8.bak^)
echo.
echo Directorios:
echo   1 - MAIN  ^(%BASEDIR%\%MAIN_DIR%^)
echo   2 - CLAUDE ^(%BASEDIR%\%CLAUDE_DIR%^)
echo.
set /p CONV_DIR_OPT=Selecciona directorio: 

IF "%CONV_DIR_OPT%"=="1" SET "CONV_DIR=%BASEDIR%\%MAIN_DIR%"
IF "%CONV_DIR_OPT%"=="2" SET "CONV_DIR=%BASEDIR%\%CLAUDE_DIR%"

IF NOT DEFINED CONV_DIR (
	echo Opcion invalida
	pause
	GOTO :RETURN_MENU
)

cd /d "%CONV_DIR%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No se puede acceder al directorio
	pause
	GOTO :RETURN_MENU
)

echo.
CALL :CONFIRM_DANGEROUS "convertir todos los ficheros VBA a ANSI"
IF ERRORLEVEL 1 GOTO :RETURN_MENU

echo.
echo Buscando y convirtiendo ficheros...
echo.

SET CONV_COUNT=0
SET CONV_ERRORS=0

FOR %%E IN (%TARGET_EXTENSIONS%) DO (
    FOR /R %%F IN (*.%%E) DO (
		CALL :CONVERT_SINGLE_FILE "%%F"
	)
)

echo.
echo ============================================================
echo  RESUMEN DE CONVERSION
echo ============================================================
echo Ficheros convertidos: %CONV_COUNT%
echo Errores: %CONV_ERRORS%
echo.
IF %CONV_COUNT% GTR 0 (
    echo RECOMENDACION: Despues de convertir, ejecuta:
    echo   git add --renormalize .
    echo   git commit -m "Convertido a ANSI"
    echo.
)
pause
GOTO :RETURN_MENU

:CONVERT_SINGLE_FILE
REM Parametro: %1 = ruta completa del fichero
SET "CONV_FILEPATH=%~1"
SET "CONV_FILENAME=%~nx1"

REM Verificar si es UTF-8 antes de convertir usando VBScript
SET "NEED_CONVERT="
CALL :CHECK_FILE_ENCODING "%CONV_FILEPATH%"
SET "NEED_CONVERT=!ENCODING_RESULT!"

REM Solo convertir si es UTF-8 (con o sin BOM)
IF NOT "%NEED_CONVERT%"=="UTF8-BOM" IF NOT "%NEED_CONVERT%"=="UTF8-NOBOM" (
	REM echo ^[SKIP^] %CONV_FILENAME% - ya esta en ANSI/ASCII
	EXIT /B 0
)

echo ^[CONV^] %CONV_FILENAME%...
copy /Y "%CONV_FILEPATH%" "%CONV_FILEPATH%.utf8.bak" >nul 2>&1

REM Convertir
CALL :CONVERT_UTF8_TO_ANSI_CMD "%CONV_FILEPATH%"

IF NOT "%CONV_RESULT%"=="OK" (
	echo        ^[ERROR^] Fallo la conversion
	SET /A CONV_ERRORS+=1
	copy /Y "%CONV_FILEPATH%.utf8.bak" "%CONV_FILEPATH%" >nul 2>&1
) ELSE (
	SET /A CONV_COUNT+=1
)
EXIT /B 0

REM ============================================================
REM SECCION: OPERACIONES SEGURAS
REM ============================================================

:FETCH
cls
echo ============================================================
echo  FETCH - Traer info remota SIN modificar archivos
echo ============================================================
echo.
echo Esta operacion es segura: no modifica tu working tree.
echo.
echo ---- MAIN ----
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul && git fetch origin
echo.
echo ---- CLAUDE ----
cd /d "%BASEDIR%\%CLAUDE_DIR%" 2>nul && git fetch origin
echo.
echo ^[OK^] Fetch completado
pause
GOTO :RETURN_MENU

:DIFF
SET "DIFF_MODE="
SET "DIFF_OPT="
cls
echo ============================================================
echo  COMPARAR CON BEYOND COMPARE
echo ============================================================
echo.
echo Modo de comparacion:
echo   1 - Comparar CARPETAS: MAIN vs CLAUDE ^(local-local^)
echo   2 - Comparar fichero a fichero via Git ^(difftool^)
echo   3 - Comparar LOCAL vs GitHub ^(clon temporal paralelo^)
echo.
set /p DIFF_MODE=Selecciona modo: 

REM ------------------------------------------------------------
REM MODO 1: LOCAL vs LOCAL
REM ------------------------------------------------------------
IF "%DIFF_MODE%"=="1" (
	echo.
	echo Abriendo Beyond Compare para comparar carpetas...
	echo   Izquierda: %BASEDIR%\%MAIN_DIR%
	echo   Derecha:   %BASEDIR%\%CLAUDE_DIR%
	echo.
	start "" "%BC%" "%BASEDIR%\%MAIN_DIR%" "%BASEDIR%\%CLAUDE_DIR%"
	echo Beyond Compare abierto. Puedes revisar las diferencias.
	pause
	GOTO :RETURN_MENU
)

REM ------------------------------------------------------------
REM MODO 2: GIT DIFFTOOL
REM ------------------------------------------------------------
IF "%DIFF_MODE%"=="2" (
	echo.
	echo Comparar via Git:
	echo   1 - MAIN: local vs GitHub
	echo   2 - CLAUDE: local vs GitHub
	echo.
	set /p DIFF_OPT=Selecciona repositorio: 
	
	IF "!DIFF_OPT!"=="1" (
		cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul && git difftool HEAD origin/main
	)
	IF "!DIFF_OPT!"=="2" (
		cd /d "%BASEDIR%\%CLAUDE_DIR%" 2>nul && git difftool HEAD origin/%CLAUDE_BRANCH%
	)
	GOTO :RETURN_MENU
)

REM ------------------------------------------------------------
REM MODO 3: LOCAL vs GITHUB (CLON TEMPORAL)
REM ------------------------------------------------------------
IF "%DIFF_MODE%"=="3" GOTO :DIFF_GITHUB

GOTO :DIFF_END

:DIFF_GITHUB
SET "DIFF_OPT="
cls
echo ============================================================
echo  COMPARAR LOCAL vs GITHUB ^(CLON TEMPORAL^)
echo ============================================================
echo.
echo   1 - MAIN
echo   2 - CLAUDE
echo.
set /p DIFF_OPT=Selecciona repositorio: 
:DIFF_GITHUB_FROM_MAIN
IF "%DIFF_OPT%"=="1" (
	SET SRC_DIR=%MAIN_DIR%
	SET SRC_BRANCH=main
)

IF "%DIFF_OPT%"=="2" (
	SET SRC_DIR=%CLAUDE_DIR%
	SET SRC_BRANCH=%CLAUDE_BRANCH%
)

IF NOT DEFINED SRC_DIR (
	echo Opcion invalida.
	pause
	GOTO :RETURN_MENU
)

SET TMP_DIR=%SRC_DIR%__GITHUB_TMP

echo.
echo Clonando rama "%SRC_BRANCH%" desde GitHub...
echo Carpeta temporal: %BASEDIR%\%TMP_DIR%
echo.

IF EXIST "%BASEDIR%\%TMP_DIR%" (
	echo ERROR: La carpeta temporal ya existe.
	echo Eliminala manualmente o cambiale el nombre.
	pause
	GOTO :RETURN_MENU
)

cd /d "%BASEDIR%"
git clone --branch "%SRC_BRANCH%" "%GITHUB_REPO%" "%TMP_DIR%"
IF ERRORLEVEL 1 (
	echo Error durante el clonado.
	pause
	GOTO :RETURN_MENU
)

echo.
echo Abriendo Beyond Compare...
echo   Izquierda ^(LOCAL^):  %BASEDIR%\!SRC_DIR!
echo   Derecha ^(GITHUB^):  %BASEDIR%\!TMP_DIR!
echo.
start /wait "" "%BC%" "%BASEDIR%\%SRC_DIR%" "%BASEDIR%\%TMP_DIR%"

echo Beyond Compare abierto. Puedes revisar las diferencias.
pause

echo.
set /p DELTMP=Eliminar carpeta temporal clonada? (S/N): 

IF /I "%DELTMP%"=="S" (
	echo Eliminando carpeta temporal...
	rmdir /s /q "%BASEDIR%\%TMP_DIR%"
	echo Carpeta eliminada.
) ELSE (
	echo Carpeta temporal conservada.
	echo %BASEDIR%\!TMP_DIR!
)

:DIFF_END
pause
GOTO :RETURN_MENU

:PULL_CLAUDE
SET "PULL_ACTION="
cls
echo ============================================================
echo  DESCARGAR NOVEDADES DE CLAUDE
echo ============================================================
echo.
echo Esta carpeta es solo para VER lo que Claude propone.
echo Normalmente NO deberias tener cambios locales aqui.
echo.

cd /d "%BASEDIR%\%CLAUDE_DIR%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No existe el directorio CLAUDE
	pause
	GOTO :RETURN_MENU
)

echo Directorio: %CD%
echo Rama: %CLAUDE_BRANCH%
echo.

REM Primero hacer fetch para ver el estado
echo Consultando GitHub...
git fetch origin

REM Verificar si hay divergencia
SET "LOCAL_AHEAD=0"
SET "LOCAL_BEHIND=0"
FOR /F "tokens=1,2" %%A IN ('git rev-list --count --left-right HEAD...origin/%CLAUDE_BRANCH% 2^>nul') DO (
	SET "LOCAL_AHEAD=%%A"
	SET "LOCAL_BEHIND=%%B"
)

REM Verificar cambios locales
SET "HAS_LOCAL_CHANGES=0"
FOR /F %%A IN ('git status --porcelain 2^>nul ^| find /c /v ""') DO (
	IF NOT "%%A"=="0" SET "HAS_LOCAL_CHANGES=1"
)

echo.
echo --- ESTADO ---
IF NOT "%LOCAL_BEHIND%"=="0" echo   GitHub tiene %LOCAL_BEHIND% cambios nuevos de Claude
IF NOT "%LOCAL_AHEAD%"=="0" echo   ^[!^] Tu carpeta tiene %LOCAL_AHEAD% cambios que GitHub no tiene
IF "%HAS_LOCAL_CHANGES%"=="1" echo   ^[!^] Hay ficheros modificados localmente
IF "%LOCAL_BEHIND%"=="0" IF "%LOCAL_AHEAD%"=="0" IF "%HAS_LOCAL_CHANGES%"=="0" (
	echo   ^[OK^] Ya estas actualizado, no hay nada que descargar
	pause
	GOTO :RETURN_MENU
)
echo.

REM Si hay divergencia o cambios locales, preguntar
IF NOT "%LOCAL_AHEAD%"=="0" GOTO :PULL_CLAUDE_DIVERGED
IF "%HAS_LOCAL_CHANGES%"=="1" GOTO :PULL_CLAUDE_DIVERGED

REM Caso simple: solo hay cambios en GitHub para descargar
echo Descargando cambios de Claude...
git pull --ff-only origin %CLAUDE_BRANCH%

IF ERRORLEVEL 1 (
	echo.
	echo ^[AVISO^] No se pudo hacer pull directo.
	GOTO :PULL_CLAUDE_DIVERGED
)

echo.
echo ^[OK^] Descarga completada
echo.
echo Los cambios de Claude estan en: %CD%
echo Usa "Comparar carpetas" para revisarlos.
pause
GOTO :RETURN_MENU

:PULL_CLAUDE_DIVERGED
echo ============================================================
echo  HAY DISCREPANCIAS
echo ============================================================
echo.
echo Tu carpeta CLAUDE tiene cambios que no deberian estar.
echo Esta carpeta es solo para LEER lo que Claude propone.
echo.
echo Opciones:
echo   1 - DESCARTAR mis cambios locales y quedarme con lo de Claude
echo       ^(Recomendado - sincroniza con GitHub^)
echo.
echo   2 - CONSERVAR mi version local y subirla a GitHub
echo       ^(Solo si sabes lo que haces^)
echo.
echo   0 - Cancelar y no hacer nada
echo.
set /p PULL_ACTION=Opcion: 

IF "%PULL_ACTION%"=="0" GOTO :RETURN_MENU

IF "%PULL_ACTION%"=="1" (
	echo.
	echo Descartando cambios locales...
	git reset --hard origin/%CLAUDE_BRANCH%
	IF ERRORLEVEL 1 (
		echo ^[ERROR^] No se pudo sincronizar
	) ELSE (
		echo ^[OK^] Sincronizado con GitHub
	)
	pause
	GOTO :RETURN_MENU
)

IF "%PULL_ACTION%"=="2" (
	echo.
	echo ^[AVISO^] Vas a sobreescribir lo que Claude ha subido a GitHub.
	choice /M "Seguro"
	IF ERRORLEVEL 2 GOTO :RETURN_MENU
	
	git add . >nul 2>&1
	git commit -m "Cambios locales en CLAUDE - %DATE%" >nul 2>&1
	git push --force origin %CLAUDE_BRANCH%
	IF ERRORLEVEL 1 (
		echo ^[ERROR^] No se pudo subir
	) ELSE (
		echo ^[OK^] Tu version subida a GitHub
	)
	pause
	GOTO :RETURN_MENU
)

GOTO :RETURN_MENU

:UPLOAD_FOR_REVIEW
cls
echo ============================================================
echo  SUBIR COPIA LOCAL PARA QUE CLAUDE LA REVISE
echo ============================================================
echo.
echo PROPOSITO:
echo   Cuando hay una DIVERGENCIA ^(tu tienes cambios Y Claude tiene otros^),
echo   puedes subir tu version a una rama temporal en GitHub para que
echo   Claude pueda verla y comparar ambas versiones.
echo.
echo COMO FUNCIONA:
echo   1. Se crea una rama temporal llamada "sergio/review-FECHA"
echo   2. Se sube tu version local de MAIN a esa rama
echo   3. Claude puede comparar su rama con la tuya en GitHub
echo   4. Despues de resolver, puedes borrar la rama temporal
echo.
echo NOTA: Esto NO modifica la rama "main" ni la rama de Claude.
echo       Es solo para revision.
echo.
pause

cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No existe el directorio MAIN
	pause
	GOTO :RETURN_MENU
)

REM Generar nombre de rama con fecha
FOR /F "tokens=1-3 delims=/" %%A IN ('date /t') DO SET "FECHA=%%C%%B%%A"
FOR /F "tokens=1-2 delims=: " %%A IN ('time /t') DO SET "HORA=%%A%%B"
SET "REVIEW_BRANCH=sergio/review-%FECHA%-%HORA%"

echo.
echo Se creara la rama: %REVIEW_BRANCH%
echo.
echo Contenido que se subira:
git status --short
echo.

choice /M "Continuar con la subida"
IF ERRORLEVEL 2 GOTO :RETURN_MENU

echo.
echo Paso 1: Guardando cambios locales ^(si los hay^)...
git add . >nul 2>&1
git stash >nul 2>&1

echo Paso 2: Creando rama temporal...
git checkout -b "%REVIEW_BRANCH%"
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No se pudo crear la rama
	git stash pop >nul 2>&1
	pause
	GOTO :RETURN_MENU
)

echo Paso 3: Recuperando cambios...
git stash pop >nul 2>&1

echo Paso 4: Haciendo commit de tu version...
git add .
git commit -m "Version de Sergio para revision - %DATE% %TIME%"

echo Paso 5: Subiendo a GitHub...
git push -u origin "%REVIEW_BRANCH%"
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No se pudo subir a GitHub
	echo Volviendo a la rama main...
	git checkout main
	git branch -D "%REVIEW_BRANCH%"
	pause
	GOTO :RETURN_MENU
)

echo.
echo ============================================================
echo  SUBIDA COMPLETADA
echo ============================================================
echo.
echo Tu version esta ahora en GitHub en la rama:
echo   %REVIEW_BRANCH%
echo.
echo URL para compartir con Claude:
echo   https://github.com/gautxori-yuyu/claude-vba-excel-abc/tree/%REVIEW_BRANCH%
echo.
echo INSTRUCCIONES PARA CLAUDE:
echo   Puedes decirle a Claude:
echo   "Compara tu rama %CLAUDE_BRANCH%
echo    con mi rama %REVIEW_BRANCH%
echo    y explicame las diferencias. Recomiendame que version
echo    deberia prevalecer para cada fichero diferente."
echo.
echo IMPORTANTE: Ahora estas en la rama temporal.
echo Para volver a trabajar normalmente, selecciona opcion V.
echo.
echo  V - Volver a la rama main ^(recomendado^)
echo  B - Borrar la rama temporal de GitHub
echo  Enter - Salir sin cambiar
echo.
SET "REVIEW_ACTION="
set /p REVIEW_ACTION=Accion: 

IF /I "%REVIEW_ACTION%"=="V" (
	echo.
	echo Volviendo a la rama main...
	git checkout main
	echo.
	echo ^[OK^] Ahora estas en la rama main
	echo.
	echo NOTA: Tus cambios siguen en la rama temporal %REVIEW_BRANCH%
	echo       Si despues de la revision quieres esos cambios en main,
	echo       tendras que publicarlos de nuevo o copiarlos manualmente.
)

IF /I "%REVIEW_ACTION%"=="B" (
	echo.
	echo Borrando rama temporal...
	git checkout main
	git branch -D "%REVIEW_BRANCH%"
	git push origin --delete "%REVIEW_BRANCH%"
	echo ^[OK^] Rama temporal borrada
)

pause
GOTO :RETURN_MENU

:MANAGE_UNTRACKED
cls
echo ============================================================
echo  GESTIONAR FICHEROS SIN SEGUIMIENTO
echo ============================================================
echo.
echo Los ficheros "sin seguimiento" ^(untracked^) son ficheros NUEVOS
echo que Git ve en tu carpeta pero que NO esta vigilando todavia.
echo.
echo Aqui puedes:
echo   - Ver cuales son
echo   - Decidir si quieres que Git los incluya en el repositorio
echo   - O ignorarlos permanentemente ^(anadirlos a .gitignore^)
echo.

cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No existe el directorio MAIN
	pause
	GOTO :RETURN_MENU
)

echo Directorio: %CD%
echo.

REM Contar ficheros sin seguimiento
SET "UNTRACKED_COUNT=0"
FOR /F %%A IN ('git ls-files --others --exclude-standard 2^>nul ^| find /c /v ""') DO SET "UNTRACKED_COUNT=%%A"

IF %UNTRACKED_COUNT% EQU 0 (
	echo ^[OK^] No hay ficheros sin seguimiento.
	echo      Todos los ficheros estan siendo vigilados por Git.
	echo.
	pause
	GOTO :RETURN_MENU
)

echo Hay %UNTRACKED_COUNT% fichero^(s^) sin seguimiento:
echo --------------------------------------------------------
git ls-files --others --exclude-standard
echo --------------------------------------------------------
echo.
echo OPCIONES:
echo   A - Anadir TODOS al seguimiento de Git
echo   S - Seleccionar ficheros uno a uno
echo   I - Ignorar ficheros ^(anadir a .gitignore^)
echo   V - Ver contenido de .gitignore actual
echo   Enter - Volver al menu
echo.
SET "UNTRACK_ACTION="
set /p UNTRACK_ACTION=Accion: 

IF /I "%UNTRACK_ACTION%"=="A" GOTO :UNTRACKED_ADD_ALL
IF /I "%UNTRACK_ACTION%"=="S" GOTO :UNTRACKED_SELECT
IF /I "%UNTRACK_ACTION%"=="I" GOTO :UNTRACKED_IGNORE
IF /I "%UNTRACK_ACTION%"=="V" GOTO :UNTRACKED_VIEW_GITIGNORE

GOTO :RETURN_MENU

REM ============================================================
REM SECCION: GESTION DE RAMAS
REM ============================================================

:CAMBIAR_RAMA_TRABAJO
SET "BRANCH_OPT="
cls
echo ============================================================
echo  CAMBIAR RAMA DE TRABAJO
echo ============================================================
echo.
echo Rama actual: %CLAUDE_BRANCH%
echo Carpeta actual: %CLAUDE_DIR%
echo.
echo Esta opcion te permite cambiar la rama remota contra la que
echo descargas los cambios de Claude u otra IA.
echo.
echo Consultando ramas disponibles en GitHub...
echo.

cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul
git fetch origin >nul 2>&1

echo Ramas remotas disponibles:
echo ------------------------------------------------------------
SET "BRANCH_COUNT=0"
FOR /F "tokens=1" %%B IN ('git branch -r 2^>nul ^| findstr /V "HEAD"') DO (
	SET /A BRANCH_COUNT+=1
	SET "BRANCH_!BRANCH_COUNT!=%%B"
	echo   !BRANCH_COUNT! - %%B
)
echo ------------------------------------------------------------
echo.
echo   0 - Cancelar ^(mantener rama actual^)
echo   M - Escribir nombre de rama manualmente
echo.
set /p BRANCH_OPT=Selecciona rama: 

IF "%BRANCH_OPT%"=="0" GOTO :RETURN_MENU
IF /I "%BRANCH_OPT%"=="M" GOTO :RAMA_MANUAL

REM Validar que sea un numero valido
SET "SELECTED_BRANCH="
IF %BRANCH_OPT% GTR 0 IF %BRANCH_OPT% LEQ %BRANCH_COUNT% (
	SET "SELECTED_BRANCH=!BRANCH_%BRANCH_OPT%!"
)

IF "%SELECTED_BRANCH%"=="" (
	echo Opcion no valida.
	pause
	GOTO :CAMBIAR_RAMA_TRABAJO
)

REM Quitar el prefijo origin/
SET "SELECTED_BRANCH=!SELECTED_BRANCH:origin/=!"
GOTO :CONFIGURAR_RAMA_SELECCIONADA

:RAMA_MANUAL
echo.
set /p SELECTED_BRANCH=Escribe el nombre de la rama (sin origin/): 
IF "%SELECTED_BRANCH%"=="" GOTO :RETURN_MENU

:CONFIGURAR_RAMA_SELECCIONADA
echo.
echo Rama seleccionada: %SELECTED_BRANCH%
echo.

REM Proponer nombre de carpeta por defecto (simplificado)
REM Extraer solo la parte final del nombre de la rama
FOR /F "tokens=1,2 delims=/" %%A IN ("%SELECTED_BRANCH%") DO (
	SET "BRANCH_PREFIX=%%A"
	SET "BRANCH_SUFFIX=%%B"
)

IF "%BRANCH_SUFFIX%"=="" (
	SET "DEFAULT_DIR=%SELECTED_BRANCH%-mirror"
) ELSE (
	SET "DEFAULT_DIR=%BRANCH_PREFIX%-mirror"
)

echo Carpeta local propuesta: %DEFAULT_DIR%
echo.
echo Puedes aceptar este nombre o escribir otro mas corto.
echo.
set /p NEW_CLAUDE_DIR=Nombre de carpeta (Enter para '%DEFAULT_DIR%'): 
IF "%NEW_CLAUDE_DIR%"=="" SET "NEW_CLAUDE_DIR=%DEFAULT_DIR%"

echo.
echo Se configurara:
echo   - Rama remota: %SELECTED_BRANCH%
echo   - Carpeta local: %NEW_CLAUDE_DIR%
echo.

choice /M "Confirmar cambio"
IF ERRORLEVEL 2 GOTO :RETURN_MENU

REM Actualizar variables
SET "CLAUDE_BRANCH=%SELECTED_BRANCH%"
SET "CLAUDE_DIR=%NEW_CLAUDE_DIR%"

REM Verificar si la carpeta existe
SET "CARPETA_NUEVA=0"
IF NOT EXIST "%BASEDIR%\%CLAUDE_DIR%" (
	SET "CARPETA_NUEVA=1"
	echo.
	echo Creando carpeta %CLAUDE_DIR%...
	mkdir "%BASEDIR%\%CLAUDE_DIR%"
	cd /d "%BASEDIR%\%CLAUDE_DIR%"
	
	echo Clonando rama %CLAUDE_BRANCH%...
	git clone --branch %CLAUDE_BRANCH% --single-branch %GITHUB_REPO% .
	IF ERRORLEVEL 1 (
		echo ^[ERROR^] No se pudo clonar la rama
		rmdir /s /q "%BASEDIR%\%CLAUDE_DIR%" 2>nul
		pause
		GOTO :RETURN_MENU
	)
	
	echo.
	echo Sincronizando con GitHub...
	git pull origin %CLAUDE_BRANCH%
) ELSE (
	echo.
	echo La carpeta %CLAUDE_DIR% ya existe.
	cd /d "%BASEDIR%\%CLAUDE_DIR%"
	
	REM Verificar si es un repo git
	git rev-parse --git-dir >nul 2>&1
	IF ERRORLEVEL 1 (
		echo ^[AVISO^] La carpeta existe pero no es un repositorio Git.
		echo Inicializando...
		git init
		git remote add origin %GITHUB_REPO%
	)
	
	REM Cambiar a la rama seleccionada
	echo Actualizando referencias...
	git fetch origin
	git checkout %CLAUDE_BRANCH% 2>nul
	IF ERRORLEVEL 1 (
		git checkout -b %CLAUDE_BRANCH% origin/%CLAUDE_BRANCH%
	)
	
	echo Sincronizando con GitHub...
	git pull origin %CLAUDE_BRANCH%
)

echo.
echo ============================================================
echo  RAMA DE TRABAJO CAMBIADA
echo ============================================================
echo.
echo Ahora trabajas con:
echo   - Rama: %CLAUDE_BRANCH%
echo   - Carpeta: %BASEDIR%\%CLAUDE_DIR%
echo.
IF "%CARPETA_NUEVA%"=="1" (
	echo La carpeta se ha creado y sincronizado automaticamente.
) ELSE (
	echo La carpeta se ha sincronizado con GitHub.
)
echo.
pause
GOTO :RETURN_MENU

REM ------------------------------------------------------------
REM SUBMENU: RAMAS HUERFANAS
REM ------------------------------------------------------------
:MENU_RAMAS_HUERFANAS
SET "ORPHAN_OPT="
cls
echo ============================================================
echo  GESTIONAR RAMAS HUERFANAS
echo ============================================================
echo.
echo Las ramas huerfanas son ramas sin historial compartido con main.
echo Utiles para codigo de otras IAs que quieres mantener separado.
echo.
echo Rama huerfana actual: %ORPHAN_BRANCH%
echo Carpeta actual: %ORPHAN_DIR%
echo.

REM Listar ramas huerfanas existentes en GitHub
echo Consultando GitHub...
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul
git fetch origin >nul 2>&1

echo.
echo Ramas en GitHub ^(posibles huerfanas^):
echo ------------------------------------------------------------
git branch -r 2>nul | findstr /V "HEAD main claude"
echo ------------------------------------------------------------
echo.
echo Opciones:
echo   1 - Crear nueva rama huerfana
echo   2 - Seleccionar rama huerfana existente para trabajar
echo   3 - Subir contenido a rama huerfana actual ^(%ORPHAN_BRANCH%^)
echo   0 - Volver
echo.
set /p ORPHAN_OPT=Opcion: 

IF "%ORPHAN_OPT%"=="1" GOTO :CREAR_RAMA_HUERFANA
IF "%ORPHAN_OPT%"=="2" GOTO :SELECCIONAR_RAMA_HUERFANA
IF "%ORPHAN_OPT%"=="3" GOTO :SUBIR_A_RAMA_HUERFANA
IF "%ORPHAN_OPT%"=="0" GOTO :MENU_HERRAMIENTAS

echo Opcion no valida.
pause
GOTO :MENU_RAMAS_HUERFANAS

:SELECCIONAR_RAMA_HUERFANA
cls
echo ============================================================
echo  SELECCIONAR RAMA HUERFANA
echo ============================================================
echo.
echo Escribe el nombre de la rama huerfana con la que quieres trabajar.
echo Ejemplo: qwen/mi-codigo, gemini/pruebas, etc.
echo.
set /p NEW_ORPHAN_BRANCH=Nombre de rama: 
IF "%NEW_ORPHAN_BRANCH%"=="" GOTO :MENU_RAMAS_HUERFANAS

REM Determinar nombre de carpeta
SET "NEW_ORPHAN_DIR=%NEW_ORPHAN_BRANCH:/=-%"
SET "NEW_ORPHAN_DIR=%NEW_ORPHAN_DIR%-mirror"

echo.
echo Se configurara:
echo   - Rama: %NEW_ORPHAN_BRANCH%
echo   - Carpeta: %NEW_ORPHAN_DIR%
echo.

choice /M "Confirmar"
IF ERRORLEVEL 2 GOTO :MENU_RAMAS_HUERFANAS

SET "ORPHAN_BRANCH=%NEW_ORPHAN_BRANCH%"
SET "ORPHAN_DIR=%NEW_ORPHAN_DIR%"

echo.
echo ^[OK^] Rama huerfana configurada: %ORPHAN_BRANCH%
echo.
pause
GOTO :MENU_RAMAS_HUERFANAS

:CREAR_RAMA_HUERFANA
cls
echo ============================================================
echo  CREAR NUEVA RAMA HUERFANA
echo ============================================================
echo.
echo Una rama huerfana no tiene historial compartido con otras ramas.
echo.
set /p NEW_ORPHAN_NAME=Nombre para la nueva rama (ej: gemini/codigo-nuevo): 
IF "%NEW_ORPHAN_NAME%"=="" GOTO :MENU_RAMAS_HUERFANAS

REM Verificar si la rama ya existe en GitHub
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul
git fetch origin >nul 2>&1
git ls-remote --heads origin %NEW_ORPHAN_NAME% | findstr "%NEW_ORPHAN_NAME%" >nul 2>&1
IF NOT ERRORLEVEL 1 (
	echo.
	echo ^[AVISO^] La rama %NEW_ORPHAN_NAME% ya existe en GitHub.
	echo.
	echo Opciones:
	echo   1 - Sobreescribir contenido de esa rama
	echo   2 - Elegir otro nombre
	echo   0 - Cancelar
	echo.
	SET "EXISTE_OPT="
	set /p EXISTE_OPT=Opcion: 
	
	IF "!EXISTE_OPT!"=="0" GOTO :MENU_RAMAS_HUERFANAS
	IF "!EXISTE_OPT!"=="2" GOTO :CREAR_RAMA_HUERFANA
	IF "!EXISTE_OPT!"=="1" (
		SET "ORPHAN_BRANCH=!NEW_ORPHAN_NAME!"
		SET "ORPHAN_DIR=!NEW_ORPHAN_NAME:/=-!-mirror"
		GOTO :PREPARAR_CARPETA_HUERFANA
	)
	GOTO :MENU_RAMAS_HUERFANAS
)

REM La rama no existe, crearla
SET "ORPHAN_BRANCH=%NEW_ORPHAN_NAME%"
SET "ORPHAN_DIR=%NEW_ORPHAN_NAME:/=-%"
SET "ORPHAN_DIR=%ORPHAN_DIR%-mirror"

:PREPARAR_CARPETA_HUERFANA
echo.
echo Configuracion:
echo   - Rama: %ORPHAN_BRANCH%
echo   - Carpeta: %ORPHAN_DIR%
echo.

cd /d "%BASEDIR%"

REM Verificar si la carpeta ya existe
IF EXIST "%ORPHAN_DIR%" (
	echo La carpeta %ORPHAN_DIR% ya existe.
	echo.
	echo Opciones:
	echo   1 - Usar carpeta existente ^(mantener ficheros^)
	echo   2 - Borrar y crear nueva
	echo   0 - Cancelar
	echo.
	SET "CARPETA_OPT="
	set /p CARPETA_OPT=Opcion: 
	
	IF "!CARPETA_OPT!"=="0" GOTO :MENU_RAMAS_HUERFANAS
	IF "!CARPETA_OPT!"=="2" (
		rmdir /s /q "%ORPHAN_DIR%"
	)
)

REM Crear carpeta si no existe
IF NOT EXIST "%ORPHAN_DIR%" (
	echo Creando carpeta %ORPHAN_DIR%...
	mkdir "%ORPHAN_DIR%"
)

cd /d "%BASEDIR%\%ORPHAN_DIR%"

REM Inicializar git si no es repo
git rev-parse --git-dir >nul 2>&1
IF ERRORLEVEL 1 (
	echo Inicializando repositorio Git...
	git init
	git remote add origin %GITHUB_REPO%
)

REM Crear rama huerfana
echo Creando rama huerfana %ORPHAN_BRANCH%...
git checkout --orphan %ORPHAN_BRANCH% 2>nul
git rm -rf --cached . >nul 2>&1
git commit --allow-empty -m "Rama inicial %ORPHAN_BRANCH%"

echo.
echo ============================================================
echo  CARPETA PREPARADA
echo ============================================================
echo.
echo Ubicacion: %CD%
echo Rama: %ORPHAN_BRANCH%
echo.
echo Ahora copia los ficheros que quieras subir a esta carpeta.
echo Despues usa "Subir contenido a rama huerfana" para publicarlos.
echo.
pause
GOTO :MENU_RAMAS_HUERFANAS

:SUBIR_A_RAMA_HUERFANA
cls
echo ============================================================
echo  SUBIR CONTENIDO A RAMA HUERFANA
echo ============================================================
echo.
echo Rama: %ORPHAN_BRANCH%
echo Carpeta: %ORPHAN_DIR%
echo.

IF NOT EXIST "%BASEDIR%\%ORPHAN_DIR%" (
	echo ^[ERROR^] La carpeta %ORPHAN_DIR% no existe.
	echo Primero debes crear la rama huerfana.
	pause
	GOTO :MENU_RAMAS_HUERFANAS
)

cd /d "%BASEDIR%\%ORPHAN_DIR%"

REM Verificar si es repo git
git rev-parse --git-dir >nul 2>&1
IF ERRORLEVEL 1 (
	echo ^[ERROR^] La carpeta no es un repositorio Git.
	echo Primero debes crear la rama huerfana.
	pause
	GOTO :MENU_RAMAS_HUERFANAS
)

echo Ficheros en la carpeta:
echo ------------------------------------------------------------
dir /b 2>nul
echo ------------------------------------------------------------
echo.

REM Contar ficheros
SET "FILE_COUNT=0"
FOR /F %%A IN ('dir /b 2^>nul ^| find /c /v ""') DO SET "FILE_COUNT=%%A"

IF "%FILE_COUNT%"=="0" (
	echo ^[AVISO^] La carpeta esta vacia.
	echo Copia ficheros antes de subir.
	pause
	GOTO :MENU_RAMAS_HUERFANAS
)

echo Se subiran %FILE_COUNT% elementos a la rama %ORPHAN_BRANCH%.
echo.

choice /M "Continuar"
IF ERRORLEVEL 2 GOTO :MENU_RAMAS_HUERFANAS

echo.
echo Preparando ficheros...
git add .

set /p COMMIT_MSG=Mensaje de commit (Enter para default): 
IF "%COMMIT_MSG%"=="" SET "COMMIT_MSG=Actualizacion %ORPHAN_BRANCH% - %DATE%"

git commit -m "%COMMIT_MSG%"

echo.
echo Subiendo a GitHub...
git push -u origin %ORPHAN_BRANCH%

IF ERRORLEVEL 1 (
	echo.
	echo ^[AVISO^] El push fallo. Puede que la rama tenga cambios diferentes.
	echo.
	choice /M "Forzar push (sobreescribir GitHub)"
	IF ERRORLEVEL 2 GOTO :MENU_RAMAS_HUERFANAS
	git push --force origin %ORPHAN_BRANCH%
)

echo.
echo ============================================================
echo  SUBIDA COMPLETADA
echo ============================================================
echo.
echo Rama %ORPHAN_BRANCH% actualizada en GitHub.
echo.
pause
GOTO :MENU_RAMAS_HUERFANAS

REM Mantener CREATE_ORPHAN_BRANCH como alias por compatibilidad
:CREATE_ORPHAN_BRANCH
GOTO :MENU_RAMAS_HUERFANAS

:UNTRACKED_ADD_ALL
echo.
echo Anadiendo TODOS los ficheros al seguimiento de Git...
echo.
git add .
echo.
echo ^[OK^] Ficheros anadidos. Ahora estan "preparados" ^(staged^).
echo     Para guardarlos permanentemente, usa "Publicar mis cambios" ^(opcion 3^).
echo.
echo Estado actual:
echo ------------------------------------------------
git status --short
echo ------------------------------------------------
echo.
pause
GOTO :RETURN_MENU

:UNTRACKED_SELECT
echo.
echo SELECCION INDIVIDUAL DE FICHEROS
echo ================================
echo Para cada fichero, indica si quieres anadirlo ^(S/N^):
echo.

FOR /F "delims=" %%F IN ('git ls-files --others --exclude-standard') DO (
	echo Fichero: %%F
	choice /M "Anadir este fichero al seguimiento"
	IF NOT ERRORLEVEL 2 (
		git add "%%F"
		echo       ^[Anadido^]
	) ELSE (
		echo       ^[Omitido^]
	)
	echo.
)

echo.
echo ^[OK^] Seleccion completada.
echo.
echo Estado actual:
echo ------------------------------------------------
git status --short
echo ------------------------------------------------
echo.
pause
GOTO :RETURN_MENU

:UNTRACKED_IGNORE
echo.
echo IGNORAR FICHEROS
echo ================
echo Los ficheros ignorados NO se subiran nunca a GitHub.
echo Se anaden al fichero .gitignore
echo.
echo Puedes ignorar:
echo   1 - Un fichero especifico ^(ej: config.local.txt^)
echo   2 - Una extension completa ^(ej: *.bak, *.tmp^)
echo   3 - Una carpeta completa ^(ej: temp/, backup/^)
echo.
set /p IGNORE_PATTERN=Escribe el patron a ignorar (o Enter para cancelar): 

IF "%IGNORE_PATTERN%"=="" GOTO :RETURN_MENU

echo.
echo Anadiendo "%IGNORE_PATTERN%" a .gitignore...
echo %IGNORE_PATTERN%>> .gitignore
echo.
echo ^[OK^] Patron anadido a .gitignore
echo.
echo Contenido actual de .gitignore:
echo --------------------------------
IF EXIST .gitignore (
	type .gitignore
) ELSE (
	echo ^[vacio^]
)
echo --------------------------------
echo.
pause
GOTO :MANAGE_UNTRACKED

:UNTRACKED_VIEW_GITIGNORE
echo.
echo CONTENIDO DE .gitignore
echo =======================
IF EXIST .gitignore (
	type .gitignore
) ELSE (
	echo ^[No existe fichero .gitignore^]
	echo.
	echo El fichero .gitignore contiene patrones de ficheros
	echo que Git debe ignorar permanentemente.
)
echo.
pause
GOTO :MANAGE_UNTRACKED

REM ============================================================
REM SECCION: CONFIGURACION
REM ============================================================

:CFG_BC
cls
echo ============================================================
echo  CONFIGURAR BEYOND COMPARE COMO DIFFTOOL
echo ============================================================
echo.
git config --global diff.tool bc
git config --global difftool.bc.cmd "\"%BC:\=/%\" \"$LOCAL\" \"$REMOTE\""
git config --global difftool.prompt false
echo.
echo ^[OK^] Beyond Compare configurado como difftool global
echo.
echo Verificando configuracion:
git config --global diff.tool
git config --global difftool.bc.cmd
echo.
pause
GOTO :RETURN_MENU

:PROTECT
cls
echo ============================================================
echo  PROTEGER FICHERO ^(skip-worktree^)
echo ============================================================
echo.
echo Esto hace que Git ignore cambios locales en el fichero.
echo Util para: configuracion local, secretos, plantillas.
echo.
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul
echo Directorio: %CD%
echo.
set /p PROTECT_FILE=Fichero a proteger (ruta relativa): 
IF "%PROTECT_FILE%"=="" GOTO :RETURN_MENU

git update-index --skip-worktree "%PROTECT_FILE%"
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No se pudo proteger el fichero
) ELSE (
	echo ^[OK^] Protegido: %PROTECT_FILE%
)
echo.
pause
GOTO :RETURN_MENU

:UNPROTECT
cls
echo ============================================================
echo  DESPROTEGER FICHERO
echo ============================================================
echo.
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul
echo Directorio: %CD%
echo.
echo Ficheros actualmente protegidos:
git ls-files -v | findstr "^S"
echo.
set /p UNPROTECT_FILE=Fichero a desproteger (ruta relativa): 
IF "%UNPROTECT_FILE%"=="" GOTO :RETURN_MENU

git update-index --no-skip-worktree "%UNPROTECT_FILE%"
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No se pudo desproteger el fichero
) ELSE (
	echo ^[OK^] Desprotegido: %UNPROTECT_FILE%
)
echo.
pause
GOTO :RETURN_MENU

REM ============================================================
REM SECCION: ETIQUETAS (TAGS)
REM ============================================================

:TAG_CREATE
cls
echo ============================================================
echo  CREAR ETIQUETA DE VERSION
echo ============================================================
echo.
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul
echo Directorio: %CD%
echo.
echo Etiquetas existentes:
git tag
echo.
set /p TAG_NAME=Nombre del tag (ej. v1.2): 
IF "%TAG_NAME%"=="" GOTO :RETURN_MENU

set /p TAG_MSG=Mensaje descriptivo: 
IF "%TAG_MSG%"=="" SET TAG_MSG=Version %TAG_NAME%

git tag -a "%TAG_NAME%" -m "%TAG_MSG%"
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No se pudo crear el tag
	pause
	GOTO :RETURN_MENU
)

echo.
choice /M "Publicar tag en GitHub"
IF ERRORLEVEL 2 (
	echo Tag creado solo localmente
) ELSE (
	git push origin "%TAG_NAME%"
	echo ^[OK^] Tag publicado
)
echo.
pause
GOTO :RETURN_MENU

:TAG_LIST
cls
echo ============================================================
echo  ETIQUETAS EXISTENTES
echo ============================================================
echo.
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul
echo.
echo --- Tags con detalle ---
git tag -n
echo.
pause
GOTO :RETURN_MENU

REM ============================================================
REM SECCION: LIMPIEZA
REM ============================================================

:CLEAN_PREVIEW
cls
echo ============================================================
echo  VISTA PREVIA DE LIMPIEZA
echo ============================================================
echo.
echo Esto muestra que ficheros NO versionados se borrarian.
echo NO borra nada todavia.
echo.
echo ---- MAIN ----
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul && git clean -nd
echo.
echo ---- CLAUDE ----
cd /d "%BASEDIR%\%CLAUDE_DIR%" 2>nul && git clean -nd
echo.
pause
GOTO :RETURN_MENU

:CLEAN_EXEC
cls
echo ============================================================
echo  EJECUTAR LIMPIEZA
echo ============================================================
echo.
echo ^[ATENCION^] Esto BORRA ficheros no versionados.
echo Los ficheros ignorados ^(.gitignore^) NO se tocan.
echo.
echo Que carpeta quieres limpiar?
echo   1 - MAIN  ^(%BASEDIR%\%MAIN_DIR%^)
echo   2 - CLAUDE ^(%BASEDIR%\%CLAUDE_DIR%^)
echo   3 - Ambas
echo   0 - Cancelar
echo.
SET "CLEAN_OPT="
set /p CLEAN_OPT=Opcion: 

IF "%CLEAN_OPT%"=="0" GOTO :RETURN_MENU
IF "%CLEAN_OPT%"=="" GOTO :RETURN_MENU

CALL :CONFIRM_DANGEROUS "ejecutar la limpieza"
IF ERRORLEVEL 1 GOTO :RETURN_MENU

echo.

IF "%CLEAN_OPT%"=="1" (
	echo Limpiando MAIN...
	cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul && git clean -fd
)
IF "%CLEAN_OPT%"=="2" (
	echo Limpiando CLAUDE...
	cd /d "%BASEDIR%\%CLAUDE_DIR%" 2>nul && git clean -fd
)
IF "%CLEAN_OPT%"=="3" (
	echo Limpiando MAIN...
	cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul && git clean -fd
	echo.
	echo Limpiando CLAUDE...
	cd /d "%BASEDIR%\%CLAUDE_DIR%" 2>nul && git clean -fd
)

echo.
echo ^[OK^] Limpieza completada
pause
GOTO :RETURN_MENU

REM ============================================================
REM SECCION: ACCIONES PELIGROSAS
REM ============================================================

:PUBLISH_MAIN
cls
echo ============================================================
echo  PUBLICAR MIS CAMBIOS EN MAIN EN GITHUB
echo ============================================================

REM --- Paso 0: Acceder a la carpeta MAIN ---
cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No existe el directorio MAIN
	pause
	GOTO :RETURN_MENU
)

REM Asegurar que estamos en la rama main
FOR /F "delims=" %%B IN ('git branch --show-current 2^>nul') DO SET "CURRENT_BRANCH=%%B"

IF NOT "%CURRENT_BRANCH%"=="main" (
    echo ^[INFO^] Cambiando a la rama main...
    git checkout main
    IF ERRORLEVEL 1 (
        echo ^[ERROR^] No se pudo cambiar a la rama main
        pause
        GOTO :RETURN_MENU
    )
)

REM --- Paso 1: Detectar y limpiar estados intermedios (rebase/merge) ---
SET "CLEANUP_DONE=0"
IF EXIST ".git\rebase-apply" (
	echo ^[AVISO^] Se ha detectado un rebase en curso. Abortando...
	git rebase --abort >nul 2>&1
	SET "CLEANUP_DONE=1"
)
IF EXIST ".git\rebase-merge" (
	echo ^[AVISO^] Se ha detectado un rebase en curso. Abortando...
	git rebase --abort >nul 2>&1
	SET "CLEANUP_DONE=1"
)
IF EXIST ".git\MERGE_HEAD" (
	echo ^[AVISO^] Se ha detectado un merge en curso. Abortando...
	git merge --abort >nul 2>&1
	SET "CLEANUP_DONE=1"
)
IF "%CLEANUP_DONE%"=="1" (
	echo Estado intermedio limpiado. Continuando...
	echo.
)

REM --- Paso 2: Confirmación explícita ---
CALL :CONFIRM_DANGEROUS "publicar MAIN en GitHub"
IF ERRORLEVEL 1 GOTO :RETURN_MENU

echo.
echo Estado actual:
echo ------------------------------------------------
git status
echo ------------------------------------------------
echo.

REM --- Paso 3: Analizar estado real del repositorio, si hay cambios REALES en archivos (modificados, nuevos, eliminados) ---
REM A. ¿Hay cambios en disco (archivos modificados, nuevos, eliminados)?
git status --porcelain | findstr /V "^$" >nul
SET "HAS_WORKING_CHANGES=!ERRORLEVEL!"

REM B. ¿Hay commits locales no subidos?
FOR /F "tokens=1,2" %%A IN ('git rev-list --count --left-right HEAD...origin/main 2^>nul') DO (
	SET "LOCAL_AHEAD=%%A"
	SET "LOCAL_BEHIND=%%B"
)
IF NOT DEFINED LOCAL_AHEAD SET "LOCAL_AHEAD=0"
IF NOT DEFINED LOCAL_BEHIND SET "LOCAL_BEHIND=0"

REM C. ¿Estamos sincronizados?
FOR /F "tokens=1" %%A IN ('git rev-parse HEAD 2^>nul') DO SET "LOCAL_HEAD=%%A"
FOR /F "tokens=1" %%A IN ('git rev-parse origin/main 2^>nul') DO SET "REMOTE_HEAD=%%A"
IF "!LOCAL_HEAD!"=="!REMOTE_HEAD!" (
	SET "IS_SYNCED=1"
) ELSE (
	SET "IS_SYNCED=0"
)

REM --- Paso 4: Actuar según el contexto real ---

REM CASO 1: No hay cambios en disco NI commits pendientes
IF "%HAS_WORKING_CHANGES%"=="1" IF "%LOCAL_AHEAD%"=="0" (
	echo ^[OK^] No hay cambios en archivos ni commits pendientes.
	echo      Tu carpeta ya está sincronizada con GitHub.
	pause
	GOTO :RETURN_MENU
)

REM CASO 2: Hay cambios en disco, hay que hacer commit primero
IF "%HAS_WORKING_CHANGES%"=="0" (
	echo Preparando ficheros...
	git add --renormalize .
	git add . 2> "%TEMP%\git_add_error.log"
	git add . 2> "%TEMP%\git_add_error.log"
	IF ERRORLEVEL 1 (
		echo.
		echo ^[ERROR^] git add ha fallado.
		echo ----------------------------------------
		type "%TEMP%\git_add_error.log"
		echo ----------------------------------------
		echo.
		SET /p DO_AUDIT_NOW=Deseas ejecutar una auditoría de codificación ahora?^(S/N^)^: 

		IF /I "!DO_AUDIT_NOW!"=="S" CALL :DO_AUDIT "%BASEDIR%\%MAIN_DIR%"
		IF /I "!DO_AUDIT_NOW!"=="SI" CALL :DO_AUDIT "%BASEDIR%\%MAIN_DIR%"

		pause
		GOTO :RETURN_MENU
	)
	echo.
	echo Ficheros que se publicarán:
	git status --porcelain
	echo.
	choice /M "Continuar con el commit y publicación"
	IF ERRORLEVEL 2 GOTO :RETURN_MENU

	set /p COMMIT_MSG=Describe los cambios ^(Enter para mensaje por defecto^):
	IF "!COMMIT_MSG!"=="" SET "COMMIT_MSG=Cambios de Sergio - %DATE%"
	echo.
	echo Guardando cambios ^(commit^)...
	git commit -m "!COMMIT_MSG!"

	REM Tras el commit, actualizar contadores
	FOR /F "tokens=1,2" %%A IN ('git rev-list --count --left-right HEAD...origin/main 2^>nul') DO (
		SET "LOCAL_AHEAD=%%A"
		SET "LOCAL_BEHIND=%%B"
	)
	IF NOT DEFINED LOCAL_AHEAD SET "LOCAL_AHEAD=0"
	IF NOT DEFINED LOCAL_BEHIND SET "LOCAL_BEHIND=0"
)

REM CASO 3: Ahora tenemos un commit local (LOCAL_AHEAD >= 1)
REM Pero ¿GitHub tiene commits que no tenemos?
IF NOT "%LOCAL_BEHIND%"=="0" IF NOT "%LOCAL_BEHIND%"=="" (
	echo.
	echo ================== ESTADO DE SINCRONIZACION ==================
	echo Commits SOLO en tu carpeta ^(local^):   !LOCAL_AHEAD!
	echo Commits SOLO en GitHub ^(sin integrar^): !LOCAL_BEHIND!
	echo ==============================================================

	echo.
	echo ^[AVISO IMPORTANTE^]
	echo GitHub tiene commits que TU NO TIENES en tu carpeta local.
	echo.
	echo Opciones:
	echo   1 - Forzar MI version ^(sobreescribir GitHub^)
	echo   2 - Descartar mis commits y alinear con GitHub
	echo   3 - Ver diferencias con Beyond Compare antes de decidir
	echo   0 - Cancelar
	SET "CONFLICT_OPT="
	set /p CONFLICT_OPT=Opcion:

	IF "!CONFLICT_OPT!"=="0" GOTO :RETURN_MENU

	IF "!CONFLICT_OPT!"=="1" (
		echo.
		echo ^[PELIGRO^] Vas a SOBRESCRIBIR GitHub con tu version local.
		choice /M "Confirmar push forzado"
		IF ERRORLEVEL 2 GOTO :RETURN_MENU
		git push --force-with-lease origin main
		IF ERRORLEVEL 1 git push --force origin main
		GOTO :PUBLISH_DONE
	)

	IF "!CONFLICT_OPT!"=="2" (
		echo.
		echo Descartando commits locales y alineando con GitHub...
		git reset --hard origin/main
		echo Carpeta MAIN ahora igual que GitHub.
		pause
		GOTO :RETURN_MENU
	)

	IF "!CONFLICT_OPT!"=="3" (
		echo.
		echo Preparando comparacion con Beyond Compare...
		SET DIFF_MODE=3
		SET DIFF_OPT=1
		GOTO :DIFF_GITHUB_FROM_MAIN
	)

	echo Opcion invalida.
	pause
	GOTO :RETURN_MENU
)

REM CASO 4: Todo listo, push normal
echo.
echo Subiendo tus commits a GitHub...
git push origin main

IF ERRORLEVEL 1 (
	echo.
	echo ^[ERROR^] GitHub no aceptó el push.
	choice /M "Forzar el push (solo si sabes lo que haces)"
	IF ERRORLEVEL 2 GOTO :RETURN_MENU
	git push --force-with-lease origin main
	IF ERRORLEVEL 1 git push --force origin main
)

:PUBLISH_DONE
echo.
echo ============================================================
echo  PUBLICACION COMPLETADA
echo ============================================================
echo Tus cambios están en GitHub. Claude puede verlos.
echo.
pause
GOTO :RETURN_MENU

:CLONE_REPOS
cls
echo ============================================================
echo  ^[!!^] CLONAR REPOSITORIOS
echo ============================================================
echo.
echo ^[PELIGRO^] Esta operacion BORRA los directorios existentes:
echo   - %BASEDIR%\%MAIN_DIR%
echo   - %BASEDIR%\%CLAUDE_DIR%
echo.
echo Y los clona de nuevo desde:
echo   %GITHUB_REPO%
echo.
CALL :CONFIRM_DANGEROUS "BORRAR y re-clonar los repositorios"
IF ERRORLEVEL 1 GOTO :RETURN_MENU

CALL :CONFIRM_DANGEROUS "CONFIRMAR por segunda vez el BORRADO"
IF ERRORLEVEL 1 GOTO :RETURN_MENU

cd /d "%BASEDIR%"

echo.
echo Borrando %MAIN_DIR%...
IF EXIST "%MAIN_DIR%" rmdir /s /q "%MAIN_DIR%"

echo Borrando %CLAUDE_DIR%...
IF EXIST "%CLAUDE_DIR%" rmdir /s /q "%CLAUDE_DIR%"

echo.
echo Clonando MAIN...
git clone %GITHUB_REPO% %MAIN_DIR%
IF ERRORLEVEL 1 (
	echo ^[ERROR^] Fallo al clonar MAIN
	pause
	GOTO :RETURN_MENU
)

echo.
echo Clonando CLAUDE...
git clone %GITHUB_REPO% %CLAUDE_DIR%
IF ERRORLEVEL 1 (
	echo ^[ERROR^] Fallo al clonar CLAUDE
	pause
	GOTO :RETURN_MENU
)

echo.
echo Configurando rama CLAUDE...
cd "%CLAUDE_DIR%"
git fetch origin
git checkout -B claude-review origin/%CLAUDE_BRANCH%

echo.
echo ^[OK^] Clonacion completada
echo.
echo MAIN: rama main
echo CLAUDE: rama %CLAUDE_BRANCH%
pause
GOTO :RETURN_MENU

:RESYNC_CLAUDE
cls
echo ============================================================
echo  ^[!!^] RE-SINCRONIZAR CLAUDE ^(reset --hard^)
echo ============================================================
echo.
echo ^[PELIGRO^] Esta operacion:
echo   - DESCARTA todos los cambios locales en CLAUDE
echo   - Sincroniza con origin/%CLAUDE_BRANCH%
echo.
CALL :CONFIRM_DANGEROUS "descartar cambios locales en CLAUDE"
IF ERRORLEVEL 1 GOTO :RETURN_MENU

cd /d "%BASEDIR%\%CLAUDE_DIR%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No existe el directorio CLAUDE
	pause
	GOTO :RETURN_MENU
)

echo.
echo Trayendo cambios remotos...
git fetch origin

echo.
echo Ejecutando reset --hard...
git reset --hard origin/%CLAUDE_BRANCH%

echo.
echo ^[OK^] CLAUDE re-sincronizado
git log --oneline -5
echo.
pause
GOTO :RETURN_MENU

:FORCE_PUSH_MAIN
cls
echo ============================================================
echo  ^[!!^] PUSH FORZADO A MAIN
echo ============================================================
echo.
echo ^[EXTREMADAMENTE PELIGROSO^]
echo.
echo Esta operacion FUERZA que tu version local PREVALEZCA
echo sobre lo que hay en GitHub, DESTRUYENDO cualquier commit
echo que exista solo en el remoto.
echo.
echo Usa esto SOLO si sabes exactamente lo que haces.
echo.
CALL :CONFIRM_DANGEROUS "ejecutar PUSH FORZADO"
IF ERRORLEVEL 1 GOTO :RETURN_MENU

CALL :CONFIRM_DANGEROUS "CONFIRMAR por segunda vez PUSH FORZADO"
IF ERRORLEVEL 1 GOTO :RETURN_MENU

echo.
echo Escribe FORZAR para continuar:
set /p FORCE_CONFIRM=
IF NOT "%FORCE_CONFIRM%"=="FORZAR" (
	echo Cancelado.
	pause
	GOTO :RETURN_MENU
)

cd /d "%BASEDIR%\%MAIN_DIR%" 2>nul
IF ERRORLEVEL 1 (
	echo ^[ERROR^] No existe el directorio MAIN
	pause
	GOTO :RETURN_MENU
)

echo.
git add --renormalize .
echo Ejecutando push --force-with-lease...
git push --force-with-lease origin main

IF ERRORLEVEL 1 (
	echo.
	echo ^[ERROR^] Git rechazo el push.
	echo El remoto cambio despues de tu ultima sincronizacion.
	echo.
	echo Opciones:
	echo   1. Vuelve a ejecutar fetch + este comando
	echo   2. Usa --force ^(MAS peligroso, sin verificacion^)
	echo.
	choice /M "Usar --force (sin verificacion)"
	IF ERRORLEVEL 2 GOTO :RETURN_MENU
	
	CALL :CONFIRM_DANGEROUS "usar --force SIN verificacion"
	IF ERRORLEVEL 1 GOTO :RETURN_MENU
	
	git push --force origin main
)

echo.
echo ^[OK^] Push forzado completado
pause
GOTO :RETURN_MENU

REM ============================================================
REM FUNCIONES AUXILIARES
REM ============================================================

:CONFIRM_DANGEROUS
REM Parametro: %1 = descripcion de la accion
echo.
echo ============================================
echo  CONFIRMACION REQUERIDA
echo ============================================
echo.
echo Estas a punto de: %~1
echo.
choice /M "Continuar"
IF ERRORLEVEL 2 (
	echo Cancelado.
	EXIT /B 1
)
EXIT /B 0

:NORMALIZE_PATH
REM Determina BASEDIR basandose en donde se ejecuta el script
SET BASEDIR=%CD%

REM Caso 1: Estamos en BASE (existen las subcarpetas)
IF EXIST "%BASEDIR%\%MAIN_DIR%" (
	EXIT /B 0
)
IF EXIST "%BASEDIR%\%CLAUDE_DIR%" (
	EXIT /B 0
)

REM Caso 2: Estamos dentro de MAIN_DIR
FOR %%I IN ("%BASEDIR%") DO SET CURRENT_FOLDER=%%~nxI
IF "%CURRENT_FOLDER%"=="%MAIN_DIR%" (
	cd /d "%BASEDIR%\.."
	SET BASEDIR=%CD%
	EXIT /B 0
)

REM Caso 3: Estamos dentro de CLAUDE_DIR
IF "%CURRENT_FOLDER%"=="%CLAUDE_DIR%" (
	cd /d "%BASEDIR%\.."
	SET BASEDIR=%CD%
	EXIT /B 0
)

REM Si llegamos aqui, carpeta incorrecta
echo.
echo ============================================
echo  ERROR: Carpeta incorrecta
echo ============================================
echo.
echo Ejecuta este script desde:
echo   - La carpeta BASE ^(que contiene %MAIN_DIR% y/o %CLAUDE_DIR%^)
echo   - %MAIN_DIR%
echo   - %CLAUDE_DIR%
echo.
echo Carpeta actual: %CD%
echo.
pause
EXIT /B 1

:RETURN_MENU
cd /d "%BASEDIR%"
GOTO :MENU_PRINCIPAL

:END_SCRIPT
cd /d "%START_DIR%"
ENDLOCAL
echo.
echo Hasta pronto.
EXIT /B 0
