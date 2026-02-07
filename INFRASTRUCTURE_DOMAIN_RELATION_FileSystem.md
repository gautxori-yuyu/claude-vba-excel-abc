# Descripcion la relación entre el nivel de infraestructura (Sistema de archivos   ) y el contexto de dominio de la aplicación

## OPORTUNIDADES
El concepto de oportunidad representa una oportunidad comercial De venta de compresores.
Ese concepto se representa físicamente, A nivel de infraestructura , en el sistema de archivos Por una carpeta Que contiene Varias subcarpetas  Y a su vez Éstas pueden contener varias subcarpetas y ficheros. En el nivel De la carpeta de oportunidad también pueden contenerse ficheros.
Para que una carpeta pueda ser considerada carpeta de oportunidad tiene que cumplir los siguientes criterios :
-  Su nombre tiene que Casar con el patrón de expresión regular FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN.
- Debe tener una o más su carpeta con los nombres siguientes : 
1.ESPECIFICACION
2.CALCULO TECNICO
3.VALORACION ECONOMICA
4.OFERTA COMERCIAL
5.COMUNICACIONES
6.PEDIDO
7.LANZAMIENTO DE PROYECTO
8.ORDENES DE COMPRA
Por tanto la identificación de si una carpeta es oportunidad se hará a nivel de infraestructura Y el método que lo identifica Será un método implementado de la interfaz IOpportunity.

### COMPRESORES
 Una oportunidad comercial puede tener asociados 1 o más compresores .  Por tanto cada oportunidad comercial gestionará una lista de instancias de clases clsCompresor , Clase todavía pendiente de implementar .
 Cada compresor Puede tener asociados 1 o varios cálculos técnicos Y tendrá asociada una valoración económica (Que puede presentar Una o varias revisiones) Y una oferta comercial; que puede venir Definida Por una oferta tipo budget (Que a su vez se puede gestionar en una o varias revisiones ), Y/O por una "cotización formal" (Que también puede tener una o varias revisiones ) .
 También la especificación Puede tener documentos Diferenciados para cada compresor , O puede tener documentos Que Especifiquen cómo deben ser todos los compresores .

### REVISIONES
Para permitir hacer un seguimiento de revisiones en el contenido de las oportunidades, dentro de las carpetas Que se han identificado como subcarpetas de la que identifica la oportunidad  Se pueden encontrar carpetas Cuyo nombre  nombre  siga un patrón (De expresión regular) de la forma "\brev\.?\s*(\d+)\.+". El valor numérico Que captura la primera su expresión de esa expresión regular representaría el valor de la revisión, Sea de las especificaciones, de los cálculos técnicos, de las valoraciones económicas, etc .Por tanto la gestión de revisiones se hace por "Subsecciones de la oportunidad", y no por la oportunidad en sí misma.
Esta estructura de su carpeta es para identificar las revisiones no es necesaria Dado que los ficheros de cálculo técnico, valoración económica y oferta comercial en su nombre tienen También Un "campo" (Una parte del nombre capturada a partir de un patrón de expresión regular  ) que define la revisión 

EJEMPLO DE ESTRUCTURA DE CARPETAS ESPERADA
El sistema trabaja con carpetas de oportunidades que deben tendran una estructura jerárquica Similar a la siguiente :
[num oportunidad 9 digitos] - [Nombre cliente] (- [breve descr. proyecto]) - [modelos compresores]
+-- 2.CALCULO TECNICO/
¦   +-- AAA12345_01_calc.xlsx
¦   +-- AAA12345_01_calc_multi.xlsx
¦   +-- AAA12345_02_calc.xlsx
¦   +-- AAA12345_03vrev01_calc.xlsx
¦   +-- [otros archivos ]
+-- 3.VALORACION ECONOMICA/
¦   +-- rev 1
¦   ¦   +-- 123456789-01.xlsx
¦   ¦   +-- 123456789-02.xlsx
¦   ¦   +-- 123456789-02.xlsm
¦   +-- rev 2
¦   ¦   +-- 123456789-01 rev.2.xlsx
¦   ¦   +-- 123456789-01 rev.2.xlsm
¦   ¦   +-- 123456789-02 rev.2.xlsx
¦   ¦   +-- 123456789-02 rev.2.xlsm
¦   +-- plantilla de referencia para calculo ingenieria y margenes.xlsm
+-- 4.OFERTA COMERCIAL/
    +-- [documentos finales de oferta]


### ESPECIFICACIONES
En la carpeta 1.ESPECIFICACION se contienen habitualmente documentos de especificación del compresor. Estos documentos Pueden ser documentos de Word, documentos de PDF, u otros; puede haber un documento, puede haber varios; Puede haber subcarpetas, puede haber ficheros empaquetados ...



### CALCULOS TECNICOS
La carpeta 2.CALCULO TECNICO Contiene ficheros de cálculos técnicos que pueden estar organizados por subcarpetas.
Los cálculos se identifican por un "código", "número" o "identificador de cálculo". Y además Un mismo cálculo puede  tener Distintas "opciones de cálculo", Que Pueden implicar distintos datos de entrada Relacionados entre sí.  
De un mismo cálculo puede haber varios ficheros, de tal manera que El contenido de esta carpeta de cálculos técnicos se organizará como una colección de ficheros Que a su vez se agrupan por el número de cálculo que les identifica.

Para que un fichero sea considerado cálculo técnico Tiene que cumplir requisitos relativos a su denominación y / o el contenido del fichero. En el caso de ficheros de Excel Esa comprobación puede venir determinada por Las hojas de Excel que contenga Y el formato Y contenido de estas.
Con objeto de determinar si un fichero es un cálculo técnico, y de qué tipo lo es, por tanto, se implementarán funciones que chequeen Tanto el nombre de los ficheros como el contenido De los Documentos .

El siguiente código extrae información del nombre Del cálculo, Con una excepción de un Tipo de fichero cuyo nombre empieza por la palabra "Ayuda_":
			Const CalcFNamePattern = "(([A-Z]{3}\d{5})_\d{2})(?:[_ \-]*rev[\._ \-]*(\d+))?.+?\.(?:txt|rtf|xlsx)$"
			regex.Pattern = CalcFNamePattern
			If regex.Test (fich.Name) And Left (fich.Name,2) <> "~$" Then
				strCalcNum = regex.Execute (fich.Name).item(0).submatches(1)
				strCalcOpc = regex.Execute (fich.Name).item(0).submatches(0)
				strCalcRev = regex.Execute (fich.Name).item(0).submatches(2) 
			ElseIf InStr (fich.Name, "Ayuda_") > 0 And InStr (fich.Name, ".txt") > 0 Then
				regex.Pattern = "Num cálculo : (([A-Z]{3}\d{5})_\d{2})"		
				For Each match In regex.Execute(fich.OpenAstextStream.ReadAll)
					strCalcNum = match.submatches(1)
					strCalcOpc = match.submatches(0)
					strCalcRev = Empty
				Next
				fich.Close
			Else
				MsgLog "Fichero descartado en carpeta de cálculos técnicos: " & fich.Name
			End If
strCalcNum Es un valor de cadena compuesto por una cadena inicial de 3 letras Y una secuencia de números que Identifica el "número de cálculo" 
strCalcOpc Es un valor numérico que Identifica la "opción de cálculo" 
strCalcRev Es un valor numérico que Identifica el cálculo al que revisa Esa opción. En este punto se distingue lo que es una revisión de los cálculos identificada como subcarpeta dentro de la carpeta 2.CALCULO TECNICO, Con el criterio antes definido De lo que sería una revisión de una opción de cálculo Que es lo que viene a identificar este campo. Si este campo toma el valor X Quiere decir que la opción de cálculo "strCalcOpc", Es una revisión de la opción de cálculo "X"; Por tanto los cambios que se hubieran hecho en Los datos de entrada para la opción de cálculo "strCalcOpc", Debe suponerse que Han tomado como partida Los datos de entrada para la opción de cálculo "X", Con objeto de modificar algunos de los resultados que aquel cálculo hubiera Obtenido.

El siguiente código determina el patrón de expresión regular de los nombres de ficheros de cálculos:
Const GasVBNetExportedFNamesPattern = "(?:Antipul_|ABC_(Gas_Cooler|Aircooler|Reducer|Main Motor|Instrumentation|Gas_Filter|Frequency Converter|Cooling Water Pump|Dryer|Piston_rider_ring_selection|Cooling Water Tower|Pressure_Safety_Valve|Valves_selection)\-)?([A-Z]{3}\d{5}_\d{2})(?:[_ \-]*rev[\._ \-]*(\d+))?(_calc(?:_multi)?)?.*?(\.(?:xlsx|rtf|txt))$"


Patrón obligatorio: [AAA12345]_[01][_calc][_rev_N].xlsx

AAA12345: Código de cálculo (3 letras + 5 dígitos)
01: Número de opción del cálculo
_calc o _calc_multi: Sufijo que identifica el archivo como resultado de cálculo
_rev_N (opcional): Número de revisión


Ejemplos válidos:

ABC12345_01_calc.xlsx
ABC12345_02_calc_multi.xlsx
ABC12345_01_rev_2_calc.xlsx


Archivos complementarios (mismo código de cálculo):

ABC12345_01.txt (exportación texto)
ABC12345_01.rtf (exportación RTF)
Antipul_ABC12345_01.xlsx (cálculo de antipulsadores)
ABC_Gas_Cooler-ABC12345_01.xlsx (enfriadores)
ABC_Aircooler-ABC12345_01.xlsx (aeroenfriadores)
Y otros componentes auxiliares


### VALORACIONES ECONOMICAS
La estructura de valoraciones económicas es similar a la de los cálculos técnicos Con la diferencia de que El patrón para identificar los ficheros De valoraciones económicas Corresponde al definido Por:
Const strQuoteNrPattern = "\d{9}(?:[\-_]\d+)?"
Dim strQuoteNrRevPattern
strQuoteNrRevPattern = "(" & strQuoteNrPattern & ")(?:[ \-_]*rev\.?[ \-_]*\d+\b)?"
regex.Pattern = "^" & strQuoteNrRevPattern & ".*\.xls[xm]$"

### OFERTAS COMERCIALES
El patrón de expresión regular de los ficheros De ofertas comerciales Sean formato budget O sea en formato cotización Es similar Al patrón de expresión regular de la oportunidad ,FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN, ya definido en el código del programa .



# funcionalidad que debe implementar el programa:

Identificar todas las carpetas de oportunidades dentro de una carpeta dada Por la variable De configuración. Esa funcionalidad ya se encuentra parcialmente implementada 
Para cada oportunidad Procesar de forma recursiva Su carpeta dentro de ella e Identificar Todos los ficheros Y carpetas Que hubiera,Comprobando los patrones antes descritos y por tanto : 
- Los cálculos técnicos Que hay dentro de la carpeta De oportunidad.
- La valoraciones económicas ,
- Y todos los demás documentos que hubiera 
- Se identificaran también las carpetas antes Descritas Como sus carpetas de oportunidad  
- Si hubiera Algún fichero fuera de la carpeta En la que se considera que debe estar, Sea de cálculos técnicos, de valoraciones económicas o de cualquier otro tipo, Se registrará Como fichero del tipo que le corresponda y se le añadirá Una marca de Que está mal ubicado     
- También se registrará si hay subcarpetas de revisiones Dentro de las subcarpetas de oportunidad antes identificadas . Por tanto tanto Las clases que posteriormente se implementarán para gestionar los cálculos técnicos como las valoraciones económicas Como las ofertas comerciales etc Gestionarán listas de revisiones , Y dentro de cada revisión se gestionarán listas de ficheros Que pertenecen a esa revisión   