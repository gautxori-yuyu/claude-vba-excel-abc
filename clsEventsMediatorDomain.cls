VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsMediatorDomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "3-Dominio.2-2-Mediadores"
'@ModuleDescription "DEPRECADA - Clase vaciada en Round 3. Sus responsabilidades han sido migradas:"
' - Eventos Template: -> FileSystemTemplateProvider (WithEvents sobre clsFSMonitoringCoord)
' - Eventos Gas: -> FileSystemTechnicalCalcProvider (WithEvents sobre clsFSMonitoringCoord)
' - Eventos Monitoring (Error/Reconnected/Failed): eliminados (clsFSWatcher gestiona internamente)
' - ActiveFileSessionChanged: eliminado (clsRibbon gestiona ActiveFileChanged directamente)
' - ContextInvalidated/Reinitialized: eliminados (no aportan valor, el host VBA falla antes)
' - OpportunityCollectionUpdated: eliminado (logging directo en AddOpportunity/RemoveOpportunity)
' - CurrentOpportunityChanged: logging directo, lógica de MoverAOp pendiente de FileLocationService
'
' Esta clase se mantiene como stub para no romper la compilación si algún módulo la referencia.
' Se puede eliminar del proyecto VBA cuando se verifique que no hay referencias externas.
Option Explicit
