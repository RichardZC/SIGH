VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOMovimientoHistoriaClinica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla MovimientosHistoriaClinica
'        Programado por: Barrantes D.
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_NroFolios As Long
Dim ml_IdServicioDestino As Long
Dim ml_IdServicioOrigen As Long
Dim ms_Observacion As String
Dim ml_IdMotivo As Long
Dim mda_FechaMovimiento As Date
Dim ml_IdPaciente As Long
Dim ml_IdMovimiento As Long
Dim ml_IdEmpleadoRecepcion As Long
Dim ml_IdEmpleadoTransporte As Long
Dim ml_IdEmpleadoArchivo As Long
Dim ml_IdGrupoMovimiento As Long
Dim ml_IdHistoriaSolicitada  As Long
'*********debb*******************
Dim ml_IdAtencion As Long
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property


Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let NroFolios(lValue As Long)
   ml_NroFolios = lValue
End Property
Property Get NroFolios() As Long
   NroFolios = ml_NroFolios
End Property
Property Let IdServicioDestino(lValue As Long)
   ml_IdServicioDestino = lValue
End Property
Property Get IdServicioDestino() As Long
   IdServicioDestino = ml_IdServicioDestino
End Property
Property Let IdServicioOrigen(lValue As Long)
   ml_IdServicioOrigen = lValue
End Property
Property Get IdServicioOrigen() As Long
   IdServicioOrigen = ml_IdServicioOrigen
End Property
Property Let Observacion(sValue As String)
   ms_Observacion = sValue
End Property
Property Get Observacion() As String
   Observacion = ms_Observacion
End Property
Property Let IdMotivo(lValue As Long)
   ml_IdMotivo = lValue
End Property
Property Get IdMotivo() As Long
   IdMotivo = ml_IdMotivo
End Property
Property Let FechaMovimiento(daValue As Date)
   mda_FechaMovimiento = daValue
End Property
Property Get FechaMovimiento() As Date
   FechaMovimiento = mda_FechaMovimiento
End Property
Property Let IdPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let IdMovimiento(lValue As Long)
   ml_IdMovimiento = lValue
End Property
Property Get IdMovimiento() As Long
   IdMovimiento = ml_IdMovimiento
End Property
Property Let IdEmpleadoRecepcion(lValue As Long)
   ml_IdEmpleadoRecepcion = lValue
End Property
Property Get IdEmpleadoRecepcion() As Long
   IdEmpleadoRecepcion = ml_IdEmpleadoRecepcion
End Property
Property Let IdEmpleadoTransporte(lValue As Long)
   ml_IdEmpleadoTransporte = lValue
End Property
Property Get IdEmpleadoTransporte() As Long
   IdEmpleadoTransporte = ml_IdEmpleadoTransporte
End Property
Property Let IdEmpleadoArchivo(lValue As Long)
   ml_IdEmpleadoArchivo = lValue
End Property
Property Get IdEmpleadoArchivo() As Long
   IdEmpleadoArchivo = ml_IdEmpleadoArchivo
End Property
Property Let IdGrupoMovimiento(lValue As Long)
   ml_IdGrupoMovimiento = lValue
End Property
Property Get IdGrupoMovimiento() As Long
   IdGrupoMovimiento = ml_IdGrupoMovimiento
End Property

Property Let IdHistoriaSolicitada(lValue As Long)
   ml_IdHistoriaSolicitada = lValue
End Property
Property Get IdHistoriaSolicitada() As Long
   IdHistoriaSolicitada = ml_IdHistoriaSolicitada
End Property

