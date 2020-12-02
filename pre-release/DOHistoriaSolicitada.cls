VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOHistoriaSolicitada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla HistoriasSolicitadas
'        Programado por: Barrantes D
'        Fecha: Octubre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdMotivo As Long
Dim ms_HoraRequerida As String
Dim mda_FechaRequerida As Date
Dim ms_HoraSolicitud As String
Dim mda_FechaSolicitud As Date
Dim ml_IdPaciente As Long
Dim ml_IdHistoriaSolicitada As Long
Dim ml_IdEmpleadoSolicita As Long
Dim ml_IdMovimiento As Long
Dim ms_Observacion As String
Dim ml_IdServicio As Long
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
Property Let IdMotivo(lValue As Long)
   ml_IdMotivo = lValue
End Property
Property Get IdMotivo() As Long
   IdMotivo = ml_IdMotivo
End Property
Property Let HoraRequerida(sValue As String)
   ms_HoraRequerida = sValue
End Property
Property Get HoraRequerida() As String
   HoraRequerida = ms_HoraRequerida
End Property
Property Let FechaRequerida(daValue As Date)
   mda_FechaRequerida = daValue
End Property
Property Get FechaRequerida() As Date
   FechaRequerida = mda_FechaRequerida
End Property
Property Let HoraSolicitud(sValue As String)
   ms_HoraSolicitud = sValue
End Property
Property Get HoraSolicitud() As String
   HoraSolicitud = ms_HoraSolicitud
End Property
Property Let FechaSolicitud(daValue As Date)
   mda_FechaSolicitud = daValue
End Property
Property Get FechaSolicitud() As Date
   FechaSolicitud = mda_FechaSolicitud
End Property
Property Let IdPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let IdHistoriaSolicitada(lValue As Long)
   ml_IdHistoriaSolicitada = lValue
End Property
Property Get IdHistoriaSolicitada() As Long
   IdHistoriaSolicitada = ml_IdHistoriaSolicitada
End Property
Property Let IdEmpleadoSolicita(lValue As Long)
   ml_IdEmpleadoSolicita = lValue
End Property
Property Get IdEmpleadoSolicita() As Long
   IdEmpleadoSolicita = ml_IdEmpleadoSolicita
End Property
Property Let IdMovimiento(lValue As Long)
   ml_IdMovimiento = lValue
End Property
Property Get IdMovimiento() As Long
   IdMovimiento = ml_IdMovimiento
End Property
Property Let Observacion(sValue As String)
   ms_Observacion = sValue
End Property
Property Get Observacion() As String
   Observacion = ms_Observacion
End Property
Property Let IdServicio(lValue As Long)
   ml_IdServicio = lValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property


