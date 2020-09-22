VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Camas
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_Y As Long
Dim ml_X As Long
Dim ml_IdPaciente As Long
Dim ml_IdServicioUbicacionActual As Long
Dim ms_Codigo As String
Dim ml_IdEstadoCama As Long
Dim ml_IdCondicionOcupacion As Long
Dim ml_IdTiposCama As Long
Dim ml_IdServicioPropietario As Long
Dim ml_IdCama As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Y(lValue As Long)
   ml_Y = lValue
End Property
Property Get Y() As Long
   Y = ml_Y
End Property
Property Let X(lValue As Long)
   ml_X = lValue
End Property
Property Get X() As Long
   X = ml_X
End Property
Property Let IdPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let IdServicioUbicacionActual(lValue As Long)
   ml_IdServicioUbicacionActual = lValue
End Property
Property Get IdServicioUbicacionActual() As Long
   IdServicioUbicacionActual = ml_IdServicioUbicacionActual
End Property
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property
Property Let IdEstadoCama(lValue As Long)
   ml_IdEstadoCama = lValue
End Property
Property Get IdEstadoCama() As Long
   IdEstadoCama = ml_IdEstadoCama
End Property
Property Let IdCondicionOcupacion(lValue As Long)
   ml_IdCondicionOcupacion = lValue
End Property
Property Get IdCondicionOcupacion() As Long
   IdCondicionOcupacion = ml_IdCondicionOcupacion
End Property
Property Let IdTiposCama(lValue As Long)
   ml_IdTiposCama = lValue
End Property
Property Get IdTiposCama() As Long
   IdTiposCama = ml_IdTiposCama
End Property
Property Let IdServicioPropietario(lValue As Long)
   ml_IdServicioPropietario = lValue
End Property
Property Get IdServicioPropietario() As Long
   IdServicioPropietario = ml_IdServicioPropietario
End Property
Property Let IdCama(lValue As Long)
   ml_IdCama = lValue
End Property
Property Get IdCama() As Long
   IdCama = ml_IdCama
End Property


