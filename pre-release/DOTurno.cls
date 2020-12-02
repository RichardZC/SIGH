VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Turnos
'        Programado por: Castro W
'        Fecha: Julio 2004
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdEspecialidad As Long
Dim ms_codigo As String
Dim ml_IdTipoServicio As Long
Dim ms_HoraFin As String
Dim ms_HoraInicio As String
Dim ms_Descripcion As String
Dim ml_idTurno As Long
Dim ml_idTipoActividades As Long

Property Let idTipoActividades(lValue As Long)
   ml_idTipoActividades = lValue
End Property
Property Get idTipoActividades() As Long
   idTipoActividades = ml_idTipoActividades
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdEspecialidad(lValue As Long)
   ml_IdEspecialidad = lValue
End Property
Property Get IdEspecialidad() As Long
   IdEspecialidad = ml_IdEspecialidad
End Property
Property Let codigo(sValue As String)
   ms_codigo = sValue
End Property
Property Get codigo() As String
   codigo = ms_codigo
End Property
Property Let IdTipoServicio(lValue As Long)
   ml_IdTipoServicio = lValue
End Property
Property Get IdTipoServicio() As Long
   IdTipoServicio = ml_IdTipoServicio
End Property
Property Let HoraFin(sValue As String)
   ms_HoraFin = sValue
End Property
Property Get HoraFin() As String
   HoraFin = ms_HoraFin
End Property
Property Let HoraInicio(sValue As String)
   ms_HoraInicio = sValue
End Property
Property Get HoraInicio() As String
   HoraInicio = ms_HoraInicio
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let idTurno(lValue As Long)
   ml_idTurno = lValue
End Property
Property Get idTurno() As Long
   idTurno = ml_idTurno
End Property


