VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOEspecialidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Especialidades
'        Programado por: Castro W
'        Fecha: Abril 2006
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_Auditoria As Long
Dim ml_IdEspecialidad As Long
Dim ms_Nombre As String
Dim ml_IdDepartamento As Long
Dim ml_TiempoPromedioConsulta As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let TiempoPromedioConsulta(lValue As Long)
   ml_TiempoPromedioConsulta = lValue
End Property
Property Get TiempoPromedioConsulta() As Long
   TiempoPromedioConsulta = ml_TiempoPromedioConsulta
End Property
Property Let IdDepartamento(lValue As Long)
   ml_IdDepartamento = lValue
End Property
Property Get IdDepartamento() As Long
   IdDepartamento = ml_IdDepartamento
End Property
Property Let Nombre(sValue As String)
   ms_Nombre = sValue
End Property
Property Get Nombre() As String
   Nombre = ms_Nombre
End Property
Property Let IdEspecialidad(lValue As Long)
   ml_IdEspecialidad = lValue
End Property
Property Get IdEspecialidad() As Long
   IdEspecialidad = ml_IdEspecialidad
End Property


