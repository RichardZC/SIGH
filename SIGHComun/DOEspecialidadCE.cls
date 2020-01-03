VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOEspecialidadCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla EspecialidadCE
'        Programado por: Castro W
'        Fecha: Abril 2006
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_Auditoria As Long
Dim ml_IdProductoInterconsulta As Long
Dim ml_IdProductoConsulta As Long
Dim ml_TiempoPromedioAtencion As Long
Dim ml_IdEspecialidad As Long
Dim ml_IdEspecialidadCE As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdProductoInterconsulta(lValue As Long)
   ml_IdProductoInterconsulta = lValue
End Property
Property Get IdProductoInterconsulta() As Long
   IdProductoInterconsulta = ml_IdProductoInterconsulta
End Property
Property Let IdProductoConsulta(lValue As Long)
   ml_IdProductoConsulta = lValue
End Property
Property Get IdProductoConsulta() As Long
   IdProductoConsulta = ml_IdProductoConsulta
End Property
Property Let TiempoPromedioAtencion(lValue As Long)
   ml_TiempoPromedioAtencion = lValue
End Property
Property Get TiempoPromedioAtencion() As Long
   TiempoPromedioAtencion = ml_TiempoPromedioAtencion
End Property
Property Let IdEspecialidad(lValue As Long)
   ml_IdEspecialidad = lValue
End Property
Property Get IdEspecialidad() As Long
   IdEspecialidad = ml_IdEspecialidad
End Property
Property Let IdEspecialidadCE(lValue As Long)
   ml_IdEspecialidadCE = lValue
End Property
Property Get IdEspecialidadCE() As Long
   IdEspecialidadCE = ml_IdEspecialidadCE
End Property


