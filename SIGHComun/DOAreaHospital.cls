VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAreaHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Areas del Hospital
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_Auditoria As Long
Dim ms_Nombre As String
Dim ml_IdArea As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Nombre(sValue As String)
   ms_Nombre = sValue
End Property
Property Get Nombre() As String
   Nombre = ms_Nombre
End Property
Property Let IdArea(lValue As Long)
   ml_IdArea = lValue
End Property
Property Get IdArea() As Long
   IdArea = ml_IdArea
End Property


