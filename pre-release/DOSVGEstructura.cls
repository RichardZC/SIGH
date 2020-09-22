VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOSVGEstructura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla SVGEstructura
'        Programado por: Castro W
'        Fecha: Julio 2004
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ms_Codigo As String
Dim ms_SVG As String
Dim ms_Tipo As String
Dim ml_Orden As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property
Property Let SVG(sValue As String)
   ms_SVG = sValue
End Property
Property Get SVG() As String
   SVG = ms_SVG
End Property
Property Let Tipo(sValue As String)
   ms_Tipo = sValue
End Property
Property Get Tipo() As String
   Tipo = ms_Tipo
End Property
Property Let Orden(lValue As Long)
   ml_Orden = lValue
End Property
Property Get Orden() As Long
   Orden = ml_Orden
End Property


