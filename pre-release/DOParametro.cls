VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOPArametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Parametros
'        Programado por: Benavides M
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ms_Descripcion As String
Dim ms_Codigo As String
Dim ms_Tipo As String
Dim ml_IdParametro As Long
Dim md_ValorFloat As Double
Dim ml_ValorInt As Long
Dim ms_ValorTexto As String
Dim ms_Grupo As String
Property Let Grupo(sValue As String)
   ms_Grupo = sValue
End Property
Property Get Grupo() As String
   Grupo = ms_Grupo
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property
Property Let Tipo(sValue As String)
   ms_Tipo = sValue
End Property
Property Get Tipo() As String
   Tipo = ms_Tipo
End Property
Property Let IdParametro(lValue As Long)
   ml_IdParametro = lValue
End Property
Property Get IdParametro() As Long
   IdParametro = ml_IdParametro
End Property
Property Let ValorFloat(dValue As Double)
   md_ValorFloat = dValue
End Property
Property Get ValorFloat() As Double
   ValorFloat = md_ValorFloat
End Property
Property Let ValorInt(lValue As Long)
   ml_ValorInt = lValue
End Property
Property Get ValorInt() As Long
   ValorInt = ml_ValorInt
End Property
Property Let ValorTexto(sValue As String)
   ms_ValorTexto = sValue
End Property
Property Get ValorTexto() As String
   ValorTexto = ms_ValorTexto
End Property







