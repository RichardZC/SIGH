VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoSisFua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para la estructura de la tabla SISFUA
'        Programado por: Barrantes D
'        Fecha: Enero 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ms_FuaDisa As String
Dim ms_FuaLote As String
Dim ms_FuaNumeroInicial As String
Dim ms_FuaNumeroFinal As String
Dim ms_FuaUltimoGenerado As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let FuaDisa(sValue As String)
   ms_FuaDisa = sValue
End Property
Property Get FuaDisa() As String
   FuaDisa = ms_FuaDisa
End Property
Property Let FuaLote(sValue As String)
   ms_FuaLote = sValue
End Property
Property Get FuaLote() As String
   FuaLote = ms_FuaLote
End Property
Property Let FuaNumeroInicial(sValue As String)
   ms_FuaNumeroInicial = sValue
End Property
Property Get FuaNumeroInicial() As String
   FuaNumeroInicial = ms_FuaNumeroInicial
End Property
Property Let FuaNumeroFinal(sValue As String)
   ms_FuaNumeroFinal = sValue
End Property
Property Get FuaNumeroFinal() As String
   FuaNumeroFinal = ms_FuaNumeroFinal
End Property
Property Let FuaUltimoGenerado(sValue As String)
   ms_FuaUltimoGenerado = sValue
End Property
Property Get FuaUltimoGenerado() As String
   FuaUltimoGenerado = ms_FuaUltimoGenerado
End Property



