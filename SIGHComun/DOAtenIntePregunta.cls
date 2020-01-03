VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenIntePregunta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenIntePregunta
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdPregunta As Long
Dim ms_Pregunta As String
Dim ml_TipoRespuesta As Long
Dim ml_TipoValorRespuesta As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdPregunta(iValue As Long)
   ml_IdPregunta = iValue
End Property
Property Get IdPregunta() As Long
   IdPregunta = ml_IdPregunta
End Property
Property Let Pregunta(sValue As String)
   ms_Pregunta = sValue
End Property
Property Get Pregunta() As String
   Pregunta = ms_Pregunta
End Property
Property Let TipoRespuesta(iValue As Long)
   ml_TipoRespuesta = iValue
End Property
Property Get TipoRespuesta() As Long
   TipoRespuesta = ml_TipoRespuesta
End Property
Property Let TipoValorRespuesta(iValue As Long)
   ml_TipoValorRespuesta = iValue
End Property
Property Get TipoValorRespuesta() As Long
   TipoValorRespuesta = ml_TipoValorRespuesta
End Property
