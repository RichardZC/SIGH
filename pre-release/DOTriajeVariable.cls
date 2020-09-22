VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOTriajeVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla TriajeVariable
'        Programado por: Garay M
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdTriajeVariable As Long
Dim ms_TriajeVariable As String
Dim mb_EsAntropometrica As Boolean
Dim mb_TieneLimiteMedicion As Boolean
Dim ml_EdadDiaLimiteMinima As Long
Dim ml_EdadDiaLimiteMaxima As Long
Dim mb_EsDatoObligatorio As Boolean
Dim mb_EsActivo As Boolean

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdTriajeVariable(iValue As Long)
   ml_IdTriajeVariable = iValue
End Property
Property Get IdTriajeVariable() As Long
   IdTriajeVariable = ml_IdTriajeVariable
End Property
Property Let TriajeVariable(sValue As String)
   ms_TriajeVariable = sValue
End Property
Property Get TriajeVariable() As String
   TriajeVariable = ms_TriajeVariable
End Property
Property Let EsAntropometrica(bValue As Boolean)
   mb_EsAntropometrica = bValue
End Property
Property Get EsAntropometrica() As Boolean
   EsAntropometrica = mb_EsAntropometrica
End Property
Property Let TieneLimiteMedicion(bValue As Boolean)
   mb_TieneLimiteMedicion = bValue
End Property
Property Get TieneLimiteMedicion() As Boolean
   TieneLimiteMedicion = mb_TieneLimiteMedicion
End Property
Property Let EdadDiaLimiteMinima(iValue As Long)
   ml_EdadDiaLimiteMinima = iValue
End Property
Property Get EdadDiaLimiteMinima() As Long
   EdadDiaLimiteMinima = ml_EdadDiaLimiteMinima
End Property
Property Let EdadDiaLimiteMaxima(iValue As Long)
   ml_EdadDiaLimiteMaxima = iValue
End Property
Property Get EdadDiaLimiteMaxima() As Long
   EdadDiaLimiteMaxima = ml_EdadDiaLimiteMaxima
End Property
Property Let EsDatoObligatorio(bValue As Boolean)
   mb_EsDatoObligatorio = bValue
End Property
Property Get EsDatoObligatorio() As Boolean
   EsDatoObligatorio = mb_EsDatoObligatorio
End Property
Property Let EsActivo(bValue As Boolean)
   mb_EsActivo = bValue
End Property
Property Get EsActivo() As Boolean
   EsActivo = mb_EsActivo
End Property
