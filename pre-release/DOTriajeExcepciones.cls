VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOTriajeExcepciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla TriajeExcepciones
'        Programado por: Garay M
'        Fecha: Octubre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdTriajeExcepciones As Long
Dim ml_IdTriajeVariable As Long
Dim ml_EdadInicialEnDia As Long
Dim ml_EdadFinalEnDia As Long
Dim mb_EsDatoObligatorio As Boolean

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdTriajeExcepciones(iValue As Long)
   ml_IdTriajeExcepciones = iValue
End Property
Property Get IdTriajeExcepciones() As Long
   IdTriajeExcepciones = ml_IdTriajeExcepciones
End Property
Property Let IdTriajeVariable(iValue As Long)
   ml_IdTriajeVariable = iValue
End Property
Property Get IdTriajeVariable() As Long
   IdTriajeVariable = ml_IdTriajeVariable
End Property
Property Let EdadInicialEnDia(iValue As Long)
   ml_EdadInicialEnDia = iValue
End Property
Property Get EdadInicialEnDia() As Long
   EdadInicialEnDia = ml_EdadInicialEnDia
End Property
Property Let EdadFinalEnDia(iValue As Long)
   ml_EdadFinalEnDia = iValue
End Property
Property Get EdadFinalEnDia() As Long
   EdadFinalEnDia = ml_EdadFinalEnDia
End Property
Property Let EsDatoObligatorio(bValue As Boolean)
   mb_EsDatoObligatorio = bValue
End Property
Property Get EsDatoObligatorio() As Boolean
   EsDatoObligatorio = mb_EsDatoObligatorio
End Property
