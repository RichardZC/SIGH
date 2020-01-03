VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoLabPruebas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla LabPruebas
'        Programado por: Bonilla A
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idPrueba As Long
Dim ms_CodigoPrueba As String
Dim ms_CodigoCPT As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idPrueba(iValue As Long)
   ml_idPrueba = iValue
End Property
Property Get idPrueba() As Long
   idPrueba = ml_idPrueba
End Property
Property Let CodigoPrueba(sValue As String)
   ms_CodigoPrueba = sValue
End Property
Property Get CodigoPrueba() As String
   CodigoPrueba = ms_CodigoPrueba
End Property
Property Let CodigoCPT(sValue As String)
   ms_CodigoCPT = sValue
End Property
Property Get CodigoCPT() As String
   CodigoCPT = ms_CodigoCPT
End Property



