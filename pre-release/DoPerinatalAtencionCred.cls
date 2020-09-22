VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoPerinatalAtencionCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla PerinatalAtencionCred
'        Programado por: Garay M
'        Fecha: Setiembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idPerinatalAtencion As Long
Dim ms_EdadEnAnios As String
Dim ml_CredNumero As Long
Dim ms_CredCheck As String
Dim ml_IdAtencion As Long
Property Let IdAtencion(iValue As Long)
   ml_IdAtencion = iValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idPerinatalAtencion(iValue As Long)
   ml_idPerinatalAtencion = iValue
End Property
Property Get idPerinatalAtencion() As Long
   idPerinatalAtencion = ml_idPerinatalAtencion
End Property
Property Let EdadEnAnios(sValue As String)
   ms_EdadEnAnios = sValue
End Property
Property Get EdadEnAnios() As String
   EdadEnAnios = ms_EdadEnAnios
End Property
Property Let CredNumero(iValue As Long)
   ml_CredNumero = iValue
End Property
Property Get CredNumero() As Long
   CredNumero = ml_CredNumero
End Property
Property Let CredCheck(sValue As String)
   ms_CredCheck = sValue
End Property
Property Get CredCheck() As String
   CredCheck = ms_CredCheck
End Property


