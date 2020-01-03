VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFactReembolsosDcto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactReembolsosDocumentos
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ml_idFactReembolso As Long
Dim ms_NroSerie As String
Dim ms_NroDocumento As String
Dim ml_idComprobantePago As Long
Dim ms_MotivoAnulacion As String

Property Let MotivoAnulacion(sValue As String)
   ms_MotivoAnulacion = sValue
End Property
Property Get MotivoAnulacion() As String
   MotivoAnulacion = ms_MotivoAnulacion
End Property
Property Let idComprobantePago(lValue As Long)
   ml_idComprobantePago = lValue
End Property
Property Get idComprobantePago() As Long
   idComprobantePago = ml_idComprobantePago
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let idFactReembolso(iValue As Long)
   ml_idFactReembolso = iValue
End Property
Property Get idFactReembolso() As Long
   idFactReembolso = ml_idFactReembolso
End Property
Property Let NroSerie(sValue As String)
   ms_NroSerie = sValue
End Property
Property Get NroSerie() As String
   NroSerie = ms_NroSerie
End Property
Property Let NroDocumento(sValue As String)
   ms_NroDocumento = sValue
End Property
Property Get NroDocumento() As String
   NroDocumento = ms_NroDocumento
End Property

