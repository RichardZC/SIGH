VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCajaNroDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla CajaNroDocumento
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdTipoComprobante As Long
Dim ms_NroDocumento As String
Dim ms_NroSerie As String
Dim ms_NroDocumentoFinal As String
Dim ml_idCaja As Long
Dim ms_NroDocumentoInicial As String
Dim ml_FacturaSinIGV As Boolean
Property Let FacturaSinIGV(lValue As Boolean)
   ml_FacturaSinIGV = lValue
End Property
Property Get FacturaSinIGV() As Boolean
   FacturaSinIGV = ml_FacturaSinIGV
End Property


Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdTipoComprobante(lValue As Long)
   ml_IdTipoComprobante = lValue
End Property
Property Get IdTipoComprobante() As Long
   IdTipoComprobante = ml_IdTipoComprobante
End Property
Property Let NroDocumento(sValue As String)
   ms_NroDocumento = sValue
End Property
Property Get NroDocumento() As String
   NroDocumento = ms_NroDocumento
End Property
Property Let NroSerie(sValue As String)
   ms_NroSerie = sValue
End Property
Property Get NroSerie() As String
   NroSerie = ms_NroSerie
End Property
Property Let NroDocumentoFinal(sValue As String)
   ms_NroDocumentoFinal = sValue
End Property
Property Get NroDocumentoFinal() As String
   NroDocumentoFinal = ms_NroDocumentoFinal
End Property
Property Let idCaja(lValue As Long)
   ml_idCaja = lValue
End Property
Property Get idCaja() As Long
   idCaja = ml_idCaja
End Property
Property Let NroDocumentoInicial(sValue As String)
   ms_NroDocumentoInicial = sValue
End Property
Property Get NroDocumentoInicial() As String
   NroDocumentoInicial = ms_NroDocumentoInicial
End Property


