VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCajaFormaPagoComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla CajaFormaPagoComprobante
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_Auditoria As Long
Dim mc_TipoCambio As Currency
Dim mc_TotalSoles As Currency
Dim mc_Importe As Currency
Dim ml_idComprobantePago As Long
Dim ml_IdTipoMoneda As Long
Dim ml_IdTipoFormaPago As Long
Dim ml_IdFormaPago As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let TipoCambio(cValue As Currency)
   mc_TipoCambio = cValue
End Property
Property Get TipoCambio() As Currency
   TipoCambio = mc_TipoCambio
End Property
Property Let TotalSoles(cValue As Currency)
   mc_TotalSoles = cValue
End Property
Property Get TotalSoles() As Currency
   TotalSoles = mc_TotalSoles
End Property
Property Let Importe(cValue As Currency)
   mc_Importe = cValue
End Property
Property Get Importe() As Currency
   Importe = mc_Importe
End Property
Property Let idComprobantePago(lValue As Long)
   ml_idComprobantePago = lValue
End Property
Property Get idComprobantePago() As Long
   idComprobantePago = ml_idComprobantePago
End Property
Property Let IdTipoMoneda(lValue As Long)
   ml_IdTipoMoneda = lValue
End Property
Property Get IdTipoMoneda() As Long
   IdTipoMoneda = ml_IdTipoMoneda
End Property
Property Let IdTipoFormaPago(lValue As Long)
   ml_IdTipoFormaPago = lValue
End Property
Property Get IdTipoFormaPago() As Long
   IdTipoFormaPago = ml_IdTipoFormaPago
End Property
Property Let IdFormaPago(lValue As Long)
   ml_IdFormaPago = lValue
End Property
Property Get IdFormaPago() As Long
   IdFormaPago = ml_IdFormaPago
End Property

