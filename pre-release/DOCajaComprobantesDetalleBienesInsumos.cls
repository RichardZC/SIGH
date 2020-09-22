VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCajaCompDetalleInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'MZD-05 [Todo el archivo]

Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOCajaComprobantesDetalleBienesInsumos
'        Autor: William Castro Grijalva
'        Fecha: 10/08/2005 07:43:00
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim mc_SubTotalPagado As Currency
Dim ml_Cantidad As Long
Dim mc_PrecioUnitario As Currency
Dim ml_IdComprobanteDetalleBienes As Long
Dim ml_IdFacturacionBienes As Long
Dim ml_EsPagoACuenta As Long
Dim ml_IdComprobantePago As Long
Dim ml_IdProducto As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let SubTotalPagado(cValue As Currency)
   mc_SubTotalPagado = cValue
End Property
Property Get SubTotalPagado() As Currency
   SubTotalPagado = mc_SubTotalPagado
End Property
Property Let Cantidad(lValue As Long)
   ml_Cantidad = lValue
End Property
Property Get Cantidad() As Long
   Cantidad = ml_Cantidad
End Property
Property Let PrecioUnitario(cValue As Currency)
   mc_PrecioUnitario = cValue
End Property
Property Get PrecioUnitario() As Currency
   PrecioUnitario = mc_PrecioUnitario
End Property
Property Let IdComprobanteDetalleBienes(lValue As Long)
   ml_IdComprobanteDetalleBienes = lValue
End Property
Property Get IdComprobanteDetalleBienes() As Long
   IdComprobanteDetalleBienes = ml_IdComprobanteDetalleBienes
End Property
Property Let IdFacturacionBienes(lValue As Long)
   ml_IdFacturacionBienes = lValue
End Property
Property Get IdFacturacionBienes() As Long
   IdFacturacionBienes = ml_IdFacturacionBienes
End Property
Property Let EsPagoACuenta(lValue As Long)
   ml_EsPagoACuenta = lValue
End Property
Property Get EsPagoACuenta() As Long
   EsPagoACuenta = ml_EsPagoACuenta
End Property
Property Let IdComprobantePago(lValue As Long)
   ml_IdComprobantePago = lValue
End Property
Property Get IdComprobantePago() As Long
   IdComprobantePago = ml_IdComprobantePago
End Property
Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property


