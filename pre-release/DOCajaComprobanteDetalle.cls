VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCajaComprobantesDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'MZD Ini 02/06/2005 [Todo el archivo]

Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOCajaComprobantesDetalleBienesInsumos
'        Autor: William Castro Grijalva
'        Fecha: 03/06/2005 10:24:57
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim mc_SubTotalPagado As Currency
Dim ml_Cantidad As Currency
Dim mc_PrecioUnitario As Currency
Dim ml_IdComprobanteDetalle As Long
Dim ml_IdComprobantePago As Long
Dim ml_IdProducto As Long
Dim mi_TipoDetalle   As sghTipoDetalleComprobante
Dim ms_NombreProducto As String
Dim ms_CodigoProducto As String
Dim ml_IdFacturacionDetalle As Long
Dim ml_IdEstadoFacturacion  As Long
Dim mc_SubTotalExonerado As Currency
Dim mc_SubTotalPagadoACuenta As Currency
Dim ml_EsPagoACuenta As Long

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
Property Let Cantidad(lValue As Currency)
   ml_Cantidad = lValue
End Property
Property Get Cantidad() As Currency
   Cantidad = ml_Cantidad
End Property
Property Let PrecioUnitario(cValue As Currency)
   mc_PrecioUnitario = cValue
End Property
Property Get PrecioUnitario() As Currency
   PrecioUnitario = mc_PrecioUnitario
End Property
Property Let IdComprobanteDetalle(lValue As Long)
   ml_IdComprobanteDetalle = lValue
End Property
Property Get IdComprobanteDetalle() As Long
   IdComprobanteDetalle = ml_IdComprobanteDetalle
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
Property Let NombreProducto(sValue As String)
   ms_NombreProducto = sValue
End Property
Property Get NombreProducto() As String
   NombreProducto = ms_NombreProducto
End Property
Property Let TipoDetalle(iValue As sghTipoDetalleComprobante)
   mi_TipoDetalle = iValue
End Property
Property Get TipoDetalle() As sghTipoDetalleComprobante
   TipoDetalle = mi_TipoDetalle
End Property
Property Let CodigoProducto(sValue As String)
   ms_CodigoProducto = sValue
End Property
Property Get CodigoProducto() As String
   CodigoProducto = ms_CodigoProducto
End Property
Property Let IdFacturacionDetalle(lValue As Long)
   ml_IdFacturacionDetalle = lValue
End Property
Property Get IdFacturacionDetalle() As Long
   IdFacturacionDetalle = ml_IdFacturacionDetalle
End Property
Property Let IdEstadoFacturacion(lValue As Long)
   ml_IdEstadoFacturacion = lValue
End Property
Property Get IdEstadoFacturacion() As Long
   IdEstadoFacturacion = ml_IdEstadoFacturacion
End Property
Property Let SubTotalExonerado(cValue As Currency)
   mc_SubTotalExonerado = cValue
End Property
Property Get SubTotalExonerado() As Currency
   SubTotalExonerado = mc_SubTotalExonerado
End Property
Property Let SubTotalPagadoACuenta(cValue As Currency)
   mc_SubTotalPagadoACuenta = cValue
End Property
Property Get SubTotalPagadoACuenta() As Currency
   SubTotalPagadoACuenta = mc_SubTotalPagadoACuenta
End Property
Property Let EsPagoACuenta(lValue As Long)
   ml_EsPagoACuenta = lValue
End Property
Property Get EsPagoACuenta() As Long
   EsPagoACuenta = ml_EsPagoACuenta
End Property

