VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCajaCompDetalleServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'MZD-05 [Todo el archivo]

Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOCajaComprobantesDetalleServicios
'        Autor: William Castro Grijalva
'        Fecha: 10/08/2005 07:43:02
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdFacturacionServicio As Long
Dim ml_EsPagoACuenta As Long
Dim ml_IdProducto As Long
Dim mc_SubTotalPagado As Currency
Dim ml_Cantidad As Currency
Dim mc_PrecioUnitario As Currency
Dim ml_IdComprobantePago As Long
Dim ml_IdComprobanteDetalleServicio As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdFacturacionServicio(lValue As Long)
   ml_IdFacturacionServicio = lValue
End Property
Property Get IdFacturacionServicio() As Long
   IdFacturacionServicio = ml_IdFacturacionServicio
End Property
Property Let EsPagoACuenta(lValue As Long)
   ml_EsPagoACuenta = lValue
End Property
Property Get EsPagoACuenta() As Long
   EsPagoACuenta = ml_EsPagoACuenta
End Property
Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
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
Property Let IdComprobantePago(lValue As Long)
   ml_IdComprobantePago = lValue
End Property
Property Get IdComprobantePago() As Long
   IdComprobantePago = ml_IdComprobantePago
End Property
Property Let IdComprobanteDetalleServicio(lValue As Long)
   ml_IdComprobanteDetalleServicio = lValue
End Property
Property Get IdComprobanteDetalleServicio() As Long
   IdComprobanteDetalleServicio = ml_IdComprobanteDetalleServicio
End Property


