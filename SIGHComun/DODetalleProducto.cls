VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DODetalleProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DODetalleProductos
'        Autor: William Castro Grijalva
'        Fecha: 01/09/2004 11:51:35 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdMotivoNoAtencion As Long
Dim ml_IdEstadoProducto As Long
Dim ml_IdDocumentoDetalle As Long
Dim mc_PrecioUnitario As Currency
Dim mc_Cantidad As Currency
Dim mc_PrecioTotal As Currency
Dim ml_IdProducto As Long
Dim ml_IdDetalleProducto As Long
Dim ml_IdCuentaAtencion As Long
Dim mb_CubiertoPorSeguro As Boolean
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdMotivoNoAtencion(lValue As Long)
   ml_IdMotivoNoAtencion = lValue
End Property
Property Get IdMotivoNoAtencion() As Long
   IdMotivoNoAtencion = ml_IdMotivoNoAtencion
End Property
Property Let IdEstadoProducto(lValue As Long)
   ml_IdEstadoProducto = lValue
End Property
Property Get IdEstadoProducto() As Long
   IdEstadoProducto = ml_IdEstadoProducto
End Property
Property Let IdDocumentoDetalle(lValue As Long)
   ml_IdDocumentoDetalle = lValue
End Property
Property Get IdDocumentoDetalle() As Long
   IdDocumentoDetalle = ml_IdDocumentoDetalle
End Property
Property Let PrecioUnitario(cValue As Currency)
   mc_PrecioUnitario = cValue
End Property
Property Get PrecioUnitario() As Currency
   PrecioUnitario = mc_PrecioUnitario
End Property
Property Let Cantidad(cValue As Currency)
   mc_Cantidad = cValue
End Property
Property Get Cantidad() As Currency
   Cantidad = mc_Cantidad
End Property
Property Let PrecioTotal(cValue As Currency)
   mc_PrecioTotal = cValue
End Property
Property Get PrecioTotal() As Currency
   PrecioTotal = mc_PrecioTotal
End Property
Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let IdDetalleProducto(lValue As Long)
   ml_IdDetalleProducto = lValue
End Property
Property Get IdDetalleProducto() As Long
   IdDetalleProducto = ml_IdDetalleProducto
End Property
Property Let IdCuentaAtencion(lValue As Long)
   ml_IdCuentaAtencion = lValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let CubiertoPorSeguro(bValue As Boolean)
   mb_CubiertoPorSeguro = bValue
End Property
Property Get CubiertoPorSeguro() As Boolean
   CubiertoPorSeguro = mb_CubiertoPorSeguro
End Property


