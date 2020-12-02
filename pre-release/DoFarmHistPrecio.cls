VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFarmHistPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para capa de estructura de la tabla farmHistPrecio
'        Programado por: Barrantes D
'        Fecha: Noviembre 2014
'
'------------------------------------------------------------------------------------
Dim ml_IdUsuarioAuditoria As Long
Dim ml_idHistPrecio As Long
Dim ml_idProducto As Long
Dim mda_fecha As Date
Dim mc_PrecioCompra As Currency
Dim mc_PrecioDistribucion As Currency
Dim mc_PrecioVenta As Currency
Dim mc_PrecioDonacion As Currency
Dim ml_idUsuario As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idHistPrecio(iValue As Long)
   ml_idHistPrecio = iValue
End Property
Property Get idHistPrecio() As Long
   idHistPrecio = ml_idHistPrecio
End Property
Property Let idProducto(iValue As Long)
   ml_idProducto = iValue
End Property
Property Get idProducto() As Long
   idProducto = ml_idProducto
End Property
Property Let fecha(daValue As Date)
   mda_fecha = daValue
End Property
Property Get fecha() As Date
   fecha = mda_fecha
End Property
Property Let PrecioCompra(cValue As Currency)
   mc_PrecioCompra = cValue
End Property
Property Get PrecioCompra() As Currency
   PrecioCompra = mc_PrecioCompra
End Property
Property Let PrecioDistribucion(cValue As Currency)
   mc_PrecioDistribucion = cValue
End Property
Property Get PrecioDistribucion() As Currency
   PrecioDistribucion = mc_PrecioDistribucion
End Property
Property Let PrecioVenta(cValue As Currency)
   mc_PrecioVenta = cValue
End Property
Property Get PrecioVenta() As Currency
   PrecioVenta = mc_PrecioVenta
End Property
Property Let PrecioDonacion(cValue As Currency)
   mc_PrecioDonacion = cValue
End Property
Property Get PrecioDonacion() As Currency
   PrecioDonacion = mc_PrecioDonacion
End Property
Property Let idUsuario(iValue As Long)
   ml_idUsuario = iValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property

