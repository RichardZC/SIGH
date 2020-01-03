VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFacturacionBienesPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FacturacionBienesPagos
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idOrden As Long
Dim ml_IdProducto As Long
Dim ml_CantidadPagar As Long
Dim mc_PrecioVenta As Currency
Dim mc_TotalPagar As Currency

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let idOrden(iValue As Long)
   ml_idOrden = iValue
End Property
Property Get idOrden() As Long
   idOrden = ml_idOrden
End Property
Property Let IdProducto(iValue As Long)
   ml_IdProducto = iValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let CantidadPagar(iValue As Long)
   ml_CantidadPagar = iValue
End Property
Property Get CantidadPagar() As Long
   CantidadPagar = ml_CantidadPagar
End Property
Property Let PrecioVenta(cValue As Currency)
   mc_PrecioVenta = cValue
End Property
Property Get PrecioVenta() As Currency
   PrecioVenta = mc_PrecioVenta
End Property
Property Let TotalPagar(cValue As Currency)
   mc_TotalPagar = cValue
End Property
Property Get TotalPagar() As Currency
   TotalPagar = mc_TotalPagar
End Property

