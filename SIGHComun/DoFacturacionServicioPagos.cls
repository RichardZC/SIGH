VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFacturacionServicioPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FacturacionServicioPagos
'        Programado por: Barrantes D
'        Fecha: Abril 2009
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ml_idOrdenPago As Long
Dim ml_IdProducto As Long
Dim ml_Cantidad As Long
Dim mv_Precio As Variant
Dim mv_Total As Variant

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let idOrdenPago(iValue As Long)
   ml_idOrdenPago = iValue
End Property
Property Get idOrdenPago() As Long
   idOrdenPago = ml_idOrdenPago
End Property
Property Let IdProducto(iValue As Long)
   ml_IdProducto = iValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let Cantidad(iValue As Long)
   ml_Cantidad = iValue
End Property
Property Get Cantidad() As Long
   Cantidad = ml_Cantidad
End Property
Property Let Precio(vValue As Variant)
   mv_Precio = vValue
End Property
Property Get Precio() As Variant
   Precio = mv_Precio
End Property
Property Let Total(vValue As Variant)
   mv_Total = vValue
End Property
Property Get Total() As Variant
   Total = mv_Total
End Property


