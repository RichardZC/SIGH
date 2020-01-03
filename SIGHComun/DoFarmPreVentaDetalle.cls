VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFarmPreVentaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla farmPreVentaDetalle
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idPreventa As Long
Dim ml_IdProducto As Long
Dim ml_item As Long
Dim ml_Cantidad As Long
Dim mc_Precio As Currency
Dim mc_Importe As Currency

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idPreventa(iValue As Long)
   ml_idPreventa = iValue
End Property
Property Get idPreventa() As Long
   idPreventa = ml_idPreventa
End Property
Property Let IdProducto(iValue As Long)
   ml_IdProducto = iValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property

Property Let Item(iValue As Long)
   ml_item = iValue
End Property
Property Get Item() As Long
   Item = ml_item
End Property
Property Let Cantidad(iValue As Long)
   ml_Cantidad = iValue
End Property
Property Get Cantidad() As Long
   Cantidad = ml_Cantidad
End Property
Property Let Precio(cValue As Currency)
   mc_Precio = cValue
End Property
Property Get Precio() As Currency
   Precio = mc_Precio
End Property
Property Let Importe(cValue As Currency)
   mc_Importe = cValue
End Property
Property Get Importe() As Currency
   Importe = mc_Importe
End Property

