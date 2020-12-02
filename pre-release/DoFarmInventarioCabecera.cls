VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFarmInventarioCabecera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla farmInventarioCabecera
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_Auditoria As Long
Dim ml_idInventario As Long
Dim ml_IdProducto As Long
Dim ml_Cantidad As Long
Dim mc_Precio As Currency
Dim mc_Total As Currency
Dim ml_CantidadSaldo As Long
Dim ml_CantidadFaltante As Long
Dim ml_CantidadSobrante As Long

Property Let CantidadSobrante(iValue As Long)
   ml_CantidadSobrante = iValue
End Property
Property Get CantidadSobrante() As Long
   CantidadSobrante = ml_CantidadSobrante
End Property

Property Let CantidadFaltante(iValue As Long)
   ml_CantidadFaltante = iValue
End Property
Property Get CantidadFaltante() As Long
   CantidadFaltante = ml_CantidadFaltante
End Property


Property Let CantidadSaldo(iValue As Long)
   ml_CantidadSaldo = iValue
End Property
Property Get CantidadSaldo() As Long
   CantidadSaldo = ml_CantidadSaldo
End Property


Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property

Property Let idInventario(iValue As Long)
   ml_idInventario = iValue
End Property
Property Get idInventario() As Long
   idInventario = ml_idInventario
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
Property Let Precio(cValue As Currency)
   mc_Precio = cValue
End Property
Property Get Precio() As Currency
   Precio = mc_Precio
End Property
Property Let Total(cValue As Currency)
   mc_Total = cValue
End Property
Property Get Total() As Currency
   Total = mc_Total
End Property



