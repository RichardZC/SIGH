VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoImagMovimientoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla ImagMovimientoDetalle
'        Programado por: Barrantes D
'        Fecha: Abril 2009
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdMovimiento As Long
Dim ml_IdProducto As Long
Dim ml_Cantidad As Long
Dim ml_CantidadFallada As Long
Dim ml_idProductoCpt As Long
Property Let idProductoCpt(iValue As Long)
   ml_idProductoCpt = iValue
End Property
Property Get idProductoCpt() As Long
   idProductoCpt = ml_idProductoCpt
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdMovimiento(iValue As Long)
   ml_IdMovimiento = iValue
End Property
Property Get IdMovimiento() As Long
   IdMovimiento = ml_IdMovimiento
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
Property Let CantidadFallada(iValue As Long)
   ml_CantidadFallada = iValue
End Property
Property Get CantidadFallada() As Long
   CantidadFallada = ml_CantidadFallada
End Property
