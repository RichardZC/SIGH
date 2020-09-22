VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFarmInventarioDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla farmInventarioDetalle
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_idInventario As Long
Dim ml_idProducto As Long
Dim ms_Lote As String
Dim mda_FechaVencimiento As Date
Dim ml_Cantidad As Long
Dim mc_Precio As Currency
Dim ms_RegistroSanitario As String
Dim ml_idTipoSalidaBienInsumo As Long
Dim ml_CantidadSaldo As Long
Dim ml_CantidadFaltante As Long
Dim ml_CantidadSobrante As Long
Dim ml_EsHistoricoSaldo As Long

Property Let EsHistoricoSaldo(iValue As Long)
   ml_EsHistoricoSaldo = iValue
End Property

Property Get EsHistoricoSaldo() As Long
   EsHistoricoSaldo = ml_EsHistoricoSaldo
End Property

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



Property Let idTipoSalidaBienInsumo(lValue As Long)
   ml_idTipoSalidaBienInsumo = lValue
End Property
Property Get idTipoSalidaBienInsumo() As Long
   idTipoSalidaBienInsumo = ml_idTipoSalidaBienInsumo
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
Property Let idProducto(iValue As Long)
   ml_idProducto = iValue
End Property
Property Get idProducto() As Long
   idProducto = ml_idProducto
End Property
Property Let Lote(sValue As String)
   ms_Lote = sValue
End Property
Property Get Lote() As String
   Lote = ms_Lote
End Property
Property Let FechaVencimiento(daValue As Date)
   mda_FechaVencimiento = daValue
End Property
Property Get FechaVencimiento() As Date
   FechaVencimiento = mda_FechaVencimiento
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
Property Let RegistroSanitario(sValue As String)
   ms_RegistroSanitario = sValue
End Property
Property Get RegistroSanitario() As String
   RegistroSanitario = ms_RegistroSanitario
End Property


