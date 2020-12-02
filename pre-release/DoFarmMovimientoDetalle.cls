VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFarmMovimientoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla farmMovimientoDetalle
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ms_movNumero As String
Dim ms_movTipo As String
Dim ml_idProducto As Long
Dim ms_Lote As String
Dim mda_FechaVencimiento As Date
Dim ml_item As Long
Dim ml_Cantidad As Long
Dim mc_Precio As Currency
Dim mc_Total As Currency
Dim ms_RegistroSanitario As String
Dim ml_idTipoSalidaBienInsumo As Long
Dim ms_DocumentoNumero As String

Property Let idTipoSalidaBienInsumo(lValue As Long)
   ml_idTipoSalidaBienInsumo = lValue
End Property
Property Get idTipoSalidaBienInsumo() As Long
   idTipoSalidaBienInsumo = ml_idTipoSalidaBienInsumo
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let movNumero(sValue As String)
   ms_movNumero = sValue
End Property
Property Get movNumero() As String
   movNumero = ms_movNumero
End Property
Property Let movTipo(sValue As String)
   ms_movTipo = sValue
End Property
Property Get movTipo() As String
   movTipo = ms_movTipo
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
Property Let Total(cValue As Currency)
   mc_Total = cValue
End Property
Property Get Total() As Currency
   Total = mc_Total
End Property
Property Let RegistroSanitario(sValue As String)
   ms_RegistroSanitario = sValue
End Property
Property Get RegistroSanitario() As String
   RegistroSanitario = ms_RegistroSanitario
End Property
Property Let DocumentoNumero(sValue As String)
   ms_DocumentoNumero = sValue
End Property
Property Get DocumentoNumero() As String
   DocumentoNumero = ms_DocumentoNumero
End Property

