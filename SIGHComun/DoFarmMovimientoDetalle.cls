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
Dim ms_LOTE As String
Dim mda_FechaVencimiento As Date
Dim ml_item As Long
Dim ml_cantidad As Long
Dim mc_precio As Currency
Dim mc_Total As Currency
Dim ms_RegistroSanitario As String
Dim ml_idTipoSalidaBienInsumo As Long
Dim ms_DocumentoNumero As String

'-----------varibles servicios
Dim cs_Id As Long
Dim cs_Id_Atencion As String
Dim cs_Descripcion_Tipo_Item As String
Dim cs_Fg_Tipo As String
Dim cs_Codigo As String
Dim cs_NombreProducto As String
Dim cs_IdSubClasificacion As String
Dim cs_LabConfHis As String

Property Let id(sValue As Long)
   cs_Id = sValue
End Property
Property Get id() As Long
   id = cs_Id
End Property

Property Let Id_Atencion(sValue As String)
   cs_Id_Atencion = sValue
End Property
Property Get Id_Atencion() As String
   Id_Atencion = cs_Id_Atencion
End Property

Property Let Descripcion_Tipo_Item(sValue As String)
   cs_Descripcion_Tipo_Item = sValue
End Property
Property Get Descripcion_Tipo_Item() As String
   Descripcion_Tipo_Item = cs_Descripcion_Tipo_Item
End Property

Property Let Fg_Tipo(sValue As String)
   cs_Fg_Tipo = sValue
End Property
Property Get Fg_Tipo() As String
   Fg_Tipo = cs_Fg_Tipo
End Property

Property Let Codigo(sValue As String)
   cs_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = cs_Codigo
End Property

Property Let NombreProducto(sValue As String)
   cs_NombreProducto = sValue
End Property
Property Get NombreProducto() As String
   NombreProducto = cs_NombreProducto
End Property

Property Let IdSubClasificacion(sValue As String)
   cs_IdSubClasificacion = sValue
End Property
Property Get IdSubClasificacion() As String
   IdSubClasificacion = cs_IdSubClasificacion
End Property

Property Let labConfHIS(sValue As String)
   cs_LabConfHis = sValue
End Property
Property Get labConfHIS() As String
   labConfHIS = cs_LabConfHis
End Property
'--------------------------------

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
Property Let LOTE(sValue As String)
   ms_LOTE = sValue
End Property
Property Get LOTE() As String
   LOTE = ms_LOTE
End Property
Property Let FechaVencimiento(daValue As Date)
   mda_FechaVencimiento = daValue
End Property
Property Get FechaVencimiento() As Date
   FechaVencimiento = mda_FechaVencimiento
End Property
Property Let item(iValue As Long)
   ml_item = iValue
End Property
Property Get item() As Long
   item = ml_item
End Property
Property Let cantidad(iValue As Long)
   ml_cantidad = iValue
End Property
Property Get cantidad() As Long
   cantidad = ml_cantidad
End Property
Property Let precio(cValue As Currency)
   mc_precio = cValue
End Property
Property Get precio() As Currency
   precio = mc_precio
End Property
Property Let TOTAL(cValue As Currency)
   mc_Total = cValue
End Property
Property Get TOTAL() As Currency
   TOTAL = mc_Total
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

