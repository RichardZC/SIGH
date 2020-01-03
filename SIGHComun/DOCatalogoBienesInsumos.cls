VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCatalogoBienesInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactCatalogoBienesInsumos
'        Programado por: Barrantes D
'        Fecha: Enero 2010
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdCentroCosto As Long
Dim ml_IdPartida As Long
Dim ml_IdSubGrupoFarmacologico As Long
Dim ml_IdGrupoFarmacologico As Long
Dim ml_IdClasificacionBienInsumo As Long
Dim ms_NombreComercial As String
Dim ms_nombre As String
Dim ms_Codigo As String
Dim ml_idProducto As Long
Dim ml_idTipoSalidaBienInsumo As Long
Dim mc_PrecioCompra As Currency
Dim mc_PrecioDistribucion As Currency
Dim mc_PrecioDonacion As Currency
Dim mc_PrecioUltCompra As Currency
Dim ml_StockMinimo As Long
Dim ml_TipoProducto As Long
Dim ms_denominacion As String
Dim ms_Concentracion As String
Dim ms_Presentacion As String
Dim ms_FormaFarmaceutica As String
Dim ms_MaterialEnvase As String
Dim ms_PresentacionEnvase As String
Dim ms_Fabricante As String
Dim ml_IdPaisOrigen As Long
Dim mb_Petitorio As Boolean
Dim ms_TipoProductoSismed As String
Dim ms_codigoSUNAT As String

Property Let codigoSUNAT(sValue As String)
   ms_codigoSUNAT = sValue
End Property
Property Get codigoSUNAT() As String
   codigoSUNAT = ms_codigoSUNAT
End Property


Property Let TipoProductoSismed(sValue As String)
   ms_TipoProductoSismed = sValue
End Property
Property Get TipoProductoSismed() As String
   TipoProductoSismed = ms_TipoProductoSismed
End Property


Property Let Petitorio(lValue As Long)
   mb_Petitorio = lValue
End Property
Property Get Petitorio() As Long
   Petitorio = mb_Petitorio
End Property


Property Let IdPaisOrigen(lValue As Long)
   ml_IdPaisOrigen = lValue
End Property
Property Get IdPaisOrigen() As Long
   IdPaisOrigen = ml_IdPaisOrigen
End Property

Property Let Fabricante(sValue As String)
   ms_Fabricante = sValue
End Property
Property Get Fabricante() As String
   Fabricante = ms_Fabricante
End Property
Property Let PresentacionEnvase(sValue As String)
   ms_PresentacionEnvase = sValue
End Property
Property Get PresentacionEnvase() As String
   PresentacionEnvase = ms_PresentacionEnvase
End Property

Property Let MaterialEnvase(sValue As String)
   ms_MaterialEnvase = sValue
End Property
Property Get MaterialEnvase() As String
   MaterialEnvase = ms_MaterialEnvase
End Property

Property Let FormaFarmaceutica(sValue As String)
   ms_FormaFarmaceutica = sValue
End Property
Property Get FormaFarmaceutica() As String
   FormaFarmaceutica = ms_FormaFarmaceutica
End Property

Property Let Presentacion(sValue As String)
   ms_Presentacion = sValue
End Property
Property Get Presentacion() As String
   Presentacion = ms_Presentacion
End Property

Property Let Concentracion(sValue As String)
   ms_Concentracion = sValue
End Property
Property Get Concentracion() As String
   Concentracion = ms_Concentracion
End Property

Property Let denominacion(sValue As String)
   ms_denominacion = sValue
End Property
Property Get denominacion() As String
   denominacion = ms_denominacion
End Property




Property Let TipoProducto(lValue As Long)
   ml_TipoProducto = lValue
End Property
Property Get TipoProducto() As Long
   TipoProducto = ml_TipoProducto
End Property
Property Let StockMinimo(lValue As Long)
   ml_StockMinimo = lValue
End Property
Property Get StockMinimo() As Long
   StockMinimo = ml_StockMinimo
End Property
Property Let PrecioUltCompra(cValue As Currency)
   mc_PrecioUltCompra = cValue
End Property
Property Get PrecioUltCompra() As Currency
   PrecioUltCompra = mc_PrecioUltCompra
End Property

Property Let PrecioDonacion(cValue As Currency)
   mc_PrecioDonacion = cValue
End Property
Property Get PrecioDonacion() As Currency
   PrecioDonacion = mc_PrecioDonacion
End Property
Property Let PrecioDistribucion(cValue As Currency)
   mc_PrecioDistribucion = cValue
End Property
Property Get PrecioDistribucion() As Currency
   PrecioDistribucion = mc_PrecioDistribucion
End Property
Property Let PrecioCompra(cValue As Currency)
   mc_PrecioCompra = cValue
End Property
Property Get PrecioCompra() As Currency
   PrecioCompra = mc_PrecioCompra
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
Property Let IdCentroCosto(lValue As Long)
   ml_IdCentroCosto = lValue
End Property
Property Get IdCentroCosto() As Long
   IdCentroCosto = ml_IdCentroCosto
End Property
Property Let IdPartida(lValue As Long)
   ml_IdPartida = lValue
End Property
Property Get IdPartida() As Long
   IdPartida = ml_IdPartida
End Property
Property Let IdSubGrupoFarmacologico(lValue As Long)
   ml_IdSubGrupoFarmacologico = lValue
End Property
Property Get IdSubGrupoFarmacologico() As Long
   IdSubGrupoFarmacologico = ml_IdSubGrupoFarmacologico
End Property
Property Let IdGrupoFarmacologico(lValue As Long)
   ml_IdGrupoFarmacologico = lValue
End Property
Property Get IdGrupoFarmacologico() As Long
   IdGrupoFarmacologico = ml_IdGrupoFarmacologico
End Property
Property Let IdClasificacionBienInsumo(lValue As Long)
   ml_IdClasificacionBienInsumo = lValue
End Property
Property Get IdClasificacionBienInsumo() As Long
   IdClasificacionBienInsumo = ml_IdClasificacionBienInsumo
End Property
Property Let NombreComercial(sValue As String)
   ms_NombreComercial = sValue
End Property
Property Get NombreComercial() As String
   NombreComercial = ms_NombreComercial
End Property
Property Let nombre(sValue As String)
   ms_nombre = sValue
End Property
Property Get nombre() As String
   nombre = ms_nombre
End Property
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property
Property Let idProducto(lValue As Long)
   ml_idProducto = lValue
End Property
Property Get idProducto() As Long
   idProducto = ml_idProducto
End Property


