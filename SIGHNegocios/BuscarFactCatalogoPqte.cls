VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscarFactCatalogoPqte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para buscar Paquete
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim oFormulario As New FacCatalogoPqteBuscar
Dim ml_IdRegistroSeleccionado As Long
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim mi_idPuntoCarga As Long
Dim mi_Descripcion As String
Dim mi_Codigo As String
Dim mi_Precio As String
Dim mi_Producto As String
Dim mi_idProducto  As Long
Dim mi_tipoProducto As Long
Dim oRsItemsMasivosElegidos As Recordset
Dim lnIdFarmaciaElegida As Long
Dim lcFarmaciaElegida As String
Dim ml_IdTipoFinanciamiento As Long
Dim ml_RegistraTodosLosItems As Boolean
Dim oRsDevuelveTodosLosItemsServ As Recordset
Dim oRsDevuelveTodosLosItemsFarm As Recordset

Property Let RegistraTodosLosItems(lValue As Boolean)
   ml_RegistraTodosLosItems = lValue
End Property


Property Let FarmaciaElegida(iValue As String)
  lcFarmaciaElegida = iValue
End Property
Property Let IdFarmaciaElegida(iValue As Long)
  lnIdFarmaciaElegida = iValue
End Property
Property Let idProducto(iValue As Long)
  mi_idProducto = iValue
End Property

Property Get idProducto() As Long
  idProducto = mi_idProducto
End Property
Property Get TipoProducto() As Long
  TipoProducto = mi_tipoProducto
End Property

Property Get Descripcion() As String
  Descripcion = mi_Descripcion
End Property
Property Get Codigo() As String
  Codigo = mi_Codigo
End Property
Property Get Precio() As String
  Precio = mi_Precio
End Property


Property Get Producto() As String
  Producto = mi_Producto
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property
Property Let idPuntoCarga(iValue As Long)
  mi_idPuntoCarga = iValue
End Property
Property Get idPuntoCarga() As Long
    idPuntoCarga = mi_idPuntoCarga
End Property
Property Let IdTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
End Property

Sub MostrarFormulario()
    oFormulario.IdTipoFinanciamiento = ml_IdTipoFinanciamiento
    oFormulario.IdFarmaciaElegida = lnIdFarmaciaElegida
    oFormulario.FarmaciaElegida = lcFarmaciaElegida
    oFormulario.idPuntoCarga = mi_idPuntoCarga
    oFormulario.RegistraTodosLosItems = ml_RegistraTodosLosItems
    oFormulario.Show 1
    mi_BotonPresionado = oFormulario.BotonPresionado
    If mi_BotonPresionado = sghAceptar Then
        mi_idProducto = oFormulario.idProducto
        mi_Producto = oFormulario.Producto
        mi_Descripcion = oFormulario.Descripcion
        mi_Precio = oFormulario.Precio
        mi_Codigo = oFormulario.Codigo
        mi_idPuntoCarga = oFormulario.idPuntoCarga
        Set oRsItemsMasivosElegidos = oFormulario.ItemsMasivosElegidos
        Set oRsDevuelveTodosLosItemsServ = oFormulario.DevuelveTodosLosItemsServ
        Set oRsDevuelveTodosLosItemsFarm = oFormulario.DevuelveTodosLosItemsFarm
    End If
    Set oFormulario = Nothing
End Sub

Property Get ItemsMasivosElegidos() As Recordset
    Set ItemsMasivosElegidos = oRsItemsMasivosElegidos
End Sub

Property Get DevuelveTodosLosItemsServ() As Recordset
    Set DevuelveTodosLosItemsServ = oRsDevuelveTodosLosItemsServ
End Property

Property Get DevuelveTodosLosItemsFarm() As Recordset
    Set DevuelveTodosLosItemsFarm = oRsDevuelveTodosLosItemsFarm
End Property

