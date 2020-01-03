VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucVentas 
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13050
   ScaleHeight     =   5685
   ScaleWidth      =   13050
   Begin VB.CommandButton cmdVerStock 
      Caption         =   "Ver stocks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9015
      TabIndex        =   6
      Top             =   5400
      Width           =   915
   End
   Begin UltraGrid.SSUltraGrid grillaBusqueda 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   3201
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   71303188
      BorderStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   "ucVentas.ctx":0000
      Caption         =   ".."
   End
   Begin UltraGrid.SSUltraGrid grdProductos 
      Height          =   5325
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   9393
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Productos"
   End
   Begin Threed.SSOption optPorCodigo 
      Height          =   255
      Left            =   5310
      TabIndex        =   4
      Top             =   5400
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   450
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Reg.Por Código"
   End
   Begin Threed.SSOption optPorDescripcion 
      Height          =   255
      Left            =   6990
      TabIndex        =   5
      Top             =   5400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Reg. Por Descripción"
      Value           =   -1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Teclas de ayuda: <F10> = Agregar         <Supr>  = Eliminar "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   5430
      Width           =   4995
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   12240
      TabIndex        =   2
      Top             =   5400
      Width           =   555
   End
End
Attribute VB_Name = "ucVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Control para Items de VENTAS
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Public Event Totalizado(lnTotalIngresado As Double)
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Public Event SeIngresoProducto(lcCodigo As String)
Dim gridInfra As New GridInfragistic
Dim mb_CargandoProductos As Boolean
Dim mRs_Productos As New Recordset
Dim mb_FilaEditable As Boolean
Dim ml_ultimoProductoEliminado As Long
Dim mo_ProductosEliminados As New Collection
Dim lnMaximoNroItems As Long
Dim ml_movNumero As String
Dim lcSql As String
Dim ml_idProducto As Long
Dim ml_IdAlmacen As Long
Dim dTotalIngresado  As Double
Dim ml_IdTipoVentaSeleccionada As Long          '0=VentaDirecta      1=PreVenta
Dim ml_idPreVenta As Long
Dim ml_IdTipoFinanciamiento As Long
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim ml_ElTipoFinanciamientoSeUsaEnFarmacia As Boolean
Dim ml_idTipoPrecioParaNiNs As Long
Dim ml_PermiteAgregarItems As Boolean
Dim ml_esReceta As Boolean
Dim lcCodigoSeleccionado As String, lcItemSeleccionado As String

Property Let esReceta(lValue As Boolean)
    ml_esReceta = lValue
End Property

Property Let PermiteAgregarItems(lValue As Boolean)
    ml_PermiteAgregarItems = lValue
End Property

Property Let TipoPrecioParaNiNs(lValue As Long)
   ml_idTipoPrecioParaNiNs = lValue
End Property



Property Let ElTipoFinanciamientoSeUsaEnFarmacia(lValue As Boolean)
   ml_ElTipoFinanciamientoSeUsaEnFarmacia = lValue
End Property
Property Let IdTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
End Property

Property Get idProducto() As Long
    idProducto = ml_idProducto
End Property

Property Let IdAlmacen(lValue As Long)
   ml_IdAlmacen = lValue
End Property
Property Let movNumero(lValue As String)
   ml_movNumero = lValue
End Property
Property Let TipoVentaSeleccionada(lValue As Long)
   ml_IdTipoVentaSeleccionada = lValue
   lnMaximoNroItems = BuscarMaximoItemsEnParametros()
End Property
Property Let idPreVenta(lValue As Long)
   ml_idPreVenta = lValue
End Property

Sub inicializar()
    Set mRs_Productos = New Recordset
    GenerarRecordsetProductos
    lnMaximoNroItems = BuscarMaximoItemsEnParametros()

End Sub

'debb-18/05/2016
Function BuscarMaximoItemsEnParametros() As Long
        ml_PermiteAgregarItems = True
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        If ml_IdTipoVentaSeleccionada = 1 Then
           'preventa
           BuscarMaximoItemsEnParametros = Val(lcBuscaParametro.SeleccionaFilaParametro(102))
           If UCase(lcBuscaParametro.SeleccionaFilaParametro(500)) = "S" Then
              BuscarMaximoItemsEnParametros = 500
           End If
        Else
           'ventas directas
           BuscarMaximoItemsEnParametros = Val(lcBuscaParametro.SeleccionaFilaParametro(247))
        End If
        Set lcBuscaParametro = Nothing
End Function

Sub AgregaProducto(lbPulsaF10 As Boolean)
    On Error GoTo ErrAddP
    If mRs_Productos.RecordCount >= lnMaximoNroItems Then
       MsgBox "Solo se permite registrar hasta " & Trim(Str(lnMaximoNroItems)) & " Items", vbExclamation, "Productos"
       Exit Sub
    End If
    grdProductos.SetFocus
    If lbPulsaF10 Then
      ' SendKeys "{Tab}"
    End If
    mb_CargandoProductos = True
    AgregaRegistro
    mb_CargandoProductos = False
    Totalizar
    mb_FilaEditable = True
ErrAddP:
End Sub


Sub AgregaRegistro()
    On Error GoTo errLetras
    With mRs_Productos
        .AddNew
        .Fields!idProducto = 0
        .Fields!codigo = ""
        .Fields!nombreProducto = ""
        .Fields!Cantidad = 0
        .Fields!Precio = 0
        .Fields!total = 0
        .Fields!saldo = 0
    End With
errLetras:
End Sub



Sub CargaProductosPorMovNumero()
   Dim rs As Recordset
   Dim oConexion As New ADODB.Connection
   oConexion.Open SIGHEntidades.CadenaConexion
   oConexion.CursorLocation = adUseClient
   Set rs = mo_ReglasFarmacia.FarmMovimientosDetalleDevuelveTodosItems(oConexion, "S", ml_movNumero)
   CargarItemsALaGrilla rs
   rs.Close
   Set rs = Nothing
   oConexion.Close
   Set oConexion = Nothing
   On Error Resume Next
   mRs_Productos.MoveFirst
End Sub
Sub CargaProductosPorIdPreVenta()
        Dim rs As Recordset
        Dim oRsTmp As New Recordset
        Set rs = mo_ReglasFarmacia.FarmPreventaDetalleDevuelveTodosItems(Val(ml_movNumero))
        mb_CargandoProductos = True
        Do While Not rs.EOF
            Set oRsTmp = mo_ReglasFarmacia.farmDevuelveSaldosSegunAlmacenProducto(ml_IdAlmacen, rs!idProducto)
            mRs_Productos.AddNew
            mRs_Productos!idProducto = rs!idProducto
            mRs_Productos!codigo = rs!codigo
            mRs_Productos!nombreProducto = rs!Nombre
            mRs_Productos!Cantidad = rs!Cantidad
            mRs_Productos!Precio = rs!Precio
            mRs_Productos!total = Round(rs!Precio * rs!Cantidad, 2)
            mRs_Productos!saldo = oRsTmp.Fields!Cantidad         '+ rs!cantidad
            mRs_Productos!tipo = mo_ReglasFarmacia.DevuelveTipoProducto(rs!idProducto)
            oRsTmp.Close
            rs.MoveNext
        Loop
        mb_CargandoProductos = False
        Totalizar
        Set grdProductos.DataSource = mRs_Productos
        mRs_Productos.MoveFirst
   
End Sub


Sub CargarItemsALaGrilla(rs As Recordset)
    Dim oRsTmp As New ADODB.Recordset
    Dim lnCantidad As Long
    Dim lcCodigo As String
    Dim lcNombreProducto As String
    Dim lnIdProducto As Long
    Dim lnPrecio As Double
    Dim lnSaldo As Long
    Dim lnPrecioSeguro As Double, lcFormaF As String
    Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    mb_CargandoProductos = True
    Do While Not rs.EOF
        lnIdProducto = rs!idProducto
        lcCodigo = rs!codigo
        lcNombreProducto = rs!Nombre
        lnPrecio = rs!Precio
        lnCantidad = 0
        Do While Not rs.EOF And lnIdProducto = rs!idProducto
            lnCantidad = lnCantidad + rs!Cantidad
            rs.MoveNext
            If rs.EOF Then
               Exit Do
            End If
        Loop
        '
        Set oRsTmp = mo_ReglasFarmacia.farmDevuelveSaldosSegunAlmacenProducto(ml_IdAlmacen, lnIdProducto)
        lnSaldo = 0
        If oRsTmp.RecordCount > 0 Then lnSaldo = oRsTmp.Fields!Cantidad
        oRsTmp.Close
        '
        lnPrecioSeguro = 0
        lcFormaF = ""
        If ml_ElTipoFinanciamientoSeUsaEnFarmacia Then   'busca precio del seguro
            Set oRsTmp = mo_ReglasFacturacion.FacturacionBienesPorCodigo(lcCodigo, ml_IdTipoFinanciamiento, oConexion)
            If oRsTmp.RecordCount > 0 Then
               lnPrecioSeguro = oRsTmp.Fields!PrecioUnitario
                lcFormaF = IIf(IsNull(oRsTmp.Fields!FormaFarmaceutica), "", oRsTmp.Fields!FormaFarmaceutica)
            End If
            oRsTmp.Close
            
        End If
        mRs_Productos.AddNew
        mRs_Productos!idProducto = lnIdProducto
        mRs_Productos!codigo = lcCodigo
        mRs_Productos!nombreProducto = lcNombreProducto
        mRs_Productos!Cantidad = lnCantidad
        mRs_Productos!Precio = lnPrecio
        mRs_Productos!total = Round(lnCantidad * lnPrecio, 2)
        mRs_Productos!saldo = lnSaldo + lnCantidad
        mRs_Productos!PrecioDelSeguro = lnPrecioSeguro
        mRs_Productos!formaF = lcFormaF
        mRs_Productos!CantidadSinEditar = lnCantidad
        mRs_Productos!tipo = mo_ReglasFarmacia.DevuelveTipoProducto(lnIdProducto, oConexion)
    Loop
    mb_CargandoProductos = False
    Totalizar
    Set grdProductos.DataSource = mRs_Productos
    oConexion.Close
    Set oConexion = Nothing
End Sub



Sub Totalizar()
    Dim rsProductos As New ADODB.Recordset
    Dim lnLin As Integer
    Dim lnLinTot As Integer
    Dim lnNro As Integer
    lnNro = 1
    lnLin = 1
    Set rsProductos = mRs_Productos.Clone
    lnLinTot = rsProductos.RecordCount
    dTotalIngresado = 0
    If rsProductos.RecordCount > 0 Then
        If Not (rsProductos.EOF And rsProductos.BOF) Then
            'rsProductos.CursorLocation = adUseClient
            rsProductos.MoveFirst
            Do While Not rsProductos.EOF
                If rsProductos.Fields!codigo = "" Or rsProductos.Fields!nombreProducto = "" Then
                    If lnLin < lnLinTot Then
                       rsProductos.Delete
                       rsProductos.Update
                    End If
                Else
                    rsProductos.Fields!Nro = lnNro
                    rsProductos.Fields!total = Round(rsProductos.Fields!Cantidad * rsProductos.Fields!Precio, 2)
                    rsProductos.Update
                    dTotalIngresado = dTotalIngresado + rsProductos!total
                    rsProductos.Update
                    lnNro = lnNro + 1
                End If
                lnLin = lnLin + 1
                rsProductos.MoveNext
            Loop
        End If
    End If
    lblTotal.Caption = "Total:    " & Format(dTotalIngresado, "####,###,##0.00")
    RaiseEvent Totalizado(dTotalIngresado)
    
End Sub







Private Sub cmdVerStock_Click()
    On Error GoTo ErrVerStock
    Dim lcMensajeLicencia As String
    'If mo_ReglasComunes.EESSconDerechosAmejoras(2, "61008", lcMensajeLicencia) = True Then
        Dim oVerStock As New FarmVerStock
        If Trim(lcCodigoSeleccionado) = "" Then
            Dim rsRecordset As ADODB.Recordset
            Set rsRecordset = grdProductos.DataSource
            lcCodigoSeleccionado = rsRecordset("codigo")
            lcItemSeleccionado = rsRecordset("NombreProducto")
            Set rsRecordset = Nothing
        End If
        oVerStock.codigo = lcCodigoSeleccionado
        oVerStock.Producto = Trim(lcCodigoSeleccionado) & " " & lcItemSeleccionado
        oVerStock.Show 1
        Set oVerStock = Nothing
    'End If
ErrVerStock:
End Sub

Private Sub grdProductos_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
Totalizar
End Sub

'Eventos de la grilla de servicios
Private Sub grdProductos_AfterRowActivate()
     If ml_PermiteAgregarItems = True Then
        If mb_CargandoProductos Then
            Exit Sub
        End If
     End If
     On Error Resume Next
     Dim oRsProd As New Recordset
     Set oRsProd = grdProductos.DataSource
     RaiseEvent SeIngresoProducto(oRsProd!codigo)
     Set oRsProd = Nothing
End Sub


Private Sub grdProductos_AfterRowsDeleted()
'   If ml_PermiteAgregarItems = True Then
    If ml_ultimoProductoEliminado > 0 Then
        mo_ProductosEliminados.Add ml_ultimoProductoEliminado
        ml_ultimoProductoEliminado = 0
        Totalizar
    Else
        Totalizar
        Set grdProductos.DataSource = mRs_Productos
    End If
 '  End If
End Sub



Private Sub grdProductos_BeforeCellActivate(ByVal Cell As UltraGrid.SSCell, ByVal Cancel As UltraGrid.SSReturnBoolean)
lcSql = ""
End Sub
'debb-29/12/2016
Private Sub grdProductos_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
     Dim oRow As SSRow
     Dim lbSeguir1 As Boolean
     If ml_PermiteAgregarItems = True Then
        If mb_FilaEditable Then
            'Si la fila es editable y estamos en la celda de codigo se completa los datos
            'del producto
            Select Case grdProductos.ActiveCell.Column.Key
            Case "Codigo"
                ConfigurarProductoPorCodigo grdProductos
            Case "Cantidad"
                Set oRow = grdProductos.ActiveCell.Row
                lbSeguir1 = True
                If ml_esReceta = True And oRow.Cells("cantidad").Value > oRow.Cells("cantidadSinEditar").Value Then
                   MsgBox "La cantidad debe ser Menor o igual a  " & Trim(Str(oRow.Cells("cantidadSinEditar").Value)), vbInformation, "RECETA"
                   oRow.Cells("cantidad").Value = oRow.Cells("cantidadSinEditar").Value
                End If
                If lbSeguir1 = True Then
                    If oRow.Cells("cantidad").Value > 0 And oRow.Cells("cantidad").Value <= oRow.Cells("saldo").Value Then
                       oRow.Cells("total").Value = oRow.Cells("cantidad").Value * oRow.Cells("precio").Value
                       'Totalizar
                    Else
                       If oRow.Cells("saldo").Value > 0 Then
                            MsgBox "La cantidad debe ser Menor o igual a  " & Trim(Str(oRow.Cells("saldo").Value)), vbInformation, "Mensaje"
                            oRow.Cells("cantidad").Value = 0
                       End If
                    End If
                End If
                Totalizar
            End Select
            grillaBusqueda.Visible = False
        End If
    Else
        If ml_esReceta = True Then
            Set oRow = grdProductos.ActiveCell.Row
            lbSeguir1 = True
            If ml_esReceta = True And oRow.Cells("cantidad").Value > oRow.Cells("cantidadSinEditar").Value Then
                 MsgBox "La cantidad debe ser Menor o igual a  " & Trim(Str(oRow.Cells("cantidadSinEditar").Value)), vbInformation, "RECETA"
                 oRow.Cells("cantidad").Value = oRow.Cells("cantidadSinEditar").Value
            End If
        End If
    End If
End Sub

Private Sub grdProductos_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
  If ml_PermiteAgregarItems = True Then
    'Si la fila no es editable, cancela cualquier cambio en la fila
    If Not mb_FilaEditable Then
        Cancel = True
        Exit Sub
    End If
  Else
   '  Cancel = True    'debb-29/12/2016
  End If
End Sub

Private Sub grdProductos_BeforeRowActivate(ByVal Row As UltraGrid.SSRow)
    If ml_PermiteAgregarItems = True Then mb_FilaEditable = True
End Sub




Private Sub grdProductos_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
'   If ml_PermiteAgregarItems = True Then
        Cancel = True
        If MsgBox("¿Desea eliminar el registro?", vbYesNo, "Ventas") = vbYes Then
            mRs_Productos.Delete
            mRs_Productos.Update
            Totalizar
            Set grdProductos.DataSource = mRs_Productos
        End If
'   Else
'        Cancel = True
 '  End If
End Sub

Private Sub grdProductos_Click()
    Dim rsRecordset As ADODB.Recordset
    On Error Resume Next
    Set rsRecordset = grdProductos.DataSource
    lcCodigoSeleccionado = rsRecordset("codigo")
    lcItemSeleccionado = rsRecordset("NombreProducto")
    Set rsRecordset = Nothing
End Sub

Private Sub grdProductos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    InicializarLaGrilla grdProductos
End Sub

Private Sub grdProductos_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    OnKeyDown grdProductos, KeyCode
    If KeyCode = vbKeyF2 Or KeyCode = vbKeyF3 Then
       Dim lnKeyCode As Integer
       lnKeyCode = KeyCode
       RaiseEvent SePresionoTeclaEspecial(lnKeyCode)
    End If
End Sub

Private Sub grdProductos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    OnKeyPress grdProductos, KeyAscii
End Sub


Sub ConfigurarProductoPorCodigo(oGrilla As SSUltraGrid)
Dim rs As Recordset
Dim oRow As SSRow
Dim lcFiltro As String
Dim lnPrecioSeguro As Double
Dim lnPrecioUnitario As Double, lcFormaF As String
Dim oConexion As New Connection
oConexion.Open SIGHEntidades.CadenaConexion
oConexion.CursorLocation = adUseClient
    Set oRow = oGrilla.ActiveCell.Row
    
    
    If IsNull(oRow.Cells("codigo").Value) Or Trim(oRow.Cells("codigo").Value) = "" Then
        Exit Sub
    End If
    lcFiltro = Trim(oRow.Cells("codigo").Value)
    '
    lnPrecioSeguro = 0: lcFormaF = ""
    If ml_ElTipoFinanciamientoSeUsaEnFarmacia Then   'busca precio del seguro
        Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigo(lcFiltro, ml_IdTipoFinanciamiento, oConexion)
        If rs.RecordCount = 0 Then
           MsgBox "Ese codigo tiene SALDO pero no está registrado en el Catálogo del TIPO DE FINANCIAMIENTO elegido"
           'AgregaProducto (True)
           Exit Sub
        Else
           lnPrecioSeguro = rs.Fields!PrecioUnitario
           lcFormaF = rs.Fields!FormaFarmaceutica
        End If
        rs.Close
    End If
    '
    Set rs = mo_ReglasFarmacia.FarmDevuelveSaldosSinLotesSegunAlmacen(ml_IdAlmacen, 0, lcFiltro)
    rs.Filter = "idTipoSalidaBienInsumoSaldo=" & sghTipoSalidaItemFarmacia.sghSoloVenta
    If rs.RecordCount > 0 Then
        If rs.Fields!idTipoSalidaBienInsumo = sghTipoSalidaItemFarmacia.sghVentaEstrategico Then
            MsgBox "Ese Medicamento es ESTRATEGICO (no se vende)"
            oRow.Cells("codigo").Value = ""
        Else
            'Busca si ya existe el producto
            If Not ItemYaExiste(rs.Fields("idproducto").Value, rs.Fields("codigo").Value) Then
                '
                lnPrecioUnitario = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(rs.Fields("idproducto").Value, ml_idTipoPrecioParaNiNs)
                oRow.Cells("Idproducto").Value = rs.Fields("idproducto").Value
                oRow.Cells("NombreProducto").Value = rs.Fields("Nombre").Value
                oRow.Cells("precio").Value = lnPrecioUnitario
                oRow.Cells("saldo").Value = rs.Fields("saldo").Value
                oRow.Cells("Total").Value = 0
                oRow.Cells("cantidad").Value = 0
                oRow.Cells("PrecioDelSeguro").Value = lnPrecioSeguro
                oRow.Cells("FormaF").Value = lcFormaF
                oRow.Cells("Tipo").Value = mo_ReglasFarmacia.DevuelveTipoProducto(rs.Fields("idproducto").Value)
                SIGHEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "+Cod:" & Trim(rs!codigo)
            End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    
    
End Sub

'***************daniel barrantes**************
'***************Verifica si YA SE REGISTRO el ITEM (al momento de registrar)
'***************
Function ItemYaExiste(lnIdProducto As Long, lcCodigo As String) As Boolean
        Dim lbExiste As Boolean
        Dim oRsTmp As New ADODB.Recordset
        Set oRsTmp = mRs_Productos.Clone
        ItemYaExiste = False
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
                If oRsTmp.Fields!idProducto = lnIdProducto Then
                   ItemYaExiste = True
                   MsgBox "Este Producto ya está registrado", vbInformation, "Farmacia"
                   Exit Do
                End If
                oRsTmp.MoveNext
           Loop
        End If
        Set oRsTmp = Nothing
        If ItemYaExiste = False Then
           RaiseEvent SeIngresoProducto(lcCodigo)
        End If
End Function


Sub OnKeyDown(oGrilla As SSUltraGrid, KeyCode As UltraGrid.SSReturnShort)
    If ml_PermiteAgregarItems = True Then
        If Not oGrilla.ActiveCell Is Nothing Then
            Select Case oGrilla.ActiveCell.Column.Key
            Case "Cantidad"
                Select Case Val(Chr(KeyCode))
                Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 0
                Case Else
                    KeyCode = 0
                End Select
            Case "NombreProducto"
                Select Case KeyCode
                Case vbKeyBack
                Case vbKeyEscape
                    Set grillaBusqueda.DataSource = Nothing
                    grillaBusqueda.Visible = False
                Case vbKeyReturn
                Case vbKeyDown
                    On Error Resume Next
                    grillaBusqueda.SetFocus
                Case vbKeyLeft
                End Select
            End Select
        End If
        Select Case KeyCode
        Case vbKeyF10
            AgregaProducto (True)
            If optPorCodigo.Value = True Then
                grdProductosFocusColumna "codigo"
            Else
                grdProductosFocusColumna "NombreProducto"
            End If
        End Select
    End If
End Sub

Sub OnKeyPress(oGrilla As SSUltraGrid, KeyAscii As UltraGrid.SSReturnShort)
    If ml_PermiteAgregarItems = True Then
        'Si la fila no es editable, cancela cualquier cambio en la fila
        If Not mb_FilaEditable Then
            Exit Sub
        End If
        
        If oGrilla.ActiveCell Is Nothing Then
            Exit Sub
        End If

        If oGrilla.ActiveCell.Column.Key = "Codigo" And KeyAscii = 13 Then
            SendKeys "{Tab}"
            If Trim(oGrilla.ActiveCell.GetText) <> "" Then
                SendKeys "{Tab}"
                SendKeys "{Tab}"
            End If
            Exit Sub
        End If
        If oGrilla.ActiveCell.Column.Key = "Cantidad" Then
            If KeyAscii = 13 Then
               'SendKeys "{Tab}"
               AgregaProducto (False)
               If optPorCodigo.Value = True Then
                   grdProductosFocusColumna "codigo"
               Else
                   grdProductosFocusColumna "NombreProducto"
               End If
            End If
            Exit Sub
        End If


        If oGrilla.ActiveCell.Column.Key = "NombreProducto" Then
            Select Case KeyAscii
            Case vbKeyEscape
                If Trim(oGrilla.ActiveCell.GetText) = "" Then
                    grillaBusqueda.Visible = False
                    Set grillaBusqueda.DataSource = Nothing
                End If
            Case vbKeyReturn
            Case vbKeyDown
            Case vbKeyLeft
            Case Else
                Dim sNombre As String
                Select Case KeyAscii
                Case vbKeyBack
                    sNombre = oGrilla.ActiveCell.GetText
                Case Else
                    sNombre = oGrilla.ActiveCell.GetText + Chr(KeyAscii)
                End Select
                
                Dim rs As New Recordset
                Set rs = mo_ReglasFarmacia.FarmDevuelveSaldosSinLotesSegunAlmacen(ml_IdAlmacen, 1, sNombre)
                rs.Filter = "idTipoSalidaBienInsumoSaldo=" & sghTipoSalidaItemFarmacia.sghSoloVenta
                Set grillaBusqueda.DataSource = rs
                'grillaBusqueda.Left = oGrilla.Left
                If mRs_Productos.RecordCount < 5 Then
                   grillaBusqueda.Top = oGrilla.ActiveCell.GetUIElement.Rect.Bottom * Screen.TwipsPerPixelY
                Else
                   grillaBusqueda.Top = 0
                End If
                grillaBusqueda.Visible = True
                grillaBusqueda.Enabled = True
                
            End Select
        End If
    End If
End Sub


Sub GenerarRecordsetProductos()
    With mRs_Productos
          .Fields.Append "Nro", adInteger, 4
          .Fields.Append "IdProducto", adInteger, 4
          .Fields.Append "Codigo", adChar, 20
          .Fields.Append "NombreProducto", adChar, 300
          .Fields.Append "Saldo", adInteger
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Precio", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "PrecioDelSeguro", adDouble
          .Fields.Append "FormaF", adVarChar, 20
          .Fields.Append "CantidadSinEditar", adInteger
          .Fields.Append "Tipo", adVarChar, 20
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    'Set grdProductos.DataSource = mRs_Productos
End Sub

Private Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
     On Error GoTo ConfigEstilo
     grdProductos.Bands(0).Columns("IdProducto").Hidden = True
     grdProductos.Bands(0).Columns("PrecioDelSeguro").Hidden = True
     grdProductos.Bands(0).Columns("formaF").Hidden = True
     grdProductos.Bands(0).Columns("CantidadSinEditar").Hidden = True
     grdProductos.Bands(0).Columns("IdProducto").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("nro").Width = 300
     grdProductos.Bands(0).Columns("nro").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("codigo").Width = 900
     grdProductos.Bands(0).Columns("NombreProducto").Width = 6800
     grdProductos.Bands(0).Columns("saldo").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("saldo").Width = 600
     grdProductos.Bands(0).Columns("cantidad").Width = 800
     grdProductos.Bands(0).Columns("cantidad").Format = "###0"
     grdProductos.Bands(0).Columns("Precio").Header.Caption = "Pr.Venta"
     grdProductos.Bands(0).Columns("Precio").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("Precio").Width = 700
     grdProductos.Bands(0).Columns("Precio").Format = "#0.000"
     grdProductos.Bands(0).Columns("Total").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("Total").Width = 1000
     grdProductos.Bands(0).Columns("Total").Format = "#0.00"
     grdProductos.Bands(0).Columns("tipo").Hidden = True
ConfigEstilo:
    gridInfra.ConfigurarFilasBiColores oGrilla, SIGHEntidades.GrillaConFilasBicolor
    
End Sub





Private Sub grillaBusqueda_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrillaBusqueda grillaBusqueda
    gridInfra.ConfigurarFilasBiColores grillaBusqueda, SIGHEntidades.GrillaConFilasBicolor
End Sub
Private Sub InicializarLaGrillaBusqueda(oGrilla As SSUltraGrid)
    On Error GoTo errInic
    oGrilla.Bands(0).Columns("IdProducto").Hidden = True
    oGrilla.Bands(0).Columns("idAlmacen").Hidden = True
    oGrilla.Bands(0).Columns("Precio").Hidden = True
    oGrilla.Bands(0).Columns("idTipoSalidaBienInsumo").Hidden = True
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Código"
    oGrilla.Bands(0).Columns("Codigo").Width = 800
    oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(0).Columns("Nombre").Header.Caption = "Descripción"
    If oGrilla.Name = "grillaBusqueda" Then
       oGrilla.Bands(0).Columns("Nombre").Width = 6000
    Else
       oGrilla.Bands(0).Columns("Nombre").Width = 7000
    End If
    oGrilla.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Saldo").Width = 800
    oGrilla.Bands(0).Columns("Saldo").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Saldo").Format = "#0"
    
    gridInfra.ConfigurarFilasBiColores oGrilla, SIGHEntidades.GrillaConFilasBicolor
errInic:
End Sub
Private Sub grillaBusqueda_DblClick()
Dim fila As New Record
Dim lnIdProductoBusqueda As Long
    'debb-hra-ya en version Polsalud
    On Error GoTo ErrGrillaBusqueda
    lnIdProductoBusqueda = grillaBusqueda.ActiveRow.Cells("idproducto").Value
    If ItemYaExiste(lnIdProductoBusqueda, grillaBusqueda.ActiveRow.Cells("codigo").Value) Then
        grdProductos.ActiveRow.Cells("codigo").Value = ""
        grdProductos.ActiveRow.Cells("idproducto").Value = 0
        grdProductos.ActiveRow.Cells("NombreProducto").Value = ""
        grdProductos.ActiveRow.Cells("saldo").Value = 0
        grdProductos.ActiveRow.Cells("precio").Value = 0
    Else
        If RefrescarDatos = True Then
            Set grillaBusqueda.DataSource = Nothing
            grillaBusqueda.Visible = False
            SendKeys "{Tab}"
            SendKeys "{Tab}"
        Else
            grdProductosFocusColumna ("NombreProducto")
        End If
    End If
ErrGrillaBusqueda:
End Sub
Function RefrescarDatos() As Boolean
    Dim fila As New Record
    Dim rs As Recordset
    Dim lnPrecioSeguro As Double
    Dim lbContinua As Boolean
    Dim lnPrecioUnitario As Double, lcFormaF As String
    Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    If (Not grillaBusqueda.ActiveRow Is Nothing) Then
       If (grillaBusqueda.ActiveRow.Cells("saldo").Value > 0) Then
                lbContinua = True
                lnPrecioSeguro = 0: lcFormaF = ""
                If ml_ElTipoFinanciamientoSeUsaEnFarmacia Then     'busca precio del seguro
                    Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigo(grillaBusqueda.ActiveRow.Cells("CODIGO").Value, ml_IdTipoFinanciamiento, oConexion)
                    If rs.RecordCount = 0 Then
                       MsgBox "Ese codigo tiene SALDO pero no está registrado en el Catálogo del TIPO DE FINANCIAMIENTO elegido"
                       Exit Function
                    Else
                       lnPrecioSeguro = rs.Fields!PrecioUnitario
                       lcFormaF = rs.Fields!FormaFarmaceutica
                    End If
                    rs.Close
                End If
                If grillaBusqueda.ActiveRow.Cells("idTipoSalidaBienInsumo") = sghTipoSalidaItemFarmacia.sghSoloEstrategico Then
                    MsgBox "Ese Medicamento es ESTRATEGICO (no se vende)"
                    lbContinua = False
                End If
                lnPrecioUnitario = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(grillaBusqueda.ActiveRow.Cells("idproducto").Value, ml_idTipoPrecioParaNiNs)
                If ml_ElTipoFinanciamientoSeUsaEnFarmacia = False Then
                   lnPrecioSeguro = lnPrecioUnitario
                End If
                If lbContinua Then
                    grdProductos.ActiveRow.Cells("codigo").Value = grillaBusqueda.ActiveRow.Cells("CODIGO").Value
                    grdProductos.ActiveRow.Cells("idproducto").Value = grillaBusqueda.ActiveRow.Cells("idproducto").Value
                    grdProductos.ActiveRow.Cells("NombreProducto").Value = grillaBusqueda.ActiveRow.Cells("nombre").Value
                    grdProductos.ActiveRow.Cells("precio").Value = lnPrecioSeguro   'lnPrecioUnitario
                    grdProductos.ActiveRow.Cells("saldo").Value = grillaBusqueda.ActiveRow.Cells("saldo").Value
                    grdProductos.ActiveRow.Cells("Total").Value = 0
                    grdProductos.ActiveRow.Cells("cantidad").Value = 0
                    grdProductos.ActiveRow.Cells("PrecioDelSeguro").Value = lnPrecioSeguro
                    grdProductos.ActiveRow.Cells("FormaF").Value = lcFormaF
                    grdProductos.ActiveRow.Cells("Tipo").Value = mo_ReglasFarmacia.DevuelveTipoProducto(grillaBusqueda.ActiveRow.Cells("idproducto").Value)
                    Totalizar
                    RefrescarDatos = True
                    SIGHEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "+Cod:" & Trim(grillaBusqueda.ActiveRow.Cells("CODIGO").Value)
                End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Function

Private Sub grillaBusqueda_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)

    Select Case KeyCode
    Case vbKeyEscape
        Set grillaBusqueda.DataSource = Nothing
        grillaBusqueda.Visible = False
    Case vbKeyReturn
        grillaBusqueda_DblClick
    Case vbKeyDown, vbKeyUp
       ' RefrescarDatos
    End Select
    
End Sub






Private Sub UserControl_Resize()
   
    On Error Resume Next
   grdProductos.Top = 0
   grdProductos.Left = 0
   grdProductos.Width = UserControl.Width
   grdProductos.Height = UserControl.Height - UserControl.Label1.Height - 100
   
   Label1.Top = UserControl.Height - UserControl.Label1.Height - 50
   lblTotal.Top = UserControl.Height - UserControl.Label1.Height - 50
   lblTotal.Left = UserControl.Width - Len(lblTotal) * 450   '150
   optPorCodigo.Top = UserControl.Height - UserControl.Label1.Height - 80
   optPorDescripcion.Top = UserControl.Height - UserControl.Label1.Height - 80
   UserControl.cmdVerStock.Top = UserControl.Height - UserControl.Label1.Height - 80
   
End Sub

Sub LimpiarGrilla()
        If mRs_Productos Is Nothing Then
            Exit Sub
        End If

        Set grdProductos.DataSource = Nothing

        If mRs_Productos.RecordCount > 0 Then
            mRs_Productos.MoveFirst
            Do While Not mRs_Productos.EOF
                mRs_Productos.Delete
                mRs_Productos.Update
                mRs_Productos.MoveNext
            Loop
        End If
        grillaBusqueda.Visible = False
        CargaProductosPorMovNumero
End Sub



Property Get DevuelveProductos() As Recordset
    'Se debe utilizar un clon del recrdset, ya que si se trabaja directamente con el recordset
    'que esta asociado a la grilla ocurre errores en los metodos movenext, movefirst, etc.
    Set DevuelveProductos = mRs_Productos.Clone()
End Property
Property Get DevuelveTotal() As Double
    DevuelveTotal = dTotalIngresado
End Property

Sub RefrescaSaldos()

End Sub



Private Sub mnuAgregarServicio_Click()
        'grdProductos.SetFocus
        'SendKeys "{Tab}"
        AgregaProducto True
End Sub



Sub TabEnDescripcion()
    On Error Resume Next
    mnuAgregarServicio_Click
    If optPorCodigo.Value = True Then
        grdProductosFocusColumna "codigo"
    Else
        grdProductosFocusColumna "NombreProducto"
    End If
End Sub

Sub grdProductosFocusColumna(lcNombreColumna As String)
    With grdProductos
        'scroll the column into view
        .ActiveColScrollRegion.ScrollColumnIntoView .Bands(0).Columns(lcNombreColumna), True
        If Not .ActiveRow Is Nothing Then
            'if there is an activerow then activate the cell from this column
            Set .ActiveCell = .ActiveRow.Cells(lcNombreColumna)
            .ActiveCell.Selected = True
        End If
        'give the grid focus
        .SetFocus
        .PerformAction ssKeyActionActivateCell
        .PerformAction ssKeyActionEnterEditMode
    End With
End Sub

'debb-18/05/2016
Public Sub cargaPaqueteElegido(lnIdPaquete As Long)
        Dim rs As New Recordset
        Dim oRsTmp As New Recordset
        Dim lnPrecio As Double, lnCantidadSaldo As Long
        Dim lcMensaje As String, lnCantidadItems As Long
        Dim mo_ReglasFacturacion As New ReglasFacturacion
        On Error Resume Next
        If mRs_Productos.RecordCount > 0 Then
           mRs_Productos.MoveFirst
           Do While Not mRs_Productos.EOF
               mRs_Productos.Delete
               mRs_Productos.Update
                mRs_Productos.MoveNext
           Loop
        End If
        Set rs = mo_ReglasFacturacion.FacturacionCatalogoPaqueteFarmSeleccionarXid(lnIdPaquete)
        mb_CargandoProductos = True
        lcMensaje = ""
        lnCantidadItems = 0
        Do While Not rs.EOF
            If rs.Fields!idTipoSalidaBienInsumo = sghTipoSalidaItemFarmacia.sghSoloVenta Or rs.Fields!idTipoSalidaBienInsumo = sghTipoSalidaItemFarmacia.sghVentaEstrategico Then
                lnPrecio = mo_ReglasFarmacia.DevuelvePrecioSegunTipoFinanciamiento(rs!idProducto, ml_IdTipoFinanciamiento)
                lnCantidadSaldo = 0
                Set oRsTmp = mo_ReglasFarmacia.farmDevuelveSaldosSegunAlmacenProducto(ml_IdAlmacen, rs!idProducto)
                If oRsTmp.RecordCount > 0 Then
                   oRsTmp.MoveFirst
                   Do While Not oRsTmp.EOF
                      If oRsTmp.Fields!idTipoSalidaBienInsumo = sghTipoSalidaItemFarmacia.sghSoloVenta Or oRsTmp.Fields!idTipoSalidaBienInsumo = sghTipoSalidaItemFarmacia.sghVentaEstrategico Then
                         lnCantidadSaldo = lnCantidadSaldo + oRsTmp.Fields!Cantidad
                      End If
                      oRsTmp.MoveNext
                   Loop
                End If
                oRsTmp.Close
                If lnCantidadSaldo >= rs!Cantidad Then
                        mRs_Productos.AddNew
                        mRs_Productos!idProducto = rs!idProducto
                        mRs_Productos!codigo = rs!codigo
                        mRs_Productos!nombreProducto = rs!Nombre
                        mRs_Productos!Cantidad = rs!Cantidad
                        mRs_Productos!Precio = lnPrecio
                        mRs_Productos!total = Round(lnPrecio * rs!Cantidad, 2)
                        mRs_Productos!saldo = lnCantidadSaldo
                        mRs_Productos!tipo = mo_ReglasFarmacia.DevuelveTipoProducto(rs!idProducto)
                        lnCantidadItems = lnCantidadItems + 1
                       If lnCantidadItems >= lnMaximoNroItems Then
                          Exit Do
                       End If
                Else
                       lcMensaje = lcMensaje & "No hay SALDO para " & rs!codigo & "-" & rs!Nombre & Chr(13)
                End If
            Else
                lcMensaje = lcMensaje & "(Se Despacha por INTERV.SANITARIA) " & rs!codigo & "-" & rs!Nombre & Chr(13)
            End If
            rs.MoveNext
        Loop
        mb_CargandoProductos = False
        Totalizar
        Set grdProductos.DataSource = mRs_Productos
        If mRs_Productos.RecordCount > 0 Then mRs_Productos.MoveFirst
        Set rs = Nothing
        Set oRsTmp = Nothing
        If lcMensaje <> "" Then
           MsgBox lcMensaje, vbInformation, "Problema con Saldos"
        End If
End Sub
'debb-29/12/2016
Sub CargaCantidadRecetada(rs As Recordset)
    If rs.RecordCount > 0 And mRs_Productos.RecordCount > 0 Then
       mRs_Productos.MoveFirst
       Do While Not mRs_Productos.EOF
            rs.MoveFirst
            rs.Find "idItem=" & mRs_Productos!idProducto
            If Not rs.EOF Then
                mRs_Productos.Fields!CantidadSinEditar = rs!cantidadPedida
                mRs_Productos.Update
            End If
            mRs_Productos.MoveNext
       Loop
    End If
End Sub

Sub CargaProductosPorIdReceta(rs As Recordset)
    Dim oRsTmp As New Recordset, lcMensaje As String, lnCantidad As Long
    mb_CargandoProductos = True
    lcMensaje = ""
    On Error Resume Next
    If mRs_Productos.RecordCount > 0 Then
       mRs_Productos.MoveFirst
       Do While Not mRs_Productos.EOF
           mRs_Productos.Delete
           mRs_Productos.Update
            mRs_Productos.MoveNext
       Loop
    End If
    Do While Not rs.EOF
        If rs!Precio = 0 Then
            lcMensaje = lcMensaje & "No hay PRECIO para " & rs!codigo & "-" & rs!Nombre & Chr(13)
        Else
            Set oRsTmp = mo_ReglasFarmacia.farmDevuelveSaldosSegunAlmacenProducto(ml_IdAlmacen, rs!iditem)
            If oRsTmp.RecordCount > 0 Then
                lnCantidad = IIf(rs!cantidadPedida > oRsTmp!Cantidad, oRsTmp!Cantidad, rs!cantidadPedida)
                mRs_Productos.AddNew
                mRs_Productos!idProducto = rs!iditem
                mRs_Productos!codigo = rs!codigo
                mRs_Productos!nombreProducto = rs!Nombre
                mRs_Productos!Cantidad = lnCantidad          'rs!cantidadPedida    'debb-31/10/2016
                mRs_Productos!Precio = rs!Precio
                mRs_Productos!total = lnCantidad * rs!Precio 'rs!total             'debb-31/10/2016
                mRs_Productos!saldo = oRsTmp.Fields!Cantidad
                mRs_Productos!PrecioDelSeguro = rs!Precio    'debb2014b
                mRs_Productos!tipo = mo_ReglasFarmacia.DevuelveTipoProducto(rs!iditem)
                mRs_Productos!CantidadSinEditar = lnCantidad
            Else
                lcMensaje = lcMensaje & "No hay SALDO para " & rs!codigo & "-" & rs!Nombre & Chr(13)
            End If
        End If
        rs.MoveNext
    Loop
    mb_CargandoProductos = False
    Totalizar
    Set grdProductos.DataSource = mRs_Productos
    If mRs_Productos.RecordCount > 0 Then mRs_Productos.MoveFirst
    Set rs = Nothing
    Set oRsTmp = Nothing
    If lcMensaje <> "" Then
       MsgBox lcMensaje, vbInformation, "Problema con Saldos"
    End If
End Sub


