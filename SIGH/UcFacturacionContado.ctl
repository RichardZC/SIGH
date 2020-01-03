VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl UcFacturacionContado 
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13050
   ScaleHeight     =   5460
   ScaleWidth      =   13050
   Begin UltraGrid.SSUltraGrid grillaBusqueda 
      Height          =   2475
      Left            =   240
      TabIndex        =   0
      Top             =   870
      Visible         =   0   'False
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   4366
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   71303188
      BorderStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   "UcFacturacionContado.ctx":0000
      Caption         =   ".."
   End
   Begin UltraGrid.SSUltraGrid grdProductos 
      Height          =   5325
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   9393
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Productos"
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
      Top             =   5430
      Width           =   555
   End
End
Attribute VB_Name = "UcFacturacionContado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para Facturación al Contado
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim gridInfra As New GridInfragistic
Dim mb_CargandoProductos As Boolean
Dim mRs_Productos As New Recordset
Dim mb_FilaEditable As Boolean
Dim ml_ultimoProductoEliminado As Long
Dim mo_ProductosEliminados As New Collection
Dim lnMaximoNroItems As Long
Dim ml_movNumero As String
Dim lcSql As String
Dim ml_IdProducto As Long
Dim ml_IdAlmacen As Long
Dim dTotalIngresado  As Double
Dim ml_IdTipoVentaSeleccionada As Long          '0=VentaDirecta      1=PreVenta
Dim ml_idPreVenta As Long
Dim ml_IdTipoFinanciamiento As Long
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim lb_inHabilitaEdicionColumnasDelGrid As Boolean


Property Let InHabilitaEdicionColumnasDelGrid(lValue As Boolean)
    lb_inHabilitaEdicionColumnasDelGrid = lValue
End Property

Property Let idTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
End Property

Property Get idProducto() As Long
    idProducto = ml_IdProducto
End Property

Property Let IdAlmacen(lValue As Long)
   ml_IdAlmacen = lValue
End Property
Property Let movNumero(lValue As String)
   ml_movNumero = lValue
End Property
Property Let TipoVentaSeleccionada(lValue As Long)
   ml_IdTipoVentaSeleccionada = lValue
End Property
Property Let idPreVenta(lValue As Long)
   ml_idPreVenta = lValue
End Property

Sub inicializar()
    Set mRs_Productos = New Recordset
    GenerarRecordsetProductos
    lnMaximoNroItems = BuscarMaximoItemsEnParametros()

End Sub

Function BuscarMaximoItemsEnParametros() As Long
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        BuscarMaximoItemsEnParametros = Val(lcBuscaParametro.SeleccionaFilaParametro(102))
        Set lcBuscaParametro = Nothing
End Function

'debb-18/05/2016
Sub CargaProductosPorIdPreVenta(ByRef lnTotalPreventa As Double)
    Dim rs As Recordset
    Dim oRsTmp As New Recordset
    Dim oRsTmp9 As New Recordset
    Dim lbContinuar As Boolean
    Set oRsTmp9 = mo_ReglasFacturacion.FacturacionBienesPagosXidPreventa(ml_idPreVenta)
    oRsTmp9.Filter = "idEstadoFacturacion<>1"
    Set rs = mo_ReglasFarmacia.FarmPreVentaDetalleDevuelveTodosItems(ml_idPreVenta)
    mb_CargandoProductos = True
    Do While Not rs.EOF
        lbContinuar = True
        If oRsTmp9.RecordCount > 0 Then
           oRsTmp9.MoveFirst
           oRsTmp9.Find "idProducto=" & rs!idProducto
           If Not oRsTmp9.EOF Then
              lbContinuar = False
           End If
        End If
        If lbContinuar = True Then
            Set oRsTmp = mo_ReglasFarmacia.farmDevuelveSaldosSegunAlmacenProducto(ml_IdAlmacen, rs!idProducto)
            mRs_Productos.AddNew
            mRs_Productos!idProducto = rs!idProducto
            mRs_Productos!Codigo = rs!Codigo
            mRs_Productos!NombreProducto = rs!Nombre
            mRs_Productos!Cantidad = rs!Cantidad
            mRs_Productos!Precio = rs!Precio
            mRs_Productos!Total = Round(rs!Precio * rs!Cantidad, 2)
            mRs_Productos!saldo = oRsTmp.Fields!Cantidad + rs!Cantidad
            oRsTmp.Close
        End If
        rs.MoveNext
    Loop
    mb_CargandoProductos = False
    Totalizar
    lnTotalPreventa = dTotalIngresado
    Set grdProductos.DataSource = mRs_Productos
    If mRs_Productos.RecordCount > 0 Then
       mRs_Productos.MoveFirst
    End If
    Set oRsTmp = Nothing
    Set oRsTmp9 = Nothing
End Sub


Sub Totalizar()
    Dim rsProductos As New ADODB.Recordset
    Set rsProductos = mRs_Productos.Clone
    dTotalIngresado = 0
    If rsProductos.RecordCount > 0 Then
        If Not (rsProductos.EOF And rsProductos.BOF) Then
            rsProductos.MoveFirst
            Do While Not rsProductos.EOF
                rsProductos.Fields!Total = Round(rsProductos.Fields!Cantidad * rsProductos.Fields!Precio, 2)
                rsProductos.Update
                dTotalIngresado = dTotalIngresado + rsProductos!Total
                rsProductos.Update
                rsProductos.MoveNext
            Loop
        End If
    End If
    lblTotal.Caption = "Total:    " & Format(dTotalIngresado, "####,###,##0.00")
End Sub


'Eventos de la grilla de servicios
Private Sub grdProductos_AfterRowActivate()
        If mb_CargandoProductos Then
            Exit Sub
        End If
End Sub


Private Sub grdProductos_AfterRowsDeleted()
    If ml_ultimoProductoEliminado > 0 Then
        mo_ProductosEliminados.Add ml_ultimoProductoEliminado
        ml_ultimoProductoEliminado = 0
        Totalizar
    Else
        Totalizar
        Set grdProductos.DataSource = mRs_Productos
    End If

End Sub

Private Sub grdProductos_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    'Si la fila no es editable, cancela cualquier cambio en la fila
    If Not mb_FilaEditable Then
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub grdProductos_BeforeRowActivate(ByVal Row As UltraGrid.SSRow)
    mb_FilaEditable = True
End Sub


Private Sub grdProductos_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If lb_inHabilitaEdicionColumnasDelGrid = False Then
    Else
       Cancel = True
    End If
End Sub

Private Sub grdProductos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdProductos
End Sub


Sub GenerarRecordsetProductos()
    With mRs_Productos
          .Fields.Append "IdProducto", adInteger, 4
          .Fields.Append "Codigo", adChar, 20
          .Fields.Append "NombreProducto", adChar, 300
          .Fields.Append "Saldo", adInteger
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Precio", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "NumeroDeItem", adInteger                                          'debb-18/05/2016
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    'Set grdProductos.DataSource = mRs_Productos
End Sub

Private Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
     On Error GoTo ConfigEstilo
     grdProductos.Bands(0).Columns("IdProducto").Hidden = True
     grdProductos.Bands(0).Columns("IdProducto").Activation = ssActivationActivateNoEdit
     
     grdProductos.Bands(0).Columns("codigo").Width = 1000
     grdProductos.Bands(0).Columns("NombreProducto").Width = 6500   '7300
     grdProductos.Bands(0).Columns("saldo").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("saldo").Width = 800
     grdProductos.Bands(0).Columns("cantidad").Width = 800
     grdProductos.Bands(0).Columns("cantidad").Format = "###0"
     If lb_inHabilitaEdicionColumnasDelGrid = True Then
        grdProductos.Bands(0).Columns("cantidad").Activation = ssActivationActivateNoEdit
     End If
     grdProductos.Bands(0).Columns("Precio").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("Precio").Width = 800
     grdProductos.Bands(0).Columns("Precio").Format = "#0.000"
     grdProductos.Bands(0).Columns("Total").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("Total").Format = "#0.00"
     
     grdProductos.Bands(0).Columns("NumeroDeItem").Hidden = True        'debb-18/05/2016
    
ConfigEstilo:
    gridInfra.ConfigurarFilasBiColores oGrilla, SIGHEntidades.GrillaConFilasBicolor
    
End Sub

Private Sub grdProductos_KeyUp(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    If KeyCode >= vbKeyF2 And KeyCode <= vbKeyF12 Then
       Dim lnKeyCode As Integer
       lnKeyCode = KeyCode
       RaiseEvent SePresionoTeclaEspecial(lnKeyCode)
    End If

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
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Código"
    oGrilla.Bands(0).Columns("Codigo").Width = 800
    oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Nombre").Header.Caption = "Descripción"
    oGrilla.Bands(0).Columns("Nombre").Width = 10000
    oGrilla.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Saldo").Width = 800
    oGrilla.Bands(0).Columns("Saldo").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Saldo").Format = "#0"
    oGrilla.Bands(0).Columns("Cantidad").Activation = ssActivationActivateNoEdit
    
    gridInfra.ConfigurarFilasBiColores oGrilla, SIGHEntidades.GrillaConFilasBicolor
errInic:
End Sub


Private Sub UserControl_Resize()
   
    On Error Resume Next
   grdProductos.Top = 0
   grdProductos.Left = 0
   grdProductos.Width = UserControl.Width
   grdProductos.Height = UserControl.Height - 100
   
   lblTotal.Top = UserControl.Height - 50
   lblTotal.Left = UserControl.Width - Len(lblTotal) * 150
End Sub

Sub LimpiarGrilla()
'        On Error GoTo ErrLimGrd
'        If mRs_Productos Is Nothing Then
'            Exit Sub
'        End If
'        Set grdProductos.DataSource = Nothing
'        If mRs_Productos.RecordCount > 0 Then
'            mRs_Productos.MoveFirst
'            Do While Not mRs_Productos.EOF
'                mRs_Productos.Delete
'                mRs_Productos.Update
'                mRs_Productos.MoveNext
'            Loop
'        End If
'        CargaProductosPorIdPreVenta
        On Error GoTo ErrLimGrd
        If mRs_Productos.RecordCount > 0 Then
            mRs_Productos.MoveFirst
            Do While Not mRs_Productos.EOF
                mRs_Productos.Delete
                mRs_Productos.Update
                mRs_Productos.MoveNext
            Loop
        End If
        Set grdProductos.DataSource = mRs_Productos
ErrLimGrd:
End Sub



Property Get DevuelveProductos() As Recordset
    'Se debe utilizar un clon del recrdset, ya que si se trabaja directamente con el recordset
    'que esta asociado a la grilla ocurre errores en los metodos movenext, movefirst, etc.
    Set DevuelveProductos = mRs_Productos.Clone()
End Property
Property Get DevuelveTotal() As Double
    DevuelveTotal = dTotalIngresado
End Property





