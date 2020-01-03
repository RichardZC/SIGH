VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucDespachoDonaciones 
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13005
   ScaleHeight     =   5685
   ScaleWidth      =   13005
   Begin VB.ComboBox cmbOrden 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "ucDespachoDonaciones.ctx":0000
      Left            =   6210
      List            =   "ucDespachoDonaciones.ctx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5340
      Width           =   1905
   End
   Begin VB.TextBox txtBusca 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8160
      TabIndex        =   0
      Top             =   5340
      Width           =   1995
   End
   Begin UltraGrid.SSUltraGrid grillaBusqueda 
      Height          =   2475
      Left            =   210
      TabIndex        =   2
      Top             =   810
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
      Appearance      =   "ucDespachoDonaciones.ctx":0026
      Caption         =   ".."
   End
   Begin UltraGrid.SSUltraGrid grdProductos 
      Height          =   5325
      Left            =   0
      TabIndex        =   3
      Top             =   0
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
      Left            =   12210
      TabIndex        =   6
      Top             =   5370
      Width           =   555
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
      Left            =   30
      TabIndex        =   5
      Top             =   5400
      Width           =   4995
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Orden:"
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
      Left            =   5610
      TabIndex        =   4
      Top             =   5400
      Width           =   555
   End
End
Attribute VB_Name = "ucDespachoDonaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Control para Items del Despacho de Donaciones
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
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
Dim ml_idProducto As Long
Dim ml_IdAlmacen As Long
Dim dTotalIngresado  As Double
Dim ml_idTipoPrecioParaNiNs As Long

Property Let TipoPrecioParaNiNs(lValue As Long)
   ml_idTipoPrecioParaNiNs = lValue
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

Sub AgregaProducto(lbPulsaF10 As Boolean)
'    If mRs_Productos.RecordCount >= lnMaximoNroItems Then
'       MsgBox "Solo se permite registrar hasta " & Trim(Str(lnMaximoNroItems)) & " Items", vbExclamation, "Productos"
'       Exit Sub
'    End If
    grdProductos.SetFocus
    If lbPulsaF10 Then
       SendKeys "{Tab}"
    End If
    mb_CargandoProductos = True
    AgregaRegistro
    mb_CargandoProductos = False
    Totalizar
    mb_FilaEditable = True
    grdProductos.PerformAction ssKeyActionActivateCell
    grdProductos.PerformAction ssKeyActionEnterEditMode
End Sub


Sub AgregaRegistro()
    On Error GoTo errAR
    With mRs_Productos
        .AddNew
        .Fields!idProducto = 0
        .Fields!codigo = ""
        .Fields!nombreProducto = ""
        .Fields!cantidad = 0
        .Fields!Precio = 0
        .Fields!total = 0
        .Fields!saldo = 0
    End With
errAR:
End Sub



Sub CargaProductosPorMovNumero()
   Dim rs As Recordset
   Dim oConexion As New ADODB.Connection
   oConexion.Open sighentidades.CadenaConexion
   oConexion.CursorLocation = adUseClient
   Set rs = mo_ReglasFarmacia.FarmMovimientosDetalleDevuelveTodosItems(oConexion, "S", ml_movNumero)
   CargarItemsALaGrilla rs
   oConexion.Close
   Set oConexion = Nothing
End Sub


Sub CargarItemsALaGrilla(rs As Recordset)
    Dim oRsTmp As New ADODB.Recordset
    Dim lnCantidad As Long
    Dim lcCodigo As String
    Dim lcNombreProducto As String
    Dim lnIdProducto As Long
    Dim lnPrecio As Double
    mb_CargandoProductos = True
    Do While Not rs.EOF
        lnIdProducto = rs!idProducto
        lcCodigo = rs!codigo
        lcNombreProducto = rs!Nombre
        lnPrecio = rs!Precio
        lnCantidad = 0
        Do While Not rs.EOF And lnIdProducto = rs!idProducto
            lnCantidad = lnCantidad + rs!cantidad
            rs.MoveNext
            If rs.EOF Then
               Exit Do
            End If
        Loop
        Set oRsTmp = mo_ReglasFarmacia.farmDevuelveSaldosSegunAlmacenProducto(ml_IdAlmacen, lnIdProducto)
        oRsTmp.Filter = "idTipoSalidaBienInsumo=" & sghTipoSalidaItemFarmacia.sghDonaciones
        mRs_Productos.AddNew
        mRs_Productos!idProducto = lnIdProducto
        mRs_Productos!codigo = lcCodigo
        mRs_Productos!nombreProducto = lcNombreProducto
        mRs_Productos!cantidad = lnCantidad
        mRs_Productos!Precio = lnPrecio
        mRs_Productos!total = Round(lnCantidad * lnPrecio, 2)
        mRs_Productos!saldo = oRsTmp.Fields!cantidad + lnCantidad
        oRsTmp.Close
    Loop
    mb_CargandoProductos = False
    Totalizar
    Set grdProductos.DataSource = mRs_Productos
End Sub



Sub Totalizar()
    Dim rsProductos As New ADODB.Recordset
    Set rsProductos = mRs_Productos.Clone
    dTotalIngresado = 0
    If rsProductos.RecordCount > 0 Then
        If Not (rsProductos.EOF And rsProductos.BOF) Then
            rsProductos.MoveFirst
            Do While Not rsProductos.EOF
                rsProductos.Fields!total = Round(rsProductos.Fields!cantidad * rsProductos.Fields!Precio, 2)
                rsProductos.Update
                dTotalIngresado = dTotalIngresado + rsProductos!total
                rsProductos.Update
                rsProductos.MoveNext
            Loop
        End If
    End If
    lblTotal.Caption = "Total:    " & Format(dTotalIngresado, "####,###,##0.00")
End Sub



Private Sub cmbOrden_Click()
  If cmbOrden.ListIndex = 0 Then
        mRs_Productos.Sort = "codigo"
    Else
        mRs_Productos.Sort = "NombreProducto"
    End If
    grdProductos.Refresh
End Sub



Private Sub grdProductos_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    Totalizar
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



Private Sub grdProductos_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
        If mb_FilaEditable Then
            'Si la fila es editable y estamos en la celda de codigo se completa los datos
            'del producto
            Select Case grdProductos.ActiveCell.Column.Key
            Case "Codigo"
                ConfigurarProductoPorCodigo grdProductos
            Case "Cantidad"
                Dim oRow As SSRow
                Set oRow = grdProductos.ActiveCell.Row
                If oRow.Cells("cantidad").Value > 0 And oRow.Cells("cantidad").Value <= oRow.Cells("saldo").Value Then
                   Totalizar
                Else
                   If oRow.Cells("saldo").Value > 0 Then
                        MsgBox "La cantidad debe ser Menor o igual a  " & Trim(Str(oRow.Cells("saldo").Value)), vbInformation, "Mensaje"
                        oRow.Cells("cantidad").Value = 0
                   End If
                End If
            End Select
            grillaBusqueda.Visible = False
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







Private Sub grdProductos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdProductos
End Sub

Private Sub grdProductos_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    OnKeyDown grdProductos, KeyCode
    If KeyCode = vbKeyF2 Then
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
Dim lnPrecioUnitario As Double
    Set oRow = oGrilla.ActiveCell.Row
    
    If IsNull(oRow.Cells("codigo").Value) Or Trim(oRow.Cells("codigo").Value) = "" Then
        Exit Sub
    End If
    lcFiltro = Trim(oRow.Cells("codigo").Value)
 
    Set rs = mo_ReglasFarmacia.FarmDevuelveSaldosSinLotesSegunAlmacen(ml_IdAlmacen, 0, lcFiltro)
    rs.Filter = "idTipoSalidaBienInsumoSaldo=" & sghTipoSalidaItemFarmacia.sghDonaciones
    If rs.RecordCount > 0 Then
            'Busca si ya existe el producto
            If Not ItemYaExiste(rs.Fields("idproducto").Value) Then
                lnPrecioUnitario = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(rs.Fields("idproducto").Value, ml_idTipoPrecioParaNiNs)
                oRow.Cells("Idproducto").Value = rs.Fields("idproducto").Value
                oRow.Cells("NombreProducto").Value = rs.Fields("Nombre").Value
                oRow.Cells("precio").Value = lnPrecioUnitario
                oRow.Cells("saldo").Value = rs.Fields("saldo").Value
                oRow.Cells("Total").Value = 0
                oRow.Cells("cantidad").Value = 0
            End If
    End If

End Sub

'***************daniel barrantes**************
'***************Verifica si YA SE REGISTRO el ITEM (al momento de registrar)
'***************
Function ItemYaExiste(lnIdProducto As Long) As Boolean
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
        oRsTmp.Close
End Function


Sub OnKeyDown(oGrilla As SSUltraGrid, KeyCode As UltraGrid.SSReturnShort)
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
        End Select
End Sub

Sub OnKeyPress(oGrilla As SSUltraGrid, KeyAscii As UltraGrid.SSReturnShort)
                
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
               SendKeys "{Tab}"
               AgregaProducto (False)
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
                rs.Filter = "idTipoSalidaBienInsumoSaldo=" & sghTipoSalidaItemFarmacia.sghDonaciones
                Set grillaBusqueda.DataSource = rs
                'grillaBusqueda.Left = oGrilla.Left
                If mRs_Productos.RecordCount < 7 Then
                   grillaBusqueda.Top = oGrilla.ActiveCell.GetUIElement.Rect.Bottom * Screen.TwipsPerPixelY
                Else
                   grillaBusqueda.Top = 0
                End If
                grillaBusqueda.Visible = True
                grillaBusqueda.Enabled = True
                
            End Select
        End If

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
     grdProductos.Bands(0).Columns("NombreProducto").Width = 6700
     grdProductos.Bands(0).Columns("saldo").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("saldo").Width = 800
     grdProductos.Bands(0).Columns("cantidad").Width = 800
     grdProductos.Bands(0).Columns("cantidad").Format = "###0"
     grdProductos.Bands(0).Columns("Precio").Header.Caption = "Pr.Donación"
     grdProductos.Bands(0).Columns("Precio").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("Precio").Width = 700
     grdProductos.Bands(0).Columns("Precio").Format = "#0.000"
     grdProductos.Bands(0).Columns("Total").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("Total").Format = "#0.00"
    
ConfigEstilo:
    gridInfra.ConfigurarFilasBiColores oGrilla, sighentidades.GrillaConFilasBicolor
    
End Sub





Private Sub grillaBusqueda_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrillaBusqueda grillaBusqueda
    gridInfra.ConfigurarFilasBiColores grillaBusqueda, sighentidades.GrillaConFilasBicolor
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
    oGrilla.Bands(0).Columns("Nombre").Width = 10000
    oGrilla.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Saldo").Width = 800
    oGrilla.Bands(0).Columns("Saldo").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Saldo").Format = "#0"
    
    gridInfra.ConfigurarFilasBiColores oGrilla, sighentidades.GrillaConFilasBicolor
errInic:
End Sub
Private Sub grillaBusqueda_DblClick()
Dim fila As New Record
Dim lnIdProductoBusqueda As Long
    'debb-hra-ya en version Polsalud
    On Error GoTo ErrGrillaBusqueda
    lnIdProductoBusqueda = grillaBusqueda.ActiveRow.Cells("idproducto").Value
    If ItemYaExiste(lnIdProductoBusqueda) Then
        grdProductos.ActiveRow.Cells("codigo").Value = ""
        grdProductos.ActiveRow.Cells("idproducto").Value = 0
        grdProductos.ActiveRow.Cells("NombreProducto").Value = ""
        grdProductos.ActiveRow.Cells("saldo").Value = 0
        grdProductos.ActiveRow.Cells("precio").Value = 0
    Else
        RefrescarDatos
        Set grillaBusqueda.DataSource = Nothing
        grillaBusqueda.Visible = False
        SendKeys "{Tab}"
        SendKeys "{Tab}"
    End If
ErrGrillaBusqueda:
End Sub
Sub RefrescarDatos()
    Dim fila As New Record
    Dim lnPrecioUnitario As Double
    
    If Not grillaBusqueda.ActiveRow Is Nothing Then
                lnPrecioUnitario = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(grillaBusqueda.ActiveRow.Cells("idproducto").Value, ml_idTipoPrecioParaNiNs)
    
               grdProductos.ActiveRow.Cells("codigo").Value = grillaBusqueda.ActiveRow.Cells("CODIGO").Value
               grdProductos.ActiveRow.Cells("idproducto").Value = grillaBusqueda.ActiveRow.Cells("idproducto").Value
               grdProductos.ActiveRow.Cells("NombreProducto").Value = grillaBusqueda.ActiveRow.Cells("nombre").Value
               grdProductos.ActiveRow.Cells("precio").Value = lnPrecioUnitario
               grdProductos.ActiveRow.Cells("saldo").Value = grillaBusqueda.ActiveRow.Cells("saldo").Value
               grdProductos.ActiveRow.Cells("Total").Value = 0
               grdProductos.ActiveRow.Cells("cantidad").Value = 0
               Totalizar
    End If

End Sub

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





Private Sub txtBusca_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtBusca.Text <> "" Then
            txtBusca.Text = Trim(txtBusca.Text)
            mRs_Productos.MoveFirst
            If cmbOrden.ListIndex = 0 Then
               mRs_Productos.Find "codigo='" & txtBusca.Text & "'"
            Else
               Do While Not mRs_Productos.EOF
                  If Left(mRs_Productos!nombreProducto, Len(txtBusca.Text)) = UCase(txtBusca.Text) Then
                     Exit Do
                  End If
                  mRs_Productos.MoveNext
               Loop
            End If
            grdProductos.Refresh
      End If
   End If
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   grdProductos.Top = 0
   grdProductos.Left = 0
   grdProductos.Width = UserControl.Width
   grdProductos.Height = UserControl.Height - UserControl.Label1.Height - 100
   
   Label1.Top = UserControl.Height - UserControl.Label1.Height - 50
   Label2.Top = UserControl.Height - UserControl.Label1.Height - 50
   cmbOrden.Top = UserControl.Height - UserControl.Label1.Height - 100
   txtBusca.Top = UserControl.Height - UserControl.Label1.Height - 100
   lblTotal.Top = UserControl.Height - UserControl.Label1.Height - 50
   lblTotal.Left = UserControl.Width - Len(lblTotal) * 150
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

Sub TabEnDescripcion()
    On Error Resume Next
    grdProductosFocusColumna "NombreProducto"
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

