VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucFarmaciaItems 
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11790
   ScaleHeight     =   5685
   ScaleWidth      =   11790
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
      Left            =   8190
      TabIndex        =   6
      Top             =   5370
      Width           =   1995
   End
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
      ItemData        =   "ucFarmaciaItems.ctx":0000
      Left            =   6240
      List            =   "ucFarmaciaItems.ctx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5370
      Width           =   1905
   End
   Begin UltraGrid.SSUltraGrid grillaBusqueda 
      Height          =   1695
      Left            =   900
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   2990
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   71303188
      BorderStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   "ucFarmaciaItems.ctx":0026
      Caption         =   ".."
   End
   Begin UltraGrid.SSUltraGrid grdProductos 
      Height          =   5325
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   9393
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Productos"
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
      Left            =   5640
      TabIndex        =   4
      Top             =   5430
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10950
      TabIndex        =   2
      Top             =   5400
      Width           =   555
   End
End
Attribute VB_Name = "ucFarmaciaItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Control para items
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Public Event OnClick(oRecordset As Recordset)
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim gridInfra As New GridInfragistic
Dim ms_TipoProducto As sghTipoProducto
Dim mb_CargandoProductos As Boolean
Dim mRs_Productos As New Recordset
Dim mb_FilaEditable As Boolean
Dim ml_ultimoProductoEliminado As Long
Dim mo_ProductosEliminados As New Collection
Dim lnMaximoNroItems As Long
Dim ml_IdInventario As Long
Dim ml_IdPuntoCarga As Long
Dim lcSql As String
Dim ml_idProducto As Long
Dim dTotalIngresado  As Double

Property Get idProducto() As sghOpciones
    idProducto = ml_idProducto
End Property

Property Let IdPuntoCarga(lValue As Long)
   ml_IdPuntoCarga = lValue
End Property
Property Let IdInventario(lValue As Long)
   ml_IdInventario = lValue
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
    If mRs_Productos.RecordCount >= lnMaximoNroItems Then
       MsgBox "Solo se permite registrar hasta " & Trim(Str(lnMaximoNroItems)) & " Items", vbExclamation, "Productos"
       Exit Sub
    End If
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
    With mRs_Productos
        .AddNew
        .Fields!idProducto = 0
        .Fields!codigo = ""
        .Fields!nombreProducto = ""
        .Fields!cantidad = 0
        .Fields!Precio = 0
        .Fields!total = 0
    End With
End Sub



Sub CargaProductosPorIdInventario()
   Dim rs As Recordset
   Set rs = mo_ReglasFarmacia.FarmInventarioDetalleDevuelveProductosLotesPorId(ml_IdInventario)
   CargarItemsALaGrilla rs
End Sub


Sub CargarItemsALaGrilla(rs As Recordset)
    mb_CargandoProductos = True
    Do While Not rs.EOF
        mRs_Productos.AddNew
        mRs_Productos!idProducto = rs!idProducto
        mRs_Productos!codigo = rs!codigo
        mRs_Productos!nombreProducto = rs!nombreProducto
        mRs_Productos!cantidad = rs!cantidad
        mRs_Productos!Precio = rs!Precio
        mRs_Productos!total = rs!total
        rs.MoveNext
    Loop
    mb_CargandoProductos = False
    Totalizar
    Set grdProductos.DataSource = mRs_Productos
End Sub



Sub Totalizar()
    Dim dTotalIngresado  As Double
    Dim rsProductos As New ADODB.Recordset
    Set rsProductos = mRs_Productos.Clone
    dTotalIngresado = 0
    If rsProductos.RecordCount > 0 Then
        If Not (rsProductos.EOF And rsProductos.BOF) Then
            rsProductos.MoveFirst
            Do While Not rsProductos.EOF
                rsProductos.Fields!total = rsProductos.Fields!cantidad * rsProductos.Fields!Precio
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






Private Sub grdProductos_DblClick()
     RaiseEvent OnClick(mRs_Productos)
End Sub

Private Sub grdProductos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdProductos
End Sub

Private Sub grdProductos_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    OnKeyDown grdProductos, KeyCode
End Sub

Private Sub grdProductos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    OnKeyPress grdProductos, KeyAscii
End Sub


Sub ConfigurarProductoPorCodigo(oGrilla As SSUltraGrid)
Dim rs As Recordset
Dim oRow As SSRow

    Set oRow = oGrilla.ActiveCell.Row
    
    If IsNull(oRow.Cells("codigo").Value) Or oRow.Cells("codigo").Value = "" Then
        Exit Sub
    End If
    If ms_TipoProducto = sghbien Then
       oRow.Cells("codigo").Value = Right("0000000000" & oRow.Cells("codigo").Value, 5)
       
    End If
    Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigodebb(oRow.Cells("codigo").Value, 1, ml_IdPuntoCarga)
    If rs.RecordCount > 0 Then
        'Busca si ya existe el producto
        If Not ItemYaExiste(rs.Fields("idproducto").Value) Then
            oRow.Cells("Idproducto").Value = rs.Fields("idproducto").Value
            oRow.Cells("NombreProducto").Value = rs.Fields("NombreProducto").Value
            oRow.Cells("precio").Value = rs.Fields("precioUnitario").Value
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
           oRsTmp.Find "idProducto=" & lnIdProducto
           If Not oRsTmp.EOF Then
              ItemYaExiste = True
              MsgBox "Este producto ya está registrado", vbInformation, "Farmacia"
           End If
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
                Dim lIdTipoFinanciamiento As Long
                Dim sNombre As String
                Select Case KeyAscii
                Case vbKeyBack
                    sNombre = oGrilla.ActiveCell.GetText
                Case Else
                    sNombre = oGrilla.ActiveCell.GetText + Chr(KeyAscii)
                End Select
                
                lIdTipoFinanciamiento = 1
                Dim rs As New Recordset
                Set rs = mo_AdminCaja.BienesFiltrarParaCajero(sNombre, lIdTipoFinanciamiento, ml_IdPuntoCarga)
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
     grdProductos.Bands(0).Columns("NombreProducto").Width = 9900
     grdProductos.Bands(0).Columns("Cantidad").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("Cantidad").Format = "#0"
     grdProductos.Bands(0).Columns("Precio").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("Precio").Format = "#0.00"
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
    oGrilla.Bands(0).Columns("Activo").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Código"
    oGrilla.Bands(0).Columns("Codigo").Width = 800
    
    oGrilla.Bands(0).Columns("Nombre").Header.Caption = "Descripción"
    oGrilla.Bands(0).Columns("Nombre").Width = 7800
    
    oGrilla.Bands(0).Columns("preciounitario").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
    
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
    
    If Not grillaBusqueda.ActiveRow Is Nothing Then
               grdProductos.ActiveRow.Cells("codigo").Value = grillaBusqueda.ActiveRow.Cells("CODIGO").Value
               grdProductos.ActiveRow.Cells("idproducto").Value = grillaBusqueda.ActiveRow.Cells("idproducto").Value
               grdProductos.ActiveRow.Cells("NombreProducto").Value = grillaBusqueda.ActiveRow.Cells("nombre").Value
               grdProductos.ActiveRow.Cells("precio").Value = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
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
        CargaProductosPorIdInventario
End Sub

Sub ActualizaCantidadTotalDeLotes(lnCantidadTotal As Long, lnIdProducto As Long)
        mRs_Productos.MoveFirst
        mRs_Productos.Find "idProducto = " & lnIdProducto
        If Not mRs_Productos.EOF Then
           mRs_Productos.Fields!cantidad = lnCantidadTotal
           mRs_Productos.Update
           Totalizar
        End If
End Sub

Property Get DevuelveProductos() As Recordset
    'Se debe utilizar un clon del recrdset, ya que si se trabaja directamente con el recordset
    'que esta asociado a la grilla ocurre errores en los metodos movenext, movefirst, etc.
    Set DevuelveProductos = mRs_Productos.Clone()
End Property
Property Get DevuelveTotal() As Double
    DevuelveTotal = dTotalIngresado
End Property

