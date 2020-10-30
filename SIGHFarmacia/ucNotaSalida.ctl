VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucNotaSalida 
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13050
   ScaleHeight     =   5880
   ScaleWidth      =   13050
   Begin VB.Frame fraSinLote 
      Height          =   2535
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   13005
      Begin VB.TextBox txtEsEstrategico 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7110
         MaxLength       =   50
         TabIndex        =   17
         Top             =   270
         Width           =   765
      End
      Begin VB.TextBox txtNombre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   16
         Top             =   270
         Width           =   4155
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10500
         MaxLength       =   30
         TabIndex        =   11
         Top             =   270
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         MaxLength       =   50
         TabIndex        =   10
         Top             =   270
         Width           =   855
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8490
         MaxLength       =   30
         TabIndex        =   9
         Top             =   270
         Width           =   855
      End
      Begin Threed.SSCommand btnAgregar 
         Height          =   465
         Left            =   11610
         TabIndex        =   12
         Top             =   165
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   820
         _Version        =   262144
         PictureFrames   =   1
         Picture         =   "ucNotaSalida.ctx":0000
         Caption         =   "Agregar"
         PictureAlignment=   9
      End
      Begin UltraGrid.SSUltraGrid grdProductosLotes 
         Height          =   1635
         Left            =   1320
         TabIndex        =   18
         Top             =   720
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   2884
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   71303188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   "ucNotaSalida.ctx":2F8C
         Caption         =   "Por Lotes/Fecha Vencimiento"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9720
         TabIndex        =   15
         Top             =   300
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   14
         Top             =   300
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8010
         TabIndex        =   13
         Top             =   300
         Width           =   435
      End
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
      ItemData        =   "ucNotaSalida.ctx":2FC8
      Left            =   6240
      List            =   "ucNotaSalida.ctx":2FD2
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5370
      Width           =   1905
   End
   Begin UltraGrid.SSUltraGrid grillaBusqueda 
      Height          =   1635
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   2884
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
      Appearance      =   "ucNotaSalida.ctx":2FEE
      Caption         =   ".."
   End
   Begin UltraGrid.SSUltraGrid grdProductos 
      Height          =   2535
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   4471
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
   Begin VB.Label lblPrecios 
      AutoSize        =   -1  'True
      Caption         =   "..."
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
      Left            =   10260
      TabIndex        =   7
      Top             =   5430
      Width           =   135
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
Attribute VB_Name = "ucNotaSalida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Control para Items de la Nota de Salida
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim gridInfra As New GridInfragistic
Dim mb_CargandoProductos As Boolean
Dim mRs_Productos As New Recordset
Dim oRsSaldosEnEsteMomento As New Recordset
Dim oRsSaldosEnEsteMoment1 As New Recordset
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
Dim mb_MuestraLoteParaDespachoNS As Boolean
Dim LdFechaMinimaDespacho As Date
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim lbDesdeCargarExcel As Boolean
Const lnTopGrilla = 3000

Property Let FechaMinimaDespacho(lValue As Date)
   LdFechaMinimaDespacho = lValue
End Property

Property Let MuestraLoteParaDespachoNS(lValue As Boolean)
   mb_MuestraLoteParaDespachoNS = lValue
End Property
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
    
    mo_Formulario.HabilitarDeshabilitar txtCodigo, False
    mo_Formulario.HabilitarDeshabilitar txtNombre, False
    mo_Formulario.HabilitarDeshabilitar txtEsEstrategico, False
    mo_Formulario.HabilitarDeshabilitar txtSaldo, False
    '
    If lcBuscaParametro.SeleccionaFilaParametro(544) = "S" Then
       fraSinLote.Top = lnTopGrilla
       grdProductosLotes.Height = fraSinLote.Height - (txtSaldo.Height + 900)
    Else
       fraSinLote.Top = UserControl.grdProductos.Top + grdProductos.Height + 70
    End If
End Sub

Function BuscarMaximoItemsEnParametros() As Long
       
        BuscarMaximoItemsEnParametros = Val(lcBuscaParametro.SeleccionaFilaParametro(102))
End Function

Sub AgregaProducto(lbPulsaF10 As Boolean)
    On Error GoTo AddP
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
AddP:
End Sub


Sub AgregaRegistro()
    On Error GoTo errAR
    With mRs_Productos
        .AddNew
        .Fields!idProducto = 0
        .Fields!codigo = ""
        .Fields!nombreProducto = ""
        .Fields!Lote = ""
'        .Fields!fechaVencimiento=null
        .Fields!Cantidad = 0
        .Fields!Precio = 0
        .Fields!Total = 0
        .Fields!saldo = 0
    End With
errAR:
End Sub



Sub CargaProductosPorMovNumero()
   Dim rs As Recordset
   Dim oConexion As New ADODB.Connection
   oConexion.Open SIGHEntidades.CadenaConexion
   oConexion.CursorLocation = adUseClient
   Set rs = mo_ReglasFarmacia.FarmMovimientosDetalleDevuelveTodosItems(oConexion, "S", ml_movNumero)
   CargarItemsALaGrilla rs
   oConexion.Close
   Set oConexion = Nothing
End Sub


Sub CargarItemsALaGrilla(rs As Recordset)
    Dim oRsTmp As New ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oFarmMovimientoDetalle As New farmMovimientoDetalle
    Dim lnSaldoDe As Long
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oFarmMovimientoDetalle.Conexion = oConexion
    mb_CargandoProductos = True
    Do While Not rs.EOF
        Set oRsTmp = oFarmMovimientoDetalle.farmDevuelveSaldosSegunAlmacenProductoLote(ml_IdAlmacen, rs.Fields!idProducto, rs.Fields!Lote, rs.Fields!FechaVencimiento, rs.Fields!idTipoSalidaBienInsumo)
        If oRsTmp.RecordCount > 0 Then
           lnSaldoDe = oRsTmp.Fields!Cantidad
        Else
           lnSaldoDe = 0
        End If
        mRs_Productos.AddNew
        mRs_Productos!idProducto = rs!idProducto
        mRs_Productos!codigo = rs!codigo
        mRs_Productos!nombreProducto = rs!Nombre
        mRs_Productos!Lote = rs!Lote
        mRs_Productos!FechaVencimiento = rs!FechaVencimiento
        mRs_Productos!Cantidad = rs!Cantidad
        mRs_Productos!Precio = rs!Precio
        mRs_Productos!Total = rs!Total
        mRs_Productos!saldo = lnSaldoDe + rs!Cantidad
        mRs_Productos.Fields("idTipoSalidaBienInsumo").Value = rs!idTipoSalidaBienInsumo
        oRsTmp.Close
        rs.MoveNext
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
                rsProductos.Fields!Total = Round(rsProductos.Fields!Cantidad * rsProductos.Fields!Precio, 2)
                rsProductos.Update
                dTotalIngresado = dTotalIngresado + rsProductos!Total
                rsProductos.Update
                rsProductos.MoveNext
            Loop
        End If
    End If
    lblTotal.Caption = "Total:    " & Format(dTotalIngresado, "####,###,##0.00")
    Select Case ml_idTipoPrecioParaNiNs
    Case 1
       'grdProductos.Bands(0).Columns("Precio").Header.Caption = "Pr.Compra"
       lblPrecios.Caption = "<Se usará Precio de Compra>"
    Case 2
       lblPrecios.Caption = "<Se usará Precio de Distribución>"
    Case 3
       lblPrecios.Caption = "<Se usará Precio de Venta>"
    Case 4
       lblPrecios.Caption = "<Se usará Precio de Donación>"
    End Select
End Sub

Sub DecargaPorCadaLote()
       Dim lbAgrego As Boolean
       Dim lnPrecioUnitario As Double
       lnPrecioUnitario = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(ml_idProducto, ml_idTipoPrecioParaNiNs)
       lbAgrego = False
       oRsSaldosEnEsteMoment1.MoveFirst
       lcSql = ""
       Do While Not oRsSaldosEnEsteMoment1.EOF
          If oRsSaldosEnEsteMoment1!Cantidad > 0 Then
               lbAgrego = True
               mRs_Productos.Fields("codigo").Value = txtCodigo.Text
               mRs_Productos.Fields("idproducto").Value = ml_idProducto
               mRs_Productos.Fields("NombreProducto").Value = txtNombre.Text
               mRs_Productos.Fields("precio").Value = lnPrecioUnitario
               mRs_Productos.Fields("lote").Value = oRsSaldosEnEsteMoment1.Fields!Lote
               mRs_Productos.Fields("fechaVencimiento").Value = oRsSaldosEnEsteMoment1.Fields!FechaVencimiento
               mRs_Productos.Fields("saldo").Value = oRsSaldosEnEsteMoment1.Fields!saldo
               mRs_Productos.Fields("idTipoSalidaBienInsumo").Value = oRsSaldosEnEsteMoment1.Fields("idTipoSalidaBienInsumoSaldo").Value
               mRs_Productos.Fields("cantidad").Value = oRsSaldosEnEsteMoment1!Cantidad
               AgregaRegistro
          End If
          oRsSaldosEnEsteMoment1.MoveNext
       Loop
       If lbAgrego = True Then
            fraSinLote.Visible = False
            SIGHEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "+Cod:" & Trim(txtCodigo.Text)
            Totalizar
            AgregaProducto (True)
       Else
            MsgBox lcSql, vbInformation, "No se puede dar salida"
       End If

End Sub

Private Sub btnAgregar_Click()
    If fraSinLote.Top = lnTopGrilla And Val(txtCantidad.Text) = 0 Then
       DecargaPorCadaLote
       Exit Sub
    End If
    Dim lnCantidadCargar As Long
    Dim lnCantSaldo As Long
    Dim lnPrecioUnitario As Double
    Dim lbAgrego As Boolean
    If Not (Val(txtCantidad.Text) > 0 And Val(txtCantidad.Text) <= Val(txtSaldo.Text)) Then
       MsgBox "La cantidad debe ser mayor a CERO y menor a " & txtSaldo.Text, vbInformation, "Farmacia"
       Exit Sub
    End If
    lbAgrego = False
    If oRsSaldosEnEsteMomento.RecordCount > 0 Then
       lnPrecioUnitario = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(ml_idProducto, ml_idTipoPrecioParaNiNs)
       lnCantSaldo = Val(txtCantidad.Text)
       oRsSaldosEnEsteMomento.MoveFirst
       lcSql = ""
       Do While Not oRsSaldosEnEsteMomento.EOF
          If oRsSaldosEnEsteMomento.Fields!saldo > 0 And oRsSaldosEnEsteMomento.Fields!FechaVencimiento >= LdFechaMinimaDespacho Then
               lbAgrego = True
               mRs_Productos.Fields("codigo").Value = txtCodigo.Text
               mRs_Productos.Fields("idproducto").Value = ml_idProducto
               mRs_Productos.Fields("NombreProducto").Value = txtNombre.Text
               mRs_Productos.Fields("precio").Value = lnPrecioUnitario
               mRs_Productos.Fields("lote").Value = oRsSaldosEnEsteMomento.Fields!Lote
               mRs_Productos.Fields("fechaVencimiento").Value = oRsSaldosEnEsteMomento.Fields!FechaVencimiento
               mRs_Productos.Fields("saldo").Value = oRsSaldosEnEsteMomento.Fields!saldo
               mRs_Productos.Fields("idTipoSalidaBienInsumo").Value = oRsSaldosEnEsteMomento.Fields("idTipoSalidaBienInsumoSaldo").Value
               If lnCantSaldo >= oRsSaldosEnEsteMomento.Fields!saldo Then
                  lnCantidadCargar = oRsSaldosEnEsteMomento.Fields!saldo
                  mRs_Productos.Fields("cantidad").Value = lnCantidadCargar
                  lnCantSaldo = lnCantSaldo - oRsSaldosEnEsteMomento.Fields!saldo
                  If lnCantSaldo <= 0 Then
                     Exit Do
                  End If
               Else
                  lnCantidadCargar = lnCantSaldo
                  mRs_Productos.Fields("cantidad").Value = lnCantidadCargar
                  Exit Do
               End If
               AgregaRegistro
          Else
               lcSql = "La fecha Vencimiento del ITEM es: " & oRsSaldosEnEsteMomento.Fields!FechaVencimiento
          End If
          oRsSaldosEnEsteMomento.MoveNext
       Loop
       If lbAgrego = True Then
            fraSinLote.Visible = False
            SIGHEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "+Cod:" & Trim(txtCodigo.Text)

            Totalizar
            AgregaProducto (True)
       Else
            MsgBox lcSql, vbInformation, "No se puede dar salida"
       End If
    End If
    
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
                If lbDesdeCargarExcel = False Then
                   ConfigurarProductoPorCodigo grdProductos
                End If
            Case "Cantidad"
                Dim oRow As SSRow
                Set oRow = grdProductos.ActiveCell.Row
                If oRow.Cells("cantidad").Value > 0 And oRow.Cells("cantidad").Value <= oRow.Cells("saldo").Value Then
                   Totalizar
                Else
                   MsgBox "La cantidad debe ser Menor o igual a  " & Trim(Str(oRow.Cells("saldo").Value)), vbInformation, "Mensaje"
                   oRow.Cells("cantidad").Value = 0
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
End Sub

Private Sub grdProductos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    OnKeyPress grdProductos, KeyAscii
End Sub

Sub CargaProductosExcel(rsExcel As Recordset)
    lbDesdeCargarExcel = True
    Set mRs_Productos = Nothing
    GenerarRecordsetProductos
    rsExcel.MoveFirst
    Do While Not rsExcel.EOF
       If rsExcel!saldo >= rsExcel!Cantidad And rsExcel!Cantidad > 0 And rsExcel!seleccionar = True Then
            Set oRsSaldosEnEsteMomento = mo_ReglasFarmacia.FarmDevuelveSaldosSinLotesSegunAlmacen(ml_IdAlmacen, 0, _
                                                           rsExcel!codigo)
            oRsSaldosEnEsteMomento.Filter = "idTipoSalidaBienInsumoSaldo=" & rsExcel!idTipoSalidaBienInsumo
            If oRsSaldosEnEsteMomento.RecordCount > 0 Then
                 txtSaldo.Text = oRsSaldosEnEsteMomento!saldo
                 oRsSaldosEnEsteMomento.Close
                 Set oRsSaldosEnEsteMomento = mo_ReglasFarmacia.FarmDevuelveSaldosConLotesSegunAlmacen(ml_IdAlmacen, 0, _
                                                                rsExcel!codigo)
                 If oRsSaldosEnEsteMomento.RecordCount > 0 Then
                     ml_idProducto = rsExcel!idProducto
                     txtCodigo.Text = rsExcel!codigo
                     txtNombre.Text = rsExcel!producto1
                     txtCantidad.Text = rsExcel!Cantidad
                     txtEsEstrategico.Text = SIGHEntidades.ElijeSiEsEstrategicoDevuelveNombre(rsExcel!idTipoSalidaBienInsumo)
                     AgregaRegistro
                     btnAgregar_Click
                 End If
            End If
       End If
       rsExcel.MoveNext
    Loop
    If mRs_Productos.RecordCount > 0 Then
       mRs_Productos.MoveFirst
       Do While Not mRs_Productos.EOF
          If mRs_Productos!idProducto = 0 Then
             mRs_Productos.Delete
             mRs_Productos.Update
          End If
          mRs_Productos.MoveNext
       Loop
       mRs_Productos.MoveFirst
    End If
    Set grdProductos.DataSource = mRs_Productos
End Sub
Sub ConfigurarProductoPorCodigo(oGrilla As SSUltraGrid)
Dim rs As Recordset
Dim oRow As SSRow
Dim lcFiltro As String
Dim oConexion As New ADODB.Connection
Dim oFarmMovimientoDetalle As New farmMovimientoDetalle
Dim lnPrecioUnitario As Double
    If ml_idTipoPrecioParaNiNs = 0 Then
       MsgBox "Debe elegir el Concepto antes de Registrar Productos", vbInformation, "Farmacia"
       Exit Sub
    End If
    
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    Set oRow = oGrilla.ActiveCell.Row
    If IsNull(oRow.Cells("codigo").Value) Or Trim(oRow.Cells("codigo").Value) = "" Then
        Exit Sub
    End If
    lcFiltro = oRow.Cells("codigo").Value
    Set oFarmMovimientoDetalle.Conexion = oConexion
    'Set rs = oFarmMovimientoDetalle.FarmDevuelveSaldosConLotesSegunAlmacen(ml_IdAlmacen, 0, lcFiltro)
    Set rs = oFarmMovimientoDetalle.FarmDevuelveSaldosConLotesSegunAlmacenCliente(ml_IdAlmacen, 0, lcFiltro)
    If rs.RecordCount > 0 Then
        'Busca si ya existe el producto
        If Not ItemYaExiste(rs.Fields("idproducto").Value, rs.Fields("lote").Value, rs.Fields("fechaVencimiento").Value, rs.Fields("idTipoSalidaBienInsumoSaldo").Value) Then
            If mb_MuestraLoteParaDespachoNS = True Then
                lnPrecioUnitario = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(rs.Fields("idproducto").Value, ml_idTipoPrecioParaNiNs)
                oRow.Cells("Idproducto").Value = rs.Fields("idproducto").Value
                oRow.Cells("NombreProducto").Value = rs.Fields("Nombre").Value
                oRow.Cells("precio").Value = lnPrecioUnitario
                oRow.Cells("saldo").Value = rs.Fields("saldo").Value
                oRow.Cells("lote").Value = rs.Fields("lote").Value
                oRow.Cells("fechaVencimiento").Value = rs.Fields("fechaVencimiento").Value
                oRow.Cells("Total").Value = 0
                oRow.Cells("cantidad").Value = 0
                oRow.Cells("idTipoSalidaBienInsumo").Value = rs.Fields("idTipoSalidaBienInsumoSaldo").Value
            Else
                ml_idProducto = rs!idProducto
                Set oRsSaldosEnEsteMomento = mo_ReglasFarmacia.FarmDevuelveSaldosSinLotesSegunAlmacen(ml_IdAlmacen, 0, lcFiltro)
                oRsSaldosEnEsteMomento.Filter = "idTipoSalidaBienInsumoSaldo=" & rs!idTipoSalidaBienInsumoSaldo
                If oRsSaldosEnEsteMomento.RecordCount > 0 Then
                    txtSaldo.Text = oRsSaldosEnEsteMomento!saldo
                    Set oRsSaldosEnEsteMomento = mo_ReglasFarmacia.FarmDevuelveSaldosConLotesSegunAlmacen(ml_IdAlmacen, 0, lcFiltro)
                    oRsSaldosEnEsteMomento.Filter = "idTipoSalidaBienInsumoSaldo=" & rs!idTipoSalidaBienInsumoSaldo
                    txtCodigo.Text = lcFiltro
                    txtNombre.Text = rs!Nombre
                    txtEsEstrategico.Text = SIGHEntidades.ElijeSiEsEstrategicoDevuelveNombre(rs!idTipoSalidaBienInsumoSaldo)
                    fraSinLote.Visible = True
                    txtCantidad.Text = ""
                    txtCantidad.SetFocus
                End If
            End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub

'***************daniel barrantes**************
'***************Verifica si YA SE REGISTRO el ITEM (al momento de registrar)
'***************
Function ItemYaExiste(lnIdProducto As Long, lcLote As String, ldFechaVencimiento As Date, lnIdTipoSalidaBienInsumoSaldo As Long) As Boolean
        Dim lbExiste As Boolean
        Dim oRsTmp As New ADODB.Recordset
        Set oRsTmp = mRs_Productos.Clone
        ItemYaExiste = False
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
                If oRsTmp.Fields!idProducto = lnIdProducto And Trim(oRsTmp.Fields!Lote) = Trim(lcLote) And oRsTmp.Fields!FechaVencimiento = ldFechaVencimiento And oRsTmp.Fields!idTipoSalidaBienInsumo = lnIdTipoSalidaBienInsumoSaldo Then
                   ItemYaExiste = True
                   MsgBox "Este Producto/Tipo/Lote/FechaVencimiento ya está registrado", vbInformation, "Farmacia"
                   Exit Do
                End If
                oRsTmp.MoveNext
           Loop
        End If
        oRsTmp.Close
End Function

Function ItemYaExisteSinLote(lnIdProducto As Long, lnIdTipoSalidaBienInsumoSaldo As Long) As Boolean
        Dim lbExiste As Boolean
        Dim oRsTmp As New ADODB.Recordset
        Set oRsTmp = mRs_Productos.Clone
        ItemYaExisteSinLote = False
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
                If oRsTmp.Fields!idProducto = lnIdProducto And oRsTmp.Fields!idTipoSalidaBienInsumo = lnIdTipoSalidaBienInsumoSaldo Then
                   ItemYaExisteSinLote = True
                   MsgBox "Este Producto/Tipo ya está registrado", vbInformation, "Farmacia"
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
            Case "codigo"

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
                SendKeys "{Tab}"
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
            If ml_idTipoPrecioParaNiNs = 0 Then
               MsgBox "Debe elegir el Concepto antes de Registrar Productos", vbInformation, "Farmacia"
               Exit Sub
            End If
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
                Dim oConexion As New ADODB.Connection
                Dim oFarmMovimientoDetalle As New farmMovimientoDetalle
                oConexion.Open SIGHEntidades.CadenaConexion
                oConexion.CursorLocation = adUseClient
                Set oFarmMovimientoDetalle.Conexion = oConexion
                If mb_MuestraLoteParaDespachoNS = True Then
                   'Set rs = oFarmMovimientoDetalle.FarmDevuelveSaldosConLotesSegunAlmacen(ml_IdAlmacen, 1, sNombre)
                   Set rs = oFarmMovimientoDetalle.FarmDevuelveSaldosConLotesSegunAlmacenCliente(ml_IdAlmacen, 1, sNombre)
                Else
                   Set rs = mo_ReglasFarmacia.FarmDevuelveSaldosSinLotesSegunAlmacen(ml_IdAlmacen, 1, sNombre)
                End If
                Set grillaBusqueda.DataSource = rs
                If mRs_Productos.RecordCount < 7 Then
                   grillaBusqueda.Top = oGrilla.ActiveCell.GetUIElement.Rect.Bottom * Screen.TwipsPerPixelY
                Else
                   grillaBusqueda.Top = 0
                End If
                grillaBusqueda.Visible = True
                grillaBusqueda.Enabled = True
                oConexion.Close
                Set oConexion = Nothing
                Set oFarmMovimientoDetalle = Nothing
            End Select
        End If

End Sub


Sub GenerarRecordsetProductos()
    With mRs_Productos
          .Fields.Append "IdProducto", adInteger, 4
          .Fields.Append "Codigo", adVarChar, 20
          .Fields.Append "NombreProducto", adChar, 300
          .Fields.Append "idTipoSalidaBienInsumo", adInteger
          .Fields.Append "Lote", adVarChar, 15
          .Fields.Append "FechaVencimiento", adDate, , adFldIsNullable
          .Fields.Append "Saldo", adInteger
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Precio", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "RegistroSanitario", adVarChar, 50, adFldIsNullable
          .Fields.Append "NumeroDocumento", adVarChar, 20, adFldIsNullable
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    'Set grdProductos.DataSource = mRs_Productos
End Sub

Private Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
     On Error GoTo ConfigEstilo
     grdProductos.Bands(0).Columns("NumeroDocumento").Hidden = True
     grdProductos.Bands(0).Columns("IdProducto").Hidden = True
     grdProductos.Bands(0).Columns("IdProducto").Activation = ssActivationActivateNoEdit
     
     grdProductos.Bands(0).Columns("codigo").Width = 1000
     grdProductos.Bands(0).Columns("NombreProducto").Width = 7000
     '
     grdProductos.ValueLists.Add "TipoSalida"
     grdProductos.ValueLists("TipoSalida").ValueListItems.Add 1, "Ventas"
     grdProductos.ValueLists("TipoSalida").ValueListItems.Add 2, "Interv.Sanit"
     grdProductos.Bands(0).Columns("idTipoSalidaBienInsumo").ValueList = "TipoSalida"
     grdProductos.Bands(0).Columns("idTipoSalidaBienInsumo").Style = ssStyleDropDownList
     grdProductos.Bands(0).Columns("idTipoSalidaBienInsumo").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("idTipoSalidaBienInsumo").Width = 800
     grdProductos.Bands(0).Columns("idTipoSalidaBienInsumo").Header.Caption = "Tipo"
     '
     grdProductos.Bands(0).Columns("Lote").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("FechaVencimiento").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("FechaVencimiento").Width = 1000
     grdProductos.Bands(0).Columns("saldo").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("saldo").Width = 800
     grdProductos.Bands(0).Columns("cantidad").Width = 800
     grdProductos.Bands(0).Columns("cantidad").Format = "###0"
     grdProductos.Bands(0).Columns("Precio").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("Precio").Width = 700
     grdProductos.Bands(0).Columns("Precio").Format = "#0.000"
     grdProductos.Bands(0).Columns("Total").Activation = ssActivationActivateNoEdit
     grdProductos.Bands(0).Columns("Total").Format = "#0.00"
    
ConfigEstilo:
    gridInfra.ConfigurarFilasBiColores oGrilla, SIGHEntidades.GrillaConFilasBicolor
    
End Sub





Private Sub grdProductosLotes_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
        If Cell.Column.Key = "Cantidad" Then
           If Cell.Row.Cells("Cantidad").Value > Cell.Row.Cells("Saldo").Value Then
              MsgBox "La CANTIDAD no puede ser mayor al SALDO", vbInformation, ""
              Cell.Row.Cells("Cantidad").Value = 0
           End If
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
    oGrilla.Bands(0).Columns("Nombre").Width = 7500
    oGrilla.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
    If mb_MuestraLoteParaDespachoNS = True Then
        oGrilla.Bands(0).Columns("Lote").Width = 1500
        oGrilla.Bands(0).Columns("Lote").Activation = ssActivationActivateNoEdit
        oGrilla.Bands(0).Columns("FechaVencimiento").Width = 1000
        oGrilla.Bands(0).Columns("FechaVencimiento").Activation = ssActivationActivateNoEdit
    End If
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
    If mb_MuestraLoteParaDespachoNS = True Then
        'Se mostro productos con Lotes y se eligio un solo LOTE
        If ItemYaExiste(lnIdProductoBusqueda, grillaBusqueda.ActiveRow.Cells("lote").Value, grillaBusqueda.ActiveRow.Cells("fechaVencimiento").Value, grillaBusqueda.ActiveRow.Cells("IdTipoSalidaBienInsumoSaldo").Value) Then
            grdProductos.ActiveRow.Cells("codigo").Value = ""
            grdProductos.ActiveRow.Cells("idproducto").Value = 0
            grdProductos.ActiveRow.Cells("NombreProducto").Value = ""
            grdProductos.ActiveRow.Cells("lote").Value = ""
            grdProductos.ActiveRow.Cells("FechaVencimiento").Value = ""
            grdProductos.ActiveRow.Cells("saldo").Value = 0
            grdProductos.ActiveRow.Cells("precio").Value = 0
        Else
            RefrescarDatos
            Set grillaBusqueda.DataSource = Nothing
            grillaBusqueda.Visible = False
            SendKeys "{Tab}"
            SendKeys "{Tab}"
            SendKeys "{Tab}"
            SendKeys "{Tab}"
        End If
    Else
        'Se mostro productos Sin Lotes y se eligio
        If ItemYaExisteSinLote(lnIdProductoBusqueda, grillaBusqueda.ActiveRow.Cells("IdTipoSalidaBienInsumoSaldo").Value) Then
             grdProductos.ActiveRow.Cells("codigo").Value = ""
             grdProductos.ActiveRow.Cells("idproducto").Value = 0
             grdProductos.ActiveRow.Cells("NombreProducto").Value = ""
             grdProductos.ActiveRow.Cells("lote").Value = ""
             grdProductos.ActiveRow.Cells("FechaVencimiento").Value = ""
             grdProductos.ActiveRow.Cells("saldo").Value = 0
             grdProductos.ActiveRow.Cells("precio").Value = 0
        Else
             ml_idProducto = grillaBusqueda.ActiveRow.Cells("idProducto").Value
             Set oRsSaldosEnEsteMomento = mo_ReglasFarmacia.FarmDevuelveSaldosConLotesSegunAlmacen(ml_IdAlmacen, 0, grillaBusqueda.ActiveRow.Cells("codigo").Value)
             oRsSaldosEnEsteMomento.Filter = "idTipoSalidaBienInsumoSaldo=" & grillaBusqueda.ActiveRow.Cells("idTipoSalidaBienInsumoSaldo").Value
             txtCodigo.Text = grillaBusqueda.ActiveRow.Cells("codigo").Value
             txtNombre.Text = grillaBusqueda.ActiveRow.Cells("nombre").Value
             txtSaldo.Text = grillaBusqueda.ActiveRow.Cells("saldo").Value
             txtEsEstrategico.Text = SIGHEntidades.ElijeSiEsEstrategicoDevuelveNombre(grillaBusqueda.ActiveRow.Cells("idTipoSalidaBienInsumoSaldo").Value)
             Set grillaBusqueda.DataSource = Nothing
             grillaBusqueda.Visible = False
             fraSinLote.Visible = True
             txtCantidad.Text = ""
             txtCantidad.SetFocus
             'debb-06/03/2018
             If fraSinLote.Top = lnTopGrilla Then
                
                If oRsSaldosEnEsteMoment1.State = 1 Then Set oRsSaldosEnEsteMoment1 = Nothing
                With oRsSaldosEnEsteMoment1
                    .Fields.Append "IdProducto", adInteger
                    .Fields.Append "tipo", adVarChar, 50, adFldIsNullable
                    .Fields.Append "IdAlmacen", adInteger
                    .Fields.Append "Precio", adDouble
                    .Fields.Append "idTipoSalidaBienInsumoSaldo", adInteger
                    .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
                    .Fields.Append "Nombre", adVarChar, 300, adFldIsNullable
                    .Fields.Append "Saldo", adInteger
                    .Fields.Append "Lote", adVarChar, 20, adFldIsNullable
                    .Fields.Append "FechaVencimiento", adDate
                    .Fields.Append "Cantidad", adInteger
                    .CursorType = adOpenKeyset
                    .LockType = adLockOptimistic
                    .Open
                End With
                If oRsSaldosEnEsteMomento.RecordCount > 0 Then
                    oRsSaldosEnEsteMomento.MoveFirst
                    Do While Not oRsSaldosEnEsteMomento.EOF
                       oRsSaldosEnEsteMoment1.AddNew
                       oRsSaldosEnEsteMoment1!idProducto = oRsSaldosEnEsteMomento!idProducto
                       oRsSaldosEnEsteMoment1!tipo = oRsSaldosEnEsteMomento!tipo
                       oRsSaldosEnEsteMoment1!IdAlmacen = oRsSaldosEnEsteMomento!IdAlmacen
                       oRsSaldosEnEsteMoment1!Precio = oRsSaldosEnEsteMomento!Precio
                       oRsSaldosEnEsteMoment1!idTipoSalidaBienInsumoSaldo = oRsSaldosEnEsteMomento!idTipoSalidaBienInsumoSaldo
                       oRsSaldosEnEsteMoment1!codigo = oRsSaldosEnEsteMomento!codigo
                       oRsSaldosEnEsteMoment1!Nombre = oRsSaldosEnEsteMomento!Nombre
                       oRsSaldosEnEsteMoment1!saldo = oRsSaldosEnEsteMomento!saldo
                       oRsSaldosEnEsteMoment1!Lote = oRsSaldosEnEsteMomento!Lote
                       oRsSaldosEnEsteMoment1!FechaVencimiento = oRsSaldosEnEsteMomento!FechaVencimiento
                       oRsSaldosEnEsteMoment1!Cantidad = 0
                       oRsSaldosEnEsteMoment1.Update
                       oRsSaldosEnEsteMomento.MoveNext
                    Loop
                   oRsSaldosEnEsteMoment1.MoveFirst
                End If
                Set grdProductosLotes.DataSource = oRsSaldosEnEsteMoment1
                
                grdProductosLotes.Bands(0).Columns("idProducto").Hidden = True
                grdProductosLotes.Bands(0).Columns("tipo").Hidden = True
                grdProductosLotes.Bands(0).Columns("idAlmacen").Hidden = True
                grdProductosLotes.Bands(0).Columns("Precio").Hidden = True
                grdProductosLotes.Bands(0).Columns("idProducto").Hidden = True
                grdProductosLotes.Bands(0).Columns("idTipoSalidaBienInsumoSaldo").Hidden = True
                grdProductosLotes.Bands(0).Columns("codigo").Width = 800
                grdProductosLotes.Bands(0).Columns("codigo").Activation = ssActivationActivateNoEdit
                grdProductosLotes.Bands(0).Columns("nombre").Width = 4000
                grdProductosLotes.Bands(0).Columns("nombre").Activation = ssActivationActivateNoEdit
                grdProductosLotes.Bands(0).Columns("saldo").Width = 800
                grdProductosLotes.Bands(0).Columns("saldo").Activation = ssActivationActivateNoEdit
                grdProductosLotes.Bands(0).Columns("lote").Width = 2000
                grdProductosLotes.Bands(0).Columns("lote").Activation = ssActivationActivateNoEdit
                grdProductosLotes.Bands(0).Columns("fechaVencimiento").Width = 1000
                grdProductosLotes.Bands(0).Columns("fechaVencimiento").Activation = ssActivationActivateNoEdit
                grdProductosLotes.Bands(0).Columns("cantidad").Width = 800
             End If
             '
        End If
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
               grdProductos.ActiveRow.Cells("lote").Value = grillaBusqueda.ActiveRow.Cells("lote").Value
               grdProductos.ActiveRow.Cells("fechaVencimiento").Value = grillaBusqueda.ActiveRow.Cells("fechaVencimiento").Value
               grdProductos.ActiveRow.Cells("Total").Value = 0
               grdProductos.ActiveRow.Cells("cantidad").Value = 0
               grdProductos.ActiveRow.Cells("idTipoSalidaBienInsumo").Value = grillaBusqueda.ActiveRow.Cells("IdTipoSalidaBienInsumoSaldo").Value
               SIGHEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "+Cod:" & Trim(grillaBusqueda.ActiveRow.Cells("CODIGO").Value)
               Totalizar
    End If

End Sub




Function AgregaProductosDesdeConsolidadoSinLote(lnIdProducto As Long, lcCodigo As String, lcNombre As String) As Boolean
    
    Dim lnCantidadCargar As Long
    Dim lnCantSaldo As Long
    Dim lnPrecioUnitario As Double
    Set oRsSaldosEnEsteMomento = mo_ReglasFarmacia.FarmDevuelveSaldosConLotesSegunAlmacen(ml_IdAlmacen, 0, lcCodigo)
    AgregaProductosDesdeConsolidadoSinLote = False
    If oRsSaldosEnEsteMomento.RecordCount > 0 Then
       lnPrecioUnitario = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(lnIdProducto, ml_idTipoPrecioParaNiNs)
       lnCantSaldo = grillaBusqueda.ActiveRow.Cells("saldo").Value
       oRsSaldosEnEsteMomento.MoveFirst
       Do While Not oRsSaldosEnEsteMomento.EOF
          If oRsSaldosEnEsteMomento.Fields!saldo > 0 And oRsSaldosEnEsteMomento.Fields!FechaVencimiento >= LdFechaMinimaDespacho Then
               AgregaProductosDesdeConsolidadoSinLote = True
               grdProductos.ActiveRow.Cells("codigo").Value = lcCodigo
               grdProductos.ActiveRow.Cells("idproducto").Value = lnIdProducto
               grdProductos.ActiveRow.Cells("NombreProducto").Value = lcNombre
               grdProductos.ActiveRow.Cells("precio").Value = lnPrecioUnitario
               grdProductos.ActiveRow.Cells("lote").Value = oRsSaldosEnEsteMomento.Fields!Lote
               grdProductos.ActiveRow.Cells("fechaVencimiento").Value = oRsSaldosEnEsteMomento.Fields!FechaVencimiento
               grdProductos.ActiveRow.Cells("saldo").Value = oRsSaldosEnEsteMomento.Fields!saldo
'                        .AddNew
'                        .Fields!idProducto = oRsConsolidado.Fields!idProducto
'                        .Fields!codigo = oRsConsolidado.Fields!codigo
'                        .Fields!nombreProducto = oRsConsolidado.Fields!nombreProducto
'                        .Fields!lote = oRsSaldosEnEsteMomento.Fields!lote
'                        .Fields!fechaVencimiento = oRsSaldosEnEsteMomento.Fields!fechaVencimiento
'                        .Fields!precio = oRsConsolidado.Fields!precio
               If lnCantSaldo >= oRsSaldosEnEsteMomento.Fields!saldo Then
                  lnCantidadCargar = oRsSaldosEnEsteMomento.Fields!saldo
                  grdProductos.ActiveRow.Cells("cantidad").Value = lnCantidadCargar
'                           .Fields!cantidad = lnCantidadCargar
'                           .Fields!total = Round(lnCantidadCargar * oRsConsolidado.Fields!precio, 2)
'                           .Update
                  lnCantSaldo = lnCantSaldo - oRsSaldosEnEsteMomento.Fields!saldo
                  If lnCantSaldo <= 0 Then
                     Exit Do
                  End If
               Else
                  lnCantidadCargar = lnCantSaldo
                  grdProductos.ActiveRow.Cells("cantidad").Value = lnCantidadCargar
'                           .Fields!cantidad = lnCantidadCargar
'                           .Fields!total = Round(lnCantidadCargar * oRsConsolidado.Fields!precio, 2)
'                           .Update
                  Exit Do
               End If
               AgregaRegistro
          End If
          oRsSaldosEnEsteMomento.MoveNext
       Loop
    End If
    oRsSaldosEnEsteMomento.Close
    'Set oRsSaldosEnEsteMomento = Nothing
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





Private Sub txtBusca_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If TxtBusca.Text <> "" Then
            TxtBusca.Text = Trim(TxtBusca.Text)
            mRs_Productos.MoveFirst
            If cmbOrden.ListIndex = 0 Then
               mRs_Productos.Find "codigo='" & TxtBusca.Text & "'"
            Else
               Do While Not mRs_Productos.EOF
                  If Left(mRs_Productos!nombreProducto, Len(TxtBusca.Text)) = UCase(TxtBusca.Text) Then
                     Exit Do
                  End If
                  mRs_Productos.MoveNext
               Loop
            End If
            grdProductos.Refresh
      End If
   End If
End Sub



Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       btnAgregar.SetFocus
    End If
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   grdProductos.Top = 0
   grdProductos.Left = 0
   grdProductos.Width = UserControl.Width
   'grdProductos.Height = UserControl.Height - UserControl.Label1.Height - 100
   grdProductos.Height = UserControl.Height - UserControl.Label1.Height - 600
   fraSinLote.Top = UserControl.Height - UserControl.Label1.Height - 600
   'fraSinLote.Visible = True
   Label1.Top = UserControl.Height - UserControl.Label1.Height - 50
   Label2.Top = UserControl.Height - UserControl.Label1.Height - 50
   cmbOrden.Top = UserControl.Height - UserControl.Label1.Height - 100
   TxtBusca.Top = UserControl.Height - UserControl.Label1.Height - 100
   lblPrecios.Top = UserControl.Height - UserControl.Label1.Height - 100
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



