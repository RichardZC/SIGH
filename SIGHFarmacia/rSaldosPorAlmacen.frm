VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form rSaldosPorAlmacen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saldos"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rSaldosPorAlmacen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   13620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmdAlmacen1 
      Height          =   330
      Left            =   1605
      TabIndex        =   21
      Top             =   75
      Width           =   4080
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   330
      ItemData        =   "rSaldosPorAlmacen.frx":0CCA
      Left            =   795
      List            =   "rSaldosPorAlmacen.frx":0CD4
      TabIndex        =   16
      Top             =   3780
      Width           =   1815
   End
   Begin VB.TextBox TxtBusca 
      Height          =   330
      Left            =   2595
      TabIndex        =   15
      Top             =   3780
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   7230
      TabIndex        =   6
      Top             =   6675
      Width           =   6330
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rSaldosPorAlmacen.frx":0CF0
         DownPicture     =   "rSaldosPorAlmacen.frx":11B4
         Height          =   700
         Left            =   3278
         Picture         =   "rSaldosPorAlmacen.frx":16A0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rSaldosPorAlmacen.frx":1B8C
         DownPicture     =   "rSaldosPorAlmacen.frx":1FEC
         Height          =   700
         Left            =   1740
         Picture         =   "rSaldosPorAlmacen.frx":2461
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame fraDatosHistoria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Left            =   7260
      TabIndex        =   3
      Top             =   3765
      Width           =   6345
      Begin VB.CheckBox chkSaldosMayorAcero 
         Caption         =   "Solo saldos mayores a CERO"
         Height          =   315
         Left            =   210
         Picture         =   "rSaldosPorAlmacen.frx":28D6
         TabIndex        =   23
         Top             =   2550
         Width           =   4665
      End
      Begin VB.ComboBox cmbTipoProducto 
         Height          =   330
         ItemData        =   "rSaldosPorAlmacen.frx":2BE8
         Left            =   1395
         List            =   "rSaldosPorAlmacen.frx":2BF2
         TabIndex        =   19
         Top             =   2235
         Width           =   4800
      End
      Begin VB.CheckBox chkSaldoMayorAcero 
         Caption         =   "Sólo muestra Productos con Saldos MAYORES A CERO"
         Height          =   225
         Left            =   180
         TabIndex        =   18
         Top             =   1980
         Width           =   5835
      End
      Begin VB.ComboBox cmbTipoSalida 
         Height          =   330
         ItemData        =   "rSaldosPorAlmacen.frx":2C14
         Left            =   4290
         List            =   "rSaldosPorAlmacen.frx":2C16
         TabIndex        =   11
         Top             =   1110
         Width           =   1920
      End
      Begin VB.CheckBox chkTodasFarmacias 
         Caption         =   "Todas las Farmacias"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   270
         Width           =   1935
      End
      Begin VB.CheckBox chkExcel 
         Caption         =   "En Excel"
         Height          =   315
         Left            =   180
         Picture         =   "rSaldosPorAlmacen.frx":2C18
         TabIndex        =   9
         Top             =   690
         Width           =   1125
      End
      Begin VB.CheckBox chkStkMinimo 
         Caption         =   "Sólo muestra Productos con Saldos menores a su STOCK MINIMO"
         Height          =   225
         Left            =   180
         TabIndex        =   8
         Top             =   1590
         Width           =   5835
      End
      Begin VB.ComboBox cmbAlmacen 
         Height          =   330
         Left            =   2130
         TabIndex        =   0
         Top             =   240
         Width           =   4080
      End
      Begin VB.ComboBox cmbOrden 
         Height          =   330
         ItemData        =   "rSaldosPorAlmacen.frx":2F2A
         Left            =   3450
         List            =   "rSaldosPorAlmacen.frx":2F34
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   690
         Width           =   2745
      End
      Begin VB.CheckBox chkLotes 
         Caption         =   "Se muestra Lotes/F.Vencimiento"
         Height          =   225
         Left            =   180
         TabIndex        =   4
         Top             =   1170
         Width           =   2985
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Producto"
         Height          =   210
         Left            =   195
         TabIndex        =   20
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Salida"
         Height          =   210
         Left            =   3420
         TabIndex        =   12
         Top             =   1170
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Orden"
         Height          =   210
         Left            =   2910
         TabIndex        =   5
         Top             =   720
         Width           =   510
      End
   End
   Begin UltraGrid.SSUltraGrid grdSaldoItem 
      Height          =   3615
      Left            =   30
      TabIndex        =   13
      Top             =   4185
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   6376
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Saldo x Item"
   End
   Begin UltraGrid.SSUltraGrid grdSaldoTotal 
      Height          =   3300
      Left            =   45
      TabIndex        =   14
      Top             =   435
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   5821
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Saldos totales del Establecimiento"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Almacen/Farmacia"
      Height          =   210
      Left            =   60
      TabIndex        =   22
      Top             =   150
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Orden"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   75
      TabIndex        =   17
      Top             =   3780
      Width           =   615
   End
End
Attribute VB_Name = "rSaldosPorAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de Saldos por Almacén
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbAlmacen As New SIGHEntidades.ListaDespleglable
Dim mo_cmbAlmace1 As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim ms_MensajeError As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_TextoDelFiltro As String
Dim ml_idUsuario As Long
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim lcSql As String
Dim oRsStockTotal As New Recordset
Dim oRsStockXitem As New Recordset

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
         Dim oRptClase As New rCrystal
         oRptClase.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
         oRptClase.IdAlmacen = Val(mo_cmbAlmacen.BoundText)
         oRptClase.OrdenadoPor = cmbOrden.ListIndex
         oRptClase.TextoDelFiltro = ml_TextoDelFiltro
         oRptClase.SeMuestraLotes = IIf(chkLotes.Value = 1, True, False)
         oRptClase.StockMinimoMayorAcantidad = IIf(chkStkMinimo.Value = 1, True, False)
         oRptClase.TipoReporte = Me.Name
         oRptClase.idTipoSalidaBienInsumo = IIf(cmbTipoSalida.ListIndex < 0, 0, cmbTipoSalida.ListIndex)
         oRptClase.CodigoItem = IIf(Me.chkSaldoMayorAcero.Value = 1, "1", "")
         oRptClase.TipoProducto = IIf(cmbTipoProducto.Text = "", 99, cmbTipoProducto.ListIndex)
         oRptClase.SoloSaldosMayoresAcero = IIf(chkSaldosMayorAcero.Value = 1, True, False)
         oRptClase.Show vbModal
         Set oRptClase = Nothing
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    ms_MensajeError = ""
    If chkTodasFarmacias.Value = 1 Then
        ml_TextoDelFiltro = "FILTROS:   Almacén: (Todos)    orden: (" & cmbOrden.Text & ")    " & _
                            IIf(chkStkMinimo.Value = 1, "(" & chkStkMinimo.Caption & ")", "") & _
                            IIf(Me.cmbTipoSalida.Text <> "", " (Tipo Salida: " & Trim(Me.cmbTipoSalida.Text) & ")", "") & _
                            IIf(cmbTipoProducto.Text = "", "", " (" & cmbTipoProducto.Text & ")")
        mo_cmbAlmacen.BoundText = ""
    Else
        ml_TextoDelFiltro = "FILTROS:   Almacén: (" & Trim(cmbAlmacen.Text) & ")    orden: (" & cmbOrden.Text & ")    " & _
                            IIf(chkStkMinimo.Value = 1, "(" & chkStkMinimo.Caption & ")", "") & _
                            IIf(Me.cmbTipoSalida.Text <> "", " (Tipo Salida: " & Trim(Me.cmbTipoSalida.Text) & ")", "") & _
                            IIf(cmbTipoProducto.Text = "", "", " (" & cmbTipoProducto.Text & ")")
        If mo_cmbAlmacen.BoundText = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén"
            cmbAlmacen.SetFocus
        End If
    End If
    If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function

Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub








Private Sub chkTodasFarmacias_Click()
   If chkTodasFarmacias.Value = 1 Then
      cmbAlmacen.Visible = False
      chkLotes.Visible = False
      chkStkMinimo.Visible = False
      Label12.Visible = False
      cmbTipoSalida.Visible = False
      chkStkMinimo.Value = 0
      chkLotes.Value = 0
   Else
      cmbAlmacen.Visible = True
      chkLotes.Visible = True
      chkStkMinimo.Visible = True
      Label12.Visible = True
      cmbTipoSalida.Visible = True
   End If
End Sub

Private Sub cmbAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacen

End Sub



Private Sub cmbAlmacen_LostFocus()
    If Val(mo_cmbAlmacen.BoundText) > 0 Then
       Dim lcIdTipoSuministro As String
       mo_ReglasComunes.LlenaDataComboTipoSalidaBienSegunAlmacen Me.cmbTipoSalida, Val(mo_cmbAlmacen.BoundText), lcIdTipoSuministro
       On Error Resume Next
       Me.cmbTipoSalida.ListIndex = 0
    End If
End Sub


'agregado por mariano 04112014
Private Sub CmbFiltro_Click()
    If CmbFiltro.ListIndex = 0 Then
            oRsStockTotal.Sort = "Codigo"
            TxtBusca.Text = ""
        Else
            oRsStockTotal.Sort = "Nombre"
            TxtBusca.Text = ""
    End If
    grdSaldoTotal.Refresh
End Sub

Private Sub cmbOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbOrden

End Sub


Private Sub cmdAlmacen1_Click()
    If cmdAlmacen1.Text <> "" Then
       Set oRsStockTotal = mo_ReglasFarmacia.SaldoTotalPorAlmacen(Val(mo_cmbAlmace1.BoundText))
    Else
       Set oRsStockTotal = mo_ReglasFarmacia.SaldoTotalDelEstablecimiento
    End If
    Set Me.grdSaldoTotal.DataSource = oRsStockTotal
End Sub

Private Sub Form_Initialize()
    Set mo_cmbAlmacen.MiComboBox = cmbAlmacen
    Set mo_cmbAlmace1.MiComboBox = cmdAlmacen1
End Sub


Private Sub Form_Load()
    cmbOrden.ListIndex = 1
    '
    mo_cmbAlmace1.BoundColumn = "IdAlmacen"
    mo_cmbAlmace1.ListField = "Descripcion"
    Set mo_cmbAlmace1.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarTodosMenosExternos
    '
    mo_cmbAlmacen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacen.ListField = "Descripcion"
    Set mo_cmbAlmacen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarTodosMenosExternos
    '
    Dim rsIdAlmacen As Recordset
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    Set oBuscaDondeLabora = Nothing
    If rsIdAlmacen.RecordCount > 0 Then
       mo_cmbAlmacen.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacen, False
       Me.chkTodasFarmacias.Enabled = False
    End If
    '
    Set oRsStockTotal = mo_ReglasFarmacia.SaldoTotalDelEstablecimiento
    
    
    Set Me.grdSaldoTotal.DataSource = oRsStockTotal
    
    mo_Apariencia.ConfigurarFilasBiColores Me.grdSaldoTotal, SIGHEntidades.GrillaConFilasBicolor
    mo_Apariencia.ConfigurarFilasBiColores Me.grdSaldoItem, SIGHEntidades.GrillaConFilasBicolor
End Sub



Private Sub grdSaldoTotal_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    'grdSaldoTotal.Bands(0).Columns("PqteGrupo").Hidden = True
    grdSaldoTotal.Bands(0).Columns("Codigo").Width = 1000
    grdSaldoTotal.Bands(0).Columns("Nombre").Width = 10400
    grdSaldoTotal.Bands(0).Columns("SaldoTotal").Width = 1500
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
'           ucListaProductos1.RealizarBusqueda
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub


Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbAlmacen = Nothing
    Set mo_Formulario = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub grdSaldoTotal_DblClick()
    If Not IsNull(oRsStockTotal.Fields!codigo) Then
        If cmdAlmacen1.Text <> "" Then
            Set oRsStockXitem = mo_ReglasFarmacia.FarmSaldoDetalladoPorIdProductoIdAlmacen(Val(mo_cmbAlmace1.BoundText), oRsStockTotal!idProducto)
            Me.grdSaldoItem.Caption = "(" & Trim(oRsStockTotal.Fields!codigo) & ") " & oRsStockTotal.Fields!Nombre
            Set Me.grdSaldoItem.DataSource = oRsStockXitem
        Else
            Set oRsStockXitem = mo_ReglasFarmacia.SaldoDetalladoPorItemSeleccionarPorCodigo(oRsStockTotal.Fields!codigo)
            Me.grdSaldoItem.Caption = "(" & Trim(oRsStockTotal.Fields!codigo) & ") " & oRsStockTotal.Fields!Nombre
            Set Me.grdSaldoItem.DataSource = oRsStockXitem
        End If
    End If
End Sub
Private Sub grdSaldoItem_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    If cmdAlmacen1.Text <> "" Then
        grdSaldoItem.Bands(0).Columns("idAlmacen").Hidden = True
        grdSaldoItem.Bands(0).Columns("idProducto").Hidden = True
        grdSaldoItem.Bands(0).Columns("idTipoSalidaBienInsumo").Hidden = True
        grdSaldoItem.Bands(0).Columns("Precio").Hidden = True
        grdSaldoItem.Bands(0).Columns("Lote").Width = 2000
        grdSaldoItem.Bands(0).Columns("fechaVencimiento").Width = 1200
        grdSaldoItem.Bands(0).Columns("Cantidad").Width = 800
        grdSaldoItem.Bands(0).Columns("Tipo").Width = 1000
    Else
        grdSaldoItem.Bands(0).Columns("Codigo").Hidden = True
        grdSaldoItem.Bands(0).Columns("Almacen").Width = 3000
        grdSaldoItem.Bands(0).Columns("Cantidad").Width = 1000
        grdSaldoItem.Bands(0).Columns("PrecPond").Width = 1000
        grdSaldoItem.Bands(0).Columns("Importe").Width = 1500
    End If
End Sub


'agregado por mariano 04112014
Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If TxtBusca.Text <> "" Then
            TxtBusca.Text = Trim(TxtBusca.Text)
            oRsStockTotal.MoveFirst
            If CmbFiltro.ListIndex = 0 Then
               oRsStockTotal.Find "Codigo='" & TxtBusca.Text & "'"
            Else
               Do While Not oRsStockTotal.EOF
                  If Left(oRsStockTotal!Nombre, Len(TxtBusca.Text)) = UCase(TxtBusca.Text) Then
                  Exit Do
                  End If
                  oRsStockTotal.MoveNext
               Loop
            End If
            grdSaldoItem.Refresh
      End If
   End If
End Sub
