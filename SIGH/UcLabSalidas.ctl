VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl UcLabSalidas 
   ClientHeight    =   8055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10935
   ScaleHeight     =   8055
   ScaleWidth      =   10935
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   60
      TabIndex        =   7
      Top             =   495
      Width           =   10830
      Begin VB.TextBox txtNroMovimiento 
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
         Left            =   180
         MaxLength       =   9
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   510
         Width           =   1455
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9420
         Picture         =   "UcLabSalidas.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   540
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   9420
         Picture         =   "UcLabSalidas.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   1305
      End
      Begin VB.ComboBox cmbIdPtoCarga 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   510
         Width           =   2085
      End
      Begin MSMask.MaskEdBox txtFinicio 
         Height          =   315
         Left            =   3840
         TabIndex        =   2
         Top             =   510
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFfinal 
         Height          =   315
         Left            =   5250
         TabIndex        =   3
         Top             =   510
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Movimiento    Punto de Carga               Fecha Inicio        Fecha Final"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   8
         Top             =   240
         Width           =   6735
      End
   End
   Begin UltraGrid.SSUltraGrid grdListaOrdenes 
      Height          =   6510
      Left            =   60
      TabIndex        =   6
      Top             =   1500
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   11483
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
      Caption         =   "LISTA DE MOVIMIENTOS"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Laboratorio - Salidas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10875
   End
End
Attribute VB_Name = "UcLabSalidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Salidas de insumos de Laboratorio
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim ml_idRegistroSeleccionado As Long
Dim ml_PuntoCarga As sghTipoFiltroPacientes
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim ml_idUsuario As Long
Dim mo_cmbIdPuntoCarga As New sighentidades.ListaDespleglable
Dim ml_IdTipoFinanciamiento As Long
Dim oRsFarmacias As New ADODB.Recordset
Dim oRsLista As New Recordset

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdListaOrdenes.DataSource = oValue
End Property

Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdListaOrdenes.DataSource
End Property

Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property

Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property

Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property

Property Get idUsuario() As Long
    idUsuario = ml_idUsuario
End Property

Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property

Property Get Titulo() As String
    Titulo = lblNombre
End Property

Property Let PuntoCarga(lValue As Long)
    ml_PuntoCarga = lValue
    mo_cmbIdPuntoCarga.BoundText = ml_PuntoCarga
End Property

Property Get PuntoCarga() As Long
    PuntoCarga = ml_PuntoCarga
End Property

Property Let HabilitarPuntoCarga(lValue As Long)
    cmbIdPtoCarga.Enabled = lValue
End Property

Property Get HabilitarPuntoCarga() As Long
    HabilitarPuntoCarga = cmbIdPtoCarga.Enabled
End Property

Property Let idTipoFinanciamiento(lValue As Long)
    ml_IdTipoFinanciamiento = lValue
End Property

Property Get idTipoFinanciamiento() As Long
    idTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property

Private Sub btnBuscar_Click()
    If UserControl.txtFinicio.Text = sighentidades.FECHA_VACIA_DMY Then Exit Sub
    If UserControl.txtFfinal.Text = sighentidades.FECHA_VACIA_DMY Then Exit Sub
    If CDate(UserControl.txtFinicio.Text) > CDate(UserControl.txtFfinal.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
  If Not (IsDate(txtFinicio.Text)) Or Not (IsDate(txtFfinal.Text)) Or mo_cmbIdPuntoCarga.BoundText = "" Then Exit Sub
  Dim ldFechaIni As Date
  Dim ldFechaFin As Date
  Dim lcFiltro As String
  'If mo_cmbIdPuntoCarga.BoundText = "" Then Exit Sub
  If mo_cmbIdPuntoCarga.BoundText = "" Then
    MsgBox "Por favor elija el filtro PUNTO DE CARGA", vbInformation, "Filtro de Busqueda"
    cmbIdPtoCarga.SetFocus
    Exit Sub
  End If
  ldFechaIni = Format(txtFinicio.Text & " 00:01", sighentidades.DevuelveFechaSoloFormato_DMY_HM)
  ldFechaFin = Format(txtFfinal.Text & " 23:59", sighentidades.DevuelveFechaSoloFormato_DMY_HM)
  lcFiltro = ""
  If txtNroMovimiento.Text <> "" Then lcFiltro = lcFiltro & "idMovimiento=" & txtNroMovimiento.Text
  Set oRsLista = mo_ReglasLaboratorio.LabMovimientoSeleccionarPorFechasPuntoCargaSalidas(Val(mo_cmbIdPuntoCarga.BoundText), ldFechaIni, ldFechaFin)
  If lcFiltro <> "" Then oRsLista.Filter = lcFiltro
  Set grdListaOrdenes.DataSource = oRsLista
  If mo_ReglasLaboratorio.MensajeError <> "" Then MsgBox mo_ReglasLaboratorio.MensajeError, vbInformation, lblNombre.Caption
  mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenes, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
  LimpiarFiltro
End Sub

Public Sub LimpiarFiltro()
  UserControl.txtNroMovimiento = ""
End Sub

Private Sub cmbIdPtoCarga_Click()
  PuntoCarga = Val(mo_cmbIdPuntoCarga.BoundText)
  btnBuscar_Click
End Sub

Private Sub cmbIdPtoCarga_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdPtoCarga
    AdministrarKeyPreview KeyCode
End Sub

Private Sub grdListaOrdenes_AfterRowActivate()
  Dim rsRecordset As ADODB.Recordset
  Set rsRecordset = grdListaOrdenes.DataSource
  On Error Resume Next
  ml_idRegistroSeleccionado = rsRecordset("IdMovimiento")
End Sub

Private Sub grdListaOrdenes_Click()
  Dim rsRecordset As ADODB.Recordset
  Set rsRecordset = grdListaOrdenes.DataSource
  On Error Resume Next
  ml_idRegistroSeleccionado = rsRecordset("IdMovimiento")
End Sub

Private Sub grdListaOrdenes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
  grdListaOrdenes.Bands(0).Columns("IdLabEstado").Hidden = True
  grdListaOrdenes.Bands(0).Columns("IdMovimiento").Header.Caption = "N° Movimiento"
  grdListaOrdenes.Bands(0).Columns("IdMovimiento").Width = 800
  grdListaOrdenes.Bands(0).Columns("Fecha").Header.Caption = "Fecha"
  grdListaOrdenes.Bands(0).Columns("Fecha").Width = 2500
  grdListaOrdenes.Bands(0).Columns("Motivo").Header.Caption = "Motivo de Salida"
  grdListaOrdenes.Bands(0).Columns("Motivo").Width = 4200
  grdListaOrdenes.Bands(0).Columns("Estado").Header.Caption = "Estado"
  grdListaOrdenes.Bands(0).Columns("Estado").Width = 2500
End Sub

Private Sub grdListaOrdenes_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
  If Val(Row.Cells("IdLabEstado").GetText()) = 0 Then Row.Appearance.ForeColor = vbRed            'Row.Appearance.Font.Strikethrough = True
End Sub

Private Sub txtFfinal_Change()
  btnBuscar_Click
End Sub

Private Sub txtFfinal_LostFocus()
If Not esfecha(txtFfinal.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFfinal.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFinicio_Change()
  btnBuscar_Click
End Sub



Private Sub txtFinicio_LostFocus()
If Not esfecha(txtFinicio.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFinicio.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtNroMovimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroMovimiento
    AdministrarKeyPreview KeyCode
End Sub

Private Sub UserControl_Initialize()
  PuntoCarga = 0
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
    Select Case KeyCode
    Case vbKeyEscape
    Case vbKeyF2
    Case vbKeyF3
     Case vbKeyF4
     Case vbKeyF5
     Case vbKeyF6
        btnBuscar_Click
     Case vbKeyF7
        btnLimpiar_Click
     Case vbKeyF8
    End Select
End Sub

Private Sub UserControl_Resize()
  On Error Resume Next
  fraBusqueda.Width = UserControl.Width - 110
  lblNombre.Width = UserControl.Width
  grdListaOrdenes.Width = fraBusqueda.Width
  grdListaOrdenes.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
End Sub

Sub inicializar()
  ConfigurarPuntosDeCarga
  txtFinicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual
  txtFfinal.Text = Date
End Sub

Sub ConfigurarPuntosDeCarga()
  Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPtoCarga
  mo_cmbIdPuntoCarga.ListField = "Descripcion"
  mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
  Set mo_cmbIdPuntoCarga.RowSource = mo_reglasComunes.SeleccionarPuntosDeCargaSegunFiltro("idUPS=2 or idUPS=3 or idUPS=4")
  If cmbIdPtoCarga.ListCount > 0 Then cmbIdPtoCarga.ListIndex = 0
    
  Dim rsAlmacen As Recordset
  Set rsAlmacen = mo_reglasComunes.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghLaboratorio, ml_idUsuario)
  If rsAlmacen.RecordCount > 0 Then
    mo_cmbIdPuntoCarga.BoundText = rsAlmacen.Fields!idLaboraSubArea
    cmbIdPtoCarga.Enabled = False
  End If
End Sub

