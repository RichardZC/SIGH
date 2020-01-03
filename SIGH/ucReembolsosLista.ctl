VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucReembolsosLista 
   ClientHeight    =   6525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11130
   ScaleHeight     =   6525
   ScaleWidth      =   11130
   Begin VB.Frame fraBusqueda 
      Caption         =   "Busqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   90
      TabIndex        =   4
      Top             =   555
      Width           =   10965
      Begin VB.CommandButton cmdBuscarPorApell 
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
         Left            =   6465
         Picture         =   "ucReembolsosLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Width           =   315
      End
      Begin VB.CommandButton bntReporte 
         Height          =   420
         Left            =   8220
         Picture         =   "ucReembolsosLista.ctx":058A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   840
         Width           =   1305
      End
      Begin VB.CheckBox chkDifImportes 
         Caption         =   "Solo mostrar 'Importe Reembolsado total' <> Suma Importes reembolsados x Cuentas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   930
         Width           =   7665
      End
      Begin VB.TextBox txtNcuenta 
         Height          =   315
         Left            =   5070
         MaxLength       =   9
         TabIndex        =   7
         Top             =   480
         Width           =   1365
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   8220
         Picture         =   "ucReembolsosLista.ctx":0A63
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   450
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9585
         Picture         =   "ucReembolsosLista.ctx":36AC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         Width           =   1275
      End
      Begin MSDataListLib.DataCombo cmbAlmacen 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "Facturacion"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Area Tramita Seguro                                                       N° Cuenta"
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
         Left            =   165
         TabIndex        =   5
         Top             =   225
         Width           =   6855
      End
   End
   Begin UltraGrid.SSUltraGrid grdLista 
      Height          =   4470
      Left            =   90
      TabIndex        =   3
      Top             =   1980
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   7885
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
      Caption         =   $"ucReembolsosLista.ctx":6288
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Reembolsos"
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
      Left            =   60
      TabIndex        =   6
      Top             =   45
      Width           =   11010
   End
End
Attribute VB_Name = "ucReembolsosLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Reembolsos Registrados
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoBusqueda As sghTipoBusquedaPrestamoHistoria
Dim mo_Teclado As New sighentidades.Teclado
Dim oRsAlmacenes As New ADODB.Recordset
Dim oRsBusqueda As New ADODB.Recordset
Dim ml_idUsuario As Long
Dim ml_lnHwnd As Long

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdLista.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdLista.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property

Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property
Property Let TipoBusqueda(lValue As sghTipoBusquedaPrestamoHistoria)
    ml_TipoBusqueda = lValue
End Property
Property Get TipoBusqueda() As sghTipoBusquedaPrestamoHistoria
    TipoBusqueda = ml_TipoBusqueda
End Property

Property Let lnHWnd(lValue As Long)
    ml_lnHwnd = lValue
End Property


Private Sub bntReporte_Click()
    'grdLista.PrintPreview True
    Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    mo_ReglasReportes.ExportarRecordSetAexcel oRsBusqueda, "Reembolsos", "", "", ml_lnHwnd
End Sub

Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
        If cmbAlmacen.Text = "" Then
           MsgBox "Por favor elija el Area que Tramita el Seguro", vbInformation, "Busqueda"
           Exit Sub
        End If
        Set oRsBusqueda = mo_ReglasFacturacion.FactReembolsosSelecionarPorAreaTramitaSegurosDEBB(Val(cmbAlmacen.BoundText), _
                                                          Val(txtNcuenta.Text), IIf(chkDifImportes.Value = 1, True, False))
        Set grdLista.DataSource = oRsBusqueda
       ' mo_Apariencia.ConfigurarFilasBiColores grdLista, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
    txtNcuenta.Text = ""
End Sub



Private Sub cmbAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacen
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdBuscarPorApell_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New doPaciente
    Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            Dim oRsTmp As New Recordset
            Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(oDOPaciente.idPaciente, oConexion, True)
            If oRsTmp.RecordCount > 0 Then
               txtNcuenta.Text = oRsTmp.Fields!idCuentaAtencion
            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
            txtNcuenta_LostFocus
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub


Private Sub grdLista_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdLista.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("idFactReembolso")
    
End Sub

Private Sub grdLista_Click()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdLista.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("idFactReembolso")
    
End Sub


Private Sub grdLista_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    If Val(txtNcuenta.Text) > 0 Then
       grdLista.Bands(0).Columns("idCuentaAtencion").Hidden = True
    End If
    grdLista.Bands(0).Columns("idFactReembolso").Header.Caption = "Id"
    grdLista.Bands(0).Columns("idFactReembolso").Width = 1000
    grdLista.Bands(0).Columns("Anio").Header.Caption = "Año"
    grdLista.Bands(0).Columns("Anio").Width = 600
    grdLista.Bands(0).Columns("Mes").Header.Caption = "Mes"
    grdLista.Bands(0).Columns("Mes").Width = 500
    grdLista.Bands(0).Columns("Plan").Header.Caption = "Fuente Financiamiento/IAFA"
    grdLista.Bands(0).Columns("Plan").Width = 2300
    grdLista.Bands(0).Columns("NroSerie").Header.Caption = "N°Serie"
    grdLista.Bands(0).Columns("NroSerie").Width = 800
    grdLista.Bands(0).Columns("nroDocumento").Header.Caption = "N°Documento"
    grdLista.Bands(0).Columns("nroDocumento").Width = 1800
    grdLista.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdLista.Bands(0).Columns("Descripcion").Width = 2000
    grdLista.Bands(0).Columns("ConsumoPorReembolsar").Header.Caption = "Consumo"
    grdLista.Bands(0).Columns("ConsumoPorReembolsar").Width = 1300
    grdLista.Bands(0).Columns("ConsumoPorReembolsar").Format = "#0.00"
    grdLista.Bands(0).Columns("ReembolsoPagado").Header.Caption = "Reembolso"
    grdLista.Bands(0).Columns("ReembolsoPagado").Width = 1300
    grdLista.Bands(0).Columns("ReembolsoPagado").Format = "#0.00"
    grdLista.Bands(0).Columns("idEstadoReembolso").Hidden = True
    'debb-24/05/2011
    grdLista.Bands(0).Columns("GrabaDefinitivamente").Hidden = True
End Sub







Private Sub grdLista_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
        'debb-24/05/2011
        If Val(Row.Cells("idEstadoReembolso").GetText()) = 0 Then
            Row.Appearance.ForeColor = vbRed
        ElseIf Row.Cells("GrabaDefinitivamente").GetText() <> True Then
            Row.Appearance.ForeColor = vbGreen
        End If
End Sub


Private Sub txtNcuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtNcuenta_LostFocus
    End If
End Sub

Private Sub txtNcuenta_LostFocus()
    If Val(txtNcuenta.Text) > 0 Then
       btnBuscar_Click
    End If
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdLista.Width = fraBusqueda.Width
   grdLista.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub



Sub CargaComboBox()
        On Error GoTo ErrFarm
        Dim oConexion As New ADODB.Connection
        Dim rsIdAlmacen As Recordset
        Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
        Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAreaTramitaSeguros, ml_idUsuario)
        Set oBuscaDondeLabora = Nothing
        oConexion.Open sighentidades.CadenaConexion
        Set oRsAlmacenes = mo_ReglasFacturacion.AreaTramitaSegurosDevuelveTodosSegunFiltro("")
        Set cmbAlmacen.RowSource = oRsAlmacenes
        cmbAlmacen.ListField = "descripcion"
        cmbAlmacen.BoundColumn = "idAreaTramitaSeguros"
        If rsIdAlmacen.RecordCount > 0 Then
           cmbAlmacen.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
           cmbAlmacen.Enabled = False
        End If
ErrFarm:
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdLista, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdLista, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub
Sub inicializar()
    SkinConfigura
    CargaComboBox
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

