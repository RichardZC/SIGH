VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucContanciasDeAtencion 
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   ScaleHeight     =   8160
   ScaleWidth      =   10995
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
      Left            =   75
      TabIndex        =   7
      Top             =   495
      Width           =   10830
      Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
         Caption         =   "..."
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Busca Cuenta por Apellidos y Nombres"
         Top             =   465
         Width           =   315
      End
      Begin VB.TextBox txtNroHistoria 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         MaxLength       =   9
         TabIndex        =   0
         Top             =   465
         Width           =   1425
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9420
         Picture         =   "ucConstanciasDeAtencion.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   540
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   9420
         Picture         =   "ucConstanciasDeAtencion.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1305
      End
      Begin VB.ComboBox cmbFecha 
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
         ItemData        =   "ucConstanciasDeAtencion.ctx":5825
         Left            =   2280
         List            =   "ucConstanciasDeAtencion.ctx":5827
         TabIndex        =   2
         Text            =   "cmbFecha"
         Top             =   465
         Width           =   1500
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Historia clínica                 Fecha"
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
   Begin UltraGrid.SSUltraGrid grdListaConstancias 
      Height          =   5535
      Left            =   90
      TabIndex        =   5
      Top             =   1560
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   9763
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
      Caption         =   "LISTA DE CONSTANCIAS DE ATENCIÓN EMITIDAS"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Constancias de Atención"
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
      TabIndex        =   6
      Top             =   0
      Width           =   10875
   End
End
Attribute VB_Name = "ucContanciasDeAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para listar constancias
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim ml_idUsuario As Long
Dim ml_idAtencion As Long
Dim oRsLista As New Recordset
Dim mrs_FacturacionProductos As New Recordset
Dim rs As Recordset

Dim ml_idRegistroSeleccionado As Long
Dim ml_Historia As Long
Dim ml_IdPaciente As Long
Dim ml_idOrden As Long
Dim ml_idTipoConstancia As Long
Dim ml_Recibo As String
Dim ml_nombrePaciente As String
Dim ml_Observaciones As String
Dim ml_IdServicio As Long

Function ConstanciaSeleccionaPorFecha(fecha As Date, HC As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  Set ConstanciaSeleccionaPorFecha = Nothing
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "ConstanciaSeleccionaPorFecha"
    Set oParameter = .CreateParameter("@HC", adInteger, adParamInput, 0, HC): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, fecha): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set ConstanciaSeleccionaPorFecha = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Property Set DataSource(oValue As ADODB.Recordset)
  Set UserControl.grdListaConstancias.DataSource = oValue
End Property

Property Get DataSource() As ADODB.Recordset
  Set DataSource = UserControl.grdListaConstancias.DataSource
End Property

Property Let idRegistroSeleccionado(lValue As Long)
  ml_idRegistroSeleccionado = lValue
End Property

Property Get idRegistroSeleccionado() As Long
  idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property

Property Get Historia() As Long
  Historia = ml_Historia
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

Property Let idAtencion(lValue As Long)
  ml_idAtencion = lValue
End Property

Property Get idAtencion() As Long
  idAtencion = ml_idAtencion
End Property

Property Get idTipoConstancia() As Long
  idTipoConstancia = ml_idTipoConstancia
End Property

Property Get Recibo() As String
  Recibo = ml_Recibo
End Property

Property Get Observaciones() As String
  Observaciones = ml_Observaciones
End Property

Property Get IdServicio() As Long
  IdServicio = ml_IdServicio
End Property

Private Sub btnBuscar_Click()
  Screen.MousePointer = vbHourglass
  ml_idRegistroSeleccionado = 0
  ml_IdPaciente = 0
  RealizarBusqueda
  Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
  On Error Resume Next
  Dim ldFechaIni As Date
  Dim ldFechaFin As Date
  Dim lcFiltro As String
  If cmbFecha.ListIndex = 0 Then
    ldFechaIni = CDate(cmbFecha.Text)
    ldFechaFin = CDate(cmbFecha.Text) + 1
  Else
    ldFechaIni = CDate("01/01/1990")
    ldFechaFin = Date + 1
  End If
  lcFiltro = Trim(HCigualDNI_AgregaNUEVEaLaHistoria(txtNroHistoria.Text))
  
  If cmbFecha.Text <> "Todas" Then
    Set rs = ConstanciaSeleccionaPorFecha(Format(CDate(cmbFecha.Text), sighentidades.DevuelveFechaSoloFormato_DMY_HM), Val(lcFiltro))
  Else
    Set rs = ConstanciaSeleccionaPorFecha(Format(CDate("01/01/1990"), sighentidades.DevuelveFechaSoloFormato_DMY_HM), Val(lcFiltro))
  End If
  Set grdListaConstancias.DataSource = rs
  
  mo_Apariencia.ConfigurarFilasBiColores grdListaConstancias, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
  LimpiarFiltro
End Sub

Public Sub LimpiarFiltro()
  UserControl.txtNroHistoria = ""
  Set grdListaConstancias.DataSource = Nothing
  cmbFecha.ListIndex = 0
End Sub

Private Sub cmbFecha_Click()
  If cmbFecha.Text <> "" Then btnBuscar_Click
End Sub

Private Sub cmbFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then btnBuscar_Click
End Sub

Private Sub cmdBuscaCuentaPorApellidos_Click()
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
      txtNroHistoria.Text = oDOPaciente.NroHistoriaClinica
      txtNroHistoria.SetFocus
      SendKeys "{TAB}"
    End If
  End If
  oConexion.Close
  Set oConexion = Nothing
End Sub

Private Sub grdListaConstancias_Click()
  If rs.EOF = True And rs.BOF = True Then Exit Sub
  ml_idRegistroSeleccionado = rs!idConstancia
  ml_idAtencion = rs!idAtencion
  ml_Historia = rs!NroHistoriaClinica
  ml_idTipoConstancia = rs!idTipoConstancia
  ml_Recibo = rs!Recibo
  ml_Observaciones = rs!Observaciones
  ml_IdServicio = rs!IdServicio
End Sub

Private Sub grdListaConstancias_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
  grdListaConstancias.Bands(0).Columns("IdConstancia").Header.Caption = "Id Constancia"
  grdListaConstancias.Bands(0).Columns("IdConstancia").Width = 1000
  grdListaConstancias.Bands(0).Columns("idAtencion").Hidden = True
  grdListaConstancias.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "H. Clínica"
  grdListaConstancias.Bands(0).Columns("NroHistoriaClinica").Width = 1000
  grdListaConstancias.Bands(0).Columns("ApNom").Header.Caption = "Apellidos y Nombres"
  grdListaConstancias.Bands(0).Columns("ApNom").Width = 3000
  grdListaConstancias.Bands(0).Columns("idPaciente").Hidden = True
  grdListaConstancias.Bands(0).Columns("idResponsable").Hidden = True
  grdListaConstancias.Bands(0).Columns("EstadoConstancia").Hidden = True
  grdListaConstancias.Bands(0).Columns("idServicio").Hidden = True
  grdListaConstancias.Bands(0).Columns("Nombre").Header.Caption = "Nombre Servicio"
  grdListaConstancias.Bands(0).Columns("Fecha").Header.Caption = "Fecha"
  grdListaConstancias.Bands(0).Columns("Fecha").Width = 2200
  grdListaConstancias.Bands(0).Columns("Nombre").Width = 2500
  grdListaConstancias.Bands(0).Columns("idMedico").Hidden = True
  grdListaConstancias.Bands(0).Columns("idTipoConstancia").Hidden = True
  grdListaConstancias.Bands(0).Columns("observaciones").Hidden = True
  grdListaConstancias.Bands(0).Columns("Recibo").Width = 1300
  grdListaConstancias.Bands(0).Columns("Recibo").Header.Caption = "Nro Recibo"
  grdListaConstancias.Bands(0).Columns("PC").Width = 1300
  grdListaConstancias.Bands(0).Columns("PC").Header.Caption = "PC Emisión"
  grdListaConstancias.Bands(0).Columns("NombreConstancia").Width = 1300
  grdListaConstancias.Bands(0).Columns("NombreConstancia").Header.Caption = "Tipo Constancia"
End Sub

Private Sub grdListaConstancias_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
  If Val(Row.Cells("estadoConstancia").GetText()) = 0 Then Row.Appearance.ForeColor = vbRed
End Sub

Private Sub txtNroHistoria_GotFocus()
  txtNroHistoria.SelStart = 0
  txtNroHistoria.SelLength = Len(txtNroHistoria.Text)
End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
  'mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then btnBuscar_Click
  If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then KeyAscii = 0
  End If
End Sub

Private Sub UserControl_GotFocus()
  'btnBuscar_Click
End Sub

Private Sub UserControl_Initialize()
  ml_idRegistroSeleccionado = 0
  ml_IdPaciente = 0
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  'btnBuscar_Click
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
    Case vbKeyF8
  End Select
End Sub
Private Sub UserControl_Resize()
  On Error Resume Next
  fraBusqueda.Width = UserControl.Width - 110
  lblNombre.Width = UserControl.Width
  grdListaConstancias.Width = fraBusqueda.Width
  grdListaConstancias.Height = (UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150))
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
        mo_Apariencia.ConfigurarFilasBiColores grdListaConstancias, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdListaConstancias, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub


Sub inicializar()
  SkinConfigura
  
  cmbFecha.Clear
  cmbFecha.AddItem Date
  cmbFecha.AddItem "Todas"
  cmbFecha.ListIndex = 0
End Sub

Private Sub UserControl_Show()
  btnBuscar_Click
End Sub

