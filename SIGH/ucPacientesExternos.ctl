VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucPacienteExternos 
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11160
   ScaleHeight     =   6555
   ScaleWidth      =   11160
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
      Height          =   885
      Left            =   90
      TabIndex        =   2
      Top             =   630
      Width           =   11025
      Begin VB.CommandButton cmdSinApellidoPaterno 
         Caption         =   "..."
         Height          =   315
         Left            =   4680
         TabIndex        =   9
         ToolTipText     =   "Sin apellido PATERNO"
         Top             =   450
         Width           =   255
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   8160
         Picture         =   "ucPacientesExternos.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   420
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9540
         Picture         =   "ucPacientesExternos.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   420
         Width           =   1275
      End
      Begin VB.TextBox txtNcuenta 
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
         Left            =   150
         MaxLength       =   9
         TabIndex        =   5
         Top             =   450
         Width           =   1425
      End
      Begin VB.TextBox txtNroHistoria 
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
         Left            =   1620
         MaxLength       =   9
         TabIndex        =   4
         Top             =   450
         Width           =   1395
      End
      Begin VB.TextBox txtApellidoPaterno 
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
         Left            =   3060
         MaxLength       =   40
         TabIndex        =   3
         Top             =   450
         Width           =   1605
      End
      Begin VB.Label Label2 
         Caption         =   "N° Cuenta          Historia clínica      Apellido paterno      F.Ingreso            Servicio          "
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
         Top             =   225
         Width           =   4635
      End
   End
   Begin UltraGrid.SSUltraGrid grdAdmision 
      Height          =   4860
      Left            =   90
      TabIndex        =   0
      Top             =   1590
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   8573
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
      Caption         =   "Lista "
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Pacientes Externos"
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
      TabIndex        =   1
      Top             =   45
      Width           =   11010
   End
End
Attribute VB_Name = "ucPacienteExternos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para buscar pacientes externos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_idRegistroSeleccionado As Long
Dim ml_IdAtencionSeleccionada As Long
Dim ml_TipoFiltro As sghTipoFiltroAdmision
Dim mo_cmbIdResponsable As New sighentidades.ListaDespleglable
Dim mrs_Tmp As New ADODB.Recordset
Public Event OnClick(oRecordset As Recordset)
Dim ml_IdServicioConCamaDisponible As Long
Dim ml_idUsuario As Long
Dim mb_EsPacienteSinSeguro As Boolean

Property Let EsPacienteSinSeguro(lValue As Boolean)
    mb_EsPacienteSinSeguro = lValue
    If mb_EsPacienteSinSeguro = True Then
       lblNombre.Caption = "Pacientes Externos con Cuenta (Particular)"
    Else
       lblNombre.Caption = "Pacientes Externos con Cuenta (Seguro)"
    End If
End Property


Property Let IdServicioConCamaDisponible(lValue As Long)
    ml_IdServicioConCamaDisponible = lValue
End Property
Property Get IdServicioConCamaDisponible() As Long
    IdServicioConCamaDisponible = ml_IdServicioConCamaDisponible
End Property

Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
    idUsuario = ml_idUsuario
End Property

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdAdmision.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdAdmision.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property
Property Let IdAtencionSeleccionada(lValue As Long)
    ml_IdAtencionSeleccionada = lValue
End Property
Property Get IdAtencionSeleccionada() As Long
    IdAtencionSeleccionada = ml_IdAtencionSeleccionada
End Property

Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property
Property Let TipoFiltro(lValue As sghTipoFiltroAdmision)
    ml_TipoFiltro = lValue
End Property
Property Get TipoFiltro() As sghTipoFiltroAdmision
    TipoFiltro = ml_TipoFiltro
End Property


Private Sub btnBuscar_Click()
   
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault

End Sub

Public Sub RealizarBusqueda()
Dim oDOPaciente As New doPaciente
Dim oDOAtencion As New DOAtencion
Dim lbSigue As Boolean
Dim lnListIndex As Integer
Dim rsRespuesta As New Recordset
Dim lbServicioVacio As Boolean
        If UserControl.txtNroHistoria.Text <> "" Then
           UserControl.txtNcuenta.Text = ""
           UserControl.txtApellidoPaterno.Text = ""
        ElseIf UserControl.txtNcuenta.Text <> "" Then
           UserControl.txtNroHistoria.Text = ""
           UserControl.txtApellidoPaterno.Text = ""
        ElseIf UserControl.txtApellidoPaterno.Text <> "" Then
           UserControl.txtNroHistoria.Text = ""
           UserControl.txtNcuenta.Text = ""
        End If
        If mb_EsPacienteSinSeguro = True Then
           Set grdAdmision.DataSource = mo_AdminAdmision.AtencionesSeleccionarPacExtPorCuentaHistoriaApellidosServPARTIC(Val(HCigualDNI_AgregaNUEVEaLaHistoria(UserControl.txtNroHistoria.Text)), Val(UserControl.txtNcuenta.Text), UserControl.txtApellidoPaterno.Text)
        Else
           Set grdAdmision.DataSource = mo_AdminAdmision.AtencionesSeleccionarPacExtPorCuentaHistoriaApellidosServSEGUROS(Val(HCigualDNI_AgregaNUEVEaLaHistoria(UserControl.txtNroHistoria.Text)), Val(UserControl.txtNcuenta.Text), UserControl.txtApellidoPaterno.Text, wxSinApellido)
        End If
        On Error Resume Next
        If mo_AdminAdmision.MensajeError <> "" Then
            MsgBox mo_AdminAdmision.MensajeError, vbInformation, "Filtro Pacientes"
        End If
        mo_Apariencia.ConfigurarFilasBiColores grdAdmision, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoPaterno = ""
        UserControl.txtNroHistoria = ""
        UserControl.txtNcuenta = ""
        UserControl.txtNcuenta.SetFocus
End Sub








Private Sub cmdSinApellidoPaterno_Click()
    txtApellidoPaterno.Text = wxSinApellido
End Sub

Private Sub grdAdmision_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    On Error Resume Next
    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdAdmision.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdAtencion")
    ml_IdAtencionSeleccionada = rsRecordset("IdAtencion")
    RaiseEvent OnClick(rsRecordset)
End Sub

Private Sub grdAdmision_Click()
Dim rsRecordset As ADODB.Recordset

    On Error Resume Next
    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdAdmision.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdAtencion")
    ml_IdAtencionSeleccionada = rsRecordset("IdCuentaAtencion")
    RaiseEvent OnClick(rsRecordset)
    
End Sub


Private Sub grdAdmision_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    On Error Resume Next
    grdAdmision.Bands(0).Columns("IdPaciente").Hidden = True
'    grdAdmision.Bands(0).Columns("IdCuentaAtencion").Hidden = True
    grdAdmision.Bands(0).Columns("IdAtencion").Hidden = True
    
    grdAdmision.Bands(0).Columns("IdTipoNumeracion").Hidden = True
    grdAdmision.Bands(0).Columns("IdServicioIngreso").Hidden = True
    
    grdAdmision.Bands(0).Columns("IdAtencion").Header.Caption = "N° atención"
    grdAdmision.Bands(0).Columns("IdAtencion").Width = 1300
    
    grdAdmision.Bands(0).Columns("FechaIngreso").Header.Caption = "Fecha Ing."
    grdAdmision.Bands(0).Columns("FechaIngreso").Width = 1000
    
    grdAdmision.Bands(0).Columns("HoraIngreso").Header.Caption = "Hr.Ing"
    grdAdmision.Bands(0).Columns("HoraIngreso").Width = 800
    
    grdAdmision.Bands(0).Columns("IdCuentaAtencion").Header.Caption = "N° Cuenta"
    grdAdmision.Bands(0).Columns("IdCuentaAtencion").Width = 1300
    
    grdAdmision.Bands(0).Columns("TipoNumeracion").Header.Caption = "Tipo Numeración"
    grdAdmision.Bands(0).Columns("TipoNumeracion").Width = 1500
    grdAdmision.Bands(0).Columns("ServicioActual").Header.Caption = "Servicio Actual"
    grdAdmision.Bands(0).Columns("ServicioActual").Width = 2500
    grdAdmision.Bands(0).Columns("Edad").Hidden = True
    
    grdAdmision.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdAdmision.Bands(0).Columns("ApellidoPaterno").Width = 1500
    
    grdAdmision.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdAdmision.Bands(0).Columns("ApellidoMaterno").Width = 1500
    
    grdAdmision.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdAdmision.Bands(0).Columns("PrimerNombre").Width = 1500

    grdAdmision.Bands(0).Columns("SegundoNombre").Hidden = True

    grdAdmision.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "N° HC"
    grdAdmision.Bands(0).Columns("NroHistoriaClinica").Width = 1200

    grdAdmision.Bands(0).Columns("IdTipoServicio").Hidden = True
        
    grdAdmision.Bands(0).Columns("Plan").Header.Caption = "Fuente Financiamiento(IAFA)"
End Sub

Private Sub grdAdmision_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
        If Val(Row.Cells("IdEstadoAtencion").GetText()) = 0 Then
            Row.Appearance.ForeColor = vbRed
        End If
End Sub




Private Sub txtApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaterno
End Sub

Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
   If KeyAscii = 13 And Len(txtApellidoPaterno.Text) > 0 Then
       btnBuscar_Click
   End If
End Sub


Private Sub txtApellidoPaterno_LostFocus()
    If txtApellidoPaterno.Text <> "" Then
       txtNroHistoria.Text = ""
       txtNcuenta.Text = ""
    End If
End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta
End Sub

Private Sub txtNcuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(txtNcuenta.Text) > 0 Then
       btnBuscar_Click
    End If
End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
   If KeyAscii = 13 And Len(txtNroHistoria.Text) > 0 Then
       btnBuscar_Click
   End If
   
End Sub


Private Sub txtNroHistoria_LostFocus()
    If txtNroHistoria.Text <> "" Then
       txtNcuenta.Text = ""
       txtApellidoPaterno.Text = ""
    End If
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
   grdAdmision.Width = fraBusqueda.Width
   grdAdmision.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)


End Sub

Public Sub Inicializar()
    On Error Resume Next
    txtNcuenta.SetFocus
End Sub


Public Sub FocusEnNroHistoria()
    txtNroHistoria.SetFocus
End Sub



