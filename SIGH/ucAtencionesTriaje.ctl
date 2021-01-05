VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucAtencionesTriaje 
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11280
   ScaleHeight     =   6150
   ScaleWidth      =   11280
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
      Left            =   30
      TabIndex        =   8
      Top             =   525
      Width           =   11190
      Begin VB.CommandButton cmdSinApellidoMaterno 
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
         Left            =   6720
         Picture         =   "ucAtencionesTriaje.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   465
         Width           =   315
      End
      Begin VB.CommandButton cmdSinApellidoPaterno 
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
         Left            =   5175
         Picture         =   "ucAtencionesTriaje.ctx":058A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   465
         Width           =   315
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
         Left            =   7140
         TabIndex        =   5
         Text            =   "cmbFecha"
         Top             =   450
         Width           =   1500
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
         Left            =   180
         MaxLength       =   9
         TabIndex        =   0
         Top             =   450
         Width           =   1185
      End
      Begin VB.TextBox txtApellidoMaterno 
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
         Left            =   5580
         MaxLength       =   40
         TabIndex        =   4
         Top             =   465
         Width           =   1125
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
         Left            =   4110
         MaxLength       =   40
         TabIndex        =   3
         Top             =   465
         Width           =   1065
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
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   2
         Top             =   465
         Width           =   1425
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   8730
         Picture         =   "ucAtencionesTriaje.ctx":0B14
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   450
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   10095
         Picture         =   "ucAtencionesTriaje.ctx":375D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   450
         Width           =   1275
      End
      Begin VB.TextBox txtDNI 
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
         Left            =   1410
         MaxLength       =   8
         TabIndex        =   1
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "N° Cuenta       DNI                Nº Historia          Apellido Paterno   Apellido Materno   Fecha Triaje"
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
         TabIndex        =   9
         Top             =   240
         Width           =   8235
      End
   End
   Begin UltraGrid.SSUltraGrid grdPacientes 
      Height          =   4590
      Left            =   45
      TabIndex        =   10
      Top             =   1500
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   8096
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
      Caption         =   "Lista de pacientes"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Triaje "
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
      TabIndex        =   11
      Top             =   0
      Width           =   11205
   End
End
Attribute VB_Name = "ucAtencionesTriaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para listar atenciones de triaje
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim lcBuscaParametro As New SIGHDatos.Parametros

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdPacientes.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdPacientes.DataSource
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
Property Let TipoFiltro(lValue As sghTipoFiltroPacientes)
    ml_TipoFiltro = lValue
End Property
Property Get TipoFiltro() As sghTipoFiltroPacientes
    TipoFiltro = ml_TipoFiltro
End Property


Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
        Dim oPaciente As New doPaciente
        Dim lcFechaTriaje As String
        Dim oConexionExterna As New Connection
        '
        oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        oConexionExterna.CursorLocation = adUseClient
        oConexionExterna.CommandTimeout = 150
        '
        If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
            UserControl.txtNroHistoria = "" And UserControl.txtDni = "" And UserControl.txtNcuenta = "" And cmbFecha.ListIndex = 1) Then
            MsgBox "Por favor ingrese algunos de los filtros (Ap. Paterno ,Ap. Materno, Nombres, DNI, Nro Historia o Nro Cuenta)", vbInformation, "Filtro de pacientes"
            btnLimpiar_Click
            Exit Sub
        End If
        oPaciente.ApellidoMaterno = UserControl.txtApellidoMaterno
        oPaciente.ApellidoPaterno = UserControl.txtApellidoPaterno
        If mo_Teclado.TextoEsSoloNumeros(UserControl.txtNroHistoria) Then
           oPaciente.NroHistoriaClinica = Val(HCigualDNI_AgregaNUEVEaLaHistoria(UserControl.txtNroHistoria.Text))
        End If
        oPaciente.IdDocIdentidad = 1
        oPaciente.nrodocumento = UserControl.txtDni.Text
        If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
            UserControl.txtNroHistoria = "" And UserControl.txtDni = "" And UserControl.txtNcuenta = "") Then
            lcFechaTriaje = IIf(cmbFecha.Text = "Todas", Date, cmbFecha.Text)
        Else
            lcFechaTriaje = ""
        End If
        Set grdPacientes.DataSource = mo_ReglasAdmision.AtencionesCEFiltrarPorPaciente(oPaciente, _
                                           Val(UserControl.txtNcuenta.Text), lcFechaTriaje, oConexionExterna, _
                                           IIf(txtApellidoMaterno.Text = wxSinApellido, True, False), _
                                           IIf(txtApellidoPaterno.Text = wxSinApellido, True, False), _
                                           wxSinApellido)

        '
        oConexionExterna.Close
        Set oConexionExterna = Nothing
End Sub

'Actualizado 14102014
Private Sub btnBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoMaterno = ""
        UserControl.txtApellidoPaterno = ""
        UserControl.txtNroHistoria = ""
        UserControl.txtDni = ""
        UserControl.txtNcuenta = ""
        cmbFecha.Clear
        cmbFecha.AddItem Date
        cmbFecha.AddItem "Todas"
        cmbFecha.ListIndex = 0
End Sub

'Actualizado 14102014
Private Sub btnLimpiar_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbFecha_KeyDown(KeyCode As Integer, Shift As Integer)
        'Actualizado 14102014
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbFecha_LostFocus()
    If cmbFecha.ListIndex <> 1 Then
        If Not EsFecha(cmbFecha.Text, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, ""
            cmbFecha.Text = Date
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdSinApellidoMaterno_Click()
    txtApellidoMaterno.Text = wxSinApellido
End Sub

Private Sub cmdSinApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
    'Actualizado 14102014
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdSinApellidoPaterno_Click()
    txtApellidoPaterno.Text = wxSinApellido
End Sub

Private Sub cmdSinApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
    'Actualizado 14102014
    AdministrarKeyPreview KeyCode
End Sub

Private Sub grdPacientes_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdPacientes.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdAtencion")
    Debug.Print rsRecordset("NroHistoriaClinica")
End Sub

'Actualizado 15102014
Private Sub grdPacientes_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdPacientes_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdPacientes.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdAtencion")
    
End Sub

Private Sub grdPacientes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdPacientes.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdAtencion")
    
End Sub

Private Sub grdPacientes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    On Error GoTo ErrGrd
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    
    grdPacientes.Bands(0).Columns("idAtencion").Hidden = True
    
    grdPacientes.Bands(0).Columns("IdCuentaAtencion").Header.Caption = "Nro Cuenta"
    grdPacientes.Bands(0).Columns("IdCuentaAtencion").Width = 1500
    
    grdPacientes.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nro Historia"
    grdPacientes.Bands(0).Columns("NroHistoriaClinica").Width = 1500
    
    grdPacientes.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdPacientes.Bands(0).Columns("ApellidoPaterno").Width = 1500
    
    grdPacientes.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdPacientes.Bands(0).Columns("ApellidoMaterno").Width = 1500
    
    grdPacientes.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdPacientes.Bands(0).Columns("PrimerNombre").Width = 1500
ErrGrd:
End Sub

Private Sub grdPacientes_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    'Actualizado 14102014
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub txtDNI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDni
    'Actualizado 14102014
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
    
'    If KeyAscii = 13 And txtDni.Text <> "" Then
'       cmbFecha.ListIndex = 1
'       btnBuscar_Click
'    End If
End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
'Actualizado 14102014
    If txtNcuenta.Text = "" Then
        mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta
    Else
        If KeyCode = vbKeyReturn Then
            btnBuscar_Click
        Else
            AdministrarKeyPreview KeyCode
        End If
    End If
End Sub

Private Sub txtNcuenta_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
'    If KeyAscii = 13 And Val(txtNcuenta.Text) > 0 Then
'       cmbFecha.ListIndex = 1
'       btnBuscar_Click
'    End If
End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
    'Actualizado 14102014
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
    'Actualizado 14102014
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtApellidoMaterno.Text <> "" Then
       cmbFecha.ListIndex = 1
    End If
End Sub


Private Sub txtApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaterno
    'Actualizado 14102014
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtApellidoPaterno.Text <> "" Then
       cmbFecha.ListIndex = 1
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
   
   grdPacientes.Width = fraBusqueda.Width
   grdPacientes.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 330)
   
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighEntidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdPacientes, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdPacientes, sighEntidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub

Public Sub Inicializar()
    SkinConfigura
    
    cmbFecha.Clear
    cmbFecha.AddItem Date
    cmbFecha.AddItem "Todas"
    cmbFecha.ListIndex = 0
   'mo_Apariencia.ConfigurarFilasBiColores grdPacientes, sighentidades.GrillaConFilasBicolor
End Sub



