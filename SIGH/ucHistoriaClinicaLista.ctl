VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucHistoriaClinicaLista 
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   ScaleHeight     =   6210
   ScaleWidth      =   10095
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
      TabIndex        =   7
      Top             =   540
      Width           =   9990
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
         Height          =   345
         Left            =   5310
         Picture         =   "ucHistoriaClinicaLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   435
         Width           =   345
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
         Height          =   345
         Left            =   3375
         Picture         =   "ucHistoriaClinicaLista.ctx":058A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   450
         Width           =   345
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9045
         Picture         =   "ucHistoriaClinicaLista.ctx":0B14
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7680
         Picture         =   "ucHistoriaClinicaLista.ctx":36F0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox txtPrimerNombre 
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
         Left            =   5760
         MaxLength       =   40
         TabIndex        =   3
         Top             =   450
         Width           =   1845
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
         Left            =   3840
         MaxLength       =   40
         TabIndex        =   2
         Top             =   450
         Width           =   1455
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
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   1
         Top             =   450
         Width           =   1455
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
         Left            =   120
         MaxLength       =   9
         TabIndex        =   0
         Top             =   450
         Width           =   1740
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Historia clínica        Apellido paterno          Apellido materno           Primer nombre                   "
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
         Left            =   270
         TabIndex        =   8
         Top             =   210
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdHistoriasClinicas 
      Height          =   4665
      Left            =   90
      TabIndex        =   6
      Top             =   1500
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8229
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
      Caption         =   "Lista de historias clínicas"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Historias Clínicas"
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
      Top             =   15
      Width           =   10080
   End
End
Attribute VB_Name = "ucHistoriaClinicaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Historia Clinica
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim ml_idRegistroSeleccionado As Long

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdHistoriasClinicas.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdHistoriasClinicas.DataSource
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


Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub
Public Sub RealizarBusqueda()
Dim oPaciente As New doPaciente
        

        If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
            UserControl.txtPrimerNombre = "" And UserControl.txtNroHistoria = "") Then
            MsgBox "Por favor ingrese algunos de los filtros (Ap. Paterno ,Ap. Materno, Nombres o Nro Historia)", vbInformation, "Filtro de pacientes"
            Exit Sub
        End If
        
        oPaciente.ApellidoMaterno = UserControl.txtApellidoMaterno
        oPaciente.ApellidoPaterno = UserControl.txtApellidoPaterno
        oPaciente.PrimerNombre = UserControl.txtPrimerNombre
        If mo_Teclado.TextoEsSoloNumeros(UserControl.txtNroHistoria) Then
           oPaciente.NroHistoriaClinica = Val(HCigualDNI_AgregaNUEVEaLaHistoria(UserControl.txtNroHistoria.Text))
        End If
        
        Set grdHistoriasClinicas.DataSource = mo_AdminArchivoClinico.HistoriaClinicaFiltrar(oPaciente, wxSinApellido)
        
        If mo_AdminArchivoClinico.MensajeError <> "" Then
            MsgBox mo_AdminArchivoClinico.MensajeError, vbInformation, "Filtro HistoriasClinicas"
        End If
        
     '   mo_Apariencia.ConfigurarFilasBiColores grdHistoriasClinicas, sighentidades.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoMaterno = ""
        UserControl.txtApellidoPaterno = ""
        UserControl.txtPrimerNombre = ""
        UserControl.txtNroHistoria = ""
End Sub

Private Sub cmdSinApellidoMaterno_Click()
    txtApellidoMaterno.Text = wxSinApellido
End Sub

Private Sub cmdSinApellidoPaterno_Click()
    txtApellidoPaterno.Text = wxSinApellido
End Sub

Private Sub grdHistoriasClinicas_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdHistoriasClinicas.DataSource
    On Error Resume Next
    'ml_idRegistroSeleccionado = IIf(IsNull(rsRecordset("NroHistoriaClinica")), -1, rsRecordset("NroHistoriaClinica"))
    ml_idRegistroSeleccionado = IIf(IsNull(rsRecordset("hc")), -1, rsRecordset("hc"))
    
End Sub

Private Sub grdHistoriasClinicas_Click()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdHistoriasClinicas.DataSource
    On Error Resume Next
    'ml_idRegistroSeleccionado = IIf(IsNull(rsRecordset("NroHistoriaClinica")), -1, rsRecordset("NroHistoriaClinica"))
    ml_idRegistroSeleccionado = IIf(IsNull(rsRecordset("hc")), -1, rsRecordset("hc"))
End Sub


Private Sub grdHistoriasClinicas_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    
    grdHistoriasClinicas.Bands(0).Columns("IdPaciente").Hidden = True
    grdHistoriasClinicas.Bands(0).Columns("hc").Hidden = True
    
    grdHistoriasClinicas.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nro Historia"
    grdHistoriasClinicas.Bands(0).Columns("NroHistoriaClinica").Width = 1000
    
    grdHistoriasClinicas.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Apellido Paterno"
    grdHistoriasClinicas.Bands(0).Columns("ApellidoPaterno").Width = 2000
    
    grdHistoriasClinicas.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Apellido Materno"
    grdHistoriasClinicas.Bands(0).Columns("ApellidoMaterno").Width = 2000
    
    grdHistoriasClinicas.Bands(0).Columns("PrimerNombre").Header.Caption = "Primer Nombre"
    grdHistoriasClinicas.Bands(0).Columns("PrimerNombre").Width = 2000

    grdHistoriasClinicas.Bands(0).Columns("SegundoNombre").Header.Caption = "Segundo Nombre"
    grdHistoriasClinicas.Bands(0).Columns("SegundoNombre").Width = 2000

    grdHistoriasClinicas.Bands(0).Columns("FechaCreacion").Header.Caption = "Fecha Creación"
    grdHistoriasClinicas.Bands(0).Columns("FechaCreacion").Width = 1500

End Sub
Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtPrimerNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombre
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtPrimerNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaterno
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdHistoriasClinicas.Width = fraBusqueda.Width
   grdHistoriasClinicas.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
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

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdHistoriasClinicas, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdHistoriasClinicas, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub
Sub Inicializar()
    SkinConfigura
End Sub
