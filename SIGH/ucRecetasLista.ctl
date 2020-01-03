VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucRecetasLista 
   ClientHeight    =   6195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10080
   ScaleHeight     =   6195
   ScaleWidth      =   10080
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
      Top             =   570
      Width           =   9930
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
         Left            =   8055
         Picture         =   "ucRecetasLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   6375
         Picture         =   "ucRecetasLista.ctx":058A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   450
         Width           =   315
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
         Left            =   1380
         MaxLength       =   9
         TabIndex        =   12
         Top             =   465
         Width           =   1215
      End
      Begin VB.TextBox txtNreceta 
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
         Left            =   90
         MaxLength       =   9
         TabIndex        =   10
         Top             =   465
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
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   0
         Top             =   465
         Width           =   1335
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8550
         Picture         =   "ucRecetasLista.ctx":0B14
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   510
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   8520
         Picture         =   "ucRecetasLista.ctx":36F0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   1305
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
         Left            =   3960
         MaxLength       =   9
         TabIndex        =   1
         Top             =   465
         Width           =   1335
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
         Left            =   5280
         MaxLength       =   40
         TabIndex        =   2
         Top             =   465
         Width           =   1080
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
         Left            =   6870
         MaxLength       =   40
         TabIndex        =   3
         Top             =   465
         Width           =   1170
      End
      Begin VB.Label lblFichaFamilar 
         BackStyle       =   0  'Transparent
         Caption         =   "N° Receta"
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
         Left            =   150
         TabIndex        =   11
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "N° Cuenta       N° DNI             N°Hist.Clínica     Apellido paterno     Apellido materno  "
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
         Left            =   1410
         TabIndex        =   8
         Top             =   240
         Width           =   7095
      End
   End
   Begin UltraGrid.SSUltraGrid grdPacientes 
      Height          =   4590
      Left            =   75
      TabIndex        =   6
      Top             =   1515
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   8096
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Caption         =   "Recetas"
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
      Left            =   30
      TabIndex        =   9
      Top             =   15
      Width           =   9975
   End
End
Attribute VB_Name = "ucRecetasLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Recetas registradas
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_IdTipoServicio As Long
Dim ml_FechaReceta As Date

Property Let FechaReceta(lValue As Date)
    ml_FechaReceta = lValue
End Property
Property Get FechaReceta() As Date
    FechaReceta = ml_FechaReceta
End Property

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
Property Let idTipoServicio(lValue As Long)
    ml_IdTipoServicio = lValue
End Property


Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
Dim oPaciente As New doPaciente
        
        If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
             UserControl.txtNroHistoria = "" And txtDNI.Text = "") And (txtNreceta.Text = "") And txtNcuenta.Text = "" Then
            MsgBox "Por favor ingrese algunos de los filtros (Ap. Paterno ,Ap. Materno, DNI, N° Receta, N° Cuenta o Nro Historia)", vbInformation, "Filtro de pacientes"
            Exit Sub
        End If
        If UserControl.txtNroHistoria = "" And txtDNI.Text = "" And (txtNreceta.Text = "") And txtNcuenta.Text = "" Then
            If UserControl.txtApellidoPaterno = "" Then
                MsgBox "Por favor ingrese Ap. Paterno", vbInformation, "Filtro de Recetas"
                Exit Sub
            End If
        End If
        oPaciente.ApellidoMaterno = UserControl.txtApellidoMaterno
        oPaciente.ApellidoPaterno = UserControl.txtApellidoPaterno
        oPaciente.NroHistoriaClinica = Val(HCigualDNI_AgregaNUEVEaLaHistoria(UserControl.txtNroHistoria.Text))
        oPaciente.nrodocumento = txtDNI.Text
        oPaciente.IdDocIdentidad = 1
        Set grdPacientes.DataSource = mo_AdminAdmision.RecetaFiltrar(oPaciente, Val(txtNreceta.Text), _
                                                                Val(txtNcuenta.Text), ml_IdTipoServicio, _
                                                                wxSinApellido)
       ' mo_Apariencia.ConfigurarFilasBiColores grdPacientes, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
        LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoMaterno = ""
        UserControl.txtApellidoPaterno = ""
        UserControl.txtNroHistoria = ""
        txtDNI.Text = ""
        txtNreceta.Text = ""
        txtNcuenta.Text = ""
        On Error Resume Next
        txtNreceta.SetFocus
End Sub

Private Sub cmdSinApellidoMaterno_Click()
    txtApellidoMaterno.Text = wxSinApellido
End Sub

Private Sub cmdSinApellidoPaterno_Click()
    txtApellidoPaterno.Text = wxSinApellido
End Sub

Private Sub grdPacientes_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdPacientes.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdCuentaAtencion")
    ml_FechaReceta = rsRecordset("FechaReceta")
End Sub

Private Sub grdPacientes_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
Cancel = True
End Sub

Private Sub grdPacientes_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdPacientes.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdCuentaAtencion")
    ml_FechaReceta = rsRecordset("FechaReceta")
    
End Sub

Private Sub grdPacientes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdPacientes.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("idCuentaAtencion")
    ml_FechaReceta = rsRecordset("FechaReceta")

End Sub

Private Sub grdPacientes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    On Error Resume Next
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    
    grdPacientes.Bands(0).Columns("idTipoServicio").Hidden = True
    
    grdPacientes.Bands(0).Columns("idCuentaAtencion").Header.Caption = "Nro Cuenta"
    grdPacientes.Bands(0).Columns("idCuentaAtencion").Width = 1000
    
    grdPacientes.Bands(0).Columns("FechaReceta").Header.Caption = "F.Receta"
    grdPacientes.Bands(0).Columns("FechaReceta").Width = 1400
    
    grdPacientes.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nro Historia"
    grdPacientes.Bands(0).Columns("NroHistoriaClinica").Width = 1300
    
    grdPacientes.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdPacientes.Bands(0).Columns("ApellidoPaterno").Width = 1500
    
    grdPacientes.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdPacientes.Bands(0).Columns("ApellidoMaterno").Width = 1500
    
    grdPacientes.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdPacientes.Bands(0).Columns("PrimerNombre").Width = 1500

    grdPacientes.Bands(0).Columns("NroDocumento").Header.Caption = "N° Documento"
    grdPacientes.Bands(0).Columns("NroDocumento").Width = 1500
End Sub




Private Sub txtDNI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDNI
    AdministrarKeyPreview KeyCode
End Sub









Private Sub txtDni_LostFocus()
    If txtDNI.Text <> "" Then
       btnBuscar_Click
    End If
End Sub



Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNcuenta_LostFocus()
    If Val(txtNcuenta.Text) > 0 Then
       btnBuscar_Click
    End If

End Sub

Private Sub txtNreceta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNreceta
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNreceta_LostFocus()
    If Val(txtNreceta.Text) > 0 Then
       btnBuscar_Click
    End If
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


Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
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
    'AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtNroHistoria_LostFocus()
    If txtNroHistoria.Text <> "" Then
       btnBuscar_Click
    End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   
   grdPacientes.Width = fraBusqueda.Width
   grdPacientes.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 330)
   
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
        mo_Apariencia.ConfigurarFilasBiColores grdPacientes, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdPacientes, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub

Public Sub inicializar()
    SkinConfigura
    Dim orsTemp As New Recordset
    Me.LimpiarFiltro
    Set grdPacientes.DataSource = orsTemp
End Sub

Public Function TieneRegistros() As Boolean
    Dim orsTemp As New Recordset
    Set orsTemp = grdPacientes.DataSource
    TieneRegistros = False
    If orsTemp.State <> 0 Then
        If orsTemp.RecordCount > 0 Then
            TieneRegistros = True
        End If
    End If
End Function
