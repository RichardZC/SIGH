VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucPacientesLista 
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
      Height          =   900
      Left            =   90
      TabIndex        =   10
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
         Left            =   6060
         Picture         =   "ucPacientesLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Left            =   4425
         Picture         =   "ucPacientesLista.ctx":058A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   450
         Width           =   315
      End
      Begin VB.TextBox txtFichaFamiliar3 
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
         Left            =   7650
         MaxLength       =   2
         TabIndex        =   6
         Top             =   465
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtFichaFamiliar2 
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
         Left            =   6900
         MaxLength       =   7
         TabIndex        =   5
         Top             =   465
         Visible         =   0   'False
         Width           =   765
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
         Left            =   150
         MaxLength       =   8
         TabIndex        =   0
         Top             =   465
         Width           =   1455
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8520
         Picture         =   "ucPacientesLista.ctx":0B14
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   510
         Width           =   1305
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   8520
         Picture         =   "ucPacientesLista.ctx":36F0
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   1
         Top             =   465
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
         Left            =   3180
         MaxLength       =   40
         TabIndex        =   2
         Top             =   465
         Width           =   1245
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
         Left            =   4830
         MaxLength       =   40
         TabIndex        =   3
         Top             =   465
         Width           =   1335
      End
      Begin VB.TextBox txtFichaFamiliar1 
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
         Left            =   6450
         MaxLength       =   4
         TabIndex        =   4
         Top             =   465
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblFichaFamilar 
         BackStyle       =   0  'Transparent
         Caption         =   "Ficha Familiar"
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
         Left            =   6510
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "N° DNI                  N°Hist.Clínica     Apellido paterno     Apellido materno  "
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
         Left            =   210
         TabIndex        =   11
         Top             =   240
         Width           =   6225
      End
   End
   Begin UltraGrid.SSUltraGrid grdPacientes 
      Height          =   4560
      Left            =   75
      TabIndex        =   9
      Top             =   1545
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   8043
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
      Caption         =   "Pacientes"
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
      TabIndex        =   12
      Top             =   15
      Width           =   9975
   End
End
Attribute VB_Name = "ucPacientesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista pacientes
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
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
        
        If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
             UserControl.txtNroHistoria = "" And txtDni.Text = "") And (txtFichaFamiliar1.Text = "" And txtFichaFamiliar2.Text = "" And txtFichaFamiliar3.Text = "") Then
            MsgBox "Por favor ingrese algunos de los filtros (Ap. Paterno ,Ap. Materno, DNI, Ficha Familiar o Nro Historia)", vbInformation, "Filtro de pacientes"
            Exit Sub
        End If
        If UserControl.txtNroHistoria = "" And txtDni.Text = "" And (txtFichaFamiliar1.Text = "" And txtFichaFamiliar2.Text = "" And txtFichaFamiliar3.Text = "") Then
            If UserControl.txtApellidoPaterno = "" Then
                MsgBox "Por favor ingrese Ap. Paterno", vbInformation, "Filtro de pacientes"
                Exit Sub
            End If
        End If
        oPaciente.ApellidoMaterno = Trim(UserControl.txtApellidoMaterno)
        oPaciente.ApellidoPaterno = Trim(UserControl.txtApellidoPaterno)
        If mo_Teclado.TextoEsSoloNumeros(UserControl.txtNroHistoria) Then
           oPaciente.NroHistoriaClinica = Val(sighEntidades.HCigualDNI_AgregaNUEVEaLaHistoria(UserControl.txtNroHistoria))
           
        End If
        oPaciente.nrodocumento = txtDni.Text
        oPaciente.IdDocIdentidad = 1
        If txtFichaFamiliar1.Text <> "" And txtFichaFamiliar2.Text <> "" And txtFichaFamiliar3.Text <> "" Then
           oPaciente.FichaFamiliar = txtFichaFamiliar1.Text & "-" & txtFichaFamiliar2.Text & "-" & txtFichaFamiliar3.Text
        Else
           oPaciente.FichaFamiliar = ""
        End If
        
        Select Case ml_TipoFiltro
        Case sghFiltrarTodos
            Set grdPacientes.DataSource = mo_AdminAdmision.PacientesFiltrar(oPaciente, _
                                                           IIf(txtApellidoMaterno.Text = wxSinApellido, True, False), _
                                                           IIf(txtApellidoPaterno.Text = wxSinApellido, True, False), _
                                                           wxSinApellido)
        Case sghFiltrarConHistoriasTemporales
            Set grdPacientes.DataSource = mo_AdminAdmision.PacientesFiltrarConHistoriasTemporales(oPaciente)
        Case sghFiltrarConHistoriasDefinitivas
            Set grdPacientes.DataSource = mo_AdminAdmision.PacientesFiltrarConHistoriasDefinitivas(oPaciente, wxSinApellido)
        Case Else
            MsgBox "Opcion no implementada", vbExclamation, Me.Titulo
        End Select
        
        Dim rsRespuesta As New Recordset
        Set rsRespuesta = grdPacientes.DataSource
        On Error Resume Next
        If rsRespuesta.RecordCount = 0 Then
            MsgBox "No se encontraron datos", vbInformation, "Búsqueda"
        Else
        '    UserControl.txtApellidoMaterno = ""
        '    UserControl.txtApellidoPaterno = ""
        '    UserControl.txtPrimerNombre = ""
        '    UserControl.txtNroHistoria = ""
        End If
        
        If mo_AdminAdmision.MensajeError <> "" Then
            MsgBox mo_AdminAdmision.MensajeError, vbInformation, "Filtro Pacientes"
        End If
        
       ' mo_Apariencia.ConfigurarFilasBiColores grdPacientes, sighentidades.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoMaterno = ""
        UserControl.txtApellidoPaterno = ""
        UserControl.txtFichaFamiliar1.Text = ""
        UserControl.txtFichaFamiliar2.Text = ""
        UserControl.txtFichaFamiliar3.Text = ""
        UserControl.txtNroHistoria = ""
        txtDni.Text = ""
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
    ml_idRegistroSeleccionado = rsRecordset("IdPaciente")
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
    ml_idRegistroSeleccionado = rsRecordset("IdPaciente")
    
End Sub

Private Sub grdPacientes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdPacientes.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdPaciente")
    
End Sub

Private Sub grdPacientes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    On Error Resume Next
    grdPacientes.Bands(0).Columns("IdPaciente").Hidden = True
    grdPacientes.Bands(0).Columns("IdTipoNumeracion").Hidden = True
    
    grdPacientes.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nro Historia"
    grdPacientes.Bands(0).Columns("NroHistoriaClinica").Width = 1300
    
    grdPacientes.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdPacientes.Bands(0).Columns("ApellidoPaterno").Width = 1500
    
    grdPacientes.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdPacientes.Bands(0).Columns("ApellidoMaterno").Width = 1500
    
    grdPacientes.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdPacientes.Bands(0).Columns("PrimerNombre").Width = 1500

    grdPacientes.Bands(0).Columns("SegundoNombre").Header.Caption = "2do Nombre"
    grdPacientes.Bands(0).Columns("SegundoNombre").Width = 1500

    grdPacientes.Bands(0).Columns("FechaNacimiento").Header.Caption = "Fecha Nac."
    grdPacientes.Bands(0).Columns("FechaNacimiento").Width = 1500

    grdPacientes.Bands(0).Columns("TipoNumeracion").Header.Caption = "Tipo Numeración"
    grdPacientes.Bands(0).Columns("TipoNumeracion").Width = 1500
    grdPacientes.Bands(0).Columns("TipoNumeracion").CellAppearance.TextAlign = ssAlignRight

  
    grdPacientes.Bands(0).Columns("TipoServicio").Header.Caption = "Ult. Tipo Serv."
    grdPacientes.Bands(0).Columns("TipoServicio").Width = 2000

    grdPacientes.Bands(0).Columns("FechaIngreso").Header.Caption = "Ult. Fec Ing."
    grdPacientes.Bands(0).Columns("FechaIngreso").Width = 1500

    grdPacientes.Bands(0).Columns("FechaEgreso").Header.Caption = "Ult. Fec Egr."
    grdPacientes.Bands(0).Columns("FechaEgreso").Width = 1500

    grdPacientes.Bands(0).Columns("ServicioIngreso").Header.Caption = "Ult. Serv. Ing."
    grdPacientes.Bands(0).Columns("ServicioIngreso").Width = 1500

End Sub




Private Sub txtDNI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDni
    AdministrarKeyPreview KeyCode
End Sub
Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48) Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub

Private Sub txtFichaFamiliar1_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFichaFamiliar1
End Sub



Private Sub txtFichaFamiliar2_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFichaFamiliar2
End Sub



Private Sub txtFichaFamiliar3_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFichaFamiliar3
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
    
    If lcBuscaParametro.SeleccionaFilaParametro(277) = "S" Then
       txtFichaFamiliar1.Visible = True: lblFichaFamilar.Visible = True
       txtFichaFamiliar2.Visible = True: txtFichaFamiliar3.Visible = True
       
    End If
    wxParametro351 = lcBuscaParametro.SeleccionaFilaParametro(351)
End Sub
