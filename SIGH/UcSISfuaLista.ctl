VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl UcSISfuaLista 
   ClientHeight    =   6105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11280
   ScaleHeight     =   6105
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
      Height          =   975
      Left            =   60
      TabIndex        =   7
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
         Left            =   9225
         Picture         =   "UcSISfuaLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   450
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
         Left            =   7620
         Picture         =   "UcSISfuaLista.ctx":058A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   465
         Width           =   315
      End
      Begin VB.TextBox txtFuaLote 
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
         Left            =   660
         MaxLength       =   9
         TabIndex        =   13
         Top             =   465
         Width           =   315
      End
      Begin VB.TextBox txtFuaDisa 
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
         TabIndex        =   12
         Top             =   465
         Width           =   405
      End
      Begin VB.TextBox txtfuaNumero 
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
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   0
         Top             =   465
         Width           =   1815
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
         Left            =   4200
         MaxLength       =   8
         TabIndex        =   2
         Top             =   465
         Width           =   1185
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9735
         Picture         =   "UcSISfuaLista.ctx":0B14
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   510
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   9750
         Picture         =   "UcSISfuaLista.ctx":36F0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   1275
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
         Left            =   5430
         MaxLength       =   9
         TabIndex        =   3
         Top             =   465
         Width           =   1185
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
         Left            =   6630
         MaxLength       =   40
         TabIndex        =   4
         Top             =   465
         Width           =   975
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
         Left            =   8010
         MaxLength       =   40
         TabIndex        =   5
         Top             =   465
         Width           =   1200
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
         Left            =   2970
         MaxLength       =   9
         TabIndex        =   1
         Top             =   465
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   990
         TabIndex        =   15
         Top             =   480
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   570
         TabIndex        =   14
         Top             =   480
         Width           =   75
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "N° de Formato FUA                    N° Cuenta       DNI                Nº Historia    Apell.Paterno      Apell.Materno  "
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
         Width           =   9285
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
      Caption         =   "Lista de FUA (grabados en tablas SIS)"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Formato FUA"
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
Attribute VB_Name = "UcSISfuaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista FUA ya registrada
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mc_FuaVersionFormato As String
Dim mc_FuaTipoAnexo2015 As String

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
Property Let FuaVersionFormato(lValue As String)
    mc_FuaVersionFormato = lValue
End Property
Property Get FuaVersionFormato() As String
    FuaVersionFormato = mc_FuaVersionFormato
End Property
Property Let FuaTipoAnexo2015(lValue As String)
    mc_FuaTipoAnexo2015 = lValue
End Property
Property Get FuaTipoAnexo2015() As String
    FuaTipoAnexo2015 = mc_FuaTipoAnexo2015
End Property

Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
        Dim lcWhereOrder As String
        
        Dim oPaciente As New doPaciente
        Dim lcFechaTriaje As String
        Dim oConexionExterna As New Connection
        '
        oConexionExterna.CursorLocation = adUseClient
        oConexionExterna.CommandTimeout = 150
        oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        '
        If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
            UserControl.txtNroHistoria = "" And UserControl.txtDNI = "" And UserControl.txtNcuenta = "" And UserControl.txtfuaNumero = "") Then
            MsgBox "Por favor ingrese algunos de los filtros (Ap. Paterno ,Ap. Materno, Nombres, DNI, Nro Historia o Nro Cuenta)", vbInformation, "Filtro de pacientes"
            btnLimpiar_Click
            Exit Sub
        End If
        lcWhereOrder = " Where "
        If mo_Teclado.TextoEsSoloNumeros(UserControl.txtfuaNumero.Text) Then
           
           lcWhereOrder = lcWhereOrder & " FuaDisa='" & UserControl.txtFuaDisa & "' and FuaLote='" & UserControl.txtFuaLote & _
                                         "' and FuaNumero='" & UserControl.txtfuaNumero & "'"
        ElseIf mo_Teclado.TextoEsSoloNumeros(UserControl.txtNcuenta.Text) Then
           lcWhereOrder = lcWhereOrder & " idCuentaAtencion=" & UserControl.txtNcuenta.Text
        ElseIf mo_Teclado.TextoEsSoloNumeros(UserControl.txtNroHistoria.Text) Then
           lcWhereOrder = lcWhereOrder & " fuaNroHistoria='" & Trim(HCigualDNI_AgregaNUEVEaLaHistoria(UserControl.txtNroHistoria.Text)) & "'"
           lcWhereOrder = lcWhereOrder & " order by FuaAtencionFecha desc  "
        ElseIf Val(UserControl.txtDNI.Text) > 0 Then
           lcWhereOrder = lcWhereOrder & " DocumentoNumero='" & Trim(UserControl.txtDNI.Text) & "'"
           lcWhereOrder = lcWhereOrder & " order by FuaAtencionFecha desc  "
        Else
           If Trim(UserControl.txtApellidoPaterno.Text) <> "" Then
              lcWhereOrder = lcWhereOrder & " aPaterno like '%" & Trim(UserControl.txtApellidoPaterno.Text) & "%'"
           End If
           If Trim(UserControl.txtApellidoMaterno.Text) <> "" Then
              If Trim(UserControl.txtApellidoPaterno.Text) <> "" Then
                 lcWhereOrder = lcWhereOrder & " and aMaterno like '%" & Trim(UserControl.txtApellidoMaterno.Text) & "%'"
              Else
                 lcWhereOrder = lcWhereOrder & "     aMaterno like '%" & Trim(UserControl.txtApellidoMaterno.Text) & "%'"
              End If
           End If
           lcWhereOrder = lcWhereOrder & " order by aPaterno,aMaterno,FuaAtencionFecha desc"
        End If
        Set grdPacientes.DataSource = mo_ReglasSISgalenhos.SisFiltraPacientesAtendidos(lcWhereOrder)
        '
        oConexionExterna.Close
        Set oConexionExterna = Nothing
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoMaterno = ""
        UserControl.txtApellidoPaterno = ""
        UserControl.txtNroHistoria = ""
        UserControl.txtDNI = ""
        UserControl.txtNcuenta = ""
        UserControl.txtfuaNumero = ""
        txtfuaNumero.SetFocus
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
    mc_FuaVersionFormato = IIf(IsNull(rsRecordset("FuaVersionFormato")), "", rsRecordset("FuaVersionFormato"))
    mc_FuaTipoAnexo2015 = IIf(IsNull(rsRecordset("FuaTipoAnexo2015")), "", rsRecordset("FuaTipoAnexo2015"))
    Debug.Print rsRecordset("NroHistoriaClinica")
End Sub

Private Sub grdPacientes_Click()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdPacientes.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdCuentaAtencion")
End Sub

Private Sub grdPacientes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim rsRecordset As ADODB.Recordset
    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdPacientes.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdCuentaAtencion")
End Sub

Private Sub grdPacientes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    On Error GoTo ErrGrd
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    'grdPacientes.Bands(0).Columns("idAtencion").Hidden = True
    grdPacientes.Bands(0).Columns("FormatoFua").Width = 3000
    grdPacientes.Bands(0).Columns("IdCuentaAtencion").Header.Caption = "Nro Cuenta"
    grdPacientes.Bands(0).Columns("IdCuentaAtencion").Width = 1500
    grdPacientes.Bands(0).Columns("NroHistoria").Header.Caption = "Nro Historia"
    grdPacientes.Bands(0).Columns("NroHistoria").Width = 1500
    grdPacientes.Bands(0).Columns("Paciente").Header.Caption = "Ap. Paterno"
    grdPacientes.Bands(0).Columns("Paciente").Width = 3000
    grdPacientes.Bands(0).Columns("FuaVersionFormato").Hidden = True
    grdPacientes.Bands(0).Columns("FuaTipoAnexo2015").Hidden = True
ErrGrd:
End Sub

Private Sub txtDNI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnBuscar.SetFocus
    ElseIf KeyCode = vbKeyF6 Then
        btnBuscar_Click
    End If
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
'    If mo_Teclado.TextoEsSoloNumeros(txtDni.Text) Then
'       btnBuscar_Click
'    End If
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub

Private Sub txtfuaNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnBuscar.SetFocus
    ElseIf KeyCode = vbKeyF6 Then
        If Len(txtfuaNumero.Text) < 8 And Len(txtfuaNumero.Text) >= 1 Then
            txtfuaNumero.Text = String(8 - Len(txtfuaNumero.Text), "0") & txtfuaNumero.Text
        End If
        btnBuscar_Click
    End If
End Sub

Private Sub txtfuaNumero_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then
'    If mo_Teclado.TextoEsSoloNumeros(txtfuaNumero.Text) Then
'       btnBuscar_Click
'    End If
'  End If
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub

Private Sub txtfuaNumero_LostFocus()
    If Len(txtfuaNumero.Text) < 8 And Len(txtfuaNumero.Text) >= 1 Then
        txtfuaNumero.Text = String(8 - Len(txtfuaNumero.Text), "0") & txtfuaNumero.Text
    End If
End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnBuscar.SetFocus
    ElseIf KeyCode = vbKeyF6 Then
        btnBuscar_Click
    End If
End Sub

Private Sub txtNcuenta_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'    If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
'       btnBuscar_Click
'    End If
'   End If
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnBuscar.SetFocus
    ElseIf KeyCode = vbKeyF6 Then
        btnBuscar_Click
    End If
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub



Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
End Sub



Private Sub txtApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaterno
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
    'mo_Apariencia.ConfigurarFilasBiColores grdPacientes, sighentidades.GrillaConFilasBicolor
    Dim oRsTmp As New Recordset
    Set oRsTmp = mo_ReglasSISgalenhos.SisFuaSeleccionarTodos
    If oRsTmp.RecordCount > 0 Then
       txtFuaDisa.Text = oRsTmp.Fields!fuaDisa
       txtFuaLote.Text = oRsTmp.Fields!fuaLote
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Sub





