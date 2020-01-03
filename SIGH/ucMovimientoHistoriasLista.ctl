VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucMovimientoHistoriasLista 
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10200
   ScaleHeight     =   6465
   ScaleWidth      =   10200
   Begin VB.Frame fraBusqueda 
      Caption         =   "Busqueda"
      Height          =   1515
      Left            =   75
      TabIndex        =   7
      Top             =   555
      Width           =   10035
      Begin VB.CommandButton cmdSinApellidoMaterno 
         Caption         =   "..."
         Height          =   315
         Left            =   5610
         TabIndex        =   15
         ToolTipText     =   "Sin apellido MATERNO"
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdSinApellidoPaterno 
         Caption         =   "..."
         Height          =   315
         Left            =   3690
         TabIndex        =   14
         ToolTipText     =   "Sin apellido PATERNO"
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtFichaFamiliar 
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
         Left            =   2130
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1050
         Visible         =   0   'False
         Width           =   1815
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
         ItemData        =   "ucMovimientoHistoriasLista.ctx":0000
         Left            =   180
         List            =   "ucMovimientoHistoriasLista.ctx":0002
         TabIndex        =   6
         Text            =   "cmbFecha"
         Top             =   1050
         Width           =   1710
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9225
         Picture         =   "ucMovimientoHistoriasLista.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7860
         Picture         =   "ucMovimientoHistoriasLista.ctx":2BE0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   450
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
         Left            =   150
         MaxLength       =   9
         TabIndex        =   0
         Top             =   480
         Width           =   1725
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
         Left            =   2100
         MaxLength       =   40
         TabIndex        =   1
         Top             =   480
         Width           =   1575
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
         Left            =   4035
         MaxLength       =   40
         TabIndex        =   2
         Top             =   480
         Width           =   1575
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
         Left            =   5970
         MaxLength       =   40
         TabIndex        =   3
         Top             =   480
         Width           =   1845
      End
      Begin VB.Label lblFichaFamiliar 
         Caption         =   "      Ficha Familiar"
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
         Left            =   2100
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "  Nº historia clínica           Apellido paterno           Apellido materno        Primer nombre                   "
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
         Width           =   7815
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha movimiento           Ficha Familiar"
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
         Top             =   810
         Width           =   1815
      End
   End
   Begin UltraGrid.SSUltraGrid grdMovimientos 
      Height          =   4230
      Left            =   75
      TabIndex        =   10
      Top             =   2160
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   7461
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de movimientos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Movimiento Historia Clínica"
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
      Left            =   45
      TabIndex        =   11
      Top             =   45
      Width           =   10200
   End
End
Attribute VB_Name = "ucMovimientoHistoriasLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Movimientos de Historia
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoBusqueda As sghTipoBusquedaPrestamoHistoria
Dim mo_Teclado As New sighentidades.Teclado
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idPacienteSeleccionado As Long

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdMovimientos.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdMovimientos.DataSource
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

Property Get idPacienteSeleccionado() As Long
    idPacienteSeleccionado = ml_idPacienteSeleccionado
End Property


Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub
Public Sub RealizarBusqueda()
Dim oPaciente As New doPaciente
Dim oMovimiento As New DOMovimientoHistoriaClinica
        
'        If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
'            UserControl.txtPrimerNombre = "" And UserControl.txtNroHistoria = "") Then
'            MsgBox "Ingrese algunos de los filtro para realizar la búsqueda", vbInformation, "Movimiento de Historias"
'            Exit Sub
'        End If
            
        
        oPaciente.ApellidoMaterno = UserControl.txtApellidoMaterno
        oPaciente.ApellidoPaterno = UserControl.txtApellidoPaterno
        oPaciente.PrimerNombre = UserControl.txtPrimerNombre
        If mo_Teclado.TextoEsSoloNumeros(UserControl.txtNroHistoria) Then
           oPaciente.NroHistoriaClinica = Val(HCigualDNI_AgregaNUEVEaLaHistoria(UserControl.txtNroHistoria))
        End If
        oPaciente.FichaFamiliar = Trim(UserControl.txtFichaFamiliar)
        If cmbFecha.ListIndex = 1 Then
           oMovimiento.FechaMovimiento = 0
        Else
           oMovimiento.FechaMovimiento = cmbFecha.Text
        End If
        
        Set grdMovimientos.DataSource = mo_AdminArchivoClinico.MovimientosHistoriaClinicaFiltrar(oPaciente, oMovimiento, wxSinApellido)
        
        If mo_AdminArchivoClinico.MensajeError <> "" Then
            MsgBox mo_AdminArchivoClinico.MensajeError, vbInformation, "Filtro PrestamosHC"
        End If
        
        'mo_Apariencia.ConfigurarFilasBiColores grdMovimientos, sighentidades.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoMaterno = ""
        UserControl.txtApellidoPaterno = ""
        UserControl.txtPrimerNombre = ""
        UserControl.txtNroHistoria = ""
        UserControl.txtFichaFamiliar.Text = ""
        
        cmbFecha.Clear
        cmbFecha.AddItem Date
        cmbFecha.AddItem "Todas"
        cmbFecha.ListIndex = 0
        
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

Private Sub cmdSinApellidoPaterno_Click()
    txtApellidoPaterno.Text = wxSinApellido
End Sub

Private Sub grdMovimientos_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset
    On Error Resume Next
    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdMovimientos.DataSource
    ml_idRegistroSeleccionado = rsRecordset("IdMovimiento")
    ml_idPacienteSeleccionado = rsRecordset!idPaciente
End Sub

Private Sub grdMovimientos_Click()
Dim rsRecordset As ADODB.Recordset

    On Error Resume Next
    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdMovimientos.DataSource
    ml_idRegistroSeleccionado = rsRecordset("IdMovimiento")
    ml_idPacienteSeleccionado = rsRecordset!idPaciente
        
End Sub

Private Sub grdMovimientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    
    grdMovimientos.Bands(0).Columns("IdMovimiento").Header.Caption = "Id"
    grdMovimientos.Bands(0).Columns("IdMovimiento").Width = 1000
    
    grdMovimientos.Bands(0).Columns("HistoriaClinica").Header.Caption = "Nro Historia"
    grdMovimientos.Bands(0).Columns("HistoriaClinica").Width = 1000
    
    grdMovimientos.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdMovimientos.Bands(0).Columns("ApellidoPaterno").Width = 1200
    
    grdMovimientos.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdMovimientos.Bands(0).Columns("ApellidoMaterno").Width = 1200
    
    grdMovimientos.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdMovimientos.Bands(0).Columns("PrimerNombre").Width = 1200
    
    grdMovimientos.Bands(0).Columns("FechaMovimiento").Header.Caption = "F.Movim"
    grdMovimientos.Bands(0).Columns("FechaMovimiento").Width = 1400
    grdMovimientos.Bands(0).Columns("FechaMovimiento").Format = "dd/mm/yyyy hh:mm"
    
    grdMovimientos.Bands(0).Columns("Origen").Header.Caption = "Origen."
    grdMovimientos.Bands(0).Columns("Origen").Width = 2000
    
    grdMovimientos.Bands(0).Columns("Destino").Header.Caption = "Destino."
    grdMovimientos.Bands(0).Columns("Destino").Width = 2000

    grdMovimientos.Bands(0).Columns("Observacion").Header.Caption = "Observación"
    grdMovimientos.Bands(0).Columns("Observacion").Width = 2500
    
    grdMovimientos.Bands(0).Columns("NroFolios").Header.Caption = "Nº Folio"
    grdMovimientos.Bands(0).Columns("NroFolios").Width = 1000
    
End Sub




Private Sub grdMovimientos_InitializePrintPreview(ByVal PreviewInfo As UltraGrid.SSPreviewInfo)
    PreviewInfo.PrintInfo.PageHeader = "Movimiento de la Historia: " & txtNroHistoria.Text
    PreviewInfo.PrintInfo.PageHeaderAppearance.Font.Name = "Arial"
    PreviewInfo.PrintInfo.PageHeaderAppearance.Font.Size = 30
    PreviewInfo.PrintInfo.Orientation = ssOrientationLandscape
End Sub

Private Sub txtFichaFamiliar_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFichaFamiliar
    AdministrarKeyPreview KeyCode
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

Private Sub txtNroHistoria_LostFocus()
   If Len(txtNroHistoria.Text) > 9 Then
      MsgBox "El Nro Historia no puede exceder de 9 caracteres", vbInformation, "Movimientos HC"
      txtNroHistoria.Text = ""
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
   grdMovimientos.Width = fraBusqueda.Width
   grdMovimientos.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
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
        mo_Apariencia.ConfigurarFilasBiColores grdMovimientos, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdMovimientos, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub
Public Sub Inicializar()
    SkinConfigura
    
    cmbFecha.Clear
    cmbFecha.AddItem Date
    cmbFecha.AddItem "Todas"
    cmbFecha.ListIndex = 0
    '
    If lcBuscaParametro.SeleccionaFilaParametro(277) = "S" Then
        UserControl.txtFichaFamiliar.Visible = True
        UserControl.lblFichaFamiliar.Visible = True
    End If
End Sub

