VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucArchivadoresLista 
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10245
   LockControls    =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   10245
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
      Height          =   915
      Left            =   90
      TabIndex        =   7
      Top             =   555
      Width           =   10035
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7260
         Picture         =   "ucArchivadoresLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   450
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8625
         Picture         =   "ucArchivadoresLista.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   1275
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   5325
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
         Left            =   3420
         MaxLength       =   40
         TabIndex        =   2
         Top             =   465
         Width           =   1845
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
         Left            =   1500
         MaxLength       =   40
         TabIndex        =   1
         Top             =   465
         Width           =   1845
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   135
         TabIndex        =   0
         Top             =   465
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Código Planilla        Apellido paterno          Apellido materno              Nombre                   "
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
         TabIndex        =   8
         Top             =   225
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdArchiveroServicio 
      Height          =   4860
      Left            =   90
      TabIndex        =   6
      Top             =   1560
      Width           =   10035
      _ExtentX        =   17701
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
      Caption         =   "Lista de responsables"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Archiveros"
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
      TabIndex        =   9
      Top             =   45
      Width           =   10200
   End
End
Attribute VB_Name = "ucArchivadoresLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para Lista archivadores
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
Dim ml_EsConsultorioAsignado As Boolean

Property Let EsConsultorioAsignado(lValue As Boolean)
    ml_EsConsultorioAsignado = lValue
    If ml_EsConsultorioAsignado = True Then
       lblNombre.Caption = "Empleados para CITAS"
       
    Else
       lblNombre.Caption = "Archiveros"
    End If
End Property


Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdArchiveroServicio.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdArchiveroServicio.DataSource
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


Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
Dim oEmpleado As New dOEmpleado
Dim oArchivero As New DOArchiveroServicio
        
        If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
            UserControl.txtNombre = "" And UserControl.txtCodigo = "") Then
        End If
            
        
        oEmpleado.ApellidoMaterno = UserControl.txtApellidoMaterno
        oEmpleado.ApellidoPaterno = UserControl.txtApellidoPaterno
        oEmpleado.Nombres = UserControl.txtNombre
        oEmpleado.CodigoPlanilla = UserControl.txtCodigo
        
        Set grdArchiveroServicio.DataSource = mo_AdminArchivoClinico.ArchiveroServicioFiltrar(oEmpleado, ml_EsConsultorioAsignado)
        
        mo_Apariencia.ConfigurarFilasBiColores grdArchiveroServicio, sighentidades.GrillaConFilasBicolor
        

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoMaterno = ""
        UserControl.txtApellidoPaterno = ""
        UserControl.txtNombre = ""
        UserControl.txtCodigo = ""
End Sub

Private Sub grdArchiveroServicio_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdArchiveroServicio.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdEmpleado")
    
End Sub

Private Sub grdArchiveroServicio_Click()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdArchiveroServicio.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdEmpleado")
    
End Sub


Private Sub grdArchiveroServicio_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdArchiveroServicio.Bands(0).Columns("IdEmpleado").Hidden = True
    
    grdArchiveroServicio.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdArchiveroServicio.Bands(0).Columns("ApellidoPaterno").Width = 2000
    
    grdArchiveroServicio.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdArchiveroServicio.Bands(0).Columns("ApellidoMaterno").Width = 2000
    
    grdArchiveroServicio.Bands(0).Columns("Nombres").Header.Caption = "Nombres"
    grdArchiveroServicio.Bands(0).Columns("Nombres").Width = 2000
    
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombre
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNombre_KeyPress(KeyAscii As Integer)
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
   grdArchiveroServicio.Width = fraBusqueda.Width
   grdArchiveroServicio.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
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
        mo_Apariencia.ConfigurarFilasBiColores grdArchiveroServicio, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdArchiveroServicio, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub
Sub Inicializar()
    SkinConfigura
End Sub




