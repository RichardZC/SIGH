VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucSupervisorLista 
   ClientHeight    =   6225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10035
   ScaleHeight     =   6225
   ScaleWidth      =   10035
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
      Height          =   945
      Left            =   45
      TabIndex        =   7
      Top             =   510
      Width           =   9930
      Begin VB.TextBox txtNombres 
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
         Left            =   4560
         MaxLength       =   30
         TabIndex        =   3
         Top             =   480
         Width           =   2145
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7110
         Picture         =   "ucSupervisorLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   450
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8475
         Picture         =   "ucSupervisorLista.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   1275
      End
      Begin VB.TextBox txtApPaterno 
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
         TabIndex        =   1
         Top             =   480
         Width           =   2145
      End
      Begin VB.TextBox txtApMaterno 
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
         Left            =   2340
         MaxLength       =   30
         TabIndex        =   2
         Top             =   480
         Width           =   2145
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   8
         Top             =   810
         Width           =   7635
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ap. Paterno                    Ap. Materno                    Nombres"
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
         TabIndex        =   0
         Top             =   210
         Width           =   6975
      End
   End
   Begin UltraGrid.SSUltraGrid grdSupervisores 
      Height          =   4590
      Left            =   60
      TabIndex        =   6
      Top             =   1545
      Width           =   9930
      _ExtentX        =   17515
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
      Caption         =   "Lista de Supervisores"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Listado de Supervisores"
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
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "ucSupervisorLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'MZD Ini 19/06/2005 [Todo el archivo]
Option Explicit
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim ml_IdRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Apariencia As New SIGHComun.GridInfragistic

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdSupervisores.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdSupervisores.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
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
    Dim oEmpleado As New dOEmpleado
    
    oEmpleado.ApellidoPaterno = UserControl.txtApPaterno
    oEmpleado.ApellidoMaterno = UserControl.txtApMaterno
    oEmpleado.Nombres = UserControl.txtNombres
    
    
    Set grdSupervisores.DataSource = mo_AdminCaja.RealizarFiltroSupervisores(oEmpleado)
    If mo_AdminCaja.MensajeError <> "" Then
        MsgBox mo_AdminCaja.MensajeError, vbCritical, "Filtro de supervisores"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdSupervisores, SIGHComun.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
    UserControl.txtApPaterno = ""
    UserControl.txtApMaterno = ""
    UserControl.txtNombres = ""
End Sub

Private Sub grdSupervisores_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdSupervisores.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdSupervisor")
    

End Sub

Private Sub grdSupervisores_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdSupervisores.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdSupervisor")
    
End Sub


Private Sub grdSupervisores_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdSupervisores.Bands(0).Columns("IdSupervisor").Hidden = True
    grdSupervisores.Bands(0).Columns("IdEmpleado").Hidden = True
    
    grdSupervisores.Bands(0).Columns("EstadoSupervisor").Header.Caption = "Activo"
    grdSupervisores.Bands(0).Columns("EstadoSupervisor").Width = 500
    
    grdSupervisores.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap.Paterno"
    grdSupervisores.Bands(0).Columns("ApellidoPaterno").Width = 2000
    
    grdSupervisores.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap.Materno"
    grdSupervisores.Bands(0).Columns("ApellidoMaterno").Width = 2000
    
    grdSupervisores.Bands(0).Columns("Nombres").Header.Caption = "Nombres"
    grdSupervisores.Bands(0).Columns("Nombres").Width = 3000

End Sub
Private Sub txtApPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, UserControl.txtApMaterno
End Sub
Private Sub txtApMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, UserControl.txtNombres
End Sub
Private Sub txtNombres_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, UserControl.btnBuscar
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
   
   grdSupervisores.Width = fraBusqueda.Width
   grdSupervisores.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub








