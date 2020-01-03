VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucInteoIntegracionSistema 
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11040
   LockControls    =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   11040
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
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   10890
      Begin VB.ComboBox cmbIdTipoSistemaSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   4965
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   7035
         Picture         =   "ucInteoIntegracionSistema.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   5670
         Picture         =   "ucInteoIntegracionSistema.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Sistema"
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
         TabIndex        =   6
         Top             =   210
         Width           =   6975
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
         TabIndex        =   5
         Top             =   810
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdIntegracionSistema 
      Height          =   4590
      Left            =   60
      TabIndex        =   3
      Top             =   1545
      Width           =   10890
      _ExtentX        =   19209
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
      Caption         =   "Lista de Tipo de Sistemas"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Listado de Integraciones con otros Sistema"
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
      TabIndex        =   7
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "ucInteoIntegracionSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para listar modalidades de salas
'        Programado por: Garay M.
'        Fecha: Noviembre 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mo_ReglasIntegracionSistema As New SIGHIntegracion.ReglasIntegracion
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_cmbIdTipoSistema As New sighentidades.ListaDespleglable

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdIntegracionSistema.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdIntegracionSistema.DataSource
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
    Dim oDOInteoIntegracionSistema As New DOInteoIntegracionSistema
    oDOInteoIntegracionSistema.IdTipoSistema = Val(mo_cmbIdTipoSistema.BoundText)
    

    Set grdIntegracionSistema.DataSource = mo_ReglasIntegracionSistema.InteoIntegracionSistemaFiltrarTodos(oDOInteoIntegracionSistema)
    If mo_ReglasIntegracionSistema.MensajeError <> "" Then
        MsgBox mo_ReglasIntegracionSistema.MensajeError, vbInformation, "Filtro de Integración de sistemas"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdIntegracionSistema, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
'    UserControl.txtCodigo = ""
'    UserControl.txtDescripcion = ""
    mo_cmbIdTipoSistema.BoundText = ""
End Sub

Private Sub cmbIdTipoSistemaSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, UserControl.cmbIdTipoSistemaSearch
End Sub

Private Sub grdIntegracionSistema_AfterRowActivate()
    Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdIntegracionSistema.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdIntegracionSistema")
    

End Sub

Private Sub grdIntegracionSistema_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdIntegracionSistema.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdIntegracionSistema")
    
End Sub


Private Sub grdIntegracionSistema_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdIntegracionSistema.Override.AllowDelete = ssAllowDeleteNo
    
    grdIntegracionSistema.Bands(0).Columns("IdIntegracionSistema").Hidden = True
    grdIntegracionSistema.Bands(0).Columns("IdTipoSistema").Hidden = True
    grdIntegracionSistema.Bands(0).Columns("IdProveedorSistema").Hidden = True
    grdIntegracionSistema.Bands(0).Columns("NombreUsuario").Hidden = True
    
    
    grdIntegracionSistema.Bands(0).Columns("EsActivo").Hidden = True
    grdIntegracionSistema.Bands(0).Columns("FechaCrea").Hidden = True
    grdIntegracionSistema.Bands(0).Columns("FechaEdita").Hidden = True
    
    
    grdIntegracionSistema.Bands(0).Columns("TipoSistema").Header.Caption = "Tipo Sistema"
    grdIntegracionSistema.Bands(0).Columns("TipoSistema").Width = 3500
    
    grdIntegracionSistema.Bands(0).Columns("ProveedorSistema").Header.Caption = "Proveedor"
    grdIntegracionSistema.Bands(0).Columns("ProveedorSistema").Width = 3500
    
    grdIntegracionSistema.Bands(0).Columns("EsProveedorActual").Header.Caption = "Proveedor Actual"
    grdIntegracionSistema.Bands(0).Columns("EsProveedorActual").Width = 1000

End Sub
'Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
'    mo_Teclado.RealizarNavegacion KeyCode, UserControl.txtDescripcion
'End Sub
'Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
'    mo_Teclado.RealizarNavegacion KeyCode, UserControl.btnBuscar
'End Sub
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
   
   grdIntegracionSistema.Width = fraBusqueda.Width
   grdIntegracionSistema.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub

Public Function inicializar()
    CargaComboBoxes
    RealizarBusqueda
End Function

Sub CargaComboBoxes()
    Set mo_cmbIdTipoSistema.MiComboBox = cmbIdTipoSistemaSearch
    
    mo_cmbIdTipoSistema.BoundColumn = "IdTipoSistema"
    mo_cmbIdTipoSistema.ListField = "TipoSistema"
    Set mo_cmbIdTipoSistema.RowSource = mo_ReglasIntegracionSistema.TipoSistemaSeleccionarTodos()
End Sub
