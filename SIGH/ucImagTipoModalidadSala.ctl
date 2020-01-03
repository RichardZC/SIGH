VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucImagTipoModalidadSala 
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10035
   LockControls    =   -1  'True
   ScaleHeight     =   6270
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
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   9930
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   5190
         Picture         =   "ucImagTipoModalidadSala.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   6555
         Picture         =   "ucImagTipoModalidadSala.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   450
         Width           =   1275
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
         Left            =   120
         MaxLength       =   5
         TabIndex        =   0
         Top             =   480
         Width           =   1125
      End
      Begin VB.TextBox txtDescripcion 
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
         MaxLength       =   50
         TabIndex        =   1
         Top             =   480
         Width           =   3585
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
         Caption         =   "Código           Descripción"
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
         TabIndex        =   7
         Top             =   210
         Width           =   6975
      End
   End
   Begin UltraGrid.SSUltraGrid grdTipoModalidadSala 
      Height          =   4590
      Left            =   60
      TabIndex        =   4
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
      Caption         =   "Lista de Tipo Modalidad Sala"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Listado de Tipo Modalidad Sala"
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
      TabIndex        =   5
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "ucImagTipoModalidadSala"
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
Dim mo_reglaImagen As New SIGHNegocios.ReglasImagenes
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdTipoModalidadSala.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdTipoModalidadSala.DataSource
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
    Dim oDOImagTipoModalidadSala As New DOImagTipoModalidadSala
    oDOImagTipoModalidadSala.Codigo = txtCodigo.Text
    oDOImagTipoModalidadSala.TipoModalidadSala = txtDescripcion.Text

    Set grdTipoModalidadSala.DataSource = mo_reglaImagen.ImagTipoModalidadSalaFiltrarTodos(oDOImagTipoModalidadSala)
    If mo_reglaImagen.MensajeError <> "" Then
        MsgBox mo_reglaImagen.MensajeError, vbInformation, "Filtro de tipo de modalidad de sala"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdTipoModalidadSala, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
    UserControl.txtCodigo = ""
    UserControl.txtDescripcion = ""
End Sub

Private Sub grdTipoModalidadSala_AfterRowActivate()
    Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdTipoModalidadSala.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdTipoModalidadSala")
    

End Sub

Private Sub grdTipoModalidadSala_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdTipoModalidadSala.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdTipoModalidadSala")
    
End Sub


Private Sub grdTipoModalidadSala_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdTipoModalidadSala.Override.AllowDelete = ssAllowDeleteNo
    
    grdTipoModalidadSala.Bands(0).Columns("IdTipoModalidadSala").Hidden = True
    grdTipoModalidadSala.Bands(0).Columns("EsActivo").Hidden = True
    grdTipoModalidadSala.Bands(0).Columns("FechaCrea").Hidden = True
    grdTipoModalidadSala.Bands(0).Columns("FechaEdita").Hidden = True
    
    
    grdTipoModalidadSala.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdTipoModalidadSala.Bands(0).Columns("Codigo").Width = 700
    
    grdTipoModalidadSala.Bands(0).Columns("TipoModalidadSala").Header.Caption = "Descripción"
    grdTipoModalidadSala.Bands(0).Columns("TipoModalidadSala").Width = 6000
    
   

End Sub
Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, UserControl.txtDescripcion
End Sub
Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
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
   
   grdTipoModalidadSala.Width = fraBusqueda.Width
   grdTipoModalidadSala.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub

Public Function inicializar()
    RealizarBusqueda
End Function
