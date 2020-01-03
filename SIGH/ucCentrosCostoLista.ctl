VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucCentrosCostoLista 
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   ScaleHeight     =   6240
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
      Height          =   945
      Left            =   45
      TabIndex        =   5
      Top             =   525
      Width           =   9930
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   480
         Width           =   1515
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         Top             =   480
         Width           =   5295
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8535
         Picture         =   "ucCentrosCostoLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   450
         Width           =   1215
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7170
         Picture         =   "ucCentrosCostoLista.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código                  Nombre"
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
         Top             =   270
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
         TabIndex        =   6
         Top             =   810
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdCentrosCosto 
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
      Caption         =   "Lista de Centros de Costo"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00808080&
      Caption         =   "Centros de Costo"
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
      TabIndex        =   8
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "ucCentrosCostoLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para buscar centro de costos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdCentrosCosto.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdCentrosCosto.DataSource
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
    Dim oDOCentrosCosto As New DOCentrosCosto
    
    
    oDOCentrosCosto.Descripcion = txtDescripcion.Text
    oDOCentrosCosto.Codigo = txtCodigo.Text
        
    Set grdCentrosCosto.DataSource = mo_AdminComun.CentrosCostoFiltrar(oDOCentrosCosto)
    If mo_AdminComun.MensajeError <> "" Then
        MsgBox mo_AdminComun.MensajeError, vbInformation, "Filtro de Centros de Costo"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdCentrosCosto, SIGHEntidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
    UserControl.txtDescripcion = ""
    UserControl.txtCodigo = ""
End Sub

Private Sub Command1_Click()
    grdCentrosCosto_InitializeLayout ssContextDisplay, Nothing
End Sub

Private Sub grdCentrosCosto_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdCentrosCosto.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdCentroCosto")

End Sub

Private Sub grdCentrosCosto_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdCentrosCosto.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdCentroCosto")
    
End Sub


Private Sub grdCentrosCosto_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdCentrosCosto.Bands(0).Columns("IdCentroCosto").Hidden = True
    
    grdCentrosCosto.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdCentrosCosto.Bands(0).Columns("Codigo").Width = 1200
   
    grdCentrosCosto.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdCentrosCosto.Bands(0).Columns("Descripcion").Width = 11000
    
End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
    AdministrarKeyPreview KeyCode

End Sub

Private Sub UserControl_Initialize()
    'CargarComboBoxes
End Sub

Public Function inicializar()
    CargarComboBoxes
End Function

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
   
   grdCentrosCosto.Width = fraBusqueda.Width
   grdCentrosCosto.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub
Private Sub CargarComboBoxes()

End Sub






