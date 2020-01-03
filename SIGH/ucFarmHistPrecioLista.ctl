VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucFarmHistPrecioLista 
   ClientHeight    =   6450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11055
   ScaleHeight     =   6450
   ScaleWidth      =   11055
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
      Height          =   750
      Left            =   30
      TabIndex        =   0
      Top             =   510
      Width           =   10965
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
         Left            =   795
         MaxLength       =   30
         TabIndex        =   4
         Top             =   285
         Width           =   1125
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   8220
         Picture         =   "ucFarmHistPrecioLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   285
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9585
         Picture         =   "ucFarmHistPrecioLista.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   285
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
         Left            =   3165
         MaxLength       =   30
         TabIndex        =   1
         Top             =   270
         Width           =   4470
      End
      Begin VB.Label lblNcuenta 
         Caption         =   "Código"
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
         Left            =   105
         TabIndex        =   6
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2475
         TabIndex        =   5
         Top             =   315
         Width           =   645
      End
   End
   Begin UltraGrid.SSUltraGrid grdLista 
      Height          =   5040
      Left            =   30
      TabIndex        =   7
      Top             =   1320
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   8890
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
      Caption         =   "Lista "
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00000000&
      Caption         =   "Histórico de Precios"
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
      Width           =   11010
   End
End
Attribute VB_Name = "ucFarmHistPrecioLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para buscar Historico de PRecios
'        Programado por: Barrantes D
'        Fecha: Noviembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoBusqueda As sghTipoBusquedaPrestamoHistoria
Dim mo_Teclado As New sighEntidades.Teclado
Dim oRsAlmacenes As New ADODB.Recordset
Dim oRsBusqueda As New ADODB.Recordset
Dim ml_idUsuario As Long

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdLista.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdLista.DataSource
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
        Set oRsBusqueda = mo_ReglasFarmacia.FarmHistPrecioSeleccionarPorCodigo(txtCodigo.Text, txtNombre.Text)
        Set grdLista.DataSource = oRsBusqueda
        mo_Apariencia.ConfigurarFilasBiColores grdLista, sighEntidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtCodigo.Text = ""
        UserControl.txtNombre.Text = ""
End Sub




Private Sub grdLista_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdLista.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = Val(rsRecordset("idHistPrecio"))
    
End Sub

Private Sub grdLista_Click()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdLista.DataSource
    On Error Resume Next
        ml_idRegistroSeleccionado = Val(rsRecordset("idHistPrecio"))
    
End Sub


Private Sub grdLista_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
'    grdLista.Bands(0).Columns("MovNumero").Hidden = True
'    grdLista.Bands(0).Columns("MovNumero").Header.Caption = "Nota Salida"
'    grdLista.Bands(0).Columns("MovNumero").Width = 1300
'    grdLista.Bands(0).Columns("fechaCreacion").Header.Caption = "Fecha"
'    grdLista.Bands(0).Columns("Total").Width = 1300
'    grdLista.Bands(0).Columns("Total").Format = "#0.00"

End Sub







Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
        AdministrarKeyPreview KeyCode
End Sub





Sub CargaComboBox()

ErrFarm:
End Sub


Sub inicializar()
'    RealizarBusqueda
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
   lblNombre.Width = UserControl.Width

   grdLista.Width = UserControl.Width
   grdLista.Height = UserControl.Height - (lblNombre.Height + 150)
   
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub



