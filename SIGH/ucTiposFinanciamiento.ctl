VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucTiposFinanciamientoLista 
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10125
   ScaleHeight     =   5985
   ScaleWidth      =   10125
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
      Height          =   705
      Left            =   75
      TabIndex        =   3
      Top             =   525
      Width           =   10035
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   4140
         Picture         =   "ucTiposFinanciamiento.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1305
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
         Left            =   1350
         TabIndex        =   0
         Top             =   240
         Width           =   2715
      End
      Begin VB.Label Label2 
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
         Height          =   345
         Left            =   300
         TabIndex        =   4
         Top             =   270
         Width           =   675
      End
   End
   Begin UltraGrid.SSUltraGrid grdTiposFinanciamiento 
      Height          =   4665
      Left            =   75
      TabIndex        =   2
      Top             =   1320
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   8229
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
      Caption         =   "Lista de Producto/Plan"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Producto/Plan"
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
      Width           =   10110
   End
End
Attribute VB_Name = "ucTiposFinanciamientoLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de tipos de Financiamientos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim ml_idRegistroSeleccionado As Long
Dim mo_Apariencia As New sighentidades.GridInfragistic
Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdTiposFinanciamiento.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdTiposFinanciamiento.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property

Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub
Public Sub RealizarBusqueda()
        Set grdTiposFinanciamiento.DataSource = mo_AdminFacturacion.TiposFinanciamientoSeleccionarTodos()
        'mo_Apariencia.ConfigurarFilasBiColores grdTiposFinanciamiento, sighentidades.GrillaConFilasBicolor

End Sub

Private Sub grdTiposFinanciamiento_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdTiposFinanciamiento.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdTipoFinanciamiento")
End Sub

Private Sub grdTiposFinanciamiento_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdTiposFinanciamiento.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdTipoFinanciamiento")
    
End Sub


Private Sub grdTiposFinanciamiento_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdTiposFinanciamiento.Bands(0).Columns("IdTipoFinanciamiento").Header.Caption = "ID"
    grdTiposFinanciamiento.Bands(0).Columns("IdTipoFinanciamiento").Width = 500

    grdTiposFinanciamiento.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdTiposFinanciamiento.Bands(0).Columns("Descripcion").Width = 5000

    grdTiposFinanciamiento.Bands(0).Columns("DescripcionLarga").Header.Caption = "Descripción Larga"
    grdTiposFinanciamiento.Bands(0).Columns("DescripcionLarga").Width = 5000


End Sub

Private Sub UserControl_Initialize()
    'mo_Formulario.ConfigurarTipoLetraDeControles UserControl.Controls
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdTiposFinanciamiento.Width = fraBusqueda.Width
   grdTiposFinanciamiento.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
'        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
'        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdTiposFinanciamiento, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdTiposFinanciamiento, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub
Sub Inicializar()
    SkinConfigura
End Sub
