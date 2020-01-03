VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form BuscaProveedores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proveedores"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BuscaProveedores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   885
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10005
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   8
         Top             =   480
         Width           =   4515
      End
      Begin VB.TextBox txtRuc 
         Height          =   315
         Left            =   180
         MaxLength       =   15
         TabIndex        =   7
         Top             =   480
         Width           =   1785
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   7920
         Picture         =   "BuscaProveedores.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   6480
         Picture         =   "BuscaProveedores.frx":38A6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "     Ruc                               Razón Social"
         Height          =   345
         Left            =   180
         TabIndex        =   9
         Top             =   240
         Width           =   3795
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   0
      TabIndex        =   0
      Top             =   4800
      Width           =   9975
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "BuscaProveedores.frx":64EF
         DownPicture     =   "BuscaProveedores.frx":694F
         Height          =   700
         Left            =   3563
         Picture         =   "BuscaProveedores.frx":6DC4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "BuscaProveedores.frx":7239
         DownPicture     =   "BuscaProveedores.frx":76FD
         Height          =   700
         Left            =   5123
         Picture         =   "BuscaProveedores.frx":7BE9
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdProveedores 
      Height          =   3810
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   6720
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
      Caption         =   "Lista de Proveedores"
   End
End
Attribute VB_Name = "BuscaProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca Receta
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic

Dim oRsProveedor As New Recordset
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes

Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim oConexion As New ADODB.Connection

Dim mi_BotonPresionado As sghBotonDetallePresionado

Dim rsTmp As New Recordset
Dim ml_ruc As String

Property Get ruc() As String
    ruc = ml_ruc
End Property

Private Sub btnAceptar_Click()
    On Error GoTo errSalir
    ml_ruc = rsTmp.Fields!ruc
errSalir:
    Me.Visible = False
End Sub

Private Sub btnBuscar_Click()
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
If txtRuc.Text = "" Then
    Set rsTmp = oConexion.Execute("select ruc,razonSocial from proveedores where razonSocial like '%" & txtDescripcion.Text & "%'")
Else
    Set rsTmp = oConexion.Execute("select ruc,razonSocial from proveedores where ruc like '%" & Me.txtRuc.Text & "%'")
End If
    Set Me.grdProveedores.DataSource = rsTmp
    mo_Apariencia.ConfigurarFilasBiColores Me.grdProveedores, SIGHEntidades.GrillaConFilasBicolor
    Set oConexion = Nothing
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub btnLimpiar_Click()
    txtRuc.Text = ""
    txtDescripcion.Text = ""
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
'oConexion.Open sighentidades.CadenaConexion
 '   oConexion.CursorLocation = adUseClient
  '  Set rsTmp = oConexion.Execute("select * from proveedores order by RazonSocial asc")
   ' Set Me.grdProveedores.DataSource = rsTmp
    'mo_Apariencia.ConfigurarFilasBiColores Me.grdProveedores, sighentidades.GrillaConFilasBicolor
    'Set oConexion = Nothing
  grdProveedores.ClearAllCachedCells
End Sub

Private Sub grdProveedores_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    'grdProveedores.Bands(0).Columns("idProveedor").Hidden = True
    '
    grdProveedores.Bands(0).Columns("Ruc").Header.Caption = "Ruc"
    grdProveedores.Bands(0).Columns("Ruc").Width = 2000
    '
    grdProveedores.Bands(0).Columns("RazonSocial").Header.Caption = "Razon Social"
    grdProveedores.Bands(0).Columns("RazonSocial").Width = 6500
End Sub

Private Sub txtDescripcion_GotFocus()
    txtRuc.Text = ""
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    btnBuscar.SetFocus
End If
End Sub



Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    btnBuscar.SetFocus
End If
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            btnBuscar_Click
        Case vbKeyEscape
            btnCancelar_Click
        Case vbKeyF7
            btnLimpiar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

