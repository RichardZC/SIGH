VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form frmBusquedaDistrito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda de Distritos"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "frmBusquedaDistrito.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   30
      TabIndex        =   9
      Top             =   4320
      Width           =   8895
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmBusquedaDistrito.frx":0CCA
         DownPicture     =   "frmBusquedaDistrito.frx":112A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   3240
         Picture         =   "frmBusquedaDistrito.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmBusquedaDistrito.frx":1A14
         DownPicture     =   "frmBusquedaDistrito.frx":1ED8
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   4800
         Picture         =   "frmBusquedaDistrito.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid ugvDistritos 
      Height          =   2775
      Left            =   30
      TabIndex        =   8
      Top             =   1500
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4895
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
      Caption         =   "Listado de Distritos"
   End
   Begin VB.Frame Frame 
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
      Height          =   975
      Left            =   30
      TabIndex        =   0
      Top             =   510
      Width           =   8895
      Begin VB.TextBox txtDistrito 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   6
         Text            =   "txtDistrito"
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7440
         Picture         =   "frmBusquedaDistrito.frx":28B0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Width           =   1305
      End
      Begin VB.ComboBox cboProvincia 
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
         Left            =   2640
         TabIndex        =   5
         Text            =   "cboProvincia"
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox cboDepartamento 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Distrito"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "Departamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Busqueda de Distritos"
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
      TabIndex        =   12
      Top             =   0
      Width           =   10110
   End
End
Attribute VB_Name = "frmBusquedaDistrito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca Distrito
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdDepartamento As Long
Dim ml_IdProvincia As Long
Dim ml_IdDistrito As Long
Dim ms_NomDistrito As String
Dim mo_Teclado As New sighentidades.Teclado
Dim ms_DescripcionDistrito As String
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes

Dim mo_cboDepartamentos As New sighentidades.ListaDespleglable
Dim mo_cboProvincias As New sighentidades.ListaDespleglable
Dim mo_Apariencia As New sighentidades.GridInfragistic

'Valores por defecto para que la busqueda sea mas rapida
Property Let IdDepartamentoPorDefecto(lValue As Long)
    ml_IdDepartamento = lValue
End Property
Property Get IdDepartamentoPorDefecto() As Long
    IdDepartamentoPorDefecto = ml_IdDepartamento
End Property
Property Let IdProvinciaPorDefecto(lValue As Long)
    ml_IdProvincia = lValue
End Property
Property Get IdProvinciaPorDefecto() As Long
    IdProvinciaPorDefecto = ml_IdProvincia
End Property

'====================================================
Property Let DescripcionDistrito(sValue As String)
    ms_DescripcionDistrito = sValue
End Property

Property Get IdDistritoSeleccionado() As Long
    IdDistritoSeleccionado = ml_IdDistrito
End Property
Property Get NombreDistritoSeleccionado() As String
    NombreDistritoSeleccionado = ms_NomDistrito
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Private Sub btnAceptar_Click()
AceptarDistrito
End Sub

Private Sub btnCancelar_Click()
CancelarDistrito
End Sub

Private Sub btnBuscar_Click()
   ' If Me.cboProvincia.ListIndex = -1 Then
   '     MsgBox "Seleccione la provincia", vbInformation, Me.Caption
   '     Exit Sub
   ' End If
    If txtDistrito.Text = "" Then
       MsgBox "Ingrese el DISTRITO a buscar", vbInformation, Me.Caption
       Exit Sub
    End If
    If Me.cboProvincia.ListIndex >= 0 Then
       Set ugvDistritos.DataSource = ObtenerDistrito(Val(Me.cboDepartamento.ItemData(Me.cboDepartamento.ListIndex)), Val(Me.cboProvincia.ItemData(Me.cboProvincia.ListIndex)), txtDistrito.Text)
    Else
       Set ugvDistritos.DataSource = ObtenerDistrito(Val(Me.cboDepartamento.ItemData(Me.cboDepartamento.ListIndex)), 0, txtDistrito.Text)
    End If
End Sub

Private Sub AceptarDistrito()
    ml_IdDistrito = CLng(ugvDistritos.ActiveRow.Cells("IdDistrito").Value)
    ms_NomDistrito = CStr(ugvDistritos.ActiveRow.Cells("Nombre").Value)
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub

Private Sub CancelarDistrito()
    ml_IdDistrito = 0
    ms_NomDistrito = ""
    mi_BotonPresionado = sghCancelar
    Me.Visible = False
End Sub


Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, Me.cboProvincia
    AdministrarKeyPreview KeyCode
End Sub


Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDistrito
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Me.txtDistrito.SetFocus
End Sub

Private Sub Form_Load()
    Set mo_cboDepartamentos.MiComboBox = cboDepartamento
    Set mo_cboProvincias.MiComboBox = cboProvincia
    CargarComboxes
    mo_cboDepartamentos.BoundText = ml_IdDepartamento
    mo_cboProvincias.BoundText = ml_IdProvincia
    txtDistrito.Text = ms_DescripcionDistrito
    'Me.txtDistrito.SetFocus
    'Busqueda de lo que digito el usuario al mandar los valors desde el formulario al dialogo
    If Me.cboProvincia.ListIndex >= 0 Then
      Set Me.ugvDistritos.DataSource = ObtenerDistrito(Val(Me.cboDepartamento.ItemData(Me.cboDepartamento.ListIndex)), Val(Me.cboProvincia.ItemData(Me.cboProvincia.ListIndex)), Me.txtDistrito.Text)
    End If
    mo_Apariencia.ConfigurarFilasBiColores ugvDistritos, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub cboDepartamento_Click()
    If cboDepartamento.ListIndex = -1 Then Exit Sub
    
    mo_cboProvincias.BoundColumn = "IdProvincia"
    mo_cboProvincias.ListField = "Nombre"
    On Error Resume Next
    Set mo_cboProvincias.RowSource = mo_AdminServiciosComunes.ConsultarProvinciasPorDepartamento(Val(Me.cboDepartamento.ItemData(Me.cboDepartamento.ListIndex)))
        
    mo_cboProvincias.BoundText = ""
    txtDistrito.Text = ""
    Me.cboProvincia.Enabled = True
End Sub

'Metodo de Busqueda de Distritos a partir del Departamento, Provincia y una descripcion del Distrito
Private Function ObtenerDistrito(lnIdDepartamento As Long, IdProvincia As Long, DescripcionDistrito As String) As Recordset
    If IdProvincia > 0 Then
       Set ObtenerDistrito = mo_AdminServiciosComunes.SeleccionarPorProvinciaYDescripcion(IdProvincia, DescripcionDistrito)
    Else
       Set ObtenerDistrito = mo_AdminServiciosComunes.DistritosSeleccionarPorDepartamentoYDescripcion(lnIdDepartamento, DescripcionDistrito)
    End If
End Function

Sub CargarComboxes()
    'Listar los Departamentos
    mo_cboDepartamentos.BoundColumn = "IdDepartamento"
    mo_cboDepartamentos.ListField = "Nombre"
    Set mo_cboDepartamentos.RowSource = mo_AdminServiciosComunes.ConsultarDepartamentos
End Sub

Private Sub txtDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDistrito
    AdministrarKeyPreview KeyCode
    If KeyCode = 13 Then
        If Me.cboProvincia.ListIndex >= 0 Then
           Set Me.ugvDistritos.DataSource = ObtenerDistrito(Val(Me.cboDepartamento.ItemData(Me.cboDepartamento.ListIndex)), Val(Me.cboProvincia.ItemData(Me.cboProvincia.ListIndex)), Me.txtDistrito.Text)
        Else
           Set Me.ugvDistritos.DataSource = ObtenerDistrito(Val(Me.cboDepartamento.ItemData(Me.cboDepartamento.ListIndex)), 0, Me.txtDistrito.Text)
        End If
        Me.ugvDistritos.SetFocus
    End If
End Sub

Private Sub ugvDistritos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    With ugvDistritos.Bands(0)
        .Columns("IdDistrito").Header.Caption = "Codigo"
        .Columns("IdDistrito").Width = 1000
        .Columns("Nombre").Header.Caption = "Nombre de Distrito"
        .Columns("Nombre").Width = 5000
        .Columns("IdProvincia").Hidden = True
    End With
End Sub

Private Sub ugvDistritos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
        AceptarDistrito
    ElseIf KeyAscii = vbKeyEscape Then
        CancelarDistrito
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
