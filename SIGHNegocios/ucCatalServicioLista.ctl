VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucCatalServicioLista 
   ClientHeight    =   8280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9990
   ScaleHeight     =   8280
   ScaleWidth      =   9990
   Begin VB.Frame fraBusqueda 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   60
      TabIndex        =   5
      Top             =   540
      Width           =   9900
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   930
         TabIndex        =   0
         Top             =   210
         Width           =   1245
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7140
         Picture         =   "ucCatalServicioLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8550
         Picture         =   "ucCatalServicioLista.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   2250
         TabIndex        =   1
         Top             =   210
         Width           =   4035
      End
      Begin VB.ComboBox cmbIdTipoCatalogo 
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
         Left            =   6360
         TabIndex        =   2
         Top             =   210
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Height          =   315
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Width           =   795
      End
   End
   Begin UltraGrid.SSUltraGrid grdServicios 
      Height          =   6960
      Left            =   60
      TabIndex        =   7
      Top             =   1260
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   12277
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   68157460
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Override        =   "ucCatalServicioLista.ctx":5825
      Caption         =   "Lista de Servicios"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00808080&
      Caption         =   "Busqueda de Catálogo de Servicios"
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
Attribute VB_Name = "ucCatalServicioLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para Listar Servicios
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Public Event SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim ml_IdRegistroSeleccionado As Long
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_cmbIdTipoCatalogo As New ListaDespleglable
Dim ml_IdDepartamentoHospital As Long
Dim ml_IdTipoCatalogo As Long
Dim mo_Teclado As New sighentidades.Teclado

Property Let IdTipoCatalogo(lValue As Long)
    ml_IdTipoCatalogo = lValue
End Property
Property Get IdTipoCatalogo() As Long
    IdTipoCatalogo = ml_IdTipoCatalogo
End Property
Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdServicios.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdServicios.DataSource
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
Property Let HabilitarTipoCatalogo(lValue As Boolean)
    cmbIdTipoCatalogo.Enabled = lValue
End Property
Property Get HabilitarTipoCatalogo() As Boolean
    HabilitarTipoCatalogo = cmbIdTipoCatalogo.Enabled
End Property

Property Let IdDepartamentoHospital(lValue As Long)
    ml_IdDepartamentoHospital = lValue
End Property
Property Get IdDepartamentoHospital() As Long
    IdDepartamentoHospital = ml_IdDepartamentoHospital
End Property


Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub
Public Sub RealizarBusqueda()
    If Trim(txtCodigo.Text) = "" And Trim(txtNombre.Text) = "" Then
        MsgBox "Tiene que registrar el Código o Parte del Nombre", vbInformation, ""
        Exit Sub
    End If
    If Trim(txtCodigo.Text) <> "" Then
       txtNombre.Text = ""
    Else
       txtCodigo.Text = ""
    End If
    Set grdServicios.DataSource = mo_ReglasCaja.FactCatalogoServiciosSeleccionarPorCodigoOnombreTipo(txtCodigo.Text, txtNombre.Text, ml_IdTipoCatalogo)
    If mo_ReglasCaja.MensajeError <> "" Then
        MsgBox mo_ReglasCaja.MensajeError, vbInformation, "Busqueda del catálogo de servicios"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdServicios, sighentidades.GrillaConFilasBicolor
End Sub



Private Sub btnLimpiar_Click()
    UserControl.txtCodigo.Text = ""
    UserControl.txtNombre.Text = ""
End Sub

Private Sub grdServicios_BeforeRowActivate(ByVal Row As UltraGrid.SSRow)
    ml_IdRegistroSeleccionado = Row.Cells(0).value
End Sub

Private Sub grdServicios_DblClick()
    'ml_IdRegistroSeleccionado = Row.Cells(1).Value
    RaiseEvent SeleccionaRegistro(ml_IdRegistroSeleccionado)
End Sub

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    ConfigurarGrilla
End Sub
Sub ConfigurarGrilla()
    
    grdServicios.Bands(0).Columns("IdProducto").Hidden = True

    grdServicios.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdServicios.Bands(0).Columns("Codigo").Width = 1200

    grdServicios.Bands(0).Columns("Nombre").Header.Caption = "Nombre"
    grdServicios.Bands(0).Columns("Nombre").Width = 10100


End Sub



Public Function SeleccionaTipoCatalogo()
    mo_cmbIdTipoCatalogo.BoundText = ml_IdTipoCatalogo
End Function

Private Sub grdServicios_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
   If KeyAscii = 13 Then
      grdServicios_DblClick
   End If
End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode
End Sub



Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombre
    AdministrarKeyPreview KeyCode
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdServicios.Width = fraBusqueda.Width
   grdServicios.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub

Public Sub ConfigurarTiposCatalogos()
    
    mo_cmbIdTipoCatalogo.BoundColumn = "IdTipoFinanciamiento"
    mo_cmbIdTipoCatalogo.ListField = "DescripcionLarga"
    Set mo_cmbIdTipoCatalogo.RowSource = mo_AdminFacturacion.TiposFinanciamientoSeleccionarTodos()
    
End Sub

Public Function inicializar()
    Set mo_cmbIdTipoCatalogo.MiComboBox = cmbIdTipoCatalogo
    
End Function

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
