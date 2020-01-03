VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucCatalogoServiciosListaBus 
   ClientHeight    =   8325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10035
   ScaleHeight     =   8325
   ScaleWidth      =   10035
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
      Height          =   1185
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   9900
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   3000
         TabIndex        =   5
         Top             =   690
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
         Left            =   1680
         TabIndex        =   4
         Top             =   270
         Width           =   5385
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8550
         Picture         =   "ucCatalogoServiciosListaBus.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   690
         Width           =   1215
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7140
         Picture         =   "ucCatalogoServiciosListaBus.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   690
         Width           =   1305
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   690
         Width           =   1245
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
         TabIndex        =   7
         Top             =   750
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de cátalogo"
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
         TabIndex        =   6
         Top             =   360
         Width           =   1545
      End
   End
   Begin UltraGrid.SSUltraGrid grdServicios 
      Height          =   6420
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   11324
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
      Override        =   "ucCatalogoServiciosListaBus.ctx":5825
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
      TabIndex        =   9
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "ucCatalogoServiciosListaBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim ml_IdRegistroSeleccionado As Long
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim mo_cmbIdTipoCatalogo As New ListaDespleglable
Dim ml_IdDepartamentoHospital As Long
Dim ml_IdTipoCatalogo As Long

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
 Dim oDOCatalogoServicios As New DOCatalogoServicio
    
    oDOCatalogoServicios.codigo = Trim(txtCodigo.Text)
    oDOCatalogoServicios.Nombre = Trim(txtNombre.Text)
        
    Set grdServicios.DataSource = mo_AdminComun.CatalogoServiciosFiltrar(oDOCatalogoServicios, ml_IdTipoCatalogo)
    ConfigurarGrilla ml_IdTipoCatalogo = 0
    
    If mo_AdminComun.MensajeError <> "" Then
        MsgBox mo_AdminComun.MensajeError, vbCritical, "Busqueda del catálogo de servicios"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdServicios, SIGHComun.GrillaConFilasBicolor
End Sub

Private Sub cmbIdTipoCatalogo_Click()

    ml_IdTipoCatalogo = Val(mo_cmbIdTipoCatalogo.BoundText)
    
    RealizarBusqueda
End Sub

Private Sub grdServicios_BeforeRowActivate(ByVal Row As UltraGrid.SSRow)
    ml_IdRegistroSeleccionado = Row.Cells(1).Value
End Sub

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    ConfigurarGrilla False
End Sub
Sub ConfigurarGrilla(lCatalogoBase As Boolean)
    
    'grdServicios.Bands(0).Columns("IdClasificacionServicio").Hidden = True
    grdServicios.Bands(0).Columns("IdProducto").Hidden = True
    
    grdServicios.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdServicios.Bands(0).Columns("Codigo").Width = 1200

    grdServicios.Bands(0).Columns("Descripcion").Header.Caption = "Nombre"
    grdServicios.Bands(0).Columns("Descripcion").Width = IIf(lCatalogoBase, 10500, 8500)

    'grdServicios.Bands(1).Columns("IdClasificacionServicio").Hidden = True
    grdServicios.Bands(1).Columns("IdProducto").Hidden = True

    grdServicios.Bands(1).Columns("Codigo").Header.Caption = "Código"
    grdServicios.Bands(1).Columns("Codigo").Width = 1200

    grdServicios.Bands(1).Columns("Nombre").Header.Caption = "Nombre"
    grdServicios.Bands(1).Columns("Nombre").Width = 6100

    grdServicios.Bands(1).Columns("PrecioUnitario").Header.Caption = "Precio Unitario (S/.)"
    grdServicios.Bands(1).Columns("PrecioUnitario").Width = 2000
    grdServicios.Bands(1).Columns("PrecioUnitario").Hidden = lCatalogoBase

    grdServicios.Bands(1).Columns("Activo").Header.Caption = "Activo"
    grdServicios.Bands(1).Columns("Activo").Width = 2500
    grdServicios.Bands(1).Columns("Activo").Hidden = lCatalogoBase

End Sub


Public Function Inicializar()
    Set mo_cmbIdTipoCatalogo.MiComboBox = cmbIdTipoCatalogo
    
End Function
Public Function SeleccionaTipoCatalogo()
    mo_cmbIdTipoCatalogo.BoundText = ml_IdTipoCatalogo
End Function
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

