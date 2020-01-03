VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucCatBienesInsumosListaBus 
   ClientHeight    =   8310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10020
   ScaleHeight     =   8310
   ScaleWidth      =   10020
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
      Left            =   0
      TabIndex        =   6
      Top             =   540
      Width           =   9900
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   690
         Width           =   1245
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7140
         Picture         =   "ucCatalogoBienesInsumosListaBus.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   690
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8550
         Picture         =   "ucCatalogoBienesInsumosListaBus.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   690
         Width           =   1215
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
         TabIndex        =   0
         Top             =   270
         Width           =   5385
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   690
         Width           =   4035
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
         TabIndex        =   8
         Top             =   360
         Width           =   1545
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
   End
   Begin UltraGrid.SSUltraGrid grdBienes 
      Height          =   6420
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   9900
      _ExtentX        =   17463
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
      Override        =   "ucCatalogoBienesInsumosListaBus.ctx":5825
      Caption         =   "Lista de Bienes e Insumos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00808080&
      Caption         =   "Busqueda de Catálogo de Bienes e Insumos"
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
Attribute VB_Name = "ucCatBienesInsumosListaBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para buscar medicamentos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
'Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim ml_idRegistroSeleccionado As Long
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_cmbIdTipoCatalogo As New ListaDespleglable
Dim ml_IdTipoCatalogo As Long
Dim mo_Teclado As New sighentidades.Teclado

Property Let IdTipoCatalogo(lValue As Long)
    ml_IdTipoCatalogo = lValue
End Property
Property Get IdTipoCatalogo() As Long
    IdTipoCatalogo = ml_IdTipoCatalogo
End Property
Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdBienes.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdBienes.DataSource
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

Property Let HabilitarTipoCatalogo(lValue As Boolean)
    cmbIdTipoCatalogo.Enabled = lValue
End Property
Property Get HabilitarTipoCatalogo() As Boolean
    HabilitarTipoCatalogo = cmbIdTipoCatalogo.Enabled
End Property


Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub
Public Sub RealizarBusqueda()
 Dim oDOCatalogoBienes As New DOCatalogoBienesInsumos
    
    oDOCatalogoBienes.Codigo = Trim(txtCodigo.Text)
    oDOCatalogoBienes.nombre = Trim(txtNombre.Text)
        
    Set grdBienes.DataSource = mo_AdminComun.CatalogoBienesInsumosFiltrar(oDOCatalogoBienes, ml_IdTipoCatalogo)
    ConfigurarGrilla ml_IdTipoCatalogo = 0
    
    If mo_AdminComun.MensajeError <> "" Then
        MsgBox mo_AdminComun.MensajeError, vbInformation, "Busqueda del catálogo de Bienes e Insumos"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdBienes, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    txtCodigo.Text = ""
    txtNombre.Text = ""
End Sub

Private Sub cmbIdTipoCatalogo_Click()

    ml_IdTipoCatalogo = Val(mo_cmbIdTipoCatalogo.BoundText)
    
    RealizarBusqueda
End Sub

Private Sub cmbIdTipoCatalogo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoCatalogo
    AdministrarKeyPreview KeyCode

End Sub

Private Sub grdBienes_BeforeRowActivate(ByVal Row As UltraGrid.SSRow)
    ml_idRegistroSeleccionado = Row.Cells(1).Value
End Sub

Private Sub grdBienes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    ConfigurarGrilla False
End Sub
Sub ConfigurarGrilla(lCatalogoBase As Boolean)
    
    grdBienes.Bands(0).Columns("IdTipoBienInsumo").Hidden = True
    grdBienes.Bands(0).Columns("IdProducto").Hidden = True

    grdBienes.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdBienes.Bands(0).Columns("Codigo").Width = 1200

    grdBienes.Bands(0).Columns("Descripcion").Header.Caption = "Nombre"
    grdBienes.Bands(0).Columns("Descripcion").Width = IIf(lCatalogoBase, 10500, 8500)

    grdBienes.Bands(1).Columns("IdTipoBienInsumo").Hidden = True
    grdBienes.Bands(1).Columns("IdProducto").Hidden = True

    grdBienes.Bands(1).Columns("Codigo").Header.Caption = "Código"
    grdBienes.Bands(1).Columns("Codigo").Width = 1200

    grdBienes.Bands(1).Columns("Nombre").Header.Caption = "Nombre"
    grdBienes.Bands(1).Columns("Nombre").Width = 6100

    grdBienes.Bands(1).Columns("PrecioUnitario").Header.Caption = "Precio Unitario (S/.)"
    grdBienes.Bands(1).Columns("PrecioUnitario").Width = 2000
    grdBienes.Bands(1).Columns("PrecioUnitario").Hidden = lCatalogoBase

    grdBienes.Bands(1).Columns("Activo").Header.Caption = "Activo"
    grdBienes.Bands(1).Columns("Activo").Width = 2500
    grdBienes.Bands(1).Columns("Activo").Hidden = lCatalogoBase

End Sub




Public Function inicializar()
    
    Set mo_cmbIdTipoCatalogo.MiComboBox = cmbIdTipoCatalogo
    
End Function
Public Function SeleccionaTipoCatalogo()
    mo_cmbIdTipoCatalogo.BoundText = ml_IdTipoCatalogo
End Function



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
   grdBienes.Width = fraBusqueda.Width
   grdBienes.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub

Public Sub ConfigurarTiposCatalogos()
    
    mo_cmbIdTipoCatalogo.BoundColumn = "IdTipoFinanciamiento"
    mo_cmbIdTipoCatalogo.ListField = "DescripcionLarga"
    Set mo_cmbIdTipoCatalogo.RowSource = mo_AdminFacturacion.TiposFinanciamientoSeleccionarTodos()
    
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
