VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucCatalogoServiciosLista 
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
      Height          =   1185
      Left            =   60
      TabIndex        =   7
      Top             =   540
      Width           =   9900
      Begin VB.CommandButton bntReporte 
         Height          =   765
         Left            =   8700
         Picture         =   "ucCatalogoServiciosLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   825
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   1
         Top             =   690
         Width           =   1245
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7230
         Picture         =   "ucCatalogoServiciosLista.ctx":04D9
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   7230
         Picture         =   "ucCatalogoServiciosLista.ctx":3122
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   690
         Width           =   1305
      End
      Begin VB.ComboBox cmbTipoCatalogo 
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
         Width           =   5415
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   3810
         MaxLength       =   50
         TabIndex        =   2
         Top             =   690
         Width           =   3255
      End
      Begin VB.Label Label3 
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
         Left            =   3060
         TabIndex        =   11
         Top             =   750
         Width           =   675
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
         Left            =   150
         TabIndex        =   10
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Height          =   315
         Left            =   150
         TabIndex        =   8
         Top             =   750
         Width           =   1545
      End
   End
   Begin UltraGrid.SSUltraGrid grdServicios 
      Height          =   6420
      Left            =   60
      TabIndex        =   6
      Top             =   1770
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   11324
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
      Caption         =   "Lista de Servicios"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00808080&
      Caption         =   "Catálogo de Servicios"
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
Attribute VB_Name = "ucCatalogoServiciosLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para listar procedimientos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim ml_idRegistroSeleccionado As Long
Dim ml_IdTipoCatalogo As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_cmbCategoria As New ListaDespleglable
Dim mo_cmbTipoCatalogo As New ListaDespleglable
Dim rsMedicamentos As New ADODB.Recordset

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdServicios.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdServicios.DataSource
End Property
Property Let IdTipoCatalogo(lValue As Long)
    ml_IdTipoCatalogo = lValue
End Property
Property Get IdTipoCatalogo() As Long
    IdTipoCatalogo = ml_IdTipoCatalogo
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
    Dim oDOCatalogoServicios As New DOCatalogoServicio
    
    oDOCatalogoServicios.Codigo = Trim(txtCodigo.Text)
    oDOCatalogoServicios.nombre = Trim(txtNombre.Text)
    'Set grdServicios.DataSource = mo_AdminComun.CatalogoServiciosFiltrar(oDOCatalogoServicios, ml_IdTipoCatalogo)
    Set rsMedicamentos = mo_AdminComun.CatalogoServiciosFiltrarDEBB(oDOCatalogoServicios, ml_IdTipoCatalogo)
    Set grdServicios.DataSource = rsMedicamentos
    
    'ConfigurarGrilla ml_IdTipoCatalogo = 0
    ConfigurarGrilla True
    
    If mo_AdminComun.MensajeError <> "" Then
        MsgBox mo_AdminComun.MensajeError, vbInformation, "Búsqueda del catálogo de servicios"
    End If
   ' mo_Apariencia.ConfigurarFilasBiColores grdServicios, sighentidades.GrillaConFilasBicolor
    grdServicios.Bands(0).Expand
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
    UserControl.txtCodigo = ""
    UserControl.txtNombre = ""
End Sub

Private Sub cmbTipoCatalogo_Click()

    ml_IdTipoCatalogo = Val(mo_cmbTipoCatalogo.BoundText)
    
    RealizarBusqueda

End Sub



Private Sub cmbTipoCatalogo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoCatalogo
    AdministrarKeyPreview KeyCode

End Sub


Private Sub grdServicios_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
    If grdServicios.ActiveCell.Column.Key = "PrecioUnitario" Or grdServicios.ActiveCell.Column.Key = "Activo" Then
        Dim oRow As SSRow
        Dim lnDctos As Double
        Dim lnCant As Long
        Dim oRsTmp As New ADODB.Recordset
        Dim lnActivo As Integer
        Dim oConexion As New ADODB.Connection
        Dim lcSql As String
        Dim lnPrecio As Double, lnIdProducto As Long
        
        Set oRow = grdServicios.ActiveCell.Row
        lnPrecio = oRow.Cells("PrecioUnitario").Value
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        lnActivo = IIf(oRow.Cells("activo").Value = True, 1, 0)
        lnIdProducto = oRow.Cells("idProducto").Value
        If lnActivo = 0 Then
            mo_AdminFacturacion.CatalogoServiciosHospEliminarXtipoFinanciamientoIdProducto ml_IdTipoCatalogo, lnIdProducto, oConexion
            grdServicios.Refresh
        Else
            Set oRsTmp = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(lnIdProducto, ml_IdTipoCatalogo)
            If oRsTmp.RecordCount > 0 Then
               oRsTmp.Fields!PrecioUnitario = lnPrecio
               oRsTmp.Fields!Activo = lnActivo
               oRsTmp.Update
            End If
            oRsTmp.Close
        End If
        Set oRsTmp = Nothing
        oConexion.Close
        Set oConexion = Nothing
    End If
End Sub

Private Sub grdServicios_BeforeRowActivate(ByVal Row As UltraGrid.SSRow)
    ml_idRegistroSeleccionado = Row.Cells(1).Value
End Sub

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    ConfigurarGrilla True
    
End Sub

Sub ConfigurarGrilla(lCatalogoBase As Boolean)
    Dim lnFilaProductos As Integer
    If ml_IdTipoCatalogo = 0 Then
        lnFilaProductos = 1
        grdServicios.Bands(0).Columns("IdServicioSubGrupo").Hidden = True
        grdServicios.Bands(0).Columns("IdProducto").Hidden = True
        grdServicios.Bands(0).Columns("idProducto").Activation = ssActivationActivateNoEdit
        
        grdServicios.Bands(0).Columns("Codigo").Header.Caption = "Código"
        grdServicios.Bands(0).Columns("Codigo").Width = 1200
        grdServicios.Bands(0).Columns("codigo").Activation = ssActivationActivateNoEdit
    
        grdServicios.Bands(0).Columns("Descripcion").Header.Caption = "Nombre"
        grdServicios.Bands(0).Columns("Descripcion").Width = IIf(lCatalogoBase, 10500, 8500)
        grdServicios.Bands(0).Columns("descripcion").Activation = ssActivationActivateNoEdit
        grdServicios.Bands(lnFilaProductos).Columns("PrecioUnitario").Hidden = lCatalogoBase
    Else
        lnFilaProductos = 0
         '
        
    End If

    grdServicios.Bands(lnFilaProductos).Columns("IdServicioSubGrupo").Hidden = True
    grdServicios.Bands(lnFilaProductos).Columns("IdProducto").Hidden = True
    grdServicios.Bands(lnFilaProductos).Columns("idProducto").Activation = ssActivationActivateNoEdit

    grdServicios.Bands(lnFilaProductos).Columns("Codigo").Header.Caption = "Código"
    grdServicios.Bands(lnFilaProductos).Columns("Codigo").Width = 1200
    grdServicios.Bands(lnFilaProductos).Columns("codigo").Activation = ssActivationActivateNoEdit

    grdServicios.Bands(lnFilaProductos).Columns("Nombre").Header.Caption = "Nombre"
    grdServicios.Bands(lnFilaProductos).Columns("Nombre").Width = 6100
    grdServicios.Bands(lnFilaProductos).Columns("nombre").Activation = ssActivationActivateNoEdit

    grdServicios.Bands(lnFilaProductos).Columns("PrecioUnitario").Header.Caption = "Precio Unitario (S/.)"
    grdServicios.Bands(lnFilaProductos).Columns("PrecioUnitario").Width = 2000
'    grdServicios.Bands(lnFilaProductos).Columns("PrecioUnitario").Hidden = lCatalogoBase

    grdServicios.Bands(lnFilaProductos).Columns("Activo").Header.Caption = "Activo"
    grdServicios.Bands(lnFilaProductos).Columns("Activo").Width = 2500
    grdServicios.Bands(lnFilaProductos).Columns("Activo").Hidden = lCatalogoBase
    
    grdServicios.Bands(0).CollapseAll
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdServicios, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdServicios, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub


Public Function inicializar()
    SkinConfigura
    Set mo_cmbTipoCatalogo.MiComboBox = UserControl.cmbTipoCatalogo
    CargarComboBoxes
    bntReporte.Enabled = True
End Function



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode
End Sub



Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNombre
    AdministrarKeyPreview KeyCode
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
        btnLimpiar_Click
     Case vbKeyF8
    End Select
       
End Sub
Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   
   grdServicios.Width = fraBusqueda.Width
   grdServicios.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub
Private Sub CargarComboBoxes()

    mo_cmbTipoCatalogo.BoundColumn = "IdTipoFinanciamiento"
    mo_cmbTipoCatalogo.ListField = "Descripcion"
    Set mo_cmbTipoCatalogo.RowSource = mo_AdminFacturacion.TiposFinanciamientoSeleccionarTodos()
    
End Sub


'***************daniel barrantes**************
'***************
Private Sub bntReporte_Click()
    'Dim oReportes As New RpCatServicios
    Dim oReportes As New SIGHReportes.clCatalogoServicios
    oReportes.EjecutaFormulario
    Set oReportes = Nothing
End Sub



