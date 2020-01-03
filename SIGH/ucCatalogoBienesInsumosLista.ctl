VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucCatalogoBienesInsumosL 
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11010
   LockControls    =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   11010
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
      Height          =   1290
      Left            =   45
      TabIndex        =   7
      Top             =   525
      Width           =   10845
      Begin VB.Frame FraPrecios 
         Caption         =   "% Precio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   9585
         TabIndex        =   14
         Top             =   255
         Width           =   1155
         Begin VB.CommandButton cmdAumentaPrecio 
            Caption         =   "..."
            Height          =   315
            Left            =   375
            TabIndex        =   16
            ToolTipText     =   "Aumenta precios a todos los ITEMS según % (toma como base PRECIO PARTICULAR)"
            Top             =   510
            Width           =   360
         End
         Begin VB.Label lblPorc 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "...."
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
            Left            =   405
            TabIndex        =   15
            Top             =   255
            Width           =   240
         End
      End
      Begin VB.CommandButton bntReporte 
         Enabled         =   0   'False
         Height          =   795
         Left            =   8640
         Picture         =   "ucCatalogoBienesInsumosLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   330
         Width           =   885
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1650
         MaxLength       =   20
         TabIndex        =   1
         Top             =   780
         Width           =   1215
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
         Left            =   1650
         TabIndex        =   0
         Top             =   360
         Width           =   5385
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   3750
         MaxLength       =   50
         TabIndex        =   2
         Top             =   780
         Width           =   3285
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   7170
         Picture         =   "ucCatalogoBienesInsumosLista.ctx":04D9
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   780
         Width           =   1305
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7170
         Picture         =   "ucCatalogoBienesInsumosLista.ctx":30B5
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label4 
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
         Height          =   285
         Left            =   3000
         TabIndex        =   11
         Top             =   840
         Width           =   735
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
         TabIndex        =   10
         Top             =   360
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
         Height          =   225
         Left            =   150
         TabIndex        =   8
         Top             =   810
         Width           =   1095
      End
   End
   Begin UltraGrid.SSUltraGrid grdBienesInsumos 
      Height          =   4710
      Left            =   60
      TabIndex        =   6
      Top             =   1920
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   8308
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
      Caption         =   "Lista de Bienes e Insumos"
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Linea en  ROJO= Precio de última Compra  >  Precio Venta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4920
      TabIndex        =   13
      Top             =   6720
      Width           =   4875
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Teclas de Ayuda:  <ENTER> = Modifica Precios "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   6720
      Width           =   3930
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00808080&
      Caption         =   "Catálogo de Bienes e Insumos"
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
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   10845
   End
End
Attribute VB_Name = "ucCatalogoBienesInsumosL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para buscar Medicamentos/Insumos
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
Dim mo_cmbIdClasificacionBienInsumo As New ListaDespleglable
Dim mo_cmbTipoCatalogo As New ListaDespleglable
Dim rsMedicamentos As New ADODB.Recordset
Dim ml_IdPlanSeleccionado As Long
Dim ml_idUsuario As Long
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim ml_lnHwnd As Long
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Property Let lnHWnd(lValue As Long)
    ml_lnHwnd = lValue
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdBienesInsumos.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdBienesInsumos.DataSource
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
    Dim oDOCatalogoBienesInsumos As New DOCatalogoBienesInsumos
    Dim rsProductos As New ADODB.Recordset
    
    '
    Set rsMedicamentos = mo_AdminComun.TiposFinanciamientoSegunFiltro("SeIngresPrecios=1")
    lblPorc.Caption = ""
    rsMedicamentos.Filter = "idTipoFinanciamiento=" & ml_IdTipoCatalogo
    If rsMedicamentos.RecordCount > 0 Then
       lblPorc.Caption = rsMedicamentos!porcPrecio
    End If
    rsMedicamentos.Close
    '
    oDOCatalogoBienesInsumos.Codigo = Trim(txtCodigo.Text)
    oDOCatalogoBienesInsumos.nombre = Trim(txtNombre.Text)
    Set rsMedicamentos = mo_AdminComun.CatalogoBienesInsumosFiltrarDEBB(oDOCatalogoBienesInsumos, ml_IdTipoCatalogo)
    Set grdBienesInsumos.DataSource = rsMedicamentos
    ConfigurarGrilla ml_IdTipoCatalogo = 0
    If mo_AdminComun.MensajeError <> "" Then
        MsgBox mo_AdminComun.MensajeError, vbInformation, "Filtro de Bienes Insumos"
    End If
    'mo_Apariencia.ConfigurarFilasBiColores grdBienesInsumos, sighentidades.GrillaConFilasBicolor
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
    If ml_IdTipoCatalogo <> 0 Then
       bntReporte.Enabled = True
    Else
       bntReporte.Enabled = False
    End If
    
'    Set grdServicios.DataSource = mo_AdminFacturacion.FacturacionSeleccionarCatalogo(mo_cmbTipoCatalogo.BoundText)
'    ConfigurarGrilla ml_IdTipoCatalogo = 0
'
'    If mo_AdminComun.MensajeError <> "" Then
'        MsgBox mo_AdminComun.MensajeError, vbInformation, "Filtro de Bienes Insumos"
'    End If
'    mo_Apariencia.ConfigurarFilasBiColores grdServicios, sighEntidades.GrillaConFilasBicolor

End Sub

Private Sub cmbTipoCatalogo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoCatalogo
    AdministrarKeyPreview KeyCode

End Sub

Private Sub cmdAumentaPrecio_Click()
    If MsgBox("Esta seguro que desea AUMENTAR EL " & lblPorc.Caption & " % DE PRECIOS A TODOS LOS ITEMS ?", vbQuestion + vbYesNo, "") = vbYes Then
       MousePointer = 11
       Dim oRsTmp1 As New Recordset
       Dim oDOCatalogoBienesInsumos As New DOCatalogoBienesInsumos
       Dim oConexion As New Connection
       Dim lnPrecioNuevo As Double
       oConexion.CommandTimeout = 900
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighentidades.CadenaConexion
       Set oRsTmp1 = mo_AdminComun.CatalogoBienesInsumosFiltrarDEBB(oDOCatalogoBienesInsumos, 1)
       If oRsTmp1.RecordCount > 0 Then
          oRsTmp1.MoveFirst
          Do While Not oRsTmp1.EOF
             lnPrecioNuevo = oRsTmp1!PrecioUnitario + Round(oRsTmp1!PrecioUnitario * CCur(lblPorc.Caption) / 100, 2)
             mo_AdminComun.FactCatalogoBienesInsumosHospActualizaPrecio oConexion, oRsTmp1!idProducto, _
                                                                        ml_IdTipoCatalogo, lnPrecioNuevo
             oRsTmp1.MoveNext
          Loop
       End If
       oRsTmp1.Close
       oConexion.Close
       Set oConexion = Nothing
       Set oRsTmp1 = Nothing
       Set oDOCatalogoBienesInsumos = Nothing
       btnBuscar_Click
       MousePointer = 1
       MsgBox "Se terminó de cambiar los precios sin problemas", vbInformation, ""
       
    End If
End Sub

Private Sub grdBienesInsumos_BeforeRowActivate(ByVal Row As UltraGrid.SSRow)
    ml_idRegistroSeleccionado = Row.Cells("idProducto").Value
    If ml_IdTipoCatalogo <> 0 Then
       ml_IdPlanSeleccionado = Row.Cells("IdPlanCatalogo").Value
    End If
End Sub

Private Sub grdBienesInsumos_DblClick()
    Dim mo_ReglasDeSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Select Case ml_IdTipoCatalogo
    Case 0   'base
    Case 1   'contado
        If mo_ReglasDeSeguridad.SoloTieneOpcionCONSULTA(ml_idUsuario, 803) = False Then
            Dim oDetalle As New SIGHCatalogos.clCatalogoBienesFarmacia
            oDetalle.Opcion = sghModificar
            oDetalle.idProducto = ml_idRegistroSeleccionado
            oDetalle.IdPlanCatalogo = ml_IdPlanSeleccionado
            oDetalle.MostrarFormulario
            Set oDetalle = Nothing
            btnBuscar_Click
        End If
    Case Else
'        If mo_ReglasDeSeguridad.SoloTieneOpcionCONSULTA(ml_idUsuario, 803) = False Then
'            Dim oDetalleFF As New SIGHCatalogos.clCatalogoBienesFinanciam
'            oDetalleFF.Opcion = sghModificar
'            oDetalleFF.idProducto = ml_idRegistroSeleccionado
'            oDetalleFF.IdPlanCatalogo = ml_IdPlanSeleccionado
'            oDetalleFF.MostrarFormulario
'            Set oDetalleFF = Nothing
'            btnBuscar_Click
'        End If
    End Select
    Set mo_ReglasDeSeguridad = Nothing
End Sub

'Private Sub grdBienesInsumos_AfterRowActivate()
'Dim rsRecordset As ADODB.Recordset
'
'    Set rsRecordset = grdBienesInsumos.DataSource
'    On Error Resume Next
'    ml_IdRegistroSeleccionado = rsRecordset("IdProducto")
'
'End Sub
'
'Private Sub grdBienesInsumos_Click()
'Dim rsRecordset As ADODB.Recordset
'
'    Set rsRecordset = grdBienesInsumos.DataSource
'    On Error Resume Next
'    ml_IdRegistroSeleccionado = rsRecordset("IdProducto")
'
'End Sub


Private Sub grdBienesInsumos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    ConfigurarGrilla True
 
'    grdBienesInsumos.Bands(0).Columns("IdProducto").Hidden = True
'
'    grdBienesInsumos.Bands(0).Columns("Codigo").Header.Caption = "Código"
'    grdBienesInsumos.Bands(0).Columns("Codigo").Width = 1200
'
'    grdBienesInsumos.Bands(0).Columns("Nombre").Header.Caption = "Nombre"
'    grdBienesInsumos.Bands(0).Columns("Nombre").Width = 4000
'
'    grdBienesInsumos.Bands(0).Columns("NombreComercial").Header.Caption = "Nombre Comercial"
'    grdBienesInsumos.Bands(0).Columns("NombreComercial").Width = 2500
'
'
'    grdBienesInsumos.Bands(0).Columns("DescTiposDeBienesEInsumos").Header.Caption = "Tipo"
'    grdBienesInsumos.Bands(0).Columns("DescTiposDeBienesEInsumos").Width = 2500
End Sub

Sub ConfigurarGrilla(lCatalogoBase As Boolean)
    Dim lnFilaProductos As Integer
    If ml_IdTipoCatalogo = 0 Then
        lnFilaProductos = 1
        grdBienesInsumos.Bands(0).Columns("IdSubGrupoFarmacologico").Hidden = True
        grdBienesInsumos.Bands(0).Columns("IdProducto").Hidden = True
    
        grdBienesInsumos.Bands(0).Columns("Codigo").Header.Caption = "Código"
        grdBienesInsumos.Bands(0).Columns("Codigo").Width = 1200
        grdBienesInsumos.Bands(0).Columns("codigo").Activation = ssActivationActivateNoEdit
    
        grdBienesInsumos.Bands(0).Columns("Descripcion").Header.Caption = "Nombre"
        grdBienesInsumos.Bands(0).Columns("Descripcion").Width = IIf(lCatalogoBase, 10500, 8500)
        grdBienesInsumos.Bands(0).Columns("descripcion").Activation = ssActivationActivateNoEdit
    Else
        lnFilaProductos = 0
        If ml_IdTipoCatalogo <> 1 Then
           grdBienesInsumos.Bands(lnFilaProductos).Columns("PrecioCompra").Hidden = True
           grdBienesInsumos.Bands(lnFilaProductos).Columns("PrecioDistribucion").Hidden = True
           grdBienesInsumos.Bands(lnFilaProductos).Columns("PrecioDonacion").Hidden = True
           grdBienesInsumos.Bands(lnFilaProductos).Columns("PrecioUltcompra").Hidden = True
        End If
        grdBienesInsumos.Bands(lnFilaProductos).Columns("PrecioCompra").Activation = ssActivationActivateNoEdit
        grdBienesInsumos.Bands(lnFilaProductos).Columns("PrecioDistribucion").Activation = ssActivationActivateNoEdit
        grdBienesInsumos.Bands(lnFilaProductos).Columns("PrecioDonacion").Activation = ssActivationActivateNoEdit
        grdBienesInsumos.Bands(lnFilaProductos).Columns("PrecioUltCompra").Activation = ssActivationActivateNoEdit
        
        grdBienesInsumos.Bands(lnFilaProductos).Columns("IdPlanCatalogo").Hidden = True
        grdBienesInsumos.Bands(lnFilaProductos).Columns("PrecioUnitario").Header.Caption = "Precio Unitario (S/.)"
        grdBienesInsumos.Bands(lnFilaProductos).Columns("PrecioUnitario").Width = 2000
        grdBienesInsumos.Bands(lnFilaProductos).Columns("PrecioUnitario").Hidden = lCatalogoBase
        grdBienesInsumos.Bands(lnFilaProductos).Columns("PrecioUnitario").Activation = ssActivationActivateNoEdit
    
        grdBienesInsumos.Bands(lnFilaProductos).Columns("Activo").Header.Caption = "Activo"
        grdBienesInsumos.Bands(lnFilaProductos).Columns("Activo").Width = 2500
        grdBienesInsumos.Bands(lnFilaProductos).Columns("Activo").Hidden = lCatalogoBase
        grdBienesInsumos.Bands(lnFilaProductos).Columns("Activo").Activation = ssActivationActivateNoEdit
    End If
    grdBienesInsumos.Bands(lnFilaProductos).Columns("IdSubGrupoFarmacologico").Hidden = True
    grdBienesInsumos.Bands(lnFilaProductos).Columns("IdProducto").Hidden = True
    grdBienesInsumos.Bands(lnFilaProductos).Columns("IdProducto").Activation = ssActivationActivateNoEdit

    grdBienesInsumos.Bands(lnFilaProductos).Columns("Codigo").Header.Caption = "Código"
    grdBienesInsumos.Bands(lnFilaProductos).Columns("Codigo").Width = 1200
    grdBienesInsumos.Bands(lnFilaProductos).Columns("codigo").Activation = ssActivationActivateNoEdit

    grdBienesInsumos.Bands(lnFilaProductos).Columns("Nombre").Header.Caption = "Nombre"
    grdBienesInsumos.Bands(lnFilaProductos).Columns("Nombre").Width = 6100
    grdBienesInsumos.Bands(lnFilaProductos).Columns("nombre").Activation = ssActivationActivateNoEdit



End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdBienesInsumos, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdBienesInsumos, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub

Function inicializar()
    SkinConfigura
    Set mo_cmbTipoCatalogo.MiComboBox = UserControl.cmbTipoCatalogo
    'Set mo_cmbIdClasificacionBienInsumo.MiComboBox = UserControl.cmbTipoBienInsumo
    CargarComboBoxes
End Function



Private Sub grdBienesInsumos_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
    If ml_IdTipoCatalogo = 1 Then
        If Val(Row.Cells("PrecioUltCompra").GetText()) > Val(Row.Cells("PrecioUnitario").GetText()) Then
            Row.Appearance.ForeColor = vbRed
        End If
    End If

End Sub

Private Sub grdBienesInsumos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
     If KeyAscii = 13 Then
        grdBienesInsumos_DblClick
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
   
   grdBienesInsumos.Width = fraBusqueda.Width
   grdBienesInsumos.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + Label3.Height + 150)
   Label3.Top = UserControl.Height - UserControl.Label3.Height - 50
   Label5.Top = Label3.Top '+ Label3.Width + 50
End Sub
Private Sub CargarComboBoxes()

    mo_cmbTipoCatalogo.BoundColumn = "IdTipoFinanciamiento"
    mo_cmbTipoCatalogo.ListField = "Descripcion"
    Set mo_cmbTipoCatalogo.RowSource = mo_AdminFacturacion.TiposFinanciamientoSeleccionarTodos()
    '
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Dim oRsAlmacen As Recordset
    Set oRsAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    If oRsAlmacen.RecordCount > 0 Then
       Set mo_cmbTipoCatalogo.RowSource = mo_ReglasFarmacia.TipoFinanciamientosParaCatalogoBienes
    End If
    Set oBuscaDondeLabora = Nothing


End Sub

'***************daniel barrantes**************
'debb-23/11/2016
Private Sub bntReporte_Click()
    
    Dim iFila As Long
    Dim lnTotal As Long
    Dim rsreporte As New Recordset
    Dim rsReporte1 As New Recordset
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim mo_ReporteUtil As New sighentidades.ReporteUtil
    Dim oRsTarifas As New Recordset
    Dim lcNombre As String
    Dim lnCant As Long
    Dim lbEsOpenOffice As Boolean
    Dim lcSql As String
    Dim lnCol As Integer
    With oRsTarifas
          .Fields.Append "id", adInteger, 4, adFldIsNullable
          .Fields.Append "Descripcion", adVarChar, 200, adFldIsNullable
          .Fields.Append "Columna", adInteger
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set rsreporte = mo_AdminComun.TiposFinanciamientoSegunFiltro("seIngresPrecios=1 and idTipoFinanciamiento>0")
    If rsreporte.RecordCount > 0 Then
       lnCol = 4
       rsreporte.MoveFirst
       Do While Not rsreporte.EOF
          oRsTarifas.AddNew
          oRsTarifas.Fields!Id = rsreporte.Fields!idTipoFinanciamiento
          oRsTarifas.Fields!Descripcion = rsreporte.Fields!Descripcion
          oRsTarifas.Fields!Columna = lnCol
          oRsTarifas.Update
          lnCol = lnCol + 1
          rsreporte.MoveNext
       Loop
    End If
    rsreporte.Close
    
    

    lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
    
    If lbEsOpenOffice = True Then
        Dim ServiceManager As Object
        Dim Desktop As Object
        Dim Document As Object
        Dim Feuille As Object
        Dim Plage As Object
        Dim args()
        Dim Chemin As String
        Dim Fichier As String
        Dim lcArchivoExcel As String
        Dim PrintArea(0)
        Dim Style As Object
        Dim Border As Object
        'encabezado
        Dim PageStyles As Object
        Dim Sheet As Object
        Dim StyleFamilies As Object
        Dim DefPage As Object
        Dim Htext As Object
        Dim Hcontent As Object
        Dim ret As Long
        Dim lnHWnd As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
    Set rsreporte = grdBienesInsumos.DataSource
    If rsreporte.RecordCount > 0 Then
        MousePointer = 11
        
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\HerrListaServiciosMedicamentos.ods"
'        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'        Chemin = "file:///" & App.Path & "\Plantillas\"
'        Chemin = Replace(Chemin, "\", "/")
'        Fichier = Chemin & "/OpenOffice.ods"
        Fichier = Format(Time, "hhmmss") & ".ods"
        FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
        lcArchivoExcel = Fichier
        Chemin = "file:///" & App.Path & "\Plantillas\"
        Chemin = Replace(Chemin, "\", "/")
        Fichier = Chemin & "/" & lcArchivoExcel
        '
        Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
        Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
        Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
        Set Feuille = Document.getSheets().getByIndex(0)
        'Encabezado de Pagina
        mo_CabeceraReportes.CabeceraReportes Document, True
        ' Pone la ventana en primer plano, pasándole el Hwnd
        ret = SetForegroundWindow(lnHWnd)
    Else
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HerrListaServiciosMedicamentos.xls")
        oWorkBookPlantilla.Worksheets("listaServiciosMedicamentos").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(1, 1).setFormula("RELACION DE MEDICAMENTOS/INSUMOS")
    Else
        oWorkSheet.Cells(2, 2).Value = "RELACION DE MEDICAMENTOS/INSUMOS"
    End If
    '
    oRsTarifas.MoveFirst
    Do While Not oRsTarifas.EOF
      If lbEsOpenOffice = True Then
         Call Feuille.getcellbyposition(oRsTarifas!Columna - 1, 4).setFormula(oRsTarifas!Descripcion)
      Else
         oWorkSheet.Cells(5, oRsTarifas!Columna).Value = oRsTarifas!Descripcion
      End If
      oRsTarifas.MoveNext
    Loop
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(lnCol - 1, 4).setFormula("PrecioCompra")
        Call Feuille.getcellbyposition(lnCol, 4).setFormula("PrecioDistribucion")
        Call Feuille.getcellbyposition(lnCol + 1, 4).setFormula("Forma Farmaceutica")
        Call Feuille.getcellbyposition(lnCol + 2, 4).setFormula("Tipo Producto Sismed")
    Else
        oWorkSheet.Cells(5, lnCol).Value = "PrecioCompra"
        oWorkSheet.Cells(5, lnCol + 1).Value = "PrecioDistribucion"
        oWorkSheet.Cells(5, lnCol + 2).Value = "Forma Farmaceutica"
        oWorkSheet.Cells(5, lnCol + 3).Value = "Tipo Producto Sismed"
    End If
    '
    
        iFila = 6
        lnTotal = 0
        rsreporte.MoveFirst
        Do While Not rsreporte.EOF
            lcNombre = rsreporte.Fields!nombre
            lnCant = 1
            Do While Not rsreporte.EOF And lcNombre = rsreporte.Fields!nombre
                If lnCant <= 1 Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(0, iFila - 1).setFormula(Trim(Str(rsreporte.Fields!TipoProducto)))
                        Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsreporte.Fields!Codigo)
                        Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsreporte.Fields!nombre)
                        'Call Feuille.getcellbyposition(5, iFila - 1).setFormula(Format(rsReporte.Fields!PrecioUnitario, "####,###.##"))
                    Else
                        oWorkSheet.Cells(iFila, 1).Value = "'" & rsreporte.Fields!TipoProducto
                        oWorkSheet.Cells(iFila, 2).Value = rsreporte.Fields!Codigo
                        oWorkSheet.Cells(iFila, 3).Value = rsreporte.Fields!nombre
                        'oWorkSheet.Cells(iFila, 6).Value = Format(rsReporte.Fields!PrecioUnitario, "####,###.##")
                    End If
                    Set rsReporte1 = mo_AdminComun.FactCatalogoBienesInsumosHospXfiltro("idProducto=" & Trim(Str(rsreporte.Fields!idProducto)))
                    If rsReporte1.RecordCount > 0 Then
                       rsReporte1.MoveFirst
                       Do While Not rsReporte1.EOF
                            oRsTarifas.MoveFirst
                            oRsTarifas.Find "id=" & rsReporte1!idTipoFinanciamiento
                            If Not oRsTarifas.EOF Then
                               If lbEsOpenOffice = True Then
                                   lcSql = Trim(Str(rsReporte1.Fields!PrecioUnitario))
                                   Call Feuille.getcellbyposition(oRsTarifas!Columna - 1, iFila - 1).setFormula(lcSql)
                               Else
                                   oWorkSheet.Cells(iFila, oRsTarifas!Columna).Value = rsReporte1!PrecioUnitario
                               End If
                            End If
                            rsReporte1.MoveNext
                       Loop
                    End If
                    rsReporte1.Close
                    Set rsReporte1 = mo_AdminComun.CatalogoBienesInsumosSeleccionarXid(rsreporte!idProducto)
                    If rsReporte1.RecordCount > 0 Then
                       If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(lnCol - 1, iFila - 1).sefformula(IIf(IsNull(rsReporte1!PrecioCompra), "0", rsReporte1!PrecioCompra))
                            Call Feuille.getcellbyposition(lnCol, iFila - 1).sefformula(IIf(IsNull(rsReporte1!PrecioDistribucion), "0", rsReporte1!PrecioDistribucion))
                            Call Feuille.getcellbyposition(lnCol + 1, iFila - 1).sefformula(IIf(IsNull(rsReporte1!formaFarmaceutica), "", rsReporte1!formaFarmaceutica))
                            Call Feuille.getcellbyposition(lnCol + 2, iFila - 1).sefformula(IIf(IsNull(rsReporte1!tipoProductoSismed), "", rsReporte1!tipoProductoSismed))
                       Else
                            oWorkSheet.Cells(iFila, lnCol).Value = IIf(IsNull(rsReporte1!PrecioCompra), "0", rsReporte1!PrecioCompra)
                            oWorkSheet.Cells(iFila, lnCol + 1).Value = IIf(IsNull(rsReporte1!PrecioDistribucion), "0", rsReporte1!PrecioDistribucion)
                            oWorkSheet.Cells(iFila, lnCol + 2).Value = IIf(IsNull(rsReporte1!formaFarmaceutica), "", rsReporte1!formaFarmaceutica)
                            oWorkSheet.Cells(iFila, lnCol + 3).Value = IIf(IsNull(rsReporte1!tipoProductoSismed), "", rsReporte1!tipoProductoSismed)
                       End If
                    End If
                    rsReporte1.Close
                    
                    
                    iFila = iFila + 1
                    lnTotal = lnTotal + 1
                End If
                lnCant = lnCant + 1
                rsreporte.MoveNext
                If rsreporte.EOF Then
                   Exit Do
                End If
            Loop
        Loop
        iFila = iFila + 1
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":G" & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nro: ")
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula(Format(lnTotal, "####,###"))
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 7
            oWorkSheet.Cells(iFila, 2).Value = "Nro: "
            oWorkSheet.Cells(iFila, 3).Value = Format(lnTotal, "####,###")
        End If
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 6
            PrintArea(0).EndRow = iFila
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            If oWorkSheet.PageSetup.PrintArea <> "" Then
               oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
            End If
            oExcel.Visible = True
            oWorkSheet.PrintPreview
            'oWorkSheet.PrintOut
        End If
        MousePointer = 1
    End If
    'rsReporte.Close
    If lbEsOpenOffice = True Then
        'Liberar Memoria
        Set Plage = Nothing
        Set Feuille = Nothing
        Set Document = Nothing
        Set Desktop = Nothing
        Set ServiceManager = Nothing
        Set Style = Nothing
        Set Border = Nothing
        'encabezado de pagina
        Set PageStyles = Nothing
        Set Sheet = Nothing
        Set StyleFamilies = Nothing
        Set DefPage = Nothing
        Set Htext = Nothing
        Set Hcontent = Nothing
    Else
        oExcel.Visible = True
        oWorkSheet.PrintPreview
        'liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    
    
End Sub






