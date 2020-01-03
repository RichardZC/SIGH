VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucConfiguraResLab 
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10035
   LockControls    =   -1  'True
   ScaleHeight     =   8340
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
         Left            =   3810
         MaxLength       =   50
         TabIndex        =   5
         Top             =   330
         Width           =   3255
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   7230
         Picture         =   "ucConfiguraResLab.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   690
         Width           =   1305
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7230
         Picture         =   "ucConfiguraResLab.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1305
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   2
         Top             =   330
         Width           =   1245
      End
      Begin VB.CommandButton bntReporte 
         Height          =   765
         Left            =   8700
         Picture         =   "ucConfiguraResLab.ctx":5825
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   825
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
         TabIndex        =   7
         Top             =   330
         Width           =   1545
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
         TabIndex        =   6
         Top             =   330
         Width           =   675
      End
   End
   Begin UltraGrid.SSUltraGrid grdServicios 
      Height          =   6420
      Left            =   60
      TabIndex        =   8
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
      Caption         =   "Configuración de Resultados de Laboratorio"
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
Attribute VB_Name = "ucConfiguraResLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para listar resultados para laboratorio
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminComun As New SIGHNegocios.ReglasConfiguarcionReslab

Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim ml_idRegistroSeleccionado As Long
Dim ml_IdTipoCatalogo As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim rsServicios As New ADODB.Recordset

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
        Set rsServicios = mo_AdminComun.FiltrarCatalogoCC(oDOCatalogoServicios)
        Set grdServicios.DataSource = rsServicios
        ConfigurarGrilla ml_IdTipoCatalogo = 0
        If mo_AdminComun.MensajeError <> "" Then
            MsgBox mo_AdminComun.MensajeError, vbInformation, "Busqueda del catálogo de servicios"
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

Private Sub grdServicios_BeforeRowActivate(ByVal Row As UltraGrid.SSRow)
        ml_idRegistroSeleccionado = Row.Cells(1).Value
End Sub

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    ConfigurarGrilla False
End Sub

Sub ConfigurarGrilla(lCatalogoBase As Boolean)
    Dim lnFilaProductos As Integer
    If ml_IdTipoCatalogo = 0 Then
        lnFilaProductos = 1
        grdServicios.Bands(0).Columns("IdServicioSubGrupo").Hidden = True
        grdServicios.Bands(0).Columns("IdProducto").Hidden = True
        grdServicios.Bands(0).Columns("Codigo").Header.Caption = "Código"
        grdServicios.Bands(0).Columns("Codigo").Width = 1100
        grdServicios.Bands(0).Columns("codigo").Activation = ssActivationActivateNoEdit
        grdServicios.Bands(0).Columns("Descripcion").Header.Caption = "Nombre"
        grdServicios.Bands(0).Columns("Descripcion").Width = 7500
        grdServicios.Bands(0).Columns("descripcion").Activation = ssActivationActivateNoEdit
    Else
        lnFilaProductos = 0
    End If
    grdServicios.Bands(lnFilaProductos).Columns("IdServicioSubGrupo").Hidden = True
    grdServicios.Bands(lnFilaProductos).Columns("IdProducto").Hidden = True
    grdServicios.Bands(0).Columns("codigo").Activation = ssActivationActivateNoEdit
    grdServicios.Bands(lnFilaProductos).Columns("Nombre").Header.Caption = "Nombre"
    grdServicios.Bands(lnFilaProductos).Columns("Nombre").Width = 7500
    grdServicios.Bands(lnFilaProductos).Columns("nombre").Activation = ssActivationActivateNoEdit
    grdServicios.Bands(lnFilaProductos).Columns("NombreMInsa").Header.Caption = "Nombre Minsa"
    grdServicios.Bands(lnFilaProductos).Columns("NombreMinsa").Width = 7000
    grdServicios.Bands(lnFilaProductos).Columns("nombreMinsa").Activation = ssActivationActivateNoEdit
    grdServicios.Bands(lnFilaProductos).Columns("Descripcion").Hidden = True
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
Public Sub Inicializar()
    SkinConfigura
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
   
   grdServicios.Width = fraBusqueda.Width
   grdServicios.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub


Private Sub bntReporte_Click()
'    'Dim oReportes As New RpCatServicios
'    Dim oReportes As New SIGHReportes.clCatalogoServicios
'    oReportes.EjecutaFormulario
'    Set oReportes = Nothing
End Sub




