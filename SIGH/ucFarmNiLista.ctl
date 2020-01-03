VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucFarmNiLista 
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11175
   ScaleHeight     =   6555
   ScaleWidth      =   11175
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
      Height          =   915
      Left            =   90
      TabIndex        =   7
      Top             =   555
      Width           =   10995
      Begin VB.TextBox txtNCuenta 
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
         Left            =   8010
         MaxLength       =   9
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   9570
         Picture         =   "ucFarmNiLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9570
         Picture         =   "ucFarmNiLista.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   510
         Width           =   1305
      End
      Begin MSMask.MaskEdBox txtFinicio 
         Height          =   315
         Left            =   5220
         TabIndex        =   1
         Top             =   480
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFfinal 
         Height          =   315
         Left            =   6630
         TabIndex        =   2
         Top             =   480
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo cmbAlmacen 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Inicial      Fecha Final           N° Cuenta"
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
         Left            =   5250
         TabIndex        =   8
         Top             =   225
         Width           =   4380
      End
   End
   Begin UltraGrid.SSUltraGrid grdLista 
      Height          =   4950
      Left            =   90
      TabIndex        =   6
      Top             =   1500
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   8731
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
      BackColor       =   &H00373842&
      Caption         =   "Nota de Ingreso"
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
      Top             =   45
      Width           =   11010
   End
End
Attribute VB_Name = "ucFarmNiLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para buscar Nota de Ingreso
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoBusqueda As sghTipoBusquedaPrestamoHistoria
Dim mo_Teclado As New sighentidades.Teclado
Dim oRsAlmacenes As New ADODB.Recordset
Dim oRsBusqueda As New ADODB.Recordset
Dim ml_idUsuario As Long
Dim lbNIsoloParaFarmacia As Boolean

Property Let NIsoloParaFarmacia(lValue As Long)
   lbNIsoloParaFarmacia = lValue
End Property


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
    cmbAlmacen.Enabled = True

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
    If CDate(UserControl.txtFinicio.Text) > CDate(UserControl.txtFfinal.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
        If cmbAlmacen.Text = "" Then
           MsgBox "Por favor elija el Almacén", vbInformation, "Busqueda"
           Exit Sub
        End If
        
        'Set oRsBusqueda = mo_ReglasFarmacia.FarmDevuelveMovimientos(Val(cmbAlmacen.BoundText), "E", CDate(Format(txtFinicio.Text & " 00:00:00", sighEntidades.DevuelveFechaSoloFormato_DMY_HMS)), CDate(Format(txtFfinal.Text & " 23:59:59", sighEntidades.DevuelveFechaSoloFormato_DMY_HMS)))
        Set oRsBusqueda = mo_ReglasFarmacia.DevuelveMovimientosDeNotaIngresos(Val(cmbAlmacen.BoundText), Format(txtFinicio.Text & " 00:00:01", sighentidades.DevuelveFechaSoloFormato_DMY_HMS), Format(txtFfinal.Text & " 23:59:59", sighentidades.DevuelveFechaSoloFormato_DMY_HMS), Val(txtNcuenta.Text))
        'oRsBusqueda.Filter = "IdTipoConcepto<>19"
        Set grdLista.DataSource = oRsBusqueda
        'mo_Apariencia.ConfigurarFilasBiColores grdLista, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.cmbAlmacen.Text = ""
        UserControl.txtFinicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual
        UserControl.txtFfinal.Text = Date
        UserControl.txtNcuenta.Text = ""
End Sub



Private Sub cmbAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacen
    AdministrarKeyPreview KeyCode
End Sub

Private Sub grdLista_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdLista.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = Val(rsRecordset("MovNumero"))
    
End Sub

Private Sub grdLista_Click()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdLista.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = Val(rsRecordset("MovNumero"))
    
End Sub


Private Sub grdLista_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    
    grdLista.Bands(0).Columns("FDestino").Hidden = True
    grdLista.Bands(0).Columns("FOrigen").Hidden = True
    grdLista.Bands(0).Columns("idAlmacenDestino").Hidden = True
    grdLista.Bands(0).Columns("idAlmacenOrigen").Hidden = True
    grdLista.Bands(0).Columns("IdEstadoMovimiento").Hidden = True
    grdLista.Bands(0).Columns("IdTipoConcepto").Hidden = True
    grdLista.Bands(0).Columns("MovNumero").Header.Caption = "Nota Ingreso"
    grdLista.Bands(0).Columns("MovNumero").Width = 1300
    grdLista.Bands(0).Columns("movTipo").Header.Caption = "Tipo"
    grdLista.Bands(0).Columns("movTipo").Width = 500
    grdLista.Bands(0).Columns("abreviatura").Header.Caption = "Doc.Tipo"
    grdLista.Bands(0).Columns("documentoNumero").Header.Caption = "Doc. N°"
    grdLista.Bands(0).Columns("fechaCreacion").Header.Caption = "Fecha"
    grdLista.Bands(0).Columns("Concepto").Width = 2400
    grdLista.Bands(0).Columns("Total").Width = 1300
    grdLista.Bands(0).Columns("Total").Format = "#0.00"
    
End Sub







Private Sub grdLista_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
        If Val(Row.Cells("IdEstadoMovimiento").GetText()) = 0 Then
            Row.Appearance.ForeColor = vbRed
            'Row.Appearance.Font.Strikethrough = True
        End If
End Sub





Private Sub txtFfinal_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFfinal
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFfinal_LostFocus()
    If Not EsFecha(txtFfinal.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFfinal.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFinicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFinicio
    AdministrarKeyPreview KeyCode
End Sub



Private Sub txtFinicio_LostFocus()
    If Not EsFecha(txtFinicio.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFinicio.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta
    AdministrarKeyPreview KeyCode

End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdLista.Width = fraBusqueda.Width
   grdLista.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub



Sub CargaComboBox()
        On Error GoTo ErrFarm
        Dim oConexion As New ADODB.Connection
        Dim rsIdAlmacen As Recordset
        Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
        Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
        Set oBuscaDondeLabora = Nothing
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        '***debb2014
        If lbNIsoloParaFarmacia = True Then
           Set oRsAlmacenes = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idEstado=1", oConexion)
        Else
           Set oRsAlmacenes = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='A' and idEstado=1", oConexion)
        End If
        '****debb2014
        Set cmbAlmacen.RowSource = oRsAlmacenes
        cmbAlmacen.ListField = "descripcion"
        cmbAlmacen.BoundColumn = "idAlmacen"
        If rsIdAlmacen.RecordCount > 0 Then
           cmbAlmacen.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
           cmbAlmacen.Enabled = False
        End If
ErrFarm:
End Sub
Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdLista, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdLista, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub


Sub inicializar()
    SkinConfigura
    CargaComboBox
    txtFinicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual
    txtFfinal.Text = Date
    lblNombre.Caption = "Nota de Ingreso " & IIf(lbNIsoloParaFarmacia = True, " de Farmacia", "del Almacén")

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


