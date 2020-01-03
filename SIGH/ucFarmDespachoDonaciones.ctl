VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucFarmDespachoDonaciones 
   ClientHeight    =   6435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11070
   ScaleHeight     =   6435
   ScaleWidth      =   11070
   Begin VB.Frame fraBusqueda 
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
      Height          =   1365
      Left            =   30
      TabIndex        =   0
      Top             =   510
      Width           =   10965
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9585
         Picture         =   "ucFarmDespachoDonaciones.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   450
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   8220
         Picture         =   "ucFarmDespachoDonaciones.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox txtNGuia 
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
         Left            =   840
         MaxLength       =   30
         TabIndex        =   2
         Top             =   900
         Width           =   1125
      End
      Begin VB.TextBox txtNHistoria 
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
         Left            =   3480
         MaxLength       =   9
         TabIndex        =   1
         Top             =   900
         Width           =   1545
      End
      Begin MSMask.MaskEdBox txtFinicio 
         Height          =   315
         Left            =   5220
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         Caption         =   " Almacén                                                                            Fecha Inicial      Fecha Final"
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
         Left            =   165
         TabIndex        =   10
         Top             =   225
         Width           =   7635
      End
      Begin VB.Label lblNcuenta 
         Caption         =   "N° Guía"
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
         TabIndex        =   9
         Top             =   930
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "N° Historia"
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
         Left            =   2280
         TabIndex        =   8
         Top             =   930
         Width           =   1125
      End
   End
   Begin UltraGrid.SSUltraGrid grdLista 
      Height          =   4440
      Left            =   30
      TabIndex        =   11
      Top             =   1935
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   7832
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
      Caption         =   "Despacho Donaciones"
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
      Width           =   11010
   End
End
Attribute VB_Name = "ucFarmDespachoDonaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para despacho de Donaciones
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
        Set oRsBusqueda = mo_ReglasFarmacia.FarmMovimientoDonacionesSeleccionarPorFechas(Val(cmbAlmacen.BoundText), CDate(Format(txtFinicio.Text & " 00:00:01", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)), CDate(Format(txtFfinal.Text & " 23:59:59", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)))
        If mo_Teclado.TextoEsSoloNumeros(txtNhistoria.Text) Then
           oRsBusqueda.Filter = "NroHistoriaClinica=" & HCigualDNI_AgregaNUEVEaLaHistoria(txtNhistoria.Text)
        ElseIf Val(txtNGuia.Text) > 0 Then
           oRsBusqueda.Filter = "guia=" & txtNGuia.Text
        End If
        Set grdLista.DataSource = oRsBusqueda
       ' mo_Apariencia.ConfigurarFilasBiColores grdLista, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        'UserControl.cmbAlmacen.Text = ""
        'UserControl.txtFinicio.Text = Date 'sighEntidades.PrimerFechaDDMMYYDelMesActual
        'UserControl.txtFfinal.Text = Date
        UserControl.txtNGuia.Text = ""
        UserControl.txtNhistoria.Text = ""
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
    grdLista.Bands(0).Columns("MovNumero").Hidden = True
    grdLista.Bands(0).Columns("IdEstadoMovimiento").Hidden = True
    grdLista.Bands(0).Columns("MovNumero").Header.Caption = "Nota Salida"
    grdLista.Bands(0).Columns("MovNumero").Width = 1300
    grdLista.Bands(0).Columns("fechaCreacion").Header.Caption = "Fecha"
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

Private Sub txtNGuia_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNGuia
    AdministrarKeyPreview KeyCode

End Sub


Private Sub txtNGuia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Val(txtNGuia.Text) > 0 Then
       btnBuscar_Click
    End If
End Sub

Private Sub txtNhistoria_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNhistoria
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNHistoria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Val(txtNhistoria.Text) > 0 Then
       btnBuscar_Click
    End If

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
        Set oRsAlmacenes = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idTipoSuministro='02' and idEstado=1 ", oConexion)
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
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
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

Sub Inicializar()
    SkinConfigura
    CargaComboBox
    txtFinicio.Text = Date
    txtFfinal.Text = Date
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



