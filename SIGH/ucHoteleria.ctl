VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.UserControl ucCamasLista 
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   8985
   ScaleWidth      =   12360
   Begin TabDlg.SSTab tabHoteleria 
      Height          =   8400
      Left            =   60
      TabIndex        =   4
      Top             =   570
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   14817
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Relación de camas"
      TabPicture(0)   =   "ucHoteleria.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdCamas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraFiltro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Plano"
      TabPicture(1)   =   "ucHoteleria.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MapaDeCamas"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraFiltro 
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
         Height          =   765
         Left            =   150
         TabIndex        =   7
         Top             =   360
         Width           =   11895
         Begin VB.CommandButton btnBuscar 
            Height          =   315
            Left            =   8730
            Picture         =   "ucHoteleria.ctx":0038
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   270
            Width           =   1305
         End
         Begin VB.ComboBox cmbIdServicio 
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
            Left            =   5190
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   285
            Width           =   3450
         End
         Begin VB.ComboBox cmbIdTipoServicio 
            Enabled         =   0   'False
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
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   300
            Width           =   3090
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo servicio"
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
            Top             =   330
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "Servicio"
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
            Left            =   4455
            TabIndex        =   8
            Top             =   330
            Width           =   915
         End
      End
      Begin UltraGrid.SSUltraGrid grdCamas 
         Height          =   6885
         Left            =   150
         TabIndex        =   3
         Top             =   1200
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   12144
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
         Caption         =   "Lista Camas"
      End
      Begin SHDocVwCtl.WebBrowser MapaDeCamas 
         Height          =   7605
         Left            =   -74820
         TabIndex        =   5
         Top             =   420
         Width           =   11895
         ExtentX         =   20981
         ExtentY         =   13414
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Administración de camas"
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
      Left            =   15
      TabIndex        =   6
      Top             =   15
      Width           =   12330
   End
End
Attribute VB_Name = "ucCamasLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para listar camas
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Public Event SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
Dim mo_AdminProgramacionMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminReglasHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim ml_idRegistroSeleccionado As Long
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim ml_idUsuario As Long
Dim mo_cmbIdTipoServicio As New ListaDespleglable
Dim mo_cmbIdServicio As New ListaDespleglable
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_NombreServicio As String
Dim ml_CodigoServicio As String
'mgaray20141014
Dim mb_PresionoBotonBuscar As Boolean
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_EsListaParaMantenimiento As Boolean

Property Let EsListaParaMantenimiento(lValue As Boolean)
   ml_EsListaParaMantenimiento = lValue
End Property

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdCamas.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdCamas.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property
Property Let idTipoServicio(lValue As Long)
   mo_cmbIdTipoServicio.BoundText = lValue
End Property
Property Get idTipoServicio() As Long
If Not mo_cmbIdTipoServicio Is Nothing Then
   idTipoServicio = Val(mo_cmbIdTipoServicio.BoundText)
End If
End Property
Property Let IdServicio(lValue As Long)
   mo_cmbIdServicio.BoundText = lValue
End Property
Property Get IdServicio() As Long
   IdServicio = Val(mo_cmbIdServicio.BoundText)
End Property
Property Let HabilitarTipoServicio(lValue As Boolean)
   cmbIdTipoServicio.Enabled = lValue
End Property
Property Get HabilitarTipoServicio() As Boolean
   HabilitarTipoServicio = cmbIdTipoServicio.Enabled
End Property
Property Let HabilitarServicio(lValue As Boolean)
   cmbIdServicio.Enabled = lValue
End Property
Property Get HabilitarServicio() As Boolean
   HabilitarServicio = cmbIdServicio.Enabled
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_idUsuario
End Property

Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property

Property Get CodigoServicio() As String
   CodigoServicio = ml_CodigoServicio
End Property
Property Get NombreServicio() As String
   NombreServicio = ml_NombreServicio
End Property


Sub ConfigurarTipoServicio()
    
    mo_cmbIdTipoServicio.ListField = "DescripcionLarga"
    mo_cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
    Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarTodos()
    
    UserControl.MapaDeCamas.Navigate App.Path + "\archivos\" + "index.html"
    
End Sub
'mgaray20141014
Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    mb_PresionoBotonBuscar = True
    RealizarBusqueda
    mb_PresionoBotonBuscar = False
    Screen.MousePointer = vbDefault
End Sub
Public Sub RealizarBusqueda()
Dim sRutaSVG As String
Dim sSVG As String
Dim lIdServicio As Long

    sRutaSVG = App.Path + "\archivos\" + "Mapa.svg"
    lIdServicio = Val(mo_cmbIdServicio.BoundText)
    ml_CodigoServicio = ""
    ml_NombreServicio = ""
    If lIdServicio <> 0 Or Val(mo_cmbIdTipoServicio.BoundText) = 0 Then
        'mgaray20141014
        Call SetDataServicioBusqueda
'        ml_NombreServicio = Mid(cmbIdServicio.Text, InStr(cmbIdServicio.Text, "=") + 2, 200)
'        ml_CodigoServicio = Left(cmbIdServicio.Text, InStr(cmbIdServicio.Text, "=") - 1)
        Dim rsCamas As New Recordset
        Set rsCamas = mo_AdminReglasHoteleria.CamasSeleccionarDisponibilidadPorServicioUbicacionActual(lIdServicio)
        Set UserControl.grdCamas.DataSource = rsCamas
      '  mo_Apariencia.ConfigurarFilasBiColores grdCamas, sighentidades.GrillaConFilasBicolor
        
        mo_AdminReglasHoteleria.CrearArchivoSVGPorServicio lIdServicio, Val(mo_cmbIdTipoServicio.BoundText), ml_idUsuario, sRutaSVG
        UserControl.MapaDeCamas.Navigate App.Path + "\archivos\" + "index.html"
        On Error Resume Next
        grdCamas.SetFocus
    Else
        'mgaray20141014
        If mb_PresionoBotonBuscar = True Then
            MsgBox "Por favor ingrese el servicio y el tipo de servicio", vbInformation, "Búsqueda de cama"
        Else
            Set rsCamas = mo_AdminReglasHoteleria.CamasSeleccionarDisponibilidadPorServicioUbicacionActual(0)
            Set UserControl.grdCamas.DataSource = rsCamas
            'mo_Apariencia.ConfigurarFilasBiColores grdCamas, sighentidades.GrillaConFilasBicolor
        End If
    End If

End Sub


Private Sub cmbIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdServicio
    AdministrarKeyPreview KeyCode

End Sub

'09/08/2011
Private Sub cmbIdTipoServicio_Click()
    mo_cmbIdServicio.BoundColumn = "IdServicio"
    mo_cmbIdServicio.ListField = "DescripcionLarga"
    Dim lcEspecialidadesDelUsuario As String
    If Val(mo_cmbIdTipoServicio.BoundText) = sghTipoServicio.sghHospitalizacion Then
       lcEspecialidadesDelUsuario = mo_ReglasAdmision.DevuelveEspecialidadesServicioSegunUsuarioSistema(sghEspecialidadesHosp, ml_idUsuario)
    Else
       lcEspecialidadesDelUsuario = mo_ReglasAdmision.DevuelveEspecialidadesServicioSegunUsuarioSistema(sghEspecialidadesEmergCons, ml_idUsuario)
    End If
    Set mo_cmbIdServicio.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarPorTipoV2debb(Val(mo_cmbIdTipoServicio.BoundText), lcEspecialidadesDelUsuario, sghFiltraSoloActivos)
End Sub

Private Sub cmbIdTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoServicio
    AdministrarKeyPreview KeyCode

End Sub

Private Sub grdCamas_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdCamas.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdCama")
End Sub

Private Sub grdCamas_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
Cancel = True
End Sub

Private Sub grdCamas_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdCamas.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdCama")
    
End Sub

Private Sub grdCamas_DblClick()
     grdCamas_Click
     RaiseEvent SeleccionaRegistro(ml_idRegistroSeleccionado)
End Sub

Private Sub grdCamas_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdCamas.Bands(0).Columns("IdCama").Hidden = True
    
    grdCamas.Bands(0).Columns("Codigo").Header.Caption = "Cod.Cama"
    grdCamas.Bands(0).Columns("Codigo").Width = 750
    
    grdCamas.Bands(0).Columns("Estado").Header.Caption = "Estado"
    grdCamas.Bands(0).Columns("Estado").Width = 2000
    
    grdCamas.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdCamas.Bands(0).Columns("ApellidoPaterno").Width = 1500
    
    grdCamas.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdCamas.Bands(0).Columns("ApellidoMaterno").Width = 1500
    
    grdCamas.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdCamas.Bands(0).Columns("PrimerNombre").Width = 1500
    
    grdCamas.Bands(0).Columns("SegundoNombre").Header.Caption = "2do Nombre"
    grdCamas.Bands(0).Columns("SegundoNombre").Width = 1500
    
    grdCamas.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nº Historia Clínica"
    grdCamas.Bands(0).Columns("NroHistoriaClinica").Width = 2000
    
    
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
'        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
'        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdCamas, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdCamas, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub

Public Function Inicializar()
    SkinConfigura
    Set mo_cmbIdTipoServicio.MiComboBox = cmbIdTipoServicio
    Set mo_cmbIdServicio.MiComboBox = cmbIdServicio

    ConfigurarTipoServicio
    ConfiguraPermisos
    If mo_cmbIdTipoServicio.BoundText <> "" Then
       cmbIdTipoServicio_Click
    End If
    'debb-29/03/2017
    If ml_EsListaParaMantenimiento = False Then
       cmbIdServicio.Enabled = IIf(lcBuscaParametro.SeleccionaFilaParametro(515) = "S", True, False)
    Else
       cmbIdServicio.Enabled = True
    End If
End Function

Private Sub grdCamas_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdCamas_DblClick
    End If
End Sub



Private Sub UserControl_Resize()
   
    On Error Resume Next
    lblNombre.Width = UserControl.Width
    
    tabHoteleria.Width = UserControl.Width - 100
    tabHoteleria.Height = UserControl.Height - 200 - lblNombre.Height
    
    fraFiltro.Width = tabHoteleria.Width - 300
    
    grdCamas.Height = tabHoteleria.Height - fraFiltro.Height - 550
    grdCamas.Width = tabHoteleria.Width - 300
    
    MapaDeCamas.Height = tabHoteleria.Height - 550
    MapaDeCamas.Width = tabHoteleria.Width - 300
   
End Sub

Public Sub ClicEnBotonBuscar()
      btnBuscar_Click
End Sub


Sub ConfiguraPermisos()
    'PERMISOS
    Dim oRsPermisos As New Recordset
    Set oRsPermisos = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_idUsuario)
    UserControl.tabHoteleria.TabVisible(1) = False
    If oRsPermisos.RecordCount > 0 Then
       Do While Not oRsPermisos.EOF
          Select Case oRsPermisos.Fields!IdPermiso
          Case 364    'Camas - Ver TAB 'Plano'
               UserControl.tabHoteleria.TabVisible(1) = True
          End Select
          oRsPermisos.MoveNext
       Loop
    End If
    Set oRsPermisos = Nothing
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
     Case vbKeyF8
    End Select
End Sub

'mgaray20141014
Public Function SetDataServicioBusqueda()
    If Val(mo_cmbIdServicio.BoundText) <> 0 Then
        ml_NombreServicio = Mid(cmbIdServicio.Text, InStr(cmbIdServicio.Text, "=") + 2, 200)
        ml_CodigoServicio = Left(cmbIdServicio.Text, InStr(cmbIdServicio.Text, "=") - 1)
    End If
End Function



Public Function TieneRegistros() As Boolean
    Dim orsTemp As New Recordset
    Set orsTemp = grdCamas.DataSource
    TieneRegistros = False
    If orsTemp.RecordCount > 0 Then
        TieneRegistros = True
    End If
End Function
