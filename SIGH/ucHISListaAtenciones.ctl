VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucHISListaAtencion 
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13275
   ScaleHeight     =   8880
   ScaleWidth      =   13275
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
      Height          =   885
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   13095
      Begin VB.TextBox txtLote 
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
         Left            =   5280
         MaxLength       =   3
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cmbMes 
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
         Left            =   7560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox cmbEstablecimiento 
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
         TabIndex        =   1
         Top             =   480
         Width           =   5055
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   10680
         Picture         =   "ucHISListaAtenciones.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   510
         Width           =   1215
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   9330
         Picture         =   "ucHISListaAtenciones.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   510
         Width           =   1305
      End
      Begin MSMask.MaskEdBox txtAnio 
         Height          =   315
         Left            =   6600
         TabIndex        =   3
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Lote"
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
         Left            =   5280
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Mes"
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
         Left            =   7560
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Año"
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
         Left            =   6600
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Establecimiento"
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
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin UltraGrid.SSUltraGrid grdListaHIS 
      Height          =   7200
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   12700
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
      Caption         =   "Lista de Registro HIS"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Hojas de Atención"
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
      TabIndex        =   0
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "ucHISListaAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para registro de  ingresados del his
'        Programado por: Cachay F
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_LoginPC As String
Dim mo_cmbEstablecimiento As New sighEntidades.ListaDespleglable
Dim mo_cmbServicio As New sighEntidades.ListaDespleglable
Dim mo_cmbMes As New sighEntidades.ListaDespleglable
Dim mo_Apariencia As New sighEntidades.GridInfragistic

Dim mo_Teclado As New sighEntidades.Teclado
Dim ml_idRegistroSeleccionado As Long
Dim lblNombre As String
Dim ml_IdEstablecimiento As Long
Dim ms_CodigoEstablecimiento As String
Dim ml_idUsuario As Long

Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos
Dim mo_DatosParametros As New SIGHDatos.Parametros
Dim oRcs_DatosEstablecimiento As New ADODB.Recordset
Dim rsRecordset As ADODB.Recordset

'========================================= PROPIEDADES ==============================
Property Set DataSource(oValue As ADODB.Recordset)
    Set grdListaHIS.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = grdListaHIS.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property
Property Let IdEstablecimiento(lValue As Long)
    ml_IdEstablecimiento = lValue
End Property
Property Get IdEstablecimiento() As Long
    IdEstablecimiento = ml_IdEstablecimiento
End Property
Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
    idUsuario = ml_idUsuario
End Property

Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property

'=========================================== EVENTOS ================================
Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Private Sub btnLimpiar_Click()
UserControl.txtAnio.Text = "____"
End Sub

Private Sub cmbEstablecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtLote
    AdministrarKeyPreview KeyCode
End Sub

Private Sub grdListaHIS_AfterRowActivate()
ml_idRegistroSeleccionado = -1
Set rsRecordset = grdListaHIS.DataSource
On Error Resume Next
ml_idRegistroSeleccionado = IIf(IsNull(rsRecordset("IdHISCabecera")), -1, rsRecordset("IdHISCabecera"))
End Sub

Private Sub grdListaHIS_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    '[IdHISCabecera], [NroFormato], [Turno], [Anio], [Mes], [IdEstablecimiento], [IdServicioPorEstablecimiento], [IdEmpleado]
    
    grdListaHIS.Bands(0).Columns("IdHISCabecera").Hidden = True
    grdListaHIS.Bands(0).Columns("IdUsuario").Hidden = True
    grdListaHIS.Bands(0).Columns("IdHIsLote").Hidden = True
    grdListaHIS.Bands(0).Columns("NroHojaHIS").Width = 700
    
    grdListaHIS.Bands(0).Columns("NroHojaHIS").Header.Caption = "Nro Hoja"
    grdListaHIS.Bands(0).Columns("NroHojaHIS").Width = 1000
    
    grdListaHIS.Bands(0).Columns("NroFormato").Header.Caption = "Nro Reg"
    grdListaHIS.Bands(0).Columns("NroFormato").Width = 1000
    
    grdListaHIS.Bands(0).Columns("Turno").Header.Caption = "Turno"
    grdListaHIS.Bands(0).Columns("Turno").Width = 1000
    
    grdListaHIS.Bands(0).Columns("Anio").Header.Caption = "Año"
    grdListaHIS.Bands(0).Columns("Anio").Width = 700
    
    grdListaHIS.Bands(0).Columns("Mes").Header.Caption = "Mes"
    grdListaHIS.Bands(0).Columns("Mes").Width = 1500
    
    grdListaHIS.Bands(0).Columns("NombreEstablecimiento").Header.Caption = "Establecimiento"
    grdListaHIS.Bands(0).Columns("NombreEstablecimiento").Width = 2500
    
    grdListaHIS.Bands(0).Columns("NombreServicio").Header.Caption = "Servicio"
    grdListaHIS.Bands(0).Columns("NombreServicio").Width = 1500
    
    grdListaHIS.Bands(0).Columns("NombreEmpleado").Header.Caption = "Especialista Encargado"
    grdListaHIS.Bands(0).Columns("NombreEmpleado").Width = 2800
End Sub

'=========================================== METODOS ================================
Public Function inicializar()
    Set mo_cmbEstablecimiento.MiComboBox = UserControl.cmbEstablecimiento
    Set mo_cmbMes.MiComboBox = UserControl.cmbMes
    txtAnio.Text = Year(CDate(mo_DatosParametros.RetornaFechaServidorSQL))
    CargarComboBoxes
End Function

Public Sub RealizarBusqueda()
    If mo_cmbEstablecimiento.BoundText = "" Then Exit Sub
    Set grdListaHIS.DataSource = mo_ReglasHIS.ConsultarRegistroFiltroAtenciones(mo_cmbEstablecimiento.BoundText, _
                                                                                  txtLote.Text, _
                                                                                UserControl.txtAnio.Text, _
                                                                                mo_cmbMes.BoundText)

    If mo_ReglasHIS.MensajeError <> "" Then
        MsgBox mo_ReglasHIS.MensajeError, vbCritical, "Filtro Registros HIS"
    End If
End Sub

Private Sub txtAnio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbMes
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtLote_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtAnio
    AdministrarKeyPreview KeyCode
End Sub

Private Sub UserControl_Resize()
   On Error Resume Next
   fraBusqueda.Width = UserControl.Width - 110
   UserControl.lblNombre.Width = UserControl.Width
   grdListaHIS.Width = fraBusqueda.Width
   grdListaHIS.Height = UserControl.Height - (UserControl.lblNombre.Height + fraBusqueda.Height + 150)
End Sub

Private Sub CargarComboBoxes()
    Dim orsTemp As New ADODB.Recordset
    Dim orstemp1 As New ADODB.Recordset
    
    mo_cmbMes.BoundColumn = "IdMes"
    mo_cmbMes.ListField = "NombreMes"
    Set mo_cmbMes.RowSource = mo_ReglasHIS.ListaMeses
    cmbMes.ListIndex = CInt(Month(CDate(mo_DatosParametros.RetornaFechaServidorSQL))) - 1
    
    mo_cmbEstablecimiento.BoundColumn = "IdEstablecimiento"
    mo_cmbEstablecimiento.ListField = "NombreEstablecimiento"
    Set orsTemp = ListadoEstablecimientos()
    Set mo_cmbEstablecimiento.RowSource = orsTemp
    If orsTemp.RecordCount = 0 Then
        MsgBox "No tiene establecimientos ni servicios configurados", vbExclamation, "HIS"
    End If
    If orsTemp.RecordCount > 0 Then
        cmbEstablecimiento.ListIndex = 0
        RealizarBusqueda
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdListaHIS, sighEntidades.GrillaConFilasBicolor
End Sub

Public Function DevuelveIdEstablecimiento() As Integer
    DevuelveIdEstablecimiento = Val(mo_cmbEstablecimiento.BoundText)
End Function

Private Function ListadoEstablecimientos() As Recordset
    Dim oTabla As New DOEstablecimiento
    Set oRcs_DatosEstablecimiento = mo_ReglasHIS.ObtenerListaEstablecimientosMR
    If oRcs_DatosEstablecimiento.RecordCount > 0 Then
        oRcs_DatosEstablecimiento.MoveFirst
    End If
    Set ListadoEstablecimientos = oRcs_DatosEstablecimiento
End Function

Sub AdministrarKeyPreview(KeyCode As Integer)
    Select Case KeyCode
     Case vbKeyF6
        btnBuscar_Click
     Case vbKeyF7
        btnLimpiar_Click
    End Select
End Sub
