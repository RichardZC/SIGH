VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form MedicosHISBusqueda 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12390
   Icon            =   "MedicosHISBusqueda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   12390
   StartUpPosition =   2  'CenterScreen
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
      Left            =   0
      TabIndex        =   6
      Top             =   510
      Width           =   12330
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtNombres 
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
         Left            =   6720
         MaxLength       =   30
         TabIndex        =   2
         Top             =   480
         Width           =   2985
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   11040
         Picture         =   "MedicosHISBusqueda.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   9720
         Picture         =   "MedicosHISBusqueda.frx":38A6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   1305
      End
      Begin VB.ComboBox cmbServicio 
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
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3000
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
         ItemData        =   "MedicosHISBusqueda.frx":64EF
         Left            =   9960
         List            =   "MedicosHISBusqueda.frx":64F1
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1545
      End
      Begin MSMask.MaskEdBox txtAnio 
         Height          =   330
         Left            =   11520
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
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
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Responsable de atención"
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
         Left            =   6720
         TabIndex        =   12
         Top             =   240
         Width           =   2955
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   30
      TabIndex        =   5
      Top             =   5280
      Width           =   12285
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "MedicosHISBusqueda.frx":64F3
         DownPicture     =   "MedicosHISBusqueda.frx":69B7
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   6555
         Picture         =   "MedicosHISBusqueda.frx":6EA3
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "MedicosHISBusqueda.frx":738F
         DownPicture     =   "MedicosHISBusqueda.frx":77EF
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   4980
         Picture         =   "MedicosHISBusqueda.frx":7C64
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdMedicos 
      Height          =   3810
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   6720
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
      Caption         =   "grdMedicos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Responsables de atención"
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
      TabIndex        =   13
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "MedicosHISBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca Médicos HIS
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim mb_Loading As Boolean
Dim mo_cmbEstablecimiento As New SIGHEntidades.ListaDespleglable
Dim mo_cmbServicio As New SIGHEntidades.ListaDespleglable
Dim mo_cmbMes As New SIGHEntidades.ListaDespleglable
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_HIS_ProgMedEstMR As New SIGHDatos.HIS_ProgMedEstMR
Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim ml_IdRegistroSeleccionado As Long
Dim mo_Teclado As New SIGHEntidades.Teclado
Public Event SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_Anio As String
Dim ml_Mes As String
Dim ml_IdEstablecimiento As Long
Dim ml_IdServicio As Long
Dim ml_IdMedico As Long
Dim ml_NombMedico As String
Dim ml_IdTurno As Integer

Property Let Anio(lValue As String)
    ml_Anio = lValue
End Property
Property Get Anio() As String
    Anio = ml_Anio
End Property
Property Let Mes(lValue As String)
    ml_Mes = lValue
End Property
Property Get Mes() As String
    Mes = ml_Mes
End Property
Property Let IdEstablecimiento(lValue As Long)
    ml_IdEstablecimiento = lValue
End Property
Property Get IdEstablecimiento() As Long
    IdEstablecimiento = ml_IdEstablecimiento
End Property
Property Let IdServicio(lValue As Long)
    ml_IdServicio = lValue
End Property
Property Get IdServicio() As Long
    IdServicio = ml_IdServicio
End Property
Property Let IdMedico(lValue As Long)
    ml_IdMedico = lValue
End Property
Property Get IdMedico() As Long
    IdMedico = ml_IdMedico
End Property
Property Let IdTurno(lValue As Integer)
    ml_IdTurno = lValue
End Property
Property Get IdTurno() As Integer
    IdTurno = ml_IdTurno
End Property
Property Let NombreMedico(lValue As String)
    ml_NombMedico = lValue
End Property
Property Get NombreMedico() As String
    NombreMedico = ml_NombMedico
End Property
Property Set DataSource(oValue As ADODB.Recordset)
    Set DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = DataSource
End Property
Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    Me.Visible = False
End Sub

Private Sub Form_Activate()
    If mb_Loading Then
        If ml_IdEstablecimiento <> 0 Then
            LimpiarFiltro
            RealizarBusqueda
            On Error Resume Next
        End If
        mb_Loading = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    InializarForm
    Me.Caption = "Responsables de atención (" & cmbEstablecimiento.Text & " " & txtAnio.Text & ")"
    mb_Loading = True
End Sub

''''
Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub InializarForm()
    Dim orsTemp As New Recordset
    Set mo_cmbEstablecimiento.MiComboBox = cmbEstablecimiento
    Set mo_cmbServicio.MiComboBox = cmbServicio
    Set mo_cmbMes.MiComboBox = cmbMes
    
    mo_cmbEstablecimiento.BoundColumn = "IdEstablecimiento"
    mo_cmbEstablecimiento.ListField = "NombreEstablecimiento"
    Set mo_cmbEstablecimiento.RowSource = mo_ReglasHIS.ObtenerListaEstablecimientosMR
    mo_cmbEstablecimiento.BoundText = ml_IdEstablecimiento
    
    mo_cmbServicio.BoundColumn = "IdServicio"
    mo_cmbServicio.ListField = "Nombre"
    Set orsTemp = mo_ReglasHIS.ListaServiciosPorEstablecimiento(ml_IdEstablecimiento)
    Set mo_cmbServicio.RowSource = orsTemp
    orsTemp.MoveFirst
    mo_cmbServicio.BoundText = orsTemp.Fields!IdServicio
    
    mo_cmbMes.BoundColumn = "IdMes"
    mo_cmbMes.ListField = "NombreMes"
    Set orsTemp = mo_ReglasHIS.ListaMeses
    Set mo_cmbMes.RowSource = orsTemp
    orsTemp.MoveFirst
    mo_cmbMes.BoundText = Val(ml_Mes)
'    cmbMes.ListIndex = 0
    txtNombres.Text = ml_NombMedico
    txtAnio.Text = ml_Anio
    
    mo_Formulario.HabilitarDeshabilitar cmbEstablecimiento, False
End Sub

Public Sub RealizarBusqueda()
    If InStr(txtAnio.Text, "_") >= 1 Then
        MsgBox "El año ingresado no tiene el formato correcto", vbInformation, "Profesional de la Salud"
        Exit Sub
    End If
    Set grdMedicos.DataSource = mo_ReglasHIS.HIS_BuscaResponsableFiltro(ml_IdEstablecimiento, IIf(mo_cmbServicio.BoundText = "", 0, mo_cmbServicio.BoundText), txtAnio.Text, mo_cmbMes.BoundText, txtNombres.Text)
    If mo_HIS_ProgMedEstMR.MensajeError <> "" Then
        MsgBox "Error leyendo datos" + Chr(13) + mo_HIS_ProgMedEstMR.MensajeError, vbInformation, "Profesional de la Salud"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdMedicos, SIGHEntidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
'    cmbServicio.ListIndex = 0
    cmbServicio.ListIndex = -1
'    cmbMes.ListIndex = 0 'Actualizado 01102014
    txtNombres = ""
'    txtAnio.Text = CStr(Year(CDate(lcBuscaParametro.RetornaFechaServidorSQL))) 'Actualizado 01102014
End Sub

Private Sub grdMedicos_AfterRowActivate()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdMedicos.DataSource
    On Error Resume Next
    ml_IdServicio = rsRecordset("idservicio")
    ml_IdMedico = rsRecordset("idmedico")
    ml_NombMedico = rsRecordset("responsable")
    ml_IdTurno = rsRecordset("idturno")
End Sub

Private Sub grdMedicos_DblClick()
    btnAceptar_Click 'Actualizado 01102014
End Sub

Private Sub grdMedicos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
'    grdMedicos.Bands(0).Columns("idhisprogmedestmr").Hidden = True
    grdMedicos.Bands(0).Columns("idestablecimiento").Hidden = True
    grdMedicos.Bands(0).Columns("Establecimiento").Header.Caption = "Establecimiento"
    grdMedicos.Bands(0).Columns("Establecimiento").Width = 2200
    grdMedicos.Bands(0).Columns("idservicio").Hidden = True
    grdMedicos.Bands(0).Columns("Servicio").Header.Caption = "Servicio"
    grdMedicos.Bands(0).Columns("Servicio").Width = 2200
    grdMedicos.Bands(0).Columns("idmedico").Hidden = True
    grdMedicos.Bands(0).Columns("colegiatura").Header.Caption = "Colegiatura"
    grdMedicos.Bands(0).Columns("colegiatura").Width = 800
    grdMedicos.Bands(0).Columns("idempleado").Hidden = True
    grdMedicos.Bands(0).Columns("responsable").Header.Caption = "Responsable"
    grdMedicos.Bands(0).Columns("responsable").Width = 2800
    grdMedicos.Bands(0).Columns("idturno").Hidden = True
    grdMedicos.Bands(0).Columns("Turno").Header.Caption = "Turno"
    grdMedicos.Bands(0).Columns("Turno").Width = 1200
    grdMedicos.Bands(0).Columns("anio").Header.Caption = "Año"
    grdMedicos.Bands(0).Columns("anio").Width = 800
    grdMedicos.Bands(0).Columns("mes").Header.Caption = "Mes"
    grdMedicos.Bands(0).Columns("mes").Width = 800
End Sub

Private Sub grdMedicos_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
If KeyCode = 13 Or KeyCode = vbKeyF2 Then
    btnAceptar_Click
End If
End Sub

'Private Sub grdMedicos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
'    If KeyAscii = 13 Then
'       grdMedicos_DblClick
'    End If
'End Sub

Private Sub txtNombres_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        btnCancelar_Click
    Case vbKeyF2
        btnAceptar_Click
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


