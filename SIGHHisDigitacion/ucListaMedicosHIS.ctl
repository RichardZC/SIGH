VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl ucMedicosHisLista 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11370
   ScaleHeight     =   5790
   ScaleWidth      =   11370
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
      Left            =   75
      TabIndex        =   2
      Top             =   510
      Width           =   11250
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
         ItemData        =   "ucListaMedicosHIS.ctx":0000
         Left            =   6120
         List            =   "ucListaMedicosHIS.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   1545
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   3000
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   8520
         Picture         =   "ucListaMedicosHIS.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9840
         Picture         =   "ucListaMedicosHIS.ctx":2C4D
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   1275
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
         Left            =   3120
         MaxLength       =   30
         TabIndex        =   1
         Top             =   480
         Width           =   3000
      End
      Begin MSMask.MaskEdBox txtAnio 
         Height          =   330
         Left            =   7680
         TabIndex        =   11
         Top             =   480
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
      Begin VB.Label Label5 
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
         Left            =   7680
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
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
         Height          =   270
         Left            =   6120
         TabIndex        =   10
         Top             =   240
         Width           =   720
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2055
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
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Width           =   2955
      End
   End
   Begin UltraGrid.SSUltraGrid grdMedicos 
      Height          =   4290
      Left            =   75
      TabIndex        =   0
      Top             =   1440
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   7567
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
      Caption         =   "Relación de médicos"
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
      Left            =   15
      TabIndex        =   4
      Top             =   0
      Width           =   11310
   End
End
Attribute VB_Name = "ucMedicosHisLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para Listar Médicos
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_HIS_ProgMedEstMR As New SIGHDatos.HIS_ProgMedEstMR
Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos
Dim ml_IdRegistroSeleccionado As Long
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_cmbServicio As New SIGHEntidades.ListaDespleglable
Dim mo_cmbMes As New SIGHEntidades.ListaDespleglable
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
Property Let Mes(lValue As String)
    ml_Mes = lValue
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
    txtNombres.Text = lValue
'    btnBuscar_Click
End Property
Property Get NombreMedico() As String
    NombreMedico = ml_NombMedico
End Property

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdMedicos.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdMedicos.DataSource
End Property
Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property

Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub InializarForm()
    Dim orsTemp As New Recordset
    Set mo_cmbServicio.MiComboBox = cmbServicio
    Set mo_cmbMes.MiComboBox = cmbMes
    
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
    txtAnio.Text = ml_Anio
End Sub

Public Sub RealizarBusqueda()
    If InStr(txtAnio.Text, "_") >= 1 Then
        MsgBox "El año ingresado no tiene el formato correcto", vbInformation, "Profesional de la Salud"
        Exit Sub
    End If
    Set grdMedicos.DataSource = mo_HIS_ProgMedEstMR.HIS_BuscaResponsableFiltro(ml_IdEstablecimiento, mo_cmbServicio.BoundText, txtAnio.Text, mo_cmbMes.BoundText, txtNombres.Text)
    If mo_HIS_ProgMedEstMR.MensajeError <> "" Then
        MsgBox "Error leyendo datos" + Chr(13) + mo_HIS_ProgMedEstMR.MensajeError, vbInformation, "Profesional de la Salud"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdMedicos, SIGHEntidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
    cmbServicio.ListIndex = 0
    cmbMes.ListIndex = 0
    txtNombres = ""
    txtAnio.Text = CStr(Year(CDate(lcBuscaParametro.RetornaFechaServidorSQL)))
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
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdMedicos.DataSource
    On Error Resume Next
    ml_IdServicio = rsRecordset("idservicio")
    ml_IdMedico = rsRecordset("idmedico")
    ml_NombMedico = rsRecordset("responsable")
    ml_IdTurno = rsRecordset("idturno")
    RaiseEvent SeleccionaRegistro(ml_IdEstablecimiento)
End Sub

Private Sub grdMedicos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdMedicos.Bands(0).Columns("idhisprogmedestmr").Hidden = True
    grdMedicos.Bands(0).Columns("idestablecimiento").Hidden = True
    grdMedicos.Bands(0).Columns("Establecimiento").Header.Caption = "Establecimiento"
    grdMedicos.Bands(0).Columns("Establecimiento").Width = 2000
    grdMedicos.Bands(0).Columns("idservicio").Hidden = True
    grdMedicos.Bands(0).Columns("Servicio").Header.Caption = "Servicio"
    grdMedicos.Bands(0).Columns("Servicio").Width = 2000
    grdMedicos.Bands(0).Columns("idmedico").Hidden = True
    grdMedicos.Bands(0).Columns("colegiatura").Header.Caption = "Colegiatura"
    grdMedicos.Bands(0).Columns("colegiatura").Width = 600
    grdMedicos.Bands(0).Columns("idempleado").Hidden = True
    grdMedicos.Bands(0).Columns("responsable").Header.Caption = "Responsable"
    grdMedicos.Bands(0).Columns("responsable").Width = 2000
    grdMedicos.Bands(0).Columns("fechaprogramada").Hidden = True
    grdMedicos.Bands(0).Columns("idturno").Hidden = True
    grdMedicos.Bands(0).Columns("Turno").Header.Caption = "Turno"
    grdMedicos.Bands(0).Columns("Turno").Width = 1200
End Sub

Private Sub grdMedicos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdMedicos_DblClick
    End If
End Sub

Private Sub txtNombres_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
'
'   fraBusqueda.Width = UserControl.Width - 150
'   lblNombre.Width = UserControl.Width
'   grdMedicos.Width = fraBusqueda.Width
'   grdMedicos.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
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

