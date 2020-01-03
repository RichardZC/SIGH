VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form AHCconVIH 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historias Clínicas de Pacientes según Tipo Historia"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   75
      TabIndex        =   17
      Top             =   5040
      Width           =   5880
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AHCconVIH.frx":0000
         DownPicture     =   "AHCconVIH.frx":0460
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
         Left            =   1538
         Picture         =   "AHCconVIH.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AHCconVIH.frx":0D4A
         DownPicture     =   "AHCconVIH.frx":120E
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
         Left            =   3068
         Picture         =   "AHCconVIH.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4950
      Left            =   60
      TabIndex        =   14
      Top             =   30
      Width           =   5880
      Begin VB.TextBox txtAnio 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2220
         MaxLength       =   4
         TabIndex        =   24
         Top             =   4005
         Width           =   1005
      End
      Begin VB.Frame frFechaNacimiento 
         Caption         =   "Fecha Nacimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   930
         TabIndex        =   4
         Top             =   1785
         Width           =   4845
         Begin VB.OptionButton optFechaNacimiento 
            Caption         =   "Ambos"
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
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optFechaNacimiento 
            Caption         =   "Real"
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
            Index           =   1
            Left            =   1920
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optFechaNacimiento 
            Caption         =   "Calculada"
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
            Index           =   2
            Left            =   3600
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txtComunidad 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2220
         MaxLength       =   7
         TabIndex        =   10
         Top             =   3210
         Width           =   1005
      End
      Begin VB.TextBox txtSector 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2220
         MaxLength       =   4
         TabIndex        =   9
         Top             =   2820
         Width           =   465
      End
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
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
         Left            =   210
         Picture         =   "AHCconVIH.frx":1BE6
         TabIndex        =   11
         Top             =   4530
         Width           =   1755
      End
      Begin VB.ComboBox cmbOrden 
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
         ItemData        =   "AHCconVIH.frx":1EF8
         Left            =   2220
         List            =   "AHCconVIH.frx":1F05
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1380
         Width           =   3555
      End
      Begin VB.ComboBox cmbIdTipoHistoria 
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
         Left            =   2220
         TabIndex        =   1
         Top             =   540
         Width           =   3540
      End
      Begin VB.ComboBox cmbIdResponsable 
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
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   945
         Width           =   3555
      End
      Begin VB.PictureBox progressRpt 
         Height          =   300
         Left            =   4410
         ScaleHeight     =   240
         ScaleWidth      =   660
         TabIndex        =   15
         Top             =   4455
         Visible         =   0   'False
         Width           =   720
      End
      Begin Threed.SSOption optTipoHistoria 
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Top             =   210
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   503
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Historias Clínicas de Pacientes según Tipo Historia"
         Value           =   -1
      End
      Begin Threed.SSOption optFichaFamiliar 
         Height          =   345
         Left            =   180
         TabIndex        =   8
         Top             =   2490
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   609
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lista Fichas Familiares por Sector y/o Comunidad"
      End
      Begin Threed.SSOption optSexoEdad 
         Height          =   345
         Left            =   180
         TabIndex        =   22
         Top             =   3660
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   609
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Indicador por Sexo y Grupo Edad"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   1845
         TabIndex        =   23
         Top             =   4050
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sector"
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
         Left            =   1680
         TabIndex        =   21
         Top             =   2865
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comunidad"
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
         Left            =   1320
         TabIndex        =   20
         Top             =   3285
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Orden del Rep"
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
         Left            =   960
         TabIndex        =   19
         Top             =   1425
         Width           =   1185
      End
      Begin VB.Label lblIdTipoHistoria 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Historia"
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
         Left            =   960
         TabIndex        =   18
         Top             =   570
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Responsable"
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
         Left            =   960
         TabIndex        =   16
         Top             =   1005
         Width           =   1005
      End
   End
End
Attribute VB_Name = "AHCconVIH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Lista Historias para pacientes con VIH
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_cmbIdTipoHistoria As New sighEntidades.ListaDespleglable
Dim mo_cmbIdResponsable As New sighEntidades.ListaDespleglable
Dim sMensaje As String
Dim mo_Teclado As New sighEntidades.Teclado



Private Sub btnAceptar_Click()
    If optTipoHistoria.Value = True Then
        If mo_cmbIdTipoHistoria.BoundText = "" Then
            MsgBox "Por favor elija el Tipo de Historia", vbInformation, Me.Caption
            Exit Sub
        End If
        Me.MousePointer = 11
        Dim oRptHistoriasClinicasConVIH As New RptAHCconVIH
        oRptHistoriasClinicasConVIH.IdResponsable = Val(mo_cmbIdResponsable.BoundText)
        oRptHistoriasClinicasConVIH.IdTipoHistoria = Val(mo_cmbIdTipoHistoria.BoundText)
        oRptHistoriasClinicasConVIH.OrdenFiltro = IIf(cmbOrden.ListIndex = 0, "HC", IIf(cmbOrden.ListIndex = 2, "AUTOGENERADO", "Paciente"))
        oRptHistoriasClinicasConVIH.TipoFechaNacimiento = getSeleccionTipoFecha()
        oRptHistoriasClinicasConVIH.TextoDelFiltro = getTextoFiltroHistoriasClinicas()
        oRptHistoriasClinicasConVIH.CrearReporte IIf(chkExcel.Value = 1, True, False), Me.hwnd
        Set oRptHistoriasClinicasConVIH = Nothing
        Me.MousePointer = 1
    ElseIf optSexoEdad.Value = True Then
        Me.MousePointer = 11
        Dim oRptXsexoGrupoEdad As New RptAHCconVIH
        oRptXsexoGrupoEdad.AfiliadosXsexoGrupoEdad Me.txtAnio.Text
        Set oRptXsexoGrupoEdad = Nothing
        Me.MousePointer = 1
    Else
        If optFichaFamiliar.Value = True And txtSector.Text = "" Then
            MsgBox "Por favor ingresar el SECTOR", vbInformation, Me.Caption
            Exit Sub
        End If
        Me.MousePointer = 11
        Dim oRptAHCpacienteHastaNanio As New RptAHCpacienteHastaNanio
        oRptAHCpacienteHastaNanio.EdadMaxima = 0
        oRptAHCpacienteHastaNanio.TextoDelFiltro = "Filtros:  Sector=" & Trim(txtSector.Text) & IIf(txtComunidad.Text <> "", " Comunidad=" & txtComunidad.Text, "")
        oRptAHCpacienteHastaNanio.CrearReporte IIf(chkExcel.Value = 1, True, False), False, txtSector.Text, txtComunidad.Text, Me.hwnd
        Set oRptAHCpacienteHastaNanio = Nothing
        Me.MousePointer = 1
    End If
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub



Private Sub chkExcel_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkExcel
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdResponsable
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoHistoria
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbOrden
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdTipoHistoria.MiComboBox = cmbIdTipoHistoria
    Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable

End Sub


Private Sub Form_Load()
       mo_cmbIdTipoHistoria.BoundColumn = "IdTipoHistoria"
       mo_cmbIdTipoHistoria.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoHistoria.RowSource = mo_AdminArchivoClinico.TiposHistoriaClinicaSeleccionarTodos()
       mo_cmbIdTipoHistoria.BoundText = 1
       sMensaje = sMensaje + mo_AdminArchivoClinico.MensajeError
       
       mo_cmbIdResponsable.BoundColumn = "IdEmpleado"
       mo_cmbIdResponsable.ListField = "ApAmNo"
       Set mo_cmbIdResponsable.RowSource = mo_AdminArchivoClinico.ArchiverosSeleccionarTodos()
       
       cmbOrden.ListIndex = 1
       optFechaNacimiento(0).Value = True
       mo_cmbIdTipoHistoria.BoundText = "10"
       
End Sub



Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub


Private Sub optFechaNacimiento_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, optFechaNacimiento
    AdministrarKeyPreview KeyCode
End Sub

Private Sub optFichaFamiliar_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, optFichaFamiliar
    AdministrarKeyPreview KeyCode
End Sub

Private Sub optTipoHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, optTipoHistoria
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtComunidad_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtComunidad
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtSector_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSector
    AdministrarKeyPreview KeyCode
End Sub

Private Function getTextoFiltroHistoriasClinicas() As String
    Dim textFilter As String
    textFilter = "Filtros:  Tipo de Historias: " & cmbIdTipoHistoria.Text
    textFilter = textFilter & "     Responsable: " & cmbIdResponsable
    textFilter = textFilter & "     Ordenado Por: " & IIf(cmbOrden.ListIndex = 0, "Historias Clínicas", IIf(cmbOrden.ListIndex = 2, "Autogenerado", "Apellidos y Nombres"))
    If optFechaNacimiento(0).Value = False Then
        textFilter = textFilter & "     Fecha Nacimiento: " & getFilterFechaNacimiento()
    End If
    getTextoFiltroHistoriasClinicas = textFilter
End Function

Private Function getFilterFechaNacimiento() As String
    Dim i As Integer
    
    i = 0
    For i = 0 To optFechaNacimiento.Count - 1
        If optFechaNacimiento(i).Value = True Then
            getFilterFechaNacimiento = optFechaNacimiento(i).Caption
            Exit For
        End If
    Next i
End Function



Private Function getSeleccionTipoFecha() As Integer
    Dim i As Integer
    
    i = 0
    For i = 0 To optFechaNacimiento.Count - 1
        If optFechaNacimiento(i).Value = True Then
            Exit For
        End If
    Next i
    getSeleccionTipoFecha = i
End Function
