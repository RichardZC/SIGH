VERSION 5.00
Begin VB.Form HEgresosHosp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Egresos Hospitalarios por Departamentos y/o Servicios"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   30
      TabIndex        =   2
      Top             =   3795
      Width           =   12465
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HEgresosHosp.frx":0000
         DownPicture     =   "HEgresosHosp.frx":04C4
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
         Left            =   6353
         Picture         =   "HEgresosHosp.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HEgresosHosp.frx":0E9C
         DownPicture     =   "HEgresosHosp.frx":12FC
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
         Left            =   4823
         Picture         =   "HEgresosHosp.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3765
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   12435
      Begin VB.ComboBox cmbDistrito 
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
         ItemData        =   "HEgresosHosp.frx":1BE6
         Left            =   1245
         List            =   "HEgresosHosp.frx":1BF0
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   705
         Width           =   1875
      End
      Begin VB.Frame frmDistrito 
         Height          =   630
         Left            =   3285
         TabIndex        =   25
         Top             =   525
         Visible         =   0   'False
         Width           =   9045
         Begin VB.ComboBox cmbIdDist 
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
            Left            =   5715
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   180
            Width           =   3240
         End
         Begin VB.ComboBox cmbIdProv 
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
            Left            =   2865
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   180
            Width           =   1950
         End
         Begin VB.ComboBox cmbIdDpto 
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
            Left            =   570
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   180
            Width           =   1605
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
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
            Left            =   5115
            TabIndex        =   31
            Top             =   225
            Width           =   570
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Prov"
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
            Left            =   2460
            TabIndex        =   30
            Top             =   225
            Width           =   360
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Dpto"
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
            Left            =   105
            TabIndex        =   29
            Top             =   225
            Width           =   405
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1905
         Left            =   1215
         TabIndex        =   8
         Top             =   1695
         Width           =   11130
         Begin VB.ComboBox cmbIdServicio2 
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
            Left            =   7440
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1290
            Width           =   3630
         End
         Begin VB.ComboBox cmbIdEspecialidad2 
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
            Left            =   7440
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   900
            Width           =   3630
         End
         Begin VB.ComboBox cmbIdDepartamento2 
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
            Left            =   7440
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   495
            Width           =   3630
         End
         Begin VB.ComboBox cmbIdServicio1 
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
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1305
            Width           =   3630
         End
         Begin VB.ComboBox cmbIdEspecialidad1 
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
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   900
            Width           =   3630
         End
         Begin VB.ComboBox cmbIdDepartamento1 
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
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   495
            Width           =   3630
         End
         Begin VB.Label lblTitulo2 
            AutoSize        =   -1  'True
            Caption         =   "Servicio2"
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
            Left            =   5805
            TabIndex        =   22
            Top             =   180
            Width           =   720
         End
         Begin VB.Label lblServicio2 
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
            Left            =   6165
            TabIndex        =   21
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label lblEspecialidad2 
            Caption         =   "Especialidad"
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
            Left            =   6165
            TabIndex        =   20
            Top             =   945
            Width           =   1395
         End
         Begin VB.Label lblDpto2 
            Caption         =   "Departamento"
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
            Left            =   6165
            TabIndex        =   19
            Top             =   555
            Width           =   1260
         End
         Begin VB.Label lblTitulo1 
            AutoSize        =   -1  'True
            Caption         =   "Servicio1"
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
            Left            =   105
            TabIndex        =   15
            Top             =   195
            Width           =   720
         End
         Begin VB.Label lblServicio1 
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
            Left            =   435
            TabIndex        =   14
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label lblEspecialidad1 
            Caption         =   "Especialidad"
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
            Left            =   435
            TabIndex        =   13
            Top             =   945
            Width           =   1395
         End
         Begin VB.Label lblDpto1 
            Caption         =   "Departamento"
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
            Left            =   435
            TabIndex        =   12
            Top             =   555
            Width           =   1260
         End
      End
      Begin VB.ComboBox cmbTipoRep 
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
         ItemData        =   "HEgresosHosp.frx":1C05
         Left            =   1230
         List            =   "HEgresosHosp.frx":1C15
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1230
         Width           =   5850
      End
      Begin VB.ComboBox cmbConsiderar 
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
         ItemData        =   "HEgresosHosp.frx":1C7F
         Left            =   4260
         List            =   "HEgresosHosp.frx":1C89
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   195
         Width           =   2865
      End
      Begin VB.ComboBox cmbAnio 
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
         ItemData        =   "HEgresosHosp.frx":1CC1
         Left            =   1245
         List            =   "HEgresosHosp.frx":1CC3
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Distr.Proced"
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
         Left            =   90
         TabIndex        =   33
         Top             =   750
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Rep"
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
         Left            =   90
         TabIndex        =   7
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Considerar"
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
         Left            =   3315
         TabIndex        =   5
         Top             =   285
         Width           =   840
      End
      Begin VB.Label Departamento 
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
         Left            =   90
         TabIndex        =   1
         Top             =   285
         Width           =   330
      End
   End
End
Attribute VB_Name = "HEgresosHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Egresos Hospitalarios
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbIdDepartamento1 As New sighEntidades.ListaDespleglable
Dim mo_cmbIdServicio1 As New sighEntidades.ListaDespleglable
Dim mo_cmbIdEspecialidad1 As New sighEntidades.ListaDespleglable
Dim mo_cmbIdDepartamento2 As New sighEntidades.ListaDespleglable
Dim mo_cmbIdServicio2 As New sighEntidades.ListaDespleglable
Dim mo_cmbIdEspecialidad2 As New sighEntidades.ListaDespleglable
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_cmbIdDpto As New sighEntidades.ListaDespleglable
Dim mo_cmbIdProv As New sighEntidades.ListaDespleglable
Dim mo_cmbIdDist As New sighEntidades.ListaDespleglable
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim ml_Titulo As String
Dim ml_TextoDelFiltro As String

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub


Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRptEgresosHosp As New RptHEgresosHosp
        oRptEgresosHosp.Anio = Val(cmbAnio.Text)
        oRptEgresosHosp.FechaAltaMedica = IIf(cmbConsiderar.ListIndex = 0, True, False)
        oRptEgresosHosp.TipoReporte = cmbTipoRep.ListIndex
        oRptEgresosHosp.IdDistrito = IIf(cmbDistrito.ListIndex = 0, 0, mo_cmbIdDist.BoundText)
        oRptEgresosHosp.idDepartamento1 = IIf(mo_cmbIdDepartamento1.BoundText = "", 0, mo_cmbIdDepartamento1.BoundText)
        oRptEgresosHosp.idEspecialidad1 = IIf(mo_cmbIdEspecialidad1.BoundText = "", 0, mo_cmbIdEspecialidad1.BoundText)
        oRptEgresosHosp.idServicio1 = IIf(mo_cmbIdServicio1.BoundText = "", 0, mo_cmbIdServicio1.BoundText)
        oRptEgresosHosp.idDepartamento2 = IIf(mo_cmbIdDepartamento2.BoundText = "", 0, mo_cmbIdDepartamento2.BoundText)
        oRptEgresosHosp.idEspecialidad2 = IIf(mo_cmbIdEspecialidad2.BoundText = "", 0, mo_cmbIdEspecialidad2.BoundText)
        oRptEgresosHosp.idServicio2 = IIf(mo_cmbIdServicio2.BoundText = "", 0, mo_cmbIdServicio2.BoundText)
        oRptEgresosHosp.Titulo = ml_Titulo
        oRptEgresosHosp.TextoDelFiltro = ml_TextoDelFiltro
        oRptEgresosHosp.CrearReporte Me.hwnd
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    Dim sMensaje As String
    sMensaje = ""
    ml_TextoDelFiltro = "FILTROS:   Año: " & cmbAnio.Text & ",   se consideró: " & cmbConsiderar.Text
    Select Case cmbDistrito.ListIndex
    Case 1
        If cmbIdDist.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Distrito de Procedencia" + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Distrito de Procedencia: (" & cmbIdDpto.Text & ")/(" & cmbIdProv.Text & ")/(" & cmbIdDist.Text
    End Select
    Select Case cmbTipoRep.ListIndex
    Case 0
        ml_Titulo = "EGRESOS HOSPITALARIOS"
    Case 1
        ml_Titulo = "EGRESOS HOSPITALARIOS CONSOLIDANDO POR DEPARTAMENTOS"
    Case 2    'por un Servicio
        ml_Titulo = "EGRESOS HOSPITALARIOS POR ESPECIALIDAD"
        If cmbIdDepartamento1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento" + Chr(13)
        End If
        If cmbIdEspecialidad1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija la Especialidad" + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Departamento: " & cmbIdDepartamento1.Text & "     Especialidad: " & cmbIdEspecialidad1.Text
    Case 3    'por 2 Especialidades
        ml_Titulo = "EGRESOS HOSPITALARIOS CONSOLIDANDO DOS SERVICIOS"
        If cmbIdDepartamento1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento (para el primer Servicio)" + Chr(13)
        End If
        If cmbIdEspecialidad1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija la Especialidad (para el primer Servicio)" + Chr(13)
        End If
        If cmbIdServicio1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Servicio (para el primer Servicio)" + Chr(13)
        End If
        If cmbIdDepartamento2.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento (para el segundo Servicio)" + Chr(13)
        End If
        If cmbIdEspecialidad2.Text = "" Then
           sMensaje = sMensaje + "Por favor elija la Especialidad (para el segundo Servicio)" + Chr(13)
        End If
        If cmbIdServicio2.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Servicio (para el segundo Servicio)" + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Servicio1: (" & cmbIdDepartamento1.Text & ")/(" & cmbIdEspecialidad1.Text & ")/(" & cmbIdServicio1.Text & "),     Servicio2: (" & cmbIdDepartamento2.Text & ")/(" & cmbIdEspecialidad2.Text & ")/(" & cmbIdServicio2.Text & ")"
    End Select
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function





Private Sub cmbDistrito_Change()
    Select Case cmbDistrito.ListIndex
    Case 0
        frmDistrito.Visible = False
    Case 1
        frmDistrito.Visible = True
    End Select

End Sub

Private Sub cmbDistrito_Click()
   cmbDistrito_Change
End Sub

Private Sub cmbIdDepartamento1_Click()
       Dim sMensaje As String
       mo_cmbIdEspecialidad1.BoundColumn = "IdEspecialidad"
       mo_cmbIdEspecialidad1.ListField = "DescripcionLarga"
       Set mo_cmbIdEspecialidad1.RowSource = mo_AdminServiciosHosp.EspecialidadesSeleccionarporDepartamento(Val(mo_cmbIdDepartamento1.BoundText))
       mo_cmbIdEspecialidad1.BoundText = ""
       If mo_AdminServiciosHosp.MensajeError <> "" Then
          MsgBox mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
       End If
End Sub


Private Sub cmbIdDepartamento1_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento1
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdDepartamento2_Click()
       Dim sMensaje As String
       mo_cmbIdEspecialidad2.BoundColumn = "IdEspecialidad"
       mo_cmbIdEspecialidad2.ListField = "DescripcionLarga"
       Set mo_cmbIdEspecialidad2.RowSource = mo_AdminServiciosHosp.EspecialidadesSeleccionarporDepartamento(Val(mo_cmbIdDepartamento2.BoundText))
       mo_cmbIdEspecialidad2.BoundText = ""
       If mo_AdminServiciosHosp.MensajeError <> "" Then
          MsgBox mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
       End If
End Sub

Private Sub cmbIdDepartamento2_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento2
    AdministrarKeyPreview KeyCode

End Sub




Private Sub cmbIdDpto_Click()
       If cmbIdDpto.ListIndex = -1 Then Exit Sub
       mo_cmbIdProv.BoundColumn = "IdProvincia"
       mo_cmbIdProv.ListField = "Nombre"
       Set mo_cmbIdProv.RowSource = mo_AdminServiciosGeograficos.ProvinciasSeleccionarPorDepartamento(Val(cmbIdDpto.ItemData(cmbIdDpto.ListIndex)))
       mo_cmbIdProv.BoundText = ""
       mo_cmbIdDist.BoundText = ""

End Sub

Private Sub cmbIdEspecialidad1_Click()
    mo_cmbIdServicio1.BoundColumn = "IdServicio"
    mo_cmbIdServicio1.ListField = "DescripcionLarga"
    Set mo_cmbIdServicio1.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarPorTipoServicioDptoEspecialidad(3, Val(mo_cmbIdDepartamento1.BoundText), Val(mo_cmbIdEspecialidad1.BoundText))
End Sub



Private Sub cmbIdEspecialidad2_Click()
    mo_cmbIdServicio2.BoundColumn = "IdServicio"
    mo_cmbIdServicio2.ListField = "DescripcionLarga"
    Set mo_cmbIdServicio2.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarPorTipoServicioDptoEspecialidad(3, Val(mo_cmbIdDepartamento2.BoundText), Val(mo_cmbIdEspecialidad2.BoundText))
End Sub


Private Sub cmbIdProv_Click()
       If cmbIdProv.ListIndex = -1 Then Exit Sub
       mo_cmbIdDist.BoundColumn = "IdDistrito"
       mo_cmbIdDist.ListField = "Nombre"
       Set mo_cmbIdDist.RowSource = mo_AdminServiciosGeograficos.DistritoSeleccionarPorProvincia(Val(cmbIdProv.ItemData(cmbIdProv.ListIndex)))
       If mo_AdminServiciosGeograficos.MensajeError <> "" Then
            MsgBox mo_AdminServiciosGeograficos.MensajeError, vbInformation, "Datos de paciente"
       End If
       mo_cmbIdDist.BoundText = ""
End Sub

Private Sub cmbTipoRep_Change()
    Select Case cmbTipoRep.ListIndex
    Case 0, 1
        lblTitulo1.Visible = False
        lblDpto1.Visible = False
        lblServicio1.Visible = False
        lblEspecialidad1.Visible = False
        cmbIdDepartamento1.Visible = False
        cmbIdServicio1.Visible = False
        cmbIdEspecialidad1.Visible = False
        lblTitulo2.Visible = False
        lblDpto2.Visible = False
        lblServicio2.Visible = False
        lblEspecialidad2.Visible = False
        cmbIdDepartamento2.Visible = False
        cmbIdServicio2.Visible = False
        cmbIdEspecialidad2.Visible = False
    Case 2
        lblTitulo1.Visible = True
        lblDpto1.Visible = True
        lblServicio1.Visible = False
        lblEspecialidad1.Visible = True
        cmbIdDepartamento1.Visible = True
        cmbIdServicio1.Visible = False
        cmbIdEspecialidad1.Visible = True
        lblTitulo2.Visible = False
        lblDpto2.Visible = False
        lblServicio2.Visible = False
        lblEspecialidad2.Visible = False
        cmbIdDepartamento2.Visible = False
        cmbIdServicio2.Visible = False
        cmbIdEspecialidad2.Visible = False
        lblTitulo1.Caption = "Elegir el Servicio:"
    Case 3
        lblTitulo1.Visible = True
        lblDpto1.Visible = True
        lblServicio1.Visible = True
        lblEspecialidad1.Visible = True
        cmbIdDepartamento1.Visible = True
        cmbIdServicio1.Visible = True
        cmbIdEspecialidad1.Visible = True
        lblTitulo2.Visible = True
        lblDpto2.Visible = True
        lblServicio2.Visible = True
        lblEspecialidad2.Visible = True
        cmbIdDepartamento2.Visible = True
        cmbIdServicio2.Visible = True
        cmbIdEspecialidad2.Visible = True
        lblTitulo1.Caption = "Elegir la primera Especialidad:"
        lblTitulo2.Caption = "Elegir la segunda Especialidad:"
    End Select
End Sub

Private Sub cmbTipoRep_Click()
    cmbTipoRep_Change
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdDepartamento1.MiComboBox = cmbIdDepartamento1
    Set mo_cmbIdEspecialidad1.MiComboBox = cmbIdEspecialidad1
    Set mo_cmbIdServicio1.MiComboBox = cmbIdServicio1
    Set mo_cmbIdDepartamento2.MiComboBox = cmbIdDepartamento2
    Set mo_cmbIdEspecialidad2.MiComboBox = cmbIdEspecialidad2
    Set mo_cmbIdServicio2.MiComboBox = cmbIdServicio2
    Set mo_cmbIdDpto.MiComboBox = cmbIdDpto
    Set mo_cmbIdProv.MiComboBox = cmbIdProv
    Set mo_cmbIdDist.MiComboBox = cmbIdDist

End Sub

Private Sub Form_Load()
    mo_Formulario.LlenaComboConAnios cmbAnio
    cmbConsiderar.ListIndex = 0
    cmbTipoRep.ListIndex = 0
    cmbDistrito.ListIndex = 0
    CargaCombos
End Sub

Sub CargaCombos()
       mo_cmbIdDepartamento1.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento1.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento1.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()
       
       mo_cmbIdDepartamento2.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento2.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento2.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()

       mo_cmbIdDpto.BoundColumn = "IdDepartamento"
       mo_cmbIdDpto.ListField = "DescripcionLarga"
       Set mo_cmbIdDpto.RowSource = mo_AdminServiciosGeograficos.DepartamentosSeleccionarTodos()
 End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub


