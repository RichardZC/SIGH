VERSION 5.00
Begin VB.Form HDiasPaciente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Días Paciente Hospitalario por Departamentos y/o Servicios"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   30
      TabIndex        =   2
      Top             =   4590
      Width           =   7230
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HDiasPaciente.frx":0000
         DownPicture     =   "HDiasPaciente.frx":04C4
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
         Left            =   3750
         Picture         =   "HDiasPaciente.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HDiasPaciente.frx":0E9C
         DownPicture     =   "HDiasPaciente.frx":12FC
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
         Left            =   2220
         Picture         =   "HDiasPaciente.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4500
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   7200
      Begin VB.Frame Frame2 
         Height          =   3405
         Left            =   1215
         TabIndex        =   8
         Top             =   960
         Width           =   5835
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   2925
            Width           =   3975
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2535
            Width           =   3975
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2130
            Width           =   3975
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
            Width           =   3975
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
            Width           =   3975
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
            Width           =   3975
         End
         Begin VB.Label lblTitulo2 
            AutoSize        =   -1  'True
            Caption         =   "Especialidad2"
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
            Left            =   45
            TabIndex        =   22
            Top             =   1815
            Width           =   1065
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
            Left            =   405
            TabIndex        =   21
            Top             =   3000
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
            Left            =   405
            TabIndex        =   20
            Top             =   2580
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
            Left            =   405
            TabIndex        =   19
            Top             =   2190
            Width           =   1260
         End
         Begin VB.Label lblTitulo1 
            AutoSize        =   -1  'True
            Caption         =   "Especialidad1"
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
            Width           =   1065
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
         ItemData        =   "HDiasPaciente.frx":1BE6
         Left            =   1230
         List            =   "HDiasPaciente.frx":1BF6
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
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
         ItemData        =   "HDiasPaciente.frx":1C60
         Left            =   4215
         List            =   "HDiasPaciente.frx":1C6A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
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
         ItemData        =   "HDiasPaciente.frx":1CA2
         Left            =   1245
         List            =   "HDiasPaciente.frx":1CA4
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   195
         Width           =   1215
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
         Top             =   690
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
         Top             =   255
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
Attribute VB_Name = "HDiasPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Días pacientes
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
        Dim oRptDiasPaciente As New RptHDiasPaciente
        oRptDiasPaciente.Anio = Val(cmbAnio.Text)
        oRptDiasPaciente.FechaAltaMedica = IIf(cmbConsiderar.ListIndex = 0, True, False)
        oRptDiasPaciente.TipoReporte = cmbTipoRep.ListIndex
        oRptDiasPaciente.idDepartamento1 = IIf(mo_cmbIdDepartamento1.BoundText = "", 0, mo_cmbIdDepartamento1.BoundText)
        oRptDiasPaciente.idEspecialidad1 = IIf(mo_cmbIdEspecialidad1.BoundText = "", 0, mo_cmbIdEspecialidad1.BoundText)
        oRptDiasPaciente.idServicio1 = IIf(mo_cmbIdServicio1.BoundText = "", 0, mo_cmbIdServicio1.BoundText)
        oRptDiasPaciente.idDepartamento2 = IIf(mo_cmbIdDepartamento2.BoundText = "", 0, mo_cmbIdDepartamento2.BoundText)
        oRptDiasPaciente.idEspecialidad2 = IIf(mo_cmbIdEspecialidad2.BoundText = "", 0, mo_cmbIdEspecialidad2.BoundText)
        oRptDiasPaciente.idServicio2 = IIf(mo_cmbIdServicio2.BoundText = "", 0, mo_cmbIdServicio2.BoundText)
        oRptDiasPaciente.Titulo = ml_Titulo
        oRptDiasPaciente.TextoDelFiltro = ml_TextoDelFiltro
        oRptDiasPaciente.CrearReporte Me.hwnd
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    Dim sMensaje As String
    sMensaje = ""
    ml_TextoDelFiltro = "(N°días FecHrAlta-FecHrIng, si HoraAlta>1pm es 1 día más) FILTROS:   Año: " & cmbAnio.Text & ",   se consideró: " & cmbConsiderar.Text
    Select Case cmbTipoRep.ListIndex
    Case 0
        ml_Titulo = "DIAS PACIENTE HOSPITALARIO"
    Case 1
        ml_Titulo = "DIAS PACIENTE HOSPITALARIO CONSOLIDANDO POR DEPARTAMENTOS"
    Case 2    'por un Servicio
        ml_Titulo = "DIAS PACIENTE HOSPITALARIO POR ESPECIALIDAD"
        If cmbIdDepartamento1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento" + Chr(13)
        End If
        If cmbIdEspecialidad1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija la Especialidad" + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Departamento: " & cmbIdDepartamento1.Text & "     Especialidad: " & cmbIdEspecialidad1.Text
    Case 3    'por 2 Especialidades
        ml_Titulo = "DIAS PACIENTE HOSPITALARIO CONSOLIDANDO DOS SERVICIOS"
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

End Sub

Private Sub Form_Load()
    mo_Formulario.LlenaComboConAnios cmbAnio
    cmbConsiderar.ListIndex = 0
    cmbTipoRep.ListIndex = 0
    CargaCombos
End Sub

Sub CargaCombos()
       mo_cmbIdDepartamento1.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento1.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento1.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()
       
       mo_cmbIdDepartamento2.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento2.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento2.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()

 End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub


