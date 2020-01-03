VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ProgMedicaReporte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de programación médica"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "ProgMedicaReporte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   105
      TabIndex        =   10
      Top             =   2295
      Width           =   5370
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ProgMedicaReporte.frx":0CCA
         DownPicture     =   "ProgMedicaReporte.frx":112A
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
         Left            =   1320
         Picture         =   "ProgMedicaReporte.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ProgMedicaReporte.frx":1A14
         DownPicture     =   "ProgMedicaReporte.frx":1ED8
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
         Left            =   2850
         Picture         =   "ProgMedicaReporte.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2250
      Left            =   105
      TabIndex        =   6
      Top             =   30
      Width           =   5370
      Begin VB.ComboBox cmbIdEspecialidad 
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
         Left            =   1500
         TabIndex        =   12
         Text            =   "cmbIdEspecialidad"
         Top             =   660
         Width           =   3705
      End
      Begin VB.CheckBox chkMostrarHoras 
         Caption         =   "Mostrar horas"
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
         Left            =   1500
         TabIndex        =   3
         Top             =   1455
         Width           =   2805
      End
      Begin VB.PictureBox progressRpt 
         Height          =   300
         Left            =   150
         ScaleHeight     =   240
         ScaleWidth      =   4950
         TabIndex        =   11
         Top             =   1830
         Visible         =   0   'False
         Width           =   5010
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.ComboBox cmbIdDepartamento 
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
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   255
         Width           =   3705
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   3915
         TabIndex        =   2
         Top             =   1065
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.Label Label8 
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
         Left            =   150
         TabIndex        =   13
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Fin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3030
         TabIndex        =   9
         Top             =   1095
         Width           =   810
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Inicio"
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
         Left            =   150
         TabIndex        =   8
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label Departamento 
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
         Left            =   150
         TabIndex        =   7
         Top             =   285
         Width           =   1260
      End
   End
End
Attribute VB_Name = "ProgMedicaReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Programación Médica
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mo_cmbIdDepartamento As New sighentidades.ListaDespleglable
Dim mo_cmbIdEspecialidad As New sighentidades.ListaDespleglable
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
'Private WithEvents oRptEgresosHosp As clReportesEgreHosp
Dim mo_Teclado As New sighentidades.Teclado

Private Sub btnAceptar_Click()
Dim oRptProgMedica As New clProgramMedica

    If mo_cmbIdDepartamento.BoundText = "" Then
        MsgBox "Por favor ingrese el departamento", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Me.txtFechaInicio = sighentidades.FECHA_VACIA_DMY Then
        MsgBox "Ingrese la fecha de inicio", vbInformation, Me.Caption
        Exit Sub
    Else
        If Not sighentidades.EsFecha(Me.txtFechaInicio, "DD/MM/AAAA") Then
            MsgBox "La fecha de inicio, no tiene el formato correcto", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    
    If Me.txtFechaFin = sighentidades.FECHA_VACIA_DMY Then
        MsgBox "Ingrese la fecha final", vbInformation, Me.Caption
        Exit Sub
    Else
        If Not sighentidades.EsFecha(Me.txtFechaFin, "DD/MM/AAAA") Then
            MsgBox "La fecha final, no tiene el formato correcto", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    If CDate(Me.txtFechaInicio.Text) > CDate(Me.txtFechaFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
       Exit Sub
    End If
    'yamill Palomino
    oRptProgMedica.IdDepartamento = Val(mo_cmbIdDepartamento.BoundText)
    oRptProgMedica.IdEspecialidad = Val(mo_cmbIdEspecialidad.BoundText)
    oRptProgMedica.FechaFin = Me.txtFechaFin.Text
    oRptProgMedica.FechaInicio = Me.txtFechaInicio.Text
    'Set oRptProgMedica.progressRpt = Me.progressRpt
    
    oRptProgMedica.CrearReporteProgramacionMedica (Me.chkMostrarHoras.Value = 1), Me.hwnd

End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub chkMostrarHoras_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkMostrarHoras
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdDepartamento_Click()
Dim sMensaje As String

       mo_cmbIdEspecialidad.BoundColumn = "IdEspecialidad"
       mo_cmbIdEspecialidad.ListField = "DescripcionLarga"
       Set mo_cmbIdEspecialidad.RowSource = mo_AdminServiciosHosp.EspecialidadesSeleccionarporDepartamento(Val(mo_cmbIdDepartamento.BoundText))
       
       mo_cmbIdEspecialidad.BoundText = ""
       
       If mo_AdminServiciosHosp.MensajeError <> "" Then
        MsgBox mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
       End If
End Sub

Private Sub cmbIdDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento
    AdministrarKeyPreview KeyCode
End Sub



Private Sub Form_Initialize()
    Set mo_cmbIdDepartamento.MiComboBox = cmbIdDepartamento
    Set mo_cmbIdEspecialidad.MiComboBox = cmbIdEspecialidad
    Me.txtFechaInicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
    Me.txtFechaFin = sighentidades.UltimaFechaDDMMYYDelMesActual()
    
End Sub

Private Sub Form_Load()
       
       mo_cmbIdDepartamento.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()
       cmbIdEspecialidad.Text = ""
End Sub

Private Sub txtFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaFin
    AdministrarKeyPreview KeyCode
End Sub



Private Sub txtFechaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaInicio
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub txtFechaInicio_LostFocus()
    If txtFechaInicio <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFechaInicio, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaInicio = sighentidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFechaFin_LostFocus()
    If txtFechaFin <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFechaFin, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaFin = sighentidades.FECHA_VACIA_DMY
        End If
    End If
End Sub
