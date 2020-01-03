VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ReportesEgresosHosp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de egresos  (Epicrisis)"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3555
      Left            =   60
      TabIndex        =   9
      Top             =   30
      Width           =   5370
      Begin VB.ComboBox cmbConsiderar 
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
         ItemData        =   "HospitalizacionReportes.frx":0000
         Left            =   1680
         List            =   "HospitalizacionReportes.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   225
         Width           =   3570
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1380
         Width           =   3570
      End
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   990
         Width           =   3570
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   3570
      End
      Begin GalenHos.XP_ProgressBar progressRpt 
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   6956042
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   2160
         Width           =   1395
         _ExtentX        =   2461
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
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   2520
         Width           =   1395
         _ExtentX        =   2461
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
      Begin MSDataListLib.DataCombo cmbFuenteFinanciamiento 
         Height          =   330
         Left            =   1680
         TabIndex        =   17
         Top             =   1770
         Width           =   3555
         _ExtentX        =   6271
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fte.Financiam/IAFA"
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
         TabIndex        =   18
         Top             =   1830
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Especialidad"
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
         TabIndex        =   16
         Top             =   285
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Serv. egreso"
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
         Left            =   105
         TabIndex        =   14
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label8 
         Caption         =   "Esp. egreso"
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
         Left            =   105
         TabIndex        =   13
         Top             =   1035
         Width           =   1395
      End
      Begin VB.Label Departamento 
         Caption         =   "Dpto egreso"
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
         Left            =   105
         TabIndex        =   12
         Top             =   660
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F.Egreso Med. Ini"
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
         TabIndex        =   11
         Top             =   2175
         Width           =   1425
      End
      Begin VB.Label Label3 
         Caption         =   "F.Egreso Med.Fin"
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
         Left            =   105
         TabIndex        =   10
         Top             =   2550
         Width           =   1410
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   8
      Top             =   3645
      Width           =   5370
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HospitalizacionReportes.frx":0043
         DownPicture     =   "HospitalizacionReportes.frx":0507
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
         Picture         =   "HospitalizacionReportes.frx":09F3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HospitalizacionReportes.frx":0EDF
         DownPicture     =   "HospitalizacionReportes.frx":133F
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
         Picture         =   "HospitalizacionReportes.frx":17B4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Width           =   1365
      End
   End
End
Attribute VB_Name = "ReportesEgresosHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************daniel barrantes**************
'***************se considera tambien EMERGENCIA
'***************
Option Explicit
Dim mo_cmbIdDepartamento As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdServicio As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdEspecialidad As New SIGHEntidades.ListaDespleglable
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_IdTipoReporte As Long
Dim lcFiltro As String
Dim mo_AdminServHosp As New ReglasServiciosHosp
Dim oRsFuentesFinanciamiento As New Recordset

Property Let IdTipoReporte(lIdValue As Long)
    ml_IdTipoReporte = lIdValue
End Property

Private Sub btnAceptar_Click()

    
    If txtFechaInicio.Text = SIGHEntidades.FECHA_VACIA_DMY Then
        MsgBox "Por favor ingrese la fecha de inicio", vbInformation, Me.Caption
        Exit Sub
    End If

    If txtFechaFin.Text = SIGHEntidades.FECHA_VACIA_DMY Then
        MsgBox "Por favor ingrese la fecha final", vbInformation, Me.Caption
        Exit Sub
    End If
    lcFiltro = "Filtros:  " & IIf(cmbConsiderar.ListIndex = 2, "F.Cita: (", "F.Egreso Médico: (") & txtFechaInicio.Text & " - " & txtFechaFin.Text & ") " & _
             "     (" & cmbConsiderar.Text & ")     " & _
             IIf(cmbIdDepartamento.Text = "", "", "     Departamento: " & cmbIdDepartamento.Text) & _
             IIf(cmbIdEspecialidad.Text = "", "", "     Especialidad: " & cmbIdEspecialidad.Text) & _
             IIf(cmbIdServicio.Text = "", "", "     Servicio: " & cmbIdServicio.Text)

    Select Case ml_IdTipoReporte
    Case sghReporteEgresosHospitalario
        Dim ldFechaIF As Date
        Me.MousePointer = 11
        Dim oRptEgresosHosp As New RptEgresosHosp
        oRptEgresosHosp.IdDepartamento = Val(mo_cmbIdDepartamento.BoundText)
        oRptEgresosHosp.IdEspecialidad = Val(mo_cmbIdEspecialidad.BoundText)
        oRptEgresosHosp.idServicio = Val(mo_cmbIdServicio.BoundText)
        ldFechaIF = CDate(Me.txtFechaFin.Text)
        oRptEgresosHosp.FechaFin = ldFechaIF
        oRptEgresosHosp.FechaInicio = Me.txtFechaInicio.Text
        oRptEgresosHosp.IdTipoEspecialidad = IIf(cmbConsiderar.ListIndex = 0, 3, IIf(cmbConsiderar.ListIndex = 2, 1, 2))
        Set oRptEgresosHosp.progressRpt = Me.progressRpt
        oRptEgresosHosp.TextoDelFiltro = lcFiltro + IIf(Val(cmbFuenteFinanciamiento.BoundText) > 0, "  (IAFA: " & Trim(cmbFuenteFinanciamiento.Text) & ")", "")
        oRptEgresosHosp.IdPlan = Val(cmbFuenteFinanciamiento.BoundText)
        oRptEgresosHosp.CrearReporteEgresosHospitalariosII
        Me.MousePointer = 1
    Case 2
    End Select

End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub cmbIdDepartamento_Click()
Dim sMensaje As String

       mo_cmbIdEspecialidad.BoundColumn = "IdEspecialidad"
       mo_cmbIdEspecialidad.ListField = "DescripcionLarga"
       Set mo_cmbIdEspecialidad.RowSource = mo_AdminServiciosHosp.EspecialidadesSeleccionarporDepartamento(Val(mo_cmbIdDepartamento.BoundText))
       
       mo_cmbIdEspecialidad.BoundText = ""
       
       If mo_AdminServiciosHosp.MensajeError <> "" Then
        MsgBox mo_AdminServiciosHosp.MensajeError, vbCritical, Me.Caption
       End If
End Sub

Private Sub cmbIdDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdEspecialidad_Click()
    mo_cmbIdServicio.BoundColumn = "IdServicio"
    mo_cmbIdServicio.ListField = "DescripcionLarga"
    If cmbConsiderar.ListIndex = 0 Then
       Set mo_cmbIdServicio.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarPorTipoServicioDptoEspecialidad(3, Val(mo_cmbIdDepartamento.BoundText), Val(mo_cmbIdEspecialidad.BoundText))
    Else
       If cmbConsiderar.ListIndex = 2 Then
          Set mo_cmbIdServicio.RowSource = mo_AdminServHosp.ServiciosSeleccionarPorTipoV2(1, sghFiltraAnuladosYactivos)
       Else
          Set mo_cmbIdServicio.RowSource = mo_ReglasComunes.ServiciosSeleccionarEmergenciaPorEspecialidad(Val(mo_cmbIdEspecialidad.BoundText))
       End If
    End If

End Sub

Private Sub Form_Activate()
       Select Case cmbConsiderar.ListIndex
       Case 0  'Hospitalizacion
       Case 1  'Emergencia
            Me.Caption = "Reporte de Egresos"
       Case 2  'CE
            Label2.Caption = "F.Cita Inicio:"
            Label3.Caption = "F.Cita Final:"
            Me.Caption = "Citados y/o atendidos x Consultorios"
       End Select
End Sub

Private Sub Form_Initialize()

    Set mo_cmbIdDepartamento.MiComboBox = cmbIdDepartamento
    Set mo_cmbIdEspecialidad.MiComboBox = cmbIdEspecialidad
    Set mo_cmbIdServicio.MiComboBox = cmbIdServicio
    
    Me.txtFechaInicio.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual()
    Me.txtFechaFin = SIGHEntidades.UltimaFechaDDMMYYDelMesActual()
    
End Sub

Private Sub Form_Load()
       

       Me.txtFechaInicio.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual()
       Me.txtFechaFin.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    
       mo_cmbIdDepartamento.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()

       Set oRsFuentesFinanciamiento = mo_ReglasComunes.FuentesFinanciamientoSegunFiltro("")
       Set cmbFuenteFinanciamiento.RowSource = oRsFuentesFinanciamiento
       cmbFuenteFinanciamiento.ListField = "Descripcion"
       cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"

       cmbConsiderar.ListIndex = 0
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

