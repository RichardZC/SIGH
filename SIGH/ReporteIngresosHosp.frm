VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ReporteIngresosHosp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de ingresos "
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   60
      TabIndex        =   18
      Top             =   4350
      Width           =   5535
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ReporteIngresosHosp.frx":0000
         DownPicture     =   "ReporteIngresosHosp.frx":0460
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
         Left            =   1410
         Picture         =   "ReporteIngresosHosp.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ReporteIngresosHosp.frx":0D4A
         DownPicture     =   "ReporteIngresosHosp.frx":120E
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
         Left            =   2940
         Picture         =   "ReporteIngresosHosp.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4305
      Left            =   30
      TabIndex        =   11
      Top             =   0
      Width           =   5565
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
         ItemData        =   "ReporteIngresosHosp.frx":1BE6
         Left            =   1695
         List            =   "ReporteIngresosHosp.frx":1BF0
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   225
         Width           =   3660
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipos de Número de Historia Clínica"
         Height          =   1140
         Left            =   120
         TabIndex        =   19
         Top             =   2580
         Width           =   5370
         Begin VB.ComboBox cmbTipoHistoria 
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
            ItemData        =   "ReporteIngresosHosp.frx":1C11
            Left            =   1545
            List            =   "ReporteIngresosHosp.frx":1C1E
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   230
            Width           =   3690
         End
         Begin VB.ComboBox cmbIdTipoGenHistoriaClinica 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1545
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   630
            Visible         =   0   'False
            Width           =   3690
         End
         Begin GalenHos.XP_ProgressBar XP_ProgressBar1 
            Height          =   300
            Left            =   135
            TabIndex        =   20
            Top             =   2280
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
         Begin VB.Label Label6 
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
            Left            =   150
            TabIndex        =   21
            Top             =   285
            Width           =   840
         End
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
         TabIndex        =   1
         Top             =   645
         Width           =   3675
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
         TabIndex        =   2
         Top             =   1035
         Width           =   3675
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
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1425
         Width           =   3690
      End
      Begin GalenHos.XP_ProgressBar progressRpt 
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   3870
         Width           =   5340
         _ExtentX        =   9419
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
         TabIndex        =   4
         Top             =   1815
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
         Left            =   3930
         TabIndex        =   5
         Top             =   1830
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
         Left            =   1665
         TabIndex        =   6
         Top             =   2190
         Width           =   3675
         _ExtentX        =   6482
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
         Left            =   150
         TabIndex        =   23
         Top             =   2250
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
         Left            =   135
         TabIndex        =   22
         Top             =   285
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "hasta"
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
         Left            =   3360
         TabIndex        =   17
         Top             =   1860
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Ingreso Ini."
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
         Left            =   135
         TabIndex        =   16
         Top             =   1860
         Width           =   1560
      End
      Begin VB.Label Departamento 
         Caption         =   "Dpto ingreso"
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
         Left            =   135
         TabIndex        =   15
         Top             =   705
         Width           =   1260
      End
      Begin VB.Label Label8 
         Caption         =   "Esp. ingreso"
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
         Left            =   135
         TabIndex        =   14
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Serv. ingreso"
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
         Left            =   135
         TabIndex        =   13
         Top             =   1485
         Width           =   1275
      End
   End
End
Attribute VB_Name = "ReporteIngresosHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************daniel barrantes**************
'***************se Considera tambien EMERGENCIA
'***************filtros: tipos Numero Historia
Dim mo_cmbIdDepartamento As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdServicio As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdEspecialidad As New SIGHEntidades.ListaDespleglable

Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_IdTipoReporte As Long
Dim mo_cmbIdTipoGenHistoriaClinica As New SIGHEntidades.ListaDespleglable
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim lcFiltro As String
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
    lcFiltro = "Filtros:  F.Ingreso: (" & txtFechaInicio.Text & " - " & txtFechaFin.Text & ") " & _
             "     (" & cmbConsiderar.Text & ")     " & _
             IIf(cmbIdDepartamento.Text = "", "", "     Departamento: " & cmbIdDepartamento.Text) & _
             IIf(cmbIdEspecialidad.Text = "", "", "     Especialidad: " & cmbIdEspecialidad.Text) & _
             IIf(cmbIdServicio.Text = "", "", "     Servicio: " & cmbIdServicio.Text)

    Select Case ml_IdTipoReporte
    Case sghReporteIngresosHospitalario
        Dim oRptIngresosHosp As New RptIngresosHosp
        Me.MousePointer = 11
        oRptIngresosHosp.IdDepartamento = Val(mo_cmbIdDepartamento.BoundText)
        oRptIngresosHosp.IdEspecialidad = Val(mo_cmbIdEspecialidad.BoundText)
        oRptIngresosHosp.idServicio = Val(mo_cmbIdServicio.BoundText)
        oRptIngresosHosp.FechaFin = Me.txtFechaFin.Text
        oRptIngresosHosp.FechaInicio = Me.txtFechaInicio.Text
        Set oRptIngresosHosp.progressRpt = Me.progressRpt
        oRptIngresosHosp.IdTipoNroHistoria = IIf(cmbTipoHistoria.ListIndex = 2, mo_cmbIdTipoGenHistoriaClinica.BoundText, IIf(cmbTipoHistoria.ListIndex = 0, 100, 200))
        oRptIngresosHosp.IdTipoEspecialidad = IIf(cmbConsiderar.ListIndex = 0, 3, 2)
        oRptIngresosHosp.TextoDelFiltro = lcFiltro
        oRptIngresosHosp.TextoDelFiltro = lcFiltro + IIf(Val(cmbFuenteFinanciamiento.BoundText) > 0, "  (IAFA: " & Trim(cmbFuenteFinanciamiento.Text) & ")", "")
        oRptIngresosHosp.IdPlan = Val(cmbFuenteFinanciamiento.BoundText)
        oRptIngresosHosp.CrearReporteIngresosHospitalarios
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
       Set mo_cmbIdServicio.RowSource = mo_ReglasComunes.ServiciosSeleccionarEmergenciaPorEspecialidad(Val(mo_cmbIdEspecialidad.BoundText))
    End If

End Sub


Private Sub cmbTipoHistoria_Change()
    cmbTipoHistoria_Click
End Sub

Private Sub cmbTipoHistoria_Click()
   If cmbTipoHistoria.ListIndex = 2 Then
      cmbIdTipoGenHistoriaClinica.Visible = True
   Else
      cmbIdTipoGenHistoriaClinica.Visible = False
   End If

End Sub

Private Sub Form_Initialize()

    Set mo_cmbIdDepartamento.MiComboBox = cmbIdDepartamento
    Set mo_cmbIdEspecialidad.MiComboBox = cmbIdEspecialidad
    Set mo_cmbIdServicio.MiComboBox = cmbIdServicio
    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica

    Me.txtFechaInicio.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual()
    Me.txtFechaFin = SIGHEntidades.UltimaFechaDDMMYYDelMesActual()
    
End Sub

Private Sub Form_Load()
       

        Me.txtFechaInicio.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual()
        Me.txtFechaFin.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
        cmbTipoHistoria.ListIndex = 0
    
        mo_cmbIdDepartamento.BoundColumn = "IdDepartamento"
        mo_cmbIdDepartamento.ListField = "DescripcionLarga"
        Set mo_cmbIdDepartamento.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()
        
        mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
        mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()
        mo_cmbIdTipoGenHistoriaClinica.BoundText = 2
        
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



