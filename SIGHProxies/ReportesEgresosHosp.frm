VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ReportesEgresosHosp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de egresos  (Epicrisis)"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4890
      Left            =   15
      TabIndex        =   8
      Top             =   -15
      Width           =   5370
      Begin VB.CheckBox chkSoloCitados 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo CITADOS y enviar MENSAJE TEXTO a CELULAR"
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
         Left            =   105
         TabIndex        =   25
         Top             =   4530
         Visible         =   0   'False
         Width           =   4605
      End
      Begin VB.TextBox txtNhistoria 
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
         Left            =   1650
         TabIndex        =   24
         Top             =   3555
         Width           =   1380
      End
      Begin SIGHProxies.XP_ProgressBar XP_ProgressBar1 
         Height          =   345
         Left            =   3255
         TabIndex        =   22
         Top             =   2775
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   609
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
         Color           =   12937777
      End
      Begin VB.CheckBox chkConsiderarAdmRepetidas 
         Alignment       =   1  'Right Justify
         Caption         =   "Considerar lista de ERRORES (Admisiones Repetidas)"
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
         Left            =   105
         TabIndex        =   21
         Top             =   4215
         Visible         =   0   'False
         Width           =   4605
      End
      Begin VB.CheckBox chkEpasaron24hr 
         Alignment       =   1  'Right Justify
         Caption         =   "Sólo los que pasaron de 24 hr de ESTANCIA"
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
         Left            =   90
         TabIndex        =   20
         Top             =   3885
         Visible         =   0   'False
         Width           =   4005
      End
      Begin VB.CheckBox chkConsiderarDx 
         Alignment       =   1  'Right Justify
         Caption         =   "Considerar Dx"
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
         Left            =   90
         TabIndex        =   19
         Top             =   3210
         Value           =   1  'Checked
         Width           =   1785
      End
      Begin VB.CheckBox chkConsiderarCPT 
         Alignment       =   1  'Right Justify
         Caption         =   "Considerar CPT"
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
         Left            =   90
         TabIndex        =   18
         Top             =   2880
         Value           =   1  'Checked
         Width           =   1785
      End
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
         ItemData        =   "ReportesEgresosHosp.frx":0000
         Left            =   1680
         List            =   "ReportesEgresosHosp.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   14
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
         TabIndex        =   16
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "N° Historia"
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
         Left            =   120
         TabIndex        =   23
         Top             =   3600
         Width           =   855
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
         TabIndex        =   17
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
         TabIndex        =   15
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   2550
         Width           =   1410
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   -30
      TabIndex        =   7
      Top             =   4830
      Width           =   5370
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ReportesEgresosHosp.frx":0043
         DownPicture     =   "ReportesEgresosHosp.frx":0507
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
         Picture         =   "ReportesEgresosHosp.frx":09F3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ReportesEgresosHosp.frx":0EDF
         DownPicture     =   "ReportesEgresosHosp.frx":133F
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
         Picture         =   "ReportesEgresosHosp.frx":17B4
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

'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Egresos Hospitalarios
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbIdDepartamento As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdServicio As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdEspecialidad As New SIGHEntidades.ListaDespleglable
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim lcFiltro As String
Dim mo_AdminServHosp As New ReglasServiciosHosp
Dim oRsFuentesFinanciamiento As New Recordset
Dim ml_IdTipoReporte As Long

Private WithEvents oRptEgresosHosp As clReportesEgreHosp
Attribute oRptEgresosHosp.VB_VarHelpID = -1

Property Let IdTipoReporte(lIdValue As Long)
    ml_IdTipoReporte = lIdValue
End Property

Private Sub btnAceptar_Click()

If wxFranklin = "*" Then Exit Sub


    If Me.txtFechaInicio = SIGHEntidades.FECHA_VACIA_DMY Then
        MsgBox "Ingrese la fecha de inicio", vbInformation, Me.Caption
        Exit Sub
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFechaInicio, "DD/MM/AAAA") Then
            MsgBox "La fecha de inicio, no tiene el formato correcto", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    
    If Me.txtFechaFin = SIGHEntidades.FECHA_VACIA_DMY Then
        MsgBox "Ingrese la fecha final", vbInformation, Me.Caption
        Exit Sub
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFechaFin, "DD/MM/AAAA") Then
            MsgBox "La fecha final, no tiene el formato correcto", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    If CDate(Me.txtFechaInicio.Text) > CDate(Me.txtFechaFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
       Exit Sub
    End If
    If CDate(Me.txtFechaInicio.Text) <= Date And Me.chkSoloCitados.Value = 1 Then
       MsgBox "La FECHA INICIAL debe ser mayor a HOY", vbInformation, "Reporte"
       Exit Sub
       
    End If
    
    lcFiltro = "Filtros:  " & IIf(cmbConsiderar.ListIndex = 2, "F.Cita: (", "F.Egreso Médico: (") & txtFechaInicio.Text & " - " & txtFechaFin.Text & ") " & _
             "     (" & cmbConsiderar.Text & ")     " & _
             IIf(cmbIdDepartamento.Text = "", "", "     Departamento: " & cmbIdDepartamento.Text) & _
             IIf(cmbIdEspecialidad.Text = "", "", "     Especialidad: " & cmbIdEspecialidad.Text) & _
             IIf(cmbIdServicio.Text = "", "", "     Servicio: " & cmbIdServicio.Text) & _
             IIf(chkEpasaron24hr.Value = 1, " (Pacientes que pasaron más 24 hr de ESTANCIA)", "") & _
             IIf(Me.txtNhistoria.Text <> "", "  (Historia: " & txtNhistoria.Text & ")", "")
             

    Select Case ml_IdTipoReporte
    Case sghReporteEgresosHospitalario
        Dim ldFechaIF As Date
        Me.MousePointer = 11
        
        
        oRptEgresosHosp.IdDepartamento = Val(mo_cmbIdDepartamento.BoundText)
        oRptEgresosHosp.IdEspecialidad = Val(mo_cmbIdEspecialidad.BoundText)
        oRptEgresosHosp.idServicio = Val(mo_cmbIdServicio.BoundText)
        ldFechaIF = CDate(Me.txtFechaFin.Text)
        oRptEgresosHosp.FechaFin = ldFechaIF
        oRptEgresosHosp.FechaInicio = Me.txtFechaInicio.Text
        oRptEgresosHosp.IdTipoEspecialidad = IIf(cmbConsiderar.ListIndex = 0, 3, IIf(cmbConsiderar.ListIndex = 2, 1, 2))
       
        oRptEgresosHosp.TextoDelFiltro = lcFiltro + IIf(Val(cmbFuenteFinanciamiento.BoundText) > 0, "  (IAFA: " & Trim(cmbFuenteFinanciamiento.Text) & ")", "") & IIf(Me.chkSoloCitados.Value = 1, " (Solo CITADOS)", "")
        oRptEgresosHosp.IdPlan = Val(cmbFuenteFinanciamiento.BoundText)
        oRptEgresosHosp.CrearReporteEgresosHospitalariosII IIf(Me.chkConsiderarDx.Value = 1, True, False), _
                                                           IIf(Me.chkConsiderarCPT.Value = 1, True, False), _
                                                           Me.hwnd, IIf(chkEpasaron24hr.Value = 1, True, False), _
                                                           IIf(chkConsiderarAdmRepetidas.Value = 1, True, False), _
                                                           txtNhistoria.Text, IIf(Me.chkSoloCitados.Value = 1, True, False)
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
        MsgBox mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
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
          Set mo_cmbIdServicio.RowSource = mo_reglasComunes.ServiciosSeleccionarEmergenciaPorEspecialidad(Val(mo_cmbIdEspecialidad.BoundText))
       End If
    End If

End Sub

Private Sub Form_Activate()
       Select Case cmbConsiderar.ListIndex
       Case 0  'Hospitalizacion
            chkConsiderarAdmRepetidas.Visible = True
       Case 1  'Emergencia
            Me.Caption = "Reporte de Egresos"
            chkEpasaron24hr.Visible = True
            chkConsiderarAdmRepetidas.Visible = True
       Case 2  'CE
            Label2.Caption = "F.Cita Inicio:"
            Label3.Caption = "F.Cita Final:"
            Me.Caption = "Citados y/o atendidos x Consultorios"
            chkSoloCitados.Visible = True
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
       '
       Set oRptEgresosHosp = New clReportesEgreHosp
       XP_ProgressBar1.ShowText = True
       '
       Me.txtFechaInicio.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual()
       Me.txtFechaFin.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    
       mo_cmbIdDepartamento.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()

       Set oRsFuentesFinanciamiento = mo_reglasComunes.FuentesFinanciamientoSegunFiltro("")
       Set cmbFuenteFinanciamiento.RowSource = oRsFuentesFinanciamiento
       cmbFuenteFinanciamiento.ListField = "Descripcion"
       cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"

       cmbConsiderar.ListIndex = 0
       
End Sub



Private Sub oRptEgresosHosp_ProgressActualizaValor(lnValorActual As Long, lnValorTotal As Long)
    XP_ProgressBar1.Max = lnValorTotal
    XP_ProgressBar1.Min = 0
    XP_ProgressBar1.Value = lnValorActual
    DoEvents
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
    If txtFechaInicio <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFechaInicio, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaInicio = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFechaFin_LostFocus()
    If txtFechaFin <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFechaFin, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaFin = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

