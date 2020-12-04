VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form EConsumoXPtoCarga 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consumo por Punto de Carga"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   1
      Top             =   4590
      Width           =   5790
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "EConsumoXptoCarga.frx":0000
         DownPicture     =   "EConsumoXptoCarga.frx":0460
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
         Left            =   1470
         Picture         =   "EConsumoXptoCarga.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "EConsumoXptoCarga.frx":0D4A
         DownPicture     =   "EConsumoXptoCarga.frx":120E
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
         Left            =   2963
         Picture         =   "EConsumoXptoCarga.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4380
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   5790
      Begin VB.CheckBox chkConDetalleItems 
         Caption         =   "Con detalle de ITEMS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   23
         Top             =   3090
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2400
      End
      Begin SIGHReportes.XP_ProgressBar XP_ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   3765
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   556
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
      Begin VB.CheckBox chkCtasSinAlta 
         Caption         =   "Agregar Cuentas que aún no tienen FECHA DE EGRESO MEDICO "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   21
         Top             =   2700
         Width           =   5625
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   90
         TabIndex        =   18
         Top             =   1020
         Width           =   5475
         Begin Threed.SSOption optPlan 
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   210
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
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
            Caption         =   "Consumo x Punto de Carga x Cuenta"
            Value           =   -1
         End
         Begin Threed.SSOption optConsumoYpago 
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Top             =   570
            Width           =   4245
            _ExtentX        =   7488
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
            Caption         =   "Consumo y Pagos por cada Cuenta"
         End
      End
      Begin VB.CheckBox chkUsaResumen 
         Caption         =   "Usa tabla RESUMEN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   17
         Top             =   2370
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   4680
         TabIndex        =   15
         Top             =   1980
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CheckBox chkExcel 
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
         Left            =   90
         Picture         =   "EConsumoXptoCarga.frx":1BE6
         TabIndex        =   9
         Top             =   2070
         Width           =   1605
      End
      Begin VB.ComboBox cmbTipoFinanciamiento 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2850
         TabIndex        =   8
         Top             =   2010
         Visible         =   0   'False
         Width           =   480
      End
      Begin MSMask.MaskEdBox txtFecha1 
         Height          =   315
         Left            =   2010
         TabIndex        =   4
         Top             =   660
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
      Begin MSMask.MaskEdBox txtFecha2 
         Height          =   315
         Left            =   4140
         TabIndex        =   5
         Top             =   660
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
         Left            =   2010
         TabIndex        =   10
         Top             =   240
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
      Begin Threed.SSOption optTipoFinanciamiento 
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
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
         Caption         =   "Por Tipo Financiamiento"
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   315
         Left            =   4680
         TabIndex        =   16
         Top             =   2370
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CE"
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
         Left            =   3600
         TabIndex        =   14
         Top             =   2370
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hosp/Emer"
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
         Left            =   3600
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fuente Financiamiento"
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
         TabIndex        =   12
         Top             =   270
         Width           =   1875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "F.Egreso Médico"
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
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "al"
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
         Left            =   3900
         TabIndex        =   6
         Top             =   690
         Width           =   120
      End
   End
End
Attribute VB_Name = "EConsumoXPtoCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: reporte de Consumo por Punto de Carga
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim sMensaje As String
Dim mo_Teclado As New sighEntidades.Teclado
Dim ml_idUsuarioConPermisoEnSISoEXOoSOAT As Long
Dim ml_idUsuario As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_cmbTipoFinanciamiento As New sighEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim lcSubTitulo As String
Dim oRsFuentesFinanciamiento As New Recordset
Dim rsTmpSOAT As New Recordset
Dim mrs_Tmp As New Recordset
Private WithEvents oGeneraDatos As RptEConsumoXptoCarga
Attribute oGeneraDatos.VB_VarHelpID = -1


Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property



Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        'mo_ReglasArchivoClinico.ActualizaDatosConProblemas
        lcSubTitulo = "Fecha Egreso Médico: " & txtFecha1.Text & " hasta " & txtFecha2.Text & "     " & IIf(optPlan = True, "Fuente Financiamiento: " & cmbFuenteFinanciamiento.Text, "") & IIf(chkUsaResumen.Value = 1, " (TR)", "") & IIf(chkCtasSinAlta.Value = 1, " (con CTAS Sin AM)", "")
       'Dim oGeneraDatos As New RptEConsumoXptoCarga
        If chkUsaResumen.Value = 1 Then
           oGeneraDatos.ProcesaDatosYllenaTmpRapidamente Val(mo_cmbTipoFinanciamiento.BoundText), CDate(txtFecha1.Text), CDate(txtFecha2.Text), Val(cmbFuenteFinanciamiento.BoundText), IIf(Me.optPlan.Value = True, False, True), IIf(chkCtasSinAlta.Value = 1, True, False)
        Else
           oGeneraDatos.ProcesaDatosYllenaTmp Val(mo_cmbTipoFinanciamiento.BoundText), CDate(txtFecha1.Text), CDate(txtFecha2.Text), Val(cmbFuenteFinanciamiento.BoundText), IIf(Me.optPlan.Value = True, False, True), IIf(chkCtasSinAlta.Value = 1, True, False)
        End If
        If optPlan.Value = True Then
            Set mrs_Tmp = oGeneraDatos.Devuelve_mrs_Tmp
            If mrs_Tmp.RecordCount = 0 Then
                MsgBox "No existe información con esos datos", vbInformation, Me.Caption
            Else
                Dim oRptClaseCry As New rCrystal
                oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
                oRptClaseCry.FechaInicio = Format(txtFecha1.Text, sighEntidades.DevuelveFechaSoloFormato_DMY)
                oRptClaseCry.FechaFin = Format(txtFecha2.Text, sighEntidades.DevuelveFechaSoloFormato_DMY)
                oRptClaseCry.IdResponsable = Val(mo_cmbTipoFinanciamiento.BoundText)
                oRptClaseCry.IdPlan = Val(cmbFuenteFinanciamiento.BoundText)
                oRptClaseCry.TextoDelFiltro = lcSubTitulo
                oRptClaseCry.TipoReporte = Me.Name
                oRptClaseCry.EnResumen = IIf(Me.optPlan.Value = True, False, True)
                Set oRptClaseCry.RecordSet_mrs_Tmp = mrs_Tmp
                oRptClaseCry.Show vbModal
                Set oRptClaseCry = Nothing
            End If
        ElseIf optConsumoYpago.Value = True Then
            Set rsTmpSOAT = oGeneraDatos.Devuelve_rsTmpSOAT
            If rsTmpSOAT.RecordCount = 0 Then
                MsgBox "No existe información con esos datos", vbInformation, Me.Caption
            Else
'                Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
'                mo_ReglasReportes.ExportarRecordSetAexcel rsTmpSOAT, Me.optConsumoYpago.Caption, lcSubTitulo, "", Me.hwnd
                 CrearReporte_excel Me.hwnd, rsTmpSOAT
            End If
        End If
        Me.MousePointer = 1
    End If
End Sub
Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    If optPlan = True Then
      If Val(cmbFuenteFinanciamiento.BoundText) = 0 Then
         sMensaje = "Por favor elija el IAFA"
      End If
    ElseIf Me.optTipoFinanciamiento.Value = True Then
      If Val(mo_cmbTipoFinanciamiento.BoundText) = 0 Then
         sMensaje = "Por favor elija el TIPO DE FINANCIAMIENTO"
      End If
    End If
    
    If Me.txtFecha1 = sighEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de egreso medico inicial"
    Else
        If Not sighEntidades.EsFecha(Me.txtFecha1, "DD/MM/AAAA") Then
            sMensaje = "La fecha de egreso medico inicial no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFecha2 = sighEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de egreso medico final"
    Else
        If Not sighEntidades.EsFecha(Me.txtFecha2, "DD/MM/AAAA") Then
            sMensaje = "La fecha de egreso medico final no tiene el formato correcto"
        End If
    End If
    If CDate(Me.txtFecha1.Text) > CDate(Me.txtFecha2.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
       Exit Function
    End If
    
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function


Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub


Private Sub Form_Initialize()
      Set mo_cmbTipoFinanciamiento.MiComboBox = cmbTipoFinanciamiento
End Sub

Private Sub Form_Load()
       '
       Set oGeneraDatos = New RptEConsumoXptoCarga
       XP_ProgressBar1.ShowText = True
       '
       Me.txtFecha1.Text = sighEntidades.PrimerFechaDDMMYYDelMesActual()
       Me.txtFecha2.Text = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
       '
       Dim lbBuscaPermisoEnFacturacion As New SIGHNegocios.ReglasFacturacion
       mo_cmbTipoFinanciamiento.BoundColumn = "idTipoFinanciamiento"
       mo_cmbTipoFinanciamiento.ListField = "Descripcion"
       Set mo_cmbTipoFinanciamiento.RowSource = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia(" and esOficina=1")
       Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
       Dim oRsDondeLabora As Recordset
       Set oRsDondeLabora = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghSeguros, ml_idUsuario)
       If oRsDondeLabora.RecordCount > 0 Then
          ml_idUsuarioConPermisoEnSISoEXOoSOAT = oRsDondeLabora.Fields!idLaboraSubArea
       End If
       Set oRsDondeLabora = Nothing
       Set oBuscaDondeLabora = Nothing
       '
       
       Set oRsFuentesFinanciamiento = mo_ReglasFacturacion.FuentesFinanciamientoSeleccionarTodos
       Set cmbFuenteFinanciamiento.RowSource = oRsFuentesFinanciamiento
       cmbFuenteFinanciamiento.ListField = "Descripcion"
       cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"
       
       
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub


Private Sub oGeneraDatos_ProgressActualizaValor(lnValorActual As Long, lnValorTotal As Long)
    XP_ProgressBar1.Max = lnValorTotal
    XP_ProgressBar1.Min = 0
    XP_ProgressBar1.Value = lnValorActual
    DoEvents
    Me.Refresh
End Sub

Private Sub optConsumoYpago_Click(Value As Integer)
    If optConsumoYpago.Value = True Then
       chkExcel.Enabled = False
       chkExcel.Value = 1
       Me.chkConDetalleItems.Visible = True
       'Me.cmbFuenteFinanciamiento.Text = ""
       'Me.cmbFuenteFinanciamiento.Enabled = False
    End If
End Sub

Private Sub optPlan_Click(Value As Integer)
    If optPlan.Value = True Then
       chkExcel.Enabled = True
       chkExcel.Value = 0
       Me.cmbFuenteFinanciamiento.Enabled = True
       Me.chkConDetalleItems.Visible = False
    End If
End Sub

Private Sub optTipoFinanciamiento_Click(Value As Integer)
   If optTipoFinanciamiento.Value = True Then
       cmbTipoFinanciamiento.Visible = True
       cmbFuenteFinanciamiento.Visible = False
   End If
End Sub

Private Sub txtFecha1_LostFocus()
    If txtFecha1 <> sighEntidades.FECHA_VACIA_DMY Then
        If Not sighEntidades.EsFecha(txtFecha1, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFecha1 = sighEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFecha2_LostFocus()
    If txtFecha2 <> sighEntidades.FECHA_VACIA_DMY Then
        If Not sighEntidades.EsFecha(txtFecha2, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFecha2 = sighEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub


Sub CrearReporte_excel(lnHwnd As Long, rsReporte As Recordset)
Dim oRsDetalle As New Recordset
Dim iFila As Long
Dim lnNumHistorias As Long
Dim lnIdPaciente As Long
Dim lcPaciente As String
Dim lnNumTotal As Long
Dim mo_ReporteUtil As New ReporteUtil
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim oConexionExterna As New Connection
Dim oConexion As New Connection
Dim lnLineas As Integer: Dim lnIdAtencion  As Long
Dim lcProcedencia As String
Dim lnGastosConsulta As Double: Dim lnTotGastosConsulta As Double
Dim lnGastosServicios As Double: Dim lnTotGastosServicios As Double
Dim lnGastosFarmacia As Double: Dim lnTotGastosFarmacia As Double
Dim lcCie10 As String: Dim lcDx As String
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String

oConexionExterna.CommandTimeout = 900
oConexionExterna.CursorLocation = adUseClient
oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)

sighEntidades.AbreConexionSIGH oConexion

lbEsOpenOffice = False
On Error GoTo ManejadorErrorExcel

    If lbEsOpenOffice = True Then
        Dim ServiceManager As Object
        Dim Desktop As Object
        Dim Document As Object
        Dim Feuille As Object
        Dim Plage As Object
        Dim args()
        Dim Chemin As String
        Dim Fichier As String
        Dim lcArchivoExcel As String
        Dim PrintArea(0)
        Dim Style As Object
        Dim Border As Object
        'encabezado
        Dim PageStyles As Object
        Dim Sheet As Object
        Dim StyleFamilies As Object
        Dim DefPage As Object
        Dim Htext As Object
        Dim Hcontent As Object
        Dim ret As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
    
    'Filtra los Datos
    lnNumHistorias = rsReporte.RecordCount
    If lnNumHistorias = 0 Then
            MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
            If lbEsOpenOffice = True Then
                'Abre el archivo ExcelOpenOffice
                lcArchivoExcel = App.Path + "\Plantillas\ceDxPacientes.ods"
'                FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'                Chemin = "file:///" & App.Path & "\Plantillas\"
'                Chemin = Replace(Chemin, "\", "/")
'                Fichier = Chemin & "/OpenOffice.ods"
                '
                Fichier = Format(Time, "hhmmss") & ".ods"
                FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
                lcArchivoExcel = Fichier
                Chemin = "file:///" & App.Path & "\Plantillas\"
                Chemin = Replace(Chemin, "\", "/")
                Fichier = Chemin & "/" & lcArchivoExcel
                '

                Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
                Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
                Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
                Set Feuille = Document.getSheets().getByIndex(0)
                'Encabezado de Pagina
                'mo_CabeceraReportes.CabeceraReportes Document, True
                ' Pone la ventana en primer plano, pasándole el Hwnd
                ret = SetForegroundWindow(lnHwnd)
            Else
                Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                'Crea nueva hoja
                Set oWorkBook = oExcel.Workbooks.Add
                'Abre, copia y cierra la plantilla
               Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
               oWorkBookPlantilla.Worksheets("Hoja_libre").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
                mo_reglasComunes.CabeceraReportes oWorkSheet, False
            End If
            If lbEsOpenOffice = True Then
            Else
                oWorkSheet.Cells(3, 1).Value = UCase(optConsumoYpago.Caption)
                oWorkSheet.Cells(4, 1).Value = lcSubTitulo
                oWorkSheet.Cells(5, 1).Value = "Cuenta"
                oWorkSheet.Cells(5, 2).Value = "Paciente"
                oWorkSheet.Cells(5, 8).Value = "Nro Historia"
                oWorkSheet.Cells(5, 9).Value = "Origen"
                oWorkSheet.Cells(5, 10).Value = "F.Alta"
                oWorkSheet.Cells(5, 11).Value = "H.Alta"
                oWorkSheet.Cells(5, 12).Value = "Financiamiento"
                oWorkSheet.Cells(5, 13).Value = "TotalFacturado"
                oWorkSheet.Cells(5, 14).Value = "FUA"
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, 5, 1, 5, 14
            End If
            iFila = 6
            lnNumTotal = 0
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                    If lbEsOpenOffice = True Then
                    Else
                       oWorkSheet.Cells(iFila, 1).Value = rsReporte!idCuentaAtencion
                       oWorkSheet.Cells(iFila, 2).Value = rsReporte!Paciente
                       oWorkSheet.Cells(iFila, 8).Value = rsReporte!nroHistoria
                       oWorkSheet.Cells(iFila, 9).Value = rsReporte!origen
                       oWorkSheet.Cells(iFila, 10).Value = rsReporte!Falta
                       oWorkSheet.Cells(iFila, 11).Value = rsReporte!Halta
                       oWorkSheet.Cells(iFila, 12).Value = rsReporte!dFinanciamiento
                       oWorkSheet.Cells(iFila, 13).Value = rsReporte!tFacturado
                       oWorkSheet.Cells(iFila, 14).Value = rsReporte!fua
                       '
                       If chkConDetalleItems.Value = 1 Then
                           If oRsDetalle.State = 1 Then oRsDetalle.Close
    '                       If rsReporte!fua <> "" Then
    '                          Set oRsDetalle = mo_ReglasFacturacion.FuaConsumoDeItemsPorCuenta(rsReporte!idCuentaAtencion, oConexionExterna)
    '                          If oRsDetalle.RecordCount = 0 Then
    '                             If oRsDetalle.State = 1 Then oRsDetalle.Close
    '                             Set oRsDetalle = mo_ReglasFacturacion.FuaConsumoDeItemsXCuentaSigh(rsReporte!idCuentaAtencion, oConexion)
    '                          End If
    '                       Else
                              Set oRsDetalle = mo_ReglasFacturacion.FuaConsumoDeItemsXCuentaSigh(rsReporte!idCuentaAtencion, rsReporte!idTipoFinanciamiento, oConexion)
    '                       End If
                           If oRsDetalle.RecordCount > 0 Then
                              iFila = iFila + 1
                              oWorkSheet.Cells(iFila, 3).Value = ""
                              oWorkSheet.Cells(iFila, 4).Value = "Código"
                              oWorkSheet.Cells(iFila, 5).Value = "Descripción"
                              oWorkSheet.Cells(iFila, 10).Value = "Cantidad"
                              oWorkSheet.Cells(iFila, 11).Value = "Precio"
                              oWorkSheet.Cells(iFila, 12).Value = "Importe"
                              mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, 12
                              oRsDetalle.MoveFirst
                              iFila = iFila + 1
                              Do While Not oRsDetalle.EOF
                                   oWorkSheet.Cells(iFila, 3).Value = oRsDetalle!PuntoCarga
                                   oWorkSheet.Cells(iFila, 4).Value = oRsDetalle!Codigo
                                   oWorkSheet.Cells(iFila, 5).Value = oRsDetalle!descripcion
                                   oWorkSheet.Cells(iFila, 10).Value = oRsDetalle!Cantidad
                                   oWorkSheet.Cells(iFila, 11).Value = oRsDetalle!precio
                                   oWorkSheet.Cells(iFila, 12).Value = Round(oRsDetalle!Cantidad * oRsDetalle!precio, 2)
                                   oRsDetalle.MoveNext
                                   iFila = iFila + 1
                              Loop
                              iFila = iFila + 1
                           End If
                       End If
                       '
                    End If
                    rsReporte.MoveNext
                    iFila = iFila + 1
            Loop
            If lbEsOpenOffice = True Then
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 1, iFila, 14
                oWorkSheet.Cells(iFila, 2).Value = "Nro Cuentas: " & Trim(Str(lnNumHistorias))
                iFila = iFila + 1
            End If
            
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 5
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                If oWorkSheet.PageSetup.PrintArea <> "" Then
                   oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                End If
                oExcel.Visible = True
                oWorkSheet.PrintPreview
            End If
                If lbEsOpenOffice = True Then
                'Liberar Memoria
                Set Plage = Nothing
                Set Feuille = Nothing
                Set Document = Nothing
                Set Desktop = Nothing
                Set ServiceManager = Nothing
                Set Style = Nothing
                Set Border = Nothing
                'encabezado de pagina
                Set PageStyles = Nothing
                Set Sheet = Nothing
                Set StyleFamilies = Nothing
                Set DefPage = Nothing
                Set Htext = Nothing
                Set Hcontent = Nothing
            Else
            'Liberar memoria
                Set oExcel = Nothing
                Set oWorkBookPlantilla = Nothing
                Set oWorkBook = Nothing
                Set oWorkSheet = Nothing
            End If
    End If
    oConexionExterna.Close
    oConexion.Close
    
    Set mo_AdminAdmision = Nothing
    Set mo_reglasComunes = Nothing
    Set oConexionExterna = Nothing
    Set oConexion = Nothing
    Set mo_ReporteUtil = Nothing
Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
    Resume
End Sub

