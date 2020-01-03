VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form RpParteDiario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ".."
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   5
      Top             =   5700
      Width           =   7035
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "RpParteDiario.frx":0000
         DownPicture     =   "RpParteDiario.frx":0460
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
         Left            =   1950
         Picture         =   "RpParteDiario.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "RpParteDiario.frx":0D4A
         DownPicture     =   "RpParteDiario.frx":120E
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
         Left            =   3555
         Picture         =   "RpParteDiario.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5610
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   7065
      Begin VB.CheckBox chkSoloCredito 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo CREDITOS"
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
         Left            =   5100
         TabIndex        =   25
         Top             =   180
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame frmProrrateo 
         Height          =   1725
         Left            =   120
         TabIndex        =   20
         Top             =   2940
         Visible         =   0   'False
         Width           =   6885
         Begin Threed.SSOption optSinProrrateo 
            Height          =   345
            Left            =   90
            TabIndex        =   21
            Top             =   930
            Width           =   6375
            _ExtentX        =   11245
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
            Caption         =   "Columna Exoneraciones según Servicio Social"
         End
         Begin Threed.SSOption optConProrrateo 
            Height          =   345
            Left            =   90
            TabIndex        =   22
            Top             =   1290
            Visible         =   0   'False
            Width           =   6375
            _ExtentX        =   11245
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
            Caption         =   "Columna Exoneraciones prorrateado (NO recomendado)"
         End
         Begin Threed.SSOption optSinProrrateoOSDesagregado 
            Height          =   345
            Left            =   90
            TabIndex        =   23
            Top             =   510
            Width           =   6375
            _ExtentX        =   11245
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
            Caption         =   "Columna Exoneraciones según Servicio Social (OS-Desag)"
         End
         Begin Threed.SSOption optPaOsDesag 
            Height          =   345
            Left            =   90
            TabIndex        =   24
            Top             =   150
            Width           =   6675
            _ExtentX        =   11774
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
            Caption         =   "Columna Exoneraciones según Servicio Social (PA,OS-Desag) (recomendado)"
            Value           =   -1
         End
      End
      Begin VB.ComboBox cmbIdTipoFinanciamiento 
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
         Left            =   1590
         TabIndex        =   16
         Top             =   2610
         Visible         =   0   'False
         Width           =   5100
      End
      Begin VB.ComboBox cmbIdTipoComprobante 
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
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2190
         Visible         =   0   'False
         Width           =   5100
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
         Left            =   1605
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1785
         Width           =   5085
      End
      Begin VB.ComboBox cmbIdTurno 
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
         Left            =   1605
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1380
         Width           =   5085
      End
      Begin VB.ComboBox cmbIdCaja 
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
         Left            =   1605
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   5085
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1605
         TabIndex        =   1
         Top             =   165
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   1605
         TabIndex        =   2
         Top             =   555
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin SISGalenPlus.XP_ProgressBar progressRpt 
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   4800
         Visible         =   0   'False
         Width           =   6570
         _ExtentX        =   11589
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
      Begin SISGalenPlus.XP_ProgressBar ProgressRpt1 
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   5220
         Visible         =   0   'False
         Width           =   6570
         _ExtentX        =   11589
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
      Begin VB.Label lblTipoFinanciamiento 
         AutoSize        =   -1  'True
         Caption         =   "Producto/Plan"
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
         TabIndex        =   17
         Top             =   2670
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblTipoDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento"
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
         TabIndex        =   15
         Top             =   2250
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cajero"
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
         TabIndex        =   13
         Top             =   1830
         Width           =   510
      End
      Begin VB.Label Label4 
         Caption         =   "Caja"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1020
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Turno"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Emisión Doc. Fin"
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
         TabIndex        =   4
         Top             =   615
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F.Emisión Doc.Ini"
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
         TabIndex        =   3
         Top             =   210
         Width           =   1380
      End
   End
End
Attribute VB_Name = "RpParteDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de Consolidado de Servicios, Farmcia
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_cmbIdCaja As New ListaDespleglable
Dim mo_cmbIdTurno As New ListaDespleglable
Dim mo_cmbIdResponsable As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdTipoComprobante As New ListaDespleglable
Dim mo_cmbIdTipoFinanciamiento As New SIGHEntidades.ListaDespleglable
Dim ml_IdTipoReporte As Long
Dim ml_IdUsuario As Long
Property Let IdUsuario(lIdValue As Long)
    ml_IdUsuario = lIdValue
End Property
Property Let IdTipoReporte(lIdValue As Long)
    ml_IdTipoReporte = lIdValue
End Property

Private Sub btnAceptar_Click()
    If txtFechaInicio.Text = SIGHEntidades.FECHA_VACIA_DMY_HM Then
        MsgBox "Por favor ingrese la fecha de inicio", vbInformation, Me.Caption
        Exit Sub
    End If
    If txtFechaFin.Text = SIGHEntidades.FECHA_VACIA_DMY_HM Then
        MsgBox "Por favor ingrese la fecha final", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Select Case ml_IdTipoReporte
    Case 1   'Parte Diario
        If cmbIdCaja.Text = "" Then
            MsgBox "Por favor elija la Caja", vbInformation, Me.Caption
            Exit Sub
        End If
        If cmbIdTurno.Text = "" Then
            MsgBox "Por favor elija el Turno", vbInformation, Me.Caption
            Exit Sub
        End If
        If cmbIdResponsable.Text = "" Then
            MsgBox "Por favor elija el Cajero", vbInformation, Me.Caption
            Exit Sub
        End If
        Me.MousePointer = 11
        Dim oRptCaja As New RptCaja
        oRptCaja.FechaInicio = txtFechaInicio.Text
        oRptCaja.FechaFin = txtFechaFin.Text
        oRptCaja.IdGestionCaja = Val(mo_cmbIdCaja.BoundText)
        oRptCaja.IdTurno = Val(mo_cmbIdTurno.BoundText)
        oRptCaja.IdCajero = Val(mo_cmbIdResponsable.BoundText)
        oRptCaja.TextoDelFiltro = "Caja: " & cmbIdCaja.Text & "     Cajero: " & cmbIdResponsable.Text & "     Turno: " & cmbIdTurno.Text & "     F.Documento: (" & txtFechaInicio.Text & " - " & txtFechaFin.Text & ")"
       ' oRptCaja.CrearParteDiario
    Case 2   'Consolidado Servicio
        Me.MousePointer = 11
        Dim oRptConsServicio As New RptCaja
        oRptConsServicio.FechaInicio = txtFechaInicio.Text
        oRptConsServicio.FechaFin = txtFechaFin.Text
        oRptConsServicio.IdGestionCaja = Val(mo_cmbIdCaja.BoundText)
        oRptConsServicio.IdTurno = Val(mo_cmbIdTurno.BoundText)
        oRptConsServicio.IdCajero = Val(mo_cmbIdResponsable.BoundText)
        oRptConsServicio.IdComprobantePago = Val(mo_cmbIdTipoComprobante.BoundText)
        oRptConsServicio.TextoDelFiltro = "(Servicio/Farmacia)  Caja: " & cmbIdCaja.Text & "     Cajero: " & cmbIdResponsable.Text & "     Turno: " & cmbIdTurno.Text & "     F.Documento: (" & txtFechaInicio.Text & " - " & txtFechaFin.Text & ")     " & IIf(Val(mo_cmbIdTipoComprobante.BoundText) = 0, "", "Tipo de Comprobante: " & cmbIdTipoComprobante.Text) & _
                                          IIf(Val(mo_cmbIdTipoFinanciamiento.BoundText) > 0, "  (Producto/Plan: " & Trim(cmbIdTipoFinanciamiento.Text) & ")", "") & IIf(Me.chkSoloCredito.Value = 1, "  (solo CREDITOS)", "  (sin considerar CREDITOS)")
        oRptConsServicio.idTipoFinanciamiento = Val(mo_cmbIdTipoFinanciamiento.BoundText)
        oRptConsServicio.lnHWnd = Me.hwnd
        oRptConsServicio.CrearReporteConsolidadoServicios IIf(chkSoloCredito.Value = 1, True, False)
    Case 3   'Consolidado Farmacia
        Me.MousePointer = 11
        Dim oRptConsFarmacia As New RptCaja
        oRptConsFarmacia.FechaInicio = txtFechaInicio.Text
        oRptConsFarmacia.FechaFin = txtFechaFin.Text
        oRptConsFarmacia.IdGestionCaja = Val(mo_cmbIdCaja.BoundText)
        oRptConsFarmacia.IdTurno = Val(mo_cmbIdTurno.BoundText)
        oRptConsFarmacia.IdCajero = Val(mo_cmbIdResponsable.BoundText)
        oRptConsFarmacia.TextoDelFiltro = "Caja: " & cmbIdCaja.Text & "     Cajero: " & cmbIdResponsable.Text & "     Turno: " & cmbIdTurno.Text & "     F.Documento: (" & txtFechaInicio.Text & " - " & txtFechaFin.Text & ")"
        oRptConsFarmacia.lnHWnd = Me.hwnd
        oRptConsFarmacia.CrearReporteConsolidadoFarmacia
    Case 4   'Consolidado Recaudacion
        Me.MousePointer = 11
        Dim oRptConsCaja As New RptCaja
        oRptConsCaja.FechaInicio = txtFechaInicio.Text
        oRptConsCaja.FechaFin = txtFechaFin.Text
        oRptConsCaja.IdGestionCaja = 0
        oRptConsCaja.IdTurno = 0
        oRptConsCaja.IdCajero = 0
        oRptConsCaja.TextoDelFiltro = "F.Apertura: (" & txtFechaInicio.Text & " - " & txtFechaFin.Text & ")"
        'oRptConsCaja.CrearConsolidadoRecaudacion
    Case 5 'Resumen por Centro de Costos
        If cmbIdCaja.Text <> "" Then
            If cmbIdTurno.Text = "" Then
                MsgBox "Por favor elija el Turno", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        If cmbIdResponsable.Text <> "" Then
            If cmbIdCaja.Text = "" Then
                MsgBox "Por favor elija la Caja", vbInformation, Me.Caption
                Exit Sub
            End If
            If cmbIdTurno.Text = "" Then
                MsgBox "Por favor elija el Turno", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        Me.MousePointer = 11
        Dim oRptClaseCry As New rCrystal
        oRptClaseCry.DestinoReporte = sghPantalla
        oRptClaseCry.FechaInicio = CDate(txtFechaInicio.Text)
        oRptClaseCry.FechaFin = CDate(txtFechaFin.Text)
        oRptClaseCry.IdCaja = Val(mo_cmbIdCaja.BoundText)
        oRptClaseCry.IdTurno = Val(mo_cmbIdTurno.BoundText)
        oRptClaseCry.IdCajero = Val(mo_cmbIdResponsable.BoundText)
        oRptClaseCry.TextoDelFiltro = "Caja: " & cmbIdCaja.Text & "     Cajero: " & cmbIdResponsable.Text & "     Turno: " & cmbIdTurno.Text & "     F.Documento: (" & txtFechaInicio.Text & " - " & txtFechaFin.Text & ")" & IIf(Me.chkSoloCredito.Value = 1, "  (solo CREDITOS)", "  (sin considerar CREDITOS)")
        oRptClaseCry.TipoReporte = "ResumenCCosto"
        Set oRptClaseCry.progressRpt = Me.progressRpt
        Set oRptClaseCry.ProgressRpt1 = Me.ProgressRpt1
        oRptClaseCry.ConProrrateoColExoneracion = IIf(optConProrrateo.Value = True, True, False)
        oRptClaseCry.ConOtrosSaludDesagregado = IIf(Me.optSinProrrateoOSDesagregado.Value = True, True, False)
        oRptClaseCry.DetallaProcAdmyOtrosServ = IIf(optPaOsDesag.Value = True, True, False)
        oRptClaseCry.TieneCredito = IIf(chkSoloCredito.Value = 1, True, False)
        oRptClaseCry.Show vbModal
        Set oRptClaseCry = Nothing
    End Select
    Me.MousePointer = 1
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub


Private Sub Form_Initialize()
    Set mo_cmbIdCaja.MiComboBox = cmbIdCaja
    Set mo_cmbIdTurno.MiComboBox = cmbIdTurno
    Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable
    Set mo_cmbIdTipoComprobante.MiComboBox = cmbIdTipoComprobante
    Set mo_cmbIdTipoFinanciamiento.MiComboBox = cmbIdTipoFinanciamiento
End Sub

Private Sub Form_Load()
    Me.txtFechaInicio.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY) & " 00:01"
    Me.txtFechaFin.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY) & " 23:59"
    Select Case ml_IdTipoReporte
    Case 1   'Parte Diario
         Me.Caption = "Reporte de Parte Diario"
    Case 2   'Consolidado Servicio
         Me.Caption = "Reporte Consolidado de Servicio y Farmacia"
         lblTipoDocumento.Visible = True
         cmbIdTipoComprobante.Visible = True
         lblTipoFinanciamiento.Visible = True
         cmbIdTipoFinanciamiento.Visible = True
         chkSoloCredito.Visible = True
    Case 3   'Consolidado Farmacia
         Me.Caption = "Reporte Consolidado de Farmacia"
    Case 4   'Consolidado de Recaudacion
         Me.Caption = "Reporte Consolidado de Recaudación"
         cmbIdCaja.Visible = False
         cmbIdTurno.Visible = False
         cmbIdResponsable.Visible = False
         Label1.Visible = False
         Label8.Visible = False
         Label4.Visible = False
         Me.txtFechaInicio.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual & " 00:01"
    Case 5   'Resumen por centro de costos
         Me.Caption = "Resumen por Centro de Costo"
         Me.progressRpt.Visible = True
         Me.ProgressRpt1.Visible = True
         frmProrrateo.Visible = True
         chkSoloCredito.Visible = True
    End Select
   
    
    mo_cmbIdTurno.BoundColumn = "IdTurno"
    mo_cmbIdTurno.ListField = "Descripcion"
    Set mo_cmbIdTurno.RowSource = mo_AdminCaja.TurnosSeleccionarTodosParaLista()
        
    mo_cmbIdCaja.BoundColumn = "IdCaja"
    mo_cmbIdCaja.ListField = "Descripcion"
    Set mo_cmbIdCaja.RowSource = mo_AdminCaja.CajaSeleccionarTodosParaLista()

    mo_cmbIdResponsable.BoundColumn = "IdEmpleado"
    mo_cmbIdResponsable.ListField = "DCajero"
    Set mo_cmbIdResponsable.RowSource = mo_AdminCaja.CajerosSeleccionarTodos()
    
    mo_cmbIdTipoComprobante.BoundColumn = "IdTipoComprobante"
    mo_cmbIdTipoComprobante.ListField = "Descripcion"
    Set mo_cmbIdTipoComprobante.RowSource = mo_AdminCaja.TiposComprobanteSeleccionarTodos()
    
    mo_cmbIdTipoFinanciamiento.BoundColumn = "IdTipoFinanciamiento"
    mo_cmbIdTipoFinanciamiento.ListField = "Descripcion"
    Set mo_cmbIdTipoFinanciamiento.RowSource = mo_ReglasFacturacion.TiposFinanciamientosSeleccionarPorGeneraPagos(sghTodosLosQuePaganEnCaja)
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub







Private Sub txtFechaFin_Change()
If Not IsDate(txtFechaInicio.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaInicio.Text = SIGHEntidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If
End Sub


Private Sub txtFechaInicio_LostFocus()
If Not IsDate(txtFechaInicio.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaInicio.Text = SIGHEntidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If
End Sub
