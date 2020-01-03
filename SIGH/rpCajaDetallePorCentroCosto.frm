VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form RpCajaDetalleCentroCosto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle por Centro de Costos"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "rpCajaDetallePorCentroCosto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3885
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   7545
      Begin VB.Frame frmProrrateo 
         Height          =   1965
         Left            =   90
         TabIndex        =   12
         Top             =   1770
         Width           =   7335
         Begin Threed.SSOption optSinProrrateo 
            Height          =   345
            Left            =   90
            TabIndex        =   13
            Top             =   1020
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
            TabIndex        =   14
            Top             =   1410
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
            TabIndex        =   15
            Top             =   630
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
            TabIndex        =   16
            Top             =   240
            Width           =   6795
            _ExtentX        =   11986
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
      Begin VB.CheckBox chkAgrupaCC 
         Caption         =   "Agrupa por Consultorio ?  (sólo para C.Costo=Hosp/Emer/CE)"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1470
         Width           =   5355
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
         Left            =   6210
         Picture         =   "rpCajaDetallePorCentroCosto.frx":0CCA
         TabIndex        =   10
         Top             =   1470
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.ComboBox cmbCentroCostos 
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
         TabIndex        =   4
         Top             =   960
         Width           =   3825
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1605
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   9
         Top             =   210
         Width           =   1380
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
         TabIndex        =   8
         Top             =   615
         Width           =   1470
      End
      Begin VB.Label Label4 
         Caption         =   "Centro de Costo"
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
         TabIndex        =   7
         Top             =   1020
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1080
      Left            =   30
      TabIndex        =   0
      Top             =   3990
      Width           =   7545
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rpCajaDetallePorCentroCosto.frx":0FDC
         DownPicture     =   "rpCajaDetallePorCentroCosto.frx":14A0
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
         Left            =   3878
         Picture         =   "rpCajaDetallePorCentroCosto.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rpCajaDetallePorCentroCosto.frx":1E78
         DownPicture     =   "rpCajaDetallePorCentroCosto.frx":22D8
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
         Left            =   2348
         Picture         =   "rpCajaDetallePorCentroCosto.frx":274D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "RpCajaDetalleCentroCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte por Centro de Costos
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_cmbCentroCostos As New ListaDespleglable
Dim ml_IdTipoReporte As Long
Dim ml_idUsuario As Long
Property Let idUsuario(lIdValue As Long)
    ml_idUsuario = lIdValue
End Property
Property Let IdTipoReporte(lIdValue As Long)
    ml_IdTipoReporte = lIdValue
End Property

Private Sub btnAceptar_Click()
    If txtFechaInicio.Text = sighEntidades.FECHA_VACIA_DMY_HM Then
        MsgBox "Por favor ingrese la fecha de inicio", vbInformation, Me.Caption
        Exit Sub
    End If
    If txtFechaFin.Text = sighEntidades.FECHA_VACIA_DMY_HM Then
        MsgBox "Por favor ingrese la fecha final", vbInformation, Me.Caption
        Exit Sub
    End If
    If cmbCentroCostos.Text = "" Then
        MsgBox "Por favor elija el Centro de Costos", vbInformation, Me.Caption
        Exit Sub
    End If
'    Me.MousePointer = 11
'    Dim oRptClaseCry As New RptCaja
'    oRptClaseCry.CentroCostoDetallado IIf(chkExcel.Value = 1, True, False), "DETALLE POR CENTRO DE COSTOS", "Centro de Costos: " & cmbCentroCostos.Text & "     F.Documento: (" & txtFechaInicio.Text & " - " & txtFechaFin.Text & ")", 0, CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text), 0, 0, Val(mo_cmbCentroCostos.BoundText)
'    Set oRptClaseCry = Nothing
'    Me.MousePointer = 1
    Me.MousePointer = 11
    Dim oRptClaseCry As New rCrystal
    oRptClaseCry.DestinoReporte = sghPantalla
    oRptClaseCry.FechaInicio = CDate(txtFechaInicio.Text)
    oRptClaseCry.FechaFin = CDate(txtFechaFin.Text)
    oRptClaseCry.TextoDelFiltro = "Centro de Costos: " & cmbCentroCostos.Text & "     F.Documento: (" & txtFechaInicio.Text & " - " & txtFechaFin.Text & ")"
    oRptClaseCry.TipoReporte = "DetalleCentroCosto"
    oRptClaseCry.EnArchivoExcel = IIf(Me.chkExcel.Value = 1, True, False)
    oRptClaseCry.idCentroCostos = mo_cmbCentroCostos.BoundText
    oRptClaseCry.TotalizarXconsultorio = IIf(chkAgrupaCC.Value = 1, True, False)
    oRptClaseCry.ConOtrosSaludDesagregado = IIf(Me.optSinProrrateoOSDesagregado.Value = True, True, False)
    oRptClaseCry.DetallaProcAdmyOtrosServ = IIf(optPaOsDesag.Value = True, True, False)
    oRptClaseCry.Show vbModal
    Set oRptClaseCry = Nothing
    Me.MousePointer = 1
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub


Private Sub Form_Initialize()
    Set mo_cmbCentroCostos.MiComboBox = cmbCentroCostos
End Sub

Private Sub Form_Load()
    Me.txtFechaInicio.Text = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY) & " 00:01"
    Me.txtFechaFin.Text = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY) & " 23:59"
   
    
    mo_cmbCentroCostos.BoundColumn = "IdCentroCosto"
    mo_cmbCentroCostos.ListField = "Descripcion"
    Set mo_cmbCentroCostos.RowSource = mo_reglasComunes.CentrosCostoSeleccionarTodos
        
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub



Private Sub txtFechaFin_LostFocus()
If Not IsDate(txtFechaFin.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaFin.Text = sighEntidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If

End Sub

Private Sub txtFechaInicio_LostFocus()
If Not IsDate(txtFechaInicio.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaInicio.Text = sighEntidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If
End Sub
