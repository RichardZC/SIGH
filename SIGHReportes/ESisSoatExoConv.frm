VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form ESisSoatExoConv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas para Liquidación"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   1
      Top             =   2700
      Width           =   5850
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ESisSoatExoConv.frx":0000
         DownPicture     =   "ESisSoatExoConv.frx":0460
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
         Left            =   1448
         Picture         =   "ESisSoatExoConv.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ESisSoatExoConv.frx":0D4A
         DownPicture     =   "ESisSoatExoConv.frx":120E
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
         Left            =   2978
         Picture         =   "ESisSoatExoConv.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2640
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   5850
      Begin SIGHReportes.XP_ProgressBar XP_ProgressBar1 
         Height          =   255
         Left            =   90
         TabIndex        =   16
         Top             =   2280
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   450
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
         TabIndex        =   15
         Top             =   1920
         Visible         =   0   'False
         Width           =   1935
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
         TabIndex        =   14
         Top             =   1560
         Width           =   5625
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
         Left            =   2970
         TabIndex        =   8
         Top             =   930
         Visible         =   0   'False
         Width           =   1110
      End
      Begin MSMask.MaskEdBox txtFecha1 
         Height          =   315
         Left            =   1200
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
         Left            =   4110
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
         Left            =   150
         TabIndex        =   9
         Top             =   210
         Width           =   5385
         _ExtentX        =   9499
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
         Left            =   4440
         TabIndex        =   10
         Top             =   990
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
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
      Begin Threed.SSOption optPlan 
         Height          =   255
         Left            =   4290
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
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
         Caption         =   "Por Plan"
         Value           =   -1
      End
      Begin MSDataListLib.DataCombo cmbEstadoCuenta 
         Height          =   330
         Left            =   1200
         TabIndex        =   13
         Top             =   1170
         Width           =   2745
         _ExtentX        =   4842
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estado Cta"
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
         TabIndex        =   12
         Top             =   1230
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "F.Alta Medica"
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
         TabIndex        =   7
         Top             =   690
         Width           =   1080
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
         Left            =   3870
         TabIndex        =   6
         Top             =   690
         Width           =   120
      End
   End
End
Attribute VB_Name = "ESisSoatExoConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: cuentas para liquidación
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim sMensaje As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_idUsuarioConPermisoEnSISoEXOoSOAT As Long
Dim ml_idUsuario As Long
Dim mo_cmbTipoFinanciamiento As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim oRsFuentesFinanciamiento As New Recordset
Dim oRsEstadoCuenta As New Recordset
Private WithEvents oRptHistorias As RptESisSoatExoConv
Attribute oRptHistorias.VB_VarHelpID = -1

Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property



Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        'Dim oRptHistorias As New RptESisSoatExoConv
        oRptHistorias.FechaInicio = txtFecha1.Text
        oRptHistorias.FechaFin = txtFecha2.Text
        oRptHistorias.TextoDelFiltro = "Fecha Alta Médica: " & txtFecha1.Text & " hasta " & txtFecha2.Text & "     (Estado Cta: " & Trim(cmbEstadoCuenta.Text) & ")"
        oRptHistorias.TextoDelFiltro1 = " Cuentas para Liquidación " & IIf(optPlan = True, cmbFuenteFinanciamiento.Text, Trim(cmbTipoFinanciamiento.Text))
        oRptHistorias.IdResponsable = Val(mo_cmbTipoFinanciamiento.BoundText)
        oRptHistorias.IdPlan = Val(cmbFuenteFinanciamiento.BoundText)
        oRptHistorias.IdEstadoCuenta = Val(cmbEstadoCuenta.BoundText)
        oRptHistorias.CrearReporte_excel IIf(chkCtasSinAlta.Value = 1, True, False), Me.hwnd
        Me.MousePointer = 1
    End If
End Sub
Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    
    If optPlan = True Then
      If Val(cmbFuenteFinanciamiento.BoundText) = 0 Then
         sMensaje = "Por favor elija la FUENTE FINANCIAMIENTO/IAFA"
      End If
    Else
      If Val(mo_cmbTipoFinanciamiento.BoundText) = 0 Then
         sMensaje = "Por favor elija el TIPO DE FINANCIAMIENTO"
      End If
    End If
    
    If Me.txtFecha1 = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de alta médica inicial"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFecha1, "DD/MM/AAAA") Then
            sMensaje = "La fecha de alta médica inicial no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFecha2 = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de alta médica final"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFecha2, "DD/MM/AAAA") Then
            sMensaje = "La fecha de alta médica final no tiene el formato correcto"
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




Private Sub cmbTipoFinanciamiento_Click()
    If mo_cmbTipoFinanciamiento.BoundText = "9" Then
       cmbEstadoCuenta.BoundText = "4"
    Else
       cmbEstadoCuenta.BoundText = "10"
    End If
End Sub

Private Sub Form_Initialize()
      Set mo_cmbTipoFinanciamiento.MiComboBox = cmbTipoFinanciamiento
End Sub

Private Sub Form_Load()
       '
       Set oRptHistorias = New RptESisSoatExoConv
       XP_ProgressBar1.ShowText = True
       '
       Me.txtFecha1.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual()
       Me.txtFecha2.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
       '
       Dim lbBuscaPermisoEnFacturacion As New SIGHNegocios.ReglasFacturacion
       
       mo_cmbTipoFinanciamiento.BoundColumn = "idTipoFinanciamiento"
       mo_cmbTipoFinanciamiento.ListField = "Descripcion"
       Set mo_cmbTipoFinanciamiento.RowSource = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia(" and esOficina=1")
       Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
       Dim oRsBuscaLabora As Recordset
       Set oRsBuscaLabora = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghSeguros, ml_idUsuario)
       If oRsBuscaLabora.RecordCount > 0 Then
          ml_idUsuarioConPermisoEnSISoEXOoSOAT = oRsBuscaLabora.Fields!idLaboraSubArea
       End If
       Set oRsBuscaLabora = Nothing
       Set oBuscaDondeLabora = Nothing
       '
       Set oRsFuentesFinanciamiento = mo_ReglasFacturacion.FuentesFinanciamientoSeleccionarTodos
       Set cmbFuenteFinanciamiento.RowSource = oRsFuentesFinanciamiento
       cmbFuenteFinanciamiento.ListField = "Descripcion"
       cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"
       '
       Set oRsEstadoCuenta = mo_ReglasComunes.EstadosCuentaDevuelveTodos
       Set cmbEstadoCuenta.RowSource = oRsEstadoCuenta
       cmbEstadoCuenta.ListField = "Descripcion"
       cmbEstadoCuenta.BoundColumn = "idEstado"
       cmbEstadoCuenta.BoundText = "10"
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub optPlan_Click(Value As Integer)
    If optPlan.Value = True Then
       cmbTipoFinanciamiento.Visible = False
       cmbFuenteFinanciamiento.Visible = True
    End If
End Sub

Private Sub optTipoFinanciamiento_Click(Value As Integer)
   If optTipoFinanciamiento.Value = True Then
       cmbTipoFinanciamiento.Visible = True
       cmbFuenteFinanciamiento.Visible = False
   End If
End Sub

Private Sub oRptHistorias_ProgressActualizaValor(lnValorActual As Long, lnValorTotal As Long)
    XP_ProgressBar1.Max = lnValorTotal
    XP_ProgressBar1.Min = 0
    XP_ProgressBar1.Value = lnValorActual
    DoEvents
    Me.Refresh
End Sub

Private Sub txtFecha1_LostFocus()
    If txtFecha1 <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFecha1, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFecha1 = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFecha2_LostFocus()
    If txtFecha2 <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFecha2, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFecha2 = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub
