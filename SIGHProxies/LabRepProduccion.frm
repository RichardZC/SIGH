VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form LabRepProduccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laboratorio: Productividad por Fechas"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   Icon            =   "LabRepProduccion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosHistoria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   11265
      Begin VB.CheckBox chkSoloGestantes 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo GESTANTES"
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
         Left            =   4935
         Picture         =   "LabRepProduccion.frx":0CCA
         TabIndex        =   21
         Top             =   660
         Width           =   2085
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   4350
         TabIndex        =   16
         Top             =   1170
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame Frame1 
         Caption         =   "Reportes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   4215
         Begin Threed.SSOption optCPT 
            Height          =   255
            Left            =   150
            TabIndex        =   14
            Top             =   285
            Width           =   3465
            _ExtentX        =   6112
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
            Caption         =   "Mostrando Procedimientos"
            Value           =   -1
         End
         Begin Threed.SSOption optPacientes 
            Height          =   255
            Left            =   150
            TabIndex        =   15
            Top             =   590
            Width           =   3465
            _ExtentX        =   6112
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
            Caption         =   "Mostrando Pacientes"
         End
         Begin Threed.SSOption optXpacienteResul 
            Height          =   255
            Left            =   150
            TabIndex        =   19
            Top             =   895
            Width           =   3825
            _ExtentX        =   6747
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
            Caption         =   "Mostrando Pacientes (solo con Resultados)"
         End
         Begin Threed.SSOption optGrupoR 
            Height          =   255
            Left            =   150
            TabIndex        =   20
            Top             =   1200
            Width           =   3945
            _ExtentX        =   6959
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
            Caption         =   "Por grupo de Exámen (solo con Resultados)"
         End
      End
      Begin VB.CheckBox chkRecalculo 
         Alignment       =   1  'Right Justify
         Caption         =   "Con Recálculo"
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
         Picture         =   "LabRepProduccion.frx":0FDC
         TabIndex        =   11
         Top             =   660
         Value           =   1  'Checked
         Width           =   1785
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
         Left            =   10080
         Picture         =   "LabRepProduccion.frx":12EE
         TabIndex        =   10
         Top             =   660
         Width           =   1035
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1350
         _ExtentX        =   2381
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   8640
         TabIndex        =   2
         Top             =   180
         Width           =   1350
         _ExtentX        =   2381
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
      Begin MSMask.MaskEdBox txtHrInicio 
         Height          =   315
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHrFin 
         Height          =   315
         Left            =   10020
         TabIndex        =   3
         Top             =   180
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   315
         Left            =   4350
         TabIndex        =   17
         Top             =   1680
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar3 
         Height          =   315
         Left            =   4350
         TabIndex        =   18
         Top             =   2160
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   8130
         TabIndex        =   9
         Top             =   210
         Width           =   435
      End
      Begin VB.Label lblFechas 
         AutoSize        =   -1  'True
         Caption         =   "F. Movimiento"
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
         Left            =   180
         TabIndex        =   8
         Top             =   270
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   5
      Top             =   2760
      Width           =   11250
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "LabRepProduccion.frx":1600
         DownPicture     =   "LabRepProduccion.frx":1A60
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
         Left            =   4208
         Picture         =   "LabRepProduccion.frx":1ED5
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "LabRepProduccion.frx":234A
         DownPicture     =   "LabRepProduccion.frx":280E
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
         Left            =   5738
         Picture         =   "LabRepProduccion.frx":2CFA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdAlertaCantidades 
      Height          =   2745
      Left            =   30
      TabIndex        =   12
      Top             =   3900
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   4842
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de MOVIMIENTOS registrados para el chequeo de CANTIDADES"
   End
End
Attribute VB_Name = "LabRepProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de producción
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_cmbIdPuntoCarga As New SIGHEntidades.ListaDespleglable
Dim sMensaje As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim ml_TextoDelFiltro As String
Dim lnIdProducto As Long
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim ml_idUsuario As Long
Dim rsTmp As New Recordset
Dim lcTitulo As String
Private WithEvents orlRepProduccion As rlRepProduccion
Attribute orlRepProduccion.VB_VarHelpID = -1

Property Let idUsuario(lValue As Long)
  ml_idUsuario = lValue
End Property

'debb-12/09/2016
Private Sub btnAceptar_Click()

If wxFranklin = "*" Then Exit Sub

  Me.ProgressBar1.Value = 0
  Me.ProgressBar2.Value = 0
  Me.ProgressBar3.Value = 0
  If ValidaDatosObligatorios Then
     If Me.optCpt.Value = True Or Me.optPacientes.Value = True Then
        lcTitulo = "Filtros:    F.Mov: " & txtFdesde.Text & " " & txtHrInicio.Text & " al " & _
                   txtFhasta.Text & " " & txtHrFin.Text & IIf(Me.chkRecalculo.Value = 1, " (Con Recálculo)", "  (Sin Recálculo)")
     ElseIf optXpacienteResul.Value = True Then
        lcTitulo = "Filtros:    F.Resultado: " & txtFdesde.Text & " " & txtHrInicio.Text & " al " & _
                   txtFhasta.Text & " " & txtHrFin.Text & IIf(Me.chkRecalculo.Value = 1, " (Con Recálculo)", "  (Sin Recálculo)")
     ElseIf optGrupoR.Value = True Then
        lcTitulo = "Filtros:    F.Movimiento: " & txtFdesde.Text & " " & txtHrInicio.Text & " al " & _
                   txtFhasta.Text & " " & txtHrFin.Text & IIf(Me.chkRecalculo.Value = 1, " (Con Recálculo)", "  (Sin Recálculo)") & _
                   " (sólo con EXAMENES CON RESULTADOS)"
     End If
     lcTitulo = lcTitulo & IIf(chkSoloGestantes.Value = 1, "   (" & chkSoloGestantes.Caption & ")", "")
     If Me.optCpt.Value = True Then
        ReportePorProcedimientos IIf(chkSoloGestantes.Value = 1, True, False)
     ElseIf Me.optPacientes.Value = True Then
        orlRepProduccion.ReportePorPacientes Me.txtFdesde.Text, Me.txtHrInicio.Text, Me.txtFhasta.Text, _
                                             Me.txtHrFin.Text, IIf(Me.chkRecalculo.Value = 1, True, False), _
                                             lcTitulo, Me.hwnd, IIf(chkSoloGestantes.Value = 1, True, False)
     ElseIf optXpacienteResul.Value = True Then
        orlRepProduccion.ReportePorPacientesResultado Me.txtFdesde.Text, Me.txtHrInicio.Text, Me.txtFhasta.Text, _
                                                      Me.txtHrFin.Text, IIf(Me.chkRecalculo.Value = 1, True, False), _
                                                      lcTitulo, Me.hwnd, IIf(chkSoloGestantes.Value = 1, True, False)
     ElseIf optGrupoR.Value = True Then
        orlRepProduccion.ReportePorGrupoExamen Me.txtFdesde.Text, Me.txtHrInicio.Text, Me.txtFhasta.Text, _
                                             Me.txtHrFin.Text, IIf(Me.chkRecalculo.Value = 1, True, False), _
                                             lcTitulo, Me.hwnd, IIf(chkSoloGestantes.Value = 1, True, False)
     End If
  End If
End Sub



Sub ReportePorProcedimientos(lbSoloGestantes As Boolean)
        On Error GoTo ErrRptLab
        Dim lnIdFF As Long: Dim lnIdTS As Long: Dim lnIdOrd As Long
        Dim lcNP As String, lcTS As String
        Dim lnSalidas As Long, lnPrecio As Long
        Dim rsReporte As New ADODB.Recordset, mrs_Tmp As New Recordset, oRsTmp1 As New Recordset
        Dim lnIdPlan As Long, lcPlan As String, lnImporte As Double, lnIdProductoCPT As Long, lcProductoCPT As String
        Dim lcCodigoCPT As String, lnIdOrden As Long, lnImporteExonerado As Double, lbNuevo As Boolean
        Dim mda_FechaInicio As Date, lbConsiderar As Boolean
        Dim mda_FechaFin As Date, lRecordCount As Long, f As Long, lnIdMovimiento As Long, ldFechaMovimiento As Date
        lnIdMovimiento = 0
        ldFechaMovimiento = 0
        mda_FechaInicio = Format(txtFdesde.Text & " " & txtHrInicio.Text, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS)
        mda_FechaFin = Format(txtFhasta.Text & " " & txtHrFin.Text, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS)
        If chkRecalculo.Value = 1 Then
            '*********************************************************** con recalculos***************************************
            Dim oConexion As New Connection
            oConexion.CommandTimeout = 300
            oConexion.CursorLocation = adUseClient
            oConexion.Open SIGHEntidades.CadenaConexion
            If mrs_Tmp.State = 1 Then
               Set mrs_Tmp = Nothing
            End If
            With mrs_Tmp
                 .Fields.Append "NroHistoria", adInteger
                 .Fields.Append "Paciente", adVarChar, 160, adFldIsNullable       'Cpt con Codigo
                 .Fields.Append "idCuentaAtencion", adInteger
                 .Fields.Append "idPuntoCarga", adInteger                         'id Plan
                 .Fields.Append "dPuntoCarga", adVarChar, 40, adFldIsNullable     'Planes
                 .Fields.Append "Consumo", adDouble                               'Cantidades o Importes
                 .Fields.Append "FechaEgreso", adDate
                 .Fields.Append "horaEgreso", adVarChar, 5
                 .Fields.Append "dFuenteFinanciamiento", adVarChar, 100           'Codigo CPT
                 .Fields.Append "idTipoServicio", adInteger
                 .LockType = adLockOptimistic
                 .Open
            End With
            'con Boletas - Externos
            lnIdPlan = 1
            lcPlan = "Particular"
            Set rsReporte = mo_ReglasLaboratorio.LabMovimientoLaboratorioSeleccionarPorFechasSoloExternosConBoletas(mda_FechaInicio, mda_FechaFin, oConexion)
            lRecordCount = rsReporte.RecordCount
            If lRecordCount > 0 Then
               Me.ProgressBar1.Min = 0: Me.ProgressBar1.Max = lRecordCount
               f = 0
               rsReporte.MoveFirst
               Do While Not rsReporte.EOF
                  f = f + 1: Me.ProgressBar1.Value = f: DoEvents: Me.Refresh
                  lnImporte = 0
                  ldFechaMovimiento = rsReporte.Fields!fecha
                  lnIdProductoCPT = rsReporte.Fields!idProductoCPT
                  lcProductoCPT = Left(Trim(rsReporte.Fields!nombre) & " " & rsReporte.Fields!Codigo, 100)
                  lcCodigoCPT = rsReporte.Fields!Codigo
                  Do While Not rsReporte.EOF And lnIdProductoCPT = rsReporte.Fields!idProductoCPT
                     '
                     lbConsiderar = True
                     If lbSoloGestantes = True Then
                        If IsNull(rsReporte!eo_eg) Then
                           lbConsiderar = False
                        End If
                     End If
                     '
                     If lbConsiderar = True Then
                        lnImporte = lnImporte + rsReporte.Fields!Importe
                     End If
                     rsReporte.MoveNext
                     If rsReporte.EOF Then
                        Exit Do
                     End If
                  Loop
                  If lnImporte > 0 Then
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!dFuenteFinanciamiento = lcCodigoCPT
                        mrs_Tmp.Fields!Paciente = lcProductoCPT
                        mrs_Tmp.Fields!IdPuntoCarga = lnIdPlan
                        mrs_Tmp.Fields!dPuntoCarga = lcPlan
                        mrs_Tmp.Fields!consumo = lnImporte
                  End If
                  mrs_Tmp.Update
               Loop
            End If
            rsReporte.Close
            'pacientes pagantes, con cuenta
            lnIdPlan = 1
            lcPlan = "Particular"
            Set rsReporte = mo_ReglasLaboratorio.LabMovimientoLaboratorioSeleccionarPorFechasSoloConCuentasPagantes(mda_FechaInicio, mda_FechaFin, oConexion)
            lRecordCount = rsReporte.RecordCount
            If lRecordCount > 0 Then
               Me.ProgressBar2.Min = 0: Me.ProgressBar2.Max = lRecordCount
               f = 0
               rsReporte.MoveFirst
               Do While Not rsReporte.EOF
                  f = f + 1: Me.ProgressBar2.Value = f: DoEvents: Me.Refresh
                  lnIdOrden = rsReporte.Fields!idOrden
                  lnImporteExonerado = rsReporte.Fields!ImporteExonerado
                  ldFechaMovimiento = rsReporte.Fields!fecha
                  If lnImporteExonerado > 0 Then
                     Set oRsTmp1 = mo_ReglasFacturacion.FacturacionServicioPagosFiltraPorIdOrdenConexion(lnIdOrden, oConexion)
                     If oRsTmp1.RecordCount > 0 Then
                        lnImporteExonerado = Round(lnImporteExonerado / oRsTmp1.RecordCount, 2)
                     End If
                     oRsTmp1.Close
                  End If
                  Do While Not rsReporte.EOF And lnIdOrden = rsReporte.Fields!idOrden
                     '
                     lbConsiderar = True
                     If lbSoloGestantes = True Then
                        If IsNull(rsReporte!eo_eg) Then
                           lbConsiderar = False
                        End If
                     End If
                     '
                     If lbConsiderar = True Then
                        lnImporte = rsReporte.Fields!Total - lnImporteExonerado
                        lnIdProductoCPT = rsReporte.Fields!idProducto
                        lcProductoCPT = Left(Trim(rsReporte.Fields!nombre) & " " & rsReporte.Fields!Codigo, 100)
                        lcCodigoCPT = rsReporte.Fields!Codigo
                        lbNuevo = True
                        If mrs_Tmp.RecordCount > 0 Then
                           mrs_Tmp.Find "dFuenteFinanciamiento='" & lcCodigoCPT & "'"
                           If Not mrs_Tmp.EOF Then
                              lbNuevo = False
                           End If
                        End If
                        If lbNuevo = True Then
                           mrs_Tmp.AddNew
                           mrs_Tmp.Fields!dFuenteFinanciamiento = lcCodigoCPT
                           mrs_Tmp.Fields!Paciente = lcProductoCPT
                           mrs_Tmp.Fields!IdPuntoCarga = lnIdPlan
                           mrs_Tmp.Fields!dPuntoCarga = lcPlan
                           mrs_Tmp.Fields!consumo = lnImporte
                        Else
                           mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + lnImporte
                        End If
                        mrs_Tmp.Update
                     End If
                     rsReporte.MoveNext
                     If rsReporte.EOF Then
                        Exit Do
                     End If
                  Loop
               Loop
            End If
            rsReporte.Close
            'con algun Seguro
            Set rsReporte = mo_ReglasLaboratorio.LabMovimientoLaboratorioSeleccionarPorFechasSoloConSeguro(mda_FechaInicio, mda_FechaFin, oConexion)
            lRecordCount = rsReporte.RecordCount
            If lRecordCount > 0 Then
               Me.ProgressBar3.Min = 0: Me.ProgressBar3.Max = lRecordCount: DoEvents: Me.Refresh
               f = 0
               rsReporte.MoveFirst
               Do While Not rsReporte.EOF
                  f = f + 1: Me.ProgressBar3.Value = f: DoEvents
                  lnIdOrden = rsReporte.Fields!idOrden
                  '
                  lbConsiderar = True
                  If lbSoloGestantes = True Then
                       If IsNull(rsReporte!eo_eg) Then
                          lbConsiderar = False
                       End If
                  End If
                  '
                  If lbConsiderar = True Then
                        'Paciente con Seguro que pagó (no cubre)
                        lnIdPlan = 1
                        lcPlan = "Particular"
                        Set oRsTmp1 = mo_ReglasFacturacion.FacturacionServicioPagosFiltraPorIdOrdenConexion(lnIdOrden, oConexion)
                        If oRsTmp1.RecordCount > 0 Then
                           lnImporteExonerado = oRsTmp1.Fields!ImporteExonerado
                           If lnImporteExonerado > 0 Then
                              lnImporteExonerado = Round(lnImporteExonerado / oRsTmp1.RecordCount)
                           End If
                           oRsTmp1.MoveFirst
                           Do While Not oRsTmp1.EOF
                              lnImporte = oRsTmp1.Fields!Total - lnImporteExonerado
                              lcProductoCPT = Left(Trim(oRsTmp1.Fields!nombre) & " " & oRsTmp1.Fields!Codigo, 100)
                              lcCodigoCPT = oRsTmp1.Fields!Codigo
                              lbNuevo = True
                              If mrs_Tmp.RecordCount > 0 Then
                                 mrs_Tmp.MoveFirst
                                 Do While Not mrs_Tmp.EOF
                                      If mrs_Tmp.Fields!dFuenteFinanciamiento = lcCodigoCPT And mrs_Tmp.Fields!IdPuntoCarga = 1 Then
                                         lbNuevo = False
                                         Exit Do
                                      End If
                                      mrs_Tmp.MoveNext
                                 Loop
                              End If
                              If lbNuevo = True Then
                                 mrs_Tmp.AddNew
                                 mrs_Tmp.Fields!dFuenteFinanciamiento = lcCodigoCPT
                                 mrs_Tmp.Fields!Paciente = lcProductoCPT
                                 mrs_Tmp.Fields!IdPuntoCarga = lnIdPlan
                                 mrs_Tmp.Fields!dPuntoCarga = lcPlan
                                 mrs_Tmp.Fields!consumo = lnImporte
                              Else
                                 mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + lnImporte
                              End If
                              mrs_Tmp.Update
                              oRsTmp1.MoveNext
                           Loop
                        End If
                        oRsTmp1.Close
                        'Seguro que reembolsará (si cubre el Seguro)
                        lnIdPlan = rsReporte.Fields!idFuenteFinanciamiento
                        lcPlan = rsReporte.Fields!dFuenteFinanciamiento
                        Set oRsTmp1 = mo_ReglasFacturacion.FacturacionServicioFinanciamientosFiltraPorIdOrdenConexion(lnIdOrden, oConexion)
                        If oRsTmp1.RecordCount > 0 Then
                           oRsTmp1.MoveFirst
                           Do While Not oRsTmp1.EOF
                              lnImporte = oRsTmp1.Fields!totalFinanciado
                              lcProductoCPT = Left(Trim(oRsTmp1.Fields!nombre) & " " & oRsTmp1.Fields!Codigo, 100)
                              lcCodigoCPT = oRsTmp1.Fields!Codigo
                              lbNuevo = True
                              If mrs_Tmp.RecordCount > 0 Then
                                 mrs_Tmp.MoveFirst
                                 Do While Not mrs_Tmp.EOF
                                      If mrs_Tmp.Fields!dFuenteFinanciamiento = lcCodigoCPT And mrs_Tmp.Fields!IdPuntoCarga = lnIdPlan Then
                                         lbNuevo = False
                                         Exit Do
                                      End If
                                      mrs_Tmp.MoveNext
                                 Loop
                              End If
                              If lbNuevo = True Then
                                 mrs_Tmp.AddNew
                                 mrs_Tmp.Fields!dFuenteFinanciamiento = lcCodigoCPT
                                 mrs_Tmp.Fields!Paciente = lcProductoCPT
                                 mrs_Tmp.Fields!IdPuntoCarga = lnIdPlan
                                 mrs_Tmp.Fields!dPuntoCarga = lcPlan
                                 mrs_Tmp.Fields!consumo = lnImporte
                              Else
                                 mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + lnImporte
                              End If
                              mrs_Tmp.Update
                              oRsTmp1.MoveNext
                           Loop
                        End If
                        oRsTmp1.Close
                  End If
                  rsReporte.MoveNext
               Loop
            End If
            rsReporte.Close
            '
            oConexion.Close
            If mrs_Tmp.RecordCount = 0 Then
              MsgBox "No hay datos para mostrar", vbInformation, "SIGH "
            Else
              Me.MousePointer = 11
              Dim oRptClaseCry1 As New frmCrystalR
              oRptClaseCry1.TextoDelFiltro = lcTitulo
              oRptClaseCry1.Excel = IIf(chkExcel.Value = 1, True, False)
              oRptClaseCry1.Archivo = "EconConsumoXpto"
              oRptClaseCry1.Tabla = mrs_Tmp
              oRptClaseCry1.Show vbModal
              Set oRptClaseCry1 = Nothing
              Me.MousePointer = 1
            End If
            
        Else
            '**************************************** sin recalculos, tal como se despacho ********************************
            Set rsReporte = mo_ReglasLaboratorio.LabMovimientoLaboratorioSeleccionarPorFechas(mda_FechaInicio, mda_FechaFin, sghPorIdProductoMasFecha)
            rsReporte.Filter = "IdFuenteFinanciamiento>=0"
            lRecordCount = rsReporte.RecordCount
            If rsReporte.RecordCount > 0 Then
               Me.ProgressBar1.Min = 0: Me.ProgressBar1.Max = lRecordCount
               f = 0
              If rsTmp.State = adStateOpen Then Set rsTmp = Nothing
              With rsTmp
                .Fields.Append "idPlan", adInteger, 4, adFldIsNullable
                .Fields.Append "Plan", adVarChar, 50, adFldIsNullable
                .Fields.Append "dTServicio", adVarChar, 10, adFldIsNullable
                .Fields.Append "Cantidad", adInteger, 4, adFldIsNullable
                .Fields.Append "Total", adDouble
                .LockType = adLockOptimistic
                .Open
              End With
              rsReporte.MoveFirst
              Do While Not rsReporte.EOF
                f = f + 1: Me.ProgressBar1.Value = f: DoEvents: Me.Refresh
                lnIdFF = rsReporte.Fields!idFuenteFinanciamiento
                lcNP = rsReporte.Fields!nombrePlan
                lnIdOrd = rsReporte.Fields!idOrden
                If IsNull(rsReporte.Fields!idTipoServicio) Then
                    lnIdTS = 0
                    lcTS = "EXTERNO"
                    lnSalidas = 0: lnPrecio = 0
                    Do While Not rsReporte.EOF And lnIdFF = rsReporte.Fields!idFuenteFinanciamiento
                        '
                        lbConsiderar = True
                        If lbSoloGestantes = True Then
                             If IsNull(rsReporte!eo_eg) Then
                                lbConsiderar = False
                             End If
                        End If
                        '
                        If lbConsiderar = True Then
                            lnSalidas = lnSalidas + rsReporte.Fields!Cantidad
                            lnPrecio = lnPrecio + rsReporte.Fields!Total
                        End If
                        rsReporte.MoveNext
                        If rsReporte.EOF Or (Not IsNull(rsReporte.Fields!idTipoServicio)) Then Exit Do
                    Loop
                Else
                    lnIdTS = rsReporte.Fields!idTipoServicio
                    lcTS = IIf(lnIdTS = 1, "CE", IIf(lnIdTS = 3, "HOSP", "EMERG"))
                    lnSalidas = 0: lnPrecio = 0
                    Do While Not rsReporte.EOF And lnIdFF = rsReporte.Fields!idFuenteFinanciamiento And lnIdTS = rsReporte.Fields!idTipoServicio
                        '
                        lbConsiderar = True
                        If lbSoloGestantes = True Then
                             If IsNull(rsReporte!eo_eg) Then
                                lbConsiderar = False
                             End If
                        End If
                        '
                        If lbConsiderar = True Then
                            lnSalidas = lnSalidas + rsReporte.Fields!Cantidad
                            lnPrecio = lnPrecio + rsReporte.Fields!Total
                        End If
                        rsReporte.MoveNext
                        If rsReporte.EOF Then Exit Do
                    Loop
                End If
                If lnSalidas > 0 Then
                    rsTmp.AddNew
                    rsTmp.Fields!IdPlan = CStr(lnIdFF)
                    rsTmp.Fields!Plan = lcNP
                    rsTmp.Fields!dTServicio = lcTS
                    rsTmp.Fields!Cantidad = lnSalidas
                    rsTmp.Fields!Total = lnPrecio
                    rsTmp.Update
                End If
              Loop
            End If
            
            If rsTmp.State = adStateClosed Then
              MsgBox "No hay datos para mostrar", vbInformation, "SIGH "
              Exit Sub
            End If
            If rsTmp.EOF = True And rsTmp.BOF = True Then
              MsgBox "No hay datos para mostrar", vbInformation, "SIGH "
            Else
              Me.MousePointer = 11
              Dim oRptClaseCry As New frmCrystalR
              oRptClaseCry.Excel = IIf(chkExcel.Value = 1, True, False)
              oRptClaseCry.Archivo = "labProducPagoDeuda"
              oRptClaseCry.Tabla = rsTmp
              oRptClaseCry.Show vbModal
              Set oRptClaseCry = Nothing
              Set rsTmp = Nothing
              Me.MousePointer = 1
            End If
        End If
        Set rsReporte = Nothing
        Set mrs_Tmp = Nothing
        Set oRsTmp1 = Nothing
        Exit Sub
ErrRptLab:
  MsgBox Err.Description & Chr(13) & "Movimiento: " & lnIdMovimiento & Chr(13) & "id CPT: " & lnIdProductoCPT & _
        Chr(13) & " Fecha Mov: " & ldFechaMovimiento
  Resume

End Sub

Function ValidaDatosObligatorios() As Boolean
  If txtFdesde.Text = "" Or txtFdesde.Text = SIGHEntidades.FECHA_VACIA_DMY Or txtFhasta.Text = "" Or txtFhasta.Text = SIGHEntidades.FECHA_VACIA_DMY Then
    MsgBox "Ingrese Fechas de Inicio y Fecha de Fin", vbInformation, "SIGH "
    txtFdesde.SetFocus
    ValidaDatosObligatorios = False
  Else
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Function
    End If
    ValidaDatosObligatorios = True
  End If
End Function

Private Sub btnCancelar_Click()
  Me.Visible = False
  LimpiarVariablesDeMemoria
End Sub

Private Sub Form_Initialize()
  'Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPuntoDeCarga
End Sub
Sub InicializaFechaHora()
  txtFdesde.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual
  txtFhasta.Text = Date
  txtHrInicio.Text = "00:00:00"
  txtHrFin.Text = "23:59:59"

End Sub
Private Sub Form_Load()
  '
  Set orlRepProduccion = New rlRepProduccion
  '
  InicializaFechaHora
  '
  Set grdAlertaCantidades.DataSource = mo_ReglasLaboratorio.LaboratorioDevuelveAlertaDeCantidades
  mo_Apariencia.ConfigurarFilasBiColores grdAlertaCantidades, SIGHEntidades.GrillaConFilasBicolor
  '
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
  Select Case KeyCode
    Case vbKeyEscape
      btnCancelar_Click
    Case vbKeyF2
      btnAceptar_Click
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  LimpiarVariablesDeMemoria
End Sub

Private Sub optCPT_Click(Value As Integer)
    If optCpt.Value = True Then
       lblFechas.Caption = "F. Movimiento"
       chkExcel.Enabled = True
       chkRecalculo.Enabled = True
       chkRecalculo.Value = 1
    End If
End Sub

Private Sub optGrupoR_Click(Value As Integer)
    If optGrupoR.Value = True Then
       lblFechas.Caption = "F. Movimiento"
       chkExcel.Enabled = False
       chkRecalculo.Enabled = True
       chkRecalculo.Value = 1
    End If
End Sub

Private Sub optPacientes_Click(Value As Integer)
    If optPacientes.Value = True Then
       lblFechas.Caption = "F.Movimiento"
       chkExcel.Enabled = False
       chkRecalculo.Enabled = True
       chkRecalculo.Value = 1
    End If
End Sub

Private Sub optXpacienteResul_Click(Value As Integer)
    If optPacientes.Value = True Then
       chkRecalculo.Value = 0
       lblFechas.Caption = "F.Resultado"
       chkExcel.Enabled = False
       chkRecalculo.Enabled = False
    End If
End Sub

Private Sub orlRepProduccion_ProgressActualizaValor(lnValorActual As Long, lnValorTotal As Long)
    ProgressBar1.Max = lnValorTotal
    ProgressBar1.Min = 0
    ProgressBar1.Value = lnValorActual
    DoEvents
    Me.Refresh
End Sub



Private Sub txtFdesde_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFdesde_LostFocus()
  If txtFdesde <> SIGHEntidades.FECHA_VACIA_DMY Then
    If Not SIGHEntidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
      MsgBox "La Fecha Inicial ingresada no es válida", vbInformation, "SIGH "
      InicializaFechaHora
      txtFdesde.SetFocus
    End If
  End If
End Sub



Private Sub txtFhasta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFhasta_LostFocus()
  If txtFhasta <> SIGHEntidades.FECHA_VACIA_DMY Then
    If Not SIGHEntidades.EsFecha(txtFhasta, "DD/MM/AAAA") Then
      MsgBox "La Fecha Final ingresada no es válida", vbInformation, "SIGH "
      InicializaFechaHora
      txtFhasta.SetFocus
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
  On Error Resume Next
  Set mo_ReglasFarmacia = Nothing
  Set mo_Teclado = Nothing
  Set mo_ReglasFacturacion = Nothing
  Set mo_reglasComunes = Nothing
  Set mo_Formulario = Nothing
End Sub



Private Sub txtHrFin_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtHrFin_LostFocus()
  If txtHrFin.Text <> "__:__:__" Then
    If Not IsDate(txtHrFin.Text) Then
      MsgBox "La Hora Final ingresada no es válida.", vbInformation, "SIGH "
      InicializaFechaHora
      txtHrFin.SetFocus
    End If
  End If
End Sub



Private Sub txtHrInicio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtHrInicio_LostFocus()
  If txtHrInicio.Text <> "__:__:__" Then
    If Not IsDate(txtHrInicio.Text) Then
      MsgBox "La Hora Inicial ingresada no es válida.", vbInformation, "SIGH "
      InicializaFechaHora
      txtHrInicio.SetFocus
    End If
  End If
End Sub
