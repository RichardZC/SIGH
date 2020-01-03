VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form EconTipoTarifa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo Tarifas"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   Icon            =   "EconTipoTarifa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
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
      Height          =   5325
      Left            =   -15
      TabIndex        =   5
      Top             =   0
      Width           =   9195
      Begin VB.Frame Frame 
         Height          =   1665
         Left            =   675
         TabIndex        =   21
         Top             =   3570
         Width           =   8340
         Begin VB.ComboBox cmbIdResponsable1 
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
            Left            =   4695
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   480
            Width           =   3585
         End
         Begin VB.ComboBox cmbFuenteFinanciamiento 
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
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   510
            Width           =   2475
         End
         Begin Threed.SSOption optCpttodos 
            Height          =   285
            Left            =   210
            TabIndex        =   22
            Top             =   210
            Width           =   870
            _ExtentX        =   1535
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
            Caption         =   "Todos"
            Value           =   -1
         End
         Begin Threed.SSOption optCptUNO 
            Height          =   285
            Left            =   1155
            TabIndex        =   23
            Top             =   210
            Width           =   3195
            _ExtentX        =   5636
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
            Caption         =   "Por una Fuente de Financiamiento"
         End
         Begin Threed.SSOption optPorMedicos 
            Height          =   285
            Left            =   4425
            TabIndex        =   25
            Top             =   210
            Width           =   1725
            _ExtentX        =   3043
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
            Caption         =   "Por un Médico"
         End
      End
      Begin VB.CheckBox chkProrrateoEX 
         Caption         =   "Incluir PRORRATEO de Exoneraciones"
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
         TabIndex        =   16
         Top             =   720
         Value           =   1  'Checked
         Width           =   7155
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   675
         TabIndex        =   11
         Top             =   1410
         Width           =   8340
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   360
            Width           =   4395
         End
         Begin VB.ComboBox cmbTipoTarifa 
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
            Left            =   3885
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1050
            Visible         =   0   'False
            Width           =   4425
         End
         Begin Threed.SSOption optResumen 
            Height          =   285
            Left            =   240
            TabIndex        =   12
            Top             =   780
            Width           =   1245
            _ExtentX        =   2196
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
            Caption         =   "Resumen"
            Value           =   -1
         End
         Begin Threed.SSOption optDetalle 
            Height          =   285
            Left            =   2670
            TabIndex        =   13
            Top             =   765
            Width           =   1245
            _ExtentX        =   2196
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
            Caption         =   "Detalle"
         End
         Begin VB.Label Label5 
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
            Left            =   240
            TabIndex        =   18
            Top             =   390
            Width           =   510
         End
         Begin VB.Label lblTipoTarifa 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Tarifa"
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
            Left            =   2955
            TabIndex        =   15
            Top             =   1095
            Visible         =   0   'False
            Width           =   870
         End
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
         Left            =   7950
         Picture         =   "EconTipoTarifa.frx":0CCA
         TabIndex        =   8
         Top             =   660
         Visible         =   0   'False
         Width           =   1125
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1500
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   6930
         TabIndex        =   1
         Top             =   210
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
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHrFin 
         Height          =   315
         Left            =   8310
         TabIndex        =   10
         Top             =   210
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin Threed.SSOption optTipoTarifa 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   3525
         _ExtentX        =   6218
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
         Caption         =   "Reporte TIPO TARIFA"
      End
      Begin Threed.SSOption optRporItems 
         Height          =   285
         Left            =   150
         TabIndex        =   20
         Top             =   3315
         Width           =   3525
         _ExtentX        =   6218
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
         Caption         =   "Reporte RENDICION POR CPT en CE"
         Value           =   -1
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
         Left            =   6420
         TabIndex        =   7
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Movimiento"
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
         TabIndex        =   6
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   -15
      TabIndex        =   3
      Top             =   5340
      Width           =   9180
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "EconTipoTarifa.frx":0FDC
         DownPicture     =   "EconTipoTarifa.frx":143C
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
         Left            =   3210
         Picture         =   "EconTipoTarifa.frx":18B1
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "EconTipoTarifa.frx":1D26
         DownPicture     =   "EconTipoTarifa.frx":21EA
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
         Left            =   4740
         Picture         =   "EconTipoTarifa.frx":26D6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "EconTipoTarifa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Consumo por Tipo Tarifa
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim oBuscaMedicos As New SIGHNegocios.ReglasDeProgMedica
Dim sMensaje As String
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_TextoDelFiltro As String
Dim lnIdProducto As Long
Dim mo_Formulario As New sighentidades.Formulario
Dim lnIdAlmacen As Long
Dim ml_idUsuario As Long
Dim mo_cmbIdResponsable As New sighentidades.ListaDespleglable
Dim mo_cmbTipoTarifa As New sighentidades.ListaDespleglable
Dim mo_cmbFuenteFinanciamiento As New sighentidades.ListaDespleglable
Dim mo_cmbIdResponsable1 As New sighentidades.ListaDespleglable

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRpt As New RptEtipoTarifa

        If optTipoTarifa.Value = True Then
            oRpt.CreaDatosReporte Val(mo_cmbIdResponsable.BoundText), Val(mo_cmbTipoTarifa.BoundText), _
                                  IIf(chkExcel.Value = 1, True, False), _
                                  Me.Caption & IIf(optResumen.Value = True, " en Resumen", " en Detalle"), _
                                  ml_TextoDelFiltro, _
                                  CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)), CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)), IIf(Me.optResumen.Value = True, 0, Val(mo_cmbTipoTarifa.BoundText)), _
                                  IIf(Me.chkProrrateoEX.Value = 1, True, False), Me.hwnd, False
        ElseIf optRporItems.Value = True Then
           oRpt.CreaDatosReporteXitemIncluyeSeguros 0, 0, _
                                  IIf(chkExcel.Value = 1, True, False), _
                                  optRporItems.Caption & _
                                  IIf(optCpttodos.Value = True, " (Todas Ftes.Financ)", " (" & cmbFuenteFinanciamiento.Text & ")"), _
                                  ml_TextoDelFiltro, _
                                  CDate(txtFdesde.Text), CDate(txtFhasta.Text), _
                                  99, IIf(Me.chkProrrateoEX.Value = 1, True, False), Me.hwnd, True, _
                                  IIf(optCpttodos.Value = True, 0, Val(mo_cmbFuenteFinanciamiento.BoundText)), _
                                  IIf(optPorMedicos.Value = True, Val(mo_cmbIdResponsable1.BoundText), 0)
        End If

        Set oRpt = Nothing
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    ml_TextoDelFiltro = "FILTROS:    F.Movimiento: (" & txtFdesde.Text & " " & txtHrInicio.Text & "   al " & _
                        txtFhasta.Text & " " & txtHrFin.Text & ") " & _
                        IIf(Me.cmbIdResponsable.Text = "", "", "(cajero: " & Trim(Me.cmbIdResponsable.Text) & ")")
    
    If Me.txtFdesde = sighentidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha desde del movimiento"
    Else
        If Not sighentidades.EsFecha(Me.txtFdesde, "DD/MM/AAAA") Then
            sMensaje = "La fecha desde del movimiento no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFhasta = sighentidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha hasta del movimiento"
    Else
        If Not sighentidades.EsFecha(Me.txtFhasta, "DD/MM/AAAA") Then
            sMensaje = "La fecha hasta del movimiento no tiene el formato correcto"
        End If
    End If
    
    If Me.txtHrInicio = sighentidades.HORA_VACIA_HM Then
        sMensaje = "Ingrese la hora desde del movimiento"
    Else
        If Not sighentidades.EsHora(txtHrInicio) Then
            sMensaje = "La hora desde del movimiento no tiene el formato correcto"
        End If
    End If
    
    If Me.txtHrFin = sighentidades.HORA_VACIA_HM Then
        sMensaje = "Ingrese la hora hasta del movimiento"
    Else
        If Not sighentidades.EsHora(txtHrFin) Then
            sMensaje = "La hora hasta del movimiento no tiene el formato correcto"
        End If
    End If
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
       Exit Function
    End If
    
    If optRporItems.Value = True And optCptUNO.Value = True And Val(mo_cmbFuenteFinanciamiento.BoundText) = 0 Then
       MsgBox "Por favor elija la FUENTE DE FINANCIAMIENTO", vbInformation, "Reporte"
       Exit Function
    End If
    If optRporItems.Value = True And optPorMedicos.Value = True And Val(mo_cmbIdResponsable1.BoundText) = 0 Then
       MsgBox "Por favor elija un MEDICO", vbInformation, "Reporte"
       Exit Function
    End If
    
    If Me.optDetalle.Value = True And Me.cmbTipoTarifa.Text = "" Then
        sMensaje = sMensaje & "Debe elegir algún TIPO TARIFA" & Chr(13)
    End If
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ml_TextoDelFiltro = ml_TextoDelFiltro & _
                           IIf(Me.cmbTipoTarifa.Text = "", "", " (Tipo Tarifa: " & Trim(Me.cmbTipoTarifa.Text) & ")") & _
                           IIf(optPorMedicos.Value = True, "  (Médico: " & cmbIdResponsable1.Text & ")", "")
       ValidaDatosObligatorios = True
    End If
End Function


Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub



Private Sub Form_Load()
    txtFdesde.Text = Date
    txtFhasta.Text = Date
    txtHrInicio.Text = "00:01"
    txtHrFin.Text = "23:59"
    Me.txtHrInicio.Visible = False
    Me.txtHrFin.Visible = False
    '
    CargaComboBoxes
End Sub

Sub CargaComboBoxes()
    Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable
    mo_cmbIdResponsable.BoundColumn = "IdEmpleado"
    mo_cmbIdResponsable.ListField = "DCajero"
    Set mo_cmbIdResponsable.RowSource = mo_AdminCaja.CajerosSeleccionarTodos()
    '
    Set mo_cmbTipoTarifa.MiComboBox = cmbTipoTarifa
    mo_cmbTipoTarifa.BoundColumn = "IdTipoTarifa"
    mo_cmbTipoTarifa.ListField = "TipoTarifa"
    Set mo_cmbTipoTarifa.RowSource = mo_reglasComunes.TiposTarifaSeleccionarTodos
    
    Set mo_cmbFuenteFinanciamiento.MiComboBox = cmbFuenteFinanciamiento
    mo_cmbFuenteFinanciamiento.BoundColumn = "idfuenteFinanciamiento"
    mo_cmbFuenteFinanciamiento.ListField = "Descripcion"
    Set mo_cmbFuenteFinanciamiento.RowSource = mo_ReglasFacturacion.FuentesFinanciamientoSeleccionarTodos
    
    Set mo_cmbIdResponsable1.MiComboBox = cmbIdResponsable1
    mo_cmbIdResponsable1.BoundColumn = "IdMedico"
    mo_cmbIdResponsable1.ListField = "Dmedico"
    Set mo_cmbIdResponsable1.RowSource = oBuscaMedicos.MedicosSeleccionarTodosOrdenadoAlfabeticamente
    
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

Private Sub optDetalle_Click(Value As Integer)
    If optDetalle.Value = True Then
       Me.lblTipoTarifa.Visible = True
       Me.cmbTipoTarifa.Visible = True
    End If

End Sub

Private Sub optResumen_Click(Value As Integer)
    If optResumen.Value = True Then
       Me.lblTipoTarifa.Visible = False
       Me.cmbTipoTarifa.Visible = False
    End If
End Sub

Private Sub optRporItems_Click(Value As Integer)
    If optRporItems.Value = True Then
       Me.txtHrInicio.Visible = False
       Me.txtHrFin.Visible = False
    End If
End Sub

Private Sub optTipoTarifa_Click(Value As Integer)
     If optTipoTarifa.Value = True Then
       Me.txtHrInicio.Visible = True
       Me.txtHrFin.Visible = True
     End If
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdesde

End Sub



Private Sub txtFdesde_LostFocus()
    If txtFdesde <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFdesde = sighentidades.FECHA_VACIA_DMY
        End If
    End If

End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFhasta

End Sub

Private Sub txtFhasta_LostFocus()
    If txtFhasta <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFhasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFhasta = sighentidades.FECHA_VACIA_DMY
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


Private Sub txtHrFin_LostFocus()
    If txtHrFin <> sighentidades.HORA_VACIA_HM Then
        If Not sighentidades.EsHora(txtHrFin) Then
            MsgBox "La hora ingresada no es válida", vbInformation, Me.Caption
            txtHrFin = sighentidades.HORA_VACIA_HM
        End If
    End If
End Sub

Private Sub txtHrInicio_LostFocus()
    If txtHrInicio <> sighentidades.HORA_VACIA_HM Then
        If Not sighentidades.EsHora(txtHrInicio) Then
            MsgBox "La hora ingresada no es válida", vbInformation, Me.Caption
            txtHrInicio = sighentidades.HORA_VACIA_HM
        End If
    End If
End Sub
