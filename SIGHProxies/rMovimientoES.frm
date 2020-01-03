VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form rMovimientoES 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos de Entrada y Salida"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rMovimientoES.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   14370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   120
      TabIndex        =   40
      Top             =   6135
      Width           =   14175
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rMovimientoES.frx":0CCA
         DownPicture     =   "rMovimientoES.frx":112A
         Height          =   700
         Left            =   5738
         Picture         =   "rMovimientoES.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rMovimientoES.frx":1A14
         DownPicture     =   "rMovimientoES.frx":1ED8
         Height          =   700
         Left            =   7268
         Picture         =   "rMovimientoES.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6075
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14175
      Begin VB.OptionButton optVentasPorProductos 
         Caption         =   "Ventas por producto según forma pago"
         Height          =   390
         Left            =   255
         TabIndex        =   48
         Top             =   4935
         Width           =   3660
      End
      Begin VB.Frame Frame 
         Height          =   975
         Left            =   9120
         TabIndex        =   43
         Top             =   1560
         Width           =   4215
         Begin VB.OptionButton optOtroMes 
            Caption         =   "Fuera del rango de FECHAS"
            Height          =   390
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   3000
         End
         Begin VB.OptionButton optMismoMes 
            Caption         =   "En el mismo rango de FECHAS"
            Height          =   390
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Value           =   -1  'True
            Width           =   3000
         End
      End
      Begin VB.OptionButton optIngrSalidas 
         Caption         =   "Consolidado de Ingresos y Salidas"
         Height          =   420
         Left            =   255
         TabIndex        =   27
         Top             =   255
         Value           =   -1  'True
         Width           =   3780
      End
      Begin VB.OptionButton optDetalladoIngresos 
         Caption         =   "Ingresos por Proveedor"
         Height          =   390
         Left            =   255
         TabIndex        =   26
         Top             =   4335
         Width           =   2310
      End
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
         Height          =   3645
         Left            =   525
         TabIndex        =   4
         Top             =   600
         Width           =   6645
         Begin VB.ComboBox cmbUsuario 
            Height          =   330
            ItemData        =   "rMovimientoES.frx":28B0
            Left            =   1560
            List            =   "rMovimientoES.frx":28BA
            TabIndex        =   12
            Text            =   "cmbUsuario"
            Top             =   2910
            Width           =   4995
         End
         Begin VB.CheckBox chkExcel 
            Alignment       =   1  'Right Justify
            Caption         =   "En Excel"
            Height          =   315
            Left            =   90
            Picture         =   "rMovimientoES.frx":28D6
            TabIndex        =   11
            Top             =   3240
            Width           =   1665
         End
         Begin VB.ComboBox cmbConcepto 
            Height          =   330
            Left            =   1560
            TabIndex        =   10
            Top             =   624
            Width           =   4980
         End
         Begin VB.ComboBox cmbmovTipo 
            Height          =   330
            ItemData        =   "rMovimientoES.frx":2BE8
            Left            =   1560
            List            =   "rMovimientoES.frx":2BF2
            TabIndex        =   9
            Top             =   1008
            Width           =   4980
         End
         Begin VB.ComboBox cmbEstado 
            Height          =   330
            ItemData        =   "rMovimientoES.frx":2C25
            Left            =   1560
            List            =   "rMovimientoES.frx":2C32
            TabIndex        =   8
            Top             =   1392
            Width           =   4980
         End
         Begin VB.ComboBox cmbAlmacenDestino 
            Height          =   330
            Left            =   1560
            TabIndex        =   7
            Top             =   2160
            Width           =   4980
         End
         Begin VB.ComboBox cmbAlmacenOrigen 
            Height          =   330
            Left            =   1560
            TabIndex        =   6
            Top             =   1776
            Width           =   4980
         End
         Begin VB.ComboBox cmbAlmacen 
            Height          =   330
            Left            =   1560
            TabIndex        =   5
            Top             =   240
            Width           =   4980
         End
         Begin MSMask.MaskEdBox txtFdesde 
            Height          =   315
            Left            =   1560
            TabIndex        =   13
            Top             =   2550
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
            Left            =   4440
            TabIndex        =   14
            Top             =   2550
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
            Left            =   2940
            TabIndex        =   15
            Top             =   2550
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
            Left            =   5790
            TabIndex        =   16
            Top             =   2550
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor Farm"
            Height          =   210
            Left            =   120
            TabIndex        =   25
            Top             =   2970
            Width           =   1260
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Movimiento"
            Height          =   210
            Left            =   120
            TabIndex        =   24
            Top             =   1035
            Width           =   1350
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Almacén Destino"
            Height          =   210
            Left            =   120
            TabIndex        =   23
            Top             =   2190
            Width           =   1365
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Almacén Origen"
            Height          =   210
            Left            =   120
            TabIndex        =   22
            Top             =   1815
            Width           =   1290
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   210
            Left            =   120
            TabIndex        =   21
            Top             =   1425
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   210
            Left            =   120
            TabIndex        =   20
            Top             =   270
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
            Height          =   210
            Left            =   120
            TabIndex        =   19
            Top             =   660
            Width           =   795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "F.Movimiento"
            Height          =   210
            Left            =   120
            TabIndex        =   18
            Top             =   2580
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "hasta"
            Height          =   210
            Left            =   3960
            TabIndex        =   17
            Top             =   2580
            Width           =   435
         End
      End
      Begin VB.OptionButton optCreditoOtorgado 
         Caption         =   "Créditos otorgados y cancelados "
         Height          =   390
         Left            =   7560
         TabIndex        =   3
         Top             =   360
         Width           =   6240
      End
      Begin VB.OptionButton optCreditoPendiente 
         Caption         =   "Créditos pendientes de pago del mes"
         Height          =   390
         Left            =   7560
         TabIndex        =   2
         Top             =   2880
         Width           =   4200
      End
      Begin VB.OptionButton optExoneracion 
         Caption         =   "Exoneraciones del mes"
         Height          =   390
         Left            =   7560
         TabIndex        =   1
         Top             =   4320
         Width           =   4200
      End
      Begin MSMask.MaskEdBox txtFexon1 
         Height          =   315
         Left            =   9120
         TabIndex        =   28
         Top             =   4800
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
      Begin MSMask.MaskEdBox txtExon2 
         Height          =   315
         Left            =   12000
         TabIndex        =   29
         Top             =   4800
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
      Begin MSMask.MaskEdBox txtfPend1 
         Height          =   315
         Left            =   9120
         TabIndex        =   30
         Top             =   3720
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
      Begin MSMask.MaskEdBox txtFpend2 
         Height          =   315
         Left            =   12000
         TabIndex        =   31
         Top             =   3720
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
      Begin MSMask.MaskEdBox txtFmovOtor1 
         Height          =   315
         Left            =   9120
         TabIndex        =   32
         Top             =   1080
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
      Begin MSMask.MaskEdBox txtFmovOtor2 
         Height          =   315
         Left            =   12000
         TabIndex        =   33
         Top             =   1080
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Pacientes: pagantes tomará F.BOLETA,  SEGUROS tomará F.REEMBOLSO"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   7920
         TabIndex        =   47
         Top             =   3360
         Width           =   6000
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Pacientes: pagantes tomará F.BOLETA,  SEGUROS tomará F.REEMBOLSO"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   7920
         TabIndex        =   46
         Top             =   720
         Width           =   6000
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
         Height          =   210
         Left            =   11520
         TabIndex        =   39
         Top             =   4830
         Width           =   435
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "F.Exoneración"
         Height          =   330
         Left            =   7920
         TabIndex        =   38
         Top             =   4830
         Width           =   1140
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
         Height          =   210
         Left            =   11520
         TabIndex        =   37
         Top             =   3750
         Width           =   435
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "F.Movimiento"
         Height          =   210
         Left            =   7920
         TabIndex        =   36
         Top             =   3750
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
         Height          =   210
         Left            =   11520
         TabIndex        =   35
         Top             =   1110
         Width           =   435
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "F.Movimiento"
         Height          =   210
         Left            =   7920
         TabIndex        =   34
         Top             =   1110
         Width           =   1080
      End
   End
End
Attribute VB_Name = "rMovimientoES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte Movimientos de Entrada y Salida
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbAlmacen As New sighentidades.ListaDespleglable
Dim mo_cmbAlmacenOrigen As New sighentidades.ListaDespleglable
Dim mo_cmbAlmacenDestino As New sighentidades.ListaDespleglable
Dim mo_cmbConceptos As New sighentidades.ListaDespleglable
Dim mo_cmbUsuario As New sighentidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim ms_MensajeError As String
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_TextoDelFiltro As String
Dim mo_Formulario As New sighentidades.Formulario
Dim ml_idUsuario As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()

If wxFranklin = "*" Then Exit Sub

    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRep As New RepMovimientoES
        If optIngrSalidas.Value = True Then
            Dim oRptClaseCry As New rCrystal
            oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
            oRptClaseCry.IdAlmacen = Val(mo_cmbAlmacen.BoundText)
            oRptClaseCry.IdAlmacenDestino = Val(mo_cmbAlmacenDestino.BoundText)
            oRptClaseCry.IdAlmacenOrigen = Val(mo_cmbAlmacenOrigen.BoundText)
            oRptClaseCry.Concepto = Val(mo_cmbConceptos.BoundText)
            oRptClaseCry.Estado = cmbEstado.ListIndex
            oRptClaseCry.MovTipo = IIf(cmbmovTipo.ListIndex = 0, "E", "S")
            oRptClaseCry.FechaInicio = CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
            oRptClaseCry.FechaFin = CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
            oRptClaseCry.TextoDelFiltro = ml_TextoDelFiltro
            oRptClaseCry.TipoReporte = Me.Name
            oRptClaseCry.idUsuario = Val(mo_cmbUsuario.BoundText)
            oRptClaseCry.Show vbModal
            Set oRptClaseCry = Nothing
        ElseIf optExoneracion.Value = True Then
            oRep.ReporteExoneraciones txtFexon1.Text, txtExon2.Text, Me.hwnd
        ElseIf optCreditoOtorgado.Value = True Then
            oRep.ReporteCreditosCancelados txtFmovOtor1.Text, txtFmovOtor2.Text & " 23:59:59", Me.hwnd, optMismoMes.Value
        ElseIf optCreditoPendiente.Value = True Then
            oRep.ReporteCreditosPendientes txtfPend1.Text, txtFpend2.Text & " 23:59:59", Me.hwnd
        End If
        Set oRep = Nothing
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    ms_MensajeError = ""
    If optIngrSalidas.Value = True Then
        ml_TextoDelFiltro = "FILTROS:   Almacén: (" & Trim(cmbAlmacen.Text) & ")     F.Movimiento: (" & txtFdesde.Text & " al " & txtFhasta.Text & ")     Tipo Movimiento: (" & Trim(cmbmovTipo.Text) & ")     Estado: (" & Trim(cmbEstado.Text) & ")"
        If cmbConcepto.Text <> "" Then
           ml_TextoDelFiltro = ml_TextoDelFiltro & "     Concepto: (" & Trim(cmbConcepto.Text) & ")"
        End If
        If cmbAlmacenOrigen.Text <> "" Then
           ml_TextoDelFiltro = ml_TextoDelFiltro & "     Alm.Origen: (" & Trim(cmbAlmacenOrigen.Text) & ")"
        End If
        If cmbAlmacenDestino.Text <> "" Then
           ml_TextoDelFiltro = ml_TextoDelFiltro & "     Alm.Destino: (" & Trim(cmbAlmacenDestino.Text) & ")"
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & IIf(Val(mo_cmbUsuario.BoundText) > 0, "     (Vendedor: " & Trim(cmbUsuario.Text) & ")", "")
        If mo_cmbAlmacen.BoundText = "" Then
            ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén" + Chr(13)
            cmbAlmacen.SetFocus
        End If
        If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
           MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
           Exit Function
        End If
    End If
    If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function



Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub



















Private Sub cmbAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacen

End Sub






Private Sub cmbAlmacenDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacenDestino

End Sub

Private Sub cmbAlmacenOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacenOrigen

End Sub

Private Sub cmbConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbConcepto

End Sub



Private Sub cmbEstado_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbEstado

End Sub

Private Sub cmbmovTipo_Click()
    If cmbmovTipo.ListIndex = 0 Then
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacenOrigen, True   'por ser Ingresos
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacenDestino, False   'por ser Ingresos
    Else
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacenOrigen, False   'por ser salidas
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacenDestino, True   'por ser salidas
    End If
End Sub

Private Sub cmbmovTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbmovTipo

End Sub





Private Sub Form_Initialize()
    Set mo_cmbAlmacen.MiComboBox = cmbAlmacen
    Set mo_cmbAlmacenOrigen.MiComboBox = cmbAlmacenOrigen
    Set mo_cmbAlmacenDestino.MiComboBox = cmbAlmacenDestino
    Set mo_cmbConceptos.MiComboBox = cmbConcepto
    Set mo_cmbUsuario.MiComboBox = cmbUsuario
End Sub


Sub InicializaFechaHora()
    txtFdesde.Text = sighentidades.PrimerFechaDDMMYYDelMesActual
    txtFhasta.Text = Date
    txtHrInicio.Text = lcBuscaParametro.SeleccionaFilaParametro(267)
    txtHrFin.Text = lcBuscaParametro.SeleccionaFilaParametro(268)
    txtFexon1.Text = sighentidades.PrimerFechaDDMMYYDelMesActual
    txtExon2.Text = Date
    txtfPend1.Text = sighentidades.PrimerFechaDDMMYYDelMesActual
    txtFpend2.Text = Date
    txtFmovOtor1.Text = sighentidades.PrimerFechaDDMMYYDelMesActual
    txtFmovOtor2.Text = Date
End Sub
Private Sub Form_Load()
    InicializaFechaHora
    
    cmbEstado.ListIndex = 2
    cmbmovTipo.ListIndex = 0    'Ingresos
    cmbmovTipo_Click
    
    '
    mo_cmbAlmacen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacen.ListField = "Descripcion"
    Set mo_cmbAlmacen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarTodosMenosExternos
    '
    mo_cmbAlmacenOrigen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenOrigen.ListField = "Descripcion"
    Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("")
    '
    mo_cmbAlmacenDestino.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenDestino.ListField = "Descripcion"
    Set mo_cmbAlmacenDestino.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("")
    '
    mo_cmbConceptos.BoundColumn = "IdTipoConcepto"
    mo_cmbConceptos.ListField = "Concepto"
    Set mo_cmbConceptos.RowSource = mo_ReglasFarmacia.FarmTipoConceptosDevuelveTodos
    '
    Dim rsIdAlmacen As Recordset
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    Set oBuscaDondeLabora = Nothing
    If rsIdAlmacen.RecordCount > 0 Then
       mo_cmbAlmacen.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacen, False
    End If
    '
    mo_cmbUsuario.BoundColumn = "IdEmpleado"
    mo_cmbUsuario.ListField = "DEmpleado"
    Set mo_cmbUsuario.RowSource = mo_reglasComunes.EmpleadosSeleccionarTodos
    
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

Private Sub optDetalladoIngresos_Click()
    Dim orProductosIngresados As New rProductosIngresados
    orProductosIngresados.NroReporte = 1
    orProductosIngresados.Show 1
    Set orProductosIngresados = Nothing

End Sub



Private Sub optVentasPorProductos_Click()
    Dim lcMensajeLicencia As String
    If mo_reglasComunes.EESSconDerechosAmejoras(2, "61008", lcMensajeLicencia) = True Then
        Dim orProductosIngresados As New rProductosIngresados
        orProductosIngresados.NroReporte = 2
        orProductosIngresados.Show 1
        Set orProductosIngresados = Nothing
    End If
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdesde

End Sub



Private Sub txtFdesde_LostFocus()
    If txtFdesde <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
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
            InicializaFechaHora
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
    Set mo_cmbAlmacen = Nothing
    Set mo_cmbAlmacenOrigen = Nothing
    Set mo_cmbAlmacenDestino = Nothing
    Set mo_cmbConceptos = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_reglasComunes = Nothing
    Set mo_Formulario = Nothing
    
End Sub



Private Sub txtHrFin_LostFocus()
         If Not sighentidades.ValidaHora(txtHrFin.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub

Private Sub txtHrInicio_LostFocus()
If Not sighentidades.ValidaHora(txtHrInicio.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub
