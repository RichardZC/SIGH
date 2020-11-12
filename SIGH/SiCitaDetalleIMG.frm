VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form SiCitaDetalleIMG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13845
   Icon            =   "SiCitaDetalleIMG.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraPaciente 
      Caption         =   "Datos del Paciente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   15
      TabIndex        =   36
      Top             =   3360
      Width           =   13755
      Begin VB.TextBox txtDireccion 
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
         Left            =   10830
         MaxLength       =   100
         TabIndex        =   39
         Top             =   510
         Width           =   2700
      End
      Begin VB.TextBox txtTelefono 
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
         Left            =   10830
         MaxLength       =   10
         TabIndex        =   37
         Top             =   150
         Width           =   2700
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Dirección domicilio"
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
         Left            =   9345
         TabIndex        =   40
         Top             =   600
         Width           =   1470
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "N° Celular"
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
         Left            =   9345
         TabIndex        =   38
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame fraDatosAtencion 
      Caption         =   "Datos de Cabecera"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   15
      TabIndex        =   5
      Top             =   15
      Width           =   13755
      Begin VB.ComboBox cmbResponsable 
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
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   2550
         Width           =   7200
      End
      Begin VB.CommandButton cmbBuscaReceta 
         Height          =   330
         Left            =   2475
         Picture         =   "SiCitaDetalleIMG.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1035
         Width           =   300
      End
      Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
         Height          =   330
         Left            =   2475
         Picture         =   "SiCitaDetalleIMG.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   615
         Width           =   300
      End
      Begin VB.CheckBox chkMuestraHistorico 
         Alignment       =   1  'Right Justify
         Caption         =   "Muestra Histórico de exámenes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10680
         TabIndex        =   19
         Top             =   180
         Width           =   2895
      End
      Begin VB.TextBox txtNreceta 
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
         Left            =   1215
         MaxLength       =   30
         TabIndex        =   18
         Top             =   1035
         Width           =   1245
      End
      Begin VB.TextBox txtEstado 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3030
         MaxLength       =   30
         TabIndex        =   17
         Top             =   255
         Width           =   750
      End
      Begin VB.TextBox txtNcita 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9210
         MaxLength       =   30
         TabIndex        =   16
         Top             =   240
         Width           =   1005
      End
      Begin VB.TextBox txtPlan 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2835
         TabIndex        =   15
         Top             =   1005
         Width           =   5550
      End
      Begin VB.TextBox txtDatosDeCuenta 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2820
         TabIndex        =   14
         Top             =   600
         Width           =   5565
      End
      Begin VB.TextBox txtNcuenta 
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
         Left            =   1215
         MaxLength       =   30
         TabIndex        =   13
         Top             =   615
         Width           =   1245
      End
      Begin VB.TextBox txtProcedencia 
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
         Left            =   1215
         TabIndex        =   12
         Top             =   1425
         Width           =   7170
      End
      Begin VB.TextBox txtNserie 
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
         Left            =   1215
         MaxLength       =   4
         TabIndex        =   11
         Top             =   2175
         Width           =   585
      End
      Begin VB.TextBox txtNboleta 
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
         Left            =   1830
         MaxLength       =   30
         TabIndex        =   10
         Top             =   2175
         Width           =   1020
      End
      Begin VB.TextBox txtDx 
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
         Left            =   1215
         MaxLength       =   30
         TabIndex        =   9
         ToolTipText     =   "Ingrese el Dx (4 dígitos)"
         Top             =   1800
         Width           =   1260
      End
      Begin VB.TextBox txtNombreDx 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   8
         Top             =   1785
         Width           =   5865
      End
      Begin VB.TextBox txtCupo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1215
         MaxLength       =   30
         TabIndex        =   6
         Top             =   240
         Width           =   810
      End
      Begin SISGalenPlus.UcPacienteDatos UcPacienteDatos1 
         Height          =   2775
         Left            =   9195
         TabIndex        =   7
         Top             =   525
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   4895
      End
      Begin UltraGrid.SSUltraGrid grdConsumoPaciente 
         Height          =   2565
         Left            =   8655
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   4524
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   71303188
         BorderStyle     =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   "SiCitaDetalleIMG.frx":13DE
         Caption         =   "Exámenes históricos del Paciente (Consulta Externa, Hospitalización, Emergencia)"
      End
      Begin MSMask.MaskEdBox txtFregistro 
         Height          =   315
         Left            =   4410
         TabIndex        =   23
         Top             =   255
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFcita 
         Height          =   315
         Left            =   6420
         TabIndex        =   24
         Top             =   255
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHoraCita 
         Height          =   315
         Left            =   7635
         TabIndex        =   25
         Top             =   255
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Responsable"
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
         TabIndex        =   42
         Top             =   2580
         Width           =   1005
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
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
         Left            =   2430
         TabIndex        =   35
         Top             =   285
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N° Cita"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8580
         TabIndex        =   34
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N° Boleta"
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
         TabIndex        =   33
         Top             =   2205
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "N° Cuenta"
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
         TabIndex        =   32
         Top             =   693
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Procedencia"
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
         TabIndex        =   31
         Top             =   1449
         Width           =   990
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Diagnóstico"
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
         TabIndex        =   30
         Top             =   1827
         Width           =   930
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "N° Receta"
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
         TabIndex        =   29
         Top             =   1071
         Width           =   870
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "F.Cita"
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
         Left            =   5925
         TabIndex        =   28
         Top             =   285
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Reg"
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
         Left            =   3915
         TabIndex        =   27
         Top             =   285
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Cupo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   26
         Top             =   285
         Width           =   705
      End
   End
   Begin SISGalenPlus.ucFacturacionIteIMG ucFacturacionProductos 
      Height          =   2430
      Left            =   15
      TabIndex        =   4
      Top             =   4395
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   4286
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   15
      TabIndex        =   0
      Top             =   6870
      Width           =   13710
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "SiCitaDetalleIMG.frx":141A
         DownPicture     =   "SiCitaDetalleIMG.frx":18DE
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
         Left            =   6870
         Picture         =   "SiCitaDetalleIMG.frx":1DCA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "SiCitaDetalleIMG.frx":22B6
         DownPicture     =   "SiCitaDetalleIMG.frx":2716
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
         Left            =   5340
         Picture         =   "SiCitaDetalleIMG.frx":2B8B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprime"
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
         Left            =   120
         Picture         =   "SiCitaDetalleIMG.frx":3000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1365
      End
   End
End
Attribute VB_Name = "SiCitaDetalleIMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Ecografía General
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_IdMovimiento As Long
Dim mi_Opcion As sghOpciones
Dim ms_MensajeError As String
Dim ml_idUsuario As Long
Dim mb_ExistenDatos As Boolean
Dim mo_AdminProgramacionMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim mo_Procesos As New SIGHProxies.Procesos
Dim wxParametro302 As String, lnIdTipoServicio As Long
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_cmbIdEstado As New sighEntidades.ListaDespleglable
Dim mo_cmbResponsable As New sighEntidades.ListaDespleglable
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim oRsProgramacion As New Recordset
Dim lbPrimeraVez As Boolean
Dim ml_IdTipoFinanciamiento As Long
Dim ml_IdPaciente As Long
Dim ml_IdComprobantePago As Long
Dim ml_IdFuenteFinanciamiento  As Long
Dim ml_IdServicioPaciente As Long
Dim ml_IdDiagnostico As Long
Dim oDOPaciente As New doPaciente
Dim DoSiCitas As New DoSiCitas
Dim rsProductosCPT As Recordset
Dim rsPuntosCarga As New Recordset
Dim ml_IdFuenteFinanciamientoDespacho As Long
Dim ml_PuntoCarga As sghPuntosCargaBasicos
Const lcConstanteMovimientoSalida As String = "S"
Dim ml_IdTipoVentaSeleccionada As Long
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim lnIdReceta As Long
Dim lnUltimaBusqueda As sghUltimaBusqueda
Dim lnIdPacienteHistorico As Long
Dim ml_SeEligioGridBoleta As Boolean
Dim wxParametro509 As String
Dim lnEpsPorcentaje As Double
Dim lcMedicoDNI As String, lcCama As String, lcMedico As String, lnMedicoId As Long
Dim ml_fechaCita As Date
Dim ml_horaCita As String
Dim lnCuposXdia As Integer, lnNroMinutos As Long, lcHoraInicioCita As String
Dim ml_idSala As Long
Dim lcFinanciamiento As String, lcHistoria As String

Property Let idSala(lValue As Long)
    ml_idSala = lValue
End Property


Property Let fechaCita(lValue As Date)
    ml_fechaCita = lValue
End Property
Property Let horaCita(lValue As String)
    ml_horaCita = lValue
End Property


Property Let PuntoCarga(lValue As sghPuntosCargaBasicos)
    ml_PuntoCarga = lValue
End Property


Property Let SeEligioGridBoleta(lValue As Boolean)
    ml_SeEligioGridBoleta = lValue
End Property
Property Get SeEligioGridBoleta() As Boolean
    SeEligioGridBoleta = ml_SeEligioGridBoleta
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property

Property Let IdMovimiento(lValue As Long)
    ml_IdMovimiento = lValue
End Property

Property Get IdMovimiento() As Long
    IdMovimiento = ml_IdMovimiento
End Property



Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   mo_ReglasComunes.DevuelveCamaYdniMedico lcMedico, lcMedicoDNI, lcCama, 0, lnMedicoId, ml_IdPaciente
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If AgregarDatos() Then
                    mo_Procesos.EnviaMensajeCelularPorCuenta Val(Me.txtNcuenta.Text), "Cita N° " & DoSiCitas.idCitaSI & _
                                              " para el " & DoSiCitas.fecha & " en " & DevuelveNombrePuntoCarga, "SiCita"
                    Me.txtNcita.Text = DoSiCitas.idCitaSI
                    MsgBox "Se agregó correctamente la CITA N° " & DoSiCitas.idCitaSI, vbInformation, Me.Caption
                    btnImprimir_Click
                    Me.Visible = False
                    LimpiarVariablesDeMemoria
                Else
                    MsgBox "No se pudo agregar los datos" & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If ModificarDatos() Then
                    
                    MsgBox "Se Modificó correctamente lc CITA N° " & DoSiCitas.idCitaSI, vbInformation, Me.Caption
                    btnImprimir_Click
                    Me.Visible = False
                    LimpiarVariablesDeMemoria
                Else
                    MsgBox "No se pudo modificar los datos" & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
            If MsgBox("¿Realmente desea Anular?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                 Exit Sub
            End If
           If ValidarReglas() Then
                CargaDatosAlObjetosDeDatos
               If EliminarDatos() Then
                    MsgBox "Los datos se Anularon correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                    LimpiarVariablesDeMemoria
                Else
                    MsgBox "No se pudo anular los datos" & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
        
End Sub



Function ValidarDatosObligatorios() As Boolean
    On Error GoTo ErrVald
    Dim lnTabError As Integer
    Dim lcCuposQueQuedan As Long
    ValidarDatosObligatorios = False
    If txtNcuenta.Text = "" And txtNboleta.Text = "" Then
       Exit Function
    End If
    ms_MensajeError = ""
    UcPacienteDatos1.CargarDatosAlObjetoDatos oDOPaciente
'    If  .Text = "" Then
'       If oDOPaciente.nrodocumento = "" Then
'           ms_MensajeError = ms_MensajeError & "Tiene que registrar el N° DNI" & Chr(13)
'           lnTabError = 1
'       End If
'        If oDOPaciente.ApellidoPaterno = "" Then
'           ms_MensajeError = ms_MensajeError & "Tiene que registrar el Apellido Paterno" & Chr(13)
'           lnTabError = 1
'       End If
'       If oDOPaciente.ApellidoMaterno = "" Then
'           ms_MensajeError = ms_MensajeError & "Tiene que registrar el Apellido Materno" & Chr(13)
'           lnTabError = 1
'       End If
'       If oDOPaciente.PrimerNombre = "" Then
'           ms_MensajeError = ms_MensajeError & "Tiene que registrar el Primer Nombre" & Chr(13)
'           lnTabError = 1
'       End If
'       If oDOPaciente.FechaNacimiento = 0 Then
'           ms_MensajeError = ms_MensajeError & "Tiene que registrar la FECHA DE NACIMIENTO" & Chr(13)
'       End If
'       If oDOPaciente.idTipoSexo = 0 Then
'           ms_MensajeError = ms_MensajeError & "Tiene que elegir el SEXO" & Chr(13)
'       End If
'    Else
'    End If
    If cmbResponsable.Text = "" Then
       ms_MensajeError = ms_MensajeError & "Tiene que elegir el RESPONSABLE" & Chr(13)
    Else
        
        lcCuposQueQuedan = Val(Mid(cmbResponsable.Text, InStr(cmbResponsable.Text, "quedan:") + 8, 2))
        If lcCuposQueQuedan <= 0 Then
             ms_MensajeError = ms_MensajeError & "No quedan CUPOS para el RESPONSABLE elejido" & Chr(13)
        Else
             oRsProgramacion.MoveFirst
             oRsProgramacion.Move cmbResponsable.ListIndex
        End If
    End If
    Select Case mi_Opcion
    Case sghAgregar, sghModificar
        'Cpt
        Set rsProductosCPT = ucFacturacionProductos.FacturacionProductos
        If Not (rsProductosCPT.EOF And rsProductosCPT.BOF) Then
            rsProductosCPT.MoveFirst
            Do While Not rsProductosCPT.EOF
                If rsProductosCPT!idProducto = 0 Then
                   rsProductosCPT.Delete
                   rsProductosCPT.Update
                Else
                   If rsProductosCPT!Cantidad <= 0 Then
                      ms_MensajeError = ms_MensajeError & "El producto CPT: " & rsProductosCPT!Codigo & " " & Trim(rsProductosCPT!NombreProducto) & "   Tiene problemas con la Cantidad" & Chr(13)
                   End If
                   If rsProductosCPT!PrecioUnitario <= 0 And rsProductosCPT!SeUsaSinPrecio = False Then
                      If Val(Me.txtNboleta.Text) = 0 Then  'debb-05/04/2011
                         ms_MensajeError = ms_MensajeError & "El producto CPT: " & rsProductosCPT!Codigo & " " & Trim(rsProductosCPT!NombreProducto) & "   Tiene problemas con el Precio" & Chr(13)
                      End If
                   End If
                   If txtNboleta.Text = "" Then
                      'chequeo solo para pacientes con  Nro Cuenta
                      rsProductosCPT.Fields!TotalPorPagar = Round(rsProductosCPT!Cantidad * rsProductosCPT!PrecioUnitario, 2)
                   End If
                End If
                rsProductosCPT.MoveNext
            Loop
        Else
            ms_MensajeError = ms_MensajeError & "No hay ITEMS" & Chr(13)
        End If
    End Select
    If ms_MensajeError = "" Then
       ValidarDatosObligatorios = True
    Else
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       Select Case lnTabError
       Case 1
           UcPacienteDatos1.SetFocusOnApellidoPaterno
       End Select
    End If
ErrVald:
End Function


Sub CargaDatosAlObjetosDeDatos()
        With DoSiCitas
             .IdProgramacion = oRsProgramacion!IdProgramacion
             .idResponsable = oRsProgramacion!idResponsable
             .fecha = txtFcita.Text
             .fechacreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
             .FechaNacimiento = oDOPaciente.FechaNacimiento
             .HoraFinal = mo_AdminProgramacionMedica.CalculaHoraFinal(txtHoraCita.Text, lnNroMinutos)
             .HoraInicio = txtHoraCita.Text
             '.idCitaSI
             If mi_Opcion = sghAgregar Then
                .IdComprobantePago = ml_IdComprobantePago
             End If
             .idCuentaAtencion = Val(txtNcuenta.Text)
             .idEstado = sghSiCitasEstados.sghSiCitaActiva
             .idMedico = lnMedicoId
             '.IdMovimiento
             .idPaciente = oDOPaciente.idPaciente
             .idPuntoCarga = ml_PuntoCarga
             .idTipoSexo = oDOPaciente.idTipoSexo
             .idUsuario = sighEntidades.Usuario
             .IdUsuarioAuditoria = sighEntidades.Usuario
             .Paciente = oDOPaciente.ApellidoPaterno & " " & oDOPaciente.ApellidoMaterno & " " & oDOPaciente.PrimerNombre
             .idReceta = Val(txtNreceta.Text)
             .idSala = ml_idSala
             '.Cupo
             .TELEFONO = Me.txtTelefono.Text
             .Direccion = Me.txtDireccion.Text
        End With
End Sub

Function ValidarReglas() As Boolean
   ValidarReglas = False
  
    
   ValidarReglas = True
End Function
Function AgregarDatos() As Boolean
    AgregarDatos = mo_ReglasImagenes.SICitasMantenimiento(DoSiCitas, rsProductosCPT, mo_lnIdTablaLISTBARITEMS, _
                                                          mo_lcNombrePc, sghAgregar, True, lnNroMinutos)
    ms_MensajeError = mo_ReglasImagenes.MensajeError
    ml_IdMovimiento = DoSiCitas.idCitaSI
End Function

Function ModificarDatos() As Boolean
    ModificarDatos = mo_ReglasImagenes.SICitasMantenimiento(DoSiCitas, rsProductosCPT, mo_lnIdTablaLISTBARITEMS, _
                                                          mo_lcNombrePc, sghModificar, False, lnNroMinutos)
    ms_MensajeError = mo_ReglasImagenes.MensajeError
End Function

Function EliminarDatos() As Boolean
    Set rsProductosCPT = ucFacturacionProductos.FacturacionProductos
    EliminarDatos = mo_ReglasImagenes.SICitasMantenimiento(DoSiCitas, rsProductosCPT, mo_lnIdTablaLISTBARITEMS, _
                                                          mo_lcNombrePc, sghEliminar, False, lnNroMinutos)
    ms_MensajeError = mo_ReglasImagenes.MensajeError
End Function





Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub




Private Sub chkMuestraHistorico_Click()
    If chkMuestraHistorico.Value = 1 Then
       Me.UcPacienteDatos1.Visible = False
       grdConsumoPaciente.Visible = True
       If lnIdPacienteHistorico > 0 Then
          If mi_Opcion = sghAgregar Then
             Set grdConsumoPaciente.DataSource = mo_ReglasImagenes.CptHistoricosPorPaciente(lnIdPacienteHistorico, ml_PuntoCarga, 0)
          Else
             Set grdConsumoPaciente.DataSource = mo_ReglasImagenes.CptHistoricosPorPaciente(lnIdPacienteHistorico, ml_PuntoCarga, ml_IdMovimiento)
          End If
         ' grdConsumoPaciente.Top = Me.UcPacienteDatos1.Top
          grdConsumoPaciente.Left = Me.UcPacienteDatos1.Left
          grdConsumoPaciente.Width = Me.UcPacienteDatos1.Width
          grdConsumoPaciente.Caption = "Históricos de exámenes: " & Me.UcPacienteDatos1.DevuelvePaciente
       End If
    Else
       grdConsumoPaciente.Visible = False
       Me.UcPacienteDatos1.Visible = True
    End If

End Sub









Function DevuelveNombrePuntoCarga() As String
    Select Case ml_PuntoCarga
    Case sghPtoCargaTomografia
         DevuelveNombrePuntoCarga = "Tomografía"
         ucFacturacionProductos.FiltraCpt = sghCptSoloTomografia
    Case sghPtoCargaRayosX
         DevuelveNombrePuntoCarga = "Rayos X"
         ucFacturacionProductos.FiltraCpt = sghCptSoloRayosX
    Case sghPtoCargaPatologiaClinica
         DevuelveNombrePuntoCarga = "Patología Clínica"
         ucFacturacionProductos.FiltraCpt = sghCptSoloLaboratorio
    Case sghPtoCargaEcogObstetrica
         DevuelveNombrePuntoCarga = "Ecografía Obstétrica"
         ucFacturacionProductos.FiltraCpt = sghCptSoloEcografiaObstetrica
    Case sghPtoCargaEcogGeneral
         DevuelveNombrePuntoCarga = "Ecografía General"
         ucFacturacionProductos.FiltraCpt = sghCptSoloEcografiaGeneral
    Case sghPtoCargaBancoSangre1
         DevuelveNombrePuntoCarga = "Banco de Sangre"
         ucFacturacionProductos.FiltraCpt = sghCptSoloLaboratorio
    Case sghPtoCargaAnatomiaPatologica1
         DevuelveNombrePuntoCarga = "Anatomía Patológica"
         ucFacturacionProductos.FiltraCpt = sghCptSoloLaboratorio
    End Select
    
End Function

Sub CargaDatosPuntoCarga()
    lnCuposXdia = 0
    lnNroMinutos = 0
    lcHoraInicioCita = ""
    Dim oRsTmp9 As New Recordset
    Dim lnNumeroCitas As Long
    Set oRsTmp9 = mo_ReglasImagenes.SiCitasSeleccionarPorIdPuntoCArgaYFecha(ml_PuntoCarga, ml_fechaCita)
    If oRsTmp9.RecordCount > 0 Then
        lnCuposXdia = IIf(IsNull(oRsTmp9!NroCupos), 0, oRsTmp9!NroCupos)
        lnNroMinutos = IIf(IsNull(oRsTmp9!nroCuposMinutos), 0, oRsTmp9!nroCuposMinutos)
        lcHoraInicioCita = IIf(IsNull(oRsTmp9!HoraInicio), sighEntidades.HORA_VACIA_HM, oRsTmp9!HoraInicio)
        lcHoraInicioCita = mo_AdminProgramacionMedica.CalculaHoraFinal(lcHoraInicioCita, lnNroMinutos)
    Else
        oRsTmp9.Close
        Set oRsTmp9 = mo_ReglasComunes.FactPuntosCargaSeleccionarPorId(ml_PuntoCarga)
        If oRsTmp9.RecordCount > 0 Then
           lnCuposXdia = IIf(IsNull(oRsTmp9!NroCupos), 0, oRsTmp9!NroCupos)
           lnNroMinutos = IIf(IsNull(oRsTmp9!nroCuposMinutos), 0, oRsTmp9!nroCuposMinutos)
           lcHoraInicioCita = IIf(IsNull(oRsTmp9!HoraInicioDiaCita), sighEntidades.HORA_VACIA_HM, oRsTmp9!HoraInicioDiaCita)
        End If
    End If
    oRsTmp9.Close
    If mi_Opcion = sghAgregar Then
'       Set oRsTmp9 = mo_ReglasImagenes.SiCitasPorDia(ml_idSala, ml_fechaCita)
'       Me.txtCupo.Text = Trim(Str(oRsTmp9.RecordCount + 1))
'       If lnCuposXdia = lnNumeroCitas Then
'          MsgBox "Llegó al máximo de CITAS", vbInformation, ""
'          Me.btnAceptar.Visible = False
'       End If
'       oRsTmp9.Close
    End If
    Set oRsTmp9 = Nothing
End Sub



Private Sub cmbResponsable_Click()
    On Error GoTo errRsp
    oRsProgramacion.MoveFirst
    oRsProgramacion.Move cmbResponsable.ListIndex
    CargaCupoYhoraCita
errRsp:
End Sub

Private Sub Form_Load()
    CargaDatosPuntoCarga
    txtFregistro.Text = lcBuscaParametro.RetornaFechaServidorSQL
    txtEstado.Text = "Registrado"
    txtFcita.Text = Format(ml_fechaCita, sighEntidades.DevuelveFechaSoloFormato_DMY)
    txtHoraCita.Text = lcHoraInicioCita
    
    CargaDataCombos
    
    ucFacturacionProductos.Opcion = mi_Opcion
    ucFacturacionProductos.idUsuario = ml_idUsuario
    ucFacturacionProductos.Inicializar
    
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Cita para " + DevuelveNombrePuntoCarga
    Case sghModificar
        Me.Caption = "Modificar Cita para " + DevuelveNombrePuntoCarga
    Case sghConsultar
        Me.Caption = "Consultar Cita para " + DevuelveNombrePuntoCarga
        btnImprimir.Visible = True
        fraDatosAtencion.Enabled = False
    Case sghEliminar
        Me.Caption = "Eliminar Cita para " + DevuelveNombrePuntoCarga
    End Select
    
    CargarDatosAlFormulario
End Sub

Sub CargarDatosAlFormulario()
 mo_Formulario.HabilitarDeshabilitar Me.txtCupo, False
 mo_Formulario.HabilitarDeshabilitar Me.txtNcita, False
 mo_Formulario.HabilitarDeshabilitar Me.txtFregistro, False
 mo_Formulario.HabilitarDeshabilitar Me.txtEstado, False
 mo_Formulario.HabilitarDeshabilitar Me.txtDatosDeCuenta, False
 mo_Formulario.HabilitarDeshabilitar Me.txtPlan, False
 mo_Formulario.HabilitarDeshabilitar Me.txtProcedencia, False
 mo_Formulario.HabilitarDeshabilitar Me.txtNombreDx, False
 wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
 wxParametro509 = lcBuscaParametro.SeleccionaFilaParametro(509)
 Me.UcPacienteDatos1.Inicializar

 Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
        CargarDatosAlosControles
     Case sghConsultar
        CargarDatosAlosControles
     Case sghEliminar
        CargarDatosAlosControles
 End Select
End Sub

Sub CargarDatosAlosControles()
        lnMedicoId = 0
        mo_Formulario.HabilitarDeshabilitar Me.txtNcuenta, False
        mo_Formulario.HabilitarDeshabilitar Me.txtNserie, False
        mo_Formulario.HabilitarDeshabilitar Me.txtNboleta, False
        txtDatosDeCuenta.Width = txtDatosDeCuenta.Width + 150
        cmdBuscaCuentaPorApellidos.Enabled = False
        
        'Carga datos de la orden
        Dim oRsTmp As New Recordset
        Dim oConexion As New Connection
        Dim oSiCitas As New SiCitas
        Dim lbSigue As Boolean, lbSeguirConCuentaCerrada As Boolean, lnIndex As Integer
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 900
        oConexion.Open sighEntidades.CadenaConexion
        Set oSiCitas.Conexion = oConexion
        DoSiCitas.IdUsuarioAuditoria = sighEntidades.Usuario
        DoSiCitas.idCitaSI = ml_IdMovimiento
        txtNcita.Text = ml_IdMovimiento
        If oSiCitas.SeleccionarPorId(DoSiCitas) = True Then
           With DoSiCitas
           
                lnIndex = 0
                oRsProgramacion.MoveFirst
                Do While Not oRsProgramacion.EOF
                     If oRsProgramacion!IdProgramacion = .IdProgramacion Then
                        Me.cmbResponsable.ListIndex = lnIndex
                        Exit Do
                     End If
                     lnIndex = lnIndex + 1
                     oRsProgramacion.MoveNext
                Loop
                
                Me.txtTelefono.Text = .TELEFONO
                Me.txtDireccion.Text = .Direccion
                Me.txtCupo.Text = .Cupo
                ml_idSala = .idSala
                txtNreceta.Text = IIf(.idReceta > 0, .idReceta, "")
                txtFregistro.Text = Format(.fechacreacion, sighEntidades.DevuelveFechaSoloFormato_DMY)
                txtFcita.Text = Format(.fecha, sighEntidades.DevuelveFechaSoloFormato_DMY)
                txtHoraCita.Text = .HoraInicio
                txtEstado.Text = IIf(.idEstado = sghSiCitasEstados.sghSiCitaActiva, "Registrado", "Con TM")
                ml_PuntoCarga = .idPuntoCarga
                lnMedicoId = .idMedico
                ml_IdPaciente = .idPaciente
                lnIdPacienteHistorico = ml_IdPaciente
           End With
           '
           UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
           If ml_IdPaciente = 0 Then
                If DoSiCitas.FechaNacimiento <> 0 Then
                   UcPacienteDatos1.FechaNacimiento = DoSiCitas.FechaNacimiento
                End If
                If DoSiCitas.idTipoSexo > 0 Then
                   UcPacienteDatos1.idTipoSexo = DoSiCitas.idTipoSexo
                End If
                UcPacienteDatos1.CargaAlgunosDatosDesdeBoleta DoSiCitas.Paciente
           Else
                UcPacienteDatos1.idPaciente = ml_IdPaciente
                UcPacienteDatos1.CargarDatosDePacienteALosControles
                
                lcHistoria = UcPacienteDatos1.NroHistoriaClinica
                
           End If
           
           If DoSiCitas.idCuentaAtencion > 0 Then
               txtNcuenta.Text = DoSiCitas.idCuentaAtencion
               Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(Val(txtNcuenta.Text), oConexion)
               lbSigue = True
               If oRsTmp.RecordCount > 0 Then
                    lnMedicoId = IIf(IsNull(oRsTmp!IdMedicoIngreso), 0, oRsTmp!IdMedicoIngreso)
                    If oRsTmp.Fields!idEstado <> 1 Then
                       If mi_Opcion <> sghConsultar Then
                          '
                          lbSeguirConCuentaCerrada = True
                          If mi_Opcion = sghModificar And oRsTmp!idTipoServicio = sghTipoServicio.sghEmergenciaConsultorios Then
                            If mo_ReglasComunes.HospitalizadoConCtaEmergNOabierta(ml_IdPaciente, _
                               Format(oRsTmp!fechaEgreso & " " & oRsTmp!HoraEgreso, sighEntidades.DevuelveFechaSoloFormato_DMY_HM), _
                               oRsTmp!IdDestinoAtencion) = True Then
                               lbSeguirConCuentaCerrada = False
                               UcPacienteDatos1.habilitar False
                               MsgBox "Ese estado de Cuenta no se encuentra ABIERTA" & Chr(13) & _
                                      "       ", vbInformation, Me.Caption
                            End If
                          End If
                          '
                          If lbSeguirConCuentaCerrada = True Then
                             MsgBox "Ese estado de Cuenta no se encuentra ABIERTA", vbInformation, Me.Caption
                             If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
                                btnAceptar.Enabled = False
                             Else
                                lbSigue = False
                             End If
                          End If
                       End If
                    End If
                    If lbSigue Then
                          lnIdTipoServicio = oRsTmp.Fields!idTipoServicio
                          txtDatosDeCuenta.Text = "F.Ing: " & oRsTmp.Fields!FechaIngreso & " - " & IIf(oRsTmp.Fields!idTipoServicio = 1, "Consultorios Externos", IIf(oRsTmp.Fields!idTipoServicio = 3, "Hospitalización", "Emergencia")) & " - (Est: " & Trim(oRsTmp.Fields!estadoCta) & ")"
                          txtPlan.Text = "IAFA Act.: " & oRsTmp!dFuenteFinanciamiento
                          lcFinanciamiento = oRsTmp.Fields!dFuenteFinanciamiento
                          'debb-14/04/2011
                    End If
               End If
               '
               
               Set oRsTmp = mo_ReglasComunes.RecetaCabeceraFiltraXcuentaYDocumentodespacho(txtNcita.Text, Val(txtNcuenta.Text))
               lnIdReceta = 0
               ucFacturacionProductos.PermiteAgregarItems = True
               If oRsTmp.RecordCount > 0 Then
                   lnIdReceta = oRsTmp.Fields!idReceta
                   ucFacturacionProductos.PermiteAgregarItems = False
               End If
               UcPacienteDatos1.DeshabilitarFrames True
               ucFacturacionProductos.CargaProductosPorIdCitaSI Val(txtNcita.Text)
               '
               chkMuestraHistorico.Value = 1
               chkMuestraHistorico_Click
           ElseIf DoSiCitas.IdComprobantePago > 0 Then
                Dim oDOCajaComprobantesPago As New DOCajaComprobantesPago
                Set oDOCajaComprobantesPago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(DoSiCitas.IdComprobantePago, oConexion)
                txtNserie.Text = oDOCajaComprobantesPago.nroSerie
                txtNboleta.Text = oDOCajaComprobantesPago.nrodocumento
                ucFacturacionProductos.PermiteAgregarItems = False
                UcPacienteDatos1.DeshabilitarFrames False
                If ml_IdServicioPaciente > 0 Then
                   'Paciente contado, con cuenta (CE), pago en CAJA
                   ml_IdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(oDOCajaComprobantesPago.idCuentaAtencion, CDate(txtFregistro.Text), lcBuscaParametro.RetornaHoraServidorSQL)
                   txtProcedencia.Text = mo_ReglasFacturacion.BuscaServicioActualDelPaciente(ml_IdServicioPaciente)
                   UcPacienteDatos1.DeshabilitarFrames True
                End If
                Set oDOCajaComprobantesPago = Nothing
                ucFacturacionProductos.CargaProductosPorIdCitaSI Val(txtNcita.Text)
            End If
            If mi_Opcion = sghConsultar Then
               btnAceptar.Enabled = False
            End If
            mb_ExistenDatos = True
        
            UcPacienteDatos1.CargarDatosAlObjetoDatos oDOPaciente
        Else
            mb_ExistenDatos = False
        End If
        oConexion.Close
        Set oConexion = Nothing
        Set oRsTmp = Nothing
        Set oSiCitas = Nothing
End Sub


Private Sub cmdBuscaCuentaPorApellidos_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New doPaciente
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            ml_IdPaciente = oDOPaciente.idPaciente
            Dim oRsTmp As New Recordset
            Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(ml_IdPaciente, oConexion, True)
            If oRsTmp.RecordCount > 0 Then
               txtNcuenta.Text = oRsTmp.Fields!idCuentaAtencion
            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
            txtNcuenta_LostFocus
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub grdConsumoPaciente_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
     grdConsumoPaciente.Top = fraDatosAtencion.Top + 1050
     grdConsumoPaciente.Bands(0).Columns("Fecha").Width = 800
     grdConsumoPaciente.Bands(0).Columns("idMovimiento").Width = 700
     grdConsumoPaciente.Bands(0).Columns("Codigo").Width = 500
     grdConsumoPaciente.Bands(0).Columns("Nombre").Width = 2500
     grdConsumoPaciente.Bands(0).Columns("Cantidad").Width = 300

End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
 mo_Teclado.RealizarNavegacion KeyCode, txtDireccion
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFcita_LostFocus()
If Not IsDate(txtFcita.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFcita.Text = sighEntidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If
End Sub

Private Sub txtNboleta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNboleta
End Sub

Private Sub txtNboleta_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
End Sub

Private Sub txtNboleta_LostFocus()
    If Trim(txtNserie.Text) <> "" And Trim(txtNboleta.Text) <> "" Then
        lcFinanciamiento = ""
        lcHistoria = ""

        lnMedicoId = 0
        lnEpsPorcentaje = 0
        lnUltimaBusqueda = sghEnBoleta
        Dim rsBuscaBoleta As New Recordset
        Dim rsBuscaBoletaEnImagenes As New Recordset
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set rsBuscaBoleta = mo_AdminCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumentoConexion(txtNserie.Text, Trim(txtNboleta.Text), oConexion)
        If rsBuscaBoleta.RecordCount > 0 Then
            '
            lnIdPacienteHistorico = 0
            If rsBuscaBoleta.Fields!idPaciente > 0 Then
               lnIdPacienteHistorico = rsBuscaBoleta.Fields!idPaciente
               chkMuestraHistorico_Click
            End If
            '
            If rsBuscaBoleta.Fields!idEstadoComprobante <> sghEstadosComprobante.sighEstadosComprobantePagado Then
                MsgBox "Esa Boleta está ANULADA", vbInformation, Me.Caption
                txtNboleta.Text = ""
                txtNserie.Text = ""
                ml_IdComprobantePago = 0
            Else
                Set rsBuscaBoletaEnImagenes = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorIdComprobantePago(rsBuscaBoleta.Fields!IdComprobantePago)
                If rsBuscaBoletaEnImagenes.RecordCount > 0 Then
                    MsgBox "Esa Boleta ya fué DESPACHADA con N° Movimiento: " & Chr(13) & rsBuscaBoletaEnImagenes.Fields!IdMovimiento & "      y fecha: " & rsBuscaBoletaEnImagenes.Fields!fecha, vbInformation, Me.Caption
                    txtNboleta.Text = ""
                    txtNserie.Text = ""
                    ml_IdComprobantePago = 0
                Else
                    UcPacienteDatos1.LimpiarDatosDePaciente
                    Set rsBuscaBoletaEnImagenes = mo_AdminCaja.FactOrdenServicioSeleccionarPuntoCargaPorIdOrden(rsBuscaBoleta.Fields!IdOrden)
                    If rsBuscaBoletaEnImagenes.RecordCount > 0 Then
                        ml_IdTipoFinanciamiento = rsBuscaBoletaEnImagenes.Fields!idTipoFinanciamiento     'Contado
                        ml_IdFuenteFinanciamiento = rsBuscaBoletaEnImagenes.Fields!IdFuenteFinanciamiento 'contado
                    End If
                    ml_IdComprobantePago = rsBuscaBoleta.Fields!IdComprobantePago
                    If rsBuscaBoleta.Fields!idPaciente > 0 And rsBuscaBoleta.Fields!idCuentaAtencion > 0 Then
                       'Paciente contado, con cuenta (CE), pago en CAJA
                       ml_IdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(rsBuscaBoleta.Fields!idCuentaAtencion, CDate(txtFregistro.Text), lcBuscaParametro.RetornaHoraServidorSQL)
                       txtProcedencia.Text = mo_ReglasFacturacion.BuscaServicioActualDelPaciente(ml_IdServicioPaciente)
                       UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
                       UcPacienteDatos1.idPaciente = rsBuscaBoleta.Fields!idPaciente
                       UcPacienteDatos1.CargarDatosDePacienteALosControles
                       UcPacienteDatos1.DeshabilitarFrames True
                    ElseIf rsBuscaBoleta.Fields!idPaciente > 0 Then
                       'Paciente con Nro Historia
                       UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
                       UcPacienteDatos1.idPaciente = rsBuscaBoleta.Fields!idPaciente
                       UcPacienteDatos1.CargarDatosDePacienteALosControles
                       UcPacienteDatos1.DeshabilitarFrames True
                    Else
                       'Paciente contado, EXTERNO
                       UcPacienteDatos1.CargaAlgunosDatosDesdeBoleta (rsBuscaBoleta.Fields!razonSocial)
                       UcPacienteDatos1.DeshabilitarFrames False
                       UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
                    End If
                    ucFacturacionProductos.LimpiarGrilla
                    ucFacturacionProductos.TipoProducto = sghServicio
                    ucFacturacionProductos.IdOrdenPago = rsBuscaBoleta!IdOrdenPago
                    ucFacturacionProductos.CargaProductosPorIdOrdenPago
                    ucFacturacionProductos.PermiteAgregarItems = False
                    txtNcuenta.Text = ""
                    txtDatosDeCuenta.Text = ""
                    txtPlan.Text = ""
                    txtProcedencia.Text = ""
                    txtDx.Text = ""
                    txtNombreDx.Text = ""
                    On Error Resume Next
                    ucFacturacionProductos.SetFocus
                End If
            End If
        End If
        Set rsBuscaBoleta = Nothing
        Set rsBuscaBoletaEnImagenes = Nothing
        oConexion.Close
        Set oConexion = Nothing
    End If
End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta
End Sub


Sub CargaBoletaAutomaticamente()
    If ml_SeEligioGridBoleta = True And ml_IdMovimiento > 0 Then
        Dim oRsTmp1 As New Recordset
        Set oRsTmp1 = mo_AdminCaja.CajaComprobantesSeleccionarPorId(ml_IdMovimiento)
        If oRsTmp1.RecordCount > 0 Then
            Me.txtNserie = oRsTmp1.Fields!nroSerie
            Me.txtNboleta = oRsTmp1.Fields!nrodocumento
            txtNboleta_LostFocus
        End If
        Set oRsTmp1 = Nothing
    End If
End Sub


Private Sub txtNcuenta_LostFocus()
   If Val(txtNcuenta.Text) = 0 And txtNcuenta.Locked = False Then
      txtNserie.SetFocus
      Exit Sub
   End If
   If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) And txtNcuenta.Locked = False Then
      lcFinanciamiento = ""
      lcHistoria = ""
   
       lnMedicoId = 0
       lnUltimaBusqueda = sghEnNroCuenta
       Dim oRsTmp As New Recordset
       Dim lbSigue As Boolean
       Dim oConexion As New Connection
       oConexion.Open sighEntidades.CadenaConexion
       oConexion.CursorLocation = adUseClient
       Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
       lbSigue = True
       If oRsTmp.RecordCount > 0 Then
          lnMedicoId = IIf(IsNull(oRsTmp!IdMedicoIngreso), 0, oRsTmp!IdMedicoIngreso)
          If oRsTmp.Fields!idEstado <> 1 Then
             If mi_Opcion <> sghConsultar Then
                MsgBox "Ese estado de Cuenta no se encuentra ABIERTA", vbInformation, Me.Caption
                If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
                   btnAceptar.Enabled = False
                Else
                   lbSigue = False
                End If
             End If
          End If
          '
          If lbSigue = True And oRsTmp!EsPacienteExterno <> True And wxParametro509 = "S" And mi_Opcion = sghAgregar Then
             If Val(txtNreceta.Text) = 0 Then
                MsgBox "No puede usar N° CUENTA, tiene que generar RECETA", vbInformation, Me.Caption
                lbSigue = False
             End If
          End If
          '
          
          If mi_Opcion = sghAgregar And _
             mo_AdminAdmision.AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS(Val(txtNcuenta.Text), wxParametro302, _
                                                                          oRsTmp.Fields!IdFuenteFinanciamiento) = False Then
                                                                       
             lbSigue = False
          End If
          
          If mi_Opcion = sghAgregar And oRsTmp.Fields!idTipoServicio = sghTipoServicio.sghConsultaExterna _
                                                                                And oRsTmp.Fields!IdFormaPago = 1 Then
                MsgBox "Es un Paciente PAGANTE y viene por CONSULTORIO EXTERNO" & Chr(13) & _
                        "Debe pagar antes en CAJA", vbInformation, "Imágenes"
                lbSigue = False
          End If
          If mi_Opcion = sghAgregar And _
                                    mo_AdminAdmision.LaFechaDespachoEsMenorAfechaCita(CDate(Format(oRsTmp!FechaIngreso, _
                                    sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & oRsTmp!HoraIngreso)) = True Then
             lbSigue = False
          End If
          
          If lbSigue Then
                lnIdTipoServicio = oRsTmp.Fields!idTipoServicio
                txtDatosDeCuenta.Text = "F.Ing: " & oRsTmp.Fields!FechaIngreso & " - " & _
                        IIf(oRsTmp!EsPacienteExterno = True, "Externo", _
                        IIf(oRsTmp.Fields!idTipoServicio = 1, "Consultorios Externos", _
                        IIf(oRsTmp.Fields!idTipoServicio = 3, "Hospitalización", "Emergencia"))) & _
                        " - (Est: " & Trim(oRsTmp.Fields!estadoCta) & ")"
                lcFinanciamiento = oRsTmp.Fields!dFuenteFinanciamiento
                lcHistoria = Trim(Str(oRsTmp!NroHistoriaClinica))
                        
                txtPlan.Text = "IAFA Act.: " & oRsTmp.Fields!dFuenteFinanciamiento
                ml_IdPaciente = oRsTmp.Fields!idPaciente
                ml_IdFuenteFinanciamiento = oRsTmp.Fields!IdFuenteFinanciamiento
                ml_IdTipoFinanciamiento = oRsTmp.Fields!IdFormaPago
                UcPacienteDatos1.idPaciente = ml_IdPaciente
                UcPacienteDatos1.FechaRegistro = CDate(txtFregistro.Text)
                UcPacienteDatos1.CargarDatosDePacienteALosControles
                UcPacienteDatos1.DeshabilitarFrames True
                txtTelefono.Text = UcPacienteDatos1.TELEFONO
                Me.txtDireccion.Text = UcPacienteDatos1.Direccion
                ucFacturacionProductos.LimpiarGrilla
                ucFacturacionProductos.TipoProducto = sghServicio
                ucFacturacionProductos.idPuntoCarga = ml_PuntoCarga
                ucFacturacionProductos.idTipoFinanciamiento = ml_IdTipoFinanciamiento
                ucFacturacionProductos.MaximoNroItems = 100
                ucFacturacionProductos.idCuentaAtencion = Val(txtNcuenta.Text)
                ucFacturacionProductos.PermiteAgregarItems = True
                ucFacturacionProductos.AgregaProducto
                ucFacturacionProductos.TabEnDescripcion
                '
                ml_IdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(Val(txtNcuenta.Text), CDate(txtFregistro.Text), lcBuscaParametro.RetornaHoraServidorSQL)
                txtProcedencia.Text = mo_ReglasFacturacion.BuscaServicioActualDelPaciente(ml_IdServicioPaciente)
                '
                txtNserie.Text = ""
                txtNboleta.Text = ""
                ml_IdComprobantePago = 0
                If mi_Opcion <> sghAgregar And ml_IdFuenteFinanciamiento <> ml_IdFuenteFinanciamientoDespacho And ml_IdFuenteFinanciamientoDespacho > 0 Then
                   txtPlan.Text = "Plan Desp: " & Trim(mo_ReglasFacturacion.FuentesFinanciamientoDevuelveNombrePlan(ml_IdFuenteFinanciamientoDespacho)) & " - " & txtPlan.Text
                End If
                '
                lnIdPacienteHistorico = ml_IdPaciente
                chkMuestraHistorico.Value = 1
                chkMuestraHistorico_Click
                '
                mo_Formulario.HabilitarDeshabilitar txtDx, True
                Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarTodosPorIdAtencion(oRsTmp.Fields!idAtencion)
                If oRsTmp.RecordCount > 0 Then
                   txtDx.Text = oRsTmp.Fields!CodigoCIE2004
                   txtNombreDx.Text = oRsTmp.Fields!descripcion
                   mo_Formulario.HabilitarDeshabilitar txtDx, False
                   ml_IdDiagnostico = oRsTmp!idDiagnostico
                End If
          Else
                txtNreceta.Text = ""
          
          End If
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
       oConexion.Close
       Set oConexion = Nothing
   End If
End Sub



Private Sub txtNserie_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNserie
End Sub

Private Sub txtNserie_KeyPress(KeyAscii As Integer)
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
End Sub


Sub CargaDataCombos()
    Dim lnNroCuposXDia As Long
    mo_ReglasImagenes.SiProgramacionLlenaComboCuposMaximo Me.cmbResponsable, ml_idSala, ml_fechaCita, lnNroCuposXDia, True, _
                      sghSiProgramacionEnDataCombo, oRsProgramacion
    If oRsProgramacion.RecordCount = 1 Then
       oRsProgramacion.MoveFirst
       CargaCupoYhoraCita
    End If
     
     
End Sub

Sub CargaCupoYhoraCita()
    txtCupo.Text = ""
    txtHoraCita.Text = sighEntidades.HORA_VACIA_HM
    lnNroMinutos = oRsProgramacion!TiempoPromedioAtencion
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = mo_ReglasImagenes.siCitasXidprogramacion(oRsProgramacion!IdProgramacion)
    txtCupo.Text = Trim(Str(oRsTmp1.RecordCount + 1))
    If oRsTmp1.RecordCount > 0 Then
       
       txtHoraCita.Text = mo_AdminProgramacionMedica.CalculaHoraFinal(oRsTmp1!HoraInicio, lnNroMinutos)
    Else
       txtHoraCita.Text = oRsProgramacion!HoraInicio
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Sub

Private Sub txtDx_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDx
    AdministrarKeyPreview KeyCode
End Sub







Private Sub txtResultadoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
' mo_Teclado.RealizarNavegacion KeyCode, txtResultadoFinal
 AdministrarKeyPreview KeyCode
End Sub







Private Sub txtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
 mo_Teclado.RealizarNavegacion KeyCode, txtTelefono
    AdministrarKeyPreview KeyCode
End Sub

Private Sub ucFacturacionProductos_SePresionoTeclaEspecial(KeyCode As Integer)
     If KeyCode = vbKeyF2 Then
        AdministrarKeyPreview KeyCode
     End If
    
End Sub


Private Sub UcPacienteDatos1_SePresionoTeclaEspecial(KeyCode As Integer)
     If KeyCode = vbKeyF2 Then
        AdministrarKeyPreview KeyCode
     End If
End Sub

Private Sub ucProductos_SePresionoTeclaEspecial(KeyCode As Integer)
     If KeyCode = vbKeyF2 Then
        AdministrarKeyPreview KeyCode
        'Me.KeyPreview = False
     End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Sub LimpiarVariablesDeMemoria()
    Set mo_ReglasImagenes = Nothing
    Set mo_AdminAdmision = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_AdminCaja = Nothing
    Set mo_ReglasComunes = Nothing
    Set mo_ReglasSeguridad = Nothing
    Set mo_AdminArchivoClinico = Nothing
    Set mo_Apariencia = Nothing
    Set mo_cmbIdEstado = Nothing
    Set mo_cmbResponsable = Nothing
    Set lcBuscaParametro = Nothing
    Set mo_Teclado = Nothing
    Set mo_Formulario = Nothing
    Set oDOPaciente = Nothing
End Sub

Private Sub txtNreceta_LostFocus()
    If Val(txtNreceta.Text) > 0 Then
       Dim lcSql As String
       Dim oRsTmp1 As New Recordset, lnRecetaProcesada As Long, lnCuenta As Long
       lnRecetaProcesada = Val(txtNreceta.Text)
       '
       Set oRsTmp1 = mo_ReglasComunes.RecetasConCabeceraYdetalleSoloCpt(lnRecetaProcesada, sghRecetaEstados.sighRecetaRegistrada)
       If oRsTmp1.RecordCount > 0 Then
            If oRsTmp1.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                mo_ReglasComunes.RecetaChequeaEstadoActual oRsTmp1.Fields!idCuentaAtencion, _
                                                           oRsTmp1.Fields!idEstado, _
                                                           0, oRsTmp1.Fields!DocumentoDespacho
                txtNreceta.Text = ""
            Else
                If oRsTmp1.Fields!idPuntoCarga <> ml_PuntoCarga Then
                     MsgBox "Esa receta no es de PUNTO CARGA", vbInformation, "Imágenes"
                     txtNreceta.Text = ""
                Else
                     txtNcuenta.Text = oRsTmp1.Fields!idCuentaAtencion
                     txtNcuenta_LostFocus
                     ucFacturacionProductos.LimpiarGrilla
                     ucFacturacionProductos.PermiteAgregarItems = False
                     ucFacturacionProductos.CargaProductosPorIdReceta oRsTmp1
                     
                     lnIdReceta = lnRecetaProcesada
                     On Error Resume Next
                     ucFacturacionProductos.SetFocus
                End If
            End If
       Else
            MsgBox "Ese N° Receta NO EXISTE", vbInformation, "Caja"
            txtNreceta.Text = ""
       End If
       oRsTmp1.Close
       Set oRsTmp1 = Nothing
    End If
End Sub


Private Sub txtNreceta_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtNreceta
       AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbBuscaReceta_Click()
    Dim oBusqueda As New SIGHNegocios.clBuscaReceta
    oBusqueda.idPuntoCarga = ml_PuntoCarga
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
       txtNreceta.Text = oBusqueda.IdRecetaSeleccionada
       txtNreceta_LostFocus
    End If
    Set oBusqueda = Nothing
End Sub

Private Sub btnImprimir_Click()
   Dim oRptCaja As New RptCaja
   oRptCaja.ImprimeCitas DoSiCitas, sghImageneología, lcHistoria, lcFinanciamiento, cmbResponsable.Text
   Set oRptCaja = Nothing
End Sub



