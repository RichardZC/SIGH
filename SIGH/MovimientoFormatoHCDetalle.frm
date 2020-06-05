VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form MovimientoFormatoHCDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11745
   Icon            =   "MovimientoFormatoHCDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2685
      Left            =   30
      TabIndex        =   43
      Top             =   1830
      Width           =   11745
      Begin VB.CheckBox chkServiciosTodos 
         Caption         =   "Todos/Ninguno"
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
         Left            =   120
         TabIndex        =   10
         Top             =   2340
         Width           =   1755
      End
      Begin UltraGrid.SSUltraGrid grdHistoriasSeleccionadas 
         Height          =   2055
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   11565
         _ExtentX        =   20399
         _ExtentY        =   3625
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Historias seleccionadas"
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   4560
      TabIndex        =   41
      Top             =   30
      Width           =   7185
      Begin VB.Frame frmFiltro2 
         Height          =   1035
         Left            =   90
         TabIndex        =   46
         Top             =   600
         Visible         =   0   'False
         Width           =   7005
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
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "Solo se muestran los servicios que corresponden al archivero"
            Top             =   180
            Width           =   5400
         End
         Begin VB.ComboBox cmbCondicionFechas 
            Height          =   315
            ItemData        =   "MovimientoFormatoHCDetalle.frx":0CCA
            Left            =   1920
            List            =   "MovimientoFormatoHCDetalle.frx":0CDD
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   570
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.ComboBox cmbFecha 
            BackColor       =   &H00E0E0E0&
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
            ItemData        =   "MovimientoFormatoHCDetalle.frx":0CFF
            Left            =   120
            List            =   "MovimientoFormatoHCDetalle.frx":0D0C
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   570
            Visible         =   0   'False
            Width           =   1755
         End
         Begin MSMask.MaskEdBox txtFechaDesde 
            Height          =   315
            Left            =   3480
            TabIndex        =   7
            Top             =   570
            Visible         =   0   'False
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MSMask.MaskEdBox txtfechaHasta 
            Height          =   315
            Left            =   5490
            TabIndex        =   8
            Top             =   540
            Visible         =   0   'False
            Width           =   1380
            _ExtentX        =   2434
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
         Begin VB.Label Label10 
            Caption         =   "Servicio destino"
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
            Left            =   150
            TabIndex        =   48
            Top             =   225
            Width           =   1365
         End
         Begin VB.Label lblHasta 
            Caption         =   "Hasta"
            Height          =   315
            Left            =   4980
            TabIndex        =   47
            Top             =   570
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.CommandButton btnBuscarPaciente 
         Caption         =   "..."
         Height          =   315
         Left            =   2070
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtIdHistoriaClinica 
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
         Left            =   990
         TabIndex        =   1
         Top             =   270
         Width           =   1050
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   5670
         Picture         =   "MovimientoFormatoHCDetalle.frx":0D3C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Haga click en este botón para filtrar las historias solicitadas"
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label lblApellidos 
         Caption         =   "..."
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
         TabIndex        =   45
         Top             =   300
         Width           =   3090
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Historia Clínica"
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
         TabIndex        =   44
         Top             =   300
         Width           =   885
      End
   End
   Begin VB.Frame fraDestino 
      Height          =   435
      Left            =   5310
      TabIndex        =   35
      Top             =   60
      Visible         =   0   'False
      Width           =   5205
      Begin VB.CommandButton btnAgregarDx 
         DisabledPicture =   "MovimientoFormatoHCDetalle.frx":3985
         DownPicture     =   "MovimientoFormatoHCDetalle.frx":3D6E
         Height          =   315
         Left            =   3180
         Picture         =   "MovimientoFormatoHCDetalle.frx":417A
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   570
         Width           =   1005
      End
      Begin VB.CommandButton btnQuitarDx 
         DisabledPicture =   "MovimientoFormatoHCDetalle.frx":4586
         DownPicture     =   "MovimientoFormatoHCDetalle.frx":4911
         Height          =   315
         Left            =   4260
         Picture         =   "MovimientoFormatoHCDetalle.frx":4CA4
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   570
         Width           =   1005
      End
      Begin VB.TextBox txtNombreServicioDestino 
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
         Left            =   3180
         TabIndex        =   37
         Top             =   180
         Width           =   5115
      End
      Begin VB.TextBox txtIdServicioDestino 
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
         Left            =   1590
         TabIndex        =   36
         Top             =   195
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "Servicio destino"
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
         Left            =   180
         TabIndex        =   40
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.CommandButton btnListarMovimientosAsoc 
      Caption         =   "Exportar a Excel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton btnBuscarRespArchivo 
      Caption         =   "..."
      Height          =   315
      Left            =   2610
      TabIndex        =   12
      Top             =   4725
      Width           =   315
   End
   Begin VB.CommandButton btnBuscarRespTransporte 
      Caption         =   "..."
      Height          =   315
      Left            =   2610
      TabIndex        =   33
      Top             =   5115
      Width           =   315
   End
   Begin VB.CommandButton btnBuscarRespRecepcion 
      Caption         =   "..."
      Height          =   315
      Left            =   2610
      TabIndex        =   19
      Top             =   5490
      Width           =   315
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   0
      TabIndex        =   30
      Top             =   5940
      Width           =   11730
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "MovimientoFormatoHCDetalle.frx":5035
         DownPicture     =   "MovimientoFormatoHCDetalle.frx":54F9
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
         Left            =   5962
         Picture         =   "MovimientoFormatoHCDetalle.frx":59E5
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "MovimientoFormatoHCDetalle.frx":5ED1
         DownPicture     =   "MovimientoFormatoHCDetalle.frx":6331
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
         Left            =   4417
         Picture         =   "MovimientoFormatoHCDetalle.frx":67A6
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1365
      End
      Begin SISGalenPlus.XP_ProgressBar progressRpt 
         Height          =   300
         Left            =   90
         TabIndex        =   49
         Top             =   300
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
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
   End
   Begin VB.Frame fraMovimiento 
      Height          =   1410
      Left            =   0
      TabIndex        =   26
      Top             =   4530
      Width           =   11760
      Begin VB.TextBox txtObservacion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   7905
         TabIndex        =   18
         Top             =   585
         Width           =   3780
      End
      Begin VB.TextBox txtNombreEmpleadoRecepcion 
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
         Left            =   3015
         TabIndex        =   20
         Top             =   975
         Width           =   3645
      End
      Begin VB.TextBox txtIdEmpleadoRecepcion 
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
         Left            =   1545
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtNombreEmpleadoTransporte 
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
         Left            =   3015
         TabIndex        =   17
         Top             =   585
         Width           =   3645
      End
      Begin VB.TextBox txtIdEmpleadoTransporte 
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
         Left            =   1545
         TabIndex        =   24
         Top             =   570
         Width           =   975
      End
      Begin VB.TextBox txtNombreEmpleadoArchivo 
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
         Left            =   3015
         TabIndex        =   13
         Top             =   210
         Width           =   3645
      End
      Begin VB.TextBox txtIdEmpleadoArchivo 
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
         Left            =   1545
         TabIndex        =   11
         Top             =   195
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtHoraMovimiento 
         Height          =   315
         Left            =   9360
         TabIndex        =   15
         Top             =   225
         Width           =   780
         _ExtentX        =   1376
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
      Begin MSMask.MaskEdBox txtFechaMovimiento 
         Height          =   315
         Left            =   7905
         TabIndex        =   14
         Top             =   225
         Width           =   1380
         _ExtentX        =   2434
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
      Begin VB.Label Label2 
         Caption         =   "Fecha y Hora"
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
         Left            =   6735
         TabIndex        =   32
         Top             =   255
         Width           =   1110
      End
      Begin VB.Label Label6 
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6735
         TabIndex        =   31
         Top             =   615
         Width           =   1065
      End
      Begin VB.Label Label8 
         Caption         =   "Resp. recepción"
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
         Left            =   165
         TabIndex        =   29
         Top             =   990
         Width           =   1380
      End
      Begin VB.Label Label7 
         Caption         =   "Resp. transporte"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   150
         TabIndex        =   28
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "Resp. salida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   150
         TabIndex        =   27
         Top             =   225
         Width           =   1230
      End
   End
   Begin VB.Frame fraPaciente 
      Height          =   1755
      Left            =   30
      TabIndex        =   25
      Top             =   30
      Width           =   4515
      Begin VB.ComboBox cmbIdMotivo 
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
         Left            =   780
         TabIndex        =   0
         ToolTipText     =   "Seleccionar el motivo por el cual se va a realizar el movimiento"
         Top             =   180
         Width           =   3690
      End
      Begin VB.Label lblArchivero 
         Caption         =   "Archivero: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   180
         TabIndex        =   42
         Top             =   540
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "Motivo"
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
         Left            =   180
         TabIndex        =   34
         Top             =   210
         Width           =   645
      End
   End
End
Attribute VB_Name = "MovimientoFormatoHCDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Movimiento de Formato de Historias
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim mo_MovimientosHistoriaClinica As New DOMovimientoFormatoHC
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim mrs_HistoriasPorMover As New ADODB.Recordset
Dim mo_Movimientos As New Collection
Dim mo_cmbIdMotivo As New sighEntidades.ListaDespleglable
Dim mo_cmbIdServicio As New sighEntidades.ListaDespleglable
Dim ml_IdMovimiento As Long
Dim ml_IdGrupoMovimiento As Long
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_Movimiento As DOMovimientoFormatoHC
Dim ml_IdPaciente As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lcHCyPaciente As String
Dim lnUsuarioFiltroCombo As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lnIdArchivoClinico As Long
Dim lnIdAtencion As Long
Dim lcSql As String

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

       mo_cmbIdMotivo.BoundColumn = "IdMotivo"
       mo_cmbIdMotivo.ListField = "DescripcionLarga"
       Set mo_cmbIdMotivo.RowSource = mo_AdminArchivoClinico.MotivosMovimientoHistoriaSeleccionarTodos()
       mo_cmbIdMotivo.BoundText = "9"
       
       sMensaje = mo_AdminArchivoClinico.MensajeError
       If sMensaje <> "" Then
           MsgBox mo_AdminArchivoClinico.MensajeError, vbInformation, Me.Caption
       End If


        

End Sub
Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
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

Private Sub btnBuscar_Click()
    On Error GoTo ErrBuscar
    If mo_cmbIdMotivo.BoundText = "" Then
        MsgBox "Ingrese el motivo del movimiento", vbInformation, Me.Caption
        Exit Sub
    End If
    LimpiarGrilla
    Dim oRsTmp As New Recordset
    Dim sOperador As String
    Select Case Val(mo_cmbIdMotivo.BoundText)
    Case 9
            If lblApellidos.Caption = "" Then
                MsgBox "Ingrese el N° Historia a Devolver", vbInformation, Me.Caption
                Exit Sub
            End If
            If cmbIdServicio.Text = "" Then
              MsgBox "Debe elegir el Servicio del Hospital", vbInformation, "Mensaje"
              Exit Sub
            End If
            mrs_HistoriasPorMover.AddNew
            mrs_HistoriasPorMover!seleccionar = True
            mrs_HistoriasPorMover!IdHistoriaSolicitada = 0
            mrs_HistoriasPorMover!idPaciente = ml_IdPaciente
            mrs_HistoriasPorMover!HistoriaClinica = Me.txtIdHistoriaClinica.Text
            mrs_HistoriasPorMover!Nombres = lblApellidos.Caption
            mrs_HistoriasPorMover!FechaSolicitud = Date
            mrs_HistoriasPorMover!FechaRequerida = Date
            mrs_HistoriasPorMover!NroFolios = 0
            mrs_HistoriasPorMover!idServicioDestino = lnIdArchivoClinico
            mrs_HistoriasPorMover!nombreServicioDestino = "Archivo Clínico"
            mrs_HistoriasPorMover!IdServicioOrigen = Val(mo_cmbIdServicio.BoundText)
            mrs_HistoriasPorMover!nombreServicioOrigen = cmbIdServicio.Text
            mrs_HistoriasPorMover!idTipoHistoria = 0
            mrs_HistoriasPorMover!IdMovimientoHistoria = 0
            mrs_HistoriasPorMover!IdEstadoregistro = "A"
            mrs_HistoriasPorMover!FormaPago = 0
            mrs_HistoriasPorMover!PagoCita = ""
            mrs_HistoriasPorMover!idAtencion = lnIdAtencion
            mrs_HistoriasPorMover.Update
    Case 4, 5, 6, 7, 8, 10
            If cmbIdServicio.Text = "" Then
              MsgBox "Debe elegir el Servicio del Hospital", vbInformation, "Mensaje"
              Exit Sub
            End If
            If lblApellidos.Caption = "" Then
              MsgBox "Debe registrar el Nro Historia Clinica", vbInformation, "Mensaje"
              Exit Sub
            End If
            mrs_HistoriasPorMover.AddNew
            mrs_HistoriasPorMover!seleccionar = True
            mrs_HistoriasPorMover!IdHistoriaSolicitada = 0
            mrs_HistoriasPorMover!idPaciente = ml_IdPaciente
            mrs_HistoriasPorMover!HistoriaClinica = Me.txtIdHistoriaClinica.Text
            mrs_HistoriasPorMover!Nombres = lblApellidos.Caption
            mrs_HistoriasPorMover!FechaSolicitud = Date
            mrs_HistoriasPorMover!FechaRequerida = Date
            mrs_HistoriasPorMover!NroFolios = 0
            mrs_HistoriasPorMover!idServicioDestino = Val(mo_cmbIdServicio.BoundText)
            mrs_HistoriasPorMover!nombreServicioDestino = cmbIdServicio.Text
            mrs_HistoriasPorMover!IdServicioOrigen = lnIdArchivoClinico
            mrs_HistoriasPorMover!nombreServicioOrigen = "Archivo Clínico"
            mrs_HistoriasPorMover!idTipoHistoria = 0
            mrs_HistoriasPorMover!IdMovimientoHistoria = 0
            mrs_HistoriasPorMover!IdEstadoregistro = "A"
            mrs_HistoriasPorMover!FormaPago = 0
            mrs_HistoriasPorMover!PagoCita = ""
            mrs_HistoriasPorMover!idAtencion = 0
            mrs_HistoriasPorMover.Update
    Case Else
            If Me.cmbCondicionFechas.ListIndex <> 0 Then
                If Me.txtFechaDesde = sighEntidades.FECHA_VACIA_DMY Then
                    MsgBox "Ingrese la Fecha Desde", vbInformation, Me.Caption
                End If
            End If
            If Me.cmbCondicionFechas.ListIndex = 4 Then
                If Me.txtFechaHasta = sighEntidades.FECHA_VACIA_DMY Then
                    MsgBox "Ingrese la Fecha Hasta", vbInformation, Me.Caption
                End If
            Else
                Me.txtFechaHasta = sighEntidades.FECHA_VACIA_DMY
            End If
            sOperador = Trim(cmbCondicionFechas.List(cmbCondicionFechas.ListIndex))
            Set oRsTmp = mo_AdminArchivoClinico.HistoriasSolicitadasSeleccionarXmotivosYfechas(mo_cmbIdMotivo.BoundText, sOperador, txtFechaDesde.Text, txtFechaHasta.Text, Val(txtIdHistoriaClinica.Text), Val(mo_cmbIdServicio.BoundText))
            If oRsTmp.RecordCount > 0 Then
               oRsTmp.MoveFirst
               Do While Not oRsTmp.EOF
                    mrs_HistoriasPorMover.AddNew
                    mrs_HistoriasPorMover!seleccionar = True
                    mrs_HistoriasPorMover!IdHistoriaSolicitada = 0
                    mrs_HistoriasPorMover!idPaciente = oRsTmp.Fields!idPaciente
                    mrs_HistoriasPorMover!HistoriaClinica = oRsTmp.Fields!NroHistoriaClinica
                    mrs_HistoriasPorMover!Nombres = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!PrimerNombre
                    mrs_HistoriasPorMover!FechaSolicitud = Date
                    mrs_HistoriasPorMover!FechaRequerida = oRsTmp.Fields!FechaIngreso
                    mrs_HistoriasPorMover!NroFolios = 0
                    mrs_HistoriasPorMover!idServicioDestino = oRsTmp.Fields!IdServicioIngreso
                    mrs_HistoriasPorMover!nombreServicioDestino = oRsTmp.Fields!nombre
                    mrs_HistoriasPorMover!IdServicioOrigen = lnIdArchivoClinico
                    mrs_HistoriasPorMover!nombreServicioOrigen = "Archivo Clínico"
                    mrs_HistoriasPorMover!idTipoHistoria = 0
                    mrs_HistoriasPorMover!IdMovimientoHistoria = 0
                    mrs_HistoriasPorMover!IdEstadoregistro = "A"
                    mrs_HistoriasPorMover!FormaPago = 0
                    mrs_HistoriasPorMover!PagoCita = ""
                    mrs_HistoriasPorMover!idAtencion = 0
                    mrs_HistoriasPorMover.Update
                    lcSql = oRsTmp.Fields!NroHistoriaClinica
                    Do While Not oRsTmp.EOF And lcSql = oRsTmp.Fields!NroHistoriaClinica
                       oRsTmp.MoveNext
                       If oRsTmp.EOF Then
                          Exit Do
                       End If
                    Loop
               Loop
            End If
            
    End Select
    Exit Sub
ErrBuscar:
    MsgBox Err.Description
End Sub

Sub LimpiarGrilla()

    
        If mrs_HistoriasPorMover Is Nothing Then
            Exit Sub
        End If

        'Set grdHistoriasSeleccionadas.DataSource = Nothing

        If mrs_HistoriasPorMover.RecordCount > 0 Then
            mrs_HistoriasPorMover.MoveFirst
            Do While Not mrs_HistoriasPorMover.EOF
                mrs_HistoriasPorMover.Delete
                mrs_HistoriasPorMover.Update
                mrs_HistoriasPorMover.MoveNext
            Loop
        End If
End Sub

Private Sub btnBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode

End Sub

Private Sub btnBuscarPaciente_Click()
Dim oBusqueda As New SIGHNegocios.BuscaPacientes
Dim oDOPaciente As New doPaciente
Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarConHistoriasDefinitivas
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            ml_IdPaciente = oDOPaciente.idPaciente
            Me.txtIdHistoriaClinica.Text = oDOPaciente.NroHistoriaClinica
            lblApellidos.Caption = Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + Trim(oDOPaciente.PrimerNombre)
            AsignaUltimoServicio ml_IdPaciente
            btnBuscar.SetFocus
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub btnBuscarRespArchivo_Click()
    CompletarDatosResponsable Me.txtIdEmpleadoArchivo, Me.txtNombreEmpleadoArchivo
End Sub

Private Sub btnBuscarRespRecepcion_Click()
    CompletarDatosResponsable Me.txtIdEmpleadoRecepcion, Me.txtNombreEmpleadoRecepcion
End Sub

Private Sub btnBuscarRespTransporte_Click()
    CompletarDatosResponsable Me.txtIdEmpleadoTransporte, Me.txtNombreEmpleadoTransporte
End Sub


Private Sub AsignaUltimoServicio(lnIdPaciente As Long)
        Dim oRsHCtieneMov As New Recordset
        Set oRsHCtieneMov = mo_AdminArchivoClinico.HistoriaUltimoServicioDondeEstubo(lnIdPaciente)
        If oRsHCtieneMov.RecordCount > 0 Then
           If Not IsNull(oRsHCtieneMov.Fields!IdServicioEgreso) Then
              mo_cmbIdServicio.BoundText = oRsHCtieneMov.Fields!IdServicioEgreso
              lnIdAtencion = oRsHCtieneMov.Fields!idAtencion
           End If
           If oRsHCtieneMov.Fields!idTipoServicio > 1 And IsNull(oRsHCtieneMov.Fields!fechaEgreso) Then
              MsgBox "No tiene Fecha de ALTA MEDICA", vbInformation, "Mensaje"
           End If
        End If
        oRsHCtieneMov.Close
        Set oRsHCtieneMov = Nothing
End Sub



Private Sub btnListarMovimientosAsoc_Click()
'Dim oRptMovimiento As New RptMovimientoHistorias
Dim oRptMovimiento As New SIGHReportes.clMovimientoHist

    oRptMovimiento.IdGrupoMovimiento = ml_IdGrupoMovimiento
    'Set oRptMovimiento.progressRpt = Me.progressRpt
    oRptMovimiento.CrearReporteMovimientoHistoria Me.hwnd
    
End Sub

Private Sub chkServiciosTodos_Click()
    mrs_HistoriasPorMover.MoveFirst
    Do While Not mrs_HistoriasPorMover.EOF
        If chkServiciosTodos.Value = 1 Then
           mrs_HistoriasPorMover.Fields!seleccionar = 1
        Else
           mrs_HistoriasPorMover.Fields!seleccionar = 0
        End If
        mrs_HistoriasPorMover.Update
        mrs_HistoriasPorMover.MoveNext
    Loop
End Sub

Private Sub chkServiciosTodos_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbCondicionFechas_Click()
    If Me.cmbCondicionFechas.ListIndex = 4 Then
        Me.lblHasta.Visible = True
        Me.txtFechaHasta.Visible = True
        Me.txtFechaHasta.Text = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
    Else
        Me.lblHasta.Visible = False
        Me.txtFechaHasta.Visible = False
        Me.txtFechaHasta.Text = sighEntidades.FECHA_VACIA_DMY
    End If
End Sub

Private Sub cmbIdMotivo_Click()
    
    Me.txtIdServicioDestino.Tag = ""
    Me.txtIdServicioDestino.Text = ""
    Me.txtNombreServicioDestino = ""
    
    fraDestino.Visible = False
    '
    LimpiarGrilla
    Select Case mi_Opcion
    Case sghAgregar
        fraFiltro.Visible = True
        Me.cmbCondicionFechas.ListIndex = 1
        txtFechaDesde.Text = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
        txtFechaHasta.Text = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
    Case Else
        fraFiltro.Visible = False
    End Select
    frmFiltro2.Visible = True
    Me.cmbCondicionFechas.Visible = True
    cmbFecha.Visible = True
    txtFechaDesde.Visible = True
    Label10.Caption = "Servicio Destino"
    Select Case Val(mo_cmbIdMotivo.BoundText)
    Case 9
        Label10.Caption = "Servicio Origen"
        mo_cmbIdServicio.BoundColumn = "IdServicio"
        mo_cmbIdServicio.ListField = "DescripcionLarga"
        Set mo_cmbIdServicio.RowSource = mo_AdminArchivoClinico.ServiciosPorArchiveroTipoServicio(lnUsuarioFiltroCombo, "")
        Me.cmbCondicionFechas.Visible = False
        cmbFecha.Visible = False
        txtFechaDesde.Visible = False
    Case Else
        Select Case Val(mo_cmbIdMotivo.BoundText)
        Case 1      'CE
            mo_cmbIdServicio.BoundColumn = "IdServicio"
            mo_cmbIdServicio.ListField = "DescripcionLarga"
            Set mo_cmbIdServicio.RowSource = mo_AdminArchivoClinico.ServiciosPorArchiveroTipoServicio(lnUsuarioFiltroCombo, "(1)")
            Me.cmbCondicionFechas.ListIndex = 1
            txtFechaDesde.Text = Date
        Case 2      'Hospitalizacion
            mo_cmbIdServicio.BoundColumn = "IdServicio"
            mo_cmbIdServicio.ListField = "DescripcionLarga"
            Set mo_cmbIdServicio.RowSource = mo_AdminArchivoClinico.ServiciosPorArchiveroTipoServicio(lnUsuarioFiltroCombo, "(3)")
            Me.cmbCondicionFechas.ListIndex = 4
            txtFechaDesde.Text = DateAdd("m", -2, Date)
            txtFechaHasta.Text = Date
        Case 3      'Emergencia
            mo_cmbIdServicio.BoundColumn = "IdServicio"
            mo_cmbIdServicio.ListField = "DescripcionLarga"
            Set mo_cmbIdServicio.RowSource = mo_AdminArchivoClinico.ServiciosPorArchiveroTipoServicio(lnUsuarioFiltroCombo, "(2,4)")
            Me.cmbCondicionFechas.ListIndex = 4
            txtFechaDesde.Text = DateAdd("m", -2, Date)
            txtFechaHasta.Text = Date
        Case Else
            mo_cmbIdServicio.BoundColumn = "IdServicio"
            mo_cmbIdServicio.ListField = "DescripcionLarga"
            Set mo_cmbIdServicio.RowSource = mo_AdminArchivoClinico.ServiciosPorArchiveroTipoServicio(lnUsuarioFiltroCombo, "")
        End Select
    End Select
            
    
End Sub
Sub LlenarGrilladeHistoriasSeleccionadas(rsSolicitudes As Recordset, idMotivoMovimiento As Long)
        Dim oRsCitaPagada As New ADODB.Recordset
        Dim oRsHCtieneMov As New ADODB.Recordset
        Dim lcSql As String
        Dim lcFormaPago As String
        Dim lnIdNroHistoria As Long
        Dim lbContinua As Boolean
        If rsSolicitudes.RecordCount > 0 Then
            ms_MensajeError = ""
            rsSolicitudes.MoveFirst
            Do While Not rsSolicitudes.EOF
               lcSql = " ": lcFormaPago = ""
               lbContinua = True
               'Chequea si el ultimo Movimiento corresponde al MOTIVO
               Set oRsHCtieneMov = mo_AdminArchivoClinico.HistoriaUltimoMovimiento(rsSolicitudes!idPaciente)
               If oRsHCtieneMov.RecordCount > 0 Then
                  If idMotivoMovimiento = 9 Then
                     If oRsHCtieneMov.Fields!idMotivo = 9 Then
                        lbContinua = False
                        ms_MensajeError = ms_MensajeError & "La HC: " & Trim(rsSolicitudes!NroHistoriaClinica) & " ya tubo RETORNO el " & oRsHCtieneMov!FechaMovimiento & Chr(13)
                     End If
                  Else
                     If oRsHCtieneMov.Fields!idMotivo <> 9 Then
                        ms_MensajeError = ms_MensajeError & "La HC: " & Trim(rsSolicitudes!NroHistoriaClinica) & " ya tubo SALIDA el " & oRsHCtieneMov!FechaMovimiento & Chr(13)
                        lbContinua = False
                     End If
                  End If
               End If
               oRsHCtieneMov.Close
               '
               If Not IsNull(rsSolicitudes!idAtencion) And lbContinua = True Then
                    Set oRsCitaPagada = mo_AdminArchivoClinico.HistoriaPagoCita(rsSolicitudes!idAtencion, mo_cmbIdMotivo.BoundText)
                     lcSql = " "
                     If oRsCitaPagada.RecordCount > 0 Then
                        If oRsCitaPagada.Fields!idestadofacturacion = 4 Then
                           lcSql = "Si"
                        End If
                        lcFormaPago = oRsCitaPagada!descripcion
                     Else
                        oRsCitaPagada.Close
                        Set oRsCitaPagada = mo_AdminArchivoClinico.HistoriasPagoCitaDescripcionTarifa(rsSolicitudes!idAtencion, mo_cmbIdMotivo.BoundText)
                        lcSql = " "
                        If oRsCitaPagada.RecordCount > 0 Then
                           lcFormaPago = oRsCitaPagada!descripcion
                        Else
                           lbContinua = False
                        End If
                     End If
                     oRsCitaPagada.Close
                End If
                If lbContinua = True Then
                    mrs_HistoriasPorMover.AddNew
                    mrs_HistoriasPorMover!seleccionar = True
                    mrs_HistoriasPorMover!IdHistoriaSolicitada = rsSolicitudes!IdHistoriaSolicitada
                    mrs_HistoriasPorMover!idPaciente = rsSolicitudes!idPaciente
                    mrs_HistoriasPorMover!HistoriaClinica = rsSolicitudes!NroHistoriaClinica
                    mrs_HistoriasPorMover!Nombres = rsSolicitudes!Nombres
                    mrs_HistoriasPorMover!FechaSolicitud = rsSolicitudes!FechaSolicitud
                    mrs_HistoriasPorMover!FechaRequerida = rsSolicitudes!FechaRequerida
                    mrs_HistoriasPorMover!NroFolios = 0
                    mrs_HistoriasPorMover!idServicioDestino = rsSolicitudes!idServicioDestino
                    mrs_HistoriasPorMover!nombreServicioDestino = rsSolicitudes!nombreServicioDestino
                    mrs_HistoriasPorMover!IdServicioOrigen = rsSolicitudes!IdServicioOrigen
                    mrs_HistoriasPorMover!nombreServicioOrigen = rsSolicitudes!nombreServicioOrigen
                    mrs_HistoriasPorMover!idTipoHistoria = rsSolicitudes!idTipoHistoria
                    mrs_HistoriasPorMover!IdMovimientoHistoria = rsSolicitudes!IdMovimientoHistoria
                    mrs_HistoriasPorMover!IdEstadoregistro = "A"
                    mrs_HistoriasPorMover!FormaPago = lcFormaPago
                    mrs_HistoriasPorMover!PagoCita = lcSql
                    mrs_HistoriasPorMover!idAtencion = IIf(IsNull(rsSolicitudes!idAtencion), 0, rsSolicitudes!idAtencion)
                End If
                lnIdNroHistoria = rsSolicitudes!NroHistoriaClinica
                Do While Not rsSolicitudes.EOF And lnIdNroHistoria = rsSolicitudes!NroHistoriaClinica
                   rsSolicitudes.MoveNext
                   If rsSolicitudes.EOF Then
                      Exit Do
                   End If
                Loop
            Loop
            If ms_MensajeError <> "" Then
               MsgBox ms_MensajeError, vbInformation, Me.Caption
               ms_MensajeError = ""
            End If
        ElseIf rsSolicitudes.RecordCount = 0 Then
            MsgBox "No existe solicitud para esta historia clinica", vbInformation, Me.Caption
            Exit Sub
        End If

End Sub


Private Sub cmbIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode

End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdMotivo.MiComboBox = cmbIdMotivo
    Set mo_cmbIdServicio.MiComboBox = cmbIdServicio
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub grdHistoriasSeleccionadas_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)

    If mi_Opcion = sghModificar Then
        If Cell.Column.Key = "Seleccionar" Then
            If NewValue = False Then
                If Cell.Row.Cells("IdMovimientoHistoria").Value <> "" Then
                    'If mo_AdminArchivoClinico.MovimientosHistoriaEsUltimoMovimiento(0, Val(Cell.Row.Cells("IdMovimientoHistoria").Value)) Then
                    '    MsgBox "No se puede eliminar existen movimientos de historias posteriores a este", vbInformation, Me.Caption
                    '    Cancel = True
                    'End If
                End If
            End If
        End If
    End If
End Sub

Private Sub grdHistoriasSeleccionadas_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdHistoriasSeleccionadas.Bands(0).Columns("IdPaciente").Hidden = True
    grdHistoriasSeleccionadas.Bands(0).Columns("IdTipoHistoria").Hidden = True
    grdHistoriasSeleccionadas.Bands(0).Columns("IdHistoriaSolicitada").Hidden = True
    grdHistoriasSeleccionadas.Bands(0).Columns("IdServicioDestino").Hidden = True
    grdHistoriasSeleccionadas.Bands(0).Columns("IdEstadoRegistro").Hidden = True
    grdHistoriasSeleccionadas.Bands(0).Columns("FechaSolicitud").Hidden = True
    
    If Val(mo_cmbIdMotivo.BoundText) = 9 Then
        grdHistoriasSeleccionadas.Bands(0).Columns("FormaPago").Hidden = True
        grdHistoriasSeleccionadas.Bands(0).Columns("PagoCita").Hidden = True
    Else
        grdHistoriasSeleccionadas.Bands(0).Columns("FormaPago").Width = 2000
        grdHistoriasSeleccionadas.Bands(0).Columns("PagoCita").Width = 800
    End If
    
    grdHistoriasSeleccionadas.Bands(0).Columns("HistoriaClinica").Header.Caption = "Nro Historia"
    grdHistoriasSeleccionadas.Bands(0).Columns("HistoriaClinica").Width = 1000
    
    grdHistoriasSeleccionadas.Bands(0).Columns("Nombres").Header.Caption = "Nombres"
    grdHistoriasSeleccionadas.Bands(0).Columns("Nombres").Width = 2500
    
    grdHistoriasSeleccionadas.Bands(0).Columns("Seleccionar").Width = 500
    grdHistoriasSeleccionadas.Bands(0).Columns("IdServicioOrigen").Hidden = True
    
    grdHistoriasSeleccionadas.Bands(0).Columns("NombreServicioOrigen").Header.Caption = "Servicio Origen"
    grdHistoriasSeleccionadas.Bands(0).Columns("NombreServicioOrigen").Width = 3000
    
    grdHistoriasSeleccionadas.Bands(0).Columns("NombreServicioDestino").Header.Caption = "Servicio Destino"
    grdHistoriasSeleccionadas.Bands(0).Columns("NombreServicioDestino").Width = 3000
    
    grdHistoriasSeleccionadas.Bands(0).Columns("IdMovimientoHistoria").Header.Caption = "IdMovimiento"
    grdHistoriasSeleccionadas.Bands(0).Columns("IdMovimientoHistoria").Width = 1500

    grdHistoriasSeleccionadas.Bands(0).Columns("NroFolios").Header.Caption = "N°Folios"
    grdHistoriasSeleccionadas.Bands(0).Columns("NroFolios").Width = 1000
    
    grdHistoriasSeleccionadas.Bands(0).Columns("FechaSolicitud").Header.Caption = "Fecha Solicitud"
    grdHistoriasSeleccionadas.Bands(0).Columns("FechaSolicitud").Width = 1250
    grdHistoriasSeleccionadas.Bands(0).Columns("FechaSolicitud").Format = sighEntidades.DevuelveFechaSoloFormato_DMY_HMS
    
    grdHistoriasSeleccionadas.Bands(0).Columns("FechaRequerida").Header.Caption = "F.Requerida"
    grdHistoriasSeleccionadas.Bands(0).Columns("FechaRequerida").Format = sighEntidades.DevuelveFechaSoloFormato_DMY
    grdHistoriasSeleccionadas.Bands(0).Columns("FechaRequerida").Width = "900"
    
    mo_Apariencia.ConfigurarFilasBiColores grdHistoriasSeleccionadas, sighEntidades.GrillaConFilasBicolor
    
End Sub



Private Sub grdHistoriasSeleccionadas_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)

     If KeyCode = vbKeyF2 Then
       Dim lnKeyCode As Integer
       lnKeyCode = KeyCode
       AdministrarKeyPreview lnKeyCode
     End If
End Sub



Private Sub txtFechaDesde_LostFocus()
If Not EsFecha(txtFechaDesde.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaDesde.Text = sighEntidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtfechaHasta_LostFocus()
If Not EsFecha(txtFechaHasta.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaHasta.Text = sighEntidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtIdHistoriaClinica_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdHistoriaClinica
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdHistoriaClinica_LostFocus()
    Dim oRsBuscar As New ADODB.Recordset
    lblApellidos.Caption = ""
    If Val(txtIdHistoriaClinica.Text) > 0 Then
         Set oRsBuscar = mo_AdminAdmision.PacientesSeleccionarPorNroHistoria(Val(HCigualDNI_AgregaNUEVEaLaHistoria(txtIdHistoriaClinica.Text)))
         If oRsBuscar.RecordCount > 0 Then
            lblApellidos.Caption = Trim(oRsBuscar.Fields!ApellidoPaterno) + " " + Trim(oRsBuscar.Fields!ApellidoMaterno) + " " + Trim(oRsBuscar.Fields!PrimerNombre)
            ml_IdPaciente = oRsBuscar.Fields!idPaciente
         End If
         oRsBuscar.Close
    End If
    Set oRsBuscar = Nothing
    CompletarDatosDeServicioEnElLostFocus txtIdServicioDestino, Me.txtNombreServicioDestino
    mo_Formulario.MarcarComoVacio txtIdHistoriaClinica
    AsignaUltimoServicio ml_IdPaciente
    btnBuscar.SetFocus
End Sub

Private Sub txtIdHistoriaClinica_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdServicioDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdServicioDestino
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdServicioDestino_LostFocus()
    CompletarDatosDeServicioEnElLostFocus txtIdServicioDestino, Me.txtNombreServicioDestino
    mo_Formulario.MarcarComoVacio txtIdServicioDestino
End Sub

Private Sub txtIdServicioDestino_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdServicioDestinoFiltro_LostFocus()
    CompletarDatosDeServicioEnElLostFocus txtIdServicioDestino, Me.txtNombreServicioDestino
    mo_Formulario.MarcarComoVacio txtIdServicioDestino
End Sub

Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtObservacion
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtObservacion_LostFocus()
   mo_Formulario.MarcarComoVacio txtObservacion
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdMotivo
   AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdMotivo_LostFocus()
   If cmbIdMotivo.Text <> "" Then
       mo_cmbIdMotivo.BoundText = Val(Split(cmbIdMotivo.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdMotivo
End Sub

Private Sub cmbIdMotivo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtHoraMovimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraMovimiento
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtHoraMovimiento_LostFocus()
If Not sighEntidades.ValidaHora(txtHoraMovimiento.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
             txtHoraMovimiento.Text = sighEntidades.HORA_VACIA_HM
        End If
End Sub

Private Sub txtHoraMovimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFechaMovimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaMovimiento
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaMovimiento_LostFocus()
If Not EsFecha(txtFechaMovimiento.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaMovimiento.Text = sighEntidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFechaMovimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdEmpleadoRecepcion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdEmpleadoRecepcion
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdEmpleadoRecepcion_LostFocus()
    CompletarDatosDeEmpleadoEnElLostFocus txtIdEmpleadoRecepcion, Me.txtNombreEmpleadoRecepcion
    mo_Formulario.MarcarComoVacio txtIdEmpleadoRecepcion
End Sub

Private Sub txtIdEmpleadoRecepcion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtIdEmpleadoTransporte_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdEmpleadoTransporte
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdEmpleadoTransporte_LostFocus()
    CompletarDatosDeEmpleadoEnElLostFocus txtIdEmpleadoTransporte, Me.txtNombreEmpleadoTransporte
    mo_Formulario.MarcarComoVacio txtIdEmpleadoTransporte
End Sub

Private Sub txtIdEmpleadoTransporte_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtIdEmpleadoArchivo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdEmpleadoArchivo
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdEmpleadoArchivo_LostFocus()
    CompletarDatosDeEmpleadoEnElLostFocus txtIdEmpleadoArchivo, Me.txtNombreEmpleadoArchivo
    mo_Formulario.MarcarComoVacio txtIdEmpleadoArchivo
    
    If Trim(Me.txtIdEmpleadoRecepcion) = "" Then Me.txtIdEmpleadoRecepcion = txtIdEmpleadoArchivo
    If Trim(Me.txtIdEmpleadoTransporte) = "" Then Me.txtIdEmpleadoTransporte = txtIdEmpleadoArchivo
    
    txtIdEmpleadoRecepcion_LostFocus
    txtIdEmpleadoTransporte_LostFocus
    
End Sub

Private Sub txtIdEmpleadoArchivo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

    Dim oDOEmpleado As New dOEmpleado
    Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(ml_idUsuario)
    lblArchivero = "Archivero:" + oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    txtIdEmpleadoArchivo.Tag = oDOEmpleado.IdEmpleado
    txtIdEmpleadoArchivo.Text = oDOEmpleado.CodigoPlanilla
    txtNombreEmpleadoArchivo = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    '
    txtIdEmpleadoRecepcion.Tag = oDOEmpleado.IdEmpleado
    txtIdEmpleadoRecepcion.Text = oDOEmpleado.CodigoPlanilla
    txtNombreEmpleadoRecepcion = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    '
    txtIdEmpleadoTransporte.Tag = oDOEmpleado.IdEmpleado
    txtIdEmpleadoTransporte.Text = oDOEmpleado.CodigoPlanilla
    txtNombreEmpleadoTransporte = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    '
    lnUsuarioFiltroCombo = IIf(lcBuscaParametro.SeleccionaFilaParametro(231) = "S", 0, ml_idUsuario)
    lnIdArchivoClinico = mo_AdminComun.ParametrosIdServicioArchivoClinico()
    '
    mo_Formulario.HabilitarDeshabilitar Me.txtNombreEmpleadoArchivo, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNombreEmpleadoRecepcion, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNombreEmpleadoTransporte, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNombreServicioDestino, False

    Select Case mi_Opcion
        Case sghAgregar
        Case sghModificar
            CargarDatosALosControles2
        Case sghConsultar
            CargarDatosALosControles2
        Case sghEliminar
            CargarDatosALosControles2
    End Select
    
    Select Case mi_Opcion
        Case sghAgregar
            Me.txtFechaMovimiento.Text = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
            Me.txtHoraMovimiento = Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
            Me.btnListarMovimientosAsoc.Visible = False
            Me.progressRpt.Visible = False
        Case sghModificar
            Me.FraPaciente.Enabled = False
            Me.txtFechaMovimiento.Enabled = False
            Me.txtHoraMovimiento.Enabled = False
        Case sghConsultar
            Me.FraPaciente.Enabled = False
            Me.fraMovimiento.Enabled = False
            Me.btnAceptar.Enabled = False
        Case sghEliminar
            Me.FraPaciente.Enabled = False
            Me.fraMovimiento.Enabled = False
    End Select
    
    
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       
       GenerarRecordsetTemporal
       cmbFecha.ListIndex = 1
        
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar movimiento de FORMATO de historia clínica"
       Case sghModificar
           Me.Caption = "Modificar movimiento de FORMATO de  historia clínica"
       Case sghConsultar
           Me.Caption = "Consultar movimiento de FORMATO de  historia clínica"
       Case sghEliminar
           Me.Caption = "Eliminar movimiento de FORMATO de  historia clínica"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
           LimpiarVariablesDeMemoria
       End If
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       Case vbKeyF6
           btnBuscar_Click
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox " Los datos se agregaron exitosamente", vbInformation, Me.Caption
                   LimpiarFormulario
                   txtIdHistoriaClinica.SetFocus
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
   LimpiarVariablesDeMemoria
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   Dim lbExisteH As Boolean
   ValidarDatosObligatorios = False
   
   If mo_cmbIdMotivo.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese el motivo" + Chr(13)
   End If
   If Me.txtHoraMovimiento.Text = sighEntidades.HORA_VACIA_HM Then
       sMensaje = sMensaje + "Ingrese la hora de movimiento" + Chr(13)
   End If
   If Me.txtFechaMovimiento.Text = sighEntidades.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Ingrese la fecha de movimiento" + Chr(13)
   End If
   If Me.txtIdEmpleadoRecepcion.Text = "" Then
       sMensaje = sMensaje + "Ingrese el código de planilla del empleado de recepcion" + Chr(13)
   End If
   If Me.txtIdEmpleadoTransporte.Text = "" Then
       sMensaje = sMensaje + "Ingrese el código de planilla del empleado de transporte" + Chr(13)
   End If
   If Me.txtIdEmpleadoArchivo.Text = "" Then
       sMensaje = sMensaje + "Ingrese el código de planilla del empleado de archivo" + Chr(13)
   End If
   '
   If mrs_HistoriasPorMover.RecordCount = 0 Then
       sMensaje = sMensaje + "No hay Historias para Seleccionar" + Chr(13)
   Else
        lbExisteH = False
        mrs_HistoriasPorMover.MoveFirst
        Do While Not mrs_HistoriasPorMover.EOF
             If mrs_HistoriasPorMover!seleccionar Then
                lbExisteH = True
             End If
             mrs_HistoriasPorMover.MoveNext
        Loop
        If lbExisteH = False Then
           sMensaje = sMensaje + "Seleccione al menos una Historia" + Chr(13)
        End If
   End If
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
Dim sMensaje As String

   ValidarReglas = False
   
   sMensaje = ""
   mrs_HistoriasPorMover.MoveFirst
   Do While Not mrs_HistoriasPorMover.EOF
        Select Case Val(mrs_HistoriasPorMover!idTipoHistoria)
        Case 3
            If Val(mrs_HistoriasPorMover!NroFolios) = 0 Then
                sMensaje = sMensaje + "La historia clínica : " & mrs_HistoriasPorMover!HistoriaClinica & " es ESPECIAL." + Chr(13)
            End If
        Case 4
            If Val(mrs_HistoriasPorMover!NroFolios) = 0 Then
                sMensaje = sMensaje + "La historia clínica : " & mrs_HistoriasPorMover!HistoriaClinica & " es JUDICIAL." + Chr(13)
            End If
        End Select
        mrs_HistoriasPorMover.MoveNext
  Loop
   
    If sMensaje <> "" Then
        MsgBox sMensaje + Chr(13) + "Por favor ingresar el nro de folios correspondientes", vbInformation, Me.Caption
        Exit Function
    End If
  ValidarReglas = True

End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS MOVIMIENTOS
    '---------------------------------------------------------------------------------
    If Not (mrs_HistoriasPorMover.BOF And mrs_HistoriasPorMover.EOF) Then
            Set mo_Movimiento = New DOMovimientoFormatoHC
            With mo_Movimiento
                .Observacion = Me.txtObservacion.Text
                .idMotivo = mo_cmbIdMotivo.BoundText
                .FechaMovimiento = Format(Me.txtFechaMovimiento.Text + " " + Me.txtHoraMovimiento.Text, sighEntidades.DevuelveFechaSoloFormato_DMY_HM)
                .IdEmpleadoRecepcion = Val(Me.txtIdEmpleadoRecepcion.Tag)
                .IdEmpleadoTransporte = Val(Me.txtIdEmpleadoTransporte.Tag)
                .IdEmpleadoArchivo = Val(Me.txtIdEmpleadoArchivo.Tag)
                .IdGrupoMovimiento = ml_IdGrupoMovimiento
                .IdUsuarioAuditoria = Val(Me.txtIdEmpleadoArchivo.Tag)
                .IdMovimiento = Me.IdMovimiento
            End With
    End If
    lcHCyPaciente = ""
    If mrs_HistoriasPorMover.RecordCount > 0 Then
       mrs_HistoriasPorMover.MoveFirst
       mrs_HistoriasPorMover.Find "seleccionar=1"
       If mrs_HistoriasPorMover.EOF Then
          lcHCyPaciente = Trim(Str(mrs_HistoriasPorMover.Fields!HistoriaClinica)) & " " & mrs_HistoriasPorMover.Fields!Nombres
       End If
    End If
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminArchivoClinico.MovimientosFormatosHCAgregar(mo_Movimiento, mrs_HistoriasPorMover, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcHCyPaciente)

End Function

Function DevolverHistoria_() As Boolean
   
   
   Dim oMovimiento As DOMovimientoHistoriaClinica
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS MOVIMIENTOS
    '---------------------------------------------------------------------------------
    Set mo_Movimientos = New Collection
    Set oMovimiento = New DOMovimientoHistoriaClinica
    With oMovimiento
        .IdMovimiento = 0
        .NroFolios = 0
        .idServicioDestino = mo_AdminComun.ParametrosIdServicioArchivoClinico()
        .IdServicioOrigen = Val(Me.txtIdServicioDestino.Tag)
        .Observacion = "Devolución al archivo"
        .idMotivo = 9
        .FechaMovimiento = Format(Now, sighEntidades.DevuelveFechaSoloFormato_DMY_HM)
        .idPaciente = mrs_HistoriasPorMover!idPaciente
        .IdEmpleadoRecepcion = ml_idUsuario
        .IdEmpleadoTransporte = 0
        .IdEmpleadoArchivo = ml_idUsuario
        .IdHistoriaSolicitada = IIf(IsNull(mrs_HistoriasPorMover!IdHistoriaSolicitada), 0, mrs_HistoriasPorMover!IdHistoriaSolicitada)
        .IdGrupoMovimiento = 0
    End With
    mo_Movimientos.Add oMovimiento
    DevolverHistoria_ = mo_AdminArchivoClinico.MovimientosHistoriaClinicaAgregar(mo_Movimiento, mrs_HistoriasPorMover, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcHCyPaciente)


End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminArchivoClinico.MovimientosFormatosHCModificar(mo_Movimiento, mrs_HistoriasPorMover, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcHCyPaciente)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminArchivoClinico.MovimientosFormatosHCEliminar(mo_Movimiento, mrs_HistoriasPorMover, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcHCyPaciente)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlosControles()
Dim oDoServicio As New doServicio
Dim oConexion As New Connection
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set mo_MovimientosHistoriaClinica = mo_AdminArchivoClinico.MovimientosFormatosHCSeleccionarPorId(Me.IdMovimiento)
        If mo_AdminArchivoClinico.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbInformation, Me.Caption"
             mb_ExistenDatos = False
             Exit Sub
        End If

       If Not mo_MovimientosHistoriaClinica Is Nothing Then
           With mo_MovimientosHistoriaClinica
           
                Me.IdMovimiento = .IdMovimiento
                ml_IdGrupoMovimiento = .IdGrupoMovimiento
                Me.txtObservacion.Text = .Observacion
                mo_cmbIdMotivo.BoundText = .idMotivo
                Me.txtHoraMovimiento.Text = Format(.FechaMovimiento, sighEntidades.DevuelveHoraSoloFormato_HM)
                Me.txtFechaMovimiento.Text = Format(.FechaMovimiento, sighEntidades.DevuelveFechaSoloFormato_DMY)
                
                 'Datos del paciente
                 Dim oDOPaciente As New doPaciente
                 Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(.idPaciente, oConexion)
                 If Not oDOPaciente Is Nothing Then
                     mrs_HistoriasPorMover.AddNew
                     mrs_HistoriasPorMover.Fields!idPaciente = oDOPaciente.idPaciente
                     mrs_HistoriasPorMover.Fields!HistoriaClinica = oDOPaciente.NroHistoriaClinica
                     mrs_HistoriasPorMover.Fields!Nombres = oDOPaciente.ApellidoPaterno + " " + oDOPaciente.ApellidoMaterno + " " + oDOPaciente.PrimerNombre + " " + oDOPaciente.SegundoNombre
                     mrs_HistoriasPorMover.Fields!IdServicioOrigen = .IdServicioOrigen
                     
                    Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(.IdServicioOrigen, oConexion)
                    If Not oDoServicio Is Nothing Then
                        mrs_HistoriasPorMover.Fields!nombreServicioOrigen = oDoServicio.Codigo + " " + oDoServicio.nombre
                    Else
                        mrs_HistoriasPorMover.Fields!nombreServicioOrigen = ""
                    End If
                     
                     mrs_HistoriasPorMover.Fields!NroFolios = .NroFolios
                     
                    mrs_HistoriasPorMover.Fields!IdHistoriaSolicitada = .IdHistoriaSolicitada
                    mrs_HistoriasPorMover.Fields!FechaSolicitud = .FechaMovimiento
                     
                    mrs_HistoriasPorMover!idTipoHistoria = 0
                 
                 End If
                                
                 'Datos del servicio destino
                 Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(.idServicioDestino, oConexion)
                 If Not oDoServicio Is Nothing Then
                     Me.txtIdServicioDestino.Tag = oDoServicio.IdServicio
                     Me.txtIdServicioDestino.Text = oDoServicio.Codigo
                     Me.txtNombreServicioDestino = oDoServicio.nombre
                 Else
                     Me.txtIdServicioDestino.Tag = ""
                     Me.txtIdServicioDestino.Text = ""
                     Me.txtNombreServicioDestino = ""
                 End If
                
                Dim oDOEmpleado As New dOEmpleado
                
                Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(.IdEmpleadoRecepcion)
                If Not oDOEmpleado Is Nothing Then
                    Me.txtIdEmpleadoRecepcion.Tag = oDOEmpleado.IdEmpleado
                    txtIdEmpleadoRecepcion.Text = oDOEmpleado.CodigoPlanilla
                    txtNombreEmpleadoRecepcion = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                End If
                
                Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(.IdEmpleadoTransporte)
                If Not oDOEmpleado Is Nothing Then
                    Me.txtIdEmpleadoTransporte.Tag = oDOEmpleado.IdEmpleado
                    txtIdEmpleadoTransporte.Text = oDOEmpleado.CodigoPlanilla
                    Me.txtNombreEmpleadoTransporte = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                End If
                
                Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(.IdEmpleadoArchivo)
                If Not oDOEmpleado Is Nothing Then
                    Me.txtIdEmpleadoArchivo.Tag = oDOEmpleado.IdEmpleado
                    txtIdEmpleadoArchivo.Text = oDOEmpleado.CodigoPlanilla
                    Me.txtNombreEmpleadoArchivo = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                End If
                 
                If Not mo_AdminArchivoClinico.MovimientosHistoriaEsUltimoMovimiento(oDOPaciente.idPaciente, Me.IdMovimiento) Then
                    Select Case mi_Opcion
                    Case sghModificar
                        MsgBox "No podrá modificar el servicio destino, esto sólo esta permitido si es el último movimiento", vbInformation, Me.Caption
                        Me.txtIdServicioDestino.Enabled = False
                    Case sghEliminar
                        MsgBox "Solo se puede eliminar el último elemento", vbInformation, Me.Caption
                        mi_Opcion = sghConsultar
                    End Select
                End If
                 
                 mb_ExistenDatos = True
           End With
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
   
End Sub

Sub CargarDatosALosControles2()
Dim oDoServicio As New doServicio

    Set mo_MovimientosHistoriaClinica = mo_AdminArchivoClinico.MovimientosFormatosHCSeleccionarPorId(Me.IdMovimiento)
    If mo_AdminArchivoClinico.MensajeError <> "" Then
         MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbInformation, Me.Caption"
         mb_ExistenDatos = False
         Exit Sub
    End If
    
    If Not mo_MovimientosHistoriaClinica Is Nothing Then
             ml_IdGrupoMovimiento = mo_MovimientosHistoriaClinica.IdGrupoMovimiento
             Me.txtObservacion.Text = mo_MovimientosHistoriaClinica.Observacion
             mo_cmbIdMotivo.BoundText = mo_MovimientosHistoriaClinica.idMotivo
             Me.txtHoraMovimiento.Text = Format(mo_MovimientosHistoriaClinica.FechaMovimiento, sighEntidades.DevuelveHoraSoloFormato_HM)
             Me.txtFechaMovimiento.Text = Format(mo_MovimientosHistoriaClinica.FechaMovimiento, sighEntidades.DevuelveFechaSoloFormato_DMY)
    
            Dim oDOEmpleado As New dOEmpleado
            
            Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(mo_MovimientosHistoriaClinica.IdEmpleadoRecepcion)
            If Not oDOEmpleado Is Nothing Then
                Me.txtIdEmpleadoRecepcion.Tag = oDOEmpleado.IdEmpleado
                txtIdEmpleadoRecepcion.Text = oDOEmpleado.CodigoPlanilla
                txtNombreEmpleadoRecepcion = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
            End If
            
            Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(mo_MovimientosHistoriaClinica.IdEmpleadoTransporte)
            If Not oDOEmpleado Is Nothing Then
                Me.txtIdEmpleadoTransporte.Tag = oDOEmpleado.IdEmpleado
                txtIdEmpleadoTransporte.Text = oDOEmpleado.CodigoPlanilla
                Me.txtNombreEmpleadoTransporte = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
            End If
            
            Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(mo_MovimientosHistoriaClinica.IdEmpleadoArchivo)
            If Not oDOEmpleado Is Nothing Then
                Me.txtIdEmpleadoArchivo.Tag = oDOEmpleado.IdEmpleado
                txtIdEmpleadoArchivo.Text = oDOEmpleado.CodigoPlanilla
                Me.txtNombreEmpleadoArchivo = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
            End If
    
            mb_ExistenDatos = True
    Else
        mb_ExistenDatos = False
    End If
    
    'Detalle del movimiento
    Dim oRsCitaPagada As New ADODB.Recordset
    Dim lcSql As String
    Dim lcFormaPago As String
    Dim rsMovimientoDeHistorias As New Recordset
    Set rsMovimientoDeHistorias = mo_AdminArchivoClinico.MovimientosFormatosHCPorIdGrupo(mo_MovimientosHistoriaClinica.IdGrupoMovimiento)
    If mo_AdminArchivoClinico.MensajeError <> "" Then
         MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbInformation, Me.Caption"
         mb_ExistenDatos = False
         Exit Sub
    End If

    Do While Not rsMovimientoDeHistorias.EOF
            lcSql = " ": lcFormaPago = ""
            
            mrs_HistoriasPorMover.AddNew
            mrs_HistoriasPorMover!seleccionar = True
            mrs_HistoriasPorMover!IdHistoriaSolicitada = rsMovimientoDeHistorias!IdHistoriaSolicitada
            mrs_HistoriasPorMover!idPaciente = rsMovimientoDeHistorias!idPaciente
            mrs_HistoriasPorMover!HistoriaClinica = rsMovimientoDeHistorias!NroHistoriaClinica
            mrs_HistoriasPorMover!Nombres = rsMovimientoDeHistorias!Nombres
            mrs_HistoriasPorMover!FechaSolicitud = rsMovimientoDeHistorias!FechaSolicitud
            mrs_HistoriasPorMover!FechaRequerida = rsMovimientoDeHistorias!FechaRequerida
            mrs_HistoriasPorMover!NroFolios = rsMovimientoDeHistorias!NroFolios
            mrs_HistoriasPorMover!idServicioDestino = IIf(IsNull(rsMovimientoDeHistorias!idServicioDestino), 0, rsMovimientoDeHistorias!idServicioDestino)
            mrs_HistoriasPorMover!nombreServicioDestino = IIf(IsNull(rsMovimientoDeHistorias!nombreServicioDestino), "", rsMovimientoDeHistorias!nombreServicioDestino)
            mrs_HistoriasPorMover!IdServicioOrigen = IIf(IsNull(rsMovimientoDeHistorias!IdServicioOrigen), 0, rsMovimientoDeHistorias!IdServicioOrigen)
            mrs_HistoriasPorMover!nombreServicioOrigen = IIf(IsNull(rsMovimientoDeHistorias!nombreServicioOrigen), "", rsMovimientoDeHistorias!nombreServicioOrigen)
            mrs_HistoriasPorMover!idTipoHistoria = rsMovimientoDeHistorias!idTipoHistoria
            mrs_HistoriasPorMover!IdMovimientoHistoria = rsMovimientoDeHistorias!IdMovimientoHistoria
            
            mrs_HistoriasPorMover!IdEstadoregistro = "M"
            mrs_HistoriasPorMover!FormaPago = lcFormaPago
            mrs_HistoriasPorMover!idAtencion = IIf(IsNull(rsMovimientoDeHistorias!idAtencion), 0, rsMovimientoDeHistorias!idAtencion)
            mrs_HistoriasPorMover!PagoCita = lcSql
            rsMovimientoDeHistorias.MoveNext
    Loop
    
   
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla MovimientosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

            Me.IdMovimiento = 0
            Me.txtIdServicioDestino.Text = ""
            Me.txtNombreServicioDestino.Text = ""
            Me.txtObservacion.Text = ""
            Me.txtFechaMovimiento.Text = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
            Me.txtHoraMovimiento.Text = Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
            txtIdHistoriaClinica.Text = ""
            lblApellidos.Caption = ""
            lnIdAtencion = 0
            If mrs_HistoriasPorMover.RecordCount > 0 Then
                mrs_HistoriasPorMover.MoveFirst
                Do While Not mrs_HistoriasPorMover.EOF
                    mrs_HistoriasPorMover.Delete
                    mrs_HistoriasPorMover.Update
                    mrs_HistoriasPorMover.MoveNext
                Loop
            End If
End Sub


Sub GenerarRecordsetTemporal()
    
    With mrs_HistoriasPorMover
        .Fields.Append "Seleccionar", adBoolean
          .Fields.Append "IdHistoriaSolicitada", adInteger, 4, adFldIsNullable
          .Fields.Append "IdPaciente", adInteger
          .Fields.Append "HistoriaClinica", adInteger
          .Fields.Append "Nombres", adVarChar, 255
          .Fields.Append "FormaPago", adVarChar, 100, adFldIsNullable
          .Fields.Append "PagoCita", adVarChar, 5, adFldIsNullable
          .Fields.Append "FechaSolicitud", adVarChar, 255, adFldIsNullable
          .Fields.Append "FechaRequerida", adVarChar, 255, adFldIsNullable
          .Fields.Append "NroFolios", adInteger, 4, adFldIsNullable
          .Fields.Append "IdServicioOrigen", adInteger, , adFldIsNullable
          .Fields.Append "NombreServicioOrigen", adVarChar, 100, adFldIsNullable
          .Fields.Append "IdServicioDestino", adInteger
          .Fields.Append "NombreServicioDestino", adVarChar, 100
          .Fields.Append "IdTipoHistoria", adInteger
          .Fields.Append "IdMovimientoHistoria", adInteger, 4, adFldIsNullable
          .Fields.Append "IdEstadoregistro", adChar, 1
          .Fields.Append "idAtencion", adInteger
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    
    Set Me.grdHistoriasSeleccionadas.DataSource = mrs_HistoriasPorMover
    
End Sub


Private Sub btnAgregarDx_Click()
Dim oDOPaciente As New doPaciente

    Me.txtIdHistoriaClinica = Trim(Me.txtIdHistoriaClinica)
    
    If Me.txtIdHistoriaClinica = "" Then
        MsgBox "Por favor ingresar el Nro de historia clínica", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Val(mo_cmbIdMotivo.BoundText) = 0 Then
        MsgBox "Ingrese el motivo del movimiento", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Select Case Val(mo_cmbIdMotivo.BoundText)
    Case 1, 2, 3, 9, 10 'Consultorios externos, Hospitalizacion, Emergencia,Devolucion archivo
    Case 4, 5, 6, 7, 8 'Investigacion, Docencia, Tramites administrativos, Interconsultas
        If Val(Me.txtIdServicioDestino.Tag) = 0 Then
            MsgBox "Por favor ingresar el servicio destino", vbInformation, Me.Caption
            Exit Sub
        End If
    End Select
    

    Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorHistoriaClinicaDefinitiva(Me.txtIdHistoriaClinica)
    If oDOPaciente.idPaciente = 0 Then
        MsgBox "No existe un paciente con el nro de historia ingresado", vbInformation, Me.Caption
        Exit Sub
    End If
    
    'Verificar si ya existe
    If mrs_HistoriasPorMover.RecordCount > 0 Then
        mrs_HistoriasPorMover.MoveFirst
        Do While Not mrs_HistoriasPorMover.EOF
            If mrs_HistoriasPorMover!HistoriaClinica = Me.txtIdHistoriaClinica Then
                MsgBox "El N° de historia clínica ingresado ya se ha seleccionado", vbInformation, Me.Caption
                Exit Sub
            End If
            mrs_HistoriasPorMover.MoveNext
        Loop
    End If
    
    With mrs_HistoriasPorMover

        Dim lIdHistoriaSolicitada As Long
        Dim daFechaSolicitud As Date
        
        'Valida que exista una solicitud
        Select Case Val(mo_cmbIdMotivo.BoundText)
        Case 1, 2, 3 'Consultorios externos, Hospitalizacion, Emergencia,Devolucion archivo
 
        Case 4, 5, 6, 7, 8, 9, 10  'Investigacion, Docencia, Tramites administrativos, Interconsultas
                    
            'Detalle del movimiento
            Dim rsMovimientoDeHistorias As New Recordset
            Set rsMovimientoDeHistorias = mo_AdminArchivoClinico.MovimientosHistoriasClinicasDetallePorIdPaciente(oDOPaciente.idPaciente)
        
            Do While Not rsMovimientoDeHistorias.EOF
                  mrs_HistoriasPorMover!IdHistoriaSolicitada = rsMovimientoDeHistorias!IdHistoriaSolicitada
                  mrs_HistoriasPorMover!idPaciente = rsMovimientoDeHistorias!idPaciente
                  mrs_HistoriasPorMover!HistoriaClinica = rsMovimientoDeHistorias!HistoriaClinica
                  mrs_HistoriasPorMover!Nombres = rsMovimientoDeHistorias!Nombres
                  mrs_HistoriasPorMover!FechaSolicitud = rsMovimientoDeHistorias!FechaSolicitud
                  mrs_HistoriasPorMover!FechaRequerida = rsMovimientoDeHistorias!FechaRequerida
                  mrs_HistoriasPorMover!NroFolios = rsMovimientoDeHistorias!NroFolios
                  mrs_HistoriasPorMover!idServicioDestino = Val(Me.txtIdServicioDestino.Tag)
                  mrs_HistoriasPorMover!nombreServicioDestino = Me.txtNombreServicioDestino
                  mrs_HistoriasPorMover!IdServicioOrigen = rsMovimientoDeHistorias!IdServicioOrigen
                  mrs_HistoriasPorMover!nombreServicioOrigen = rsMovimientoDeHistorias!nombreServicioOrigen
                  mrs_HistoriasPorMover!idTipoHistoria = rsMovimientoDeHistorias!idTipoHistoria
                  mrs_HistoriasPorMover!IdMovimientoHistoria = rsMovimientoDeHistorias!IdMovimientoHistoria
                  mrs_HistoriasPorMover!IdEstadoregistro = "A"
            Loop
        End Select

        End With
    
    Me.txtIdHistoriaClinica = ""

End Sub

Private Sub btnQuitarDx_Click()
    On Error Resume Next
    With mrs_HistoriasPorMover
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With
End Sub

Sub CompletarDatosResponsable(txtIdResponsable As TextBox, txtNombreResponsable As TextBox)
'Dim oBusqueda As New EmpleadosBusqueda
Dim oBusqueda As New SIGHNegocios.BuscaEmpleados
Dim oDOEmpleado As New dOEmpleado
    oBusqueda.MostrarFormulario
    'oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDOEmpleado Is Nothing Then
            txtIdResponsable.Tag = oDOEmpleado.IdEmpleado
            txtIdResponsable.Text = oDOEmpleado.CodigoPlanilla
            txtNombreResponsable = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
            txtIdResponsable.SetFocus
        End If
    End If

End Sub
Sub CompletarDatosDeServicio(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
Dim oBusqueda As New SIGHNegocios.BuscaServicioHosp
Dim oDoServicio As New doServicio
Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oBusqueda.HabilitarTipoServicio = True
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDoServicio Is Nothing Then
            txtIdServicio.Text = oDoServicio.Codigo
            txtIdServicio.Tag = oDoServicio.IdServicio
            lblDescripcionServicio = oDoServicio.nombre
        Else
            txtIdServicio.Text = ""
            txtIdServicio.Tag = ""
            lblDescripcionServicio = ""
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDoServicio = Nothing

End Sub

Sub CompletarDatosDeEmpleadoEnElLostFocus(txtCodigoPlanilla As TextBox, txtNombre As TextBox)
Dim oDOEmpleado As New dOEmpleado

        If mo_AdminComun.EmpleadosSeleccionarPorCodigo(txtCodigoPlanilla.Text, oDOEmpleado) Then
            txtCodigoPlanilla.Tag = oDOEmpleado.IdEmpleado
            txtCodigoPlanilla.Text = oDOEmpleado.CodigoPlanilla
            txtNombre = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        Else
            txtCodigoPlanilla.Tag = ""
            txtCodigoPlanilla = ""
            txtNombre = ""
        End If
End Sub

Sub CompletarDatosDeServicioEnElLostFocus(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
    
    txtIdServicio.Text = UCase(txtIdServicio.Text)
    If txtIdServicio.Text <> "" Then
        Dim oDoServicio As doServicio
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(txtIdServicio.Text)
        If Not oDoServicio Is Nothing Then
            txtIdServicio.Tag = oDoServicio.IdServicio
            lblDescripcionServicio.Text = oDoServicio.nombre
        Else
            txtIdServicio.Tag = ""
            lblDescripcionServicio.Text = ""
        End If
   End If

End Sub


Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Teclado = Nothing
    Set mo_Formulario = Nothing
    Set mo_MovimientosHistoriaClinica = Nothing
    Set mo_AdminAdmision = Nothing
    Set mo_AdminArchivoClinico = Nothing
    Set mo_AdminServiciosHosp = Nothing
    Set mo_AdminComun = Nothing
    Set mrs_HistoriasPorMover = Nothing
    Set mo_Movimientos = Nothing
    Set mo_cmbIdMotivo = Nothing
    Set mo_cmbIdServicio = Nothing
    Set mo_Apariencia = Nothing
    Set lcBuscaParametro = Nothing
End Sub


