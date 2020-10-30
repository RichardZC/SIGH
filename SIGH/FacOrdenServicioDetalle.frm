VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form FacOrdenServicioDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13860
   Icon            =   "FacOrdenServicioDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   13860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2505
      Left            =   30
      TabIndex        =   20
      Top             =   0
      Width           =   13755
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
         Left            =   6405
         MaxLength       =   30
         TabIndex        =   43
         Top             =   705
         Width           =   1125
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
         Left            =   5790
         MaxLength       =   4
         TabIndex        =   42
         Top             =   705
         Width           =   585
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
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   41
         Top             =   705
         Width           =   1245
      End
      Begin VB.CommandButton cmbBuscaReceta 
         Height          =   330
         Left            =   2625
         Picture         =   "FacOrdenServicioDetalle.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   705
         Width           =   300
      End
      Begin VB.CheckBox chkPlanNoCubre 
         Alignment       =   1  'Right Justify
         Caption         =   "Plan NO cubre"
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
         Left            =   6030
         TabIndex        =   3
         Top             =   1140
         Width           =   1485
      End
      Begin VB.ComboBox cmbServicioIngreso 
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
         Left            =   9330
         TabIndex        =   15
         Top             =   1560
         Width           =   4350
      End
      Begin VB.TextBox txtNroOrdenPago 
         Alignment       =   2  'Center
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
         Left            =   4050
         TabIndex        =   10
         Top             =   270
         Width           =   1095
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
         Left            =   9330
         TabIndex        =   35
         Top             =   1560
         Width           =   4305
      End
      Begin VB.ComboBox cmbIdPuntoDeCarga 
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
         Left            =   9330
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txtIdOrden 
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
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   9
         Top             =   270
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
         Caption         =   "..."
         Height          =   315
         Left            =   2655
         TabIndex        =   1
         ToolTipText     =   "Busca Cuenta por Apellidos y Nombres"
         Top             =   1155
         Width           =   315
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
         Height          =   330
         Left            =   3870
         TabIndex        =   5
         Top             =   1530
         Width           =   3645
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
         Left            =   9330
         TabIndex        =   14
         Top             =   1110
         Width           =   4335
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
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   0
         Top             =   1140
         Width           =   1245
      End
      Begin VB.TextBox txtNombrePaciente 
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
         Height          =   330
         Left            =   3000
         TabIndex        =   2
         Top             =   1140
         Width           =   2865
      End
      Begin VB.TextBox txtNroOrdenSisSoat 
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
         Left            =   6060
         TabIndex        =   8
         Top             =   1950
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame fraTipoVenta 
         Enabled         =   0   'False
         Height          =   525
         Left            =   10575
         TabIndex        =   21
         Top             =   1920
         Visible         =   0   'False
         Width           =   3045
         Begin Threed.SSOption optVentas 
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   180
            Width           =   1785
            _ExtentX        =   3149
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
            Caption         =   "Venta Directa"
            Value           =   -1
         End
         Begin Threed.SSOption optPreventa 
            Height          =   255
            Left            =   4410
            TabIndex        =   23
            Top             =   180
            Width           =   1125
            _ExtentX        =   1984
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
            Caption         =   "PreVenta"
         End
      End
      Begin VB.TextBox txtEstado 
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
         Height          =   315
         Left            =   9330
         MaxLength       =   30
         TabIndex        =   12
         Top             =   330
         Width           =   1185
      End
      Begin MSMask.MaskEdBox txtHentrega 
         Height          =   315
         Left            =   2820
         TabIndex        =   7
         Top             =   2010
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
      Begin MSMask.MaskEdBox txtFentrega 
         Height          =   315
         Left            =   1380
         TabIndex        =   6
         Top             =   2010
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
      Begin MSDataListLib.DataCombo cmbFormaPago 
         Height          =   360
         Left            =   1380
         TabIndex        =   4
         Top             =   1530
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txtFregistro 
         Height          =   315
         Left            =   6180
         TabIndex        =   11
         Top             =   270
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox txtFrealizaCpt 
         Height          =   315
         Left            =   11850
         TabIndex        =   38
         Top             =   330
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
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
         Left            =   4980
         TabIndex        =   45
         Top             =   765
         Width           =   780
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
         Left            =   90
         TabIndex        =   44
         Top             =   765
         Width           =   870
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "F.Realiza CPT"
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
         Left            =   10740
         TabIndex        =   39
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "N° Orden Pago"
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
         Left            =   2790
         TabIndex        =   37
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label Label3 
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
         Left            =   8310
         TabIndex        =   36
         Top             =   1590
         Width           =   990
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "F.Registro"
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
         Left            =   5370
         TabIndex        =   34
         Top             =   330
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N° Orden"
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
         TabIndex        =   33
         Top             =   270
         Width           =   780
      End
      Begin VB.Label lblEstadoOrden 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11490
         TabIndex        =   32
         Top             =   510
         Width           =   1995
      End
      Begin VB.Label Label2 
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
         Left            =   90
         TabIndex        =   31
         Top             =   1575
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "F.Despacho"
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
         TabIndex        =   30
         Top             =   2040
         Width           =   945
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
         Left            =   90
         TabIndex        =   29
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "N° Orden"
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
         Left            =   5250
         TabIndex        =   28
         Top             =   2040
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Venta"
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
         Left            =   9330
         TabIndex        =   27
         Top             =   2130
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label7 
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
         Left            =   7740
         TabIndex        =   26
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pto. Carga"
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
         Left            =   8445
         TabIndex        =   25
         Top             =   780
         Width           =   855
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
         Height          =   210
         Left            =   8700
         TabIndex        =   24
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   19
      Top             =   7650
      Width           =   13710
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FacOrdenServicioDetalle.frx":1254
         DownPicture     =   "FacOrdenServicioDetalle.frx":1718
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
         Picture         =   "FacOrdenServicioDetalle.frx":1C04
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FacOrdenServicioDetalle.frx":20F0
         DownPicture     =   "FacOrdenServicioDetalle.frx":2550
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
         Picture         =   "FacOrdenServicioDetalle.frx":29C5
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1365
      End
   End
   Begin SISGalenPlus.ucFacturacionItems ucFacturacionProductos 
      Height          =   4995
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   13755
      _extentx        =   24262
      _extenty        =   8811
   End
End
Attribute VB_Name = "FacOrdenServicioDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Procedimientos Médicos en un Servicio al Paciente con cuenta
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_PuntoCarga As Long
Dim ml_idOrden As Long
Dim mi_Opcion As sghOpciones
Dim ms_MensajeError As String
Dim ml_idUsuario As Long
Dim mb_ExistenDatos As Boolean
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim wxParametro302 As String, lnIdTipoServicio As Long
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_cmbIdPuntoCarga As New sighEntidades.ListaDespleglable
Dim mo_cmbIdEstado As New sighEntidades.ListaDespleglable
Dim mo_cmbFechaIngreso As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoGenHistoriaClinica As New sighEntidades.ListaDespleglable
Dim mo_cmbServicioIngreso As New sighEntidades.ListaDespleglable
Dim mo_DOFactOrdenServicio As New DoFactOrdenServ
Dim mo_DoAtencion As New DOAtencion
Dim ml_IdPaciente As Long
Dim ml_IdTipoFinanciamiento As Long
Dim ml_IdFuenteFinanciamiento As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim oRsFormaPago As New ADODB.Recordset
Dim lbDocumentoYaRegistradoEnSeguros As Boolean
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim ln_IdOrdenPago As Long
Dim lbPrimeraVez As Boolean
Dim oRsTipoFinanciamiento As New Recordset
Dim lcPosicionDefaultCombo As String
Dim ml_IdFuenteFinanciamientoDespacho As Long
Dim ml_IdServicioPaciente As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_IdPuntoCargaServicioHosp As Long
Dim ml_idCuentaAtencion As Long
Dim lbCargaVentanaDesdeOtraVentana As Boolean
Dim mo_lbNOValidaCodigoPrestacion As Boolean
'mgaray201411a desde donde se esta llamando el formulario
Dim ml_FormMostradoDesde As Long
Dim lnIdReceta As Long, lnIdComprobantePagoDeReceta As Long
Dim lbCuentaDeEmergenciaCerrada As Boolean

Property Let lbNOValidaCodigoPrestacion(lValue As String)
   mo_lbNOValidaCodigoPrestacion = lValue
End Property

Property Let idCuentaAtencion(lValue As Long)
    lbCargaVentanaDesdeOtraVentana = True
    ml_idCuentaAtencion = lValue
    txtNcuenta.Text = ml_idCuentaAtencion
    txtNcuenta_LostFocus
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

Property Let PuntoCarga(lValue As Long)
    ml_PuntoCarga = lValue
End Property

Property Get PuntoCarga() As Long
    PuntoCarga = ml_PuntoCarga
End Property

Property Let idTipoFinanciamiento(lValue As Long)
    ml_IdTipoFinanciamiento = lValue
End Property

Property Get idTipoFinanciamiento() As Long
    idTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property

Property Let IdOrden(lValue As Long)
    ml_idOrden = lValue
End Property

Property Get IdOrden() As Long
    IdOrden = ml_idOrden
End Property
'mgaray201411a
Property Let FormMostradoDesde(lValue As Long)
    ml_FormMostradoDesde = lValue
End Property

Property Get FormMostradoDesde() As Long
    FormMostradoDesde = ml_FormMostradoDesde
End Property


Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If AgregarDatos() Then

                    If optVentas.Value = True Then
                       MsgBox "Se Agregó correctamente la Orden N° " & mo_DOFactOrdenServicio.IdOrden & Chr(13) & _
                       IIf(Val(txtNroOrdenPago.Text) > 0, "  (Orden de Pago N° " & txtNroOrdenPago.Text & ")", ""), vbInformation, Me.Caption
                    Else
                       MsgBox "Se Agregó correctamente la Orden de Pago N° " & DevuelveNroOrdenPago(mo_DOFactOrdenServicio.IdOrden), vbInformation, Me.Caption
                    End If
                    Me.txtIdOrden = mo_DOFactOrdenServicio.IdOrden
                    If txtFentrega.Text <> sighEntidades.FECHA_VACIA_DMY Then
                        LimpiarFormulario
                    Else
                        Me.Visible = False
                        LimpiarVariablesDeMemoria
                    End If
                    If lbCargaVentanaDesdeOtraVentana = True Then
                        Me.Visible = False
                        LimpiarVariablesDeMemoria
                    End If
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

                    If optVentas.Value = True Then
                       MsgBox "Se Modificó correctamente la Orden N° " & mo_DOFactOrdenServicio.IdOrden, vbInformation, Me.Caption
                    Else
                       MsgBox "Se Modificó correctamente la Orden de Pago N° " & txtNroOrdenPago.Text, vbInformation, Me.Caption
                    End If
                    Me.Visible = False
                    LimpiarVariablesDeMemoria
                Else
                    MsgBox "No se pudo modificar los datos" & Chr(13) & ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
            If MsgBox("¿Realmente desea Eliminar?", vbQuestion + vbYesNo, "Estado de Cuenta") = vbNo Then
                 Exit Sub
            End If
           If ValidarReglas() Then
                CargaDatosAlObjetosDeDatos
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                    LimpiarVariablesDeMemoria
                Else
                    MsgBox "No se pudo eliminar los datos"
               End If
           End If
   End Select
        
End Sub

Sub LimpiarFormulario()
    lbCuentaDeEmergenciaCerrada = False
    txtNcuenta.Text = ""
    txtNombrePaciente.Text = ""
    cmbFormaPago.Text = ""
    txtHentrega.Text = lcBuscaParametro.RetornaHoraServidorSQL   'Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
    txtFentrega.Text = lcBuscaParametro.RetornaFechaServidorSQL    'Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
    txtDatosDeCuenta.Text = ""
    txtNroOrdenSisSoat.Text = ""
    txtProcedencia.Text = ""
    txtPlan.Text = ""
    ln_IdOrdenPago = 0
    txtNroOrdenPago.Text = ""
    ml_IdServicioPaciente = 0
    ml_IdPaciente = 0
    ml_IdPuntoCargaServicioHosp = 0
    Me.cmbServicioIngreso.Text = ""
    chkPlanNoCubre.Value = 0
    Me.ucFacturacionProductos.LimpiarGrilla
    txtFrealizaCpt.Text = lcBuscaParametro.RetornaFechaHoraServidorSQL
    txtNcuenta.SetFocus
End Sub

Function ValidarDatosObligatorios() As Boolean
    On Error Resume Next
    ValidarDatosObligatorios = False
    If optVentas.Value = True Then
        If txtDatosDeCuenta.Text = "" Then
           MsgBox "Tiene problemas con el N° Cuenta", vbInformation, Me.Caption
           Exit Function
        ElseIf txtFentrega.Text = sighEntidades.FECHA_VACIA_DMY Or txtHentrega.Text = sighEntidades.HORA_VACIA_HM Then
           MsgBox "Tiene que registrar la Fecha y Hora de despacho", vbInformation, Me.Caption
           Exit Function
        ElseIf cmbServicioIngreso.Text = "" Then
            MsgBox "Tiene que elegir el Servicio donde está el Paciente (procedencia) en el momento del Consumo", vbInformation, Me.Caption
            Exit Function
        End If
    Else
        If cmbFormaPago.Text = "" Then
           MsgBox "Tiene que elegir la Forma de Pago", vbInformation, Me.Caption
           Exit Function
        End If
    End If
    Select Case mi_Opcion
    Case sghAgregar, sghModificar
        Dim rsProductos As Recordset
        Set rsProductos = Me.ucFacturacionProductos.FacturacionProductos
        If Not (rsProductos.EOF And rsProductos.BOF) Then
            rsProductos.MoveFirst
            Do While Not rsProductos.EOF
                If rsProductos!idProducto = 0 Then
                   rsProductos.Delete
                   rsProductos.Update
                ElseIf rsProductos!Cantidad <= 0 Then
                    MsgBox "El producto: " & rsProductos!Codigo & " " & Trim(rsProductos!NombreProducto) & "   Tiene problemas con la Cantidad", vbInformation, Me.Caption
                    Exit Function
                ElseIf rsProductos!PrecioUnitario <= 0 And rsProductos!SeUsaSinPrecio = False Then
                    MsgBox "El producto: " & rsProductos!Codigo & " " & Trim(rsProductos!NombreProducto) & "   Tiene problemas con el Precio", vbInformation, Me.Caption
                    Exit Function
                End If
                rsProductos.MoveNext
            Loop
        End If
    End Select
    ValidarDatosObligatorios = True
End Function

Sub CargaDatosAlObjetosDeDatos()
    Select Case mi_Opcion
    Case sghAgregar
        With mo_DOFactOrdenServicio
             .fechacreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL      'Now
             .idCuentaAtencion = Val(txtNcuenta.Text)
             .idestadofacturacion = sghEstadoFacturacion.sghAtendido
             .IdFuenteFinanciamiento = ml_IdFuenteFinanciamiento
             .idPaciente = ml_IdPaciente
             .idPuntoCarga = Val(mo_cmbIdPuntoCarga.BoundText)
             .idTipoFinanciamiento = Val(cmbFormaPago.BoundText)
             .idUsuario = ml_idUsuario
             .IdUsuarioAuditoria = ml_idUsuario
             .FechaDespacho = .fechacreacion       '(txtFentrega.Text & " " & txtHentrega.Text)
             .IdUsuarioDespacho = ml_idUsuario
             .FechaHoraRealizaCpt = txtFrealizaCpt.Text
        End With
    Case sghModificar
        mo_DOFactOrdenServicio.idTipoFinanciamiento = Val(cmbFormaPago.BoundText)
        mo_DOFactOrdenServicio.IdUsuarioAuditoria = ml_idUsuario
        mo_DOFactOrdenServicio.FechaHoraRealizaCpt = txtFrealizaCpt.Text
        If txtHentrega.Text <> sighEntidades.HORA_VACIA_HM Then
           mo_DOFactOrdenServicio.idestadofacturacion = sghEstadoFacturacion.sghAtendido
           mo_DOFactOrdenServicio.FechaDespacho = CDate(txtFentrega.Text & " " & txtHentrega.Text) 'Now
           mo_DOFactOrdenServicio.IdUsuarioDespacho = ml_idUsuario
        End If
    End Select
End Sub

Function ValidarReglas() As Boolean
   ValidarReglas = False
    

    
   ValidarReglas = True
End Function
Function AgregarDatos() As Boolean

    AgregarDatos = mo_ReglasFacturacion.FactOrdenServicioAgregar(mo_DOFactOrdenServicio, Me.ucFacturacionProductos.FacturacionProductos, _
                   mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtNombrePaciente.Text, Val(mo_cmbServicioIngreso.BoundText), lnIdReceta, lnIdComprobantePagoDeReceta)
    ms_MensajeError = mo_ReglasFacturacion.MensajeError
    ln_IdOrdenPago = Val(mo_ReglasFacturacion.Texto)
    txtNroOrdenPago.Text = ln_IdOrdenPago
    '
    mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar mo_DOFactOrdenServicio.idCuentaAtencion, False, 0
    mo_ReglasSISgalenhos.FuaActualizaDespachosEnServicios mo_DOFactOrdenServicio.idCuentaAtencion, wxParametro302, lnIdTipoServicio, ml_IdFuenteFinanciamiento
End Function

Function ModificarDatos() As Boolean
    ModificarDatos = mo_ReglasFacturacion.FactOrdenServicioModificar(mo_DOFactOrdenServicio, Me.ucFacturacionProductos.FacturacionProductos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtNombrePaciente.Text)
    ms_MensajeError = mo_ReglasFacturacion.MensajeError
    ln_IdOrdenPago = Val(mo_ReglasFacturacion.Texto)
    '
    mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar mo_DOFactOrdenServicio.idCuentaAtencion, False, 0
    mo_ReglasSISgalenhos.FuaActualizaDespachosEnServicios mo_DOFactOrdenServicio.idCuentaAtencion, wxParametro302, lnIdTipoServicio, ml_IdFuenteFinanciamiento
End Function

Function EliminarDatos() As Boolean
    mo_DOFactOrdenServicio.IdUsuarioAuditoria = ml_idUsuario
    EliminarDatos = mo_ReglasFacturacion.FactOrdenServicioEliminar(mo_DOFactOrdenServicio, mo_lnIdTablaLISTBARITEMS, _
                    mo_lcNombrePc, txtNombrePaciente.Text, lnIdReceta, lnIdComprobantePagoDeReceta)
    '
    mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar mo_DOFactOrdenServicio.idCuentaAtencion, False, 0
    mo_ReglasSISgalenhos.FuaActualizaDespachosEnServicios mo_DOFactOrdenServicio.idCuentaAtencion, wxParametro302, lnIdTipoServicio, ml_IdFuenteFinanciamiento
End Function

Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub

Private Sub chkPlanNoCubre_Click()
    If chkPlanNoCubre.Value = 1 Then
        cmbFormaPago.BoundText = 1
        ml_IdFuenteFinanciamiento = 1 'contado
        ucFacturacionProductos.idTipoFinanciamiento = 1 'contado
        ucFacturacionProductos.LimpiarGrilla
        Me.ucFacturacionProductos.IdOrden = -999
        Me.ucFacturacionProductos.CargaProductosPorIdOrden
        ucFacturacionProductos.TabEnDescripcionParaFactuacion
    Else
       txtNcuenta_LostFocus
       ucFacturacionProductos.LimpiarGrilla
    End If
    
End Sub

Private Sub cmbServicioIngreso_Click()
    If lcBuscaParametro.SeleccionaFilaParametro(246) <> "N" Then
        ml_IdServicioPaciente = Val(mo_cmbServicioIngreso.BoundText)
        Dim oRsTmp As New Recordset
        Set oRsTmp = mo_ReglasComunes.FactPuntosCargaSeleccionarPorFiltro("idServicio=" & Trim(Str(ml_IdServicioPaciente)))
        If oRsTmp.RecordCount > 0 Then
           ml_IdPuntoCargaServicioHosp = oRsTmp.Fields!idPuntoCarga
        End If
        'para poder registrar cualquier CPT en ese SERVICIO
        '(lo registra FACTURACION)
        ucFacturacionProductos.IdPuntoCargaServicioHosp = ml_IdPuntoCargaServicioHosp '0 <--para poder registrar cualquier CPT en ese SERVICIO
        ucFacturacionProductos.TabEnDescripcionParaFactuacion
        oRsTmp.Close
        Set oRsTmp = Nothing
        On Error Resume Next
        ucFacturacionProductos.SetFocus
    End If
End Sub

Private Sub cmbServicioIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbServicioIngreso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmbServicioIngreso_Click
    End If
End Sub

Private Sub Form_Activate()
    If txtNcuenta.Text <> "" And mi_Opcion = sghAgregar And lbDocumentoYaRegistradoEnSeguros = True Then
        txtHentrega.Text = sighEntidades.HORA_VACIA_HM
        txtFentrega.Text = sighEntidades.FECHA_VACIA_DMY
        txtHentrega.Enabled = False
        txtFentrega.Enabled = False
        cmbFormaPago.Enabled = False
        txtNcuenta_LostFocus
    Else
        ucFacturacionProductos.SetFocus
    End If
    If mi_Opcion = sghAgregar Then
       txtNcuenta.SetFocus
    End If
'    CargaCptPorAtencion 'Actualizado 09102014
End Sub

Private Sub Form_Initialize()
      Set mo_cmbServicioIngreso.MiComboBox = cmbServicioIngreso
End Sub

Private Sub Form_Load()
    lbDocumentoYaRegistradoEnSeguros = False
    txtFregistro.Text = lcBuscaParametro.RetornaFechaServidorSQL
    txtEstado.Text = "Registrado"
    txtFrealizaCpt.Text = lcBuscaParametro.RetornaFechaHoraServidorSQL
    
    Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPuntoDeCarga
    
    ConfigurarPuntosDeCarga
    CargaDataCombos
    
    mo_cmbIdPuntoCarga.BoundText = ml_PuntoCarga

    Me.ucFacturacionProductos.idUsuario = ml_idUsuario
    Me.ucFacturacionProductos.Inicializar
    Me.ucFacturacionProductos.idTipoFinanciamiento = ml_IdTipoFinanciamiento
    Me.ucFacturacionProductos.TipoProducto = sghServicio
    Me.ucFacturacionProductos.idPuntoCarga = ml_PuntoCarga
    Me.ucFacturacionProductos.MostrarColumnaLab = GetMostrarLabEnListaProductos()

    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Consumo en el Servicio"
    Case sghModificar
        Me.Caption = "Modificar Consumo en el Servicio"
    Case sghConsultar
        Me.Caption = "Consultar Consumo en el Servicio"
    Case sghEliminar
        Me.Caption = "Eliminar Consumo en el Servicio"
    End Select
    
    CargarDatosAlFormulario
    If mi_Opcion = sghAgregar Then
        ConfiguraPermisosDelUsuario
    End If

End Sub


Sub ConfiguraPermisosDelUsuario()
    Dim ms_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim oRsPermisosUsuario As New Recordset
    Set oRsPermisosUsuario = ms_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_idUsuario)
    optPreventa.Enabled = False
    optVentas.Enabled = False
    If oRsPermisosUsuario.RecordCount > 0 Then
       Do While Not oRsPermisosUsuario.EOF
          Select Case oRsPermisosUsuario.Fields!IdPermiso
          Case 116    'Facturacion - Sólo realiza PreVenta de Servicios
               optPreventa.Enabled = True
               optPreventa.Value = True
          Case 117    'Facturacion - Sólo realiza VentaDirecta de Servicios
               optVentas.Enabled = True
               optVentas.Value = True
          End Select
          oRsPermisosUsuario.MoveNext
       Loop
    End If
    Set oRsPermisosUsuario = Nothing
End Sub

Sub CargarDatosAlFormulario()
 mo_Formulario.HabilitarDeshabilitar Me.txtIdOrden, False
 mo_Formulario.HabilitarDeshabilitar Me.txtFregistro, False
 mo_Formulario.HabilitarDeshabilitar Me.txtEstado, False
 mo_Formulario.HabilitarDeshabilitar Me.cmbIdPuntoDeCarga, False
 mo_Formulario.HabilitarDeshabilitar Me.txtDatosDeCuenta, False
 mo_Formulario.HabilitarDeshabilitar Me.txtPlan, False
 mo_Formulario.HabilitarDeshabilitar Me.txtNombrePaciente, False
 mo_Formulario.HabilitarDeshabilitar Me.txtProcedencia, False
 mo_Formulario.HabilitarDeshabilitar Me.txtNroOrdenPago, False
 mo_Formulario.HabilitarDeshabilitar Me.txtFentrega, False
 mo_Formulario.HabilitarDeshabilitar Me.txtHentrega, False
 wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
 Me.cmbFormaPago.Enabled = False
 If lcBuscaParametro.SeleccionaFilaParametro(246) = "N" Then
    mo_Formulario.HabilitarDeshabilitar Me.cmbServicioIngreso, False
 End If
 If mi_Opcion <> sghAgregar Then
    mo_Formulario.HabilitarDeshabilitar Me.cmbServicioIngreso, False
 End If
 
 Select Case mi_Opcion
     Case sghAgregar
        txtHentrega.Text = lcBuscaParametro.RetornaHoraServidorSQL   'Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
        txtFentrega.Text = lcBuscaParametro.RetornaFechaServidorSQL    'Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
        Me.ucFacturacionProductos.IdOrden = -999
        Me.ucFacturacionProductos.CargaProductosPorIdOrden
     Case sghModificar
        CargarDatosAlosControles
     Case sghConsultar
        CargarDatosAlosControles
     Case sghEliminar
        CargarDatosAlosControles
 End Select
End Sub

Sub CargarDatosAlosControles()
        If Me.IdOrden = 0 And lbCargaVentanaDesdeOtraVentana = True Then Exit Sub
        
        Dim oConexion As New Connection
        Dim oRsTmp As New Recordset
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        cmdBuscaCuentaPorApellidos.Enabled = False
        chkPlanNoCubre.Visible = False: txtNombrePaciente.Width = txtNombrePaciente.Width + chkPlanNoCubre.Width + 170
        'Carga datos de la orden
        Set mo_DOFactOrdenServicio = mo_ReglasFacturacion.FactOrdenServicioSeleccionarPorId(Me.IdOrden)
        
        If Not mo_DOFactOrdenServicio Is Nothing Then
             With mo_DOFactOrdenServicio
                  txtFregistro.Text = Format(.fechacreacion, sighEntidades.DevuelveFechaSoloFormato_DMY)
                  txtEstado.Text = mo_ReglasFarmacia.DevuelveEstadoActualDeFacturacion("idEstadoFacturacion=" & mo_DOFactOrdenServicio.idestadofacturacion)
                  mo_cmbIdPuntoCarga.BoundText = mo_DOFactOrdenServicio.idPuntoCarga
                  Me.txtIdOrden = Me.IdOrden
                  txtNcuenta.Text = .idCuentaAtencion
                  ml_IdFuenteFinanciamientoDespacho = .IdFuenteFinanciamiento
                  txtNcuenta_LostFocus
                  cmbFormaPago.BoundText = .idTipoFinanciamiento
                  mo_cmbServicioIngreso.BoundText = .idServicioPaciente
                  
                  ml_IdPaciente = .idPaciente
                  Me.txtFrealizaCpt.Text = Format(.FechaHoraRealizaCpt, sighEntidades.DevuelveFechaSoloFormato_DMY_HM)
            End With
            If mo_DOFactOrdenServicio.idestadofacturacion <> 1 And mo_DOFactOrdenServicio.idestadofacturacion <> 11 Then
               btnAceptar.Enabled = False
            End If
            If ml_IdPaciente = 0 Then
               'Preventa
               optPreventa.Value = True
            Else
               'Venta directa
               optVentas.Value = True
            End If
            txtFentrega.Text = Format(mo_DOFactOrdenServicio.FechaDespacho, sighEntidades.DevuelveFechaSoloFormato_DMY)
            txtHentrega.Text = Format(mo_DOFactOrdenServicio.FechaDespacho, sighEntidades.DevuelveHoraSoloFormato_HM)
            fraTipoVenta.Enabled = False
            ml_IdServicioPaciente = Val(mo_cmbServicioIngreso.BoundText)
            'debb-14/04/2011
            If Val(txtNcuenta.Text) > 0 Then
                
                Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(Val(txtNcuenta.Text), oConexion)
                If mi_Opcion = sghModificar And oRsTmp.Fields!IdFuenteFinanciamiento <> ml_IdFuenteFinanciamientoDespacho Then
                   MsgBox "No se podrá modificar datos, porque el despacho tubo otra PRODUCTO/PLAN" & Chr(13) & "hubo RECALCULO", vbInformation, Me.Caption
                   btnAceptar.Enabled = False
                End If
                lnIdTipoServicio = oRsTmp.Fields!idTipoServicio
                oRsTmp.Close
                
            End If
         Else
            mb_ExistenDatos = False
            Exit Sub
        End If
        'carga Nro Orden Pago
        If optPreventa.Value = True Then
           txtNroOrdenPago.Text = DevuelveNroOrdenPago(Me.IdOrden)
        End If
        'Cargar datos de los servicios
        Me.ucFacturacionProductos.LimpiarGrilla
        Me.ucFacturacionProductos.IdOrden = Me.IdOrden
        Me.ucFacturacionProductos.idTipoFinanciamiento = Val(cmbFormaPago.BoundText)
'        Me.ucFacturacionProductos.IdEstadoOrden = mo_DOFactOrdenServicio.IdEstadoOrden
        Me.ucFacturacionProductos.CargaDespachosPorIdOrden
         
        '
        Set oRsTmp = mo_ReglasComunes.RecetaCabeceraFiltraXpuntoCargaYDocumentodespacho(Trim(Str(mo_DOFactOrdenServicio.IdOrden)), sghPtoCargaServicioHospitalizacion)
        lnIdReceta = 0
        lnIdComprobantePagoDeReceta = 0
        If oRsTmp.RecordCount > 0 Then
           lnIdReceta = oRsTmp.Fields!idReceta
           Me.txtNreceta.Text = Trim(Str(lnIdReceta))
           Me.ucFacturacionProductos.PermiteAgregarItems = False
           oRsTmp.Close
           Set oRsTmp = mo_ReglasComunes.RecetaDetalleItemPorIdReceta(lnIdReceta)
           If oRsTmp.RecordCount > 0 Then
              If Not IsNull(oRsTmp!IdComprobantePago) Then
                 lnIdComprobantePagoDeReceta = oRsTmp!IdComprobantePago
                 lnIdReceta = 0
                 oRsTmp.Close
                 Set oRsTmp = mo_AdminCaja.CajaComprobantePagoServiciosPorIdComprobante(lnIdComprobantePagoDeReceta)
                 If oRsTmp.RecordCount > 0 Then
                    Me.txtNserie.Text = oRsTmp!nroSerie
                    Me.txtNboleta.Text = oRsTmp!nrodocumento
                 End If
              End If
           End If
        End If
        oRsTmp.Close
        '
        txtNcuenta.Enabled = False
        txtNroOrdenSisSoat.Enabled = False
        
        oConexion.Close
        Set oConexion = Nothing
        Set oRsTmp = Nothing
        
        Select Case mi_Opcion
        Case sghModificar
        Case sghEliminar
        Case sghConsultar
        End Select
   
   
End Sub

Function DevuelveNroOrdenPago(lnIdOrden As Long) As Long
        Dim oRsTmp1 As New Recordset
        Dim oConexion As New Connection
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        DevuelveNroOrdenPago = 0
        Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(lnIdOrden, oConexion)
        If oRsTmp1.RecordCount > 0 Then
           DevuelveNroOrdenPago = oRsTmp1.Fields!IdOrdenPago
        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
        oConexion.Close
        Set oConexion = Nothing
End Function




Sub ConfigurarPuntosDeCarga()
    
    mo_cmbIdPuntoCarga.ListField = "Descripcion"
    mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
    Set mo_cmbIdPuntoCarga.RowSource = mo_ReglasComunes.SeleccionarPuntosDeCarga()

End Sub











Sub Impresion()
    Dim oExcel As Excel.Application
    Dim oWorkBookPlantilla As Workbook
    Dim oWorkBook As Workbook
    Dim oWorkSheet As Worksheet
    Dim iFila As Long: Dim lnTotal As Double
    Dim rsreporte As New Recordset
    Dim mo_ReporteUtil As New ReporteUtil

        MousePointer = 11
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\EFacturacion_bs.xls")
        oWorkBookPlantilla.Worksheets("facturacion_bs").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
        '*******************************************Inicio de Impresion
        oWorkSheet.Cells(1, 2).Value = "Nº de Orden: " & mo_DOFactOrdenServicio.IdOrden
        oWorkSheet.Cells(3, 3).Value = txtNombrePaciente.Text
        oWorkSheet.Cells(3, 6).Value = txtNcuenta.Text
        oWorkSheet.Cells(4, 3).Value = cmbIdPuntoDeCarga.Text
        'oWorkSheet.Cells(4, 6).Value = txtNroHistoriaBusqueda.Text
        oWorkSheet.Cells(5, 3).Value = cmbFormaPago.Text
        If txtHentrega.Text <> sighEntidades.HORA_VACIA_HM Then
           oWorkSheet.Cells(5, 6).Value = CDate(txtFentrega.Text & " " & txtHentrega.Text)
        End If
        oWorkSheet.Cells(6, 2).Value = ""
        oWorkSheet.Cells(6, 3).Value = ""

        Set rsreporte = Me.ucFacturacionProductos.FacturacionProductos
        iFila = 9: lnTotal = 0
        rsreporte.MoveFirst
        Do While Not rsreporte.EOF
           oWorkSheet.Cells(iFila, 2).Value = rsreporte.Fields!Codigo
           oWorkSheet.Cells(iFila, 3).Value = rsreporte.Fields!NombreProducto
           oWorkSheet.Cells(iFila, 5).Value = Format(rsreporte.Fields!Cantidad, "####,###")
           oWorkSheet.Cells(iFila, 6).Value = Format(rsreporte.Fields!PrecioUnitario, "####,##0.000")
           oWorkSheet.Cells(iFila, 7).Value = Format(rsreporte.Fields!TotalPorPagar, "####,##0.00")
           lnTotal = lnTotal + rsreporte.Fields!TotalPorPagar
           iFila = iFila + 1
           rsreporte.MoveNext
        Loop
        iFila = iFila + 1
        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 7
        oWorkSheet.Cells(iFila, 2).Value = "Total: "
        oWorkSheet.Cells(iFila, 7).Value = Format(lnTotal, "####,##0.00")
        If oWorkSheet.PageSetup.PrintArea <> "" Then
           oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
        End If
        oExcel.Visible = True
        oWorkSheet.PrintPreview
        'oWorkSheet.PrintOut
        'oWorkBook.Close SaveChanges:=False
    Set oWorkSheet = Nothing
    Set oExcel = Nothing
        MousePointer = 1
End Sub

Sub CargaDataCombos()
    Set oRsFormaPago = mo_ReglasComunes.TiposFinanciamientoSegunFiltro("esFarmacia=1")
    Set cmbFormaPago.RowSource = oRsFormaPago
    cmbFormaPago.ListField = "Descripcion"
    cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
    '
    CargaServicio "(1,2,3,4)"
End Sub

Sub CargaServicio(lcFiltro As String)
    Dim oBuscaServicios As New SIGHNegocios.ReglasAdmision
    mo_cmbServicioIngreso.BoundColumn = "IdServicio"
    mo_cmbServicioIngreso.ListField = "DservicioHosp"
    Set mo_cmbServicioIngreso.RowSource = oBuscaServicios.DevuelveServiciosQueSonPuntosCarga(lcFiltro, sghFiltraSoloActivos, sghPorDescTipoServicio)
    Set oBuscaServicios = Nothing
End Sub


Sub CargaOrdenYaRegistradaEnSisSoat()
    If mi_Opcion = sghAgregar Then
        Dim LnIdFormaPago As Long
        mi_Opcion = sghModificar
        Me.IdOrden = txtNroOrdenSisSoat.Text
        CargarDatosAlFormulario
        LnIdFormaPago = Me.ucFacturacionProductos.OrdenRegistradaYaprobadaPorSisSoat
        If EsFecha(Left(mo_DOFactOrdenServicio.FechaDespacho, 10), "DD/MM/AAAA") = False Then
           If LnIdFormaPago > 1 Then
                cmbFormaPago.BoundText = LnIdFormaPago  '3=sis, 4=soat
                cmbFormaPago.Enabled = False
                Frame3.Enabled = True
                txtHentrega.Text = lcBuscaParametro.RetornaHoraServidorSQL 'Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
                txtFentrega.Text = lcBuscaParametro.RetornaFechaServidorSQL  'Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
                txtHentrega.Enabled = True
                txtFentrega.Enabled = True
                lbDocumentoYaRegistradoEnSeguros = True
                'Cargar datos de los servicios
                Me.ucFacturacionProductos.LimpiarGrilla
                Me.ucFacturacionProductos.DocumentoYaRegistradoEnSeguros = lbDocumentoYaRegistradoEnSeguros
                Me.ucFacturacionProductos.IdOrden = Me.IdOrden
                'Me.ucFacturacionProductos.IdEstadoOrden = mo_DOFactOrdenServicio.IdEstadoOrden
                Me.ucFacturacionProductos.idTipoFinanciamiento = LnIdFormaPago
                Me.ucFacturacionProductos.CargaProductosPorIdOrden
            Else
                MsgBox "Ese 'Nro de Orden' NO ha sido Aprobado en SIS o SOAT", vbInformation, "Mensaje"
                btnCancelar_Click
            End If
        Else
            MsgBox "Ese 'Nro de Orden' ha sido Registrado y Aprobado en SIS o SOAT, pero ya fue despachado", vbInformation, "Mensaje"
            btnCancelar_Click
        End If
   End If
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
            txtNombrePaciente.Text = oDOPaciente.NroHistoriaClinica & " " & Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + oDOPaciente.PrimerNombre
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
    oConexion.Close
    Set oConexion = Nothing
End Sub





Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub



Private Sub optPreventa_Click(Value As Integer)
    If optPreventa.Value = True And mi_Opcion = sghAgregar Then
        mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, True
        mo_Formulario.HabilitarDeshabilitar Me.txtFentrega, False
        mo_Formulario.HabilitarDeshabilitar Me.txtHentrega, False
        mo_Formulario.HabilitarDeshabilitar Me.txtNcuenta, False
        '
        txtFentrega.Text = sighEntidades.FECHA_VACIA_DMY
        txtHentrega.Text = sighEntidades.HORA_VACIA_HM
        '
        Set oRsTipoFinanciamiento = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia(" and dbo.TiposFinanciamiento.TipoVenta='P'")
        lcPosicionDefaultCombo = ""
        If oRsTipoFinanciamiento.RecordCount = 1 Then
            lcPosicionDefaultCombo = Trim(Str(oRsTipoFinanciamiento.Fields!idTipoFinanciamiento))
        End If
        cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
        cmbFormaPago.ListField = "Descripcion"
        Set cmbFormaPago.RowSource = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia(" and dbo.TiposFinanciamiento.TipoVenta='P'")
        If lcPosicionDefaultCombo <> "" Then
           cmbFormaPago.BoundText = lcPosicionDefaultCombo
           ml_IdFuenteFinanciamiento = 1 'contado
        End If
        ml_IdPuntoCargaServicioHosp = 0
        chkPlanNoCubre.Enabled = False
    End If
End Sub

Private Sub optVentas_Click(Value As Integer)
   If optVentas.Value = True And mi_Opcion = sghAgregar Then
      mo_Formulario.HabilitarDeshabilitar Me.txtNcuenta, True
      '
      txtHentrega.Text = lcBuscaParametro.RetornaHoraServidorSQL   'Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
      txtFentrega.Text = lcBuscaParametro.RetornaFechaServidorSQL    'Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
      '
      cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
      cmbFormaPago.ListField = "Descripcion"
      Set cmbFormaPago.RowSource = mo_ReglasFarmacia.TipoFinanciamientosDevuelveSoloFarmacia("")
      chkPlanNoCubre.Enabled = True
   End If
End Sub

Private Sub txtFentrega_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFentrega

End Sub



Private Sub txtHentrega_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHentrega
End Sub

Private Sub txtNboleta_LostFocus()
    If Trim(txtNserie.Text) <> "" And Val(txtNboleta.Text) > 0 And mi_Opcion = sghAgregar Then
        lnIdComprobantePagoDeReceta = 0
        lnIdReceta = 0
        Dim rsBuscaBoleta As New Recordset
        Dim rsBuscaBoletaEnImagenes As New Recordset
        Dim oConexion As New Connection
        Dim oRsTmp1 As New Recordset
        oConexion.CommandTimeout = 300
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set rsBuscaBoleta = mo_AdminCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumentoConexion(txtNserie.Text, Trim(txtNboleta.Text), oConexion)
        If rsBuscaBoleta.RecordCount > 0 Then
            '
            'lnIdPacienteHistorico = 0
            If rsBuscaBoleta.Fields!idPaciente > 0 Then
              ' lnIdPacienteHistorico = rsBuscaBoleta.Fields!idPaciente
              ' chkMuestraHistorico_Click
            End If
            '
            If rsBuscaBoleta.Fields!idEstadoComprobante <> sghEstadosComprobante.sighEstadosComprobantePagado Then
                MsgBox "Esa Boleta está ANULADA", vbInformation, Me.Caption
                txtNboleta.Text = ""
                txtNserie.Text = ""
            Else
                Set rsBuscaBoletaEnImagenes = mo_ReglasComunes.RecetaDetalleItemFiltraXidComprobantePagoYpuntoCarga(rsBuscaBoleta!IdComprobantePago, sghPtoCargaServicioHospitalizacion)
                If rsBuscaBoletaEnImagenes.RecordCount = 0 Then
                   MsgBox "Esa Boleta no tiene RECETA", vbInformation, ""
                ElseIf rsBuscaBoletaEnImagenes!idEstado <> sghRecetaEstados.sighRecetaConBoleta Then
                   MsgBox "Esa Boleta tiene RECETA pero ya se DESPACHO/ANULO", vbInformation, ""
                ElseIf ml_IdPaciente <> rsBuscaBoleta!idPaciente And lbCargaVentanaDesdeOtraVentana = True Then
                     MsgBox "El Paciente de la RECETA no es el mismo", vbInformation, ""
                     txtNreceta.Text = ""
                Else
                   lnIdComprobantePagoDeReceta = rsBuscaBoleta!IdComprobantePago
                   Set oRsTmp1 = mo_ReglasComunes.RecetasConCabeceraYdetalleSoloCpt(rsBuscaBoletaEnImagenes!idReceta, 0)
                   If lbCargaVentanaDesdeOtraVentana = False Then
                        lbCuentaDeEmergenciaCerrada = mo_ReglasComunes.CuentaDeEmergenciaCerrada(oRsTmp1!idCuentaAtencion, sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
                        txtNcuenta.Text = oRsTmp1.Fields!idCuentaAtencion
                        txtNcuenta_LostFocus
                   End If
                   ucFacturacionProductos.PermiteAgregarItems = False
                   ucFacturacionProductos.CargaProductosDeLaBoleta oRsTmp1
                End If
            End If
        End If
        Set rsBuscaBoleta = Nothing
        Set rsBuscaBoletaEnImagenes = Nothing
        oConexion.Close
        Set oConexion = Nothing
        Set oRsTmp1 = Nothing
    End If

End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNcuenta_LostFocus()
   If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
       Dim oRsTmp As New Recordset
       Dim lbSigue As Boolean
       Dim oConexion As New Connection
       oConexion.Open sighEntidades.CadenaConexion
       oConexion.CursorLocation = adUseClient
       Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
       txtDatosDeCuenta.Text = ""
       If mi_Opcion = sghAgregar Then
          cmbFormaPago.Text = ""
       End If
       ml_IdPaciente = 0
       ml_IdFuenteFinanciamiento = 0
       ml_IdPuntoCargaServicioHosp = 99999
       txtPlan.Text = ""
       lbSigue = True
       If oRsTmp.RecordCount > 0 Then
          If oRsTmp.Fields!idEstado <> 1 Then
             If mi_Opcion <> sghConsultar And lbCuentaDeEmergenciaCerrada = False Then
                MsgBox "Ese estado de Cuenta no se encuentra ABIERTA", vbInformation, Me.Caption
                If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
                   btnAceptar.Enabled = False
                Else
                   lbSigue = False
                End If
             End If
          End If
          If mo_lbNOValidaCodigoPrestacion = False Then
                If mi_Opcion = sghAgregar And _
                   mo_AdminAdmision.AtencionesDatosAdicionalesSItieneCodigoPrestacionSIS(Val(txtNcuenta.Text), _
                                                                  wxParametro302, _
                                                                  oRsTmp.Fields!IdFuenteFinanciamiento) = False Then
                                                                             
                   lbSigue = False
                End If
          End If
          If lbSigue Then
                lnIdTipoServicio = oRsTmp.Fields!idTipoServicio
                txtDatosDeCuenta.Text = "F.Ing: " & oRsTmp.Fields!FechaIngreso & " - " & IIf(oRsTmp.Fields!idTipoServicio = 1, "Consultorios Externos", IIf(oRsTmp.Fields!idTipoServicio = 3, "Hospitalización", "Emergencia")) & " - (Est: " & Trim(oRsTmp.Fields!estadoCta) & ")"
                txtPlan.Text = "IAFA Act.: " & oRsTmp.Fields!dFuenteFinanciamiento
                ml_IdPaciente = oRsTmp.Fields!idPaciente
                ml_IdFuenteFinanciamiento = oRsTmp.Fields!IdFuenteFinanciamiento
                txtNombrePaciente.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oRsTmp!NroHistoriaClinica)), False) & _
                                         " " & Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!PrimerNombre
                
                
                ucFacturacionProductos.idTipoFinanciamiento = oRsTmp.Fields!IdFormaPago
                If mi_Opcion = sghAgregar Then
                   cmbFormaPago.BoundText = oRsTmp.Fields!IdFormaPago
                End If
                If mi_Opcion <> sghAgregar And ml_IdFuenteFinanciamiento <> ml_IdFuenteFinanciamientoDespacho And ml_IdFuenteFinanciamientoDespacho > 0 Then
                   txtPlan.Text = "Plan Desp: " & Trim(mo_ReglasFacturacion.FuentesFinanciamientoDevuelveNombrePlan(ml_IdFuenteFinanciamientoDespacho)) & " - " & txtPlan.Text
                End If
                '
                ml_IdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(txtNcuenta.Text, CDate(txtFregistro.Text), lcBuscaParametro.RetornaHoraServidorSQL)
                txtProcedencia.Text = mo_ReglasFacturacion.BuscaServicioActualDelPaciente(ml_IdServicioPaciente)
                oRsTmp.Close
                '
                If lcBuscaParametro.SeleccionaFilaParametro(246) <> "N" Then
                   'para poder registrar cualquier CPT en ese SERVICIO
                   '(lo registra FACTURACION)
                    Set oRsTmp = mo_ReglasFacturacion.ServiciosSeleccionarPorFiltro("idServicio=" & Trim(Str(ml_IdServicioPaciente)), sghPorDescripcion)
                    If oRsTmp.RecordCount > 0 Then
                       CargaServicio "(" & Trim(Str(oRsTmp.Fields!idTipoServicio)) & ")"
                    End If
                    oRsTmp.Close
                End If
                mo_cmbServicioIngreso.BoundText = ml_IdServicioPaciente
                '
                Set oRsTmp = mo_ReglasComunes.FactPuntosCargaSeleccionarPorFiltro("idServicio=" & Trim(Str(ml_IdServicioPaciente)))
                If oRsTmp.RecordCount > 0 Then
                   ml_IdPuntoCargaServicioHosp = oRsTmp.Fields!idPuntoCarga
                Else
                   ml_IdPuntoCargaServicioHosp = 9999
                End If
                If lcBuscaParametro.SeleccionaFilaParametro(246) <> "N" Then
                   'para poder registrar cualquier CPT en ese SERVICIO
                   '(lo registra FACTURACION)
                   ucFacturacionProductos.IdPuntoCargaServicioHosp = ml_IdPuntoCargaServicioHosp   '0 <--para poder registrar cualquier CPT en ese SERVICIO
                Else
                   'solo registra CPT de ese Servicio (Punto de carga)
                   '(lo registra el mismo SERVICIO HOSP)
                   ucFacturacionProductos.IdPuntoCargaServicioHosp = ml_IdPuntoCargaServicioHosp
                End If
                ucFacturacionProductos.TabEnDescripcionParaFactuacion
                If mi_Opcion = sghAgregar And chkPlanNoCubre.Value = 1 Then
                   chkPlanNoCubre_Click
                End If
          End If
       End If
       
       oRsTmp.Close
       Set oRsTmp = Nothing
       oConexion.Close
       Set oConexion = Nothing
   End If
End Sub

Public Sub CargaCptPorAtencion()
    'Actualizado 09102014
    If lbCargaVentanaDesdeOtraVentana = True Then
         Me.ucFacturacionProductos.idCuentaAtencion = ml_idCuentaAtencion
         Me.ucFacturacionProductos.CargaCptPorAtencion
         
        Dim rsProductos As Recordset
        Set rsProductos = Me.ucFacturacionProductos.FacturacionProductos
        If Not (rsProductos.EOF And rsProductos.BOF) Then
            rsProductos.MoveFirst
            Do While Not rsProductos.EOF
                If rsProductos!idProducto = 0 Then
                   rsProductos.Delete
                   rsProductos.Update
                End If
                rsProductos.MoveNext
            Loop
        End If
    End If
End Sub

Private Sub txtNroOrdenSisSoat_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroOrdenSisSoat
End Sub

Private Sub txtNroOrdenSisSoat_LostFocus()
    CargaOrdenYaRegistradaEnSisSoat
End Sub

Private Sub ucFacturacionProductos_SePresionoTeclaEspecial(KeyCode As Integer)
     If KeyCode = vbKeyF2 Then
        AdministrarKeyPreview KeyCode
        Me.KeyPreview = False
     End If

End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub


Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_AdminAdmision = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_AdminCaja = Nothing
    Set mo_ReglasComunes = Nothing
    Set mo_ReglasSeguridad = Nothing
    Set mo_AdminArchivoClinico = Nothing
    Set mo_Apariencia = Nothing
    Set mo_cmbIdPuntoCarga = Nothing
    Set mo_cmbIdEstado = Nothing
    Set mo_cmbFechaIngreso = Nothing
    Set mo_cmbIdTipoGenHistoriaClinica = Nothing
    Set mo_cmbServicioIngreso = Nothing
    Set mo_DOFactOrdenServicio = Nothing
    Set mo_DoAtencion = Nothing
    Set lcBuscaParametro = Nothing
    Set oRsFormaPago = Nothing
    Set mo_Teclado = Nothing
    Set mo_Formulario = Nothing
    Set oRsTipoFinanciamiento = Nothing
End Sub

'mgaray201411a
Private Function GetMostrarLabEnListaProductos() As Boolean
    Dim returnValue As Boolean
    returnValue = False
    Select Case ml_FormMostradoDesde
        Case 1:
            'mgaray201412a para ocultar lab
            returnValue = False
        Case Else:
            returnValue = False
    End Select
    GetMostrarLabEnListaProductos = returnValue
End Function


Private Sub txtNreceta_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtNreceta
       AdministrarKeyPreview KeyCode
End Sub



Private Sub cmbBuscaReceta_Click()
    Dim oBusqueda As New SIGHNegocios.clBuscaReceta
    oBusqueda.idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
       txtNreceta.Text = oBusqueda.IdRecetaSeleccionada
       txtNreceta_LostFocus
    End If
    Set oBusqueda = Nothing
End Sub


Private Sub txtNreceta_LostFocus()
    If Val(txtNreceta.Text) > 0 And mi_Opcion = sghAgregar Then
       lnIdReceta = 0
       lnIdComprobantePagoDeReceta = 0
       Dim lcSql As String
       Dim oRsTmp1 As New Recordset, lnRecetaProcesada As Long, lnCuenta As Long
       
       lnRecetaProcesada = Val(txtNreceta.Text)
       '
       ucFacturacionProductos.LimpiarGrilla


       
       Set oRsTmp1 = mo_ReglasComunes.RecetasConCabeceraYdetalleSoloCpt(lnRecetaProcesada, sghRecetaEstados.sighRecetaRegistrada)
       If oRsTmp1.RecordCount > 0 Then
            If oRsTmp1.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                mo_ReglasComunes.RecetaChequeaEstadoActual oRsTmp1.Fields!idCuentaAtencion, _
                                                           oRsTmp1.Fields!idEstado, _
                                                           0, oRsTmp1.Fields!DocumentoDespacho
                txtNreceta.Text = ""
            Else
                If oRsTmp1.Fields!idPuntoCarga <> sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion Then
                     MsgBox "Esa receta no es de CONSUMO EN EL SERVICIO", vbInformation, ""
                     txtNreceta.Text = ""
                ElseIf ml_IdFuenteFinanciamiento = 1 And lnIdTipoServicio = 1 Then
                     MsgBox "El Paciente es PAGANTE debe ingresar en N°BOLETA", vbInformation, ""
                     txtNreceta.Text = ""
                ElseIf ml_IdPaciente <> oRsTmp1!idPaciente And lbCargaVentanaDesdeOtraVentana = True Then
                     MsgBox "El Paciente de la RECETA no es el mismo", vbInformation, ""
                     txtNreceta.Text = ""
                Else
                     If lbCargaVentanaDesdeOtraVentana = False Then
                        lbCuentaDeEmergenciaCerrada = mo_ReglasComunes.CuentaDeEmergenciaCerrada(oRsTmp1!idCuentaAtencion, sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
                        txtNcuenta.Text = oRsTmp1.Fields!idCuentaAtencion
                        txtNcuenta_LostFocus
                     End If
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
Private Sub txtNserie_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNserie
End Sub

Private Sub txtNboleta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNboleta
End Sub

Private Sub txtNboleta_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
End Sub



