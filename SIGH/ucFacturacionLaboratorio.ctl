VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucFacturacionLaboratorio 
   ClientHeight    =   8190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10980
   ScaleHeight     =   8190
   ScaleWidth      =   10980
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   75
      TabIndex        =   20
      Top             =   495
      Width           =   6780
      Begin VB.CheckBox chkPorFcpt 
         Caption         =   "Filtrar por Fechas realizar CPT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   39
         Top             =   1320
         Width           =   2835
      End
      Begin VB.TextBox txtNombres 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5280
         MaxLength       =   40
         TabIndex        =   6
         Top             =   420
         Width           =   1395
      End
      Begin VB.TextBox txtNCuenta 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         MaxLength       =   9
         TabIndex        =   2
         Top             =   390
         Width           =   1155
      End
      Begin VB.ComboBox cmbFecha 
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
         ItemData        =   "ucFacturacionLaboratorio.ctx":0000
         Left            =   4125
         List            =   "ucFacturacionLaboratorio.ctx":0002
         TabIndex        =   11
         Text            =   "cmbFecha"
         Top             =   735
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.ComboBox cmbIdPtoCarga 
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
         Left            =   3210
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   420
         Width           =   2085
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   5415
         Picture         =   "ucFacturacionLaboratorio.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   5430
         Picture         =   "ucFacturacionLaboratorio.ctx":2C4D
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   1275
      End
      Begin VB.TextBox txtNroOrden 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1020
         MaxLength       =   9
         TabIndex        =   1
         Top             =   405
         Width           =   1005
      End
      Begin VB.TextBox txtNroHistoria 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         MaxLength       =   9
         TabIndex        =   0
         Top             =   405
         Width           =   945
      End
      Begin MSDataListLib.DataCombo cmbFarmacia 
         Height          =   330
         Left            =   4635
         TabIndex        =   12
         Top             =   735
         Visible         =   0   'False
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox txtFF 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   900
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFI 
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   900
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin Threed.SSOption SSApellido 
         Height          =   240
         Left            =   5280
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   423
         _Version        =   262144
         Caption         =   "Apellido Paterno"
      End
      Begin Threed.SSOption ssNombre 
         Height          =   240
         Left            =   5280
         TabIndex        =   9
         Top             =   750
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   423
         _Version        =   262144
         Caption         =   "Primer Nombre"
         Value           =   -1
      End
      Begin MSMask.MaskEdBox txtFcpt1 
         Height          =   315
         Left            =   60
         TabIndex        =   40
         Top             =   1530
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
      Begin MSMask.MaskEdBox txtFCpt2 
         Height          =   315
         Left            =   1440
         TabIndex        =   41
         Top             =   1530
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
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "              Fechas de movimientos"
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
         Left            =   75
         TabIndex        =   38
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label lblFarmacia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Farmacia"
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
         Left            =   7470
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "  Nº H.C.     N° Movim    Nº Cuenta      Punto de Carga            Paciente"
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
         Left            =   60
         TabIndex        =   21
         Top             =   210
         Width           =   6795
      End
   End
   Begin UltraGrid.SSUltraGrid grdListaOrdenes 
      Height          =   1710
      Left            =   90
      TabIndex        =   13
      Top             =   2385
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   3016
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "LISTADO DE ÓRDENES"
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "DETALLE DE LA ORDEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3735
      Left            =   45
      TabIndex        =   24
      Top             =   4320
      Width           =   10815
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprime"
         Height          =   645
         Left            =   9330
         Picture         =   "ucFacturacionLaboratorio.ctx":5829
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   210
         Width           =   1080
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7980
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   540
         Width           =   1305
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7980
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   210
         Width           =   1305
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   210
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   210
         Width           =   4425
      End
      Begin UltraGrid.SSUltraGrid grdListaOrdenesDetalle 
         Height          =   2550
         Left            =   60
         TabIndex        =   19
         Top             =   990
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   4498
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "grdListaOrdenesDetalle"
      End
      Begin VB.Label lblIdOrden 
         AutoSize        =   -1  'True
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   10455
         TabIndex        =   43
         Top             =   585
         Width           =   135
      End
      Begin VB.Label lblEPSpago 
         AutoSize        =   -1  'True
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   10470
         TabIndex        =   42
         Top             =   225
         Width           =   135
      End
      Begin VB.Label lblPRes 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6420
         TabIndex        =   36
         Top             =   750
         Width           =   990
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Items con Resultados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   435
         Left            =   4380
         TabIndex        =   35
         Top             =   750
         Width           =   1980
      End
      Begin VB.Label lblPReg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3600
         TabIndex        =   34
         Top             =   750
         Width           =   990
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Items Registrados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   435
         Left            =   1920
         TabIndex        =   33
         Top             =   750
         Width           =   1710
      End
      Begin VB.Label lblT 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   960
         TabIndex        =   32
         Top             =   750
         Width           =   750
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   75
         TabIndex        =   31
         Top             =   750
         Width           =   780
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Historia Clínica"
         Height          =   255
         Left            =   6840
         TabIndex        =   28
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Id Paciente"
         Height          =   255
         Left            =   6480
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Sexo"
         Height          =   255
         Left            =   60
         TabIndex        =   26
         Top             =   210
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombres y Apellidos"
         Height          =   255
         Left            =   60
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Ayuda: <Doble Click sobre el nombre de la prueba> = Ingreso / Modificación de resultados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   190
         Left            =   120
         TabIndex        =   29
         Top             =   3510
         Width           =   8625
      End
   End
   Begin UltraGrid.SSUltraGrid grdBoletas 
      Height          =   1890
      Left            =   6870
      TabIndex        =   37
      Top             =   480
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   3334
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Boletas pendientes  (marcar y  F2->Agregar)"
   End
   Begin VB.Label Label7 
      Caption         =   "Ayuda: <Click sobre la Órden de Laboratorio> = Ver detalle de la Órden de Laboratorio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   4095
      Width           =   8625
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Facturación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   10875
   End
End
Attribute VB_Name = "ucFacturacionLaboratorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para buscar movimientos en Laboratorio
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminCaja As New SIGHDatos.CatalogoServicios
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_ReglasDeSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim ml_idUsuario As Long
Dim mo_cmbIdPuntoCarga As New sighentidades.ListaDespleglable
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim ml_IdTipoFinanciamiento As Long
Dim oRsFarmacias As New ADODB.Recordset
Dim oRsBoletas As New ADODB.Recordset
Dim oRsLista As New Recordset
Dim mrs_FacturacionProductos As New Recordset
Dim rs As Recordset
Dim rsTmp As New Recordset

Dim ml_resultado As String
Dim ml_observacion As String
Dim ml_realiza As Long
Dim ml_NombreMedico As String
Dim ml_idRegistroSeleccionado As Long
Dim ml_IdPaciente As Long
Dim ml_IdPruebaSeleccionada As String
Dim ml_CodigoPruebaSeleccionada As String
Dim ml_idOrden As Long
Dim ml_idOrdenLab As Long
Dim ml_NombrePruebaSeleccionada As String
Dim ml_nombrePaciente As String
Dim ml_PuntoCarga As sghTipoFiltroPacientes
Dim ml_areaTrabajo As Long
Dim ldFechaNacimiento As Date, lnIdTipoSexo As Long
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS  As Long
Dim ml_SeEligioGridBoleta As Boolean
Dim md_fechaNacimiento As Date
Dim lcServicioActualPaciente As String
Dim lcEdadEnAtencion As String

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
Property Let AreaTrabajo(lValue As Long)
    ml_areaTrabajo = lValue
End Property

Property Get AreaTrabajo() As Long
  AreaTrabajo = ml_areaTrabajo
End Property

Property Set DataSource(oValue As ADODB.Recordset)
  Set UserControl.grdListaOrdenes.DataSource = oValue
End Property

Property Get DataSource() As ADODB.Recordset
  Set DataSource = UserControl.grdListaOrdenes.DataSource
End Property

Property Let idRegistroSeleccionado(lValue As Long)
  ml_idRegistroSeleccionado = lValue
End Property

Property Get idRegistroSeleccionado() As Long
  idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property

Property Let idUsuario(lValue As Long)
  ml_idUsuario = lValue
End Property

Property Get idUsuario() As Long
  idUsuario = ml_idUsuario
End Property

Property Let Titulo(lValue As String)
  lblNombre = lValue
End Property

Property Get Titulo() As String
  Titulo = lblNombre
End Property

Property Let PuntoCarga(lValue As Long)
  ml_PuntoCarga = lValue
  mo_cmbIdPuntoCarga.BoundText = ml_PuntoCarga
End Property

Property Get PuntoCarga() As Long
  PuntoCarga = ml_PuntoCarga
End Property

Property Let HabilitarPuntoCarga(lValue As Long)
  cmbIdPtoCarga.Enabled = lValue
End Property

Property Get HabilitarPuntoCarga() As Long
  HabilitarPuntoCarga = cmbIdPtoCarga.Enabled
End Property

Property Let idTipoFinanciamiento(lValue As Long)
  ml_IdTipoFinanciamiento = lValue
End Property

Property Get idTipoFinanciamiento() As Long
  idTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property

Sub ConfiguraFechaHora()
  Dim Hora As Date
  Hora = Format(Now, sighentidades.DevuelveHoraSoloFormato_HMS)
  If Hora >= "07:00:00" And Hora < "19:00:00" Then
    txtFI.Text = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY) & " 07:00:00"
    txtFF.Text = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY) & " 18:59:59"
  ElseIf Hora < "07:00:00" Then
    txtFI.Text = Format(DateAdd("d", -1, Now), sighentidades.DevuelveFechaSoloFormato_DMY) & " 19:00:00"
    txtFF.Text = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY) & " 06:59:59"
  ElseIf Hora >= "19:00:00" Then
    txtFI.Text = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY) & " 19:00:00"
    txtFF.Text = Format(DateAdd("d", 1, Now), sighentidades.DevuelveFechaSoloFormato_DMY) & " 06:59:59"
  End If
End Sub

Sub CargarItemsALaGrillaS(rs As Recordset)
  Dim mrs_FacturacionProductos As Recordset
  Set mrs_FacturacionProductos = Nothing
  Do While Not rs.EOF
    mrs_FacturacionProductos.AddNew
    mrs_FacturacionProductos!idProducto = rs!idProducto
    mrs_FacturacionProductos!Codigo = rs!Codigo
    mrs_FacturacionProductos!NombreProducto = rs!nombre
    mrs_FacturacionProductos!idTipoFinanciamiento = ml_IdTipoFinanciamiento
    mrs_FacturacionProductos!Cantidad = rs!Cantidad
    mrs_FacturacionProductos!PrecioUnitario = rs!Precio
    mrs_FacturacionProductos!TotalPorPagar = rs!Total
    rs.MoveNext
  Loop
  mrs_FacturacionProductos.Close
  Set grdListaOrdenesDetalle.DataSource = mrs_FacturacionProductos
End Sub

Private Sub btnBuscar_Click()
  Screen.MousePointer = vbHourglass
  Set grdListaOrdenesDetalle.DataSource = Nothing
  Set rs = Nothing
  Text1.Text = ""
  Text2.Text = ""
  Text3.Text = ""
  Text4.Text = ""
  ml_idRegistroSeleccionado = 0
  ml_IdPaciente = 0
  lblT.Caption = "0.00"
  lblPReg.Caption = "0"
  lblPRes.Caption = "0"
  RealizarBusqueda
  Screen.MousePointer = vbDefault
  ml_SeEligioGridBoleta = False
End Sub

Public Sub RealizarBusqueda()
  On Error Resume Next
  Dim ldFechaIni As Date
  Dim ldFechaFin As Date
  Dim lcFiltro As String
  Dim oRsTmp1 As New Recordset
  Dim lcBoleta As String, lnTotal As Double, ldFecha As Date, lnIdProducto As Long
  Dim lnIdComprobantePago As Long, lbEsDelPuntoCarga As Boolean, lcFiltroPaciente As String
  lcFiltroPaciente = ""
  If txtFI.Enabled = True Then
        If Not IsDate(txtFI.Text) Then
          MsgBox "Fecha Inicial no válida.", vbInformation, "SIGH "
          Exit Sub
        End If
        If Not IsDate(txtFF.Text) Then
          MsgBox "Fecha Final no válida.", vbInformation, "SIGH "
          Exit Sub
        End If
          If CDate(txtFI.Text) > CDate(txtFF.Text) Then
             MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
             Exit Sub
          End If
  End If
  If cmbIdPtoCarga.Text = "" Then
    MsgBox "Punto de Carga no válido.", vbInformation, "SIGH "
    Exit Sub
  End If
  If (UserControl.txtNroHistoria.Text = "" And UserControl.txtNroOrden.Text = "" And UserControl.txtNcuenta.Text = "" And IsDate(txtFI.Text) = False And IsDate(txtFF.Text) = False And txtNombres.Text = "") Then
    MsgBox "Por favor ingrese algunos de los filtros (Nº Historia, Nº cuenta, Nº Orden, Fecha ó Nombres)", vbInformation, "Filtro de ordenes de procedimientos"
    Exit Sub
  End If
        
  Dim oDOFactOrdenServicio As New DOFactOrdenServicio
  Dim oDOFactOrdenBienInsumo As New DOFactOrdenBienInsumo
  Dim oDOPaciente As New doPaciente
        
  Select Case ml_PuntoCarga
    Case 5
      oDOFactOrdenBienInsumo.IdOrden = Val(UserControl.txtNroOrden)
      oDOFactOrdenBienInsumo.idPuntoCarga = Val(mo_cmbIdPuntoCarga.BoundText)
      If mo_Teclado.TextoEsSoloNumeros(UserControl.txtNroHistoria) Then
         oDOPaciente.NroHistoriaClinica = Val(UserControl.txtNroHistoria)
      End If
    Case Else
      If txtFI.Enabled = True Then
        ldFechaIni = CDate(txtFI.Text)
        ldFechaFin = CDate(txtFF.Text)
      Else
        ldFechaIni = CDate(txtFcpt1.Text & " 00:00:00")
        ldFechaFin = CDate(txtFCpt2.Text & " 23:59:59")
        If ldFechaIni > ldFechaFin Then
           MsgBox "La FECHA FINAL debe ser mayor a la FECHA INICIAL", vbInformation, ""
           Exit Sub
        End If
      End If
      lcFiltro = ""
      If mo_Teclado.TextoEsSoloNumeros(txtNroHistoria.Text) Then
          'lcFiltro = "NroHistoriaClinica=" & HCigualDNI_AgregaNUEVEaLaHistoria(txtNroHistoria.Text)
          lcFiltro = "NroHistoriaClinica=" & txtNroHistoria.Text
      End If
      If mo_Teclado.TextoEsSoloNumeros(txtNroOrden.Text) Then
        If lcFiltro = "" Then
          lcFiltro = "idMovimiento=" & Val(txtNroOrden.Text)
        Else
          lcFiltro = lcFiltro & " AND idMovimiento=" & Val(txtNroOrden.Text)
        End If
      End If
      If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
        If lcFiltro = "" Then
          lcFiltro = "idCuentaAtencion=" & Val(txtNcuenta.Text)
        Else
          lcFiltro = lcFiltro & " AND idCuentaAtencion=" & Val(txtNcuenta.Text)
        End If
      End If
      If Trim(txtNombres.Text) <> "" Then
        lcFiltroPaciente = Left(Trim(txtNombres.Text), 50)
      End If
      oDOFactOrdenServicio.IdOrden = Val(UserControl.txtNroOrden)
      oDOFactOrdenServicio.idPuntoCarga = Val(mo_cmbIdPuntoCarga.BoundText)
      oDOPaciente.NroHistoriaClinica = Val(UserControl.txtNroHistoria)
      Set oRsLista = mo_ReglasFacturacion.FactOrdenServicioPorFechasLabPaciente(ldFechaIni, ldFechaFin, _
                                                    Val(mo_cmbIdPuntoCarga.BoundText), lcFiltroPaciente, _
                                                    IIf(txtFI.Enabled = True, 1, 0))
      If lcFiltro <> "" Then oRsLista.Filter = lcFiltro
      Set grdListaOrdenes.DataSource = oRsLista
      Set grdBoletas.DataSource = mo_ReglasCaja.BoletasServicioPorPuntoCarga(ldFechaIni, ldFechaFin, Val(mo_cmbIdPuntoCarga.BoundText))
      UserControl.txtNroHistoria.SetFocus 'Actuaizado Frank 16092014
  End Select
  If mo_ReglasFacturacion.MensajeError <> "" Then MsgBox mo_ReglasFacturacion.MensajeError, vbInformation, "Filtro órdenes de procedimientos"
  'mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenes, sighentidades.GrillaConFilasBicolor
  Set oRsTmp1 = Nothing
End Sub

Private Sub btnLimpiar_Click()
  LimpiarFiltro
End Sub

Public Sub LimpiarFiltro()
  ConfiguraFechaHora
  UserControl.txtNroHistoria.Text = ""
  UserControl.txtNroOrden.Text = ""
  UserControl.txtNcuenta.Text = ""
  UserControl.txtNombres.Text = ""
  SSApellido.Value = False
  ssNombre.Value = True
  ml_SeEligioGridBoleta = False
End Sub

Private Sub chkPorFcpt_Click()
    If chkPorFcpt.Value = 1 Then
        txtFcpt1.Enabled = True
        txtFCpt2.Enabled = True
        txtFI.Enabled = False
        txtFF.Enabled = False
    Else
        txtFcpt1.Enabled = False
        txtFCpt2.Enabled = False
        txtFI.Enabled = True
        txtFF.Enabled = True
    End If
End Sub

Private Sub cmbFecha_Click()
  'If cmbIdPtoCarga.Text <> "" And cmbFecha.Text <> "" Then btnBuscar_Click
End Sub

Private Sub cmbFecha_KeyPress(KeyAscii As Integer)
  'If KeyAscii = 13 Then btnBuscar_Click
End Sub

Private Sub cmbIdPtoCarga_Click()
  'If cmbIdPtoCarga.Text <> "" And IsDate(txtFI.Text) And IsDate(txtFF.Text) Then btnBuscar_Click
End Sub

Private Sub cmbIdPtoCarga_KeyPress(KeyAscii As Integer)
 ' If KeyAscii = 13 Then btnBuscar_Click
End Sub


Private Sub cmdImprimir_Click()
  If Val(lblPRes.Caption) = 0 Then
    MsgBox "No hay pruebas con resultados para poder imprimir", vbInformation, "Laboratorio"
    Exit Sub
  End If
  rsTmp.Filter = "Imprime=true"
  If rsTmp.RecordCount = 0 Then
     rsTmp.Filter = ""
     Exit Sub
  End If
  rsTmp.MoveFirst
  If Not (rsTmp.EOF = True And rsTmp.BOF = True) Then
    Dim lbExisteAlgunResultado As Boolean
    Dim ldFechaResultado As Date
    lbExisteAlgunResultado = False
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
      ml_IdPruebaSeleccionada = rsTmp!Codigo
      ml_NombrePruebaSeleccionada = rsTmp("nombre")
      ml_idRegistroSeleccionado = rsTmp("idOrden")
      ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(ml_IdPruebaSeleccionada, CDbl(ml_idRegistroSeleccionado))
      ml_observacion = mo_ReglasLaboratorio.LabRecuperaResultados_Obs(ml_IdPruebaSeleccionada, CDbl(ml_idRegistroSeleccionado))
      'Dim ldFechaResultado As Date
      ml_realiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(ml_IdPruebaSeleccionada, CDbl(ml_idRegistroSeleccionado), ldFechaResultado)
      ml_CodigoPruebaSeleccionada = mo_ReglasLaboratorio.LabAveriguaCodigoPrueba(ml_IdPruebaSeleccionada)
      If ml_IdPruebaSeleccionada <> "" And Trim(ml_resultado) <> "" Then
         lbExisteAlgunResultado = True
         Exit Do
      End If
      rsTmp.MoveNext
    Loop
    
    If lbExisteAlgunResultado = True Then
    rsTmp.MoveFirst
    If ml_PuntoCarga <> 11 Then mo_ReglasLaboratorio.LabImprimeCabeceraResultados UserControl.Text3.Text, UserControl.Text1.Text, UserControl.Text4.Text, Now, Trim(ml_NombreMedico) & "  " & lcServicioActualPaciente
    
    Do While Not rsTmp.EOF
      ml_IdPruebaSeleccionada = rsTmp!Codigo
      ml_NombrePruebaSeleccionada = rsTmp("nombre")
      ml_idRegistroSeleccionado = rsTmp("idOrden")
      ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(ml_IdPruebaSeleccionada, CDbl(ml_idRegistroSeleccionado))
      ml_observacion = mo_ReglasLaboratorio.LabRecuperaResultados_Obs(ml_IdPruebaSeleccionada, CDbl(ml_idRegistroSeleccionado))
      'Dim ldFechaResultado As Date
      ml_realiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(ml_IdPruebaSeleccionada, CDbl(ml_idRegistroSeleccionado), ldFechaResultado)
      ml_CodigoPruebaSeleccionada = mo_ReglasLaboratorio.LabAveriguaCodigoPrueba(ml_IdPruebaSeleccionada)
  
      If ml_IdPruebaSeleccionada <> "" And Trim(ml_resultado) <> "" Then
        If Left(ml_CodigoPruebaSeleccionada, 3) = "BQM" Then
          'Bioquímica
          mo_ReglasLaboratorio.LabImprimeResultadosBQM ml_resultado, CStr(ml_CodigoPruebaSeleccionada), ml_NombrePruebaSeleccionada, ml_observacion, ml_realiza
        ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "HEM" Then
          'Hematología
          mo_ReglasLaboratorio.LabImprimeResultadosHEM ml_resultado, CStr(ml_CodigoPruebaSeleccionada), ml_NombrePruebaSeleccionada, ml_observacion, ml_realiza
        ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "INM" Then
          'Inmunoserología
          mo_ReglasLaboratorio.LabImprimeResultadosINM ml_resultado, CStr(ml_CodigoPruebaSeleccionada), ml_NombrePruebaSeleccionada, ml_observacion, ml_realiza
        ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "MIC" Then
          'Microbiología
          mo_ReglasLaboratorio.LabImprimeResultadosMIC ml_resultado, CStr(ml_CodigoPruebaSeleccionada), ml_NombrePruebaSeleccionada, ml_observacion, ml_realiza
        ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "COA" Then
          'Parasitología
          mo_ReglasLaboratorio.LabImprimeResultadosCOA ml_resultado, CStr(ml_CodigoPruebaSeleccionada), ml_NombrePruebaSeleccionada, ml_observacion, ml_realiza
        ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "ANA" Then
          'Urianálisis
          mo_ReglasLaboratorio.LabImprimeResultadosANA ml_resultado, CStr(ml_CodigoPruebaSeleccionada), ml_NombrePruebaSeleccionada, ml_observacion, ml_realiza
        ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "CPA" Then
          'Citopatología
          mo_ReglasLaboratorio.LabImprimeResultadosCPA ml_resultado, CStr(ml_CodigoPruebaSeleccionada), ml_NombrePruebaSeleccionada, ml_observacion, ml_realiza
        ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "PAQ" Then
          'Anatomía Patológica
          mo_ReglasLaboratorio.LabImprimeResultadosPAQ ml_resultado, CStr(ml_CodigoPruebaSeleccionada), ml_NombrePruebaSeleccionada, ml_observacion, ml_realiza
        ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "BSA" Then
          'Banco de Sangre
          mo_ReglasLaboratorio.LabImprimeResultadosBS ml_resultado, CStr(ml_CodigoPruebaSeleccionada), ml_NombrePruebaSeleccionada, ml_observacion, ml_realiza
        End If
      End If
      rsTmp.MoveNext
    Loop
    mo_ReglasLaboratorio.LabImprimePieResultados
    End If
    '
    'cmdImprimir_LabResultadosItems
    mo_ReglasLaboratorio.Imprimir_LabResultadosItems rsTmp, lcEdadEnAtencion, Text4.Text, lcServicioActualPaciente
  End If
  rsTmp.Filter = ""
End Sub

Private Sub grdListaOrdenes_AfterRowActivate()
  'grdListaOrdenes_Click
End Sub

Private Sub grdListaOrdenes_Click()
  ml_SeEligioGridBoleta = False
  Dim rsRecordset As ADODB.Recordset
  'Dim rsTmp As ADODB.Recordset
  'Dim rsTmp1 As Recordset
  
  Text1.Text = ""
  Text2.Text = ""
  Text3.Text = ""
  Text4.Text = ""
  ml_idRegistroSeleccionado = 0
  ml_IdPaciente = 0
  ml_NombreMedico = ""
  lblT.Caption = "0.00"
  lblPReg.Caption = "0"
  lblPRes.Caption = "0"
  
  Set rsRecordset = grdListaOrdenes.DataSource
  If rsRecordset.State = adStateClosed Then Exit Sub
  If Not (rsRecordset.EOF = True And rsRecordset.BOF = True) Then
'    If rsRecordset("IdLabEstado") = "0" Then
'      Set grdListaOrdenesDetalle.DataSource = Nothing
'      MsgBox "Esta Orden fue anulada.", vbInformation, "SIGH "
'      Exit Sub
'    End If
    If rsRecordset("IdLabEstado") = "0" Then
      grdListaOrdenesDetalle.Enabled = False
      MsgBox "Esta Orden fue anulada.", vbInformation, "SIGH "
    Else
        grdListaOrdenesDetalle.Enabled = True
    End If
    
    mo_ReglasLaboratorio.ResultadosAutomaticosActualizaHaciaGalenhos rsRecordset!IdOrden
    
    ml_idRegistroSeleccionado = rsRecordset("IdOrden")
    idRegistroSeleccionado = rsRecordset("IdMovimiento")
    ml_idOrdenLab = rsRecordset("idmovimiento")
    ml_IdPaciente = IIf(IsNull(rsRecordset("idpaciente")), 0, rsRecordset("idpaciente"))
    'Actualizado 30102014 yamill palomino
    md_fechaNacimiento = IIf(IsNull(rsRecordset("fechanacimiento")), 0, rsRecordset("fechanacimiento"))
    If Not IsNull(rsRecordset("ordenaPrueba")) Then
      ml_NombreMedico = rsRecordset("ordenaPrueba")
    Else
      ml_NombreMedico = ""
    End If
    '
    lblIdOrden.Caption = "IDENTIFICADOR: " & Trim(Str(rsRecordset!IdOrden))
    '
    lcEdadEnAtencion = ""
    If md_fechaNacimiento <> 0 Then
       Dim oEdad As Edad
       oEdad = sighentidades.CalcularEdad(md_fechaNacimiento, rsRecordset!FechaDespacho)
       lcEdadEnAtencion = oEdad.Edad & " " & oEdad.NombreEdad
    End If
    '
    lcServicioActualPaciente = mo_ReglasLaboratorio.DevuelveDatosParaImpresionResultadoLaboratorio(rsRecordset("IdOrden"))
    '
    Set rs = mo_ReglasFacturacion.LabFacturacionServicioDespachoFiltraPorIdOrdenIdPuntoCarga(rsRecordset("IdOrden"), ml_PuntoCarga, rsRecordset("IdMovimiento"))  'Set rs = mo_ReglasFacturacion.FacturacionServicioDespachoFiltraPorIdOrden(ml_IdRegistroSeleccionado)
    If rs.State = adStateOpen Then
      If rsTmp.State = adStateOpen Then Set rsTmp = Nothing
      With rsTmp
        .Fields.Append "Imprime", adBoolean
        .Fields.Append "NroOrden", adDouble
        .Fields.Append "idOrden", adDouble
        .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
        .Fields.Append "Nombre", adVarChar, 50, adFldIsNullable
        .Fields.Append "idProducto", adDouble
        .Fields.Append "Cantidad", adInteger
        .Fields.Append "Precio", adDouble
        .Fields.Append "Total", adDouble
        .Fields.Append "idPuntoCarga", adDouble
        .Fields.Append "Resultado", adVarChar, 2, adFldIsNullable
        .Fields.Append "ResultadoAutomatico", adBoolean
        .Fields.Append "ObsReceta", adVarChar, 300, adFldIsNullable
        .LockType = adLockOptimistic
        .Open
      End With
  
      Dim Tot As Double
      Dim TotP As Integer
      Dim TotRes As Integer
      Dim T As Integer
      
      Tot = 0: TotP = 0: TotRes = 0: T = 0
      If rs.RecordCount > 0 Then
      rs.MoveFirst
      Do While Not rs.EOF
      
        Tot = Tot + rs!Cantidad * rs!Precio
        TotP = TotP + 1
        rsTmp.AddNew
        T = T + 1
        rsTmp!NroOrden = T
        rsTmp!IdOrden = rs!IdOrden
        rsTmp!idProducto = rs!idProducto
        rsTmp!Cantidad = rs!Cantidad
        rsTmp!Precio = rs!Precio
        rsTmp!Total = rs!Total
        rsTmp!Codigo = rs!Codigo
        rsTmp!nombre = Left(rs!nombre, 50)
        rsTmp!idPuntoCarga = rs!idPuntoCarga
        If mo_ReglasLaboratorio.PruebaTieneResultado(rsTmp!Codigo, rsTmp!IdOrden) = True Then
          rsTmp!resultado = "SI"
          rsTmp!Imprime = True
          TotRes = TotRes + 1
        Else
          rsTmp!resultado = "NO"
        End If
        If Not IsNull(rs!ResultadoAutomatico) Then
           If rs!ResultadoAutomatico = 1 Then
              rsTmp!ResultadoAutomatico = True
           End If
        End If
        rsTmp.Update
        rs.MoveNext
      Loop
      End If
      lblT.Caption = Format(Tot, "0.00")
      lblPReg.Caption = Format(TotP, "0")
      lblPRes.Caption = Format(TotRes, "0")
      '*********Proviene de una Receta (inicio)
      lblEPSpago.Caption = ""
      If rsRecordset!idCuentaAtencion > 0 Then
      
            Dim oRsTmp1 As New Recordset
            Dim lcBoletaEPS  As String, lcOrdenPago
            Dim oConexion As New Connection
            oConexion.CommandTimeout = 300
            oConexion.CursorLocation = adUseClient
            oConexion.Open sighentidades.CadenaConexion
            '
            Set oRsTmp1 = mo_reglasComunes.AtencionesSeleccionarMedicoPorCuenta(rsRecordset!idCuentaAtencion)
            If oRsTmp1.RecordCount > 0 Then
                lcOrdenPago = mo_ReglasFacturacion.DevuelveOrdenPago(oRsTmp1!idAtencion, sghPtoCargaCaja, _
                                                                     rsRecordset!FechaDespacho, oConexion, lcBoletaEPS)
                If lcBoletaEPS <> "" Then
                    lblEPSpago.Caption = "Pagó EPS - " & lcBoletaEPS
                    lblEPSpago.ForeColor = vbBlack
                ElseIf lcOrdenPago <> "" Then
                    lblEPSpago.Caption = "NO PAGO EPS - N°OrdenPago: " & lcOrdenPago
                    lblEPSpago.ForeColor = vbRed
                End If
            End If
            '
            Dim lnidReceta As Long
            lnidReceta = 0
            Set oRsTmp1 = mo_reglasComunes.RecetaCabeceraFiltraXcuentaYDocumentodespacho(Trim(Str(rsRecordset!IdMovimiento)), rsRecordset!idCuentaAtencion)
            If oRsTmp1.RecordCount > 0 Then
               lnidReceta = oRsTmp1.Fields!idReceta
            End If
            oRsTmp1.Close
            If lnidReceta > 0 Then
                Set oRsTmp1 = mo_reglasComunes.RecetaDetalleSeleccioarPorIdReceta(lnidReceta, oConexion)
                oRsTmp1.Filter = "observaciones<>''"
                If oRsTmp1.RecordCount > 0 Then
                   oRsTmp1.MoveFirst
                   Do While Not oRsTmp1.EOF
                      If Not IsNull(oRsTmp1!Observaciones) Then
                         rsTmp.MoveFirst
                         rsTmp.Find "idProducto=" & oRsTmp1!idItem
                         If Not rsTmp.EOF Then
                            rsTmp!obsReceta = oRsTmp1!Observaciones
                            rsTmp.Update
                         End If
                      End If
                      oRsTmp1.MoveNext
                   Loop
                End If
                oRsTmp1.Close
             End If
             oConexion.Close
             Set oConexion = Nothing
             Set oRsTmp1 = Nothing
       End If
       '*********Proviene de una Receta (fin)
    End If
    'Set rsTmp1 = rsTmp
    Set rs = Nothing
    Set grdListaOrdenesDetalle.DataSource = rsTmp
    'rsTmp.Close
    'mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenesDetalle, sighentidades.GrillaConFilasBicolor
    '
    'Text1.Text = UCase(rsRecordset("ApellidoPaterno") & " " & rsRecordset("ApellidoMaterno")) & ", " & mo_Teclado.CapitalizarNombres(rsRecordset("PrimerNombre"))
    Text1.Text = UCase(rsRecordset("Paciente"))
    If Not IsNull(rsRecordset("fechaNacimiento")) Then
       ldFechaNacimiento = rsRecordset("fechaNacimiento")
    End If
    If Not IsNull(rsRecordset("idTipoSexo")) Then
       lnIdTipoSexo = rsRecordset("idTipoSexo")
    End If
    '
    If rsRecordset("idpaciente") <> "" Then Text3.Text = rsRecordset("idpaciente")
    Text4.Text = IIf(IsNull(rsRecordset("NroHistoriaClinica")), "", rsRecordset("NroHistoriaClinica"))
    If rsTmp.RecordCount > 0 Then
       rsTmp.MoveFirst
    End If
  End If
End Sub

Private Sub grdListaOrdenes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
  grdListaOrdenes.Bands(0).Columns("IdPuntoCarga").Hidden = True
  grdListaOrdenes.Bands(0).Columns("IdOrden").Hidden = True
  grdListaOrdenes.Bands(0).Columns("IdPaciente").Hidden = True
  grdListaOrdenes.Bands(0).Columns("IdEstadoFacturacion").Hidden = True
  grdListaOrdenes.Bands(0).Columns("ApellidoPaterno").Hidden = True
  grdListaOrdenes.Bands(0).Columns("ApellidoMaterno").Hidden = True
  grdListaOrdenes.Bands(0).Columns("PrimerNombre").Hidden = True
  grdListaOrdenes.Bands(0).Columns("IdOrden").Header.Caption = "Nº Orden"
  grdListaOrdenes.Bands(0).Columns("IdOrden").Width = 1000
  grdListaOrdenes.Bands(0).Columns("FechaDespacho").Header.Caption = "F.Registro"
  grdListaOrdenes.Bands(0).Columns("FechaDespacho").Width = 2500
  grdListaOrdenes.Bands(0).Columns("FechaCreacion").Hidden = True
  grdListaOrdenes.Bands(0).Columns("idCuentaAtencion").Header.Caption = "Nº Cuenta"
  grdListaOrdenes.Bands(0).Columns("idCuentaAtencion").Width = 1000
  grdListaOrdenes.Bands(0).Columns("idPaciente").Header.Caption = "Id Paciente"
  grdListaOrdenes.Bands(0).Columns("idPaciente").Width = 1000
  grdListaOrdenes.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nº H.C."
  grdListaOrdenes.Bands(0).Columns("NroHistoriaClinica").Width = 1000
  grdListaOrdenes.Bands(0).Columns("Paciente").Width = 3500
  grdListaOrdenes.Bands(0).Columns("EstadoOrden").Header.Caption = "Estado Orden"
  grdListaOrdenes.Bands(0).Columns("EstadoOrden").Width = 1500
  grdListaOrdenes.Bands(0).Columns("idmovimiento").Header.Caption = "Nº Movimiento"
  grdListaOrdenes.Bands(0).Columns("idmovimiento").Width = 1000
  grdListaOrdenes.Bands(0).Columns("idlabEstado").Hidden = True
  grdListaOrdenes.Bands(0).Columns("ordenaPrueba").Header.Caption = "Ordena Prueba"
  grdListaOrdenes.Bands(0).Columns("ordenaPrueba").Width = 3000
End Sub

Private Sub grdListaOrdenes_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
  If Val(Row.Cells("IdLabEstado").GetText()) = 0 Then Row.Appearance.ForeColor = vbRed
End Sub

Private Sub grdListaOrdenes_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn
        grdListaOrdenes_Click
    Case vbKeyDown, vbKeyUp
       ' RefrescarDatos
  End Select
End Sub

Private Sub grdListaOrdenesDetalle_DblClick()
  On Error GoTo Fin
  If rsTmp!ResultadoAutomatico = True Then
     Exit Sub
  End If
  Dim lbEsAntiguoFormato As Boolean
  lbEsAntiguoFormato = True
  If mo_ReglasLaboratorio.UsaNuevaVentanaResultadosLaboratorio(rsTmp!IdOrden, rsTmp!Codigo) = True Then
    '************(inicio) usa el nuevo formulario para llenar e imprimir RESULTADOS **********************
    Dim oRsTmp1 As New Recordset
    Dim lbSoloTieneOpcionConsulta As Boolean
    lbSoloTieneOpcionConsulta = mo_ReglasDeSeguridad.SoloTieneOpcionCONSULTA(ml_idUsuario, mo_lnIdTablaLISTBARITEMS)
    Set oRsTmp1 = mo_ReglasLaboratorio.LabItemsCptSeleccionarXfiltro("dbo.FactCatalogoServicios.Codigo='" & rsTmp("codigo") & "'")
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.Close
       
       Dim oResultadoXitems As New SIGHLaboratorio.ResultadoXitems
       oResultadoXitems.IdOrden = rsTmp("idOrden")
       oResultadoXitems.idProductoCpt = rsTmp("idProducto")
       oResultadoXitems.idUsuario = ml_idUsuario
       oResultadoXitems.NoMuestraBotonGrabar = lbSoloTieneOpcionConsulta  'debb2014d
       oResultadoXitems.lcNombrePc = mo_lcNombrePc
       oResultadoXitems.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
       oResultadoXitems.MostrarFormulario
       Set oResultadoXitems = Nothing
       Set oRsTmp1 = Nothing
       lbEsAntiguoFormato = False
       Exit Sub
    Else
       lbEsAntiguoFormato = True
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
    '************(fin) usa el nuevo formulario para llenar e imprimir RESULTADOS **********************
  End If
  If lbEsAntiguoFormato = True Then
    'Cargar los formularios para el resultado
    ml_IdPruebaSeleccionada = rsTmp("codigo")
    ml_NombrePruebaSeleccionada = rsTmp("nombre")
    ml_nombrePaciente = Text1.Text
    ml_idOrden = rsTmp("idOrden")
    
    'debb2014d
    Dim oMuestraResultado As New SIGHLaboratorio.Ingresos
    oMuestraResultado.MuestraResultadoDelExamen ml_IdPruebaSeleccionada, ml_NombrePruebaSeleccionada, _
                                                ml_nombrePaciente, ml_idOrden, ml_IdPaciente, ml_NombreMedico, _
                                                ml_areaTrabajo, ml_idOrdenLab, lnIdTipoSexo, lbSoloTieneOpcionConsulta, grdListaOrdenesDetalle.DataSource, md_fechaNacimiento
    Set oMuestraResultado = Nothing
  End If

  Exit Sub
  
Fin:
  Exit Sub
  
End Sub




Private Sub grdListaOrdenesDetalle_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
  Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
  
  grdListaOrdenesDetalle.Bands(0).Columns("idOrden").Hidden = True
  grdListaOrdenesDetalle.Bands(0).Columns("ResultadoAutomatico").Activation = ssActivationActivateNoEdit
  
  grdListaOrdenesDetalle.Bands(0).Columns("NroOrden").Header.Caption = "Nº"
  grdListaOrdenesDetalle.Bands(0).Columns("NroOrden").Width = 700
  grdListaOrdenesDetalle.Bands(0).Columns("NroOrden").Activation = ssActivationActivateNoEdit ' = ssActivationAllowEdit
    
  grdListaOrdenesDetalle.Bands(0).Columns("idProducto").Hidden = True
  grdListaOrdenesDetalle.Bands(0).Columns("idProducto").Header.Caption = "C.Producto"
  grdListaOrdenesDetalle.Bands(0).Columns("idProducto").Width = 900
  grdListaOrdenesDetalle.Bands(0).Columns("idProducto").Activation = ssActivationActivateNoEdit
    
  grdListaOrdenesDetalle.Bands(0).Columns("cantidad").Header.Caption = "Cantidad"
  grdListaOrdenesDetalle.Bands(0).Columns("cantidad").Width = 1000
  grdListaOrdenesDetalle.Bands(0).Columns("cantidad").Activation = ssActivationActivateNoEdit
    
  grdListaOrdenesDetalle.Bands(0).Columns("precio").Width = 1000
  grdListaOrdenesDetalle.Bands(0).Columns("precio").Header.Caption = "Precio"
  grdListaOrdenesDetalle.Bands(0).Columns("precio").Format = "#0.000"
  grdListaOrdenesDetalle.Bands(0).Columns("precio").Hidden = False
  grdListaOrdenesDetalle.Bands(0).Columns("precio").Activation = ssActivationActivateNoEdit
    
  grdListaOrdenesDetalle.Bands(0).Columns("total").Header.Caption = "Total"
  grdListaOrdenesDetalle.Bands(0).Columns("total").Format = "#0.000"
  grdListaOrdenesDetalle.Bands(0).Columns("total").Width = 1000
  grdListaOrdenesDetalle.Bands(0).Columns("total").Activation = ssActivationActivateNoEdit
    
  grdListaOrdenesDetalle.Bands(0).Columns("codigo").Header.Caption = "C.Prueba"
  grdListaOrdenesDetalle.Bands(0).Columns("codigo").Width = "1000"
  grdListaOrdenesDetalle.Bands(0).Columns("codigo").Hidden = False
  grdListaOrdenesDetalle.Bands(0).Columns("codigo").Activation = ssActivationActivateNoEdit
    
  grdListaOrdenesDetalle.Bands(0).Columns("nombre").Header.Caption = "Nombre de Prueba"
  grdListaOrdenesDetalle.Bands(0).Columns("nombre").Activation = ssActivationActivateNoEdit
  grdListaOrdenesDetalle.Bands(0).Columns("nombre").Width = 6000
  
  grdListaOrdenesDetalle.Bands(0).Columns("resultado").Width = 1000
    
  grdListaOrdenesDetalle.Bands(0).Columns("idpuntocarga").Hidden = True
  grdListaOrdenesDetalle.Bands(0).Columns("obsReceta").Width = 4200
  
  grdListaOrdenesDetalle.Bands(0).Columns("Imprime").Width = 800
  grdListaOrdenesDetalle.Bands(0).Columns("Imprime").Style = ssStyleCheckBox
End Sub

Private Sub grdListaOrdenesDetalle_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn
        grdListaOrdenesDetalle_DblClick
    Case vbKeyDown, vbKeyUp
       ' RefrescarDatos
  End Select
End Sub

Private Sub Text1_GotFocus()
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text2_GotFocus()
  Text2.SelStart = 0
  Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text3_GotFocus()
  Text3.SelStart = 0
  Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text4_GotFocus()
  Text4.SelStart = 0
  Text4.SelLength = Len(Text4.Text)
End Sub

Private Sub txtFF_Change()
 'If cmbIdPtoCarga.Text <> "" And IsDate(txtFI.Text) And IsDate(txtFF.Text) Then btnBuscar_Click
End Sub

Private Sub txtFF_GotFocus()
  txtFF.SelStart = 0
  txtFF.SelLength = Len(txtFF.Text)
End Sub

Private Sub txtFF_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If cmbIdPtoCarga.Text <> "" And IsDate(txtFI.Text) And IsDate(txtFF.Text) Then btnBuscar_Click
  End If
End Sub

Private Sub txtFF_LostFocus()
     
    If Not IsDate(txtFF.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        
        txtFF.Text = sighentidades.FECHA_VACIA_DMY_HMS
        Exit Sub
    End If
End Sub

Private Sub txtFI_Change()
  'If cmbIdPtoCarga.Text <> "" And IsDate(txtFI.Text) And IsDate(txtFF.Text) Then btnBuscar_Click
End Sub

Private Sub txtFI_GotFocus()
  txtFI.SelStart = 0
  txtFI.SelLength = Len(txtFI.Text)
End Sub

Private Sub txtFI_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    'If cmbIdPtoCarga.Text <> "" And IsDate(txtFI.Text) And IsDate(txtFF.Text) Then btnBuscar_Click
  End If
End Sub

Private Sub txtFI_LostFocus()
    If Not IsDate(txtFI.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFI.Text = sighentidades.FECHA_VACIA_DMY_HM
        Exit Sub
    End If
End Sub

Private Sub txtNCuenta_GotFocus()
  txtNcuenta.SelStart = 0
  txtNcuenta.SelLength = Len(txtNcuenta.Text)
End Sub

Private Sub txtNcuenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then btnBuscar_Click
  If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then KeyAscii = 0
  End If
End Sub

Private Sub txtNombres_GotFocus()
  txtNombres.SelStart = 0
  txtNombres.SelLength = Len(txtNombres.Text)
End Sub

Private Sub txtNombres_KeyPress(KeyAscii As Integer)
  'If KeyAscii = 13 Then btnBuscar_Click
End Sub

Private Sub txtNroHistoria_GotFocus()
  txtNroHistoria.SelStart = 0
  txtNroHistoria.SelLength = Len(txtNroHistoria.Text)
End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
  'mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then btnBuscar_Click
  If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then KeyAscii = 0
  End If
End Sub

Private Sub txtNroOrden_GotFocus()
  txtNroOrden.SelStart = 0
  txtNroOrden.SelLength = Len(txtNroOrden.Text)
End Sub

Private Sub txtNroOrden_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then btnBuscar_Click
  If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then KeyAscii = 0
  End If
End Sub

Private Sub UserControl_GotFocus()
  'btnBuscar_Click
End Sub

Private Sub UserControl_Initialize()
  ml_idRegistroSeleccionado = 0
  ml_IdPaciente = 0
  grdListaOrdenesDetalle.Caption = ""
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  'btnBuscar_Click
  AdministrarKeyPreview KeyCode
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
  Select Case KeyCode
    Case vbKeyEscape
    Case vbKeyF2
    Case vbKeyF3
    Case vbKeyF4
    Case vbKeyF5
    Case vbKeyF6
      btnBuscar_Click
    Case vbKeyF7
    Case vbKeyF8
  End Select
End Sub

Private Sub UserControl_Resize()
  On Error Resume Next
  'fraBusqueda.Width = (UserControl.Width - 110) - (2 * grdBoletas.Width) - 2400
  
  lblNombre.Width = UserControl.Width
  
  grdListaOrdenes.Width = UserControl.Width - 110
  FraDetalle.Width = grdListaOrdenes.Width
  grdListaOrdenes.Height = (UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)) / 2 - Label7.Height
  Label7.Top = grdListaOrdenes.Top + grdListaOrdenes.Height
  FraDetalle.Top = Label7.Top + 285 ' grdListaOrdenes.Top + grdListaOrdenes.Height + 100
  FraDetalle.Height = (UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)) / 2 - 150
  grdListaOrdenesDetalle.Width = FraDetalle.Width - 120
  grdListaOrdenesDetalle.Height = FraDetalle.Height - grdListaOrdenesDetalle.Top - 100 - Label6.Height
  Label6.Top = grdListaOrdenesDetalle.Top + grdListaOrdenesDetalle.Height '+10
End Sub


Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenes, "99"
        mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenesDetalle, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenes, sighentidades.GrillaConFilasBicolor
        mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenesDetalle, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub
Sub Inicializar()
  SkinConfigura
  Dim lcBuscaParametro As New SIGHDatos.Parametros
  txtFcpt1.Text = Date
  txtFCpt2.Text = Date
  If lcBuscaParametro.SeleccionaFilaParametro(531) = "S" Then
     chkPorFcpt.Value = 1
     chkPorFcpt_Click
  End If
  Set lcBuscaParametro = Nothing
  
  ConfiguraFechaHora
  'cmbFecha.Clear
  'cmbFecha.AddItem Date
  'cmbFecha.AddItem "Todas"
  'cmbFecha.ListIndex = 0
  
  ConfigurarPuntosDeCarga
  If ml_PuntoCarga = 5 Then
    cmbFarmacia.Visible = True
    lblFarmacia.Visible = True
    CargaFarmacias
  Else
    cmbFarmacia.Visible = False
    lblFarmacia.Visible = False
  End If
  mo_Formulario.HabilitarDeshabilitar Text1, False
  mo_Formulario.HabilitarDeshabilitar Text4, False
  
End Sub

Sub CargaFarmacias()
  On Error GoTo ErrFarm
  Dim oConexion As New ADODB.Connection
  Dim lnCodigoFarmacia  As Long
  Dim lcSql As String
        
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  Set oRsFarmacias = mo_ReglasDeSeguridad.UsuariosRolesXidEmpleadoEsDeFarmacia(ml_idUsuario, oConexion)

  If oRsFarmacias.RecordCount > 0 Then lnCodigoFarmacia = oRsFarmacias.Fields!IdPermiso
  oRsFarmacias.Close
  Set oRsFarmacias = mo_ReglasDeSeguridad.PermisosSoloFarmacia(oConexion)
  Set cmbFarmacia.RowSource = oRsFarmacias
  cmbFarmacia.ListField = "descripcion"
  cmbFarmacia.BoundColumn = "idPermiso"
  If lnCodigoFarmacia > 0 Then cmbFarmacia.BoundText = lnCodigoFarmacia

ErrFarm:
  'cmbFarmacia.BoundText = ""
  'cmbFarmacia.Text = ""
  'oRsFarmacias.Close
  'Resume
End Sub

Sub ConfigurarPuntosDeCarga()
  Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPtoCarga
  mo_cmbIdPuntoCarga.ListField = "Descripcion"
  mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
  Set mo_cmbIdPuntoCarga.RowSource = mo_reglasComunes.SeleccionarPuntosDeCarga()
End Sub

Private Sub UserControl_Show()
  'btnBuscar_Click
End Sub

Private Sub grdBoletas_AfterRowActivate()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdBoletas.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("idComprobantePago")
    'ml_SeEligioGridBoleta = True 'Actualizado 16092014
End Sub

Private Sub grdBoletas_Click()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdBoletas.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdComprobantePago")
    ml_SeEligioGridBoleta = True

End Sub

Private Sub grdBoletas_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdBoletas.Bands(0).Columns("IdComprobantePago").Hidden = True
    grdBoletas.Bands(0).Columns("Boleta").Width = 1300
    grdBoletas.Bands(0).Columns("total").Width = 500
    grdBoletas.Bands(0).Columns("Fecha").Header.Caption = "Fecha"
    grdBoletas.Bands(0).Columns("Fecha").Width = 1500
    grdBoletas.Bands(0).Columns("Fecha").Format = sighentidades.DevuelveFechaSoloFormato_DMY  'debb-10/08/2016
    mo_Apariencia.ConfigurarFilasBiColores grdBoletas, sighentidades.GrillaConFilasBicolor
End Sub

'debb-28/03/2016
'Private Sub cmdImprimir_LabResultadosItems()
'        Dim lbHalloUnCPT As Boolean
'        Dim oRsResultados As New Recordset
'        Dim oRsResultadosCPT As New Recordset
'        Dim oRsTmp1 As New Recordset
'        Dim oRsTmp2 As New Recordset
'        Dim lcBuscaParametro As New SIGHDatos.Parametros
'        Dim lcSql As String, lcTexto As String
'        Dim lnIdProducto As Long, lnIdItem As Long, ldFechaResultado As Date, lnCantidadCPT As Integer
'        Dim lcPaciente As String, ml_idTipoSexo As Long, ml_FechaNacimiento As Date, lcMedico As String
'        lbHalloUnCPT = False
'        If rsTmp.RecordCount > 0 Then
'            With oRsResultados
'                  .Fields.Append "idProducto", adInteger
'                  .Fields.Append "SoloNumero", adBoolean
'                  .Fields.Append "SoloTexto", adBoolean
'                  .Fields.Append "SoloCombo", adBoolean
'                  .Fields.Append "SoloCheck", adBoolean
'                  .Fields.Append "ordenXresultado", adInteger, 4, adFldIsNullable
'                  .Fields.Append "Grupo", adVarChar, 100, adFldIsNullable
'                  .Fields.Append "Item", adVarChar, 100, adFldIsNullable
'                  .Fields.Append "idItem", adInteger
'                  .Fields.Append "ValorNumero", adDouble
'                  .Fields.Append "ValorTexto", adVarChar, 500, adFldIsNullable
'                  .Fields.Append "ValorCombo", adVarChar, 100, adFldIsNullable
'                  .Fields.Append "ValorCheck", adVarChar, 1, adFldIsNullable
'                  .Fields.Append "ValorReferencial", adVarChar, 100, adFldIsNullable
'                  .Fields.Append "Metodo", adVarChar, 100, adFldIsNullable
'                  .LockType = adLockOptimistic
'                  .Open
'           End With
'
'           rsTmp.MoveFirst
'           Do While Not rsTmp.EOF
'              Set oRsTmp1 = mo_ReglasLaboratorio.LabResultadosPorItemsSeleccionarPORfiltro("dbo.LabResultadoPorItems.idOrden=" & rsTmp("idOrden") & _
'                                                 " and dbo.LabResultadoPorItems.idProductoCpt=" & rsTmp!idProducto)
'              If oRsTmp1.RecordCount > 0 Then
'                   oRsTmp1.MoveFirst
'                   If lbHalloUnCPT = False Then
'                        'Medico que ordena
'                        Set oRsTmp2 = mo_ReglasLaboratorio.LabMovimientoLaboratorioSeleccionarXidOrden(rsTmp("idOrden"))
'                        lcPaciente = "": ml_idTipoSexo = 0: ml_FechaNacimiento = 0: lcMedico = ""
'                        If oRsTmp2.RecordCount > 0 Then
'                           ml_idTipoSexo = IIf(IsNull(oRsTmp2.Fields!idTipoSexo), 1, oRsTmp2.Fields!idTipoSexo)
'                           ml_FechaNacimiento = IIf(IsNull(oRsTmp2.Fields!FechaNacimiento), Date, oRsTmp2.Fields!FechaNacimiento)
'                           If lcEdadEnAtencion = "" Then
'                              lcEdadEnAtencion = "F.Nacim: " & ml_FechaNacimiento
'                           End If
'                           lcMedico = oRsTmp2.Fields!OrdenaPrueba
'                           lcPaciente = "(" & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Text4.Text, False) & ")" & oRsTmp2.Fields!Paciente
'                        End If
'                    End If
'                    lbHalloUnCPT = True
'                    'Llena Temporal
'                    oRsResultados.AddNew
'                    oRsResultados.Fields!idProducto = rsTmp!idProducto
'                    oRsResultados.Fields!ValorTexto = "ANALISIS: " & rsTmp!Codigo & " - " & rsTmp!Nombre
'                    oRsResultados.Update
'                    oRsResultados.AddNew
'                    oRsResultados.Fields!idProducto = rsTmp!idProducto
'                    oRsResultados.Fields!ValorTexto = "Personal q realizó la prueba: " & mo_reglasComunes.EmpleadosDevuelveNombre(oRsTmp1!realizaAnalisis)
'                    oRsResultados.Update
'                    oRsResultados.AddNew
'                    oRsResultados.Fields!idProducto = rsTmp!idProducto
'                    oRsResultados.Fields!ValorTexto = "F.resultado: " & oRsTmp1!fecha
'                    oRsResultados.Update
'                    Do While Not oRsTmp1.EOF
'                        oRsResultados.AddNew
'                        oRsResultados.Fields!idProducto = rsTmp!idProducto
'                        oRsResultados.Fields!ordenXresultado = oRsTmp1!ordenXresultado
'                        oRsResultados.Fields!Grupo = oRsTmp1!nombreGrupo
'                        oRsResultados.Fields!Item = oRsTmp1!Item
'                        oRsResultados.Fields!idItem = oRsTmp1!idItem
'                        oRsResultados.Fields!ValorReferencial = oRsTmp1!ValorReferencial
'                        oRsResultados.Fields!Metodo = oRsTmp1!Metodo
'                        If Not IsNull(oRsTmp1.Fields!ValorNumero) Then
'                           oRsResultados.Fields!ValorNumero = oRsTmp1.Fields!ValorNumero
'                           oRsResultados.Fields!SoloNumero = True
'                        End If
'                        If Not IsNull(oRsTmp1.Fields!ValorTexto) Then
'                           oRsResultados.Fields!ValorTexto = oRsTmp1.Fields!ValorTexto
'                           oRsResultados.Fields!Solotexto = True
'                        End If
'                        If Not IsNull(oRsTmp1.Fields!ValorCombo) Then
'                           oRsResultados.Fields!ValorCombo = oRsTmp1.Fields!ValorCombo
'                           oRsResultados.Fields!SoloCombo = True
'                        End If
'                        If Not IsNull(oRsTmp1.Fields!ValorCheck) Then
'                           oRsResultados.Fields!ValorCheck = oRsTmp1.Fields!ValorCheck
'                           oRsResultados.Fields!SoloCheck = True
'                        End If
'                        oRsResultados.Update
'                        oRsTmp1.MoveNext
'                    Loop
'              End If
'              rsTmp.MoveNext
'           Loop
'
''              Set oRsTmp1 = mo_ReglasLaboratorio.LabResultadosPorItemsSeleccionarXfiltro("idOrden=" & rsTmp("idOrden") & _
''                                                 " and idProductoCpt=" & rsTmp!idProducto)
''              If oRsTmp1.RecordCount > 0 Then
''                    If lbHalloUnCPT = False Then
''                        'Medico que ordena
''                        Set oRsTmp2 = mo_ReglasLaboratorio.LabMovimientoLaboratorioSeleccionarXidOrden(rsTmp("idOrden"))
''                        lcPaciente = "": ml_idTipoSexo = 0: ml_FechaNacimiento = 0: lcMedico = ""
''                        If oRsTmp2.RecordCount > 0 Then
''                           ml_idTipoSexo = IIf(IsNull(oRsTmp2.Fields!idTipoSexo), 1, oRsTmp2.Fields!idTipoSexo)
''                           ml_FechaNacimiento = IIf(IsNull(oRsTmp2.Fields!FechaNacimiento), Date, oRsTmp2.Fields!FechaNacimiento)
''                           If lcEdadEnAtencion = "" Then
''                              lcEdadEnAtencion = "F.Nacim: " & ml_FechaNacimiento
''                           End If
''                           lcMedico = oRsTmp2.Fields!OrdenaPrueba
''                           lcPaciente = "(" & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Text4.Text, False) & ")" & oRsTmp2.Fields!Paciente
''                        End If
''                    End If
''                    lbHalloUnCPT = True
''                    'Llena Temporal
''                    Set oRsResultadosCPT = mo_ReglasLaboratorio.LabItemsCptSeleccionarXfiltro("dbo.LabItemsCpt.idProductoCpt=" & rsTmp!idProducto)
''                    If oRsResultadosCPT.RecordCount > 0 Then
''                       oRsResultadosCPT.MoveFirst
''                       oRsResultados.AddNew
''                       oRsResultados.Fields!idProducto = rsTmp!idProducto
''                       oRsResultados.Fields!ValorTexto = "ANALISIS: " & oRsResultadosCPT.Fields!Codigo & " - " & oRsResultadosCPT.Fields!Nombre
''                       oRsResultados.Update
''                       oRsResultados.AddNew
''                       oRsResultados.Fields!idProducto = rsTmp!idProducto
''                       oRsResultados.Fields!ValorTexto = "Personal q realizó la prueba: " & mo_reglasComunes.EmpleadosDevuelveNombre(oRsTmp1.Fields!realizaAnalisis)
''                       oRsResultados.Update
''                       oRsResultados.AddNew
''                       oRsResultados.Fields!idProducto = rsTmp!idProducto
''                       oRsResultados.Fields!ValorTexto = "F.resultado: " & oRsTmp1.Fields!fecha
''                       oRsResultados.Update
''                       Do While Not oRsResultadosCPT.EOF
''                          lnIdItem = oRsResultadosCPT.Fields!idItem
''                          oRsResultados.AddNew
''                          oRsResultados.Fields!idProducto = rsTmp!idProducto
''                          oRsResultados.Fields!ordenXresultado = oRsResultadosCPT.Fields!ordenXresultado
''                          oRsResultados.Fields!Grupo = oRsResultadosCPT.Fields!Grupo
''                          oRsResultados.Fields!Item = oRsResultadosCPT.Fields!Item
''                          oRsResultados.Fields!idItem = oRsResultadosCPT.Fields!idItem
''                          oRsResultados.Fields!ValorReferencial = oRsResultadosCPT.Fields!ValorReferencial
''                          oRsResultados.Fields!Metodo = oRsResultadosCPT.Fields!Metodo
''                          oRsResultados.Fields!SoloNumero = IIf(oRsResultadosCPT.Fields!SoloNumero = True, True, False)
''                          oRsResultados.Fields!Solotexto = IIf(oRsResultadosCPT.Fields!Solotexto = True, True, False)
''                          oRsResultados.Fields!SoloCombo = IIf(oRsResultadosCPT.Fields!SoloCombo = True, True, False)
''                          oRsResultados.Fields!SoloCheck = IIf(oRsResultadosCPT.Fields!SoloCheck = True, True, False)
''                          oRsResultados.Update
''                          Do While Not oRsResultadosCPT.EOF And lnIdItem = oRsResultadosCPT.Fields!idItem
''                             oRsResultadosCPT.MoveNext
''                             If oRsResultadosCPT.EOF Then
''                                Exit Do
''                             End If
''                          Loop
''                       Loop
''                     End If
''                     'Llena Resultados del Paciente
''                     oRsTmp1.MoveFirst
''                     ldFechaResultado = oRsTmp1.Fields!fecha
''                     Do While Not oRsTmp1.EOF
''                       oRsResultados.MoveFirst
''                       oRsResultados.Find "ordenXresultado=" & oRsTmp1.Fields!ordenXresultado
''                       If Not oRsResultados.EOF Then
''                           Do While Not oRsResultados.EOF
''                              If oRsResultados.Fields!idProducto = rsTmp!idProducto And oRsResultados.Fields!ordenXresultado = oRsTmp1.Fields!ordenXresultado Then
''                                 Exit Do
''                              End If
''                              oRsResultados.MoveNext
''                           Loop
''                           If Not IsNull(oRsTmp1.Fields!ValorNumero) Then
''                              oRsResultados.Fields!ValorNumero = oRsTmp1.Fields!ValorNumero
''                           End If
''                           If Not IsNull(oRsTmp1.Fields!ValorTexto) Then
''                              oRsResultados.Fields!ValorTexto = oRsTmp1.Fields!ValorTexto
''                           End If
''                           If Not IsNull(oRsTmp1.Fields!ValorCombo) Then
''                              oRsResultados.Fields!ValorCombo = oRsTmp1.Fields!ValorCombo
''                           End If
''                           If Not IsNull(oRsTmp1.Fields!ValorCheck) Then
''                              oRsResultados.Fields!ValorCheck = oRsTmp1.Fields!ValorCheck
''                           End If
''                           oRsResultados.Update
''                       End If
''                       oRsTmp1.MoveNext
''                     Loop
''              End If
''              rsTmp.MoveNext
''           Loop
'
'
'
'
'
'        End If
'        If lbHalloUnCPT = False Then
'           Exit Sub
'        End If
'        '
'        Dim iFila As Long, iColumna As Integer
'        Dim lbEsOpenOffice As Boolean
'
'        lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
'        On Error GoTo ManejadorError
'
'        If lbEsOpenOffice = True Then
'            Dim ServiceManager As Object
'            Dim Desktop As Object
'            Dim Document As Object
'            Dim Feuille As Object
'            Dim Plage As Object
'            Dim args()
'            Dim Chemin As String
'            Dim Fichier As String
'            Dim lcArchivoExcel As String
'            Dim PrintArea(0)
'            Dim Style As Object
'            Dim Border As Object
'            'encabezado
'            Dim PageStyles As Object
'            Dim Sheet As Object
'            Dim StyleFamilies As Object
'            Dim DefPage As Object
'            Dim Htext As Object
'            Dim Hcontent As Object
'            Dim lnHWnd As Long
'            Dim ret As Long
'        Else
'            Dim oExcel As Excel.Application
'            Dim oWorkBookPlantilla As Workbook
'            Dim oWorkBook As Workbook
'            Dim oWorkSheet As Worksheet
'        End If
'
'        If lbEsOpenOffice = True Then
'            'Abre el archivo ExcelOpenOffice
'            lcArchivoExcel = App.Path + "\Plantillas\LabResultadoXitem.ods"
'            Fichier = Format(Time, "hhmmss") & ".ods"
'            FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
'            lcArchivoExcel = Fichier
'            Chemin = "file:///" & App.Path & "\Plantillas\"
'            Chemin = Replace(Chemin, "\", "/")
'            Fichier = Chemin & "/" & lcArchivoExcel
'            Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
'            Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
'            Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
'            Set Feuille = Document.getSheets().getByIndex(0)
'            'Encabezado de Pagina
'            mo_CabeceraReportes.CabeceraReportes Document, True
'            ' Pone la ventana en primer plano, pasándole el Hwnd
'            ret = SetForegroundWindow(lnHWnd)
'        Else
'            'Crea nueva hoja
'            Set oExcel = GalenhosExcelApplication()
'            Set oWorkBook = oExcel.Workbooks.Add
'            'Abre, copia y cierra la plantilla
'            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\LabResultadoXitem.xls")
'
'            oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
'            oWorkBookPlantilla.Close
'            'Activa la primera hoja
'            Set oWorkSheet = oWorkBook.Sheets(1)
'            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
'        End If
'        If lbEsOpenOffice = True Then
'            Call Feuille.getcellbyposition(3, 0).setFormula(lcPaciente & _
'                                       " (Sexo: " & IIf(ml_idTipoSexo = 1, "Masculino", "Femenino") & ") ( " & _
'                                       lcEdadEnAtencion & ")  " & lcServicioActualPaciente)
'            Call Feuille.getcellbyposition(3, 1).setFormula(lcMedico)
'            Call Feuille.getcellbyposition(5, 1).setFormula("")
'            Call Feuille.getcellbyposition(2, 2).setFormula("")
'        Else
'            oWorkSheet.Cells(1, 4).Value = lcPaciente & _
'                                           " (Sexo: " & IIf(ml_idTipoSexo = 1, "Masculino", "Femenino") & ") (" & _
'                                           lcEdadEnAtencion & ")  " & lcServicioActualPaciente
'            oWorkSheet.Cells(2, 4).Value = lcMedico
'            oWorkSheet.Cells(2, 6).Value = ""
'            oWorkSheet.Cells(3, 3).Value = " "
'        End If
'        iFila = 9
'        lnCantidadCPT = 0
'        If oRsResultados.RecordCount > 0 Then
'           oRsResultados.MoveFirst
'           Do While Not oRsResultados.EOF
'              lnCantidadCPT = lnCantidadCPT + 1
'              lnIdProducto = oRsResultados.Fields!idProducto
'              Do While Not oRsResultados.EOF And lnIdProducto = oRsResultados.Fields!idProducto
'                    If lbEsOpenOffice = True Then
'                        Call Feuille.getcellbyposition(3, 0).setFormula(IIf(IsNull(oRsResultados.Fields!Grupo), "", oRsResultados.Fields!Grupo))
'                        Call Feuille.getcellbyposition(3, 0).setFormula(IIf(IsNull(oRsResultados.Fields!Item), "", oRsResultados.Fields!Item))
'                    Else
'                        oWorkSheet.Cells(iFila, 2).Value = oRsResultados.Fields!Grupo
'                        oWorkSheet.Cells(iFila, 3).Value = oRsResultados.Fields!Item
'                    End If
'
'                    lcTexto = ""
'                    If oRsResultados.Fields!ValorNumero > 0 Then
'                      lcTexto = lcTexto & Trim(Str(IIf(IsNull(oRsResultados.Fields!ValorNumero), "", oRsResultados.Fields!ValorNumero))) & "| "
'                    End If
'                    If Len(Trim(oRsResultados.Fields!ValorTexto)) > 0 Then
'                      lcTexto = lcTexto & Trim(IIf(IsNull(oRsResultados.Fields!ValorTexto), "", oRsResultados.Fields!ValorTexto)) & "| "
'                    End If
'                    If Len(Trim(oRsResultados.Fields!ValorCombo)) > 0 Then
'                      lcTexto = lcTexto & Trim(IIf(IsNull(oRsResultados.Fields!ValorCombo), "", oRsResultados.Fields!ValorCombo)) & "| "
'                    End If
'                    If Not IsNull(oRsResultados.Fields!ValorCheck) Then
'                      lcTexto = lcTexto & IIf(oRsResultados.Fields!ValorCheck = True, "x", "")
'                    End If
'                    If lcTexto <> "" Then
'                        lcTexto = Trim(lcTexto)
'                        If Right(lcTexto, 1) = "|" Then
'                            lcTexto = Left(lcTexto, Len(lcTexto) - 1)
'                        End If
'
'                        If lbEsOpenOffice = True Then
'                            Call Feuille.getcellbyposition(3, iFila - 1).setFormula(lcTexto)
'                        Else
'                            oWorkSheet.Cells(iFila, 4).Value = lcTexto
'                        End If
'                    End If
'                    If lbEsOpenOffice = True Then
'                        Call Feuille.getcellbyposition(7, iFila - 1).setFormula(IIf(IsNull(oRsResultados.Fields!ValorReferencial), "", oRsResultados.Fields!ValorReferencial))
'                        Call Feuille.getcellbyposition(8, iFila - 1).setFormula(IIf(IsNull(oRsResultados.Fields!Metodo), "", oRsResultados.Fields!Metodo))
'                    Else
'                        oWorkSheet.Cells(iFila, 8).Value = oRsResultados.Fields!ValorReferencial
'                        oWorkSheet.Cells(iFila, 9).Value = oRsResultados.Fields!Metodo
'                    End If
'                    iFila = iFila + 1
'                    oRsResultados.MoveNext
'                    If oRsResultados.EOF Then
'                       Exit Do
'                    End If
'              Loop
'                If lbEsOpenOffice = True Then
'                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula("*************************************")
'                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula("*************************************")
'                    Call Feuille.getcellbyposition(3, iFila - 1).setFormula("*************************************")
'                    Call Feuille.getcellbyposition(4, iFila - 1).setFormula("*************************************************************************************")
'                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula("*************************************")
'                    Call Feuille.getcellbyposition(6, iFila - 1).setFormula("*************************************")
'                    Call Feuille.getcellbyposition(7, iFila - 1).setFormula("**************************************************************************")
'                    Call Feuille.getcellbyposition(8, iFila - 1).setFormula("*************************************")
'                Else
'
'                    oWorkSheet.Cells(iFila, 2).Value = "*************************************"
'                    oWorkSheet.Cells(iFila, 3).Value = "*************************************"
'                    oWorkSheet.Cells(iFila, 4).Value = "*************************************"
'                    oWorkSheet.Cells(iFila, 5).Value = "*************************************************************************************"
'                    oWorkSheet.Cells(iFila, 6).Value = "*************************************"
'                    oWorkSheet.Cells(iFila, 7).Value = "*************************************"
'                    oWorkSheet.Cells(iFila, 8).Value = "**************************************************************************"
'                    oWorkSheet.Cells(iFila, 9).Value = "*************************************"
'
'                End If
'              iFila = iFila + 2
'           Loop
'        End If
'        If lbEsOpenOffice = True Then
'        Else
'            oWorkSheet.Cells(iFila, 2).Value = "Digitador: " & InicialesDelDigitador
'            oWorkSheet.range(oWorkSheet.Cells(iFila, 3), oWorkSheet.Cells(iFila + 2, 100)).Select
'        End If
'        If lbEsOpenOffice = True Then
'            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
'            PrintArea(0).Sheet = 0
'            PrintArea(0).startcolumn = 1
'            PrintArea(0).StartRow = 0
'            PrintArea(0).EndColumn = 9
'            PrintArea(0).EndRow = iFila
'            Call Feuille.SetPrintAreas(PrintArea())
'            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
'            MsgBox "El reporte se generó con exito: " & lcArchivoExcel, vbInformation, "Reporte"
'        Else
'            If oWorkSheet.PageSetup.PrintArea <> "" Then
'
'               oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
'
'            End If
'            oExcel.Visible = True
'            If lnCantidadCPT = 1 Then
'                oWorkSheet.PageSetup.Zoom = 73
'                oWorkSheet.PageSetup.Orientation = xlLandscape
'            End If
'            oWorkSheet.PrintPreview
'
'    '        oWorkSheet.PrintOut
'    '        oWorkBook.Close SaveChanges:=False
'        End If
'    If lbEsOpenOffice = True Then
'        'Liberar Memoria
'        Set Plage = Nothing
'        Set Feuille = Nothing
'        Set Document = Nothing
'        Set Desktop = Nothing
'        Set ServiceManager = Nothing
'        Set Style = Nothing
'        Set Border = Nothing
'        'encabezado de pagina
'        Set PageStyles = Nothing
'        Set Sheet = Nothing
'        Set StyleFamilies = Nothing
'        Set DefPage = Nothing
'        Set Htext = Nothing
'        Set Hcontent = Nothing
'    Else
'        'Liberar memoria
'        Set oExcel = Nothing
'        Set oWorkBookPlantilla = Nothing
'        Set oWorkBook = Nothing
'        Set oWorkSheet = Nothing
'    End If
'        Set oRsResultados = Nothing
'        Set oRsResultadosCPT = Nothing
'        Set oRsTmp1 = Nothing
'        Set oRsTmp2 = Nothing
'        Set oExcel = Nothing
'        Exit Sub
'ManejadorError:
'    Select Case Err.Number
'    Case 1004
'        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
'        Resume
'    Case Else
'        MsgBox Err.Description
'    End Select
'
'End Sub

Function InicialesDelDigitador() As String
    Dim oConexion As New Connection
    Dim oReglasCaja As New ReglasCaja
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    InicialesDelDigitador = oReglasCaja.SeleccionaDatosCajeroConexion(sighentidades.USUARIO, sghIniciales, oConexion)
    oConexion.Close
    Set oConexion = Nothing
    Set oReglasCaja = Nothing
End Function
