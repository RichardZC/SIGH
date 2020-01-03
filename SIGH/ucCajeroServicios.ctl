VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl ucCajeroServicios 
   ClientHeight    =   9555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12480
   LockControls    =   -1  'True
   ScaleHeight     =   9555
   ScaleWidth      =   12480
   Begin VB.CommandButton btnLeerProductos 
      Caption         =   "Leer ..."
      Height          =   345
      Left            =   10740
      TabIndex        =   35
      Top             =   2220
      Width           =   1185
   End
   Begin VB.Frame frmFiltro 
      Height          =   735
      Left            =   90
      TabIndex        =   33
      Top             =   1980
      Width           =   12255
      Begin VB.ComboBox cmbIdEstado 
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
         ItemData        =   "ucCajeroServicios.ctx":0000
         Left            =   7440
         List            =   "ucCajeroServicios.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   240
         Width           =   3075
      End
      Begin VB.ComboBox cmbIdPuntosDeCarga 
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
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
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
         Height          =   225
         Left            =   6660
         TabIndex        =   36
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Punto de carga"
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
         TabIndex        =   34
         Top             =   240
         Width           =   1290
      End
   End
   Begin VB.CommandButton btnBuscar 
      Caption         =   "Seleccionar Cuenta de Atencion ..."
      Height          =   345
      Left            =   9060
      TabIndex        =   3
      Top             =   660
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.CommandButton btnHistoria 
      Caption         =   "..."
      Height          =   345
      Left            =   3120
      TabIndex        =   15
      Top             =   660
      Width           =   345
   End
   Begin VB.Frame fraOtros 
      Height          =   2235
      Left            =   120
      TabIndex        =   27
      Top             =   7170
      Width           =   8655
      Begin MSFlexGridLib.MSFlexGrid grdSubtotales 
         Height          =   2025
         Left            =   90
         TabIndex        =   28
         Top             =   180
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   3572
         _Version        =   393216
         FixedRows       =   0
         ForeColor       =   0
         BackColorFixed  =   8476221
         ForeColorFixed  =   16777215
         BackColorBkg    =   15198696
         GridColorFixed  =   16777215
         GridLines       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdCaja 
         Height          =   1845
         Left            =   4440
         TabIndex        =   29
         Top             =   210
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   3254
         _Version        =   393216
         BackColorFixed  =   8476221
         ForeColorFixed  =   16777215
         BackColorBkg    =   15198696
         GridColor       =   16777215
         GridColorFixed  =   16777215
         GridLinesFixed  =   1
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame frmAccion 
      Height          =   2205
      Left            =   8820
      TabIndex        =   21
      Top             =   7170
      Width           =   3555
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ucCajeroServicios.ctx":0004
         DownPicture     =   "ucCajeroServicios.ctx":04C8
         Height          =   700
         Left            =   1470
         Picture         =   "ucCajeroServicios.ctx":09B4
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   210
         Width           =   1275
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Limpiar [F5]"
         DisabledPicture =   "ucCajeroServicios.ctx":0EA0
         DownPicture     =   "ucCajeroServicios.ctx":1289
         Height          =   700
         Left            =   1470
         Picture         =   "ucCajeroServicios.ctx":1695
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   990
         Width           =   1275
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir [F3]"
         Height          =   705
         Left            =   120
         Picture         =   "ucCajeroServicios.ctx":1AA1
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   990
         Width           =   1275
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ucCajeroServicios.ctx":1F7A
         DownPicture     =   "ucCajeroServicios.ctx":23DA
         Height          =   700
         Left            =   120
         Picture         =   "ucCajeroServicios.ctx":284F
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   210
         Width           =   1275
      End
   End
   Begin TabDlg.SSTab tabCuentas 
      Height          =   4335
      Left            =   90
      TabIndex        =   10
      Top             =   2790
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   7646
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Servicios"
      TabPicture(0)   =   "ucCajeroServicios.ctx":2CC4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdServicios"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Bienes e Insumos"
      TabPicture(1)   =   "ucCajeroServicios.ctx":2CE0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grillaBusqueda"
      Tab(1).Control(1)=   "grdBienes"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Pagos a Cuenta"
      TabPicture(2)   =   "ucCajeroServicios.ctx":2CFC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdExoneraciones"
      Tab(2).Control(1)=   "grdACuenta"
      Tab(2).ControlCount=   2
      Begin UltraGrid.SSUltraGrid grillaBusqueda 
         Height          =   2415
         Left            =   -74160
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   4260
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
         Caption         =   "grillaBusqueda"
      End
      Begin UltraGrid.SSUltraGrid grdServicios 
         Height          =   3735
         Left            =   150
         TabIndex        =   11
         Top             =   450
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   6588
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
         Caption         =   "Servicios"
      End
      Begin UltraGrid.SSUltraGrid grdBienes 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   12
         Top             =   480
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   6588
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
         Caption         =   "Bienes e Insumos"
      End
      Begin UltraGrid.SSUltraGrid grdExoneraciones 
         Height          =   1935
         Left            =   -74850
         TabIndex        =   14
         Top             =   2250
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   3413
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
         Caption         =   "Exoneraciones"
      End
      Begin UltraGrid.SSUltraGrid grdACuenta 
         Height          =   1695
         Left            =   -74850
         TabIndex        =   13
         Top             =   450
         Width           =   12030
         _ExtentX        =   21220
         _ExtentY        =   2990
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
         Caption         =   "Pagos a Cuenta"
      End
   End
   Begin VB.Frame fraPaciente 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   90
      TabIndex        =   16
      Top             =   30
      Width           =   12270
      Begin VB.TextBox txtIdCuentaAtencion 
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
         Left            =   7440
         TabIndex        =   2
         Top             =   660
         Width           =   1440
      End
      Begin VB.TextBox txtRuc 
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
         Left            =   7440
         TabIndex        =   5
         Top             =   1050
         Width           =   1440
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
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   3705
      End
      Begin VB.TextBox txtNroDocumento 
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
         Left            =   8520
         MaxLength       =   30
         TabIndex        =   8
         Top             =   1440
         Width           =   2385
      End
      Begin VB.ComboBox cmbIdTipoPaciente 
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
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3705
      End
      Begin VB.TextBox txtNroSerie 
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
         Left            =   7440
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1440
         Width           =   825
      End
      Begin VB.TextBox txtNroHistoria 
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
         Left            =   1650
         MaxLength       =   9
         TabIndex        =   1
         Top             =   660
         Width           =   1335
      End
      Begin VB.TextBox txtPaciente 
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
         Left            =   1650
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1050
         Width           =   3705
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre paciente"
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
         Left            =   150
         TabIndex        =   32
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Nº Cuenta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6270
         TabIndex        =   30
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label Label5 
         Caption         =   "R.U.C."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   26
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo documento"
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
         TabIndex        =   25
         Top             =   1440
         Width           =   1380
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8340
         TabIndex        =   20
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "N° de Historia"
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
         Left            =   180
         TabIndex        =   19
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Documento"
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
         Left            =   5940
         TabIndex        =   18
         Top             =   1500
         Width           =   1395
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de paciente"
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
         Left            =   150
         TabIndex        =   17
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.Menu mnuServicio 
      Caption         =   "Menu Servicio"
      Visible         =   0   'False
      Begin VB.Menu mnuAgregarServ 
         Caption         =   "Agregar Servicio"
      End
   End
   Begin VB.Menu mnuBienes 
      Caption         =   "Menu Bienes Insumos"
      Begin VB.Menu mnuAgregarBien 
         Caption         =   "Agregar Bien Insumo"
      End
   End
   Begin VB.Menu mnuACuenta 
      Caption         =   "Menu Pago A Cuenta"
      Begin VB.Menu mnuAgreACuenta 
         Caption         =   "Agregar Pago a Cuenta"
      End
   End
End
Attribute VB_Name = "ucCajeroServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_cmbIdTipoPaciente As New SIGHComun.ListaDespleglable
Dim mo_cmbIdTipoComprobante As New SIGHComun.ListaDespleglable

'efgl 060606
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_cmbIdPuntosDeCarga As New SIGHComun.ListaDespleglable
Dim mo_cmbIdEstado As New SIGHComun.ListaDespleglable
'fin efgl 666

Dim gridInfra As New GridInfragistic
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim NroHistoriaClinica As Long
Dim mo_CuentaAtencion As New DOCuentaAtencion
Dim oComprobantePago As New DOCajaComprobantesPago
Dim oCajaNroDocumento As New DOCajaNroDocumento
Dim ml_IdTipoComprobante As Integer

Dim mb_AgregoAtencion As Boolean
Dim oAtencion As DOAtencion
Dim ml_IdPaciente As Long
Dim ml_IdUsuario As Long

Dim mb_TransaccionDeNuevoRegistroEnProceso  As Boolean
Dim mb_PresionoEscape As Boolean

Const ID_TIPO_COMPROBANTE_FACTURA = 2

Dim ml_IdCaja As Integer
Dim ml_IdCajero As Long
Dim ml_IdTurno As Integer
Dim oCajaGestion As New DOCajaGestion
Dim mrs_Servicios As New Recordset

Public Event HizoClickEnEscape()

'EFGL 14/06/2006
'variables para borrar items de la base de datos
'por lo pronto no se deberian borrar
'Dim mo_ReglasFacturacionServiciosBorrar As Collection
'Dim mo_ReglasFacturacionBienesBorrar  As Collection
'Dim idProductoSelecto() As Long
'Dim nombreProductoSelecto() As String
'Dim numeroProductosSelectos As Integer
'EFGL 14/06/2006

'variables por comprobante de pago
Dim md_Subtotal As Double
Dim md_IGV As Double
Dim md_Exoneraciones As Double
Dim md_PagosACuenta As Double
Dim md_Total As Double
Dim ml_IdComprobantePago As Long

Dim md_Recibido As Double
Dim md_Falta As Double
Dim md_Vuelto As Double

Dim mi_Opcion As sghOpciones

Dim mo_ReglasFacturacionServicios As Collection
Dim mo_ReglasFacturacionBienes  As Collection
Dim mo_ReglasFacturacionACuenta  As Collection

Dim mrs_FacturacionServicios As New ADODB.Recordset
Dim mrs_FacturacionBienes As New ADODB.Recordset
Dim mrs_ACuenta As New ADODB.Recordset

Dim numeroFilaActiva As Integer
Dim ms_TipoProducto As String

Dim mb_NoEditar As Boolean

Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let IdComprobantePago(lValue As Long)
   ml_IdComprobantePago = lValue
End Property
Property Get IdComprobantePago() As Long
   IdComprobantePago = ml_IdComprobantePago
End Property

Property Get IdTipoComprobante() As Integer
    IdTipoComprobante = ml_IdTipoComprobante
End Property

Property Let IdTipoComprobante(oValue As Integer)
    ml_IdTipoComprobante = oValue
End Property

Property Get IdCaja() As Integer
    IdCaja = ml_IdCaja
End Property

Property Let IdCaja(oValue As Integer)
    ml_IdCaja = oValue
End Property

Property Get IdCajero() As Integer
    IdCajero = ml_IdCajero
End Property

Property Let IdCajero(oValue As Integer)
    ml_IdCajero = oValue
End Property

Property Get IdTurno() As Integer
    IdTurno = ml_IdTurno
End Property

Property Let IdTurno(oValue As Integer)
    ml_IdTurno = oValue
End Property

Property Get IdPaciente() As Long
    IdPaciente = ml_IdPaciente
End Property

Property Let IdPaciente(oValue As Long)
    ml_IdPaciente = oValue
End Property

Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property

'----------------------------------------generales---------------------------
Sub ObtenerNombrePaciente(IdPaciente As Long)
    Dim oPaciente As doPaciente
    Set oPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(IdPaciente)
    txtPaciente.Text = oPaciente.ApellidoPaterno & " " & oPaciente.ApellidoMaterno & ", " & oPaciente.PrimerNombre & " " & oPaciente.SegundoNombre
    txtNroHistoria.Text = oPaciente.NroHistoriaClinica
    NroHistoriaClinica = Val(oPaciente.NroHistoriaClinica)
    ml_IdPaciente = IdPaciente
End Sub

Function AgregarNuevaAtencion() As Boolean
       
    AgregarNuevaAtencion = True
    If (Not mb_AgregoAtencion) Then
        Set oAtencion = mo_ReglasFacturacion.SeleccionarUltimaAtencion(IdPaciente, mo_CuentaAtencion.IdCuentaAtencion)
        oAtencion.IdCuentaAtencion = mo_CuentaAtencion.IdCuentaAtencion
'        If (mo_ReglasFacturacion.AgregarAtencion(oAtencion)) Then
'            mb_AgregoAtencion = True
'        Else
'            MsgBox "No se agrego una nueva atencion", vbExclamation, "Consulte al administrador"
'            AgregarNuevaAtencion = False
'        End If
    End If
End Function

Public Function AbrirCaja(oCajero As DOCajaCajero)
On Error GoTo errdes

Set oCajaGestion = New DOCajaGestion
With oCajaGestion
    .FechaApertura = Now
    .EstadoLote = "A"
    '.FechaCierre = ""
    .IdCaja = oCajero.IdCaja
    .IdCajero = oCajero.IdCajero
    .IdGestionCaja = 0
    .IdTurno = IdTurno
    .IdUsuarioAuditoria = IdUsuario
    .TotalCobrado = ""
End With


mo_AdminCaja.CajaGestionAgregar oCajaGestion
If mo_AdminCaja.MensajeError <> "" Then
    MsgBox mo_AdminCaja.MensajeError, vbExclamation, "Error"
End If
Exit Function
errdes:
MsgBox Err.Description, vbCritical, Err.Number
End Function

Public Function ModificarCaja(oExternoCajaGestion As DOCajaGestion)
    Set oCajaGestion = oExternoCajaGestion
End Function

Public Function CerrarCaja()
On Error GoTo errdes

With oCajaGestion
    '.FechaApertura = Now
    .EstadoLote = "C"
    .FechaCierre = Now
    '.IdCaja = oCajero.IdCaja
    '.IdCajero = oCajero.IdCajero
    '.IdGestionCaja = 0
    '.IdTurno = IdTurno
    '.IdUsuarioAuditoria = IdUsuario
    '.TotalCobrado = ""
End With


mo_AdminCaja.CajaGestionModificar oCajaGestion
If mo_AdminCaja.MensajeError <> "" Then
    MsgBox mo_AdminCaja.MensajeError, vbExclamation, "Error"
End If
Exit Function
errdes:
MsgBox Err.Description, vbCritical, Err.Number
End Function
Public Sub BusquedaCuentasAtencion()
    
    Dim rsRespuesta As New Recordset
    Dim frm As New SelecccionProductos
        
        Set oComprobantePago = Nothing
        If (NroHistoriaClinica > 0) Then
            Set rsRespuesta = mo_AdminCaja.ObtenerUltimaCuentaAtencionPorIdPaciente(IdPaciente)
            If rsRespuesta.RecordCount <= 0 Then
                Exit Sub
            End If
            
            If (rsRespuesta.RecordCount = 1) Then
                Set mo_CuentaAtencion = mo_ReglasFacturacion.CuentasAtencionSeleccionarPorId(rsRespuesta!IdCuentaAtencion)
                cmbIdTipoPaciente.Enabled = False
                txtNroHistoria.Enabled = False
                UserControl.btnBuscar.Visible = False
            Else
                UserControl.btnBuscar.Visible = True
            End If
            
'            If (Not oCuentaAtencion Is Nothing) Then
'                'cargar BIENES
'                Set rsRespuesta = mo_AdminCaja.CatalogoBienesInsumosPorCuentaAtencion(oCuentaAtencion.IdCuentaAtencion, Val(mo_cmbIdTipoPaciente.BoundText), Val(mo_cmbIdEspecialidad.BoundText), Trim(cmbIdEstado.Text))
'                frm.BienesDataSource = rsRespuesta
'                'Set grdBienes.DataSource = rsRespuesta
'
'                'cargar servicios
'                Set rsRespuesta = mo_AdminCaja.CatalogoServicioPorCuentaAtencion(oCuentaAtencion.IdCuentaAtencion, Val(mo_cmbIdTipoPaciente.BoundText), Val(mo_cmbIdEspecialidad.BoundText), Trim(cmbIdEstado.Text))
'                frm.ServiciosDataSource = rsRespuesta
'                'Set grdServicios.DataSource = rsRespuesta
'
'                'cargar exoneraciones
'                If Trim(cmbIdEstado.Text) = "3" Then
'                    Set rsRespuesta = mo_AdminCaja.ExoneracionesPorCuentaAtencion(oCuentaAtencion.IdCuentaAtencion)
'                Else
'                    Set rsRespuesta = mo_AdminCaja.ExoneracionesPorCuentaAtencion(0)
'                End If
'                Set grdExoneraciones.DataSource = rsRespuesta
'
'                'cargar pagos a cuenta
'                If Trim(cmbIdEstado.Text) = "3" Then
'                    Set rsRespuesta = mo_AdminCaja.PagosACuentaPorCuentaAtencion(oCuentaAtencion.IdCuentaAtencion)
'                Else
'                    Set rsRespuesta = mo_AdminCaja.PagosACuentaPorCuentaAtencion(0)
'                End If
'                Set grdACuenta.DataSource = rsRespuesta
'
'                frm.Inicializar
'                frm.Show 1
'                If frm.Acepta Then
'                    Set grdBienes.DataSource = frm.BienesDataSource
'                    Set grdServicios.DataSource = frm.ServiciosDataSource
'                End If
'                ActualizaTotales
'            End If
        Else
            MsgBox "Debe seleccionar un Paciente", vbCritical, "Cuentas de Atencion"
            Exit Sub
        End If
        
        txtIdCuentaAtencion.Text = mo_CuentaAtencion.IdCuentaAtencion
        
            
'        If mo_AdminAdmision.MensajeError <> "" Then
'            MsgBox mo_AdminAdmision.MensajeError, vbCritical, "Filtro Pacientes"
'        End If
        
End Sub
Sub ActualizaTotales()
    Dim i As Integer
    Dim rsAux As ADODB.Recordset
    On Error Resume Next
    md_Subtotal = 0
    md_Exoneraciones = 0
    md_PagosACuenta = 0
    md_IGV = 0
    md_Total = 0
    
    Set rsAux = grdServicios.DataSource
    If Not rsAux Is Nothing Then
        If Not (rsAux.BOF And rsAux.EOF) Then
            rsAux.MoveFirst
            For i = 0 To rsAux.RecordCount - 1
                md_Subtotal = md_Subtotal + Val(rsAux!TotalPorPagar)
                rsAux.MoveNext
            Next
        End If
    End If
    
    
    Set rsAux = grdBienes.DataSource
    If Not rsAux Is Nothing Then
        If Not (rsAux.BOF And rsAux.EOF) Then
            rsAux.MoveFirst
            For i = 0 To rsAux.RecordCount - 1
                md_Subtotal = md_Subtotal + Val(rsAux!TotalPorPagar)
                rsAux.MoveNext
            Next
        End If
    End If
    
    
    Set rsAux = grdExoneraciones.DataSource
    If Not rsAux Is Nothing Then
        If Not (rsAux.BOF And rsAux.EOF) Then
            rsAux.MoveFirst
            For i = 0 To rsAux.RecordCount - 1
                md_Exoneraciones = md_Exoneraciones + Val(rsAux!TotalExonerado)
                rsAux.MoveNext
            Next
        End If
    End If
    
    Set rsAux = grdACuenta.DataSource
    If Not rsAux Is Nothing Then
        If Not (rsAux.BOF And rsAux.EOF) Then
            rsAux.MoveFirst
            For i = 0 To rsAux.RecordCount - 1
                md_PagosACuenta = md_PagosACuenta + Val(rsAux!TotalPagado)
                rsAux.MoveNext
            Next
        End If
    End If
    
    
    grdSubtotales.TextMatrix(0, 1) = Format(md_Subtotal, "#0.000")
    'md_IGV = 0.19 * md_Subtotal
    md_IGV = 0
    grdSubtotales.TextMatrix(1, 1) = Format(md_IGV, "#0.000")
    grdSubtotales.TextMatrix(2, 1) = Format(md_Exoneraciones, "#0.000")
    grdSubtotales.TextMatrix(3, 1) = Format(md_PagosACuenta, "#0.000")
    
    If cmbIdEstado.Text = "3" Then
        md_Total = md_Subtotal + md_IGV - md_Exoneraciones - md_PagosACuenta
    Else
        md_Total = md_Subtotal + md_IGV - md_Exoneraciones + md_PagosACuenta
    End If
    grdSubtotales.TextMatrix(4, 1) = Format(md_Total, "#0.000")
    
End Sub
Public Sub BusquedaPaciente()
Dim oPaciente As New doPaciente
Dim rsRespuesta As New Recordset
        
        Dim oFrm As New PacientesBusqueda
        oFrm.TipoFiltro = sghFiltrarTodos
        oFrm.Caption = "Seleccione el paciente"
        oFrm.Show vbModal
        If oFrm.IdRegistroSeleccionado <> 0 Then
            IdPaciente = oFrm.IdRegistroSeleccionado
            ObtenerNombrePaciente oFrm.IdRegistroSeleccionado
            BusquedaCuentasAtencion
        End If
            
End Sub

Private Sub btnHistoria_Click()
    Me.BusquedaPaciente
End Sub

Public Function Inicializar()

    Set mo_cmbIdTipoPaciente.MiComboBox = cmbIdTipoPaciente
    Set mo_cmbIdTipoComprobante.MiComboBox = cmbIdTipoComprobante
    Set mo_cmbIdPuntosDeCarga.MiComboBox = cmbIdPuntosDeCarga
    Set mo_cmbIdEstado.MiComboBox = cmbIdEstado
    
    ConfigurarTipoPaciente
    ConfigurarTipoComprobante
    ConfigurarPuntosDeCarga
    ConfigurarEstado
    ConfigurarSubtotales
    ConfigurarCaja
    
    txtNroHistoria.Enabled = False
    txtPaciente.Enabled = False
    btnHistoria.Enabled = False
    txtNroSerie.Enabled = False
    txtNroDocumento.Enabled = False
    grdBienes.Override.AllowAddNew = ssAllowAddNewYes
    grdServicios.Override.AllowAddNew = ssAllowAddNewYes
    
    cmbIdTipoPaciente.ListIndex = 0
    cmbIdPuntosDeCarga.ListIndex = 0
    cmbIdEstado.ListIndex = 0
    
    On Error Resume Next
    txtNroHistoria.SetFocus
    
    mb_AgregoAtencion = False
    
End Function

Sub ConfigurarSubtotales()
    
    grdSubtotales.Clear
    grdSubtotales.Cols = 2
    grdSubtotales.Rows = 5
    grdSubtotales.ColWidth(0) = 2000
    grdSubtotales.ColWidth(1) = 2000
    
    grdSubtotales.Col = 0
    
    grdSubtotales.Row = 0
    grdSubtotales.Text = "Subtotal"
    
    grdSubtotales.Row = 1
    grdSubtotales.Text = "IGV"
    
    grdSubtotales.Row = 2
    grdSubtotales.Text = "Exoneraciones"
    
    grdSubtotales.Row = 3
    grdSubtotales.Text = "Pagos A Cuenta"
    
    grdSubtotales.Row = 4
    grdSubtotales.Text = "Total"
End Sub
Sub ConfigurarCaja()

    grdCaja.Clear
    grdCaja.Cols = 2
    grdCaja.Rows = 4
    grdCaja.MergeCells = flexMergeRestrictRows
    grdCaja.ColWidth(0) = 2000
    grdCaja.ColWidth(1) = 2000
    
    grdCaja.Row = 0
    grdCaja.Col = 0
    grdCaja.Text = "Caja"
    grdCaja.Col = 1
    grdCaja.Text = "Caja"
    grdCaja.MergeRow(0) = True
    grdCaja.Col = 0
    
    grdCaja.Row = 1
    grdCaja.Text = "Recibido"
    
    grdCaja.Row = 2
    grdCaja.Text = "Falta"
    
    grdCaja.Row = 3
    grdCaja.Text = "Vuelto"
    
End Sub

Sub ConfigurarTipoPaciente()
    mo_cmbIdTipoPaciente.ListField = "DescripcionLarga"
    mo_cmbIdTipoPaciente.BoundColumn = "IdTipoFinanciamiento"
    Set mo_cmbIdTipoPaciente.RowSource = mo_ReglasFacturacion.TiposFinanciamientoSeleccionarParaCaja()
End Sub
Sub ConfigurarTipoComprobante()
    mo_cmbIdTipoComprobante.BoundColumn = "IdTipoComprobante"
    mo_cmbIdTipoComprobante.ListField = "Descripcion"
    Set mo_cmbIdTipoComprobante.RowSource = mo_AdminCaja.TiposComprobanteSeleccionarTodos()
End Sub
Sub ConfigurarPuntosDeCarga()
    
    mo_cmbIdPuntosDeCarga.ListField = "Descripcion"
    mo_cmbIdPuntosDeCarga.BoundColumn = "IdPuntoCarga"
    
    cmbIdPuntosDeCarga.AddItem "<Todos>"
    Set mo_cmbIdPuntosDeCarga.RowSourceSinClear = mo_ReglasComunes.SeleccionarPuntosDeCarga()

End Sub
Sub ConfigurarEstado()
        
    mo_cmbIdEstado.ListField = "Descripcion"
    mo_cmbIdEstado.BoundColumn = "IdEstadoFacturacion"
    
    cmbIdEstado.AddItem "<Todos>"
    Set mo_cmbIdEstado.RowSourceSinClear = mo_ReglasFacturacion.EstadosFacturacionObtenerTodosExceptoPagado()
    
   
End Sub

Private Sub btnLeerProductos_Click()
Dim rsRespuesta As New Recordset
    
    
    If (Not mo_CuentaAtencion Is Nothing) Then
        'cargar BIENES
        Set rsRespuesta = mo_AdminCaja.CatalogoBienesInsumosPorCuentaAtencion(mo_CuentaAtencion.IdCuentaAtencion, Val(mo_cmbIdTipoPaciente.BoundText), Val(mo_cmbIdPuntosDeCarga.BoundText), Val(mo_cmbIdEstado.BoundText))
        Set grdBienes.DataSource = rsRespuesta

        'cargar servicios
        Set rsRespuesta = mo_AdminCaja.CatalogoServicioPorCuentaAtencion(mo_CuentaAtencion.IdCuentaAtencion, Val(mo_cmbIdTipoPaciente.BoundText), Val(mo_cmbIdPuntosDeCarga.BoundText), Val(mo_cmbIdEstado.BoundText))
        Set grdServicios.DataSource = rsRespuesta

        'cargar exoneraciones
        If Trim(cmbIdEstado.Text) = "3" Then    'En el caso de el pago de cuenta final carga las exoneraciones
            'Set rsRespuesta = mo_AdminCaja.ExoneracionesPorCuentaAtencion(mo_CuentaAtencion.IdCuentaAtencion)
        Else
            'Solo se usa para configurar la grilla
            'Set rsRespuesta = mo_AdminCaja.ExoneracionesPorCuentaAtencion(0)
        End If
        Set grdExoneraciones.DataSource = rsRespuesta

        'cargar pagos a cuenta
        If Trim(cmbIdEstado.Text) = "3" Then    'En el caso de el pago de cuenta final carga los pagos a cuenta
            'Set rsRespuesta = mo_AdminCaja.PagosACuentaPorCuentaAtencion(mo_CuentaAtencion.IdCuentaAtencion)
        Else
            'Set rsRespuesta = mo_AdminCaja.PagosACuentaPorCuentaAtencion(0)
        End If
        Set grdACuenta.DataSource = rsRespuesta

        ActualizaTotales
    End If


End Sub


Private Sub cmbIdTipoComprobante_Change()
    txtNroSerie.Text = ""
    txtNroDocumento.Text = ""
    txtRuc.Enabled = True
    IdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
    If IdTipoComprobante > 0 Then
        Set oCajaNroDocumento = mo_AdminCaja.NroDocumentoSeleccionarPorIdCajaYTipoComprobante(IdCaja, IdTipoComprobante)
        txtNroSerie.Text = Trim(oCajaNroDocumento.NroSerie)
        txtNroDocumento.Text = Trim(oCajaNroDocumento.NroDocumento)
        If IdTipoComprobante = ID_TIPO_COMPROBANTE_FACTURA Then
            'txtRuc.Text = ""
            txtRuc.Enabled = False
        End If
    End If
End Sub

Private Sub cmbIdTipoComprobante_Click()
    txtRuc.Enabled = True
    txtNroSerie.Text = ""
    txtNroDocumento.Text = ""
    IdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
    If IdTipoComprobante > 0 Then
        Set oCajaNroDocumento = mo_AdminCaja.NroDocumentoSeleccionarPorIdCajaYTipoComprobante(IdCaja, IdTipoComprobante)
        txtNroSerie.Text = Trim(oCajaNroDocumento.NroSerie)
        txtNroDocumento.Text = Trim(oCajaNroDocumento.NroDocumento)
        If IdTipoComprobante <> ID_TIPO_COMPROBANTE_FACTURA Then
            'txtRuc.Text = ""
            txtRuc.Enabled = False
        End If
    End If
End Sub

Private Sub cmbIdTipoPaciente_Click()
        txtNroHistoria.Enabled = False
        btnHistoria.Enabled = False
        txtPaciente.Enabled = False
        txtNroHistoria.Text = ""
        txtPaciente.Text = ""
        
    If (mo_cmbIdTipoPaciente.BoundText = "1") Then
        txtNroHistoria.Enabled = True
        btnHistoria.Enabled = True
    ElseIf (mo_cmbIdTipoPaciente.BoundText = "5") Then
        txtPaciente.Enabled = True
    End If
End Sub

Private Sub FormatoGrilla(oGrilla As SSUltraGrid)
Dim oColumnProducto As SSColumn
   
    oGrilla.Bands(0).Columns("IdProducto").Hidden = True
    oGrilla.Bands(0).Columns("id").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Código"
    oGrilla.Bands(0).Columns("Codigo").Width = 800
    
    Set oColumnProducto = oGrilla.Bands(0).Columns("Descripcion")
    oColumnProducto.Header.Caption = "Descripción"
    oColumnProducto.Width = 7800
    
    oGrilla.Bands(0).Columns("Cantidad").Header.Caption = "Cantidad"
    oGrilla.Bands(0).Columns("Cantidad").Width = 800
    
    oGrilla.Bands(0).Columns("preciounitario").Header.Caption = "P.U.(S/.)"
    oGrilla.Bands(0).Columns("preciounitario").Format = "#0.00"
    oGrilla.Bands(0).Columns("preciounitario").Width = 1000
    
    oGrilla.Bands(0).Columns("totalporpagar").Header.Caption = "Subtotal"
    oGrilla.Bands(0).Columns("totalporpagar").Width = 1000
    oGrilla.Bands(0).Columns("totalporpagar").Format = "#0.00"
    
    oGrilla.Bands(0).Columns("Idestadofacturacion").Hidden = True
    oGrilla.Bands(0).Columns("IdAtencion").Hidden = True
    
    
    gridInfra.ConfigurarFilasBiColores oGrilla, SIGHComun.GrillaConFilasBicolor
End Sub

Sub BuscaProductos(sNombre As String, lIdTipoFinanciamiento As Long, lIdPuntoCarga As Long)
    
    Dim rs As New Recordset
    
    If ms_TipoProducto = "servicios" Then
        grillaBusqueda.Left = grdServicios.Left
        Set rs = mo_AdminCaja.ServiciosFiltrarParaCajero(sNombre, lIdTipoFinanciamiento, lIdPuntoCarga)
        grillaBusqueda.Top = grdServicios.ActiveCell.GetUIElement.RECT.Bottom * Screen.TwipsPerPixelY + 500
    Else
        grillaBusqueda.Left = grdBienes.Left
        Set rs = mo_AdminCaja.BienesFiltrarParaCajero(sNombre, lIdTipoFinanciamiento, lIdPuntoCarga)
        grillaBusqueda.Top = grdBienes.ActiveCell.GetUIElement.RECT.Bottom * Screen.TwipsPerPixelY + 500
    End If
    
    Set grillaBusqueda.DataSource = rs
    'grillaBusqueda.Refresh ssRefreshDisplay
    grillaBusqueda.Visible = True
    grillaBusqueda.Enabled = True
    'grillaBusqueda.Left = 300
    
End Sub
Private Sub FormatoGrillaBusqueda(oGrilla As SSUltraGrid)
   
    oGrilla.Bands(0).Columns("IdProducto").Hidden = True
    oGrilla.Bands(0).Columns("Activo").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Código"
    oGrilla.Bands(0).Columns("Codigo").Width = 800
    
    
    oGrilla.Bands(0).Columns("Nombre").Header.Caption = "Descripción"
    oGrilla.Bands(0).Columns("Nombre").Width = 7800
    
    
    oGrilla.Bands(0).Columns("preciounitario").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
    
    gridInfra.ConfigurarFilasBiColores oGrilla, SIGHComun.GrillaConFilasBicolor
End Sub


Private Sub cmdGrabar_Click()
    
    Select Case mi_Opcion
    Case sghAgregar
       If ValidarDatosObligatorios() Then
           CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If AgregarDatos() Then
                    MsgBox "Los datos se agregaron correctamente", vbInformation, "Comprobante de Pago"
                    NuevoComprobante
                Else
                    MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminCaja.MensajeError, vbExclamation, "Comprobante de Pago"
               End If
           End If
       End If
    Case sghModificar
    Case sghEliminar
    End Select

End Sub

Private Sub cmdNuevo_Click()
    NuevoComprobante
End Sub

Private Sub Command1_Click()
    FormatoGrilla grdServicios
End Sub

Private Sub cmdSalir_Click()
    RaiseEvent HizoClickEnEscape
End Sub

Private Sub grdACuenta_AfterRowUpdate(ByVal Row As UltraGrid.SSRow)
    ActualizaTotales
End Sub

Private Sub grdACuenta_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)

    grdACuenta.Bands(0).Columns("id").Hidden = True
    grdACuenta.Bands(0).Columns("idAtencion").Hidden = True
    grdACuenta.Bands(0).Columns("fechaPago").Header.Caption = "Fecha"
    grdACuenta.Bands(0).Columns("fechaPago").Width = 3000
    grdACuenta.Bands(0).Columns("NombresEmpleado").Header.Caption = "Cajero"
    grdACuenta.Bands(0).Columns("NombresEmpleado").Width = 3000
    grdACuenta.Bands(0).Columns("Empleado").Hidden = True
    grdACuenta.Bands(0).Columns("IdEmpleadoCajero").Hidden = True
    grdACuenta.Bands(0).Columns("IdComprobantePago").Hidden = True
    grdACuenta.Bands(0).Columns("totalPagado").Header.Caption = "SubTotal"
    grdACuenta.Bands(0).Columns("totalPagado").Format = "#0.00"
    gridInfra.ConfigurarFilasBiColores grdACuenta, SIGHComun.GrillaConFilasBicolor
    
'
'
'
'
'    grdACuenta.Bands(0).Columns("IdAtencion").Hidden = True
'    grdACuenta.Bands(0).Columns("Id").Hidden = True
'    grdACuenta.Bands(0).Columns("IdEmpleadoCajero").Hidden = True
'    grdACuenta.Bands(0).Columns("Empleado").Hidden = True
'    grdACuenta.Bands(0).Columns("NombresEmpleado").Hidden = True
'    grdACuenta.Bands(0).Columns("FechaPago").Header.Caption = "Fecha"
'    grdACuenta.Bands(0).Columns("FechaPago").Width = 2500
'    grdACuenta.Bands(0).Columns("IdComprobantePago").Hidden = True
'    grdACuenta.Bands(0).Columns("TotalPagado").Header.Caption = "Subtotal"
'    grdACuenta.Bands(0).Columns("TotalPagado").Format = "#0.00"
'
'    gridInfra.ConfigurarFilasBiColores grdACuenta, SIGHComun.GrillaConFilasBicolor
End Sub

Private Sub grdACuenta_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Or KeyAscii = 9 Then
        ActualizaTotales
    End If
End Sub

Private Sub grdACuenta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuACuenta
    End If
End Sub







Private Sub grdCaja_KeyPress(KeyAscii As Integer)
    If grdCaja.Col = 1 And grdCaja.Row = 1 Then
        If Not ((KeyAscii > 47 And KeyAscii < 59) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 42)) Then
            KeyAscii = 0
            Exit Sub
        End If
        If KeyAscii <> 8 Then
            grdCaja.Text = grdCaja.Text & Chr(KeyAscii)
        Else
            If (grdCaja.Text <> "") Then
                grdCaja.Text = Mid(grdCaja.Text, 1, Len(grdCaja.Text) - 1)
            End If
        End If
        
        grdCaja.TextMatrix(2, 1) = ""
        grdCaja.TextMatrix(3, 1) = ""
        md_Recibido = Val(grdCaja.Text)
        If (md_Recibido < md_Total) Then
            md_Falta = md_Total - md_Recibido
            grdCaja.TextMatrix(2, 1) = Format(md_Falta, "#0.000")
        Else
            md_Vuelto = md_Recibido - md_Total
            grdCaja.TextMatrix(3, 1) = Format(md_Vuelto, "#0.000")
        End If
        
    Else
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub grdExoneraciones_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdExoneraciones.Bands(0).Columns("IdAtencion").Hidden = True
    grdExoneraciones.Bands(0).Columns("Id").Hidden = True
    grdExoneraciones.Bands(0).Columns("Empleado").Hidden = True
    grdExoneraciones.Bands(0).Columns("NombresEmpleado").Hidden = True
    grdExoneraciones.Bands(0).Columns("FechaExoneracion").Header.Caption = "Fecha"
    grdExoneraciones.Bands(0).Columns("IdEmpleadoExonera").Hidden = True
    grdExoneraciones.Bands(0).Columns("TotalExonerado").Header.Caption = "Subtotal"
    
    gridInfra.ConfigurarFilasBiColores grdExoneraciones, SIGHComun.GrillaConFilasBicolor
End Sub








Private Sub grdSubtotales_KeyPress(KeyAscii As Integer)
    If grdSubtotales.Col = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    
        KeyAscii = 0
    
    
End Sub

Private Sub grillaBusqueda_DblClick()
    Dim fila As New Record
'    If Not grillaBusqueda.ActiveCell Is Nothing Then
'       Set fila.Source = grillaBusqueda.ActiveCell.Row
'       Exit Sub
'    End If
    If Not grillaBusqueda.ActiveRow Is Nothing Then
        If ms_TipoProducto = "bienes" Then
           grdBienes.ActiveRow.Cells("codigo").Value = grillaBusqueda.ActiveRow.Cells("CODIGO").Value
           grdBienes.ActiveRow.Cells("idproducto").Value = grillaBusqueda.ActiveRow.Cells("idproducto").Value
           grdBienes.ActiveRow.Cells("descripcion").Value = grillaBusqueda.ActiveRow.Cells("nombre").Value
           grdBienes.ActiveRow.Cells("preciounitario").Value = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
           grdBienes.ActiveRow.Cells("totalporpagar").Value = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
           grdBienes.ActiveRow.Cells("idestadofacturacion").Value = 4
           grdBienes.ActiveRow.Cells("cantidad").Value = 1
        Else
           grdServicios.ActiveRow.Cells("codigo").Value = grillaBusqueda.ActiveRow.Cells("CODIGO").Value
           grdServicios.ActiveRow.Cells("idproducto").Value = grillaBusqueda.ActiveRow.Cells("idproducto").Value
           grdServicios.ActiveRow.Cells("descripcion").Value = grillaBusqueda.ActiveRow.Cells("nombre").Value
           grdServicios.ActiveRow.Cells("preciounitario").Value = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
           grdServicios.ActiveRow.Cells("totaLPORPagar").Value = grillaBusqueda.ActiveRow.Cells("preciounitario").Value
           grdServicios.ActiveRow.Cells("idestadofacturacion").Value = 4
           grdServicios.ActiveRow.Cells("cantidad").Value = 1
        End If
        
        Set grillaBusqueda.DataSource = Nothing
        grillaBusqueda.Visible = False
    
        Exit Sub
    End If
End Sub

Private Sub grillaBusqueda_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    FormatoGrillaBusqueda grillaBusqueda
    gridInfra.ConfigurarFilasBiColores grillaBusqueda, SIGHComun.GrillaConFilasBicolor
End Sub

Private Sub mnuAgreACuenta_Click()
    NuevoPagoACuenta
End Sub

Private Sub txtNroHistoria_LostFocus()
Dim oPaciente As New doPaciente
Dim rsRespuesta As New ADODB.Recordset

        IdPaciente = 0
        txtPaciente.Text = ""
        txtPaciente.Tag = ""
        If (txtNroHistoria <> "") Then
            oPaciente.NroHistoriaClinica = Val(UserControl.txtNroHistoria)
                
            Set rsRespuesta = mo_AdminAdmision.PacientesFiltrar(oPaciente)
            On Error Resume Next
            If rsRespuesta.RecordCount = 0 Then
                MsgBox "No se encontraron datos", vbInformation, "Búsqueda"
            ElseIf rsRespuesta.RecordCount = 1 Then
                IdPaciente = rsRespuesta!IdPaciente
                ObtenerNombrePaciente rsRespuesta!IdPaciente
                BusquedaCuentasAtencion
            End If
            
            If mo_AdminAdmision.MensajeError <> "" Then
                MsgBox mo_AdminAdmision.MensajeError, vbCritical, "Filtro Pacientes"
            End If
        End If
        
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
   
'   fraPaciente.Width = UserControl.Width - fraEstadoCuenta.Width - 360
'   tabCuentas.Width = fraPaciente.Width
'
'   fraPaciente.Left = 80
'   tabCuentas.Left = 80
'   fraEstadoCuenta.Left = UserControl.Width - fraEstadoCuenta.Width - 160
'   fraOtros.Left = fraEstadoCuenta.Left
'
'   frmAccion.Top = UserControl.Height - frmAccion.Height - 80
'   frmAccion.Left = 80
'   frmAccion.Width = UserControl.Width - 360
'
'   tabCuentas.Height = UserControl.Height - fraPaciente.Height - frmAccion.Height - 480
'   fraOtros.Height = tabCuentas.Height
'
'   grdServicios.Width = tabCuentas.Width - 240
'   grdBienes.Width = tabCuentas.Width - 240
'
'   grdServicios.Height = tabCuentas.Height - 480
'   grdBienes.Height = tabCuentas.Height - 480
'
'   grdExoneraciones.Height = tabCuentas.Height / 4
'   grdExoneraciones.Top = 0
'
'   grdACuenta.Height = tabCuentas.Height / 4
'   grdACuenta.Top = grdExoneraciones.Height + 40
'
'   grdSubtotales.Height = tabCuentas.Height / 5
'   grdSubtotales.Top = grdACuenta.Top + grdACuenta.Height + 60
'
'   grdCaja.Height = tabCuentas.Height / 5
'   grdCaja.Top = grdSubtotales.Top + grdSubtotales.Height + 40
   
End Sub

Function ValidarDatosObligatorios() As Boolean
Dim sMensaje As String
Dim bFound As Boolean

    ValidarDatosObligatorios = False
    
    If txtNroSerie.Text = "" Then
        sMensaje = sMensaje + "Ingrese el Nº de Serie" + Chr(13)
    End If
    If txtNroDocumento.Text = "" Then
        sMensaje = sMensaje + "Ingrese el Nº de Documento" + Chr(13)
    End If
    If mo_cmbIdTipoComprobante.BoundText = "" Then
        sMensaje = sMensaje + "Ingrese el tipo de Comprobante" + Chr(13)
    End If
    If Val(mo_cmbIdTipoComprobante.BoundText) = ID_TIPO_COMPROBANTE_FACTURA Then
        If txtRuc = "" Then
            sMensaje = sMensaje + "Ingrese el RUC para la Factura" + Chr(13)
        End If
    Else
         txtRuc = ""
    End If
    If Trim(txtPaciente.Text) = "" Then
        sMensaje = sMensaje + "Ingrese la Razón Social" + Chr(13)
    End If
    
    If sMensaje <> "" Then
         MsgBox sMensaje, vbExclamation, "Comprobantes de Pago"
         Exit Function
    End If
    
    ValidarDatosObligatorios = True
End Function

Sub CargaDatosAlObjetosDeDatos()
On Error GoTo errDescription
    
    Set oComprobantePago = New DOCajaComprobantesPago
    With oComprobantePago
        .IdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
        .NroSerie = Trim(txtNroSerie.Text)
        .NroDocumento = Trim(txtNroDocumento.Text)
        .IdCuentaAtencion = txtIdCuentaAtencion
        .RazonSocial = txtPaciente.Text
        .Observaciones = ""
        .IdGestionCaja = oCajaGestion.IdGestionCaja
        .IdUsuarioAuditoria = ml_IdUsuario
        .RUC = txtRuc.Text
        
        .subtotal = md_Subtotal
        .IGV = md_IGV
        .Total = md_Total
        
        '.FechaCobranza
        .IdComprobantePago = 0
        
    End With
    
    Set mrs_FacturacionServicios = grdServicios.DataSource
    Set mrs_FacturacionBienes = grdBienes.DataSource
    Set mrs_ACuenta = grdACuenta.DataSource
    
    Set mo_ReglasFacturacionServicios = New Collection
    Set mo_ReglasFacturacionBienes = New Collection
    Set mo_ReglasFacturacionACuenta = New Collection
    
    Dim oDOFacturacionServicios As DOFacturacionServicios
    Dim odoFacturacionBienesInsumos As DOFacturacionBienesInsumos
    'Dim odoFacturacionACuenta As DOFacturacionPAgosACuenta
    
   ' mrs_FacturacionServicios.Open
    If Not (mrs_FacturacionServicios.EOF And mrs_FacturacionServicios.BOF) Then
        mrs_FacturacionServicios.MoveFirst
        Do While Not mrs_FacturacionServicios.EOF
            'If mrs_FacturacionServicios!EstadoRegistro = "M" Then
            If mrs_FacturacionServicios!Id <= 0 Then
                Set oDOFacturacionServicios = New DOFacturacionServicios
                oDOFacturacionServicios.IdAtencion = oAtencion.IdAtencion
                oDOFacturacionServicios.IdProducto = mrs_FacturacionServicios!IdProducto
                oDOFacturacionServicios.IdTipoFinanciamiento = Val(mo_cmbIdTipoPaciente.BoundText)
                oDOFacturacionServicios.PrecioUnitario = mrs_FacturacionServicios!PrecioUnitario
                oDOFacturacionServicios.IdPuntoCarga = 99
                If (Val(mo_cmbIdTipoPaciente.BoundText) = 1) Then
                oDOFacturacionServicios.IdFuenteFinanciamiento = 1 'debe haber una querie que saque segun tipo de paciente
                Else
                oDOFacturacionServicios.IdFuenteFinanciamiento = 11
                End If
            Else
                'Set oDOFacturacionServicios = mo_ReglasFacturacion.FacturacionServiciosSeleccionarPorId(Val(mrs_FacturacionServicios!Id))  'New DOFacturacionServicios
            End If
            
            oDOFacturacionServicios.cantidad = mrs_FacturacionServicios!cantidad
            oDOFacturacionServicios.TotalPorPagar = mrs_FacturacionServicios!TotalPorPagar
            oDOFacturacionServicios.IdEstadoFacturacion = 4 'Los almacena como estado pagado
            If oDOFacturacionServicios.IdAtencion <= 0 Then
                oDOFacturacionServicios.IdAtencion = oAtencion.IdAtencion
            End If
            oDOFacturacionServicios.IdUsuarioAuditoria = Me.IdUsuario
            
            mo_ReglasFacturacionServicios.Add oDOFacturacionServicios
                        
            mrs_FacturacionServicios.MoveNext
        Loop
        mrs_FacturacionServicios.MoveFirst
    End If
    
    If Not (mrs_FacturacionBienes.EOF And mrs_FacturacionBienes.BOF) Then
        mrs_FacturacionBienes.MoveFirst
        Do While Not mrs_FacturacionBienes.EOF
            'If mrs_FacturacionBienes!EstadoRegistro = "M" Then
            If mrs_FacturacionBienes!Id <= 0 Then
                Set odoFacturacionBienesInsumos = New DOFacturacionBienesInsumos
                odoFacturacionBienesInsumos.IdAtencion = oAtencion.IdAtencion
                odoFacturacionBienesInsumos.IdProducto = mrs_FacturacionBienes!IdProducto
                odoFacturacionBienesInsumos.IdTipoFinanciamiento = Val(mo_cmbIdTipoPaciente.BoundText)
                odoFacturacionBienesInsumos.PrecioUnitario = mrs_FacturacionBienes!PrecioUnitario
                odoFacturacionBienesInsumos.IdPuntoCarga = 99
            Else
                Set odoFacturacionBienesInsumos = mo_ReglasFacturacion.FacturacionBienesInsumosSeleccionarPorId(Val(mrs_FacturacionBienes!Id))
            End If
                odoFacturacionBienesInsumos.IdEstadoFacturacion = 4
                odoFacturacionBienesInsumos.cantidad = mrs_FacturacionBienes!cantidad
                odoFacturacionBienesInsumos.TotalPorPagar = mrs_FacturacionBienes!TotalPorPagar
                If odoFacturacionBienesInsumos.IdAtencion <= 0 Then
                    odoFacturacionBienesInsumos.IdAtencion = oAtencion.IdAtencion
                End If
                odoFacturacionBienesInsumos.IdUsuarioAuditoria = Me.IdUsuario
                mo_ReglasFacturacionBienes.Add odoFacturacionBienesInsumos
            'End If
            mrs_FacturacionBienes.MoveNext
        Loop
        mrs_FacturacionBienes.MoveFirst
    End If


'If Not (mrs_ACuenta.EOF And mrs_ACuenta.BOF) Then
'        mrs_ACuenta.MoveFirst
'        Do While Not mrs_ACuenta.EOF
'            'If mrs_ACuenta!EstadoRegistro = "M" Then
'                'Set odoFacturacionACuenta = mo_ReglasFacturacion.FacturacionACuentaSeleccionarPorId(Val(mrs_ACuenta!Id))
'                If odoFacturacionACuenta Is Nothing Then
'                    Set odoFacturacionACuenta = New DOFacturacionPAgosACuenta
'                End If
'                'odoFacturacionBienesInsumos.IdFacturacionBienes = mrs_ACuenta!id
'                odoFacturacionACuenta.FechaPago = mrs_ACuenta!FechaPago
'                odoFacturacionACuenta.IdAtencion = mrs_ACuenta!IdAtencion
'                odoFacturacionACuenta.IdEmpleadoCajero = mrs_ACuenta!IdEmpleadoCajero
''                If odoFacturacionACuenta.IdAtencion <= 0 Then
''                    odoFacturacionACuenta.IdAtencion = oAtencion.IdAtencion
''                End If
'                odoFacturacionACuenta.IdComprobantePago = mrs_ACuenta!IdComprobantePago
'                odoFacturacionACuenta.IdUsuarioAuditoria = Me.IdUsuario
'                odoFacturacionACuenta.TotalPagado = mrs_ACuenta!TotalPagado
'                mo_ReglasFacturacionACuenta.Add odoFacturacionACuenta
'            'End If
'            mrs_ACuenta.MoveNext
'        Loop
'        mrs_ACuenta.MoveFirst
'    End If

Exit Sub
errDescription:
    Set mo_ReglasFacturacionServicios = New Collection
    Set mo_ReglasFacturacionBienes = New Collection
    Set mo_ReglasFacturacionACuenta = New Collection
End Sub

Function ValidarReglas() As Boolean



    ValidarReglas = False
   
'    If mi_Opcion = sghAgregar Then
'
'    End If
    ActualizaTotales
    md_Falta = Val(grdCaja.TextMatrix(2, 1))
    md_Recibido = Val(grdCaja.TextMatrix(1, 1))
    md_Vuelto = Val(grdCaja.TextMatrix(3, 1))
    
    If md_Falta > 0 Then
        MsgBox "El valor pagado no cubre el valor de la cuenta", vbExclamation, "Facturacion"
        Exit Function
    End If
    If md_Recibido < md_Total Then
        MsgBox "El valor pagado no cubre el valor de la cuenta", vbExclamation, "Facturacion"
        Exit Function
    End If
   
   If md_Total = 0 Then
        MsgBox "No hay cuenta a pagar. No debe generar un comprobante", vbExclamation, "Facturacion"
        Exit Function
   End If
   ValidarReglas = True
End Function
Function AgregarDatos() As Boolean
    AgregarDatos = mo_AdminCaja.CajaComprobantePagoAgregar(oComprobantePago, mo_ReglasFacturacionServicios, mo_ReglasFacturacionBienes, mo_ReglasFacturacionACuenta, IdCaja)
End Function
Function NuevoComprobante()

    'limpiar variables
    NroHistoriaClinica = 0
    Set oComprobantePago = Nothing
    Set mo_CuentaAtencion = Nothing
    Set oCajaNroDocumento = Nothing
    ml_IdTipoComprobante = 0
    mb_AgregoAtencion = False
    
    Set oAtencion = Nothing
    IdPaciente = 0
    
    'ReDim idProductoSelecto(0)
    'ReDim nombreProductoSelecto(0)
    'numeroProductosSelectos = 0
    
    'variables por comprobante de pago
    md_Subtotal = 0
    md_IGV = 0
    md_Exoneraciones = 0
    md_PagosACuenta = 0
    md_Total = 0
    ml_IdComprobantePago = 0
    
    md_Recibido = 0
    md_Falta = 0
    md_Vuelto = 0
    
    mi_Opcion = sghAgregar
    
    txtNroHistoria.Text = ""
    txtNroHistoria.Tag = ""
    txtPaciente.Text = ""
    txtRuc.Text = ""
    txtPaciente.Tag = ""
    mo_cmbIdTipoPaciente.BoundText = 0
    mo_cmbIdTipoComprobante.BoundText = 0
    txtNroSerie.Text = ""
    txtNroDocumento.Text = ""
    Set grdServicios.DataSource = Nothing
    Set mrs_FacturacionServicios = Nothing
    Set grdBienes.DataSource = Nothing
    Set mrs_FacturacionBienes = Nothing
    
    Set mo_ReglasFacturacionBienes = New Collection
    Set mo_ReglasFacturacionServicios = New Collection
    
    txtIdCuentaAtencion.Text = ""
    Set grdExoneraciones.DataSource = Nothing
    Set grdACuenta.DataSource = Nothing
    
    Me.ConfigurarCaja
    Me.ConfigurarSubtotales

    cmbIdTipoPaciente.Enabled = True
    txtNroHistoria.Enabled = True
    btnHistoria.Enabled = True
    
    cmbIdTipoPaciente.SetFocus
    
    
End Function

'------------------------------------servicios------------------------------
Sub AgregaServicios()
'EFGL 14/06/2006
    If Val(mo_cmbIdTipoPaciente.BoundText) = 0 Then
        MsgBox "Debe Seleccionar un tipo de paciente", vbExclamation, "No puede agregar un servicio"
        Exit Sub
    End If
    If Val(mo_cmbIdTipoPaciente.BoundText) = 1 Then
    If mo_CuentaAtencion Is Nothing Then
        MsgBox "Debe seleccionar una cuenta de atencion", vbCritical, "Filtro Servicios"
        Exit Sub
    End If
    
    If mo_CuentaAtencion.IdCuentaAtencion <= 0 Then
        MsgBox "Debe seleccionar una cuenta de atencion", vbCritical, "Filtro Servicios"
        Exit Sub
    End If
    
    If (Not AgregarNuevaAtencion) Then
        Exit Sub
    End If
    End If
    'Obtiene el recordset de la grilla de servicios
    Set mrs_FacturacionServicios = grdServicios.DataSource
    
    'Agrega un nuevo registro al recordset
    With mrs_FacturacionServicios
        .AddNew
        .Fields!IdProducto = 0
        .Fields!codigo = ""
        .Fields!descripcion = ""
        .Fields!cantidad = 1
        .Fields!PrecioUnitario = 0
        .Fields!TotalPorPagar = 0
        .Fields!IdEstadoFacturacion = 4
        If Val(mo_cmbIdTipoPaciente.BoundText) = 1 Then
            .Fields!IdAtencion = oAtencion.IdAtencion
        Else
            .Fields!IdAtencion = 0
        End If
    End With
    
     mb_TransaccionDeNuevoRegistroEnProceso = True
    mb_NoEditar = True
    Set grdServicios.DataSource = mrs_FacturacionServicios
    mb_NoEditar = False
    
    
    'Obtiene la ultima fila agregada
    numeroFilaActiva = mrs_FacturacionServicios.RecordCount - 1
    
    mb_TransaccionDeNuevoRegistroEnProceso = True
    
    grdServicios.PerformAction ssKeyActionActivateCell
    grdServicios.PerformAction ssKeyActionEnterEditMode
    'EFGL 14/06/2006
End Sub




Public Sub BusquedaServicios()
    
    If mo_CuentaAtencion Is Nothing Then
        MsgBox "Debe seleccionar una cuenta de atencion", vbCritical, "Filtro Servicios"
        Exit Sub
    End If
    If mo_CuentaAtencion.IdCuentaAtencion <= 0 Then
        MsgBox "Debe seleccionar una cuenta de atencion", vbCritical, "Filtro Servicios"
        Exit Sub
    End If
    Dim oCatalogoServicio As New DOCatalogoServicio
    Dim rsRespuesta As New Recordset
        
        If (mo_cmbIdTipoPaciente.BoundText <> "") Then
            Dim oFrm As New CatalogoServiciosBusqueda
            oFrm.IdTipoCatalogo = Val(mo_cmbIdTipoPaciente.BoundText)
            oFrm.HabilitarTipoCatalogo = False
            oFrm.Caption = "Seleccione el Servicio"
            oFrm.Show vbModal
            If oFrm.IdRegistroSeleccionado <> 0 Then
                Call CargaDatosServicio(oFrm.IdRegistroSeleccionado)
            End If
        Else
            MsgBox "Debe seleccionar un Tipo de Paciente", vbCritical, "Filtro Servicios"
        End If
            
'        If mo_AdminAdmision.MensajeError <> "" Then
'            MsgBox mo_AdminAdmision.MensajeError, vbCritical, "Filtro Pacientes"
'        End If
        
End Sub
Sub CargaDatosServicio(IdCatalogoServicio As Long)
Dim rsRespuesta As New ADODB.Recordset
    If (Not AgregarNuevaAtencion) Then
        Exit Sub
    End If
    
    'insertando un nuevo servicio
     Dim oFacturacionServicios As New DOFacturacionServicios
     Dim oCatalogoServicios As New DOCatalogoServicio
     Dim oCatalogoServiciosHosp As New DOFinanciamientoCatalogoServ
     
     Set oCatalogoServicios = mo_ReglasFacturacion.CatalogoServiciosSeleccionarPorId(IdCatalogoServicio)
     Set oCatalogoServiciosHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(IdCatalogoServicio, Val(mo_cmbIdTipoPaciente.BoundText))
     
     With oFacturacionServicios
        .IdAtencion = oAtencion.IdAtencion
        .cantidad = 1
        .IdCentroCosto = oCatalogoServicios.IdCentroCosto
        '.IdClasificacionServicio = oCatalogoServicios.IdClasificacionServicio
        .IdComprobantePago = 0
        .IdEstadoFacturacion = 1
        .IdFuenteFinanciamiento = 1
        .IdProducto = oCatalogoServicios.IdProducto
        .IdTipoFinanciamiento = 1
        .IdUsuarioAuditoria = ml_IdUsuario
        .PrecioUnitario = oCatalogoServiciosHosp.PrecioUnitario
        .TotalPorPagar = 1 * .PrecioUnitario
     End With
      If (Not mo_ReglasFacturacion.AgregarFacturacionServicios(oFacturacionServicios)) Then
        MsgBox mo_ReglasFacturacion.MensajeError, vbCritical, "Consulte al administrador"
        Exit Sub
      End If
    'cargando nuevamente la tabla
    
    Set rsRespuesta = mo_AdminCaja.CatalogoServicioPorCuentaAtencion(mo_CuentaAtencion.IdCuentaAtencion, Val(mo_cmbIdTipoPaciente.BoundText), Val(mo_cmbIdPuntosDeCarga.BoundText), cmbIdEstado.Text)
    Set grdServicios.DataSource = rsRespuesta
    ActualizaTotales
    
End Sub



'-------------------Servicios------------------------------------------

Private Sub grdServicios_AfterRowsDeleted()
'EFGL 14/06/200
On Error GoTo errDescription
'If numeroProductosSelectos <= 0 Then
'    Exit Sub
'End If
'Dim oFacturacionServicios As New DOFacturacionServicios
'Dim i As Integer
'For i = 0 To numeroProductosSelectos - 1
'    Set oFacturacionServicios = New DOFacturacionServicios
'    oFacturacionServicios.IdFacturacionServicio = idProductoSelecto(i)
'    oFacturacionServicios.IdUsuarioAuditoria = IdUsuario
'    mo_ReglasFacturacionServiciosBorrar.Add oFacturacionServicios
'
'Next
'numeroProductosSelectos = 0
ActualizaTotales
errDescription:
'EFGL 14/06/2006
End Sub
Private Sub grdServicios_AfterRowUpdate(ByVal Row As UltraGrid.SSRow)
On Error GoTo errDescription
    ActualizaTotales
    
Exit Sub
errDescription:

End Sub
'A este evento entra cada vez que se cambia de Celda de la misma fila
Private Sub grdServicios_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
Dim precio As Double
Dim cantidad As Integer
Dim subtotal As Double
    
    If (grdServicios.ActiveCell Is Nothing) Then
        Exit Sub
    End If
    
    If grdServicios.ActiveCell.Value = "" Then
        Exit Sub
    End If
    
    
    'Si la transaccion esta en proceso y se presiona la tecla ESCAPE,
    'se eleimina el registro agregado
    If mb_TransaccionDeNuevoRegistroEnProceso Then
'        If mb_PresionoEscape Then
'            grdServicios.ActiveRow.Delete
'            mb_PresionoEscape = False
'            Exit Sub
'        End If
'
        'A este codigo entra luega que se coloca el CODIGO y se presiona TAB
        If grdServicios.ActiveCell.Column.Key = "codigo" Then
            ms_TipoProducto = "servicios"
            If Not SeteaProducto(grdServicios.ActiveCell.Value) Then
                Cancel.Value = True
            End If
        End If
    
    End If
    
    'A este codigo entra luego que se modifica la CANTDAD y se presiona TAB
    If (grdServicios.ActiveCell.Column.Key = "cantidad") Then
        cantidad = grdServicios.ActiveRow.Cells("cantidad").Value
        precio = grdServicios.ActiveRow.Cells("preciounitario").Value
        subtotal = precio * cantidad
        grdServicios.ActiveRow.Cells("totalporpagar").Value = subtotal
    End If
    
End Sub
Private Sub grdServicios_BeforeRowDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
On Error GoTo errDescription
    
    Set grillaBusqueda.DataSource = Nothing
    grillaBusqueda.Visible = False
    
    If grdServicios.ActiveRow Is Nothing Then
        Exit Sub
    End If
    
'    If grdServicios.ActiveCell Is Nothing Then
'        Exit Sub
'    End If
    
    If grdServicios.ActiveRow.Cells("idproducto") = 0 And grdServicios.ActiveRow.Cells("codigo") = "" And grdServicios.ActiveRow.Cells("descripcion") = "" Then
       If Not mb_NoEditar Then
         mb_NoEditar = True
            grdServicios.ActiveRow.Delete
            mb_NoEditar = False
       End If
    End If
    
    Exit Sub
errDescription:

End Sub
Private Sub grdServicios_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
'EFGL 14/06/2006
'    If (Cancel.Value = True) Then
'        Exit Sub
'    End If
'
'   If (grdServicios.Selected Is Nothing) Then
'        Exit Sub
'    End If
'    If (grdServicios.Selected.Rows Is Nothing) Then
'        Exit Sub
'    End If
'    If (grdServicios.Selected.Rows.Count <= 0) Then
'        Exit Sub
'    End If
'    Dim i As Integer
'
'
'    ReDim idProductoSelecto(grdServicios.Selected.Rows.Count)
'    ReDim nombreProductoSelecto(grdServicios.Selected.Rows.Count)
'    numeroProductosSelectos = grdServicios.Selected.Rows.Count
'    For i = 0 To grdServicios.Selected.Rows.Count - 1
'        idProductoSelecto(i) = grdServicios.Selected.Rows(i).Cells("id").Value
'        nombreProductoSelecto(i) = grdServicios.Selected.Rows(i).Cells("descripcion").Value
'    Next
'EFGL 14/06/2006
End Sub

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    FormatoGrilla grdServicios
End Sub
Private Sub grdServicios_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
On Error GoTo errDescription
Dim sNombre As String
    
    If grdServicios.ActiveRow Is Nothing Then
        Exit Sub
    End If
    If grdServicios.ActiveCell Is Nothing Then
        Exit Sub
    End If
    If grdServicios.ActiveCell.Column.Key <> "descripcion" And grdServicios.ActiveCell.Column.Key <> "codigo" And grdServicios.ActiveCell.Column.Key <> "cantidad" Then
        KeyAscii = 0
        Exit Sub
    End If
    
    'If mb_TransaccionDeNuevoRegistroEnProceso Then
    If grdServicios.ActiveCell.Row.Cells("idproducto").Value = 0 Then
        mb_PresionoEscape = False
        
        'El cajero presiono ESCAPE
        If KeyAscii = vbKeyEscape Then
            mb_PresionoEscape = True
            If grdServicios.ActiveCell.Row.Cells("id").Value = 0 Then
                mb_NoEditar = True
                grdServicios.ActiveRow.Delete
                mb_NoEditar = False
                grillaBusqueda.Visible = False
                Set grillaBusqueda.DataSource = Nothing
                Exit Sub
            End If
        End If
        
        'El cajero esta editando la parte de la DESCRIPCION
        If grdServicios.ActiveCell.Column.Key = "descripcion" Then
            ms_TipoProducto = "servicios"
            Select Case KeyAscii
            Case 8
                'El cajero ha presionado BACKSPACE
                sNombre = grdServicios.ActiveCell.GetText
                If Len(sNombre) > 1 Then
                    sNombre = Mid(sNombre, 1, Len(sNombre) - 1)
                End If
            Case 13, 9, 10
            Case Else
                sNombre = grdServicios.ActiveCell.GetText + Chr(KeyAscii)
            End Select
            
            Dim lIdTipoFinanciamiento As Long
            Dim lIdPuntoCarga As Long
            'EFGL 14/06/2006
            
            lIdTipoFinanciamiento = Val(mo_cmbIdTipoPaciente.BoundText)
            lIdPuntoCarga = Val(mo_cmbIdPuntosDeCarga.BoundText)
            
            BuscaProductos sNombre, lIdTipoFinanciamiento, lIdPuntoCarga
        End If
    Else
        If grdServicios.ActiveCell.Column.Key <> "cantidad" Then
            KeyAscii = 0
        End If
    End If
    
Exit Sub
errDescription:

End Sub

Private Sub grdServicios_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuServicio
    End If
End Sub

Private Sub mnuAgregarServ_Click()
    
    AgregaServicios

End Sub

Function SeteaProducto(codigo As String) As Boolean
Dim rs As New ADODB.Recordset
Dim IdTipoFinanciamiento  As Long
    'EFGL 14/06/2006
    SeteaProducto = False
    
    IdTipoFinanciamiento = Val(mo_cmbIdTipoPaciente.BoundText)
    
    If ms_TipoProducto = "bienes" Then
        Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigo(codigo, IdTipoFinanciamiento)
        Select Case rs.RecordCount
        Case 0
            MsgBox "El bien no esta disponible en el catalogó de bienes e insumos para este tipo de paciente", vbInformation, "Registro de bienes"
        Case 1
           'grdBienes.ActiveRow.Cells("id").Value = 0
           grdBienes.ActiveRow.Cells("codigo").Value = rs.Fields("CODIGO").Value
           grdBienes.ActiveRow.Cells("idproducto").Value = rs.Fields("idproducto").Value
           grdBienes.ActiveRow.Cells("descripcion").Value = rs.Fields("descripcion").Value
           grdBienes.ActiveRow.Cells("preciounitario").Value = rs.Fields("preciounitario").Value
           grdBienes.ActiveRow.Cells("totalporpagar").Value = rs.Fields("preciounitario").Value
           grdBienes.ActiveRow.Cells("cantidad").Value = 1
           grdBienes.ActiveRow.Cells("idestadofacturacion").Value = 4
           SeteaProducto = True
        Case Else
            MsgBox "El bien esta disponible mas de una vez en el catàlogo de bienes e insumos para este tipo de paciente, revise el catalogo de bienes e insumos", vbInformation, "Registro de bienes e insumos"
            grdBienes.PerformAction ssKeyActionActivateCell
            grdBienes.PerformAction ssKeyActionEnterEditMode
        
        End Select
           
    Else
        Set rs = mo_ReglasFacturacion.FacturacionServicioPorCodigo(codigo, IdTipoFinanciamiento)
        Select Case rs.RecordCount
        Case 0
            MsgBox "El servicio no esta disponible en el catalogó de servicios para este tipo de paciente", vbInformation, "Registro de servicios"
            'grdServicios.PerformAction ssKeyActionActivateCell
            'grdServicios.PerformAction ssKeyActionEnterEditMode
        
        Case 1
            'grdServicios.ActiveRow.Cells("id").Value = 0
            grdServicios.ActiveRow.Cells("codigo").Value = rs.Fields("codigo").Value
            grdServicios.ActiveRow.Cells("idproducto").Value = rs.Fields("idproducto").Value
            grdServicios.ActiveRow.Cells("descripcion").Value = rs.Fields("descripcion").Value
            grdServicios.ActiveRow.Cells("preciounitario").Value = rs.Fields("preciounitario").Value
            grdServicios.ActiveRow.Cells("totalporpagar").Value = rs.Fields("preciounitario").Value
            grdServicios.ActiveRow.Cells("cantidad").Value = 1
            grdServicios.ActiveRow.Cells("idestadofacturacion").Value = 4
            SeteaProducto = True
        Case Else
            MsgBox "El servicio esta disponible mas de una vez en el catalogó de servicios para este tipo de paciente, revise el catalogó de servicios", vbInformation, "Registro de servicios"
            grdServicios.PerformAction ssKeyActionActivateCell
            grdServicios.PerformAction ssKeyActionEnterEditMode
        
        End Select
    
    
    End If
       grillaBusqueda.Visible = False
       Set grillaBusqueda.DataSource = Nothing
       'EFGL 14/06/2006
End Function


'-----------------------------------------Bienes Insumos-------------------------
Sub AgregaBienesInsumos()
'EFGL 14/06/2006
    If Val(mo_cmbIdTipoPaciente.BoundText) = 0 Then
        MsgBox "Debe Seleccionar un tipo de paciente", vbExclamation, "No puede agregar un bien"
        Exit Sub
    End If
    If Val(mo_cmbIdTipoPaciente.BoundText) = 1 Then
     If mo_CuentaAtencion Is Nothing Then
        MsgBox "Debe seleccionar una cuenta de atencion", vbCritical, "Filtro Bienes"
        Exit Sub
    End If
    
    If mo_CuentaAtencion.IdCuentaAtencion <= 0 Then
        MsgBox "Debe seleccionar una cuenta de atencion", vbCritical, "Filtro Bienes"
        Exit Sub
    End If
    
    If (Not AgregarNuevaAtencion) Then
        Exit Sub
    End If
    End If
    'Obtiene el recordset de la grilla de servicios
    Set mrs_FacturacionBienes = grdBienes.DataSource
    
    'Agrega un nuevo registro al recordset
    With mrs_FacturacionBienes
        .AddNew
        .Fields!IdProducto = 0
        .Fields!codigo = ""
        .Fields!descripcion = ""
        .Fields!cantidad = 1
        .Fields!PrecioUnitario = 0
        .Fields!TotalPorPagar = 0
        .Fields!IdEstadoFacturacion = 4
        If Val(mo_cmbIdTipoPaciente.BoundText) = 1 Then
            .Fields!IdAtencion = oAtencion.IdAtencion
        Else
            .Fields!IdAtencion = 0
        End If
    End With
    
     mb_TransaccionDeNuevoRegistroEnProceso = True
    mb_NoEditar = True
    Set grdBienes.DataSource = mrs_FacturacionBienes
    mb_NoEditar = False
    
    
    'Obtiene la ultima fila agregada
    numeroFilaActiva = mrs_FacturacionBienes.RecordCount - 1
    
    mb_TransaccionDeNuevoRegistroEnProceso = True
    
    grdBienes.PerformAction ssKeyActionActivateCell
    grdBienes.PerformAction ssKeyActionEnterEditMode
'EFGL 14/06/2006
End Sub
Public Sub BusquedaBienesInsumos()
    
    If mo_CuentaAtencion Is Nothing Then
        MsgBox "Debe seleccionar una cuenta de atencion", vbCritical, "Filtro Bienes e Insumos"
        Exit Sub
    End If
    If mo_CuentaAtencion.IdCuentaAtencion <= 0 Then
        MsgBox "Debe seleccionar una cuenta de atencion", vbCritical, "Filtro Bienes e Insumos"
        Exit Sub
    End If

    Dim oCatalogoBien As New DOCatalogoBienesInsumos
    Dim rsRespuesta As New Recordset
        
        If (mo_cmbIdTipoPaciente.BoundText <> "") Then
            Dim oFrm As New CatalogoBienesInsumosBusqueda
            oFrm.IdTipoCatalogo = Val(mo_cmbIdTipoPaciente.BoundText)
            oFrm.HabilitarTipoCatalogo = False
            oFrm.Caption = "Seleccione el BienInsumo"
            oFrm.Show vbModal
            If oFrm.IdRegistroSeleccionado <> 0 Then
                Call CargaDatosBienesInsumos(oFrm.IdRegistroSeleccionado)
            End If
        Else
            MsgBox "Debe seleccionar un Tipo de Paciente", vbCritical, "Filtro Bienes e Insumos"
        End If
            
'        If mo_AdminAdmision.MensajeError <> "" Then
'            MsgBox mo_AdminAdmision.MensajeError, vbCritical, "Filtro Pacientes"
'        End If
        
End Sub
Sub CargaDatosBienesInsumos(IdCatalogoBienInsumo As Long)
Dim rsRespuesta As New ADODB.Recordset
    If (Not AgregarNuevaAtencion) Then
        Exit Sub
    End If
    
    'insertando un nuevo servicio
     Dim oFacturacionBienes As New DOFacturacionBienesInsumos
     Dim oCatalogoBienes As New DOCatalogoBienesInsumos
     Dim oCatalogoBienesHosp As New DoFinanciamientoCatalogoBien
     
     Set oCatalogoBienes = mo_ReglasFacturacion.CatalogoBienesSeleccionarPorId(IdCatalogoBienInsumo)
     Set oCatalogoBienesHosp = mo_ReglasFacturacion.CatalogoBienesHospSeleccionarPorId(IdCatalogoBienInsumo, Val(mo_cmbIdTipoPaciente.BoundText))
     
     With oFacturacionBienes
        .IdAtencion = oAtencion.IdAtencion
        .cantidad = 1
        .IdCentroCosto = oCatalogoBienes.IdCentroCosto
        .IdComprobantePago = 0
        .IdEstadoFacturacion = 1
        .IdFuenteFinanciamiento = 1
        .IdPartidaPresupuestal = oCatalogoBienes.IdPartida
        .IdProducto = oCatalogoBienes.IdProducto
        .IdReceta = 0
        '.IdTipoBienInsumo = oCatalogoBienes.IdClasificacionBienInsumo
        .IdTipoFinanciamiento = 1
        .IdUsuarioAuditoria = ml_IdUsuario
        .PrecioUnitario = oCatalogoBienesHosp.PrecioUnitario
        .TotalPorPagar = 1 * .PrecioUnitario
     End With
      If (Not mo_ReglasFacturacion.AgregarFacturacionBienes(oFacturacionBienes)) Then
        MsgBox mo_ReglasFacturacion.MensajeError, vbCritical, "Consulte al administrador"
        Exit Sub
      End If
    'cargando nuevamente la tabla
    
    Set rsRespuesta = mo_AdminCaja.CatalogoBienesInsumosPorCuentaAtencion(mo_CuentaAtencion.IdCuentaAtencion, Val(mo_cmbIdTipoPaciente.BoundText), Val(mo_cmbIdPuntosDeCarga.BoundText), cmbIdEstado.Text)
    Set grdBienes.DataSource = rsRespuesta
    ActualizaTotales
End Sub



Private Sub grdBienes_AfterRowsDeleted()
'EFGL 14/06/2006
On Error GoTo errDescription
'If numeroProductosSelectos <= 0 Then
'    Exit Sub
'End If
'Dim oFacturacionBienes As New doFacturacionBienesInsumos
'Dim i As Integer
'For i = 0 To numeroProductosSelectos - 1
'    Set oFacturacionBienes = New doFacturacionBienesInsumos
'    oFacturacionBienes.IdFacturacionBienes = idProductoSelecto(i)
'    oFacturacionBienes.IdUsuarioAuditoria = IdUsuario
'    'If (Not mo_ReglasFacturacion.EliminarFacturacionBienes(oFacturacionBienes)) Then
'    '   MsgBox "No se pudo  eliminar el bien e insumo: " & Trim(nombreProductoSelecto(i))
'    'End If
'    mo_ReglasFacturacionBienesBorrar.Add oFacturacionBienes
'Next
'numeroProductosSelectos = 0
ActualizaTotales
errDescription:
'EFGL 14/06/2006
End Sub
Private Sub grdBienes_AfterRowUpdate(ByVal Row As UltraGrid.SSRow)
    On Error GoTo errDescription
    ActualizaTotales
    
Exit Sub
errDescription:

End Sub
Private Sub grdBienes_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
   
Dim precio As Double
Dim cantidad As Integer
Dim subtotal As Double
    
    If (grdBienes.ActiveCell Is Nothing) Then
        Exit Sub
    End If
    
    If grdBienes.ActiveCell.Value = "" Then
        Exit Sub
    End If
    
    
    'Si la transaccion esta en proceso y se presiona la tecla ESCAPE,
    'se eleimina el registro agregado
    If mb_TransaccionDeNuevoRegistroEnProceso Then
'        If mb_PresionoEscape Then
'            grdbienes.ActiveRow.Delete
'            mb_PresionoEscape = False
'            Exit Sub
'        End If
'
        'A este codigo entra luega que se coloca el CODIGO y se presiona TAB
        If grdBienes.ActiveCell.Column.Key = "codigo" Then
            ms_TipoProducto = "bienes"
            If Not SeteaProducto(grdBienes.ActiveCell.Value) Then
                Cancel.Value = True
            End If
        End If
    
    End If
    
    'A este codigo entra luego que se modifica la CANTDAD y se presiona TAB
    If (grdBienes.ActiveCell.Column.Key = "cantidad") Then
        cantidad = grdBienes.ActiveRow.Cells("cantidad").Value
        precio = grdBienes.ActiveRow.Cells("preciounitario").Value
        subtotal = precio * cantidad
        grdBienes.ActiveRow.Cells("totalporpagar").Value = subtotal
    End If
    
End Sub
Private Sub grdBienes_BeforeRowDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
On Error GoTo errDescription
    
    Set grillaBusqueda.DataSource = Nothing
    grillaBusqueda.Visible = False
    
    If grdBienes.ActiveRow Is Nothing Then
        Exit Sub
    End If
    
    If grdBienes.ActiveRow.Cells("idproducto") = 0 And grdBienes.ActiveRow.Cells("codigo") = "" And grdBienes.ActiveRow.Cells("descripcion") = "" Then
       If Not mb_NoEditar Then
         mb_NoEditar = True
            grdBienes.ActiveRow.Delete
            mb_NoEditar = False
       End If
    End If
    
    Exit Sub
errDescription:

End Sub
Private Sub grdBienes_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
'EFGL 14/06/2006
'    If (Cancel.Value = True) Then
'        Exit Sub
'    End If
'   If (grdBienes.Selected Is Nothing) Then
'        Exit Sub
'    End If
'    If (grdBienes.Selected.Rows Is Nothing) Then
'        Exit Sub
'    End If
'    If (grdBienes.Selected.Rows.Count <= 0) Then
'        Exit Sub
'    End If
'    Dim i As Integer
'
'    ReDim idProductoSelecto(grdBienes.Selected.Rows.Count)
'    ReDim nombreProductoSelecto(grdBienes.Selected.Rows.Count)
'    numeroProductosSelectos = grdBienes.Selected.Rows.Count
'    For i = 0 To grdBienes.Selected.Rows.Count - 1
'        idProductoSelecto(i) = grdBienes.Selected.Rows(i).Cells("id").Value
'        nombreProductoSelecto(i) = grdBienes.Selected.Rows(i).Cells("descripcion").Value
'    Next
'EFGL 14/06/2006
End Sub
Private Sub grdBienes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    FormatoGrilla grdBienes
End Sub
Private Sub grdBienes_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
On Error GoTo errDescription
Dim sNombre As String
    
    If grdBienes.ActiveRow Is Nothing Then
        Exit Sub
    End If
    If grdBienes.ActiveCell Is Nothing Then
        Exit Sub
    End If
    If grdBienes.ActiveCell.Column.Key <> "descripcion" And grdBienes.ActiveCell.Column.Key <> "codigo" And grdBienes.ActiveCell.Column.Key <> "cantidad" Then
        KeyAscii = 0
        Exit Sub
    End If
    
    'If mb_TransaccionDeNuevoRegistroEnProceso Then
    If grdBienes.ActiveCell.Row.Cells("idproducto").Value = 0 Then
        mb_PresionoEscape = False
        
        'El cajero presiono ESCAPE
        If KeyAscii = vbKeyEscape Then
            mb_PresionoEscape = True
            If grdBienes.ActiveCell.Row.Cells("id").Value = 0 Then
                mb_NoEditar = True
                grdBienes.ActiveRow.Delete
                mb_NoEditar = False
                grillaBusqueda.Visible = False
                Set grillaBusqueda.DataSource = Nothing
                Exit Sub
            End If
        End If
        
        'El cajero esta editando la parte de la DESCRIPCION
        If grdBienes.ActiveCell.Column.Key = "descripcion" Then
            ms_TipoProducto = "bienes"
            'EFGL 14/06/2006
            Select Case KeyAscii
            Case 8
                'El cajero ha presionado BACKSPACE
                sNombre = grdBienes.ActiveCell.GetText
                If Len(sNombre) > 1 Then
                    sNombre = Mid(sNombre, 1, Len(sNombre) - 1)
                End If
            Case 13, 9, 10
            Case Else
                sNombre = grdBienes.ActiveCell.GetText + Chr(KeyAscii)
            End Select
            
            Dim lIdTipoFinanciamiento As Long
            Dim lIdPuntoCarga As Long
            
            lIdTipoFinanciamiento = Val(mo_cmbIdTipoPaciente.BoundText)
            lIdPuntoCarga = Val(mo_cmbIdPuntosDeCarga.BoundText)
            
            BuscaProductos sNombre, lIdTipoFinanciamiento, lIdPuntoCarga
            
            'EFGL 14/06/2006
        End If
    Else
        If grdBienes.ActiveCell.Column.Key <> "cantidad" Then
            KeyAscii = 0
        End If
    End If
    
Exit Sub
errDescription:

End Sub
Private Sub grdBienes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuBienes
    End If
End Sub


Private Sub mnuAgregarBien_Click()
    'BusquedaBienesInsumos
    Me.AgregaBienesInsumos
    
End Sub
Function NuevoPagoACuenta()
    
'EFGL 14/06/2006
    If Val(mo_cmbIdTipoPaciente.BoundText) <> 1 Then
        MsgBox "Solo puede agregar un pago a cuenta para un paciente normal", vbExclamation, "No se agregarà el pago a cuenta"
        Exit Function
    End If
    If (Not AgregarNuevaAtencion) Then
        Exit Function
    End If
    
    If mo_CuentaAtencion Is Nothing Then
        Exit Function
    End If
    
    If mo_CuentaAtencion.IdCuentaAtencion = 0 Then
        Exit Function
    End If
    
    Set mrs_ACuenta = grdACuenta.DataSource
    
    With mrs_ACuenta
                 .AddNew
                 '.Fields!id = 0
                 .Fields!IdAtencion = oAtencion.IdAtencion
                 .Fields!FechaPago = Now
                 '.Fields("Empleado").Value = Me.NombreUsuario
                 '.Fields!NombresEmpleado = ""
                 .Fields!IdEmpleadoCajero = Me.IdUsuario
                 .Fields!IdComprobantePago = 0
                 .Fields!TotalPagado = 0
   End With
   Set grdACuenta.DataSource = mrs_ACuenta
   ActualizaTotales
End Function
