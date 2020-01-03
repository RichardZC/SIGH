VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form RegistroComprobantes 
   Caption         =   "Form1"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14820
   LinkTopic       =   "Form1"
   ScaleHeight     =   10305
   ScaleWidth      =   14820
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   2265
      Left            =   90
      TabIndex        =   41
      Top             =   7980
      Width           =   7245
      Begin VB.CommandButton btnAgregarDinero 
         Caption         =   "Agregar[F6]"
         DisabledPicture =   "RegistroComprobantes.frx":0000
         DownPicture     =   "RegistroComprobantes.frx":03E9
         Height          =   615
         Left            =   6090
         Picture         =   "RegistroComprobantes.frx":07F5
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   210
         Width           =   1005
      End
      Begin VB.CommandButton btnQuitarDinero 
         Caption         =   "Quitar[F7]"
         DisabledPicture =   "RegistroComprobantes.frx":0C01
         DownPicture     =   "RegistroComprobantes.frx":0F8C
         Height          =   615
         Left            =   6090
         Picture         =   "RegistroComprobantes.frx":131F
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   900
         Width           =   1005
      End
      Begin UltraGrid.SSUltraGrid grdDinero 
         Height          =   1920
         Left            =   120
         TabIndex        =   7
         Top             =   210
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   3387
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Dinero Recibido"
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6495
      Left            =   13560
      TabIndex        =   36
      Top             =   1470
      Width           =   1605
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "RegistroComprobantes.frx":16B0
         DownPicture     =   "RegistroComprobantes.frx":1B74
         Height          =   700
         Left            =   120
         Picture         =   "RegistroComprobantes.frx":2060
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2520
         Width           =   1365
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "RegistroComprobantes.frx":254C
         DownPicture     =   "RegistroComprobantes.frx":29AC
         Height          =   700
         Left            =   120
         Picture         =   "RegistroComprobantes.frx":2E21
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir [F3]"
         Height          =   705
         Left            =   120
         Picture         =   "RegistroComprobantes.frx":3296
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1020
         Width           =   1365
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo [F5]"
         DisabledPicture =   "RegistroComprobantes.frx":376F
         DownPicture     =   "RegistroComprobantes.frx":3B58
         Height          =   700
         Left            =   120
         Picture         =   "RegistroComprobantes.frx":3F64
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1800
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   7410
      TabIndex        =   24
      Top             =   7980
      Width           =   7755
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto Facturado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   35
         Top             =   120
         Width           =   1770
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Recibido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3330
         TabIndex        =   34
         Top             =   120
         Width           =   900
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Faltan (S/.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         TabIndex        =   33
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Vuelto (S/.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6270
         TabIndex        =   32
         Top             =   120
         Width           =   1230
      End
      Begin VB.Label lblMontoRecibidoDolares 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F2DED9&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2970
         TabIndex        =   31
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label Label16 
         Caption         =   "SOLES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   30
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label Label17 
         Caption         =   "DOLARES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   29
         Top             =   1110
         Width           =   1335
      End
      Begin VB.Label lblMontoFacturadoSoles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F2DED9&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1380
         TabIndex        =   28
         Top             =   450
         Width           =   1500
      End
      Begin VB.Label lblMontoRecibidoSoles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F2DED9&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2970
         TabIndex        =   27
         Top             =   450
         Width           =   1500
      End
      Begin VB.Label lblMontoFaltanteSoles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F2DED9&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   465
         Left            =   4560
         TabIndex        =   26
         Top             =   450
         Width           =   1500
      End
      Begin VB.Label lblMontoVueltoSoles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F2DED9&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   465
         Left            =   6120
         TabIndex        =   25
         Top             =   450
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   8460
      TabIndex        =   16
      Top             =   60
      Width           =   6705
      Begin VB.ComboBox cmbIdTipoComprobante 
         Height          =   315
         Left            =   1650
         TabIndex        =   5
         Top             =   210
         Width           =   3765
      End
      Begin VB.TextBox txtNroSerie 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F2DED9&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1650
         TabIndex        =   18
         Top             =   600
         Width           =   825
      End
      Begin VB.TextBox txtNroDocumento 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F2DED9&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2580
         TabIndex        =   17
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc:"
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
         TabIndex        =   23
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
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
         Left            =   180
         TabIndex        =   21
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Cambio:"
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
         TabIndex        =   20
         Top             =   1020
         Width           =   1200
      End
      Begin VB.Label lblTipoCambio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F2DED9&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1650
         TabIndex        =   19
         Top             =   990
         Width           =   1605
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Datos del paciente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   90
      TabIndex        =   10
      Top             =   60
      Width           =   8310
      Begin VB.CommandButton btnLeerDatos 
         Caption         =   "..."
         Height          =   345
         Left            =   2460
         TabIndex        =   42
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtRUC 
         Height          =   315
         Left            =   1110
         TabIndex        =   2
         Top             =   990
         Width           =   1605
      End
      Begin VB.TextBox txtIdNroHistoria 
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
         Left            =   4035
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1320
      End
      Begin VB.ComboBox cmbIdTipoGenHistoriaClinica 
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
         Left            =   5460
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   2715
      End
      Begin VB.TextBox txtIdCuentaAtencion 
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
         Left            =   1125
         TabIndex        =   0
         Top             =   255
         Width           =   1260
      End
      Begin VB.TextBox txtRazonSocial 
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
         Left            =   1110
         TabIndex        =   1
         Top             =   630
         Width           =   7050
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "RUC:"
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
         TabIndex        =   22
         Top             =   1020
         Width           =   390
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Razón Social"
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
         Left            =   90
         TabIndex        =   12
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Nº historia"
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
         Left            =   3090
         TabIndex        =   11
         Top             =   285
         Width           =   975
      End
   End
   Begin TabDlg.SSTab tabExoneracion 
      Height          =   6465
      Left            =   90
      TabIndex        =   14
      Top             =   1500
      Width           =   13395
      _ExtentX        =   23627
      _ExtentY        =   11404
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabPicture(0)   =   "RegistroComprobantes.frx":4370
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdServicios"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Bienes e Insumos"
      TabPicture(1)   =   "RegistroComprobantes.frx":438C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdBienes"
      Tab(1).ControlCount=   1
      Begin UltraGrid.SSUltraGrid grdServicios 
         Height          =   5895
         Left            =   150
         TabIndex        =   6
         Top             =   420
         Width           =   13125
         _ExtentX        =   23151
         _ExtentY        =   10398
         _Version        =   131072
         GridFlags       =   17040388
         UpdateMode      =   2
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lista de Servicios"
      End
      Begin UltraGrid.SSUltraGrid grdBienes 
         Height          =   7005
         Left            =   -74880
         TabIndex        =   15
         Top             =   420
         Width           =   13125
         _ExtentX        =   23151
         _ExtentY        =   12356
         _Version        =   131072
         GridFlags       =   17040388
         UpdateMode      =   2
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lista de Bienes e Insumos"
      End
   End
End
Attribute VB_Name = "RegistroComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: POAtencionesInterconsultas
'        Autor: William Castro Grijalva
'        Fecha: 31/10/2004 09:32:29 a.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------
Const ID_TIPO_MONEDA_SOLES = 1
Const ID_TIPO_MONEDA_DOLAR = 2

Const ID_TIPO_COMPROBANTE_FACTURA = 2

Dim ml_IdComprobantePago As Long
Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim mo_AdminSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_CajaComprobantesPago As New DOCajaComprobantesPago
Dim mo_ItemsAFacturar As New Collection
Dim mo_ItemsDinero As New Collection

Dim mo_cmbIdTipoComprobante As New SIGHComun.ListaDespleglable

Dim mrs_ComprobantesDetalle As ADODB.Recordset
Dim mrs_FormaPago As New ADODB.Recordset
Dim mo_Apariencia As New SIGHComun.GridInfragistic

Dim md_PorcentajeIGV  As Double
Dim md_TipoCambioDolar As Double
Dim mo_CajaLoteActual As New DOCajaLote
Dim mo_CajaCajaActual As New DOCajaCaja

Dim bCalculandoSubTotales As Boolean

Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico

Dim ml_IdCuentaAtencion As Long
Dim mo_cmbIdTipoGenHistoriaClinica As New ListaDespleglable

Property Set CajaLoteActual(oValue As DOCajaLote)
   Set mo_CajaLoteActual = oValue
   'Ubicamos la caja en función del Lote
   Set mo_CajaCajaActual = mo_AdminCaja.CajaSeleccionarPorId(mo_CajaLoteActual.IdCaja)
End Property
Property Get CajaLoteActual() As DOCajaLote
   Set CajaLoteActual = mo_CajaLoteActual
End Property

Property Let IdCuentaAtencion(Value As Long)
    ml_IdCuentaAtencion = Value
End Property
Property Get IdCuentaAtencion() As Long
    IdCuentaAtencion = ml_IdCuentaAtencion
End Property

Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
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
Private Sub btnAceptar_Click()
    
    If MsgBox("Por favor confirmar, ¿Realmente desea grabar los cambios que ha realizado?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        Me.grdServicios.Update
        Me.grdBienes.Update
        
        Set Me.grdServicios.DataSource = mo_AdminFacturacion.FacturacionServiciosObtenerParaPendientePago(ml_IdCuentaAtencion)
        Set Me.grdBienes.DataSource = mo_AdminFacturacion.FacturacionBienesInsumosObtenerParaPendientePago(ml_IdCuentaAtencion)
    End If
    
End Sub

Private Sub btnCancelar_Click()
    If MsgBox("Por favor confirmar, ¿Realmente desea salir?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        Me.Visible = False
    End If
End Sub
Sub CargarComboBoxes()
Dim sSQL As String
       
    mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
    mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
    Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()

    mo_cmbIdTipoComprobante.BoundColumn = "IdTipoComprobante"
    mo_cmbIdTipoComprobante.ListField = "Descripcion"
    Set mo_cmbIdTipoComprobante.RowSource = mo_AdminCaja.TiposComprobanteSeleccionarTodos()

End Sub

Private Sub btnAgregarDinero_Click()
    With mrs_FormaPago
        .AddNew
        If Me.grdDinero.ValueLists("TipoFormaPago").ValueListItems.Count > 0 Then
            .Fields!IdTipoFormaPago = Me.grdDinero.ValueLists("TipoFormaPago").ValueListItems(0).DataValue
        End If
        If Me.grdDinero.ValueLists("TipoMoneda").ValueListItems.Count > 0 Then
            .Fields!IdTipoMoneda = Me.grdDinero.ValueLists("TipoMoneda").ValueListItems(0).DataValue
        End If
        .Update
    End With
    grdDinero.SetFocus
    CalcularSubTotalesDinero
End Sub

Private Sub btnLeerDatos_Click()
    
    ObtenerDatosDePaciente

    'Cargar datos de servicios
    Set Me.grdServicios.DataSource = mo_AdminFacturacion.FacturacionServiciosObtenerParaCaja(ml_IdCuentaAtencion)
    Set Me.grdBienes.DataSource = mo_AdminFacturacion.FacturacionBienesInsumosObtenerParaCaja(ml_IdCuentaAtencion)
    
    ConfigurarListaDesplegablesDeServicio
    ConfigurarListaDesplegablesDeBienes
    
    'mo_Formulario.HabilitarDeshabilitar txtIdCuentaAtencion, False
    mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    'mo_Formulario.HabilitarDeshabilitar txtRazonSocial, False
    
End Sub

Private Sub btnQuitarDinero_Click()
    On Error Resume Next
    With mrs_FormaPago
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
        .MoveFirst
    End With
    CalcularSubTotalesDinero
End Sub

Private Sub cmdGrabar_Click()
   
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If AgregarDatos() Then
                    MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                    NuevoComprobante
                Else
                    MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If ModificarDatos() Then
                    MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
                CargaDatosAlObjetosDeDatos
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select

End Sub

Private Sub cmdNuevo_Click()
    NuevoComprobante
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    CargarComboBoxes
    
    GenerarRecordsetTemporal
    
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Comprobante de Pago"
    Case sghModificar
        Me.Caption = "Modificar Comprobante de Pago"
    Case sghConsultar
        Me.Caption = "Consultar Comprobante de Pago"
    Case sghEliminar
        Me.Caption = "Eliminar Comprobante de Pago"
    End Select
    CargarComboBoxes
    CargarDatosAlFormulario
    
End Sub
Private Sub CargarDatosGenerales()
    Dim oTipoMoneda As New DOCajaTiposMoneda
    oTipoMoneda.IdTipoMoneda = ID_TIPO_MONEDA_DOLAR
    md_TipoCambioDolar = mo_AdminCaja.CajaTipoCambioActualMoneda(oTipoMoneda)
    lblTipoCambio.Caption = Format(md_TipoCambioDolar, "0.00")
    md_PorcentajeIGV = mo_AdminCaja.ImpuestoIGV / 100#
End Sub
Private Sub NuevoComprobante()
    ml_IdComprobantePago = 0
    mi_Opcion = sghOpciones.sghAgregar
    If cmbIdTipoComprobante.ListCount > 0 Then
        cmbIdTipoComprobante.ListIndex = 0
    End If
    Me.txtIdCuentaAtencion = ""
    Me.txtRazonSocial = ""
    Me.txtRUC = ""
    
    'GenerarRecordsetTemporal
    'CalcularSubTotalesItems
    'CalcularSubTotalesDinero
    'CalcularVuelto
    
End Sub
Sub CargarDatosAlFormulario()
    
    CargarDatosGenerales
        
    Select Case mi_Opcion
        Case sghAgregar
            NuevoComprobante
        Case sghModificar
            CargarDatosALosControles
        Case sghConsultar
            CargarDatosALosControles
        Case sghEliminar
            CargarDatosALosControles
    End Select
End Sub
Private Sub ConfigurarListaDesplegablesDeServicio()
Dim oValueList As SSValueList
Dim i As Long
Dim rsEstado As ADODB.Recordset

    On Error Resume Next
    'Crea lista de estadps de facturacion
    Set oValueList = Me.grdServicios.ValueLists.Add("EstadoFacturacion")
    Set rsEstado = mo_AdminFacturacion.EstadosFacturacionObtenerTodos()
    rsEstado.MoveFirst
    For i = 0 To rsEstado.RecordCount - 1
        oValueList.ValueListItems.Add rsEstado.Fields("IdEstadoFacturacion").Value, rsEstado.Fields("Descripcion").Value
        rsEstado.MoveNext
    Next i
    Me.grdServicios.Bands(1).Columns("IdEstadoFacturacion").ValueList = oValueList
    Me.grdServicios.Bands(1).Columns("IdEstadoFacturacion").ValueList.DisplayStyle = ssValueListDisplayStyleDisplayText
    Me.grdServicios.Bands(1).Columns("IdEstadoFacturacion").Style = ssStyleDropDownList
    rsEstado.Close

    'Crea lista de empleados que han autorizado pendiente incluyendo el nuevo empleado
    Set oValueList = Me.grdServicios.ValueLists.Add("Empleados")
    Set rsEstado = mo_AdminFacturacion.EmpleadosSeleccionarParaPendientePagoServicio(ml_IdCuentaAtencion, ml_IdUsuario)
    rsEstado.MoveFirst
    For i = 0 To rsEstado.RecordCount - 1
        oValueList.ValueListItems.Add rsEstado.Fields("IdEmpleado").Value, rsEstado.Fields("Nombre").Value
        rsEstado.MoveNext
    Next i
    Me.grdServicios.Bands(1).Columns("IdEmpleadoModifica").ValueList = oValueList
    Me.grdServicios.Bands(1).Columns("IdEmpleadoModifica").ValueList.DisplayStyle = ssValueListDisplayStyleDisplayText
    Me.grdServicios.Bands(1).Columns("IdEmpleadoModifica").Style = ssStyleDropDownList
    rsEstado.Close

End Sub
Private Sub ConfigurarListaDesplegablesDeBienes()
Dim oValueList As SSValueList
Dim i As Long
Dim rsEstado As ADODB.Recordset

    On Error Resume Next
    'Crea lista de estadps de facturacion
    Set oValueList = Me.grdBienes.ValueLists.Add("EstadoFacturacion")
    Set rsEstado = mo_AdminFacturacion.EstadosFacturacionObtenerTodos()
    rsEstado.MoveFirst
    For i = 0 To rsEstado.RecordCount - 1
        oValueList.ValueListItems.Add rsEstado.Fields("IdEstadoFacturacion").Value, rsEstado.Fields("Descripcion").Value
        rsEstado.MoveNext
    Next i
    Me.grdBienes.Bands(1).Columns("IdEstadoFacturacion").ValueList = oValueList
    Me.grdBienes.Bands(1).Columns("IdEstadoFacturacion").ValueList.DisplayStyle = ssValueListDisplayStyleDisplayText
    Me.grdBienes.Bands(1).Columns("IdEstadoFacturacion").Style = ssStyleDropDownList
    rsEstado.Close

    'Crea lista de empleados que han autorizado pendiente incluyendo el nuevo empleado
    Set oValueList = Me.grdBienes.ValueLists.Add("Empleados")
    Set rsEstado = mo_AdminFacturacion.EmpleadosSeleccionarParaPendientePagoBienInsumo(ml_IdCuentaAtencion, ml_IdUsuario)
    rsEstado.MoveFirst
    For i = 0 To rsEstado.RecordCount - 1
        oValueList.ValueListItems.Add rsEstado.Fields("IdEmpleado").Value, rsEstado.Fields("Nombre").Value
        rsEstado.MoveNext
    Next i
    Me.grdBienes.Bands(1).Columns("IdEmpleadoAutorizaPendiente").ValueList = oValueList
    Me.grdBienes.Bands(1).Columns("IdEmpleadoAutorizaPendiente").ValueList.DisplayStyle = ssValueListDisplayStyleDisplayText
    Me.grdBienes.Bands(1).Columns("IdEmpleadoAutorizaPendiente").Style = ssStyleDropDownList
    rsEstado.Close

End Sub

Private Sub Form_Resize()

'    On Error Resume Next
'    Me.tabExoneracion.Width = Me.Width - 240
'    Me.tabExoneracion.Height = Me.Height - Me.Frame4.Height - Me.fraDatos.Height - 640
'
'    Me.grdServicios.Width = Me.tabExoneracion.Width - 240
'    Me.grdServicios.Height = Me.tabExoneracion.Height - 560
'
'    Me.grdBienes.Width = Me.tabExoneracion.Width - 240
'    Me.grdBienes.Height = Me.tabExoneracion.Height - 560
'
'    Me.fraDatos.Width = Me.tabExoneracion.Width
'
'    Me.Frame4.Width = Me.tabExoneracion.Width
'    Me.Frame4.Left = Me.tabExoneracion.Left
'    Me.Frame4.Top = Me.tabExoneracion.Top + Me.tabExoneracion.Height
End Sub

Function CalculaTotalPendientePorCategoria(oRowParent As SSRow)
Dim oRow As SSRow

    Dim cTotal As Currency
    Set oRow = oRowParent.GetChild(ssChildRowFirst)
    cTotal = oRow.Cells("SubTotalPendientePago").Value
    Do While oRow.HasNextSibling
        Set oRow = oRow.GetSibling(ssSiblingRowNext)
        cTotal = cTotal + oRow.Cells("SubTotalPendientePago").Value
    Loop
    CalculaTotalPendientePorCategoria = cTotal
End Function
Function CalculaTotalPorPagarPorCategoria(oRowParent As SSRow)
Dim oRow As SSRow

    Dim cTotal As Currency
    Set oRow = oRowParent.GetChild(ssChildRowFirst)
    cTotal = oRow.Cells("SubTotalPorPagar").Value
    Do While oRow.HasNextSibling
        Set oRow = oRow.GetSibling(ssSiblingRowNext)
        cTotal = cTotal + oRow.Cells("SubTotalPorPagar").Value
    Loop
    CalculaTotalPorPagarPorCategoria = cTotal
End Function
Function CalculaTotalPagadoACuentaPorCategoria(oRowParent As SSRow)
Dim oRow As SSRow

    Dim cTotal As Currency
    Set oRow = oRowParent.GetChild(ssChildRowFirst)
    cTotal = oRow.Cells("SubTotalPagadoACuenta").Value
    Do While oRow.HasNextSibling
        Set oRow = oRow.GetSibling(ssSiblingRowNext)
        cTotal = cTotal + oRow.Cells("SubTotalPagadoACuenta").Value
    Loop
    CalculaTotalPagadoACuentaPorCategoria = cTotal
End Function
Function CalculaTotalPorCobrarPorCategoria(oRowParent As SSRow)
Dim oRow As SSRow

    Dim cTotal As Currency
    Set oRow = oRowParent.GetChild(ssChildRowFirst)
    cTotal = oRow.Cells("SubTotalPagado").Value
    Do While oRow.HasNextSibling
        Set oRow = oRow.GetSibling(ssSiblingRowNext)
        cTotal = cTotal + oRow.Cells("SubTotalPagado").Value
    Loop
    CalculaTotalPorCobrarPorCategoria = cTotal
End Function


Private Sub grdServicios_AfterCellListCloseUp(ByVal Cell As UltraGrid.SSCell)
Dim oRow As SSRow
Dim oRowParent As SSRow

    If Cell.Column.BaseColumnName = "IdEstadoFacturacion" Then
    
        Set oRow = Cell.Row
        Select Case Cell.GetText
        Case "Pagado"
            Select Case oRow.Cells("IdEstadofacturacion").Value
            Case "1"
                oRow.Cells("SubTotalPagado").Value = oRow.Cells("SubTotalPorPagar").Value
            Case "2"
                oRow.Cells("SubTotalPagado").Value = oRow.Cells("SubTotalPendiente").Value
            Case "7"
                oRow.Cells("SubTotalPendientePago").Value = oRow.Cells("SubTotalPorPagar").Value - oRow.Cells("SubTotalPagadoACuenta").Value
                oRow.Cells("SubTotalPagado").Value = oRow.Cells("SubTotalPendientePago").Value
            Case Else
            End Select
        Case "Pago a cuenta"
            Select Case oRow.Cells("IdEstadofacturacion").Value
            Case "1"
                oRow.Cells("SubTotalPagado").Value = oRow.Cells("SubTotalPagadoACuenta").Value
            Case "7"
                oRow.Cells("SubTotalPendientePago").Value = oRow.Cells("SubTotalPorPagar").Value - oRow.Cells("SubTotalPagadoACuenta").Value
                oRow.Cells("SubTotalPagado").Value = oRow.Cells("SubTotalPagadoACuenta").Value
            Case Else
            End Select
        Case "Emitido"
            oRow.Cells("SubTotalPagado").Value = 0
            oRow.Cells("SubTotalPagadoACuenta").Value = 0
            oRow.Cells("SubTotalPendientePago").Value = 0
        Case "Pendiente Pago"
            oRow.Cells("SubTotalPendientePago").Value = oRow.Cells("SubTotalPorPagar").Value
            oRow.Cells("SubTotalPagado").Value = 0
            oRow.Cells("SubTotalPagadoACuenta").Value = 0
        Case Else
            MsgBox "Ud solo puede modificar el estado si esta en Emitido o Pendiente de Pago", vbInformation, Me.Caption
            Me.grdServicios.PerformAction ssKeyActionUndoCell
            Exit Sub
        End Select
    
        oRow.Cells("IdEmpleadoModifica").Value = ml_IdUsuario
        oRow.Cells("FechaModificacion").Value = Format(Now, "dd/MM/yyyy hh:mm:ss")
    
        oRow.Cells("SubTotalPendientePago").Refresh
        oRow.Cells("IdEmpleadoModifica").Refresh
        oRow.Cells("FechaModificacion").Refresh
        Me.grdServicios.PerformAction ssKeyActionExitEditMode
        
        Dim oFirstRow As SSRow
        Dim cTotal As Currency
        Dim cTotalCategoria As Currency
        
        Set oRowParent = oRow.GetParent()
        
        Set oFirstRow = oRowParent.GetSibling(ssSiblingRowFirst)
        Set oRow = oFirstRow
        cTotal = 0
        cTotalCategoria = 0
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            cTotalCategoria = CalculaTotalPendientePorCategoria(oRow)
            cTotal = cTotal + cTotalCategoria
            oRow.Cells("SubTotalPendientePagoAux").Value = cTotalCategoria
            oRow.Cells("SubTotalPendientePagoAux").Refresh
        Loop
        oFirstRow.Cells("SubTotalPendientePagoAux").Value = cTotal
        oFirstRow.Cells("SubTotalPendientePagoAux").Refresh
    
        'Calcula total por pagar
'        Set oFirstRow = oRowParent.GetSibling(ssSiblingRowFirst)
'        Set oRow = oFirstRow
'        cTotal = 0
'        cTotalCategoria = 0
'        Do While oRow.HasNextSibling
'            Set oRow = oRow.GetSibling(ssSiblingRowNext)
'            cTotalCategoria = CalculaTotalPorPagarPorCategoria(oRow)
'            cTotal = cTotal + cTotalCategoria
'            oRow.Cells("SubTotalPorPagar").Value = cTotalCategoria
'            oRow.Cells("SubTotalPorPagar").Refresh
'        Loop
'        oFirstRow.Cells("SubTotalPorPagar").Value = cTotal
'        oFirstRow.Cells("SubTotalPorPagar").Refresh
    
        'Calcula total pagado a cuenta
        Set oFirstRow = oRowParent.GetSibling(ssSiblingRowFirst)
        Set oRow = oFirstRow
        cTotal = 0
        cTotalCategoria = 0
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            cTotalCategoria = CalculaTotalPagadoACuentaPorCategoria(oRow)
            cTotal = cTotal + cTotalCategoria
            oRow.Cells("SubTotalPagadoACuentaAux").Value = cTotalCategoria
            oRow.Cells("SubTotalPagadoACuentaAux").Refresh
        Loop
        oFirstRow.Cells("SubTotalPagadoACuentaAux").Value = cTotal
        oFirstRow.Cells("SubTotalPagadoACuentaAux").Refresh
    
        'Calcula total pagado
        Set oFirstRow = oRowParent.GetSibling(ssSiblingRowFirst)
        Set oRow = oFirstRow
        cTotal = 0
        cTotalCategoria = 0
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            cTotalCategoria = CalculaTotalPorCobrarPorCategoria(oRow)
            cTotal = cTotal + cTotalCategoria
            oRow.Cells("SubTotalPagadoAux").Value = cTotalCategoria
            oRow.Cells("SubTotalPagadoAux").Refresh
        Loop
        oFirstRow.Cells("SubTotalPagadoAux").Value = cTotal
        oFirstRow.Cells("SubTotalPagadoAux").Refresh
        
        Me.lblMontoFacturadoSoles = cTotal
    End If

End Sub

Private Sub grdBienes_AfterCellListCloseUp(ByVal Cell As UltraGrid.SSCell)
Dim oRow As SSRow
Dim oRowParent As SSRow

    If Cell.Column.BaseColumnName = "IdEstadoFacturacion" Then
    
        Set oRow = Cell.Row
        Select Case Cell.GetText
        Case "Emitido"
            oRow.Cells("SubTotalPendientePago").Value = 0
            oRow.Cells("IdEmpleadoAutorizaPendiente").Value = Null
            oRow.Cells("FechaAutorizaPendiente").Value = Null
        Case "Pendiente Pago"
            oRow.Cells("SubTotalPendientePago").Value = oRow.Cells("SubTotalPorPagar").Value
            oRow.Cells("IdEmpleadoAutorizaPendiente").Value = ml_IdUsuario
            oRow.Cells("FechaAutorizaPendiente").Value = Format(Now, "dd/MM/yyyy hh:mm:ss")
        Case Else
            MsgBox "Ud solo puede modificar el estado si esta en Emitido o Pendiente de Pago", vbInformation, Me.Caption
            Me.grdBienes.PerformAction ssKeyActionUndoCell
            Exit Sub
        End Select
    
        oRow.Cells("IdEmpleadoModifica").Value = ml_IdUsuario
        oRow.Cells("FechaModificacion").Value = Format(Now, "dd/MM/yyyy hh:mm:ss")
    
        oRow.Cells("SubTotalPendientePago").Refresh
        oRow.Cells("IdEmpleadoAutorizaPendiente").Refresh
        oRow.Cells("FechaAutorizaPendiente").Refresh
        Me.grdBienes.PerformAction ssKeyActionExitEditMode
        
        Dim oFirstRow As SSRow
        Set oRowParent = oRow.GetParent()
        Set oFirstRow = oRowParent.GetSibling(ssSiblingRowFirst)
        Set oRow = oFirstRow
        Dim cTotal As Currency
        Dim cTotalCategoria As Currency
        cTotal = 0
        cTotalCategoria = 0
        
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            cTotalCategoria = CalculaTotalPendientePorCategoria(oRow)
            cTotal = cTotal + cTotalCategoria
            
            oRow.Cells("SubTotalPendientePagoAux").Value = cTotalCategoria
            oRow.Cells("SubTotalPendientePagoAux").Refresh
        Loop
        
        oFirstRow.Cells("SubTotalPendientePagoAux").Value = cTotal
        oFirstRow.Cells("SubTotalPendientePagoAux").Refresh
    End If

End Sub

Private Sub grdServicios_BeforeCellListDropDown(ByVal Cell As UltraGrid.SSCell, ByVal Cancel As UltraGrid.SSReturnBoolean)
Dim oRow As SSRow

    Set oRow = Cell.Row
    Select Case oRow.Cells("IdEstadoFacturacion").Value
    Case 1, 3
    Case Else
'        MsgBox "Ud solo puede modificar el estado si esta en Emitido o Pendiente de Pago ", vbInformation, Me.Caption
'        Cancel = True
    End Select


End Sub
Private Sub grdBienes_BeforeCellListDropDown(ByVal Cell As UltraGrid.SSCell, ByVal Cancel As UltraGrid.SSReturnBoolean)
Dim oRow As SSRow

    Set oRow = Cell.Row
    Select Case oRow.Cells("IdEstadoFacturacion").Value
    Case 1, 3
    Case Else
        MsgBox "Ud solo puede modificar el estado si esta en Emitido o Pendiente de Pago ", vbInformation, Me.Caption
        Cancel = True
    End Select


End Sub

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)

    Layout.ViewStyleBand = ssViewStyleBandVertical
    Layout.Override.ExpandRowsOnLoad = ssExpandOnLoadNo
    Layout.Override.FetchRows = ssFetchRowsPreloadWithParent
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti


    With Layout.Override
        .ExpandRowsOnLoad = ssExpandOnLoadNo
        .CellClickAction = ssClickActionEdit
        '.RowSelectors = ssRowSelectorsOff
        .CellSpacing = 60 '75
        .CellPadding = 30 '45
        .RowAppearance.BackColor = &H44F4F9 '&HCDEBFF
        .CellAppearance.BackColor = vbWhite
        .BorderStyleCell = ssBorderStyleNone
        .BorderStyleRow = ssBorderStyleNone
        
        .RowAppearance.AlphaLevel = 192
        .RowAppearance.BackColorAlpha = ssAlphaUseAlphaLevel
        .CellAppearance.AlphaLevel = 192
        .CellAppearance.BackColorAlpha = ssAlphaUseAlphaLevel
        
        .ActiveRowAppearance.BackColorAlpha = ssAlphaOpaque
        .ActiveCellAppearance.BackColorAlpha = ssAlphaOpaque
        
    End With
    
    InitializeServicios
    
End Sub

Sub InitializeServicios()
    
    'Banda 0
    Me.grdServicios.Bands(0).Override.HeaderAppearance.Font.Name = "Tahoma"
    Me.grdServicios.Bands(0).Override.HeaderAppearance.Font.Size = 8
    'Me.grdServicios.Bands(0).Override.HeaderAppearance.Font.Bold = True
    
    Me.grdServicios.Bands(0).Override.RowAppearance.Font.Name = "Tahoma"
    Me.grdServicios.Bands(0).Override.RowAppearance.Font.Size = 8
    Me.grdServicios.Bands(0).Override.RowAppearance.BackColor = &HDEB59E
    
    Me.grdServicios.Bands(0).Columns("IdCategoriaProducto").Hidden = True
    Me.grdServicios.Bands(0).Columns("Descripcion").Width = 4000
    
    Me.grdServicios.Bands(0).Columns("SubTotalPorPagar").Hidden = True
    
    Me.grdServicios.Bands(0).Columns.Add "SubTotalPorPagarAux"
    Me.grdServicios.Bands(0).Columns("SubTotalPorPagarAux").Header.Caption = "Por Pagar"
    Me.grdServicios.Bands(0).Columns("SubTotalPorPagarAux").Width = 1000
    Me.grdServicios.Bands(0).Columns("SubTotalPorPagarAux").CellAppearance.TextAlign = ssAlignRight
    Me.grdServicios.Bands(0).Columns("SubTotalPorPagarAux").Activation = ssActivationActivateOnly
    
    Me.grdServicios.Bands(0).Columns("SubTotalPagadoACuenta").Hidden = True
    
    Me.grdServicios.Bands(0).Columns.Add "SubTotalPagadoACuentaAux"
    Me.grdServicios.Bands(0).Columns("SubTotalPagadoACuentaAux").Header.Caption = "A Cuenta"
    Me.grdServicios.Bands(0).Columns("SubTotalPagadoACuentaAux").Width = 1000
    Me.grdServicios.Bands(0).Columns("SubTotalPagadoACuentaAux").CellAppearance.TextAlign = ssAlignRight
    Me.grdServicios.Bands(0).Columns("SubTotalPagadoACuentaAux").Activation = ssActivationActivateOnly
    
    Me.grdServicios.Bands(0).Columns("SubTotalPendientePago").Hidden = True
    
    Me.grdServicios.Bands(0).Columns.Add "SubTotalPendientePagoAux"
    Me.grdServicios.Bands(0).Columns("SubTotalPendientePagoAux").Header.Caption = "Pendiente"
    Me.grdServicios.Bands(0).Columns("SubTotalPendientePagoAux").Width = 1000
    Me.grdServicios.Bands(0).Columns("SubTotalPendientePagoAux").CellAppearance.TextAlign = ssAlignRight
    Me.grdServicios.Bands(0).Columns("SubTotalPendientePagoAux").Activation = ssActivationActivateOnly
    
    Me.grdServicios.Bands(0).Columns.Add "SubTotalIGV"
    Me.grdServicios.Bands(0).Columns("SubTotalIGV").Header.Caption = "IGV"
    Me.grdServicios.Bands(0).Columns("SubTotalIGV").Width = 1000
    Me.grdServicios.Bands(0).Columns("SubTotalIGV").CellAppearance.TextAlign = ssAlignRight
    Me.grdServicios.Bands(0).Columns("SubTotalIGV").Activation = ssActivationActivateOnly
    
    Me.grdServicios.Bands(0).Columns("SubTotalPagado").Hidden = True
    
    Me.grdServicios.Bands(0).Columns.Add "SubTotalPagadoAux"
    Me.grdServicios.Bands(0).Columns("SubTotalPagadoAux").Header.Caption = "Por Cobrar"
    Me.grdServicios.Bands(0).Columns("SubTotalPagadoAux").Width = 1000
    Me.grdServicios.Bands(0).Columns("SubTotalPagadoAux").CellAppearance.TextAlign = ssAlignRight
    Me.grdServicios.Bands(0).Columns("SubTotalPagadoAux").Activation = ssActivationActivateOnly
    
    'Banda 1
    Me.grdServicios.Bands(1).Override.HeaderAppearance.Font.Name = "Tahoma"
    Me.grdServicios.Bands(1).Override.HeaderAppearance.Font.Size = 8
    'Me.grdServicios.Bands(1).Override.HeaderAppearance.Font.Bold = True
    
    Me.grdServicios.Bands(1).Override.RowAppearance.Font.Name = "Tahoma"
    Me.grdServicios.Bands(1).Override.RowAppearance.Font.Size = 8

    Me.grdServicios.Bands(1).Columns("IdCategoriaProducto").Hidden = True
    Me.grdServicios.Bands(1).Columns("IdProducto").Hidden = True
    Me.grdServicios.Bands(1).Columns("IdFacturacionServicio").Hidden = True
    
    Me.grdServicios.Bands(1).Columns("Nombre").Width = 4000
    Me.grdServicios.Bands(1).Columns("Nombre").DisplayEllipses = ssDisplayEllipsesYes
    Me.grdServicios.Bands(1).Columns("Nombre").Activation = ssActivationActivateOnly
    
    Me.grdServicios.Bands(1).Columns("PrecioUnitario").Header.Caption = "P.U."
    Me.grdServicios.Bands(1).Columns("PrecioUnitario").Width = 250
    Me.grdServicios.Bands(1).Columns("PrecioUnitario").Activation = ssActivationActivateNoEdit
    
    Me.grdServicios.Bands(1).Columns("Cantidad").Header.Caption = "Cant."
    Me.grdServicios.Bands(1).Columns("Cantidad").Width = 250
    Me.grdServicios.Bands(1).Columns("Cantidad").Activation = ssActivationActivateNoEdit
    

    Me.grdServicios.Bands(1).Columns("SubTotalPorPagar").Header.Caption = "Por Pagar"
    Me.grdServicios.Bands(1).Columns("SubTotalPorPagar").Width = 1000
    Me.grdServicios.Bands(1).Columns("SubTotalPorPagar").Activation = ssActivationActivateOnly
    
    Me.grdServicios.Bands(1).Columns("SubTotalPendientePago").Header.Caption = "Pendiente"
    Me.grdServicios.Bands(1).Columns("SubTotalPendientePago").Width = 1000
    Me.grdServicios.Bands(1).Columns("SubTotalPendientePago").Activation = ssActivationActivateOnly
    
    Me.grdServicios.Bands(1).Columns("SubTotalPagadoACuenta").Header.Caption = "A Cuenta"
    Me.grdServicios.Bands(1).Columns("SubTotalPagadoACuenta").Width = 1000
    Me.grdServicios.Bands(1).Columns("SubTotalPagadoACuenta").Activation = ssActivationAllowEdit
    
    Me.grdServicios.Bands(1).Columns("SubTotalPagado").Header.Caption = "Por Cobrar"
    Me.grdServicios.Bands(1).Columns("SubTotalPagado").Width = 1000
    Me.grdServicios.Bands(1).Columns("SubTotalPagado").Activation = ssActivationActivateOnly
    
    Me.grdServicios.Bands(1).Columns("IdEstadoFacturacion").Header.Caption = "Estado"
    Me.grdServicios.Bands(1).Columns("IdEstadoFacturacion").Width = 1800
        
    Me.grdServicios.Bands(1).Columns("IdEmpleadoModifica").Header.Caption = "Resp."
    Me.grdServicios.Bands(1).Columns("IdEmpleadoModifica").Width = 3000
    Me.grdServicios.Bands(1).Columns("IdEmpleadoModifica").Activation = ssActivationActivateOnly
    
    Me.grdServicios.Bands(1).Columns("FechaModificacion").Header.Caption = "Fecha"
    Me.grdServicios.Bands(1).Columns("FechaModificacion").Width = 2500
    Me.grdServicios.Bands(1).Columns("FechaModificacion").Activation = ssActivationActivateOnly
    Me.grdServicios.Bands(1).Columns("FechaModificacion").Format = "dd/MM/yyyy hh:mm:ss"

End Sub

Private Sub grdBienes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)

    Layout.ViewStyleBand = ssViewStyleBandVertical
    Layout.Override.ExpandRowsOnLoad = ssExpandOnLoadNo
    Layout.Override.FetchRows = ssFetchRowsPreloadWithParent
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti


    With Layout.Override
        .ExpandRowsOnLoad = ssExpandOnLoadNo
        .CellClickAction = ssClickActionEdit
        '.RowSelectors = ssRowSelectorsOff
        .CellSpacing = 75
        .CellPadding = 45
        .RowAppearance.BackColor = &H44F4F9 '&HCDEBFF
        .CellAppearance.BackColor = vbWhite
        .BorderStyleCell = ssBorderStyleNone
        .BorderStyleRow = ssBorderStyleNone
        
        .RowAppearance.AlphaLevel = 192
        .RowAppearance.BackColorAlpha = ssAlphaUseAlphaLevel
        .CellAppearance.AlphaLevel = 192
        .CellAppearance.BackColorAlpha = ssAlphaUseAlphaLevel
        
        .ActiveRowAppearance.BackColorAlpha = ssAlphaOpaque
        .ActiveCellAppearance.BackColorAlpha = ssAlphaOpaque
        
    End With
    
    InitializeServiciosBienes
    
End Sub

Sub InitializeServiciosBienes()
    
    'Banda 0
    Me.grdBienes.Bands(0).Override.HeaderAppearance.Font.Name = "Tahoma"
    Me.grdBienes.Bands(0).Override.HeaderAppearance.Font.Size = 10
    Me.grdBienes.Bands(0).Override.HeaderAppearance.Font.Bold = True
    
    Me.grdBienes.Bands(0).Override.RowAppearance.Font.Name = "Tahoma"
    Me.grdBienes.Bands(0).Override.RowAppearance.Font.Size = 10
    Me.grdBienes.Bands(0).Override.RowAppearance.BackColor = &HDEB59E
    
    Me.grdBienes.Bands(0).Columns("IdTipoBienInsumo").Hidden = True
    Me.grdBienes.Bands(0).Columns("Descripcion").Width = 5000
    
    Me.grdBienes.Bands(0).Columns("SubTotalPorPagar").Header.Caption = "Por Pagar S/."
    Me.grdBienes.Bands(0).Columns("SubTotalPorPagar").Width = 1500
    Me.grdBienes.Bands(0).Columns("SubTotalPorPagar").Activation = ssActivationActivateOnly
    
    Me.grdBienes.Bands(0).Columns("SubTotalPendientePago").Hidden = True
    
    Me.grdBienes.Bands(0).Columns.Add "SubTotalPendientePagoAux"
    Me.grdBienes.Bands(0).Columns("SubTotalPendientePagoAux").Header.Caption = "Pendiente S/."
    Me.grdBienes.Bands(0).Columns("SubTotalPendientePagoAux").Width = 1500
    Me.grdBienes.Bands(0).Columns("SubTotalPendientePagoAux").CellAppearance.TextAlign = ssAlignRight
    Me.grdBienes.Bands(0).Columns("SubTotalPendientePagoAux").Activation = ssActivationActivateOnly
    
    'Banda 1
    Me.grdBienes.Bands(1).Override.HeaderAppearance.Font.Name = "Tahoma"
    Me.grdBienes.Bands(1).Override.HeaderAppearance.Font.Size = 10
    Me.grdBienes.Bands(1).Override.HeaderAppearance.Font.Bold = True
    
    Me.grdBienes.Bands(1).Override.RowAppearance.Font.Name = "Tahoma"
    Me.grdBienes.Bands(1).Override.RowAppearance.Font.Size = 10

    Me.grdBienes.Bands(1).Columns("IdTipoBienInsumo").Hidden = True
    Me.grdBienes.Bands(1).Columns("Nombre").Width = 5000
    Me.grdBienes.Bands(1).Columns("Nombre").DisplayEllipses = ssDisplayEllipsesYes
    Me.grdBienes.Bands(1).Columns("Nombre").Activation = ssActivationActivateOnly
    
    Me.grdBienes.Bands(1).Columns("PrecioUnitario").Header.Caption = "P.U. S/."
    Me.grdBienes.Bands(1).Columns("PrecioUnitario").Width = 1000
    Me.grdBienes.Bands(1).Columns("PrecioUnitario").Activation = ssActivationActivateNoEdit
    
    Me.grdBienes.Bands(1).Columns("Cantidad").Header.Caption = "Cantidad S/."
    Me.grdBienes.Bands(1).Columns("Cantidad").Width = 1000
    Me.grdBienes.Bands(1).Columns("Cantidad").Activation = ssActivationActivateNoEdit
    

    Me.grdBienes.Bands(1).Columns("SubTotalPorPagar").Header.Caption = "Por Pagar S/."
    Me.grdBienes.Bands(1).Columns("SubTotalPorPagar").Width = 1500
    Me.grdBienes.Bands(1).Columns("SubTotalPorPagar").Activation = ssActivationActivateOnly
    
    Me.grdBienes.Bands(1).Columns("SubTotalPendientePago").Header.Caption = "Pendiente S/."
    Me.grdBienes.Bands(1).Columns("SubTotalPendientePago").Width = 1500
    Me.grdBienes.Bands(1).Columns("SubTotalPendientePago").Activation = ssActivationActivateOnly
    
    Me.grdBienes.Bands(1).Columns("IdEstadoFacturacion").Header.Caption = "Estado"
    Me.grdBienes.Bands(1).Columns("IdEstadoFacturacion").Width = 1800
        
    Me.grdBienes.Bands(1).Columns("IdEmpleadoAutorizaPendiente").Header.Caption = "Resp.de Pend."
    Me.grdBienes.Bands(1).Columns("IdEmpleadoAutorizaPendiente").Width = 3000
    Me.grdBienes.Bands(1).Columns("IdEmpleadoAutorizaPendiente").Activation = ssActivationActivateOnly
    
    Me.grdBienes.Bands(1).Columns("FechaAutorizaPendiente").Header.Caption = "Fecha Pendiente"
    Me.grdBienes.Bands(1).Columns("FechaAutorizaPendiente").Width = 2500
    Me.grdBienes.Bands(1).Columns("FechaAutorizaPendiente").Activation = ssActivationActivateOnly
    Me.grdBienes.Bands(1).Columns("FechaAutorizaPendiente").Format = "dd/MM/yyyy hh:mm:ss"

    Me.grdBienes.Bands(1).Columns("IdEmpleadoModifica").Hidden = True
    Me.grdBienes.Bands(1).Columns("FechaModificacion").Hidden = True

End Sub


Private Sub grdServicios_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)

    If Row.Band.Index = 0 Then
            Row.Cells("SubTotalPorPagarAux").Value = Row.Cells("SubTotalPorPagar").Value
            Row.Cells("SubTotalPendientePagoAux").Value = Row.Cells("SubTotalPendientePago").Value
            Row.Cells("SubTotalPagadoACuentaAux").Value = Row.Cells("SubTotalPagadoACuenta").Value
            Row.Cells("SubTotalPagadoAux").Value = Row.Cells("SubTotalPagado").Value
    Else
        'Row.Cells("SubTotalPendientePago").Appearance.Font.Size = 11
        'Row.Cells("SubTotalPendientePago").Appearance.Font.Bold = True
        'Row.Cells("SubTotalPendientePago").Appearance.ForeColor = RGB(255, 0, 0)
        
        If Row.Cells("IdEstadoFacturacion").Value = 6 Then
            'Row.Cells("IdEstadoFacturacion").Activation = ssActivationDisabled
        End If
        
    End If
    
    
    
End Sub
Private Sub grdBienes_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)

    If Row.Band.Index = 0 Then
        Row.Cells("SubTotalPendientePagoAux").Value = Row.Cells("SubTotalPendientePago").Value
    Else
        Row.Cells("SubTotalPendientePago").Appearance.Font.Size = 11
        Row.Cells("SubTotalPendientePago").Appearance.Font.Bold = True
        Row.Cells("SubTotalPendientePago").Appearance.ForeColor = RGB(255, 0, 0)
        
        If Row.Cells("IdEstadoFacturacion").Value = 6 Then
            Row.Cells("IdEstadoFacturacion").Activation = ssActivationDisabled
        End If
        
    End If
    
    
    
End Sub


Sub ObtenerDatosDePaciente()
Dim rsPaciente  As New Recordset

    Screen.MousePointer = vbHourglass
    ml_IdCuentaAtencion = Val(Me.txtIdCuentaAtencion.Text)
    Set rsPaciente = mo_AdminAdmision.CuentasAtencionDatosPacientePorIdCuentaAtencion(ml_IdCuentaAtencion)
    Screen.MousePointer = vbDefault
    
    'Si hay una sola coincidencia
    If rsPaciente.RecordCount = 1 Then
        rsPaciente.MoveFirst
        Me.txtIdNroHistoria.Text = rsPaciente!NroHistoriaClinica
        mo_cmbIdTipoGenHistoriaClinica.BoundText = rsPaciente!IdTipoNumeracion
        Me.txtRazonSocial = rsPaciente!ApellidoPaterno + " " + rsPaciente!ApellidoMaterno + " " + rsPaciente!PrimerNombre + " " + ("" & rsPaciente!SegundoNombre)
    ElseIf rsPaciente.RecordCount = 0 Then
        MsgBox "No se encontraron atenciones para el nro de cuenta ingresado", vbInformation, Me.Caption
    End If

End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica
    Set mo_cmbIdTipoComprobante.MiComboBox = cmbIdTipoComprobante
End Sub

Sub GenerarRecordsetTemporal()
'    Set mrs_ComprobantesDetalle = New Recordset
'    With mrs_ComprobantesDetalle
'          '.Fields.Append "CheckSeleccionado", adBoolean, 4, adFldIsNullable
'          .Fields.Append "TipoDetalle", adVarChar, 4, adFldIsNullable
'          .Fields.Append "IdFacturacionDetalle", adInteger, 4, adFldIsNullable
'          .Fields.Append "IdComprobanteDetalle", adInteger, 4, adFldIsNullable
'          .Fields.Append "IdProducto", adInteger, 4, adFldIsNullable
'          .Fields.Append "CodigoProducto", adVarChar, 20, adFldIsNullable
'          .Fields.Append "Producto", adVarChar, 200, adFldIsNullable
'          .Fields.Append "Cantidad", adCurrency, 8, adFldIsNullable
'          .Fields.Append "PrecioUnitario", adCurrency, 8, adFldIsNullable
'          .Fields.Append "SubTotalExonerado", adCurrency, 8, adFldIsNullable
'          .Fields.Append "SubTotalPendientePago", adCurrency, 8, adFldIsNullable
'
'          .LockType = adLockOptimistic
'          .Open
'    End With
'
'    Set Me.grdServicios.DataSource = mrs_ComprobantesDetalle
    
    Set mrs_FormaPago = New Recordset
    With mrs_FormaPago
          .Fields.Append "IdFormaPago", adInteger, 4, adFldIsNullable
          .Fields.Append "IdTipoFormaPago", adInteger, 4, adFldIsNullable
          .Fields.Append "Importe", adCurrency, 8, adFldIsNullable
          .Fields.Append "IdTipoMoneda", adInteger, 4, adFldIsNullable
          
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdDinero.DataSource = mrs_FormaPago
End Sub

Sub CargarDatosALosControles()
Dim oDOCajaComprobantesPago  As DOCajaComprobantesPago

       Set oDOCajaComprobantesPago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(Me.IdComprobantePago)
       
       If mo_AdminCaja.MensajeError <> "" Then
            MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminCaja.MensajeError, vbCritical, Me.Caption
            mb_ExistenDatos = False
            Exit Sub
       End If
       
       If Not oDOCajaComprobantesPago Is Nothing Then
            With oDOCajaComprobantesPago
                mo_cmbIdTipoComprobante.BoundText = .IdTipoComprobante
                txtNroSerie = .NroSerie
                txtNroDocumento = .NroDocumento
                txtRazonSocial = .RazonSocial
                txtIdCuentaAtencion = IIf(.IdCuentaAtencion = 0, "", CStr(.IdCuentaAtencion))
                txtRUC = .RUC
                'lblSubTotal = Format(.SubTotal, "0.00")
                'lblIGV = Format(.IGV, "0.00")
                'lblTotal = Format(.Total, "0.00")
                
                Set mo_CajaComprobantesPago = oDOCajaComprobantesPago
                mb_ExistenDatos = True
            End With
            '-------------------------------------
            'Cargamos del Items de la Factura
            '-------------------------------------
            Dim rsDetalle As New Recordset
            Dim oDOCajaDetalle As New DOCajaComprobantesDetalle
            oDOCajaDetalle.IdComprobantePago = oDOCajaComprobantesPago.IdComprobantePago
            Set rsDetalle = mo_AdminCaja.CajaComprobantesDetalle(oDOCajaDetalle)
            Do While Not rsDetalle.EOF
                With mrs_ComprobantesDetalle
                    .AddNew
                    .Fields!CheckSeleccionado = True
                    .Fields!IdFacturacionDetalle = rsDetalle!IdFacturacionDetalle
                    .Fields!TipoDetalle = rsDetalle!TipoDetalle
                    .Fields!CodigoProducto = rsDetalle!Codigo
                    .Fields!IdProducto = rsDetalle!IdProducto
                    .Fields!Producto = rsDetalle!Producto
                    .Fields!cantidad = rsDetalle!cantidad
                    .Fields!precioUnitario = rsDetalle!precioUnitario
                    .Fields!SubTotalExonerado = rsDetalle!SubTotalExonerado
                    .Fields!SubTotalPagado = rsDetalle!SubTotalPagado
                
                End With
                rsDetalle.MoveNext
            Loop
            rsDetalle.Close
            mo_Apariencia.ConfigurarFilasBiColores Me.grdServicios, SIGHComun.GrillaConFilasBicolor
            '-------------------------------------
            
            '-------------------------------------
            'Cargamos las Formas de Pago de la Factura
            '-------------------------------------
            Dim rsFormaPago As New Recordset
            Dim oDOCajaFormaPago As New DOCajaFormaPagoComprobante
            oDOCajaFormaPago.IdComprobantePago = oDOCajaComprobantesPago.IdComprobantePago
            Set rsFormaPago = mo_AdminCaja.CajaFormaPagoComprobante(oDOCajaFormaPago)
            Do While Not rsFormaPago.EOF
                With mrs_FormaPago
                    .AddNew
                    .Fields!IdFormaPago = rsFormaPago!IdFormaPago
                    .Fields!IdTipoFormaPago = rsFormaPago!IdTipoFormaPago
                    .Fields!Importe = rsFormaPago!Importe
                    .Fields!IdTipoMoneda = rsFormaPago!IdTipoMoneda
                
                End With
                rsFormaPago.MoveNext
            Loop
            rsFormaPago.Close
            mo_Apariencia.ConfigurarFilasBiColores Me.grdDinero, SIGHComun.GrillaConFilasBicolor
            '-------------------------------------
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
       
    CalcularSubTotalesItems
    CalcularSubTotalesDinero
   
    cmdImprimir.Enabled = False
    cmdNuevo.Enabled = False

    'fraDinero.Enabled = False
    Select Case mi_Opcion
        Case sghAgregar
            cmdNuevo.Enabled = True
            cmdImprimir.Enabled = True
            
            'fraDatosGenerales.Enabled = True
            'fraItems.Enabled = True
            'fraDinero.Enabled = True
        Case sghModificar
            cmdImprimir.Enabled = True
        
            'fraDatosGenerales.Enabled = True
            'fraItems.Enabled = True
            'fraDinero.Enabled = True
        Case sghEliminar
        Case sghConsultar
            Me.cmdGrabar.Enabled = False
    End Select
   
End Sub

Private Sub grdDinero_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    If bCalculandoSubTotales Then Exit Sub
    CalcularSubTotalesDinero
End Sub


Private Sub grdDinero_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdDinero.Bands(0).Columns("IdFormaPago").Hidden = True
    
    Dim rs As New Recordset
    
    Set rs = mo_AdminCaja.CajaTiposFormasPago()
    With grdDinero.ValueLists.Add("TipoFormaPago").ValueListItems
        Do Until rs.EOF
            .Add Trim(Str(rs.Fields!IdTipoFormaPago)), rs.Fields!Descripcion
            rs.MoveNext
        Loop
    End With
    rs.Close
    Set rs = mo_AdminCaja.CajaTiposMoneda
    With grdDinero.ValueLists.Add("TipoMoneda").ValueListItems
        Do Until rs.EOF
            .Add Trim(Str(rs.Fields!IdTipoMoneda)), rs.Fields!Descripcion
            rs.MoveNext
        Loop
    End With
    rs.Close
    grdDinero.Bands(0).Columns("IdTipoFormaPago").Header.Caption = "Forma Pago"
    grdDinero.Bands(0).Columns("IdTipoFormaPago").Width = 1500
    grdDinero.Bands(0).Columns("IdTipoFormaPago").ValueList = "TipoFormaPago"
    grdDinero.Bands(0).Columns("IdTipoFormaPago").ButtonDisplayStyle = ssButtonDisplayStyleOnCellActivate
    
    grdDinero.Bands(0).Columns("IdTipoMoneda").Header.Caption = "Moneda"
    grdDinero.Bands(0).Columns("IdTipoMoneda").Width = 1500
    grdDinero.Bands(0).Columns("IdTipoMoneda").ValueList = "TipoMoneda"
    grdDinero.Bands(0).Columns("IdTipoMoneda").ButtonDisplayStyle = ssButtonDisplayStyleOnCellActivate
        
    grdDinero.Bands(0).Columns("Importe").Header.Caption = "Importe"
    grdDinero.Bands(0).Columns("Importe").Width = 1500

End Sub

Private Sub grdDinero_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode.Value, cmdGrabar
    AdministrarKeyPreview KeyCode.Value
End Sub

Private Sub CalcularSubTotalesItems()
    Dim dSubTotal As Double
    Dim dImpuesto As Double
    Dim dTotal As Double
    
    bCalculandoSubTotales = True
        
    dSubTotal = 0
    
    
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = mrs_ComprobantesDetalle.Clone(adLockReadOnly)
    
    If Not (rsTemp.BOF And rsTemp.EOF) Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If rsTemp.Fields!CheckSeleccionado Then
                dSubTotal = dSubTotal + rsTemp.Fields!SubTotalPendientePago
            End If
            rsTemp.MoveNext
        Loop
        rsTemp.MoveFirst
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    
    bCalculandoSubTotales = False
    'Me.lblSubTotal = Format(dSubTotal, "0.00")
    'Me.lblIGV = Format(dSubTotal * md_PorcentajeIGV, "0.00")
    'Me.lblTotal = Format(dSubTotal * (1 + md_PorcentajeIGV), "0.00")
    'Me.lblMontoFacturadoSoles = Me.lblTotal
    
    CalcularVuelto
End Sub
Private Sub CalcularSubTotalesDinero()
    Dim dSubTotalSoles As Double
    Dim dSubTotalDolar As Double
    dSubTotalSoles = 0
    dSubTotalDolar = 0
    bCalculandoSubTotales = True
    
    
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = mrs_FormaPago.Clone(adLockReadOnly)
    
    If Not (rsTemp.BOF And rsTemp.EOF) Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If rsTemp.Fields!IdTipoMoneda = ID_TIPO_MONEDA_DOLAR Then
                dSubTotalDolar = dSubTotalDolar + IIf(IsNull(rsTemp.Fields!Importe), 0, rsTemp.Fields!Importe)
            Else
                dSubTotalSoles = dSubTotalSoles + IIf(IsNull(rsTemp.Fields!Importe), 0, rsTemp.Fields!Importe)
            End If
            rsTemp.MoveNext
        Loop
        rsTemp.MoveFirst
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    
    bCalculandoSubTotales = False
    lblMontoRecibidoDolares = Format(dSubTotalDolar, "0.00")
    lblMontoRecibidoSoles = Format(dSubTotalSoles, "0.00")
    
    CalcularVuelto
End Sub
Private Sub CalcularVuelto()
    Dim dTotalRecibidoSoles As Currency
    Dim dTotalFacturadoSoles As Currency

    
    dTotalRecibidoSoles = CCurrency(Me.lblMontoRecibidoSoles) + CCurrency(Me.lblMontoRecibidoDolares) * md_TipoCambioDolar
    dTotalFacturadoSoles = CCurrency(Me.lblMontoFacturadoSoles)
    If dTotalFacturadoSoles > dTotalRecibidoSoles Then
        lblMontoFaltanteSoles = Format(dTotalFacturadoSoles - dTotalRecibidoSoles, "0.00")
        lblMontoVueltoSoles = Format(0, "0.00")
    Else
        lblMontoFaltanteSoles = Format(0, "0.00")
        lblMontoVueltoSoles = Format(dTotalRecibidoSoles - dTotalFacturadoSoles, "0.00")
    End If
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           cmdGrabar.Value = True
       End Select
End Sub
Private Function CCurrency(sValor As String) As Currency
    If Trim(sValor) = "" Then
        CCurrency = 0
    Else
        CCurrency = CCur(sValor)
    End If
End Function

Private Sub txtIdCuentaAtencion_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        btnLeerDatos_Click
        Exit Sub
    End If

    mo_Teclado.RealizarNavegacion KeyCode, txtIdCuentaAtencion
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdCuentaAtencion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtRazonSocial_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtRazonSocial
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtRazonSocial_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsParaDireccion(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtRUC_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtRUC
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtRUC_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Function ValidarReglas() As Boolean

    ValidarReglas = False
   
    If mi_Opcion = sghAgregar Then
    
    End If
   
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla CajaComprobantesPago
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

    With mo_CajaComprobantesPago
        .IdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
        .NroSerie = Me.txtNroSerie
        .NroDocumento = Me.txtNroDocumento
        .IdCuentaAtencion = Val(Me.txtIdCuentaAtencion)
        .RazonSocial = Me.txtRazonSocial
        .Observaciones = ""
        .IdGestionCaja = Me.CajaLoteActual.IdLote
        .IdUsuarioAuditoria = ml_IdUsuario
        
        .subtotal = CCurrency(Me.lblMontoFacturadoSoles)
        .IGV = CCurrency(0)
        .Total = CCurrency(Me.lblMontoFacturadoSoles)
    End With
    
    'mo_CajaCajaActual.NroSerie = mo_CajaComprobantesPago.NroSerie
    'mo_CajaCajaActual.NroComprobante = mo_CajaComprobantesPago.NroDocumento
    
    '------------------------------------
    'Cargamos los Items a Facturar
    '------------------------------------
    Set mo_ItemsAFacturar = New Collection
    Dim oItemFactura As DOCajaComprobantesDetalle
    
    Dim oParentRow As SSRow
    
    Set oParentRow = Me.grdServicios.GetRow(ssChildRowFirst)
    CargarDatosDeLosHijos oParentRow
    
    Do While oParentRow.HasNextSibling
        'Obtiene el siguiente registro de la banda 0
        Set oParentRow = oParentRow.GetSibling(ssSiblingRowNext)
        CargarDatosDeLosHijos oParentRow
    Loop
        
'    If Not (mrs_ComprobantesDetalle.BOF And mrs_ComprobantesDetalle.EOF) Then
'        mrs_ComprobantesDetalle.MoveFirst
'        Do While Not mrs_ComprobantesDetalle.EOF
'            If mrs_ComprobantesDetalle!CheckSeleccionado Then
'                Set oItemFactura = New DOCajaComprobantesDetalle
'                oItemFactura.IdFacturacionDetalle = IIf(IsNull(mrs_ComprobantesDetalle!IdFacturacionDetalle), 0, mrs_ComprobantesDetalle!IdFacturacionDetalle)
'                oItemFactura.IdProducto = mrs_ComprobantesDetalle!IdProducto
'                oItemFactura.TipoDetalle = mrs_ComprobantesDetalle!TipoDetalle
'                oItemFactura.Cantidad = mrs_ComprobantesDetalle!Cantidad
'                oItemFactura.PrecioUnitario = mrs_ComprobantesDetalle!PrecioUnitario
'                oItemFactura.SubTotalPagado = mrs_ComprobantesDetalle!SubTotalPagado
'                oItemFactura.IdUsuarioAuditoria = ml_IdUsuario
'
'                mo_ItemsAFacturar.Add oItemFactura
'            End If
'            mrs_ComprobantesDetalle.MoveNext
'        Loop
'        mrs_ComprobantesDetalle.MoveFirst
'    End If
    
    '------------------------------------
    'Cargamos los Items de Dinero
    '------------------------------------
    Set mo_ItemsDinero = New Collection
    Dim oItemDinero  As DOCajaFormaPagoComprobante
    
    If Not (mrs_FormaPago.BOF And mrs_FormaPago.EOF) Then
        mrs_FormaPago.MoveFirst
        Do While Not mrs_FormaPago.EOF
            Set oItemDinero = New DOCajaFormaPagoComprobante
            oItemDinero.IdTipoFormaPago = mrs_FormaPago!IdTipoFormaPago
            oItemDinero.IdTipoMoneda = mrs_FormaPago!IdTipoMoneda
            oItemDinero.Importe = mrs_FormaPago!Importe
            oItemDinero.IdUsuarioAuditoria = ml_IdUsuario
            If oItemDinero.IdTipoMoneda = ID_TIPO_MONEDA_DOLAR Then
                oItemDinero.TipoCambio = md_TipoCambioDolar
            Else
                oItemDinero.TipoCambio = 1
            End If
            oItemDinero.TotalSoles = IIf(IsNull(mrs_FormaPago!Importe), 0, mrs_FormaPago!Importe) * oItemDinero.TipoCambio
            
            mo_ItemsDinero.Add oItemDinero
            mrs_FormaPago.MoveNext
        Loop
    End If
    
End Sub
Sub CargarDatosDeLosHijos(oParentRow As SSRow)
Dim oChildRow As SSRow
    
    If oParentRow.HasChild Then
        Set oChildRow = oParentRow.GetChild(ssChildRowFirst)
        If oChildRow.Cells("IdEstadoFacturacion").Value = 6 Or oChildRow.Cells("IdEstadoFacturacion").Value = 7 Then
            CargarDatosDeHijo oChildRow
        End If
        Do While oChildRow.HasNextSibling
            Set oChildRow = oChildRow.GetSibling(ssSiblingRowNext)
            If oChildRow.Cells("IdEstadoFacturacion").Value = 6 Or oChildRow.Cells("IdEstadoFacturacion").Value = 7 Then
                CargarDatosDeHijo oChildRow
            End If
        Loop
    End If

End Sub
Sub CargarDatosDeHijo(oChildRow As SSRow)
Dim oItemFactura As DOCajaComprobantesDetalle

    Set oItemFactura = New DOCajaComprobantesDetalle
    oItemFactura.IdFacturacionDetalle = IIf(IsNull(oChildRow.Cells("IdFacturacionServicio").Value), 0, oChildRow.Cells("IdFacturacionServicio").Value)
    oItemFactura.IdProducto = oChildRow.Cells("IdProducto").Value
    oItemFactura.TipoDetalle = SIGHComun.sghDetalleComprobanteServicios
    oItemFactura.cantidad = oChildRow.Cells("Cantidad").Value
    oItemFactura.precioUnitario = oChildRow.Cells("PrecioUnitario").Value
    oItemFactura.SubTotalPagado = oChildRow.Cells("SubTotalPagado").Value
    oItemFactura.IdEstadoFacturacion = oChildRow.Cells("IdEstadoFacturacion").Value
    oItemFactura.IdUsuarioAuditoria = ml_IdUsuario

    mo_ItemsAFacturar.Add oItemFactura
End Sub
Sub LLenarComprobanteDetalle(oRow As SSRow)
'Dim oItemFactura As DOCajaComprobantesDetalle
'
'        If oRow("IdEstadoFacturacion") = 6 Or oRow("IdEstadoFacturacion") = 5 Then
'            Set oItemFactura = New DOCajaComprobantesDetalle
'            oItemFactura.IdFacturacionDetalle = IIf(IsNull(mrs_ComprobantesDetalle!IdFacturacionDetalle), 0, mrs_ComprobantesDetalle!IdFacturacionDetalle)
'            oItemFactura.IdProducto = mrs_ComprobantesDetalle!IdProducto
'            oItemFactura.TipoDetalle = mrs_ComprobantesDetalle!TipoDetalle
'            oItemFactura.Cantidad = mrs_ComprobantesDetalle!Cantidad
'            oItemFactura.PrecioUnitario = mrs_ComprobantesDetalle!PrecioUnitario
'            oItemFactura.SubTotalPagado = mrs_ComprobantesDetalle!SubTotalPagado
'            oItemFactura.IdUsuarioAuditoria = ml_IdUsuario
'
'            mo_ItemsAFacturar.Add oItemFactura
'        End If
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    AgregarDatos = mo_AdminCaja.ComprobantePagoAgregar(mo_CajaComprobantesPago, mo_ItemsAFacturar, mo_ItemsDinero, mo_CajaCajaActual)
    
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    ModificarDatos = mo_AdminCaja.ComprobantePagoModificar(mo_CajaComprobantesPago, mo_ItemsAFacturar, mo_ItemsDinero)
       
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    EliminarDatos = mo_AdminCaja.ComprobantePagoEliminar(mo_CajaComprobantesPago)
End Function

Function ValidarDatosObligatorios() As Boolean
Dim sMensaje As String
Dim bFound As Boolean

    ValidarDatosObligatorios = False
    
    If Me.txtNroSerie.Text = "" Then
        sMensaje = sMensaje + "Ingrese el Nº de Serie" + Chr(13)
    End If
    If Me.txtNroDocumento.Text = "" Then
        sMensaje = sMensaje + "Ingrese el Nº de Documento" + Chr(13)
    End If
    If mo_cmbIdTipoComprobante.BoundText = "" Then
        sMensaje = sMensaje + "Ingrese el tipo de Comprobante" + Chr(13)
    End If
    If mo_cmbIdTipoComprobante.BoundText = ID_TIPO_COMPROBANTE_FACTURA Then
        If Me.txtRUC = "" Then
            sMensaje = sMensaje + "Ingrese el RUC para la Factura" + Chr(13)
        End If
    Else
         Me.txtRUC = ""
    End If
    If Trim(Me.txtRazonSocial.Text) = "" Then
        sMensaje = sMensaje + "Ingrese la Razón Social" + Chr(13)
    End If
    
    bFound = False
    Dim oParentRow As SSRow
    Dim oChildRow As SSRow
    
    Set oParentRow = Me.grdServicios.GetRow(ssChildRowFirst)
    If oParentRow.HasChild Then
        Set oChildRow = oParentRow.GetChild(ssChildRowFirst)
        If oChildRow.Cells("IdEstadoFacturacion").Value = 6 Or oChildRow.Cells("IdEstadoFacturacion").Value = 7 Then
            bFound = True
        End If
        Do While oChildRow.HasNextSibling
            Set oChildRow = oChildRow.GetSibling(ssSiblingRowNext)
            If oChildRow.Cells("IdEstadoFacturacion").Value = 6 Or oChildRow.Cells("IdEstadoFacturacion").Value = 7 Then
                bFound = True
            End If
        Loop
    End If
    
    Do While oParentRow.HasNextSibling
        'Obtiene el siguiente registro de la banda 0
        Set oParentRow = oParentRow.GetSibling(ssSiblingRowNext)
    
        'Procesa todo los hijos de la banda 1
        If oParentRow.HasChild Then
            Set oChildRow = oParentRow.GetChild(ssChildRowFirst)
            If oChildRow.Cells("IdEstadoFacturacion").Value = 6 Or oChildRow.Cells("IdEstadoFacturacion").Value = 7 Then
                bFound = True
            End If
            Do While oChildRow.HasNextSibling
                Set oChildRow = oChildRow.GetSibling(ssSiblingRowNext)
                If oChildRow.Cells("IdEstadoFacturacion").Value = 6 Or oChildRow.Cells("IdEstadoFacturacion").Value = 7 Then
                    bFound = True
                End If
            Loop
        End If
    Loop
    
'    If mrs_ComprobantesDetalle.EOF = False And mrs_ComprobantesDetalle.BOF = False Then
'    mrs_ComprobantesDetalle.MoveFirst
'    Do Until mrs_ComprobantesDetalle.EOF
'        If mrs_ComprobantesDetalle.Fields!CheckSeleccionado Then
'            bFound = True
'            Exit Do
'        End If
'    Loop
'    End If
    If Not bFound Then
        sMensaje = sMensaje + "Ingrese los Items a Facturar" + Chr(13)
    End If
    
    If Me.lblMontoFaltanteSoles <> "" Then
        If CCurrency(lblMontoFaltanteSoles) > 0 Then
            sMensaje = sMensaje + "No puede registrar un comprobante con faltante de dinero" + Chr(13)
        End If
    End If
    
    If sMensaje <> "" Then
         MsgBox sMensaje, vbExclamation, Me.Caption
         Exit Function
    End If
    
    ValidarDatosObligatorios = True
End Function

