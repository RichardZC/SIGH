VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form InterconsultasDetalle 
   Caption         =   "Form1"
   ClientHeight    =   8760
   ClientLeft      =   915
   ClientTop       =   675
   ClientWidth     =   11535
   Icon            =   "Interconsultas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11535
   Begin UltraGrid.SSUltraGrid grdPacientesEncontrados 
      Height          =   2775
      Left            =   11580
      TabIndex        =   41
      Top             =   990
      Visible         =   0   'False
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   4895
      _Version        =   131072
      GridFlags       =   17040384
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
      Caption         =   "Lista de Pacientes Encontrados"
   End
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
      Height          =   915
      Left            =   30
      TabIndex        =   32
      Top             =   -30
      Width           =   11445
      Begin VB.TextBox txtPrimerNombreBusqueda 
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
         Left            =   4215
         TabIndex        =   39
         Top             =   450
         Width           =   1350
      End
      Begin VB.TextBox txtApellidoPaternoBusqueda 
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
         Left            =   1560
         TabIndex        =   38
         Top             =   450
         Width           =   1185
      End
      Begin VB.TextBox txtApellidoMaternoBusqueda 
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
         Left            =   2805
         TabIndex        =   37
         Top             =   450
         Width           =   1350
      End
      Begin VB.TextBox txtSegundoNombreBusqueda 
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
         Left            =   5625
         TabIndex        =   36
         Top             =   450
         Width           =   1350
      End
      Begin VB.TextBox txtNroDNIBusqueda 
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
         Left            =   7035
         TabIndex        =   35
         Top             =   450
         Width           =   1185
      End
      Begin VB.CommandButton btnBuscarPaciente 
         Height          =   315
         Left            =   8310
         Picture         =   "Interconsultas.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   420
         Width           =   1305
      End
      Begin VB.TextBox txtNroHistoriaBusqueda 
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
         Left            =   120
         TabIndex        =   33
         Top             =   450
         Width           =   1365
      End
      Begin VB.Label Label50 
         Caption         =   "Nº Historia      Ap. paterno      Ap. materno       1er nombre      2do nombre         DNI"
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
         Left            =   450
         TabIndex        =   40
         Top             =   240
         Width           =   8445
      End
   End
   Begin TabDlg.SSTab tab 
      Height          =   4020
      Left            =   0
      TabIndex        =   31
      Top             =   3480
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   7091
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Diagnóstico realizado en la interconsulta"
      TabPicture(0)   =   "Interconsultas.frx":3913
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ucDiagnosticoDetalle1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin Galenhos.ucDiagnosticoDetalle ucDiagnosticoDetalle1 
         Height          =   3525
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   6218
      End
   End
   Begin VB.CommandButton btnBuscarMedicoRealiza 
      Caption         =   "..."
      Height          =   315
      Left            =   6210
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2640
      Width           =   315
   End
   Begin VB.CommandButton btnBuscarMedicoSolicita 
      Caption         =   "..."
      Height          =   315
      Left            =   6210
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2265
      Width           =   315
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la atención"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   60
      TabIndex        =   17
      Top             =   900
      Width           =   11415
      Begin VB.TextBox lblServicioIngreso 
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
         Left            =   6855
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   630
         Width           =   4395
      End
      Begin VB.TextBox lblFechaIngreso 
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
         Left            =   9690
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   255
         Width           =   1560
      End
      Begin VB.TextBox lblPaciente 
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
         Left            =   3420
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   255
         Width           =   4770
      End
      Begin VB.TextBox lblNroCuentaAtencion 
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
         Left            =   1485
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   255
         Width           =   1050
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
         Left            =   2610
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   630
         Width           =   2790
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
         Left            =   1485
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "Servicio Ingreso"
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
         Left            =   5475
         TabIndex        =   23
         Top             =   675
         Width           =   1305
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
         Left            =   150
         TabIndex        =   21
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Ingreso"
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
         Left            =   8370
         TabIndex        =   20
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Paciente"
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
         Left            =   2640
         TabIndex        =   19
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Nº cuenta "
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
         Left            =   150
         TabIndex        =   18
         Top             =   300
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   75
      TabIndex        =   16
      Top             =   7560
      Width           =   11355
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "Interconsultas.frx":392F
         DownPicture     =   "Interconsultas.frx":3DF3
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
         Left            =   5805
         Picture         =   "Interconsultas.frx":42DF
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "Interconsultas.frx":47CB
         DownPicture     =   "Interconsultas.frx":4C2B
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
         Left            =   4260
         Picture         =   "Interconsultas.frx":50A0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame fraInterconsulta 
      Caption         =   "Interconsulta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   60
      TabIndex        =   11
      Top             =   1980
      Width           =   11415
      Begin VB.ComboBox cmbIdTipoConsulta 
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
         Left            =   1500
         TabIndex        =   42
         Top             =   1000
         Width           =   9800
      End
      Begin VB.TextBox lblMedicoRealiza 
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
         Left            =   6540
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   675
         Width           =   4710
      End
      Begin VB.TextBox lblNombreMedicoSolicita 
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
         Left            =   6555
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   300
         Width           =   4695
      End
      Begin VB.TextBox txtIdMedicoSolicita 
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
         Left            =   5190
         TabIndex        =   2
         Top             =   300
         Width           =   885
      End
      Begin VB.TextBox txtIdMedicoRealiza 
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
         Left            =   5190
         TabIndex        =   6
         Top             =   660
         Width           =   885
      End
      Begin MSMask.MaskEdBox txtHoraSolicitud 
         Height          =   315
         Left            =   2955
         TabIndex        =   1
         Top             =   300
         Width           =   720
         _ExtentX        =   1270
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
      Begin MSMask.MaskEdBox txtFechaSolicitud 
         Height          =   315
         Left            =   1500
         TabIndex        =   0
         Top             =   300
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
      Begin MSMask.MaskEdBox txtHoraRealizacion 
         Height          =   315
         Left            =   2955
         TabIndex        =   5
         Top             =   660
         Width           =   735
         _ExtentX        =   1296
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
      Begin MSMask.MaskEdBox txtFechaRealizacion 
         Height          =   315
         Left            =   1500
         TabIndex        =   4
         Top             =   660
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
      Begin VB.Label Label5 
         Caption         =   "Tipo Consulta"
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
         TabIndex        =   43
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label Label45 
         Caption         =   "Medico solicita"
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
         Left            =   3810
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label55 
         Caption         =   "Fecha solicitud"
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
         Left            =   135
         TabIndex        =   14
         Top             =   330
         Width           =   1320
      End
      Begin VB.Label Label63 
         Caption         =   "Medico realiza"
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
         Left            =   3795
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label64 
         Caption         =   "Fecha realización"
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
         Left            =   135
         TabIndex        =   12
         Top             =   690
         Width           =   1440
      End
   End
End
Attribute VB_Name = "InterconsultasDetalle"
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

Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim mo_AtencionesInterconsultas As New DOAtencionInterconsulta
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim mo_Diagnosticos As New Collection
Dim mo_Procedimientos As New Collection
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdDetalleProducto As Long
Dim ml_IdCuentaAtencion As Long
Dim ml_IdInterconsulta As Long
Dim ml_IdTipoServicio As Long
Dim ml_IdAtencion   As Long
Dim mo_cmbIdTipoGenHistoriaClinica As New SIGHComun.ListaDespleglable
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim mo_cmbIdTipoConsulta As New SIGHComun.ListaDespleglable 'WCG20060313 (para el tipo de producto)
Dim mo_FacturacionServicios As New Collection 'WCG20060313 (para los servicios a facturar)
Dim oCuentaAtencion As New DOCuentaAtencion 'WCG20060313 (para los datos de la cuenta del paciente)
Dim mo_FacturacionServicioAsociada As New DOFacturacionServicios 'WCG20060314
Dim ml_IdServicioIngreso As Long 'WCG20060315


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
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let IdDetalleProducto(lValue As Long)
   ml_IdDetalleProducto = lValue
End Property
Property Get IdDetalleProducto() As Long
   IdDetalleProducto = ml_IdDetalleProducto
End Property
Property Let IdCuentaAtencion(lValue As Long)
   ml_IdCuentaAtencion = lValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let IdInterconsulta(lValue As Long)
   ml_IdInterconsulta = lValue
End Property
Property Get IdInterconsulta() As Long
   IdInterconsulta = ml_IdInterconsulta
End Property
Property Let IdTipoServicio(lValue As Long)
   ml_IdTipoServicio = lValue
End Property
Property Get IdTipoServicio() As Long
   IdTipoServicio = ml_IdTipoServicio
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String
       
       mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
       mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()

        Me.ucDiagnosticoDetalle1.TipoDiagnostico = sghInterconsultas
        Me.ucDiagnosticoDetalle1.ConfigurarComboBoxes
        'WCG20060313 (cargamos la lista de productos asociados a la interconsulta)
        mo_cmbIdTipoConsulta.BoundColumn = "IdProducto"
        mo_cmbIdTipoConsulta.ListField = "Descripcion"
        Set mo_cmbIdTipoConsulta.RowSource = mo_AdminFacturacion.FacturacionSeleccionarTiposConsultaInterconsulta()
        'WCG20060313


End Sub

Private Sub btnBuscarMedicoRealiza_Click()
    CompletarDatosDeMedico Me.txtIdMedicoRealiza, Me.lblMedicoRealiza
    txtIdMedicoRealiza.SetFocus
End Sub

Private Sub btnBuscarMedicoSolicita_Click()
    CompletarDatosDeMedico Me.txtIdMedicoSolicita, Me.lblNombreMedicoSolicita
    txtIdMedicoSolicita.SetFocus
End Sub

Private Sub btnBuscarPaciente_Click()
Dim rsHistorias As New Recordset
Dim oDOPaciente As New doPaciente
Dim oDOAtencion As New DOAtencion
    
    oDOPaciente.ApellidoPaterno = Me.txtApellidoPaternoBusqueda.Text
    oDOPaciente.ApellidoMaterno = Me.txtApellidoMaternoBusqueda
    oDOPaciente.PrimerNombre = Me.txtPrimerNombreBusqueda
    oDOPaciente.SegundoNombre = Me.txtSegundoNombreBusqueda
    oDOPaciente.NroHistoriaClinica = Val(Me.txtNroHistoriaBusqueda.Text)
    oDOPaciente.IdDocIdentidad = 1
    oDOPaciente.NroDocumento = Me.txtNroDNIBusqueda
    
    If (oDOPaciente.ApellidoPaterno + oDOPaciente.ApellidoMaterno + _
    oDOPaciente.PrimerNombre + oDOPaciente.SegundoNombre = "") And _
    (Val(Me.txtNroHistoriaBusqueda.Text) = 0) And _
    (oDOPaciente.NroDocumento = "") Then
        MsgBox "Ingrese alguno de los valores de búsqueda", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    Select Case ml_IdTipoServicio
    Case 1
        Set rsHistorias = mo_AdminAdmision.AtencionesFiltrarConsultaExterna(oDOPaciente, oDOAtencion)
    Case 2
        Set rsHistorias = mo_AdminAdmision.AtencionesFiltrarConsultorioEmergencia(oDOPaciente, oDOAtencion)
    Case 3
        Set rsHistorias = mo_AdminAdmision.AtencionesFiltrarHospitalizacion(oDOPaciente, oDOAtencion)
    Case 4
        Set rsHistorias = mo_AdminAdmision.AtencionesFiltrarObservacionEmergencia(oDOPaciente, oDOAtencion)
    End Select
    Screen.MousePointer = vbDefault
    
    txtNroHistoriaBusqueda.Text = ""
    Set grdPacientesEncontrados.DataSource = rsHistorias
    
    'Si hay una sola coincidencia
    If rsHistorias.RecordCount = 1 Then
        Me.grdPacientesEncontrados.Visible = False
        
        rsHistorias.MoveFirst
        LimpiarDatosDeAtencion
        
        Me.txtIdNroHistoria.Text = rsHistorias!NroHistoriaClinica
        mo_cmbIdTipoGenHistoriaClinica.BoundText = rsHistorias!IdTipoNumeracion
        Me.lblFechaIngreso = rsHistorias!FechaIngreso
        Me.lblServicioIngreso = rsHistorias!ServicioIngreso
        Me.lblPaciente = rsHistorias!ApellidoPaterno + " " + rsHistorias!ApellidoMaterno + " " + rsHistorias!PrimerNombre + " " + ("" & rsHistorias!SegundoNombre)
        Me.IdAtencion = rsHistorias!IdAtencion
        Me.IdCuentaAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarIdPorIdAtencion(Me.IdAtencion)
        Me.lblNroCuentaAtencion = IdCuentaAtencion
        ml_IdServicioIngreso = rsHistorias!IdServicioIngreso
    ElseIf rsHistorias.RecordCount > 1 Then
        Me.grdPacientesEncontrados.Visible = True

    ElseIf rsHistorias.RecordCount = 0 Then
        Me.grdPacientesEncontrados.Visible = False

        Select Case ml_IdTipoServicio
        Case 1
            MsgBox "No se encontraron atenciones de consulta externa", vbInformation, Me.Caption
        Case 2
            MsgBox "No se encontraron atenciones de consultorio de emergencia", vbInformation, Me.Caption
        Case 3
            MsgBox "No se encontraron atenciones de hospitalización", vbInformation, Me.Caption
        Case 4
            MsgBox "No se encontraron atenciones de observación de emergencia", vbInformation, Me.Caption
        End Select
    
        LimpiarDatosDeAtencion
    End If

    mo_Apariencia.ConfigurarFilasBiColores Me.grdPacientesEncontrados, SIGHComun.GrillaConFilasBicolor
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmbNroHistoria_Click()
End Sub
Sub LimpiarDatosDeAtencion()
        
        Me.txtIdNroHistoria.Text = ""
        mo_cmbIdTipoGenHistoriaClinica.BoundText = ""
        Me.lblFechaIngreso = ""
        Me.lblServicioIngreso = ""
        Me.lblPaciente = ""
        Me.IdAtencion = 0
        Me.IdCuentaAtencion = 0

End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = Me.cmbIdTipoGenHistoriaClinica
    Set mo_cmbIdTipoConsulta.MiComboBox = Me.cmbIdTipoConsulta 'WCG20060313 (para los productso asociados a este servicio)
End Sub

Private Sub grdPacientesEncontrados_DblClick()
Dim rsHistorias As Recordset

    On Error Resume Next
    Set rsHistorias = Me.grdPacientesEncontrados.DataSource
    
    Me.txtIdNroHistoria.Text = rsHistorias!NroHistoriaClinica
    mo_cmbIdTipoGenHistoriaClinica.BoundText = rsHistorias!IdTipoNumeracion
    Me.lblFechaIngreso = rsHistorias!FechaIngreso
    Me.lblServicioIngreso = rsHistorias!ServicioIngreso
    Me.lblPaciente = rsHistorias!ApellidoPaterno + " " + rsHistorias!ApellidoMaterno + " " + rsHistorias!PrimerNombre + " " + ("" & rsHistorias!SegundoNombre)
    Me.IdAtencion = rsHistorias!IdAtencion
    Me.IdCuentaAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarIdPorIdAtencion(Me.IdAtencion)
    Me.lblNroCuentaAtencion = IdCuentaAtencion
    ml_IdServicioIngreso = rsHistorias!IdServicioIngreso 'WCG20060317
    
    Me.grdPacientesEncontrados.Visible = False: DoEvents
End Sub

Private Sub grdPacientesEncontrados_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdPacientesEncontrados.Bands(0).Columns("IdPaciente").Hidden = True
    grdPacientesEncontrados.Bands(0).Columns("IdTipoNumeracion").Hidden = True
    grdPacientesEncontrados.Bands(0).Columns("IdepisodioAtencion").Hidden = True
    grdPacientesEncontrados.Bands(0).Columns("IdAtencion").Hidden = True
    grdPacientesEncontrados.Bands(0).Columns("IdCita").Hidden = True
    
    grdPacientesEncontrados.Bands(0).Columns("IdCuentaAtencion").Header.Caption = "Nro Cuenta"
    grdPacientesEncontrados.Bands(0).Columns("IdCuentaAtencion").Width = 1300
    
    grdPacientesEncontrados.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nro Historia"
    grdPacientesEncontrados.Bands(0).Columns("NroHistoriaClinica").Width = 1300
    
    grdPacientesEncontrados.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdPacientesEncontrados.Bands(0).Columns("ApellidoPaterno").Width = 1500
    
    grdPacientesEncontrados.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdPacientesEncontrados.Bands(0).Columns("ApellidoMaterno").Width = 1500
    
    grdPacientesEncontrados.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdPacientesEncontrados.Bands(0).Columns("PrimerNombre").Width = 1500

    grdPacientesEncontrados.Bands(0).Columns("SegundoNombre").Header.Caption = "2do Nombre"
    grdPacientesEncontrados.Bands(0).Columns("SegundoNombre").Width = 1500

    'grdPacientesEncontrados.Bands(0).Columns("FechaNacimiento").Header.Caption = "Fecha Nac."
    'grdPacientesEncontrados.Bands(0).Columns("FechaNacimiento").Width = 1500

    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Header.Caption = "Tipo Numeración"
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Width = 1500
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").CellAppearance.TextAlign = ssAlignRight

    'grdPacientesEncontrados.Bands(0).Columns("TipoServicio").Header.Caption = "Ult. Tipo Serv."
    'grdPacientesEncontrados.Bands(0).Columns("TipoServicio").Width = 2000

    grdPacientesEncontrados.Bands(0).Columns("FechaIngreso").Header.Caption = "Ult. Fec Ing."
    grdPacientesEncontrados.Bands(0).Columns("FechaIngreso").Width = 1500

    'grdPacientesEncontrados.Bands(0).Columns("FechaEgreso").Header.Caption = "Ult. Fec Egr."
    'grdPacientesEncontrados.Bands(0).Columns("FechaEgreso").Width = 1500

    grdPacientesEncontrados.Bands(0).Columns("ServicioIngreso").Header.Caption = "Ult. Serv. Ing."
    grdPacientesEncontrados.Bands(0).Columns("ServicioIngreso").Width = 1500
End Sub

Private Sub grdPacientesEncontrados_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = vbKeyReturn Then
        grdPacientesEncontrados_DblClick
    End If
End Sub

Private Sub txtHoraSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraSolicitud
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtHoraSolicitud_LostFocus()
   mo_Formulario.MarcarComoVacio txtHoraSolicitud
End Sub

Private Sub txtHoraSolicitud_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtHoraRealizacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraRealizacion
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtHoraRealizacion_LostFocus()
   mo_Formulario.MarcarComoVacio txtHoraRealizacion
End Sub

Private Sub txtHoraRealizacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFechaSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaSolicitud
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaSolicitud_LostFocus()

       If txtFechaSolicitud <> SIGHComun.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaSolicitud, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, "Datos de paciente"
                 txtFechaSolicitud = SIGHComun.FECHA_VACIA_DMY
            End If
        End If
        
        mo_Formulario.MarcarComoVacio txtFechaSolicitud
End Sub

Private Sub txtFechaSolicitud_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFechaRealizacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaRealizacion
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaRealizacion_LostFocus()
       If txtFechaRealizacion <> SIGHComun.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaRealizacion, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, "Datos de paciente"
                 txtFechaRealizacion = SIGHComun.FECHA_VACIA_DMY
            End If
        End If

   mo_Formulario.MarcarComoVacio txtFechaRealizacion
End Sub

Private Sub txtFechaRealizacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtIdMedicoRealiza_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdMedicoRealiza
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdMedicoRealiza_LostFocus()
    CompletarDatosDeMedicoEnElLostFocus txtIdMedicoRealiza, Me.lblMedicoRealiza
    mo_Formulario.MarcarComoVacio txtIdMedicoRealiza
End Sub

Private Sub txtIdMedicoRealiza_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtIdMedicoSolicita_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdMedicoSolicita
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdMedicoSolicita_LostFocus()
        CompletarDatosDeMedicoEnElLostFocus txtIdMedicoSolicita, Me.lblNombreMedicoSolicita
        mo_Formulario.MarcarComoVacio txtIdMedicoSolicita
        
        'Por defecto
        If Trim(txtIdMedicoRealiza) = "" Then
            txtIdMedicoRealiza = txtIdMedicoSolicita
            txtIdMedicoRealiza_LostFocus
        End If
End Sub

Private Sub txtIdMedicoSolicita_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

    Me.grdPacientesEncontrados.Left = 210
    Me.grdPacientesEncontrados.Top = 780
    
    mo_Formulario.HabilitarDeshabilitar lblNroCuentaAtencion, False
    mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    mo_Formulario.HabilitarDeshabilitar lblPaciente, False
    mo_Formulario.HabilitarDeshabilitar lblFechaIngreso, False
    mo_Formulario.HabilitarDeshabilitar lblServicioIngreso, False
    mo_Formulario.HabilitarDeshabilitar lblNombreMedicoSolicita, False
    mo_Formulario.HabilitarDeshabilitar lblMedicoRealiza, False
    'WCG20060314
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoConsulta, True

 Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
 End Select
 
    Select Case mi_Opcion
    Case sghAgregar
    Case sghModificar
    Case sghConsultar
            Me.fraBusqueda.Enabled = False
            Me.fraInterconsulta.Enabled = False
            Me.ucDiagnosticoDetalle1.BotonAgregarEnabled = False
            Me.ucDiagnosticoDetalle1.BotonQuitarEnabled = False
            'WCG comentado por facturacion
            'Me.ucProcedimientoDetalle1.BotonAgregarEnabled = False
            'Me.ucProcedimientoDetalle1.BotonQuitarEnabled = False
            Me.btnAceptar.Enabled = False
            
    Case sghEliminar
            Me.fraBusqueda.Enabled = False
            Me.fraInterconsulta.Enabled = False
            Me.ucDiagnosticoDetalle1.BotonAgregarEnabled = False
            Me.ucDiagnosticoDetalle1.BotonQuitarEnabled = False
            'WCG comentado por facturacion
            'Me.ucProcedimientoDetalle1.BotonAgregarEnabled = False
            'Me.ucProcedimientoDetalle1.BotonQuitarEnabled = False
    
    End Select
 
 'WCG comentado por facturacion
 'Me.ucProcedimientoDetalle1.TipoServicio = ml_IdTipoServicio
 
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Interconsultas"
       Case sghModificar
           Me.Caption = "Modificar Interconsultas"
       Case sghConsultar
           Me.Caption = "Consultar Interconsultas"
       Case sghEliminar
           Me.Caption = "Eliminar Interconsultas"
       End Select

        CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
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
            btnBuscarPaciente_Click
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox " Los datos se agregaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron exitosamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   'If IdDetalleProducto = 0 Then
   '    sMensaje = sMensaje + "Ingrese el valor de IdDetalleProducto" + Chr(13)
   'End If
   If IdCuentaAtencion = 0 Then
       sMensaje = sMensaje + "Ingrese la cuenta de atención" + Chr(13)
   End If
   'If IdInterconsulta = 0 Then
   '    sMensaje = sMensaje + "Ingrese el valor de IdInterconsulta" + Chr(13)
   'End If
   If Me.txtHoraSolicitud.Text = "__:__" Then
       sMensaje = sMensaje + "Ingrese la hora de solicitud" + Chr(13)
   End If
   If Me.txtHoraRealizacion.Text = "__:__" Then
       sMensaje = sMensaje + "Ingrese la hora de realización" + Chr(13)
   End If
   If Me.txtFechaSolicitud.Text = SIGHComun.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Ingrese la fecha de solicitud" + Chr(13)
   End If
   If Me.txtFechaRealizacion.Text = SIGHComun.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Ingrese la fecha de realización" + Chr(13)
   End If
   If Val(Me.txtIdMedicoRealiza.Text) = 0 Then
       sMensaje = sMensaje + "Ingrese el médico que realiza la interconsulta" + Chr(13)
   End If
   If Val(Me.txtIdMedicoSolicita.Text) = 0 Then
       sMensaje = sMensaje + "Ingrese el médico que solicita la interconsulta" + Chr(13)
   End If
   'WCG20060313
   If (mi_Opcion = sghAgregar) Or (mi_Opcion = sghModificar) Then
        If Val(mo_cmbIdTipoConsulta.BoundText) = 0 Then
             sMensaje = sMensaje + "Seleccione el tipo de consulta" + Chr(13)
        End If
   End If
   'WCG20060313
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If

   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   
'    If CDate(Me.txtFechaSolicitud + " " + Me.txtHoraSolicitud) > Date Then
'        MsgBox "La fecha de solicitud no puede ser mayor que la fecha de hoy", vbExclamation, Me.Caption
'        Exit Function
'    End If
   
    If CDate(Me.txtFechaSolicitud + " " + Me.txtHoraSolicitud) < CDate(Me.lblFechaIngreso) Then
        MsgBox "La fecha de solicitud no puede ser menor que la fecha de ingreso de la atención", vbExclamation, Me.Caption
        Exit Function
    End If
   
    'If CDate(Me.txtFechaRealizacion + " " + Me.txtHoraRealizacion) > Now Then
    '    MsgBox "La fecha de realización no puede ser mayor que la fecha de hoy", vbExclamation, Me.Caption
    '    Exit Function
    'End If
   
    If CDate(Me.txtFechaRealizacion + " " + Me.txtHoraRealizacion) < CDate(Me.lblFechaIngreso) Then
        MsgBox "La fecha de realización no puede ser menor que la fecha de ingreso de la atención", vbExclamation, Me.Caption
        Exit Function
    End If
   
    If CDate(Me.txtFechaSolicitud + " " + Me.txtHoraSolicitud) > CDate(Me.txtFechaRealizacion + " " + Me.txtHoraRealizacion) Then
        MsgBox "La fecha de solicitud no puede ser menor que la fecha de realización de la interconsulta", vbExclamation, Me.Caption
        Exit Function
    End If
    'WCG20060313
    If (mi_Opcion = sghAgregar) Or (mi_Opcion = sghModificar) Then
        If Val(mo_cmbIdTipoConsulta.BoundText) = 0 Then
             MsgBox "Seleccione el tipo de consulta", vbExclamation, Me.Caption
             Exit Function
        End If
   End If
   'WCG20060313
   
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_AtencionesInterconsultas
           .IdDetalleProducto = Me.IdDetalleProducto 'WCG20060313
           '.IdCuentaAtencion = Me.IdCuentaAtencion
           .IdInterconsulta = Me.IdInterconsulta
           .HoraSolicitud = Me.txtHoraSolicitud.Text
           .HoraRealizacion = Me.txtHoraRealizacion.Text
           .FechaSolicitud = Me.txtFechaSolicitud.Text
           .FechaRealizacion = Me.txtFechaRealizacion.Text
           .IdMedicoRealiza = Me.txtIdMedicoRealiza.Tag
           .IdMedicoSolicita = Me.txtIdMedicoSolicita.Tag
           .IdUsuarioAuditoria = Me.IdUsuario
   End With
   
   Me.ucDiagnosticoDetalle1.IdUsuario = Me.IdUsuario
   Me.ucDiagnosticoDetalle1.CargarDiagnosticosAlObjetoDatos mo_Diagnosticos
   'WCG20060313
   CargaDeServiciosAFacturar
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminAdmision.AtencionesInterconsultasAgregar(mo_AtencionesInterconsultas, mo_Diagnosticos, mo_FacturacionServicios, oCuentaAtencion) 'WCG20060314

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminAdmision.AtencionesInterconsultasModificar(mo_AtencionesInterconsultas, mo_Diagnosticos, mo_FacturacionServicioAsociada, oCuentaAtencion) 'WCG20060314

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminAdmision.AtencionesInterconsultasEliminar(mo_AtencionesInterconsultas, mo_Diagnosticos, ml_IdAtencion, mo_FacturacionServicioAsociada) 'WCG20060314

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oDOServicio As New DOServicio

        Set mo_AtencionesInterconsultas = mo_AdminAdmision.AtencionesInterconsultasSeleccionarPorId(Me.IdInterconsulta)
        If mo_AdminAdmision.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminAdmision.MensajeError, vbCritical, Me.Caption"
             mb_ExistenDatos = False
             Exit Sub
        End If
        
       If Not mo_AtencionesInterconsultas Is Nothing Then
           With mo_AtencionesInterconsultas
                Me.IdDetalleProducto = .IdDetalleProducto
                'Me.IdCuentaAtencion = .IdCuentaAtencion
                Me.IdInterconsulta = .IdInterconsulta
                Me.txtHoraSolicitud.Text = .HoraSolicitud
                Me.txtHoraRealizacion.Text = .HoraRealizacion
                Me.txtFechaSolicitud.Text = .FechaSolicitud
                Me.txtFechaRealizacion.Text = .FechaRealizacion
                Me.txtIdMedicoRealiza.Tag = .IdMedicoRealiza
                Me.txtIdMedicoSolicita.Tag = .IdMedicoSolicita
                
                'CargarDatosDelaAtencion .IdCuentaAtencion
                
                If mo_AdminProgramacion.MedicosSeleccionarPorId(.IdMedicoRealiza, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
                    Me.txtIdMedicoRealiza = oDOEmpleado.CodigoPlanilla
                    Me.lblMedicoRealiza = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                Else
                    Me.lblMedicoRealiza = ""
                End If
                
                If mo_AdminProgramacion.MedicosSeleccionarPorId(.IdMedicoSolicita, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
                    Me.txtIdMedicoSolicita = oDOEmpleado.CodigoPlanilla
                    Me.lblNombreMedicoSolicita = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                Else
                    Me.lblNombreMedicoSolicita = ""
                End If
                
                mb_ExistenDatos = True
           End With
           
           Me.ucDiagnosticoDetalle1.IdAtencion = Me.IdAtencion
           Me.ucDiagnosticoDetalle1.CargarDiagnosticosDeInterconsultas Me.IdInterconsulta
           'WCG20060314 traemos los datos de la facturacion asociada
           Set mo_FacturacionServicioAsociada = mo_AdminFacturacion.FacturacionServiciosSeleccionarPorId(Me.IdDetalleProducto)
           With mo_FacturacionServicioAsociada
                mo_cmbIdTipoConsulta.BoundText = mo_FacturacionServicioAsociada.IdProducto
                If .IdEstadoFacturacion <> 1 Then
                    mo_Formulario.HabilitarDeshabilitar cmbIdTipoConsulta, False
                Else
                    mo_Formulario.HabilitarDeshabilitar cmbIdTipoConsulta, True
                End If
           End With
           'WCG20060314
           
           
           'WCG comentado por facturacion
           'Me.ucProcedimientoDetalle1.IdCuentaAtencion = Me.IdCuentaAtencion
           'Me.ucProcedimientoDetalle1.IdInterconsulta = Me.IdInterconsulta
           'Me.ucProcedimientoDetalle1.CargarDatosDeProcedimientos
           
           
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
        'WCG20060315
       If mi_Opcion = sghEliminar Then
            If mo_FacturacionServicioAsociada.IdEstadoFacturacion <> sghEstadoFacturacion.sghPendientePago Then
                MsgBox "No puede eliminar la Interconsulta porque el Producto ya está pagado", vbInformation + vbOKOnly, "Mensaje Informativo"
                btnAceptar.Enabled = False
                Exit Sub
            End If
       End If
End Sub
Sub CargarDatosDelaAtencion(lIdAtencion As Long)
Dim oDOPaciente As New doPaciente
Dim oDOAtencion As New DOAtencion
Dim rsHistorias As New ADODB.Recordset

    oDOAtencion.IdAtencion = lIdAtencion
    
    Select Case ml_IdTipoServicio
    Case 1
        Set rsHistorias = mo_AdminAdmision.AtencionesFiltrarConsultaExterna(oDOPaciente, oDOAtencion)
    Case 2
        Set rsHistorias = mo_AdminAdmision.AtencionesFiltrarConsultorioEmergencia(oDOPaciente, oDOAtencion)
    Case 3
        Set rsHistorias = mo_AdminAdmision.AtencionesFiltrarHospitalizacion(oDOPaciente, oDOAtencion)
    Case 4
        Set rsHistorias = mo_AdminAdmision.AtencionesFiltrarObservacionEmergencia(oDOPaciente, oDOAtencion)
    End Select
    
    'Si hay una sola coincidencia
    If Not (rsHistorias.EOF And rsHistorias.BOF) Then
        LimpiarDatosDeAtencion
        Me.txtIdNroHistoria.Text = rsHistorias!NroHistoriaClinica
        mo_cmbIdTipoGenHistoriaClinica.BoundText = rsHistorias!IdTipoNumeracion
        Me.lblFechaIngreso = rsHistorias!FechaIngreso
        Me.lblServicioIngreso = rsHistorias!ServicioIngreso
        Me.lblPaciente = rsHistorias!ApellidoPaterno + " " + rsHistorias!ApellidoMaterno + " " + rsHistorias!PrimerNombre + " " + ("" & rsHistorias!SegundoNombre)
        Me.IdAtencion = rsHistorias!IdAtencion
        Me.IdCuentaAtencion = rsHistorias!IdCuentaAtencion
        Me.lblNroCuentaAtencion = rsHistorias!IdCuentaAtencion
        ml_IdServicioIngreso = rsHistorias!IdServicioIngreso 'WCG20060317
    End If
    rsHistorias.Close

End Sub
'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           Me.IdDetalleProducto = 0
           Me.IdCuentaAtencion = 0
           Me.IdInterconsulta = 0
           Me.txtHoraSolicitud.Text = SIGHComun.HORA_VACIA_HM
           Me.txtHoraRealizacion.Text = SIGHComun.HORA_VACIA_HM
           Me.txtFechaSolicitud.Text = SIGHComun.FECHA_VACIA_DMY
           Me.txtFechaRealizacion.Text = SIGHComun.FECHA_VACIA_DMY
           Me.txtIdMedicoRealiza.Text = ""
           Me.txtIdMedicoSolicita.Text = ""
   
End Sub

Sub CompletarDatosDeMedico(txtMedico As TextBox, lblNombreMedico As TextBox)
Dim oBusqueda As New MedicosBusqueda
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection

    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
            txtMedico.Text = oDOEmpleado.CodigoPlanilla
            txtMedico.Tag = oDoMedico.IdMedico
            lblNombreMedico.Text = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If

End Sub
Sub CompletarDatosDeMedicoEnElLostFocus(txtMedico As TextBox, lblNombreMedico As TextBox)
Dim oMedicosEspecialidad As New Collection

    txtMedico = Trim(txtMedico)
    If txtMedico <> "" Then
        Dim oDOEmpleado As New dOEmpleado
        Dim oDoMedico As New DOMedico
        If mo_AdminProgramacion.MedicosSeleccionarPorCodigo(CStr(txtMedico), oDoMedico, oDOEmpleado, oMedicosEspecialidad) Then
            txtMedico.Tag = oDoMedico.IdMedico
            Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(oDoMedico.IdEmpleado)
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        Else
            txtMedico.Tag = ""
            lblNombreMedico = ""
        End If
    Else
            txtMedico.Tag = ""
            lblNombreMedico = ""
    End If
    
End Sub

Private Sub txtNroDNIBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroDNIBusqueda
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNroDNIBusqueda_LostFocus()
txtNroDNIBusqueda.Text = mo_Teclado.CapitalizarNombres(txtNroDNIBusqueda.Text)
   'mo_Formulario.MarcarComoVacio txtNroDNIBusqueda
End Sub

Private Sub txtNroDNIBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtApellidoPaternoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaternoBusqueda
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoPaternoBusqueda_LostFocus()
txtApellidoPaternoBusqueda.Text = mo_Teclado.CapitalizarNombres(txtApellidoPaternoBusqueda.Text)
   'mo_Formulario.MarcarComoVacio txtApellidoPaternoBusqueda
End Sub

Private Sub txtApellidoPaternoBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtApellidoMaternoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaternoBusqueda
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoMaternoBusqueda_LostFocus()
txtApellidoMaternoBusqueda.Text = mo_Teclado.CapitalizarNombres(txtApellidoMaternoBusqueda.Text)
   'mo_Formulario.MarcarComoVacio txtApellidoMaternoBusqueda
End Sub

Private Sub txtApellidoMaternoBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtPrimerNombreBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombreBusqueda
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtPrimerNombreBusqueda_LostFocus()
txtPrimerNombreBusqueda.Text = mo_Teclado.CapitalizarNombres(txtPrimerNombreBusqueda.Text)
   'mo_Formulario.MarcarComoVacio txtPrimerNombreBusqueda
End Sub

Private Sub txtPrimerNombreBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtSegundoNombreBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtSegundoNombreBusqueda
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtSegundoNombreBusqueda_LostFocus()
txtSegundoNombreBusqueda.Text = mo_Teclado.CapitalizarNombres(txtSegundoNombreBusqueda.Text)
   'mo_Formulario.MarcarComoVacio txtSegundoNombreBusqueda
End Sub

Private Sub txtSegundoNombreBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Sub CargaDeServiciosAFacturar()
Dim oCatalogoServicios As New DOCatalogoServicio
    
    'WCG_2006
    If mi_Opcion = sghAgregar Then
        Set oCuentaAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(ml_IdCuentaAtencion)
        Dim oConsulta As New DOFacturacionServicios
        Set oCatalogoServicios = mo_AdminServiciosComunes.CatalogoServiciosSeleccionarPorId(Val(mo_cmbIdTipoConsulta.BoundText))
        With oConsulta
            '.IdCuentaAtencion = ml_IdCuentaAtencion
            .IdFacturacionServicio = Val(mo_cmbIdTipoConsulta.BoundText)
            '.IdFacturacionServicio = ml_IdTipoServicio '(preguntar xq) WCG20060313
            '.IdFuenteFinanciamiento = oCuentaAtencion.IdFuenteFinanciamiento
            '.IdTipoFinanciamiento = oCuentaAtencion.IdTipoFinanciamiento
            .Cantidad = 1
            .IdProducto = Val(mo_cmbIdTipoConsulta.BoundText)
            .IdUsuarioAuditoria = ml_IdUsuario
            .IdEstadoFacturacion = 2
            .FechaAutorizaPendiente = 0
            .FechaAutorizaSeguro = 0
            .IdCentroCosto = 0
            .IdEmpAutorizaPendiente = 0
            .IdEmpAutorizaSeguro = 0
            .PrecioUnitario = 0
            .TotalPorPagar = 0
        End With
        mo_FacturacionServicios.Add oConsulta
    End If
    
    If mi_Opcion = sghModificar Then
        Set oCuentaAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(ml_IdCuentaAtencion)
        With mo_FacturacionServicioAsociada
            .IdProducto = Val(mo_cmbIdTipoConsulta.BoundText)
            .IdUsuarioAuditoria = ml_IdUsuario
            .IdEstadoFacturacion = 2
        End With
    End If
End Sub
