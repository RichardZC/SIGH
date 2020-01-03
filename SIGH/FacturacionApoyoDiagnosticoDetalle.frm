VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FacturacionApoyoDxDetalle 
   Caption         =   "Form1"
   ClientHeight    =   9015
   ClientLeft      =   1095
   ClientTop       =   615
   ClientWidth     =   11520
   Icon            =   "FacturacionApoyoDiagnosticoDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   11520
   Begin VB.Frame fraProcedimiento 
      Caption         =   "Procedimientos"
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
      Left            =   60
      TabIndex        =   42
      Top             =   1950
      Width           =   11415
      Begin VB.CommandButton btnBusquedaServicio 
         Caption         =   "..."
         Height          =   315
         Left            =   2700
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   960
         Width           =   345
      End
      Begin VB.CommandButton btnBusquedaMedico 
         Caption         =   "..."
         Height          =   315
         Left            =   2700
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   600
         Width           =   345
      End
      Begin VB.TextBox txtIdServicio 
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
         Left            =   1680
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox lblDescServicio 
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
         Left            =   3120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   960
         Width           =   5325
      End
      Begin VB.TextBox lblDescMedicoOrdena 
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
         Left            =   3120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   600
         Width           =   5325
      End
      Begin VB.TextBox txtIdMedicoOrdena 
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
         Left            =   1680
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNroOrden 
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
         Left            =   1680
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtHoraOrden 
         Height          =   315
         Left            =   5310
         TabIndex        =   18
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
      Begin MSMask.MaskEdBox txtFechaOrden 
         Height          =   315
         Left            =   3870
         TabIndex        =   17
         Top             =   240
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
      Begin VB.Label Label63 
         Caption         =   "Servicio ordena"
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
         Left            =   150
         TabIndex        =   23
         Top             =   990
         Width           =   1425
      End
      Begin VB.Label Label65 
         Caption         =   "Fecha orden"
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
         Left            =   2790
         TabIndex        =   16
         Top             =   270
         Width           =   1530
      End
      Begin VB.Label Label66 
         Caption         =   "Médico ordena"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   150
         TabIndex        =   19
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label Label69 
         Caption         =   "Orden Nro"
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
         Left            =   150
         TabIndex        =   14
         Top             =   300
         Width           =   1260
      End
   End
   Begin VB.Frame fraExamen 
      Height          =   945
      Left            =   60
      TabIndex        =   41
      Top             =   3390
      Width           =   11415
      Begin VB.CommandButton btnBusquedaProcedimiento 
         Caption         =   "..."
         Height          =   315
         Left            =   2670
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   180
         Width           =   375
      End
      Begin VB.ComboBox cmbIdServicioRealiza 
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
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   540
         Width           =   4485
      End
      Begin VB.CommandButton btnQuitarDx 
         DisabledPicture =   "FacturacionApoyoDiagnosticoDetalle.frx":0CCA
         DownPicture     =   "FacturacionApoyoDiagnosticoDetalle.frx":1055
         Height          =   315
         Left            =   10260
         Picture         =   "FacturacionApoyoDiagnosticoDetalle.frx":13E8
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   540
         Width           =   1005
      End
      Begin VB.CommandButton btnAgregarDx 
         DownPicture     =   "FacturacionApoyoDiagnosticoDetalle.frx":1779
         Height          =   315
         Left            =   9180
         Picture         =   "FacturacionApoyoDiagnosticoDetalle.frx":1BAB
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   540
         Width           =   1005
      End
      Begin VB.TextBox txtIdProcedimiento 
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
         TabIndex        =   28
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox lblDescProcedimiento 
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
         Left            =   3120
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   180
         Width           =   8145
      End
      Begin VB.Label Label6 
         Caption         =   "Serv. Realiza"
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
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Examen"
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
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   90
      TabIndex        =   40
      Top             =   7890
      Width           =   11355
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FacturacionApoyoDiagnosticoDetalle.frx":3DFC
         DownPicture     =   "FacturacionApoyoDiagnosticoDetalle.frx":425C
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
         Left            =   4305
         Picture         =   "FacturacionApoyoDiagnosticoDetalle.frx":46D1
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FacturacionApoyoDiagnosticoDetalle.frx":4B46
         DownPicture     =   "FacturacionApoyoDiagnosticoDetalle.frx":500A
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
         Left            =   5850
         Picture         =   "FacturacionApoyoDiagnosticoDetalle.frx":54F6
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1065
      Left            =   45
      TabIndex        =   39
      Top             =   885
      Width           =   11430
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
         Left            =   3915
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1140
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
         Left            =   5130
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   3315
      End
      Begin VB.TextBox lblNroCuenta 
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
         Left            =   1695
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   255
         Width           =   1140
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
         Left            =   1695
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   645
         Width           =   4020
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
         Left            =   9825
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   255
         Width           =   1425
      End
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
         Left            =   7200
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   630
         Width           =   4065
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
         Left            =   150
         TabIndex        =   3
         Top             =   300
         Width           =   1065
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
         Left            =   165
         TabIndex        =   10
         Top             =   675
         Width           =   1005
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
         Left            =   8535
         TabIndex        =   8
         Top             =   300
         Width           =   1155
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
         Left            =   3000
         TabIndex        =   5
         Top             =   285
         Width           =   975
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
         Left            =   5805
         TabIndex        =   12
         Top             =   675
         Width           =   1305
      End
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
      Height          =   885
      Left            =   45
      TabIndex        =   38
      Top             =   0
      Width           =   11430
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   3120
         Picture         =   "FacturacionApoyoDiagnosticoDetalle.frx":59E2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   330
         Width           =   1305
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
         Left            =   1695
         TabIndex        =   1
         Top             =   330
         Width           =   1350
      End
      Begin VB.Label Label50 
         Caption         =   "Nro Historia"
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
         TabIndex        =   0
         Top             =   390
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList lstOpciones 
      Left            =   240
      Top             =   5940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacturacionApoyoDiagnosticoDetalle.frx":862B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacturacionApoyoDiagnosticoDetalle.frx":8A47
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacturacionApoyoDiagnosticoDetalle.frx":8F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FacturacionApoyoDiagnosticoDetalle.frx":9331
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin UltraGrid.SSUltraGrid grdProcedimientos 
      Height          =   3435
      Left            =   60
      TabIndex        =   35
      Top             =   4380
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6059
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
      Caption         =   "Lista de examenes"
   End
End
Attribute VB_Name = "FacturacionApoyoDxDetalle"
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
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_Diagnosticos As New Collection
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdAtencionApoyoDiagnostico As Long
Dim ml_IdTipoServicio As Long
Dim mo_cmbIdTipoGenHistoriaClinica As New SIGHComun.ListaDespleglable
Dim mo_cmbIdServicioRealiza As New SIGHComun.ListaDespleglable
Dim mo_ApoyoDiagDetalle As New Collection
Dim mo_ApoyoDiagnostico As New DOAtencionApoyoDiagnostico
Dim mrs_ApoyoDiagnostico As New ADODB.Recordset
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim mrs_ApoyoDiagnosticoEliminados As New Recordset
Dim ml_IdDepartamento As Long

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
Property Let IdAtencionApoyoDiagnostico(lValue As Long)
   ml_IdAtencionApoyoDiagnostico = lValue
End Property
Property Get IdAtencionApoyoDiagnostico() As Long
   IdAtencionApoyoDiagnostico = ml_IdAtencionApoyoDiagnostico
End Property
Property Let IdTipoServicio(lValue As Long)
   ml_IdTipoServicio = lValue
End Property
Property Get IdTipoServicio() As Long
   IdTipoServicio = ml_IdTipoServicio
End Property

Property Let IdDepartamento(lValue As Long)
   ml_IdDepartamento = lValue
End Property
Property Get IdDepartamento() As Long
   IdDepartamento = ml_IdDepartamento
End Property

Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String
       
       mo_cmbIdServicioRealiza.BoundColumn = "IdServicio"
       mo_cmbIdServicioRealiza.ListField = "DescripcionLarga"
       Set mo_cmbIdServicioRealiza.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarPorTipoServicioYDpto(5, ml_IdDepartamento)
       
       mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
       mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()

End Sub

Private Sub btnAgregarDx_Click()
    'Validamos que no exista el procedimiento en la lista de procedimientos
    If Val(Me.txtIdProcedimiento.Tag) <= 0 Then
        MsgBox "Debe ingresar el examen", vbExclamation, Me.Caption
        Exit Sub
    End If
    If Val(mo_cmbIdServicioRealiza.BoundText) <= 0 Then
        MsgBox "Debe ingresar el Servicio que realiza", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    If mrs_ApoyoDiagnostico.EOF = False Or mrs_ApoyoDiagnostico.BOF = False Then
        mrs_ApoyoDiagnostico.MoveFirst
        Do While Not mrs_ApoyoDiagnostico.EOF
            If mrs_ApoyoDiagnostico.Fields!IdProcedimiento = Val(Me.txtIdProcedimiento.Tag) Then
                mrs_ApoyoDiagnostico.MoveFirst
                Exit Sub
            End If
            mrs_ApoyoDiagnostico.MoveNext
        Loop
        mrs_ApoyoDiagnostico.MoveFirst
    End If
    
    With mrs_ApoyoDiagnostico
        .AddNew
        .Fields!IdProcedimiento = Val(Me.txtIdProcedimiento.Tag)
        .Fields!CodigoCPT = Me.txtIdProcedimiento
        .Fields!Descripcion = Me.lblDescProcedimiento
        '.Fields!IdMedicoRealiza = 0
        '.Fields!NombreMedico = ""
        .Fields!IdServicioRealiza = Val(mo_cmbIdServicioRealiza.BoundText)
        .Fields!NombreServicio = Trim(Split(cmbIdServicioRealiza.Text, "=")(1))
        .Fields!FechaResultado = 0
        .Fields!HoraResultado = ""
        .Fields!IdFacturacionServicio = 0
        .Fields!EstadoRegistro = "A"
    End With
End Sub

Private Sub btnBuscar_Click()


Dim oDOPaciente As New doPaciente
Dim oDOCuentaAtencion As New DOCuentaAtencion
Dim lIdCuentaAtencionActual As Long
    
    LimpiarDatosDeAtencion
    If (Me.txtNroHistoria) = "" Then
        MsgBox "Ingrese la Historia Clínica a buscar", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Dim rsCuentasAtencion As New ADODB.Recordset
    Dim iCount As Integer

    lIdCuentaAtencionActual = 0
    Set rsCuentasAtencion = mo_AdminCaja.ObtenerCuentasAtencionPorHistoriaClinica(Val(Me.txtNroHistoria))
    iCount = 0
    Do While Not rsCuentasAtencion.EOF
        iCount = iCount + 1
        lIdCuentaAtencionActual = rsCuentasAtencion!IdCuentaAtencion
        rsCuentasAtencion.MoveNext
    Loop
    If iCount > 1 Then
        'Levantamos el formulario para seleccionar la cuenta de atención
        Dim oFrmCuentasAtencion As New CuentasAtencionSeleccionar
        Set oFrmCuentasAtencion.DataSource = rsCuentasAtencion
        oFrmCuentasAtencion.Show vbModal
        If oFrmCuentasAtencion.BotonPresionado = sghCancelar Then
            lIdCuentaAtencionActual = 0
        Else
            lIdCuentaAtencionActual = oFrmCuentasAtencion.IdRegistroSeleccionado
        End If
    End If
    RecuperarDatosCuentaAtencion lIdCuentaAtencionActual

End Sub
Private Sub RecuperarDatosCuentaAtencion(lIdCuentaAtencion As Long)
Dim rsPaciente As New Recordset
Dim oDOPaciente As New doPaciente
Dim oDOCuentaAtencion As New DOCuentaAtencion
    
    
    'oDOPaciente.NroHistoriaClinica = Val(Me.cmbNroHistoriaBusqueda.Text)
    oDOCuentaAtencion.IdCuentaAtencion = lIdCuentaAtencion
    
    Screen.MousePointer = vbHourglass
    Set rsPaciente = mo_AdminAdmision.AtencionesFiltrarPacientesParaIngresarProcedimientos(oDOPaciente, oDOCuentaAtencion)
    Screen.MousePointer = vbDefault
    
    'cmbNroHistoriaBusqueda.BoundColumn = ""
    'Set cmbNroHistoriaBusqueda.ListSource = rsPaciente
    
    'Si hay una sola coincidencia
    If rsPaciente.RecordCount = 1 Then
        rsPaciente.MoveFirst
        LimpiarDatosDeAtencion
        
        Me.txtIdNroHistoria.Text = rsPaciente!NroHistoriaClinica
        mo_cmbIdTipoGenHistoriaClinica.BoundText = rsPaciente!IdTipoNumeracion
        Me.lblFechaIngreso = rsPaciente!FechaIngreso
        Me.lblServicioIngreso = rsPaciente!ServicioIngreso
        Me.lblPaciente = rsPaciente!ApellidoPaterno + " " + rsPaciente!ApellidoMaterno + " " + rsPaciente!PrimerNombre + " " + ("" & rsPaciente!SegundoNombre)
        Me.lblNroCuenta = rsPaciente!IdCuentaAtencion
        Me.txtNroOrden.SetFocus
    ElseIf rsPaciente.RecordCount > 1 Then
        'cmbNroHistoriaBusqueda.ShowDropDown
        
    ElseIf rsPaciente.RecordCount = 0 Then
        MsgBox "No se encontraron atenciones para el nro de historia o nro de cuenta ingresado", vbInformation, Me.Caption
        LimpiarDatosDeAtencion
    End If

End Sub

Private Sub cmbNroHistoria_Click()
End Sub
Sub LimpiarDatosDeAtencion()
        
        Me.txtIdNroHistoria.Text = ""
        mo_cmbIdTipoGenHistoriaClinica.BoundText = ""
        Me.lblFechaIngreso = ""
        Me.lblServicioIngreso = ""
        Me.lblPaciente = ""
        Me.lblNroCuenta = ""

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

Private Sub btnBusquedaMedico_Click()
       CompletarDatosDeMedico txtIdMedicoOrdena, lblDescMedicoOrdena
End Sub

Private Sub btnBusquedaServicio_Click()
    CompletarDatosDeServicio txtIdServicio, lblDescServicio
    Me.txtIdProcedimiento.SetFocus
End Sub

Private Sub btnQuitarDx_Click()
    
    Dim doFacturacionServicio As DOFacturacionServicios
    On Error Resume Next
    With mrs_ApoyoDiagnostico
        If Not .EOF And Not .BOF Then
            If mrs_ApoyoDiagnostico!IdAtencionApoyoDetalle <> 0 Then
                'Verificamos que el detalle esté como emitido para poder eliminarse
                Set doFacturacionServicio = mo_AdminFacturacion.FacturacionServiciosSeleccionarPorId(mrs_ApoyoDiagnostico!IdFacturacionServicio)
                If Not doFacturacionServicio Is Nothing Then
                    If Not (doFacturacionServicio.IdEstadoFacturacion = sghEstadoFacturacion.sghPendientePago And doFacturacionServicio.TotalPorPagar = 0) Then
                        If doFacturacionServicio.IdEstadoFacturacion = sghEstadoFacturacion.sghPendientePago Then
                            MsgBox "No se puede eliminar el item seleccionado por que ya se encuentra en proceso de facturación [Con un importe de S/. " & doFacturacionServicio.TotalPorPagar & " ]", vbExclamation, Me.Caption
                        Else
                            MsgBox "No se puede eliminar el item seleccionado por que ya se encuentra Facturado", vbExclamation, Me.Caption
                        End If
                        Exit Sub
                    End If
                End If
                mrs_ApoyoDiagnosticoEliminados.AddNew
                
                mrs_ApoyoDiagnosticoEliminados!IdAtencionProcDetalle = mrs_ApoyoDiagnostico!IdAtencionApoyoDetalle
                mrs_ApoyoDiagnosticoEliminados!IdFacturacionServicio = mrs_ApoyoDiagnostico!IdFacturacionServicio
            End If
           .Delete
           .Update
        End If
        .MoveFirst
    End With

End Sub

Private Sub cmbIdServicioRealiza_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdServicioRealiza
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = Me.cmbIdTipoGenHistoriaClinica
    Set mo_cmbIdServicioRealiza.MiComboBox = Me.cmbIdServicioRealiza
End Sub

Private Sub toolProcedimientos_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub txtFechaOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaOrden
End Sub
Private Sub txtFechaOrden_LostFocus()

       If txtFechaOrden <> SIGHComun.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaOrden, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, "Datos de paciente"
                 txtFechaOrden = SIGHComun.FECHA_VACIA_DMY
            End If
        End If
        
        mo_Formulario.MarcarComoVacio txtFechaOrden
End Sub

Private Sub txtFechaOrden_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtHoraOrden_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraOrden
End Sub


Private Sub txtHoraOrden_LostFocus()
    If txtHoraOrden <> SIGHComun.HORA_VACIA_HM Then
         If Not SIGHComun.ValidaHora(txtHoraOrden) Then
             MsgBox "La hora ingresada no es válida", vbInformation, "Datos de paciente"
             txtHoraOrden = SIGHComun.HORA_VACIA_HM
         End If
     End If
   mo_Formulario.MarcarComoVacio txtHoraOrden
End Sub

Private Sub txtHoraOrden_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdMedicoOrdena_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdMedicoOrdena
    If KeyCode = vbKeyF1 Then
        btnBusquedaMedico_Click
    End If
End Sub


Private Sub txtIdMedicoOrdena_LostFocus()
    CompletarDatosDeMedicoEnElLostFocus txtIdMedicoOrdena, lblDescMedicoOrdena
    mo_Formulario.MarcarComoVacio txtIdMedicoOrdena
End Sub

Private Sub txtIdMedicoOrdena_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
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
    End If
    
End Sub

Private Sub txtIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdServicio
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdServicio_LostFocus()
    CompletarDatosDeServicioEnElLostFocus txtIdServicio, lblDescServicio
    mo_Formulario.MarcarComoVacio txtIdServicio
End Sub

Private Sub txtIdServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNroHistoria_LostFocus()
   mo_Formulario.MarcarComoVacio txtNroHistoria
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
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

    mo_Formulario.HabilitarDeshabilitar lblNroCuenta, False
    mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    mo_Formulario.HabilitarDeshabilitar lblPaciente, False
    mo_Formulario.HabilitarDeshabilitar lblFechaIngreso, False
    mo_Formulario.HabilitarDeshabilitar lblServicioIngreso, False
    

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
            Me.fraBusqueda.Enabled = False
    Case sghConsultar
            Me.fraBusqueda.Enabled = False
            Me.fraProcedimiento.Enabled = False
            Me.fraExamen.Enabled = False
            Me.grdProcedimientos.Enabled = False
            'WCG comentado por facturacion
            'Me.ucProcedimientoDetalle1.BotonAgregarEnabled = False
            'Me.ucProcedimientoDetalle1.BotonQuitarEnabled = False
            Me.btnAceptar.Enabled = False
            
    Case sghEliminar
            Me.fraBusqueda.Enabled = False
            Me.fraProcedimiento.Enabled = False
            Me.fraExamen.Enabled = False
            Me.grdProcedimientos.Enabled = False
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
Dim sTitulo As String

        Select Case ml_IdDepartamento
        Case 7
            sTitulo = "Facturacion de exámenes (Patología Clínica)"
        Case 8
            sTitulo = "Facturacion de exámenes (Anatomía Patológica)"
        Case 9
            sTitulo = "Facturacion de exámenes (Imaginología)"
        End Select

       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar orden - " + sTitulo
       Case sghModificar
           Me.Caption = "Modificar orden - " + sTitulo
       Case sghConsultar
           Me.Caption = "Consultar orden - " + sTitulo
       Case sghEliminar
           Me.Caption = "Eliminar orden - " + sTitulo
       End Select

        GenerarRecordsetTemporal
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
            btnBuscar_Click
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
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
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
   
    If Me.lblNroCuenta = "" Then
        MsgBox "Seleccione el paciente", vbInformation, Me.Caption
        Exit Function
    End If
    
    If txtNroOrden = "" Then
        MsgBox "Ingrese el nro de orden de procedimiento", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    
    If txtIdMedicoOrdena = "" Then
        MsgBox "Ingrese el médico que ordena el procedimiento", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    
    If txtFechaOrden = SIGHComun.FECHA_VACIA_DMY Then
        MsgBox "Ingrese la fecha de orden del procedimiento", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    If Not SIGHComun.EsFecha(txtFechaOrden, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada para la orden no es válida", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    
    If txtHoraOrden = SIGHComun.HORA_VACIA_HM Then
        MsgBox "Ingrese la hora de orden del procedimiento", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    If Not SIGHComun.ValidaHora(txtHoraOrden) Then
        MsgBox "La hora ingresada para la orden no es válida", vbInformation, "Validación de órdenes"
        Exit Function
    End If
    
    'Validamos que existan detalles
    Dim bFound  As Boolean
    bFound = False
    If mrs_ApoyoDiagnostico.EOF = False And mrs_ApoyoDiagnostico.BOF = False Then
        mrs_ApoyoDiagnostico.MoveFirst
        Do Until mrs_ApoyoDiagnostico.EOF
            If mrs_ApoyoDiagnostico.Fields!IdProcedimiento <> 0 Then
                bFound = True
                Exit Do
            End If
        Loop
    End If
    If Not bFound Then
        MsgBox "Ingrese los examenes de la orden", vbInformation, "Validación de órdenes"
        Exit Function
    End If
   
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   
    If txtFechaOrden < CDate(Me.lblFechaIngreso) Then
        MsgBox "La fecha de la orden del procedimiento no puede ser menor que la fecha de ingreso de la atención", vbExclamation, Me.Caption
        Exit Function
    End If
   
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

    'WCG Comentado por facturacion
   mo_ApoyoDiagnostico.IdCuentaAtencion = Val(Me.lblNroCuenta)
   
   CargarProcedimientosAlObjetoDatos mo_ApoyoDiagnostico, mo_ApoyoDiagDetalle
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminFacturacion.AtencionApoyoDxAgregar(mo_ApoyoDiagnostico, mo_ApoyoDiagDetalle)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminFacturacion.AtencionApoyoDxModificar(mo_ApoyoDiagnostico, mo_ApoyoDiagDetalle, mrs_ApoyoDiagnosticoEliminados)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminFacturacion.AtencionApoyoDxEliminar(mo_ApoyoDiagnostico, mo_ApoyoDiagDetalle)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
    
    '1ro
    Dim oDOAtencionApoyoDiagnostico As New DOAtencionApoyoDiagnostico
    Set oDOAtencionApoyoDiagnostico = mo_AdminFacturacion.AtencionApoyoDxSeleccionarPorId(Me.IdAtencionApoyoDiagnostico)
    If Not oDOAtencionApoyoDiagnostico Is Nothing Then
        CargarDatosDelaAtencion oDOAtencionApoyoDiagnostico.IdCuentaAtencion
        mb_ExistenDatos = True
    Else
        mb_ExistenDatos = False
    End If
    
    '2do
    CargarDatosDeDeProcedimientos
   
End Sub
Sub CargarDatosDelaAtencion(lIdCuentaAtencion As Long)
Dim oDOPaciente As New doPaciente
Dim rsPaciente As New ADODB.Recordset
Dim oDOCuentaAtencion As New DOCuentaAtencion
    
    oDOPaciente.NroHistoriaClinica = 0
    oDOCuentaAtencion.IdCuentaAtencion = lIdCuentaAtencion
    Set rsPaciente = mo_AdminAdmision.AtencionesFiltrarPacientesParaIngresarProcedimientos(oDOPaciente, oDOCuentaAtencion)
    
    'Si hay una sola coincidencia
    If Not (rsPaciente.EOF And rsPaciente.BOF) Then
        LimpiarDatosDeAtencion
        Me.txtIdNroHistoria.Text = rsPaciente!NroHistoriaClinica
        mo_cmbIdTipoGenHistoriaClinica.BoundText = rsPaciente!IdTipoNumeracion
        Me.lblFechaIngreso = rsPaciente!FechaIngreso
        Me.lblServicioIngreso = rsPaciente!ServicioIngreso
        Me.lblPaciente = rsPaciente!ApellidoPaterno + " " + rsPaciente!ApellidoMaterno + " " + rsPaciente!PrimerNombre + " " + ("" & rsPaciente!SegundoNombre)
        Me.lblNroCuenta = rsPaciente!IdCuentaAtencion
    End If
    rsPaciente.Close

End Sub
'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla AtencionesInterconsultas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()
   
End Sub
Sub GenerarRecordsetTemporal()
    
    With mrs_ApoyoDiagnostico
        .Fields.Append "IdAtencionApoyoDetalle", adInteger
        .Fields.Append "IdProcedimiento", adInteger
        .Fields.Append "CodigoCPT", adVarChar, 10
        .Fields.Append "Descripcion", adVarChar, 255
        .Fields.Append "FechaResultado", adChar, 10
        .Fields.Append "HoraResultado", adChar, 5
        '.Fields.Append "IdMedicoRealiza", adInteger, , adFldIsNullable
        '.Fields.Append "NombreMedico", adVarChar, 100, adFldIsNullable
        .Fields.Append "IdServicioRealiza", adInteger, , adFldIsNullable
        .Fields.Append "NombreServicio", adVarChar, 100, adFldIsNullable
        .Fields.Append "IdFacturacionServicio", adInteger, , adFldIsNullable
        .Fields.Append "EstadoRegistro", adChar, 1
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    
    'Para los procedimientos eliminados
    With mrs_ApoyoDiagnosticoEliminados
        .Fields.Append "IdAtencionProcDetalle", adInteger
        .Fields.Append "IdFacturacionServicio", adInteger, , adFldIsNullable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    Set grdProcedimientos.DataSource = mrs_ApoyoDiagnostico
    
End Sub

Public Sub CargarDatosDeDeProcedimientos()
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oDOAtencionApoyoDiagnostico As New DOAtencionApoyoDiagnostico

    'Carga datos de la cabecera
    Dim rsProcedimiento As New Recordset
    Set oDOAtencionApoyoDiagnostico = mo_AdminFacturacion.AtencionApoyoDxSeleccionarPorId(Me.IdAtencionApoyoDiagnostico)
    
    If oDOAtencionApoyoDiagnostico.IdAtencionApoyoDx = 0 Then
        MsgBox "No existe datos de procedimientos", vbInformation, Me.Caption
        Exit Sub
    End If
    
    txtFechaOrden = oDOAtencionApoyoDiagnostico.FechaOrden
    txtHoraOrden = oDOAtencionApoyoDiagnostico.HoraOrden
    txtNroOrden = oDOAtencionApoyoDiagnostico.OrdenNro

    'Completa datos de medico
    If mo_AdminProgramacion.MedicosSeleccionarPorId(oDOAtencionApoyoDiagnostico.IdMedicoOrdena, oDoMedico, oDOEmpleado, oDOEspecialidades) Then
        txtIdMedicoOrdena.Text = oDOEmpleado.CodigoPlanilla
        txtIdMedicoOrdena.Tag = oDoMedico.IdMedico
        lblDescMedicoOrdena = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    End If
    
    Me.txtIdServicio.Tag = IIf(oDOAtencionApoyoDiagnostico.IdServicioOrdena = 0, "", oDOAtencionApoyoDiagnostico.IdServicioOrdena)
    Dim oDOServicio As New DOServicio
    If Me.txtIdServicio.Tag <> "" Then
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oDOAtencionApoyoDiagnostico.IdServicioOrdena)
        If Not oDOServicio Is Nothing Then
            Me.txtIdServicio.Text = oDOServicio.Codigo
            Me.lblDescServicio = oDOServicio.Nombre
        End If
    End If
  
    Dim rsProcedimientos As New Recordset
    Set rsProcedimientos = mo_AdminFacturacion.AtencionApoyoDxDetalleSeleccionarPorIdAtencionApoyoDx(Me.IdAtencionApoyoDiagnostico)
    Do While Not rsProcedimientos.EOF
        With mrs_ApoyoDiagnostico
            .AddNew
            .Fields!IdAtencionApoyoDetalle = rsProcedimientos!IdAtencionApoyoDetalle
            .Fields!IdProcedimiento = rsProcedimientos!IdProcedimiento
            .Fields!CodigoCPT = rsProcedimientos!CodigoCPT
            .Fields!Descripcion = rsProcedimientos!Descripcion
            '.Fields!IdMedicoRealiza = rsProcedimientos!IdMedicoRealiza
            '.Fields!NombreMedico = rsProcedimientos!NombreMedico
            .Fields!IdServicioRealiza = rsProcedimientos!IdServicioRealiza
            .Fields!NombreServicio = rsProcedimientos!NombreServicio
            .Fields!FechaResultado = Format(rsProcedimientos!FechaResultado, "dd/mm/yyyy")
            .Fields!HoraResultado = rsProcedimientos!HoraResultado
            .Fields!IdFacturacionServicio = rsProcedimientos!IdFacturacionServicio
            .Fields!EstadoRegistro = "M"
        End With
        rsProcedimientos.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores grdProcedimientos, SIGHComun.GrillaConFilasBicolor
    
End Sub

Sub CargarProcedimientosAlObjetoDatos(oAtencionApoyoDiagnostico As DOAtencionApoyoDiagnostico, oAtencionApoyoDiagnosticoDetalle As Collection)
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS ProcedimientoS
    '---------------------------------------------------------------------------------
    'Datos de la cabecera
    oAtencionApoyoDiagnostico.IdAtencionApoyoDx = Me.IdAtencionApoyoDiagnostico
    oAtencionApoyoDiagnostico.IdCuentaAtencion = Val(lblNroCuenta)
    oAtencionApoyoDiagnostico.IdMedicoOrdena = Val(txtIdMedicoOrdena.Tag)
    oAtencionApoyoDiagnostico.IdServicioOrdena = Val(Me.txtIdServicio.Tag)
    oAtencionApoyoDiagnostico.FechaOrden = txtFechaOrden.Text
    oAtencionApoyoDiagnostico.HoraOrden = txtHoraOrden.Text
    oAtencionApoyoDiagnostico.OrdenNro = txtNroOrden.Text
    oAtencionApoyoDiagnostico.IdUsuarioAuditoria = ml_IdUsuario
    
    'Datos del detalle
    Dim oFacturacionProcDetalle As DOAtencionApoyoDiagDetalle
    If Not (mrs_ApoyoDiagnostico.BOF And mrs_ApoyoDiagnostico.EOF) Then
        Set oFacturacionProcDetalle = New DOAtencionApoyoDiagDetalle
        mrs_ApoyoDiagnostico.MoveFirst
        Do While Not mrs_ApoyoDiagnostico.EOF
            Set oFacturacionProcDetalle = New DOAtencionApoyoDiagDetalle
            
            oFacturacionProcDetalle.IdAtencionApoyoDetalle = mrs_ApoyoDiagnostico!IdAtencionApoyoDetalle
            oFacturacionProcDetalle.IdAtencionApoyoDx = Me.IdAtencionApoyoDiagnostico
            oFacturacionProcDetalle.FechaResultado = IIf(Trim(mrs_ApoyoDiagnostico!FechaResultado) <> "__/__/____" And Trim(mrs_ApoyoDiagnostico!FechaResultado) <> "", mrs_ApoyoDiagnostico!FechaResultado, 0)
            oFacturacionProcDetalle.HoraResultado = mrs_ApoyoDiagnostico!HoraResultado
            'oFacturacionProcDetalle.IdMedicoRealiza = IIf(IsNull(mrs_ApoyoDiagnostico!IdMedicoRealiza), 0, mrs_ApoyoDiagnostico!IdMedicoRealiza)
            oFacturacionProcDetalle.IdProcedimiento = mrs_ApoyoDiagnostico!IdProcedimiento
            oFacturacionProcDetalle.IdServicioRealiza = IIf(IsNull(mrs_ApoyoDiagnostico!IdServicioRealiza), 0, mrs_ApoyoDiagnostico!IdServicioRealiza)
            oFacturacionProcDetalle.IdUsuarioAuditoria = ml_IdUsuario
            oFacturacionProcDetalle.IdFacturacionServicio = IIf(IsNull(mrs_ApoyoDiagnostico!IdFacturacionServicio), 0, mrs_ApoyoDiagnostico!IdFacturacionServicio)
            oFacturacionProcDetalle.EstadoRegistro = mrs_ApoyoDiagnostico!EstadoRegistro
            oAtencionApoyoDiagnosticoDetalle.Add oFacturacionProcDetalle
            mrs_ApoyoDiagnostico.MoveNext
        Loop
    End If
    
End Sub

Private Sub grdProcedimientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdProcedimientos.Bands(0).Columns("IdAtencionApoyoDetalle").Hidden = True
    grdProcedimientos.Bands(0).Columns("IdProcedimiento").Hidden = True
    grdProcedimientos.Bands(0).Columns("IdFacturacionServicio").Hidden = True
    'grdProcedimientos.Bands(0).Columns("IdMedicoRealiza").Hidden = True
    grdProcedimientos.Bands(0).Columns("IdServicioRealiza").Hidden = True
    
    grdProcedimientos.Bands(0).Columns("CodigoCPT").Header.Caption = "CPT"
    grdProcedimientos.Bands(0).Columns("CodigoCPT").Width = 1000
    
    'grdProcedimientos.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdProcedimientos.Bands(0).Columns("Descripcion").Width = 7500
    
    'grdProcedimientos.Bands(0).Columns("FechaRealizacion").Header.Caption = "Fecha"
    grdProcedimientos.Bands(0).Columns("FechaResultado").Hidden = True
    
    'grdProcedimientos.Bands(0).Columns("HoraRealizacion").Header.Caption = "Hora"
    grdProcedimientos.Bands(0).Columns("HoraResultado").Hidden = True
    
    'grdProcedimientos.Bands(0).Columns("NombreMedico").Header.Caption = "Médico"
    'grdProcedimientos.Bands(0).Columns("NombreMedico").Hidden = True

    grdProcedimientos.Bands(0).Columns("NombreServicio").Header.Caption = "Servicio"
    grdProcedimientos.Bands(0).Columns("NombreServicio").Width = 2500
    'grdProcedimientos.Bands(0).Columns("NombreServicio").Hidden = True
    grdProcedimientos.Bands(0).Columns("EstadoRegistro").Hidden = True


End Sub
Private Sub btnBusquedaProcedimiento_Click()
Dim oBusqueda As New ProcedimientosBusqueda
Dim oDOProcedimiento As DOProcedimiento
    
    oBusqueda.IdDiferenciacion = ml_IdDepartamento
    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOProcedimiento = mo_AdminServiciosComunes.ProcedimientosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOProcedimiento Is Nothing Then
            If oDOProcedimiento.IdProducto = 0 Then
                MsgBox "El examen ingresado no se encuentra en el catálogo de servicios (Tarifario)", vbInformation, Me.Caption
                txtIdProcedimiento.Tag = ""
                txtIdProcedimiento.Text = ""
                lblDescProcedimiento = ""
            Else
                txtIdProcedimiento.Text = oDOProcedimiento.CodigoCPT2004
                txtIdProcedimiento.Tag = oDOProcedimiento.IdProcedimiento
                lblDescProcedimiento = oDOProcedimiento.Descripcion
            End If
        Else
            txtIdProcedimiento.Text = ""
            txtIdProcedimiento.Tag = ""
            lblDescProcedimiento = ""
        End If
    End If
    
End Sub

Private Sub txtIdProcedimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdProcedimiento
End Sub

Private Sub txtIdProcedimiento_LostFocus()

    txtIdProcedimiento.Text = UCase(txtIdProcedimiento.Text)

   If txtIdProcedimiento.Text <> "" Then
    Dim oDOProcedimiento As DOProcedimiento
        Set oDOProcedimiento = mo_AdminServiciosComunes.ProcedimientosSeleccionarPorCodigoCPT(txtIdProcedimiento.Text)
        If Not oDOProcedimiento Is Nothing Then
            If oDOProcedimiento.IdProducto = 0 Then
                MsgBox "El examen ingresado no se encuentra en el catálogo de servicios (Tarifario)", vbInformation, Me.Caption
                txtIdProcedimiento.Tag = ""
                txtIdProcedimiento.Text = ""
                lblDescProcedimiento = ""
            Else
                txtIdProcedimiento.Tag = oDOProcedimiento.IdProcedimiento
                lblDescProcedimiento = oDOProcedimiento.Descripcion
            End If
        Else
            txtIdProcedimiento.Tag = ""
            lblDescProcedimiento = ""
        End If
    Else
        txtIdProcedimiento.Tag = ""
        lblDescProcedimiento = ""
   End If
   
   'mo_Formulario.MarcarComoVacio txtIdProcedimiento
End Sub

Private Sub txtIdProcedimiento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtIdProcedimiento_LostFocus
        'btnAgregarDx_Click
        Exit Sub
    End If
    
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Sub CompletarDatosDeServicio(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
Dim oBusqueda As New ServiciosBusqueda
Dim oDOServicio As New DOServicio

    oBusqueda.IdTipoServicio = 0
    oBusqueda.HabilitarTipoServicio = True
    oBusqueda.Show 1
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOServicio Is Nothing Then
            txtIdServicio.Text = oDOServicio.Codigo
            txtIdServicio.Tag = oDOServicio.IdServicio
            lblDescripcionServicio = oDOServicio.Nombre
        End If
    End If

End Sub

Sub CompletarDatosDeServicioEnElLostFocus(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
    
    txtIdServicio.Text = UCase(txtIdServicio.Text)
    If txtIdServicio.Text <> "" Then
        Dim oDOServicio As DOServicio
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(txtIdServicio.Text)
        If Not oDOServicio Is Nothing Then
            txtIdServicio.Tag = oDOServicio.IdServicio
            lblDescripcionServicio = oDOServicio.Nombre
        Else
            txtIdServicio.Tag = ""
            lblDescripcionServicio = ""
        End If
   End If

End Sub

Private Sub txtNroOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroOrden
    AdministrarKeyPreview KeyCode
End Sub
