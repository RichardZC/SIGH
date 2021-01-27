VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form EmpleadoDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EmpleadoDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   12150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   2532
      Left            =   60
      TabIndex        =   43
      Top             =   5160
      Width           =   6390
      Begin VB.CheckBox chkAutorizadoReniec 
         Alignment       =   1  'Right Justify
         Caption         =   "Autorizado por RENIEC "
         Height          =   585
         Left            =   180
         TabIndex        =   53
         Top             =   1860
         Width           =   1695
      End
      Begin VB.TextBox txtCodigoHisDelDigitador 
         Height          =   315
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   14
         Top             =   1380
         Width           =   1590
      End
      Begin VB.TextBox txtUsuario 
         Height          =   315
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   12
         Top             =   210
         Width           =   1590
      End
      Begin VB.TextBox txtClave 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   570
         Width           =   1575
      End
      Begin VB.CheckBox chkUsaGalenHos 
         Alignment       =   1  'Right Justify
         Caption         =   "Esta usando Galenhos"
         Height          =   495
         Left            =   180
         TabIndex        =   15
         Top             =   870
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Codigo Digitador (HIS)"
         Height          =   435
         Left            =   180
         TabIndex        =   49
         Top             =   1410
         Width           =   1500
      End
      Begin VB.Label lblNombrePc 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1980
         TabIndex        =   46
         Top             =   990
         Width           =   4290
      End
      Begin VB.Label Label3 
         Caption         =   "Usuario"
         Height          =   315
         Left            =   180
         TabIndex        =   45
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   "Clave"
         Height          =   315
         Left            =   180
         TabIndex        =   44
         Top             =   615
         Width           =   1260
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Labora en:"
      Height          =   2784
      Left            =   6510
      TabIndex        =   42
      Top             =   4905
      Width           =   5550
      Begin VB.CommandButton cmdAddLabora 
         DisabledPicture =   "EmpleadoDetalle.frx":0CCA
         DownPicture     =   "EmpleadoDetalle.frx":10B3
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         Picture         =   "EmpleadoDetalle.frx":14BF
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   615
         Width           =   795
      End
      Begin VB.CommandButton cmdDelLabora 
         DisabledPicture =   "EmpleadoDetalle.frx":18CB
         DownPicture     =   "EmpleadoDetalle.frx":1C56
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Picture         =   "EmpleadoDetalle.frx":1FE9
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   615
         Width           =   795
      End
      Begin VB.ComboBox cmbArea 
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
         ItemData        =   "EmpleadoDetalle.frx":237A
         Left            =   120
         List            =   "EmpleadoDetalle.frx":239C
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   255
         Width           =   2775
      End
      Begin MSDataListLib.DataCombo cmbSubArea 
         Height          =   315
         Left            =   2910
         TabIndex        =   25
         Top             =   255
         Visible         =   0   'False
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UltraGrid.SSUltraGrid grdLaboraEn 
         Height          =   1650
         Left            =   75
         TabIndex        =   28
         Top             =   1020
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   2910
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
         Caption         =   "Labora en:"
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cargos"
      Height          =   2388
      Left            =   6510
      TabIndex        =   40
      Top             =   2445
      Width           =   5565
      Begin VB.CommandButton btnQuitaCargo 
         DisabledPicture =   "EmpleadoDetalle.frx":249E
         DownPicture     =   "EmpleadoDetalle.frx":2829
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4635
         Picture         =   "EmpleadoDetalle.frx":2BBC
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   255
         Width           =   795
      End
      Begin VB.CommandButton btnAgreaCargo 
         DisabledPicture =   "EmpleadoDetalle.frx":2F4D
         DownPicture     =   "EmpleadoDetalle.frx":3336
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3825
         Picture         =   "EmpleadoDetalle.frx":3742
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   795
      End
      Begin VB.ComboBox cmbCargos 
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
         Left            =   120
         TabIndex        =   20
         Top             =   270
         Width           =   3705
      End
      Begin UltraGrid.SSUltraGrid grdCargos 
         Height          =   1620
         Left            =   120
         TabIndex        =   23
         Top             =   660
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   2858
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
         Caption         =   "Cargos asignados"
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Roles"
      Height          =   2220
      Left            =   6510
      TabIndex        =   39
      Top             =   180
      Width           =   5565
      Begin VB.ComboBox cmbIdRol 
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
         Left            =   60
         TabIndex        =   16
         Top             =   210
         Width           =   3690
      End
      Begin VB.CommandButton btnAgregarDx 
         DisabledPicture =   "EmpleadoDetalle.frx":3B4E
         DownPicture     =   "EmpleadoDetalle.frx":3F37
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3765
         Picture         =   "EmpleadoDetalle.frx":4343
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   195
         Width           =   825
      End
      Begin VB.CommandButton btnQuitarDx 
         DisabledPicture =   "EmpleadoDetalle.frx":474F
         DownPicture     =   "EmpleadoDetalle.frx":4ADA
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4620
         Picture         =   "EmpleadoDetalle.frx":4E6D
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   195
         Width           =   825
      End
      Begin UltraGrid.SSUltraGrid grdRoles 
         Height          =   1425
         Left            =   60
         TabIndex        =   19
         Top             =   705
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   2514
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
         Caption         =   "Roles asignados"
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1044
      Left            =   60
      TabIndex        =   37
      Top             =   7740
      Width           =   12030
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "EmpleadoDetalle.frx":51FE
         DownPicture     =   "EmpleadoDetalle.frx":565E
         Height          =   700
         Left            =   4582
         Picture         =   "EmpleadoDetalle.frx":5AD3
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "EmpleadoDetalle.frx":5F48
         DownPicture     =   "EmpleadoDetalle.frx":640C
         Height          =   700
         Left            =   6127
         Picture         =   "EmpleadoDetalle.frx":68F8
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5010
      Left            =   60
      TabIndex        =   31
      Top             =   150
      Width           =   6375
      Begin VB.ComboBox cmbIdNacionalidad 
         Height          =   330
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   1650
         Width           =   4605
      End
      Begin VB.ComboBox cmbIdTipoSexo 
         Height          =   330
         Left            =   3810
         TabIndex        =   6
         Top             =   2040
         Width           =   2475
      End
      Begin VB.CheckBox chkEsActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo"
         Height          =   435
         Left            =   195
         TabIndex        =   59
         Top             =   4650
         Width           =   1650
      End
      Begin VB.CommandButton cmdSupervisorAdd 
         DisabledPicture =   "EmpleadoDetalle.frx":6DE4
         DownPicture     =   "EmpleadoDetalle.frx":71CD
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5370
         Picture         =   "EmpleadoDetalle.frx":75D9
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4245
         Width           =   435
      End
      Begin VB.CommandButton cmdSupervisorDel 
         DisabledPicture =   "EmpleadoDetalle.frx":79E5
         DownPicture     =   "EmpleadoDetalle.frx":7D70
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5835
         Picture         =   "EmpleadoDetalle.frx":8103
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   4245
         Width           =   435
      End
      Begin VB.TextBox txtSupervisor 
         Height          =   315
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   56
         Top             =   4260
         Width           =   3675
      End
      Begin VB.ComboBox cmbIdDocIdentidad 
         Height          =   330
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   600
         Width           =   1905
      End
      Begin VB.TextBox txtEstablecimientoExterno 
         Height          =   345
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   51
         Top             =   3570
         Width           =   3675
      End
      Begin VB.CommandButton BtnDelCsPS 
         DisabledPicture =   "EmpleadoDetalle.frx":8494
         DownPicture     =   "EmpleadoDetalle.frx":881F
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5835
         Picture         =   "EmpleadoDetalle.frx":8BB2
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3570
         Width           =   435
      End
      Begin VB.CommandButton BtnAdicionarCsPs 
         DisabledPicture =   "EmpleadoDetalle.frx":8F43
         DownPicture     =   "EmpleadoDetalle.frx":932C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5370
         Picture         =   "EmpleadoDetalle.frx":9738
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3585
         Width           =   435
      End
      Begin VB.ComboBox cmbTipoDestacado 
         Height          =   330
         Left            =   1665
         TabIndex        =   9
         Top             =   3150
         Width           =   4605
      End
      Begin VB.TextBox txtDNI 
         Height          =   315
         Left            =   3555
         MaxLength       =   8
         TabIndex        =   1
         Top             =   600
         Width           =   2730
      End
      Begin VB.ComboBox cmbIdCondicionTrabajo 
         Height          =   330
         Left            =   1665
         TabIndex        =   8
         Top             =   2760
         Width           =   4620
      End
      Begin VB.ComboBox cmbIdTipoEmpleado 
         Height          =   330
         Left            =   1665
         TabIndex        =   7
         Top             =   2400
         Width           =   4620
      End
      Begin VB.TextBox txtCodigoPlanilla 
         Height          =   315
         Left            =   1665
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1590
      End
      Begin VB.TextBox txtApellidoPaterno 
         Height          =   315
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   2
         Top             =   960
         Width           =   1590
      End
      Begin VB.TextBox txtApellidoMaterno 
         Height          =   315
         Left            =   4725
         MaxLength       =   50
         TabIndex        =   3
         Top             =   915
         Width           =   1575
      End
      Begin VB.TextBox txtNombres 
         Height          =   315
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1305
         Width           =   4620
      End
      Begin MSMask.MaskEdBox txtFechaNacimiento 
         Height          =   315
         Left            =   1650
         TabIndex        =   5
         Top             =   2040
         Width           =   1560
         _ExtentX        =   2752
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
      Begin VB.Label lblNacionalidad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nacionalidad"
         Height          =   210
         Left            =   195
         TabIndex        =   61
         Top             =   1725
         Width           =   990
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3375
         TabIndex        =   60
         Top             =   2085
         Width           =   405
      End
      Begin VB.Label Label8 
         Caption         =   "Supervisor (HisGalenhos)"
         Height          =   435
         Left            =   195
         TabIndex        =   57
         Top             =   4260
         Width           =   1500
      End
      Begin VB.Label lblId 
         Alignment       =   1  'Right Justify
         Caption         =   "...."
         Height          =   285
         Left            =   5025
         TabIndex        =   54
         Top             =   195
         Width           =   1245
      End
      Begin VB.Label Label9 
         Caption         =   "Cs,Ps externo donde labora (HisGalenhos)"
         Height          =   615
         Left            =   195
         TabIndex        =   52
         Top             =   3540
         Width           =   1470
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo Destacado"
         Height          =   315
         Left            =   195
         TabIndex        =   48
         Top             =   3210
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Nacimiento"
         Height          =   315
         Left            =   195
         TabIndex        =   47
         Top             =   2070
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         Height          =   315
         Left            =   195
         TabIndex        =   41
         Top             =   645
         Width           =   1170
      End
      Begin VB.Label Label2 
         Caption         =   "Código planilla"
         Height          =   285
         Left            =   195
         TabIndex        =   38
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label lblApellidoPaterno 
         Caption         =   "Apellido paterno"
         Height          =   315
         Left            =   195
         TabIndex        =   36
         Top             =   990
         Width           =   1500
      End
      Begin VB.Label lblApellidoMaterno 
         Caption         =   "Apellido materno"
         Height          =   315
         Left            =   3270
         TabIndex        =   35
         Top             =   960
         Width           =   1470
      End
      Begin VB.Label lblNombres 
         Caption         =   "Nombres"
         Height          =   315
         Left            =   195
         TabIndex        =   34
         Top             =   1365
         Width           =   1500
      End
      Begin VB.Label lblIdCondicionTrabajo 
         Caption         =   "Condición trabajo"
         Height          =   315
         Left            =   195
         TabIndex        =   33
         Top             =   2820
         Width           =   1500
      End
      Begin VB.Label lblIdTipoEmpleado 
         Caption         =   "Tipo empleado"
         Height          =   315
         Left            =   195
         TabIndex        =   32
         Top             =   2460
         Width           =   1500
      End
   End
End
Attribute VB_Name = "EmpleadoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Empleados
'        Programado por: Barrantes D
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------
Dim lbHuboCambioEnDato As Boolean
Dim ml_IdEmpleado As Long
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim mo_AdminSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_Empleado As New dOEmpleado
Dim mo_CmbIdTipoSexo As New sighentidades.ListaDespleglable
Dim mo_cmbIdTipoEmpleado As New sighentidades.ListaDespleglable
Dim mo_cmbIdCondicionTrabajo  As New sighentidades.ListaDespleglable
Dim mo_cmbIdRol As New sighentidades.ListaDespleglable
Dim mo_cmbCargos As New sighentidades.ListaDespleglable
Dim mo_cmbTipoDestacado As New sighentidades.ListaDespleglable
Dim mo_cmbIdDocIdentidad As New sighentidades.ListaDespleglable
Dim mo_UsuarioRoles As New Collection
Dim mrs_UsuariosRoles As New ADODB.Recordset
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim oRsFormaPago As New ADODB.Recordset
Dim oRsAlmacen As New ADODB.Recordset
Dim mrs_UsuariosCargos As New ADODB.Recordset
Dim mrs_LaboraEn As New ADODB.Recordset
Dim oRsSubArea As New Recordset
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Const lnDefaultEmpleado As Long = 738
Dim lnIdEstablecimientoExterno As Long
'--GLCC Agregar objeto cmbNacionaliudad - Cambio5 Inicio
Dim mo_cmbIdNacionalidad As New sighentidades.ListaDespleglable
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
'<(Inicio) Añadido Por: WABG el: 26/01/2021-12:07:00 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
Dim mo_Reniec As New ReniecGalenhosNegocios
Dim lbBuscaDNIenReniec As Boolean
Dim lcBuscaParametro As New SIGHDatos.Parametros
'</(Fin) Añadido Por: WABG el: 26/01/2021-12:07:00 p.m. en el Equipo: SISGALENPLUS-PC<CAMBIO-37>
'Dim mo_cmbIdNacionalidad As New sighentidades.ListaDespleglable
'GLCC Agregar objeto cmbNacionaliudad - Cambio 5 Fin
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Sub CargarComboBoxes()
Dim sSQL As String
Dim oConexion As New ADODB.Connection

       oConexion.Open sighentidades.CadenaConexion
       oConexion.CursorLocation = adUseClient
       
       mo_CmbIdTipoSexo.BoundColumn = "IdtipoSexo"
       mo_CmbIdTipoSexo.ListField = "DescripcionLarga"
       Set mo_CmbIdTipoSexo.RowSource = mo_AdminServiciosComunes.TiposSexoSeleccionarTodos()
       
       
       mo_cmbIdTipoEmpleado.BoundColumn = "IdTipoEmpleado"
       mo_cmbIdTipoEmpleado.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoEmpleado.RowSource = mo_AdminServiciosComunes.TiposEmpleadosSeleccionarSegunFiltro("")
       
       mo_cmbIdCondicionTrabajo.BoundColumn = "IdCondicionTrabajo"
       mo_cmbIdCondicionTrabajo.ListField = "DescripcionLarga"
       Set mo_cmbIdCondicionTrabajo.RowSource = mo_AdminServiciosComunes.TiposCondicionTrabajoSeleccionarTodos

       mo_cmbIdRol.BoundColumn = "IdRol"
       mo_cmbIdRol.ListField = "Nombre"
       Set mo_cmbIdRol.RowSource = mo_AdminSeguridad.RolesSeleccionarTodos()
     
       mo_cmbCargos.BoundColumn = "IdTipoCargo"
       mo_cmbCargos.ListField = "Cargo"
       Set mo_cmbCargos.RowSource = mo_ReglasFarmacia.TiposCargoSeleccionarTodos
       
       mo_cmbTipoDestacado.BoundColumn = "idDestacado"
       mo_cmbTipoDestacado.ListField = "Destacado"
       Set mo_cmbTipoDestacado.RowSource = mo_AdminServiciosComunes.TiposDestacadosSeleccionarTodos()
       mo_cmbTipoDestacado.BoundText = "3"

       mo_cmbIdDocIdentidad.BoundColumn = "IdDocIdentidad"
       mo_cmbIdDocIdentidad.ListField = "DescripcionLarga"
       Set mo_cmbIdDocIdentidad.RowSource = mo_AdminServiciosComunes.TiposDocIdentidadSeleccionarTodos()
       mo_cmbIdDocIdentidad.BoundText = "1"
       
       'GLCC- Cargar combo nacionalidad-06/07/2020
       mo_cmbIdNacionalidad.BoundColumn = "IdPais"
       mo_cmbIdNacionalidad.ListField = "Nombre"
       Set mo_cmbIdNacionalidad.RowSource = mo_AdminServiciosGeograficos.PaisesSeleccionarTodos()
       mo_cmbIdNacionalidad.BoundText = "166"
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
Property Let IdEmpleado(lValue As Long)
   ml_IdEmpleado = lValue
End Property
Property Get IdEmpleado() As Long
   IdEmpleado = ml_IdEmpleado
End Property
Private Sub BtnAdicionarCsPs_Click()
        'Dim oBusqueda As New EstablecimientosBusqueda
        Dim oBusqueda As New SIGHNegocios.BuscaEstablecimientos
        
        oBusqueda.MostrarFormulario
        If oBusqueda.BotonPresionado = sghAceptar Then
           lnIdEstablecimientoExterno = oBusqueda.idRegistroSeleccionado
           CargaNombreEstablecimiento
        End If
        Set oBusqueda = Nothing
End Sub
Private Sub btnAgreaCargo_Click()
    
    If mo_cmbCargos.BoundText = "" Then
        MsgBox "Ingrese el Cargo", vbInformation, Me.Caption
        Exit Sub
    End If
    
    On Error Resume Next
    mrs_UsuariosCargos.MoveFirst
    Do While Not mrs_UsuariosCargos.EOF
        If mo_cmbCargos.BoundText = mrs_UsuariosCargos!IdTipoCargo Then
            MsgBox "El CARGO ya fué asignado", vbExclamation, Me.Caption
            Exit Sub
        End If
        mrs_UsuariosCargos.MoveNext
    Loop
    
    With mrs_UsuariosCargos
        .AddNew
        .Fields!IdTipoCargo = mo_cmbCargos.BoundText
        .Fields!cargo = cmbCargos.Text
        .Update
    End With
End Sub

Private Sub BtnDelCsPS_Click()
    lnIdEstablecimientoExterno = 0
    CargaNombreEstablecimiento
End Sub

Private Sub btnQuitaCargo_Click()
    On Error Resume Next
    With mrs_UsuariosCargos
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With
End Sub
Private Sub cmbArea_Click()
    On Error GoTo ErrCombo
    Dim lcSql As String
    cmbSubArea.Text = ""
    Select Case cmbArea.ListIndex
    Case 0    'En otro Lugar
        Set cmbSubArea.RowSource = Nothing
        cmbSubArea.Visible = False
    Case sghAreasLaboraEmpleado.sghAlmacenFarmacia
        Set oRsSubArea = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales<>'X' and idEstado=1")
        Set cmbSubArea.RowSource = oRsSubArea
        cmbSubArea.ListField = "Descripcion"
        cmbSubArea.BoundColumn = "idAlmacen"
        cmbSubArea.Visible = True
    Case sghAreasLaboraEmpleado.sghImageneología
        Set oRsSubArea = mo_AdminServiciosComunes.FactPuntosCargaSeleccionarPorFiltro("TipoPunto='I'")
        Set cmbSubArea.RowSource = oRsSubArea
        cmbSubArea.ListField = "Descripcion"
        cmbSubArea.BoundColumn = "IdPuntoCarga"
        cmbSubArea.Visible = True
    Case sghAreasLaboraEmpleado.sghLaboratorio
        Set oRsSubArea = mo_AdminServiciosComunes.FactPuntosCargaSeleccionarPorFiltro("TipoPunto='L' or idPuntoCarga=2 or idPuntoCarga=3 or idPuntoCarga=11")
        Set cmbSubArea.RowSource = oRsSubArea
        cmbSubArea.ListField = "Descripcion"
        cmbSubArea.BoundColumn = "IdPuntoCarga"
        cmbSubArea.Visible = True
    Case sghAreasLaboraEmpleado.sghSeguros
        Set oRsSubArea = mo_AdminServiciosComunes.TiposFinanciamientoSegunFiltro("esOficina=1")
        Set cmbSubArea.RowSource = oRsSubArea
        cmbSubArea.ListField = "Descripcion"
        cmbSubArea.BoundColumn = "idTipoFinanciamiento"
        cmbSubArea.Visible = True
'    Case sghareaslaboraempleado.sghServiciosHosp
'        Dim oBuscaServicios As New SIGHNegocios.ReglasAdmision
'        Set oRsSubArea = oBuscaServicios.DevuelveServiciosDelHospital("(1,2,3,4)")
'        Set cmbSubArea.RowSource = oRsSubArea
'        cmbSubArea.ListField = "DservicioHosp"
'        cmbSubArea.BoundColumn = "idServicio"
'        cmbSubArea.Visible = True
'        Set oBuscaServicios = Nothing
    Case sghAreasLaboraEmpleado.sghEspecialidadesCE
        Dim oBuscaEspecialidadesCE As New SIGHNegocios.ReglasAdmision
        Set oRsSubArea = oBuscaEspecialidadesCE.DevuelveEspecialidadesDelHospital("(1)")
        Set cmbSubArea.RowSource = oRsSubArea
        cmbSubArea.ListField = "DescripcionLarga"
        cmbSubArea.BoundColumn = "IdEspecialidad"
        cmbSubArea.Visible = True
        Set oBuscaEspecialidadesCE = Nothing
    Case sghAreasLaboraEmpleado.sghEspecialidadesHosp
        Dim oBuscaEspecialidadesHOSP As New SIGHNegocios.ReglasAdmision
        Set oRsSubArea = oBuscaEspecialidadesHOSP.DevuelveEspecialidadesDelHospital("(3)")
        Set cmbSubArea.RowSource = oRsSubArea
        cmbSubArea.ListField = "DescripcionLarga"
        cmbSubArea.BoundColumn = "IdEspecialidad"
        cmbSubArea.Visible = True
        Set oBuscaEspecialidadesHOSP = Nothing
    Case sghAreasLaboraEmpleado.sghEspecialidadesEmergCons
        Dim oBuscaEspecialidadesEMERG As New SIGHNegocios.ReglasAdmision
        Set oRsSubArea = oBuscaEspecialidadesEMERG.DevuelveEspecialidadesDelHospital("(2)")
        Set cmbSubArea.RowSource = oRsSubArea
        cmbSubArea.ListField = "DescripcionLarga"
        cmbSubArea.BoundColumn = "IdEspecialidad"
        cmbSubArea.Visible = True
        Set oBuscaEspecialidadesEMERG = Nothing
    Case sghAreasLaboraEmpleado.sghEspecialidadesEmergObs
        Dim oBuscaEspecialidadesEMERGobs As New SIGHNegocios.ReglasAdmision
        Set oRsSubArea = oBuscaEspecialidadesEMERGobs.DevuelveEspecialidadesDelHospital("(4)")
        Set cmbSubArea.RowSource = oRsSubArea
        cmbSubArea.ListField = "DescripcionLarga"
        cmbSubArea.BoundColumn = "IdEspecialidad"
        cmbSubArea.Visible = True
        Set oBuscaEspecialidadesEMERGobs = Nothing
    Case sghAreasLaboraEmpleado.sghAreaTramitaSeguros
        Set oRsSubArea = mo_ReglasFacturacion.AreaTramitaSegurosDevuelveTodosSegunFiltro("")
        Set cmbSubArea.RowSource = oRsSubArea
        cmbSubArea.ListField = "Descripcion"
        cmbSubArea.BoundColumn = "IdAreaTramitaSeguros"
        cmbSubArea.Visible = True
    End Select
    Exit Sub
ErrCombo:
    If Err.Number = 3705 Then
       oRsSubArea.Close
       Resume
    End If
End Sub
Private Sub cmbIdDocIdentidad_Click()
    lbHuboCambioEnDato = True
End Sub
Private Sub cmbIdDocIdentidad_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbIdDocIdentidad.Text
      lbHuboCambioEnDato = False
    End If
    Select Case mo_cmbIdDocIdentidad.BoundText
    Case 1    'dni
         txtDNI.MaxLength = 8
    Case Else
         txtDNI.MaxLength = 20
    End Select
End Sub

Private Sub cmbIdTipoEmpleado_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoEmpleado
       AdministrarKeyPreview KeyCode
End Sub
Private Sub cmbIdTipoEmpleado_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdCondicionTrabajo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdCondicionTrabajo
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdCondicionTrabajo_LostFocus()
   If cmbIdCondicionTrabajo.Text <> "" Then
       mo_cmbIdCondicionTrabajo.BoundText = Val(Split(cmbIdCondicionTrabajo.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdCondicionTrabajo
End Sub

Private Sub cmbIdCondicionTrabajo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdRol_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdRol
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdRol_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdTipoSexo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoSexo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbTipoDestacado_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbTipoDestacado
   AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdAddLabora_Click()
    If cmbArea.Text = "" Then
        MsgBox "Ingrese el Area", vbInformation, Me.Caption
        Exit Sub
    End If
    If cmbSubArea.Text = "" Then
        MsgBox "Ingrese SubArea", vbInformation, Me.Caption
        Exit Sub
    End If
    
    On Error Resume Next
    mrs_LaboraEn.MoveFirst
    Do While Not mrs_LaboraEn.EOF
        If cmbArea.ListIndex = mrs_LaboraEn.Fields!idLaboraArea And Val(cmbSubArea.BoundText) = mrs_LaboraEn.Fields!idLaboraSubArea Then
            MsgBox "El Area/SubArea ya fué asignado", vbExclamation, Me.Caption
            Exit Sub
        End If
        mrs_LaboraEn.MoveNext
    Loop
    
    With mrs_LaboraEn
        .AddNew
        .Fields!idLaboraArea = cmbArea.ListIndex
        .Fields!idLaboraSubArea = Val(cmbSubArea.BoundText)
        .Fields!LaboraArea = cmbArea.Text
        .Fields!LaboraSubArea = cmbSubArea.Text
        .Update
    End With

End Sub

Private Sub cmdDelLabora_Click()
    On Error Resume Next
    With mrs_LaboraEn
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With

End Sub


Private Sub cmdSupervisorAdd_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaEmpleados
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        BuscaEmpleadoYllenaDatosDelSupervisor oBusqueda.idRegistroSeleccionado
    End If
    Set oBusqueda = Nothing

End Sub

Sub BuscaEmpleadoYllenaDatosDelSupervisor(lnIdEmpleado As Long)
    Dim oDOEmpleado As New dOEmpleado
    Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(lnIdEmpleado)
    txtSupervisor.Tag = 0
    txtSupervisor.Text = ""
    If Not oDOEmpleado Is Nothing Then
        txtSupervisor.Tag = oDOEmpleado.IdEmpleado
        txtSupervisor.Text = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    End If
    Set oDOEmpleado = Nothing
End Sub

Private Sub cmdSupervisorDel_Click()
    txtSupervisor.Tag = 0
    txtSupervisor.Text = ""
End Sub
Private Sub Form_Initialize()
    Set mo_cmbIdTipoEmpleado.MiComboBox = cmbIdTipoEmpleado
    Set mo_cmbIdCondicionTrabajo.MiComboBox = cmbIdCondicionTrabajo
    Set mo_cmbIdRol.MiComboBox = cmbIdRol
    Set mo_cmbCargos.MiComboBox = cmbCargos
    Set mo_cmbTipoDestacado.MiComboBox = cmbTipoDestacado
    Set mo_cmbIdDocIdentidad.MiComboBox = cmbIdDocIdentidad
    Set mo_CmbIdTipoSexo.MiComboBox = cmbIdTipoSexo
    'llenar campos en cmbNacionalidad--*GLCC***07/07/2020
    Set mo_cmbIdNacionalidad.MiComboBox = cmbIdNacionalidad
End Sub
Private Sub grdCargos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdCargos.Bands(0).Columns("IdTipoCargo").Hidden = True
    
    grdCargos.Bands(0).Columns("Cargo").Header.Caption = "Cargo"
    grdCargos.Bands(0).Columns("Cargo").Width = 5500

End Sub
Private Sub grdLaboraEn_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdLaboraEn.Bands(0).Columns("idLaboraArea").Hidden = True
    grdLaboraEn.Bands(0).Columns("idLaboraSubArea").Hidden = True
    '
    grdLaboraEn.Bands(0).Columns("LaboraArea").Header.Caption = "Area"
    grdLaboraEn.Bands(0).Columns("LaboraArea").Width = 2100
    '
    grdLaboraEn.Bands(0).Columns("LaboraSubArea").Header.Caption = "SubArea"
    grdLaboraEn.Bands(0).Columns("LaboraSubArea").Width = 3700

End Sub

Private Sub grdRoles_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdRoles.Bands(0).Columns("IdRol").Hidden = True
    
    grdRoles.Bands(0).Columns("Nombre").Header.Caption = "Rol"
    grdRoles.Bands(0).Columns("Nombre").Width = 5500
    
End Sub
Private Sub txtApellidoMaterno_Change()
lbHuboCambioEnDato = True
End Sub
Private Sub txtApellidoPaterno_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtClave
    AdministrarKeyPreview KeyCode
End Sub
Private Sub txtClave_LostFocus()
   mo_Formulario.MarcarComoVacio txtClave
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtCodigoHisDelDigitador_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigoHisDelDigitador
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtCodigoPlanilla_Change()
  lbHuboCambioEnDato = True
End Sub
Private Sub txtCodigoPlanilla_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigoPlanilla
AdministrarKeyPreview KeyCode
End Sub
Private Sub txtCodigoPlanilla_LostFocus()
   If lbHuboCambioEnDato = True Then
      sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtCodigoPlanilla.Text
      lbHuboCambioEnDato = False
   End If
   mo_Formulario.MarcarComoVacio txtCodigoPlanilla
End Sub
Private Sub txtCodigoPlanilla_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtDNI_Change()
lbHuboCambioEnDato = True
End Sub
Private Sub txtFechaNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaNacimiento
AdministrarKeyPreview KeyCode

End Sub
Private Sub txtFechaNacimiento_LostFocus()
    If Not EsFecha(txtFechaNacimiento.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFechaNacimiento.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
    If Year(Date) - Val(Right(txtFechaNacimiento.Text, 4)) < 15 Then
        MsgBox "No debe existir empleados menores a 15 años", vbInformation, Me.Caption
        txtFechaNacimiento.Text = sighentidades.FECHA_VACIA_DMY
        txtFechaNacimiento.SetFocus
        Exit Sub
    End If

End Sub
Private Sub txtNombres_Change()
lbHuboCambioEnDato = True
End Sub
Private Sub txtNombres_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombres
AdministrarKeyPreview KeyCode
End Sub
Private Sub txtNombres_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtNombres.Text
      lbHuboCambioEnDato = False
    End If

txtNombres.Text = mo_Teclado.CapitalizarNombres(txtNombres.Text)
   mo_Formulario.MarcarComoVacio txtNombres
End Sub
Private Sub txtNombres_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
AdministrarKeyPreview KeyCode
End Sub
Private Sub txtApellidoMaterno_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtApellidoMaterno.Text
      lbHuboCambioEnDato = False
    End If

   txtApellidoMaterno.Text = mo_Teclado.CapitalizarNombres(txtApellidoMaterno.Text)
   mo_Formulario.MarcarComoVacio txtApellidoMaterno
End Sub
Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaterno
AdministrarKeyPreview KeyCode
End Sub
Private Sub txtApellidoPaterno_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtApellidoPaterno.Text
      lbHuboCambioEnDato = False
    End If

    txtApellidoPaterno.Text = mo_Teclado.CapitalizarNombres(txtApellidoPaterno.Text)
   mo_Formulario.MarcarComoVacio txtApellidoPaterno
End Sub
Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Sub CargarDatosAlFormulario()
 mo_Formulario.HabilitarDeshabilitar txtEstablecimientoExterno, False
 mo_Formulario.HabilitarDeshabilitar txtSupervisor, False
 Select Case mi_Opcion
     Case sghAgregar
         cmbArea.ListIndex = 0
         chkUsaGalenHos.Visible = False
         chkEsActivo = 1
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
 End Select
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Sub Form_Load()
       sighentidades.ParaAuditoriaPorCadaDato sghAudLimpiar, ""
      
        GenerarRecordsetTemporal
       
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar empleado"
           chkEsActivo = 1
       Case sghModificar
           Me.Caption = "Modificar empleado"
       Case sghConsultar
           Me.Caption = "Consultar empleado"
       Case sghEliminar
           Me.Caption = "Eliminar empleado"
       End Select
       
        'Me.txtDNI.MaxLength = 8
       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
           If ml_IdEmpleado = lnDefaultEmpleado And ml_idUsuario <> lnDefaultEmpleado Then
              MsgBox "solo el usuario Administrador trabajará con este EMPLEADO", vbInformation, Me.Caption
              Me.Visible = False
           End If
       End If
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
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'AdministrarKeyPreview KeyCode
End Sub
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
                    MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                    LimpiarFormulario
                    Me.txtCodigoPlanilla.SetFocus
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
Private Sub btnCancelar_Click()
   If sighentidades.ParaAuditoria = "" Then
      Me.Visible = False
   ElseIf MsgBox("Hubo cambios, desea salir de todas maneras ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
      Me.Visible = False
   End If
End Sub
Function ValidarDatosObligatorios() As Boolean
Dim sMensaje As String
   ValidarDatosObligatorios = False
   'Validar Nacionalidad - GLCC - 08/07/2020
'   If mo_cmbIdNacionalidad.BoundText <> "166" Then
'        sMensaje = sMensaje + "Ingrese un pais" + Chr(13)
'        Exit Function
'   End If
   If Me.txtDNI.Text = "" Then
       sMensaje = sMensaje + "Ingrese el nro de DNI" + Chr(13)
   End If
   If cmbIdTipoSexo.Text = "" Then
       sMensaje = sMensaje + "Elija el sexo del Empleado" + Chr(13)
   End If
   If Me.txtCodigoPlanilla.Text = "" Then
       sMensaje = sMensaje + "Ingrese el código de planilla" + Chr(13)
   End If
   If mo_cmbIdTipoEmpleado.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese el tipo de empleado" + Chr(13)
   End If
   If mo_cmbIdCondicionTrabajo.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese la condición de trabajo" + Chr(13)
   End If
   If Me.txtNombres.Text = "" Then
       sMensaje = sMensaje + "Ingrese el nombre" + Chr(13)
   End If
   If Me.txtApellidoMaterno.Text = "" Then
       sMensaje = sMensaje + "Ingrese el apellido materno" + Chr(13)
   End If
   If Me.txtApellidoPaterno.Text = "" Then
       sMensaje = sMensaje + "Ingrese el apellido paterno" + Chr(13)
   End If
   If Me.txtFechaNacimiento.Text = sighentidades.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Ingrese una fecha de nacimiento Válida" + Chr(13)
   End If
   If cmbArea.ListIndex > 0 And cmbSubArea.Text = "" Then
       sMensaje = sMensaje + "Elija el Lugar donde Labora" + Chr(13)
   End If
   If sMensaje <> "" Then
        MsgBox sMensaje, vbExclamation, Me.Caption
        Exit Function
   End If
   If chkEsActivo.Value = 0 And mi_Opcion = sghModificar Then
      txtClave.Text = ""
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
Dim rsEmpleado As Recordset

    ValidarReglas = False
    '
    If mi_Opcion = sghAgregar And mo_AdminServiciosComunes.TiposEmpleadosSeleccionarSiSeProgramaPorId(Val(mo_cmbIdTipoEmpleado.BoundText)) = True Then
        MsgBox "El TIPO DE EMPLEADO elegido se programa" & Chr(13) & "para registrarlo debe utilizar el módulo de PROFESIONALES DE SALUD", vbInformation, Me.Caption
        Exit Function
    End If
'    If mi_Opcion = sghAgregar Then
'        If Val(mo_cmbIdTipoEmpleado.BoundText) >= 100 Then
'            MsgBox "Para ingresar médicos debe utilizar el módulo de médicos", vbInformation, Me.Caption
'            Exit Function
'        End If
'    End If
    '
    Set rsEmpleado = mo_AdminServiciosComunes.EmpleadosObtenerConElMismoCodigoPlanilla(mo_Empleado)
    If Not (rsEmpleado.EOF And rsEmpleado.BOF) Then
        MsgBox "Ya existe un empleado con el mismo CODIGO PLANILLA" + Chr(13) + rsEmpleado!ApellidoPaterno + " " + rsEmpleado!ApellidoMaterno + " " + rsEmpleado!Nombres, vbExclamation, Me.Caption
        rsEmpleado.Close
        Exit Function
    End If
    '
    If txtDNI.Text <> "" Then
        If mo_cmbIdDocIdentidad.BoundText = "1" And Len(Trim(txtDNI.Text)) <> 8 Then
            MsgBox "El DNI debe tener longitud 8", vbInformation, Me.Caption
            Exit Function
        End If
        Set rsEmpleado = mo_AdminServiciosComunes.EmpleadosObtenerConelMismoDNI(txtDNI.Text, Val(mo_cmbIdDocIdentidad.BoundText))
        Select Case mi_Opcion
        Case sghAgregar
             If rsEmpleado.RecordCount > 0 Then
                 MsgBox "Ese N° DOCUMENTO ya esta Registrado para: " + Trim(rsEmpleado.Fields!ApellidoPaterno) & " " & Trim(rsEmpleado.Fields!ApellidoMaterno) & rsEmpleado.Fields!Nombres, vbInformation, Me.Caption
                 Exit Function
             End If
        Case sghModificar
             If rsEmpleado.RecordCount > 0 Then
                rsEmpleado.MoveFirst
                Do While Not rsEmpleado.EOF
                   If Trim(rsEmpleado.Fields!DNI) = Trim(Me.txtDNI.Text) And rsEmpleado.Fields!IdEmpleado <> ml_IdEmpleado Then
                      MsgBox "Ese N° DOCUMENTO ya esta Registrado para: " + Trim(rsEmpleado.Fields!ApellidoPaterno) & " " & Trim(rsEmpleado.Fields!ApellidoMaterno) & rsEmpleado.Fields!Nombres, vbInformation, Me.Caption
                      Exit Function
                   End If
                   rsEmpleado.MoveNext
                Loop
             End If
        End Select
    End If
    '
    Set rsEmpleado = mo_AdminServiciosComunes.EmpleadosObtenerConElMismoUsuario(mo_Empleado)
    If Not (rsEmpleado.EOF And rsEmpleado.BOF) Then
        MsgBox "Ya existe un empleado con el mismo USUARIO" + Chr(13) + rsEmpleado!ApellidoPaterno + " " + rsEmpleado!ApellidoMaterno + " " + rsEmpleado!Nombres, vbExclamation, Me.Caption
        rsEmpleado.Close
        Exit Function
    End If
    '
    Set rsEmpleado = Nothing
    ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

   With mo_Empleado
            .CodigoPlanilla = txtCodigoPlanilla
            .DNI = Me.txtDNI
           .IdEmpleado = Me.IdEmpleado
           .IdTipoEmpleado = Val(mo_cmbIdTipoEmpleado.BoundText)
           .IdCondicionTrabajo = Val(mo_cmbIdCondicionTrabajo.BoundText)
           .Nombres = Me.txtNombres.Text
           .ApellidoMaterno = Me.txtApellidoMaterno.Text
           .ApellidoPaterno = Me.txtApellidoPaterno.Text
           .IdUsuarioAuditoria = Me.idUsuario
           .Usuario = Me.txtUsuario
           .Clave = Me.txtClave
           .LoginEstado = IIf(chkUsaGalenHos.Value, 1, 0)
           If chkUsaGalenHos.Value = 0 Then
              .LoginPc = ""
           End If
           .FechaNacimiento = IIf(Me.txtFechaNacimiento.Text = sighentidades.FECHA_VACIA_DMY, 0, Me.txtFechaNacimiento.Text)
           .idTipoDestacado = Val(mo_cmbTipoDestacado.BoundText)
           .HisCodigoDigitador = txtCodigoHisDelDigitador.Text
           .IdEstablecimientoExterno = lnIdEstablecimientoExterno
           .ReniecAutorizado = IIf(chkAutorizadoReniec.Value = 1, True, False)
           .idTipoDocumento = Val(mo_cmbIdDocIdentidad.BoundText)
           .IdSupervisor = Val(txtSupervisor.Tag)
           .EsActivo = IIf(Me.chkEsActivo.Value = 1, True, False)
           'Nacionalidad Peru-GLCC-08/07/2020
          .IdPais = Val(mo_cmbIdNacionalidad.BoundText)
           If mi_Opcion = sghAgregar Then
              .fechaingreso = Date
           End If
           .idTipoSexo = Val(mo_CmbIdTipoSexo.BoundText)
   End With
   CargarRolItemsAlObjetoDatos mo_UsuarioRoles
End Sub

Function EncriptaDNIsiTieneAccesoRENIEC() As String
    EncriptaDNIsiTieneAccesoRENIEC = ""
    If Me.chkAutorizadoReniec.Value = 1 Then
        Dim oCrypKey As New CrypKey.Util
        EncriptaDNIsiTieneAccesoRENIEC = oCrypKey.EncryptString(mo_Empleado.DNI)
        Set oCrypKey = Nothing
    End If
End Function

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    
    AgregarDatos = mo_AdminServiciosComunes.EmpleadosAgregar(mo_Empleado, mo_UsuarioRoles, mrs_UsuariosCargos, mrs_LaboraEn, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & " " & Trim(txtNombres.Text), EncriptaDNIsiTieneAccesoRENIEC)
   
End Function



'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    
    ModificarDatos = mo_AdminServiciosComunes.EmpleadosModificar(mo_Empleado, mo_UsuarioRoles, mrs_UsuariosCargos, mrs_LaboraEn, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & " " & Trim(txtNombres.Text), EncriptaDNIsiTieneAccesoRENIEC)
   
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    Dim lcMensaje As String
    lcMensaje = mo_AdminServiciosComunes.EmpleadosChequeaNoTengaMovimientos(mo_Empleado.IdEmpleado)
    If lcMensaje <> "" Then
       MsgBox "No se puede eliminar" & Chr(13) & lcMensaje, vbInformation, Me.Caption
    Else
       EliminarDatos = mo_AdminServiciosComunes.EmpleadosEliminar(mo_Empleado, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & " " & Trim(txtNombres.Text))
    End If
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
Dim oDOEmpleado As dOEmpleado
       lblId.Caption = ""
       Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(Me.IdEmpleado)
       
       If mo_AdminServiciosComunes.MensajeError <> "" Then
            MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption
            mb_ExistenDatos = False
            Exit Sub
       End If
       
       If Not oDOEmpleado Is Nothing Then
            With oDOEmpleado
                lblId.Caption = .IdEmpleado
                Me.IdEmpleado = .IdEmpleado
                mo_cmbIdTipoEmpleado.BoundText = .IdTipoEmpleado
                mo_cmbIdCondicionTrabajo.BoundText = .IdCondicionTrabajo
                Me.txtNombres.Text = .Nombres
                Me.txtApellidoMaterno.Text = .ApellidoMaterno
                Me.txtApellidoPaterno.Text = .ApellidoPaterno
                Me.txtDNI = .DNI
                Me.txtCodigoPlanilla = .CodigoPlanilla
                Me.txtUsuario = .Usuario
                Me.txtClave = .Clave
                txtCodigoHisDelDigitador = .HisCodigoDigitador
                Me.chkUsaGalenHos = IIf(.LoginEstado = 1, 1, 0)
                chkAutorizadoReniec.Value = IIf(.ReniecAutorizado = True, 1, 0)
                Me.chkEsActivo.Value = IIf(.EsActivo = True, 1, 0)
                If .LoginEstado = 1 Then
                    lblNombrePc = IIf(IsNull(.LoginPc), "", .LoginPc)
                End If
                If .FechaNacimiento <> 0 Then
                   Me.txtFechaNacimiento.Text = .FechaNacimiento
                End If
                mo_cmbTipoDestacado.BoundText = .idTipoDestacado
                mo_cmbIdDocIdentidad.BoundText = .idTipoDocumento
                'cmbNacionalidad-09/07/2020
'                mo_cmbIdNacionalidad.BoundText = .idPais
                mo_cmbIdNacionalidad.BoundText = .IdPais
                '
                lnIdEstablecimientoExterno = .IdEstablecimientoExterno
                CargaNombreEstablecimiento
                '
                BuscaEmpleadoYllenaDatosDelSupervisor .IdSupervisor
                '
                mo_CmbIdTipoSexo.BoundText = .idTipoSexo
                Set mo_Empleado = oDOEmpleado
                mb_ExistenDatos = True
                lbHuboCambioEnDato = False
            End With
            
            CargarDatosDeRolItems
            CargarDatosDeCargos
            CargarDatosDeDondeLabora
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
   
    Select Case mi_Opcion
    Case sghModificar
        If mo_AdminServiciosComunes.TiposEmpleadosSeleccionarSiSeProgramaPorId(Val(mo_cmbIdTipoEmpleado.BoundText)) = True Then
            Me.cmbIdTipoEmpleado.Enabled = False
            MsgBox "Para modificar DATOS PERSONALES de médicos ingrese a la opción 'Profesionales de Salud' del módulo 'Programación General", vbInformation, Me.Caption
        End If
    Case sghEliminar
        If mo_AdminServiciosComunes.TiposEmpleadosSeleccionarSiSeProgramaPorId(Val(mo_cmbIdTipoEmpleado.BoundText)) = True Then
            Me.Frame1.Enabled = False
            MsgBox "Para eliminar médicos ingrese a la opción médicos del módulo programación médica ", vbInformation, Me.Caption
            'Me.btnAceptar.Enabled = False
        Else
            Me.Frame1.Enabled = False
        End If
    Case sghConsultar
        Me.Frame1.Enabled = False
        Me.btnAceptar.Enabled = False
    End Select
End Sub

Sub CargaNombreEstablecimiento()
    Dim oRsTmp1 As New Recordset
    txtEstablecimientoExterno.Text = ""
    If lnIdEstablecimientoExterno > 0 Then
       Set oRsTmp1 = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorFiltro("idEstablecimiento=" & Trim(Str(lnIdEstablecimientoExterno)))
       If oRsTmp1.RecordCount > 0 Then
          txtEstablecimientoExterno.Text = oRsTmp1.Fields!Nombre
       End If
       oRsTmp1.Close
    End If
    Set oRsTmp1 = Nothing
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()
           sighentidades.ParaAuditoriaPorCadaDato sghAudLimpiar, ""
           
           Me.IdEmpleado = 0
           mo_cmbIdTipoEmpleado.BoundText = ""
           mo_cmbIdCondicionTrabajo.BoundText = ""
           mo_cmbTipoDestacado.BoundText = "3"
           mo_cmbIdDocIdentidad.BoundText = "1"
           'cmbnacionalidad - GLCC-09/07/2020
          mo_cmbIdNacionalidad.BoundText = "1"
           Me.txtNombres.Text = ""
           Me.txtApellidoMaterno.Text = ""
           Me.txtApellidoPaterno.Text = ""
           Me.txtDNI = ""
           Me.txtCodigoPlanilla = ""
           Me.txtUsuario = ""
           Me.txtClave = ""
           txtCodigoHisDelDigitador.Text = ""
           chkAutorizadoReniec.Value = 0
           Me.txtFechaNacimiento.Text = sighentidades.FECHA_VACIA_DMY
           txtSupervisor.Text = ""
           txtSupervisor.Tag = 0
           If mrs_UsuariosRoles.RecordCount > 0 Then
                mrs_UsuariosRoles.MoveFirst
                Do While Not mrs_UsuariosRoles.EOF
                   mrs_UsuariosRoles.Delete
                   mrs_UsuariosRoles.Update
                   mrs_UsuariosRoles.MoveNext
                Loop
           End If
           Set Me.grdRoles.DataSource = mrs_UsuariosRoles
           '
           If mrs_UsuariosCargos.RecordCount > 0 Then
                mrs_UsuariosCargos.MoveFirst
                Do While Not mrs_UsuariosCargos.EOF
                   mrs_UsuariosCargos.Delete
                   mrs_UsuariosCargos.Update
                   mrs_UsuariosCargos.MoveNext
                Loop
           End If
           Set Me.grdCargos.DataSource = mrs_UsuariosCargos
           '
           If mrs_LaboraEn.RecordCount > 0 Then
              mrs_LaboraEn.MoveFirst
              Do While Not mrs_LaboraEn.EOF
                 mrs_LaboraEn.Delete
                 mrs_LaboraEn.Update
                 mrs_LaboraEn.MoveNext
              Loop
           End If
           Set Me.grdLaboraEn.DataSource = mrs_LaboraEn
End Sub


Private Sub txtDNI_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDNI
   AdministrarKeyPreview KeyCode
End Sub


Private Sub txtDni_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtDNI.Text
      lbHuboCambioEnDato = False
    End If
   
   If Len(Trim(txtDNI.Text)) > 0 Then
        If mo_cmbIdDocIdentidad.BoundText = "1" And Len(Trim(txtDNI.Text)) <> 8 Then
           MsgBox "Si el Documento es DNI debe tener 8 dígitos", vbInformation, "Mensaje"
           On Error Resume Next
           txtDNI.SetFocus
           Exit Sub
        End If
        
'<(Inicio) Añadido Por: WABG el: 26/01/2021-12:08:01 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
        lbBuscaDNIenReniec = IIf(lcBuscaParametro.SeleccionaFilaParametro(296) = "S", True, False)
        If mo_cmbIdDocIdentidad.BoundText = "1" And lbBuscaDNIenReniec = True And Len(Trim(txtApellidoPaterno.Text)) = 0 And Len(Trim$(txtApellidoMaterno.Text)) = 0 And mi_Opcion = sghAgregar Then
                            
                mo_Reniec.SeAccesaAlaWebDesdeGalenhos = True
                mo_Reniec.Inicializar
           
               mo_Reniec.ConsultarDNIenReniec txtDNI.Text
            If mo_Reniec.ApellidoPaterno <> "" Then
                  
                  txtCodigoPlanilla.Text = txtDNI.Text
                  txtApellidoPaterno.Text = mo_Reniec.ApellidoPaterno
                  txtApellidoMaterno.Text = mo_Reniec.ApellidoMaterno
                  txtNombres.Text = Trim(mo_Reniec.PrimerNombre) + " " + Trim(mo_Reniec.SegundoNombre) + " " + Trim$(mo_Reniec.TercerNombre)
                  txtFechaNacimiento.Text = mo_Reniec.FechaNacimiento
                  
                  'CARGAR SEXO
                  If mo_Reniec.idTipoSexo = 1 Then
                    cmbIdTipoSexo.ListIndex = 0
                  Else
                    cmbIdTipoSexo.ListIndex = 1
                  End If
                  
            End If
        End If
'</(Fin) Añadido Por: WABG el: 26/01/2021-12:08:01 p.m. en el Equipo: SISGALENPLUS-PC<CAMBIO-37>
        
        
   End If
   mo_Formulario.MarcarComoVacio txtDNI
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Val(mo_cmbIdDocIdentidad.BoundText) <> 2 And Val(mo_cmbIdDocIdentidad.BoundText) <> 3 Then
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If
       Else
            If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
                KeyAscii = 0
            End If
       End If
   End If
End Sub

Private Sub txtUsuario_LostFocus()
   mo_Formulario.MarcarComoVacio txtUsuario
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtUsuario
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub CargarDatosDeRolItems()
Dim rsRolItems As New Recordset

    Set rsRolItems = mo_AdminSeguridad.UsuariosRolesSeleccionarPorEmpleado(ml_IdEmpleado)
    Do While Not rsRolItems.EOF
        With mrs_UsuariosRoles
            .AddNew
            .Fields!IdRol = rsRolItems!IdRol
            .Fields!Nombre = rsRolItems!Nombre
        End With
        rsRolItems.MoveNext
    Loop
    
    
End Sub
Sub CargarDatosDeDondeLabora()
    Dim oRsTmp As New Recordset
    Dim lcSql As String
    With mrs_LaboraEn
        Set oRsTmp = mo_AdminServiciosComunes.EmpleadosLugarDeTrabajoSeleccionarPorFiltro("idEmpleado=" & Trim(Str(ml_IdEmpleado)))
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
                Me.cmbArea.ListIndex = oRsTmp.Fields!idLaboraArea
                Me.cmbSubArea.BoundText = oRsTmp.Fields!idLaboraSubArea
                .AddNew
                .Fields!idLaboraArea = oRsTmp.Fields!idLaboraArea
                .Fields!idLaboraSubArea = oRsTmp.Fields!idLaboraSubArea
                .Fields!LaboraArea = cmbArea.Text
                .Fields!LaboraSubArea = cmbSubArea.Text
                .Update
                oRsTmp.MoveNext
           Loop
        End If
        oRsTmp.Close
    End With
    Set oRsTmp = Nothing
End Sub
Sub CargarDatosDeCargos()
    Dim oRsTmp As New Recordset
    Dim lcSql As String
    Set oRsTmp = mo_AdminServiciosComunes.EmpleadosCargosSeleccionarPorFiltro("dbo.EmpleadosCargos.idEmpleado=" & Trim(Str(ml_IdEmpleado)))
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
           mrs_UsuariosCargos.AddNew
           mrs_UsuariosCargos.Fields!IdTipoCargo = oRsTmp.Fields!IdCargo
           mrs_UsuariosCargos.Fields!cargo = oRsTmp.Fields!cargo
           mrs_UsuariosCargos.Update
           oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Sub
Sub CargarRolItemsAlObjetoDatos(oUsuarioRoles As Collection)
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS ExamenS
    '---------------------------------------------------------------------------------
    Dim oUsuarioRol As DOUsuarioRol
    Dim lnRegRol As Integer
'    If oUsuarioRoles.Count > 0 Then
'       For lnRegRol = 1 To oUsuarioRoles.Count
'           oUsuarioRoles.Remove (lnRegRol)
'       Next
'    End If
    Set oUsuarioRoles = New Collection
    If mrs_UsuariosRoles.RecordCount > 0 Then
        mrs_UsuariosRoles.MoveFirst
        Do While Not mrs_UsuariosRoles.EOF
            Set oUsuarioRol = New DOUsuarioRol
            oUsuarioRol.IdUsuarioRol = 0
            oUsuarioRol.IdEmpleado = ml_IdEmpleado
            oUsuarioRol.IdRol = mrs_UsuariosRoles!IdRol
            oUsuarioRol.IdUsuarioAuditoria = ml_idUsuario
            
            oUsuarioRoles.Add oUsuarioRol
            mrs_UsuariosRoles.MoveNext
        Loop
    End If
End Sub


Sub GenerarRecordsetTemporal()
    
    With mrs_UsuariosRoles
          .Fields.Append "IdRol", adInteger, 4, adFldIsNullable
          .Fields.Append "Nombre", adVarChar, 100, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdRoles.DataSource = mrs_UsuariosRoles
    mo_Apariencia.ConfigurarFilasBiColores Me.grdRoles, sighentidades.GrillaConFilasBicolor

    '
    With mrs_UsuariosCargos
          .Fields.Append "IdTipoCargo", adInteger, 4, adFldIsNullable
          .Fields.Append "Cargo", adVarChar, 100, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdCargos.DataSource = mrs_UsuariosCargos
    mo_Apariencia.ConfigurarFilasBiColores Me.grdCargos, sighentidades.GrillaConFilasBicolor
    '
    With mrs_LaboraEn
          .Fields.Append "idLaboraArea", adInteger, 4, adFldIsNullable
          .Fields.Append "LaboraArea", adVarChar, 100, adFldIsNullable
          .Fields.Append "idLaboraSubArea", adInteger, 4, adFldIsNullable
          .Fields.Append "LaboraSubArea", adVarChar, 100, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdLaboraEn.DataSource = mrs_LaboraEn
    mo_Apariencia.ConfigurarFilasBiColores Me.grdLaboraEn, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub btnAgregarDx_Click()
    
    If mo_cmbIdRol.BoundText = "" Then
        MsgBox "Ingrese el rol", vbInformation, Me.Caption
        Exit Sub
    End If
    If ml_idUsuario <> lnDefaultEmpleado Then
       If Val(mo_cmbIdRol.BoundText) = 1 Or Val(mo_cmbIdRol.BoundText) = 12 Then
          MsgBox "Solo el Administrador puede usar este ROL", vbInformation, Me.Caption
          Exit Sub
       End If
    End If
    On Error Resume Next
    mrs_UsuariosRoles.MoveFirst
    Do While Not mrs_UsuariosRoles.EOF
        If mo_cmbIdRol.BoundText = mrs_UsuariosRoles!IdRol Then
            MsgBox "El rol ya fue asignado", vbExclamation, Me.Caption
            Exit Sub
        End If
        mrs_UsuariosRoles.MoveNext
    Loop
    
    With mrs_UsuariosRoles
        .AddNew
        .Fields!IdRol = mo_cmbIdRol.BoundText
        .Fields!Nombre = cmbIdRol.Text
        .Update
    End With
    sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "+Rol: " & cmbIdRol.Text
End Sub

Private Sub btnQuitarDx_Click()
    On Error GoTo errBtnQ
    With mrs_UsuariosRoles
        If Not .EOF And Not .BOF Then
           sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "-Rol: " & mrs_UsuariosRoles!Nombre
           .Delete
           .Update
        End If
    End With
errBtnQ:
End Sub

