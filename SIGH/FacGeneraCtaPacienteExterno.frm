VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form FacGeneraCtaPacienteExterno 
   Caption         =   "Genera Cuenta para un Paciente Externo - PARTICULAR"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14745
   Icon            =   "FacGeneraCtaPacienteExterno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   14745
   StartUpPosition =   2  'CenterScreen
   Begin UltraGrid.SSUltraGrid grdPacientesEncontrados 
      Height          =   225
      Left            =   0
      TabIndex        =   7
      Top             =   810
      Visible         =   0   'False
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   397
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
      Caption         =   "Lista de pacientes encontrados"
   End
   Begin VB.Frame Frame1 
      Caption         =   "PreVentas generadas en CAJA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8235
      Left            =   10740
      TabIndex        =   35
      Top             =   60
      Width           =   4005
      Begin VB.TextBox txtConsideraciones 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   765
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   38
         Text            =   "FacGeneraCtaPacienteExterno.frx":000C
         Top             =   240
         Width           =   3765
      End
      Begin UltraGrid.SSUltraGrid grdPreVentaCab 
         Height          =   3255
         Left            =   90
         TabIndex        =   36
         Top             =   1110
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   5741
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
         Caption         =   "Cabecera"
      End
      Begin UltraGrid.SSUltraGrid grdPreVentaDet 
         Height          =   3675
         Left            =   90
         TabIndex        =   37
         Top             =   4470
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   6482
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
         Caption         =   "Detalle"
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
      Height          =   975
      Left            =   30
      TabIndex        =   31
      Top             =   60
      Width           =   10695
      Begin VB.Frame fraPacienteNuevo 
         Height          =   795
         Left            =   8700
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   1905
         Begin VB.CheckBox chkPacienteNuevo 
            Alignment       =   1  'Right Justify
            Caption         =   "Paciente &nuevo"
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
            Left            =   90
            TabIndex        =   33
            Top             =   210
            Width           =   1605
         End
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
         TabIndex        =   6
         Top             =   450
         Width           =   1065
      End
      Begin VB.CommandButton btnBuscarPaciente 
         Height          =   315
         Left            =   7350
         Picture         =   "FacGeneraCtaPacienteExterno.frx":005D
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   1305
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
         Left            =   6120
         TabIndex        =   4
         Top             =   450
         Width           =   1185
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
         Left            =   5010
         TabIndex        =   3
         Top             =   450
         Width           =   1080
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
         Left            =   2475
         TabIndex        =   1
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
         Left            =   1230
         TabIndex        =   0
         Top             =   450
         Width           =   1185
      End
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
         Left            =   3885
         TabIndex        =   2
         Top             =   450
         Width           =   1080
      End
      Begin VB.Label Label50 
         Caption         =   "Nº Historia      Ap. paterno      Ap. materno   1er nombre    2do Nombre       DNI"
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
         TabIndex        =   34
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de la Atención:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   30
      TabIndex        =   12
      Top             =   1080
      Width           =   10665
      Begin VB.Frame Frame6 
         Enabled         =   0   'False
         Height          =   3195
         Left            =   120
         TabIndex        =   46
         Top             =   300
         Width           =   4455
         Begin GalenHos.UcPacienteDatosAloj UcPacienteDatosAloj1 
            Height          =   3195
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   5636
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Atención"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   4590
         TabIndex        =   13
         Top             =   300
         Width           =   5985
         Begin VB.TextBox txtNroOrdenPago 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1740
            TabIndex        =   44
            Top             =   2730
            Width           =   4125
         End
         Begin VB.TextBox txtEdadEnDias 
            BackColor       =   &H00FFFFFF&
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
            Left            =   4440
            TabIndex        =   20
            Top             =   1080
            Width           =   405
         End
         Begin VB.TextBox lblNombreMedico 
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
            Left            =   1740
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   690
            Width           =   3255
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
            Left            =   1740
            TabIndex        =   18
            Top             =   270
            Width           =   4170
         End
         Begin VB.ComboBox cmbIdTipoEdad 
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
            ItemData        =   "FacGeneraCtaPacienteExterno.frx":2CA6
            Left            =   4860
            List            =   "FacGeneraCtaPacienteExterno.frx":2CA8
            TabIndex        =   17
            Top             =   1080
            Width           =   1035
         End
         Begin VB.TextBox txtIdMedicoIngreso 
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
            Left            =   4980
            TabIndex        =   16
            Top             =   690
            Width           =   885
         End
         Begin VB.TextBox txtNroCuenta 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1740
            TabIndex        =   15
            Top             =   2340
            Width           =   1635
         End
         Begin VB.Frame Frame5 
            Height          =   30
            Left            =   -60
            TabIndex        =   14
            Top             =   1410
            Width           =   6015
         End
         Begin MSMask.MaskEdBox txtHoraIngreso 
            Height          =   315
            Left            =   3000
            TabIndex        =   21
            Top             =   1080
            Width           =   660
            _ExtentX        =   1164
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
         Begin MSMask.MaskEdBox txtFechaIngreso 
            Height          =   315
            Left            =   1740
            TabIndex        =   22
            Top             =   1080
            Width           =   1230
            _ExtentX        =   2170
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
         Begin MSDataListLib.DataCombo cmbFuenteFinanciamiento 
            Height          =   330
            Left            =   1740
            TabIndex        =   23
            Top             =   1500
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
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
         Begin MSDataListLib.DataCombo cmbFormaPago 
            Height          =   360
            Left            =   1740
            TabIndex        =   42
            Top             =   1890
            Width           =   4155
            _ExtentX        =   7329
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Orden Pago"
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
            TabIndex        =   45
            Top             =   2760
            Width           =   1275
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Producto/Plan"
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
            TabIndex        =   43
            Top             =   1950
            Width           =   1185
         End
         Begin VB.Label lblFecha 
            Caption         =   "Fecha ingreso"
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
            TabIndex        =   30
            Top             =   1110
            Width           =   1215
         End
         Begin VB.Label lblEdadEnDias 
            Alignment       =   1  'Right Justify
            Caption         =   "Edad"
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
            Left            =   3900
            TabIndex        =   29
            Top             =   1110
            Width           =   495
         End
         Begin VB.Label lblIdServicioIngreso 
            Caption         =   "Servicio ingreso"
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
            TabIndex        =   28
            Top             =   300
            Width           =   1395
         End
         Begin VB.Label lblIdMedicoIngreso 
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
            Height          =   315
            Left            =   120
            TabIndex        =   27
            Top             =   735
            Width           =   1335
         End
         Begin VB.Label lblEstadoCta 
            AutoSize        =   -1  'True
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   3480
            TabIndex        =   26
            Top             =   2400
            Width           =   2310
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   210
            Left            =   120
            TabIndex        =   25
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label6 
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
            Left            =   120
            TabIndex        =   24
            Top             =   1560
            Width           =   1575
         End
      End
      Begin UltraGrid.SSUltraGrid grdProductos 
         Height          =   2235
         Left            =   120
         TabIndex        =   39
         Top             =   3540
         Width           =   10425
         _ExtentX        =   18389
         _ExtentY        =   3942
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
         Caption         =   "PreVentas de la Cuenta"
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "..."
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
         Left            =   8130
         TabIndex        =   40
         Top             =   5790
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   8
      Top             =   7200
      Width           =   10665
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FacGeneraCtaPacienteExterno.frx":2CAA
         DownPicture     =   "FacGeneraCtaPacienteExterno.frx":316E
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
         Left            =   5430
         Picture         =   "FacGeneraCtaPacienteExterno.frx":365A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar e Imprimir (F2)"
         DisabledPicture =   "FacGeneraCtaPacienteExterno.frx":3B46
         DownPicture     =   "FacGeneraCtaPacienteExterno.frx":3FA6
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
         Left            =   3900
         Picture         =   "FacGeneraCtaPacienteExterno.frx":441B
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimePreCta 
         Caption         =   "Imprime Cuenta"
         Height          =   700
         Left            =   120
         Picture         =   "FacGeneraCtaPacienteExterno.frx":4890
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   1245
      End
      Begin GalenHos.ucMensajeParpadeando ucMensajeParpadeando1 
         Height          =   885
         Left            =   6870
         TabIndex        =   41
         Top             =   180
         Visible         =   0   'False
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   1561
      End
   End
End
Attribute VB_Name = "FacGeneraCtaPacienteExterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mi_Opcion As sghOpciones
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
'
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
'
Dim oRsPreVentaCab As New Recordset
Dim oRsPreVentaDet As New Recordset
Dim mrs_FacturacionProductos As New Recordset
Dim oRsFormaPago As New Recordset
'
Dim mo_cmbServicioIngreso As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdTipoEdad As New SIGHEntidades.ListaDespleglable
'
Dim oRsFuentesFinanciamiento As New Recordset
'
Dim mo_CuentasAtencion As New DOCuentaAtencion
Dim mo_Atenciones As New DOAtencion
Dim mo_Pacientes  As New doPaciente
'
Dim ml_idCuentaAtencion As Long
Dim ml_idPaciente As Long
'
Dim lcBuscaParametro As New SIGHDatos.Parametros
'
Dim lcSql As String
Dim ms_MensajeError As String
Dim lnEspecialidadServicio As Long
Dim lbUltimaTeclaPulsoENTER As Boolean
'
Dim ml_lbPacienteTieneSeguro As Boolean
Dim ml_idUsuario As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_idAtencion As Long
Dim mo_lnIdPuntoCarga As Long

Property Let idPuntoCarga(lValue As Long)
  mo_lnIdPuntoCarga = lValue
End Property


Property Let lcNombrePc(lValue As String)
  mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
  mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let idUsuario(lValue As Long)
  ml_idUsuario = lValue
End Property

Property Let Opcion(iValue As sghOpciones)
  mi_Opcion = iValue
End Property

Property Get Opcion() As sghOpciones
  Opcion = mi_Opcion
End Property

Property Let lbPacienteTieneSeguro(lValue As Boolean)
  ml_lbPacienteTieneSeguro = lValue
End Property

Property Let idAtencion(lValue As Long)
  ml_idAtencion = lValue
End Property

Property Get idAtencion() As Long
  idAtencion = ml_idAtencion
End Property


Sub LimpiarBusqueda()
    Me.txtNroHistoriaBusqueda.Text = ""
    Me.txtApellidoPaternoBusqueda.Text = ""
    Me.txtApellidoMaternoBusqueda.Text = ""
    Me.txtPrimerNombreBusqueda.Text = ""
    Me.txtSegundoNombreBusqueda.Text = ""
    Me.txtNroDNIBusqueda.Text = ""

End Sub

Private Sub btnAceptar_Click()
  If btnAceptar.Enabled = False Then Exit Sub
  Dim oConexion As New Connection
  oConexion.Open SIGHEntidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  Select Case mi_Opcion
    Case sghAgregar
      If ValidarDatosObligatorios() Then
        CargaDatosAlObjetosDeDatos
        If ValidarReglas() Then
          If AgregarDatos() Then
            Me.txtNroCuenta = mo_Atenciones.idCuentaAtencion
            
            txtNroOrdenPago.Text = mo_AdminFacturacion.DevuelveOrdenesPagoSegunCuenta(mo_Atenciones.idCuentaAtencion, oConexion)
            ms_MensajeError = " Los datos se agregaron correctamente, para la Historia Nª: " & mo_Pacientes.NroHistoriaClinica & Chr(13) & Chr(13) & "N° Cuenta " & txtNroCuenta.Text & Chr(13) & Chr(13) & txtNroOrdenPago.Text
            If mo_lnIdPuntoCarga = 99 Then
                'esta en CAJA
                If MsgBox(ms_MensajeError & Chr(13) & Chr(13) & "¿Desea Imprimir TICKET ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                     Me.Visible = False
                     Exit Sub
                End If
            Else
                MsgBox ms_MensajeError, vbInformation, Me.Caption
                Me.Visible = False
            End If
            If Me.txtNroCuenta.Text <> "" Then
               ImprimePreCuenta
            End If
            Me.Visible = False
          Else
            MsgBox "No se pudo agregar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
          End If
        End If
      End If
    Case sghModificar
       If ValidarDatosObligatorios() Then
         CargaDatosAlObjetosDeDatos
         If ValidarReglas() Then
           If ModificarDatos() Then
             MsgBox " Los datos se modificaron correctamente, para la Cuenta N° " & txtNroCuenta.Text, vbInformation, Me.Caption
             Me.Visible = False
           Else
             MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
           End If
         End If
       End If
    Case sghEliminar
      'If ValidarReglas() Then
        CargaDatosAlObjetosDeDatos
        If EliminarDatos() Then
          MsgBox "Los datos se eliminaron correctamente, para la Cuenta N° " & txtNroCuenta.Text, vbInformation, Me.Caption
          Me.Visible = False
        Else
          MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
        End If
      'End If
  End Select
  oConexion.Close
  Set oConexion = Nothing
End Sub



Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   UcPacienteDatosAloj1.CargarDatosAlObjetoDatos mo_Pacientes
   If mo_Pacientes.ApellidoPaterno = "" Then
       sMensaje = sMensaje + "Ingrese el Apellido Paterno " + Chr(13)
   End If
   If mo_Pacientes.ApellidoMaterno = "" Then
       sMensaje = sMensaje + "Ingrese el Apellido Materno " + Chr(13)
   End If
   If mo_Pacientes.PrimerNombre = "" Then
       sMensaje = sMensaje + "Ingrese el Apellido Primer Nombre" + Chr(13)
   End If
   If mo_Pacientes.idTipoSexo = 0 Then
       sMensaje = sMensaje + "Elija el Sexo" + Chr(13)
   End If
   If cmbServicioIngreso.Text = "" Then
       sMensaje = sMensaje + "Elija el Servicio de Ingreso" + Chr(13)
   End If
   If txtIdMedicoIngreso.Text = "" Then
       sMensaje = sMensaje + "Ingrese el Responsable" + Chr(13)
   End If
   If txtFechaIngreso.Text = SIGHEntidades.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Registre la Fecha de Ingreso " + Chr(13)
   End If
   If txtHoraIngreso.Text = SIGHEntidades.HORA_VACIA_HM Then
       sMensaje = sMensaje + "Registre la Hora de Ingreso" + Chr(13)
   End If
   If txtEdadEnDias.Text = "" Then
       sMensaje = sMensaje + "Ingrese la Edad" + Chr(13)
   End If
   If cmbIdTipoEdad.Text = "" Then
       sMensaje = sMensaje + "Elija el Tipo de Edad" + Chr(13)
   End If
   If mrs_FacturacionProductos.RecordCount = 0 Then
      sMensaje = sMensaje + "Debe haber al menos 1 preventa" + Chr(13)
   End If
   If Me.cmbFuenteFinanciamiento.Text = "" Then
      sMensaje = sMensaje + "Elija el Plan de Atención" + Chr(13)
   End If
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
    ValidarReglas = False
    ValidarReglas = UcPacienteDatosAloj1.ValidarReglas(mo_Pacientes)
End Function

Sub CargaDatosAlObjetosDeDatos()
    Dim oRsTmp As New Recordset
    Dim lnIdTipoServicio As Long, lnIdTipoFinanciamiento As Long
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DEL PACIENTE
    '---------------------------------------------------------------------------------
    '********mo_Pacientes****** YA SE CARGO EN VALIDADATOSOBLIGATORIOS()
    '
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA CUENTA ATENCION
    '---------------------------------------------------------------------------------
    Select Case mi_Opcion
    Case sghAgregar
        With mo_CuentasAtencion
                .idPaciente = mo_Pacientes.idPaciente
                .TotalAsegurado = 0
                .TotalExonerado = 0
                .TotalPagado = 0
                .TotalPorPagar = 0
                .idEstado = sghEstadoCuenta.sghAbierto
                .FechaApertura = Me.txtFechaIngreso.Text
                .HoraApertura = Me.txtHoraIngreso.Text
                .FechaCierre = 0
                .HoraCierre = ""
                .IdUsuarioAuditoria = ml_idUsuario
        End With
    Case Else
    End Select
   
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
    lnIdTipoServicio = 0
    oRsTmp.Open "select * from Servicios where idServicio=" & mo_cmbServicioIngreso.BoundText, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       lnIdTipoServicio = oRsTmp.Fields!idTipoServicio
    End If
    oRsTmp.Close
    With mo_Atenciones
           .idAtencion = Me.idAtencion
           .IdEspecialidadMedico = 0
           .idMedicoIngreso = Val(Me.txtIdMedicoIngreso.Tag)
           .IdMedicoRespNacimiento = 0
           .IdServicioIngreso = Val(mo_cmbServicioIngreso.BoundText)
           .IdTipoReferenciaOrigen = 0
           .idEstablecimientoOrigen = 0
           .IdEstablecimientoNoMinsaOrigen = 0
           .IdTipoReferenciaDestino = 0
           .idEstablecimientoDestino = 0
           .IdEstablecimientoNoMinsaDestino = 0
           .HoraIngreso = IIf(Me.txtHoraIngreso.Text = SIGHEntidades.HORA_VACIA_HM, "", Me.txtHoraIngreso.Text)
           .FechaIngreso = IIf(Me.txtFechaIngreso.Text = SIGHEntidades.HORA_VACIA_HM, "", Me.txtFechaIngreso.Text)
           .idTipoServicio = lnIdTipoServicio
           .Edad = Me.txtEdadEnDias.Text
           .IdTipoEdad = Val(mo_cmbIdTipoEdad.BoundText)
           .idPaciente = mo_Pacientes.idPaciente
           .IdUsuarioAuditoria = ml_idUsuario
           .RecienNacido = 0
            .IdTipoCondicionALEstab = 1
            .IdTipoCondicionAlServicio = 1
            .FechaEgreso = 0
            .horaEgreso = SIGHEntidades.HORA_VACIA_HM
            .IdCondicionAlta = 0
            .IdTipoAlta = 0
            .TieneNecropsia = False
            .HuboInfeccionIntraHospitalaria = False
            .IdTipoGravedad = 0
            .IdFormaPago = Val(Me.cmbFormaPago.BoundText)
            .IdFuenteFinanciamiento = Val(cmbFuenteFinanciamiento.BoundText)
            .idCuentaAtencion = Val(Me.txtNroCuenta.Text)
            .IdEstadoAtencion = 1
            .EsPacienteExterno = True
            
   End With
   Set oRsTmp = Nothing
   '
   mo_Atenciones.RecienNacido = SIGHEntidades.CalculaSiEsRecienNacido(mo_Pacientes.FechaNacimiento, CDate(mo_Atenciones.FechaIngreso & " " & mo_Atenciones.HoraIngreso))
   '
End Sub


Private Sub btnBuscarPaciente_Click()
    Dim rsHistorias As New Recordset
    Dim oDOPaciente As New doPaciente
    On Error GoTo ErrBusq
    oDOPaciente.NroHistoriaClinica = Val(Me.txtNroHistoriaBusqueda.Text)
    oDOPaciente.ApellidoPaterno = Me.txtApellidoPaternoBusqueda
    oDOPaciente.ApellidoMaterno = Me.txtApellidoMaternoBusqueda
    oDOPaciente.PrimerNombre = Me.txtPrimerNombreBusqueda
    oDOPaciente.SegundoNombre = Me.txtSegundoNombreBusqueda
    oDOPaciente.IdDocIdentidad = 1
    oDOPaciente.NroDocumento = Me.txtNroDNIBusqueda
    If (oDOPaciente.ApellidoPaterno + oDOPaciente.ApellidoMaterno + _
    oDOPaciente.PrimerNombre + oDOPaciente.SegundoNombre = "") And _
    (Val(Me.txtNroHistoriaBusqueda.Text) = 0) And _
    (oDOPaciente.NroDocumento = "") Then
        MsgBox "Ingrese alguno de los valores de búsqueda", vbInformation, Me.Caption
        Exit Sub
    End If
    Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set rsHistorias = mo_AdminAdmision.PacientesFiltrarTodosSoloHistoriasDefinitivas(oDOPaciente)
    Set grdPacientesEncontrados.DataSource = rsHistorias
    With grdPacientesEncontrados
        .Left = 240
        .Top = 780
        .Width = 11775
        .Height = 4455
    End With
    ml_idPaciente = 0
    'Si hay una sola coincidencia
    If rsHistorias.RecordCount = 1 Then
        If mo_AdminAdmision.BuscaSiEstaHospitalizado(rsHistorias!idPaciente, oConexion) = False Then
            Me.grdPacientesEncontrados.Visible = False
            rsHistorias.MoveFirst
            chkPacienteNuevo.Value = 0
            ml_idPaciente = rsHistorias!idPaciente
            UcPacienteDatosAloj1.idPaciente = rsHistorias!idPaciente
            UcPacienteDatosAloj1.CargarDatosDePacienteALosControles
            CalculaEdadEnLaAtencion
            DeudasPendientesDeAnterioresAtenciones oConexion
            LimpiarBusqueda
            cmbServicioIngreso.SetFocus
        End If
    ElseIf rsHistorias.RecordCount > 1 Then
        Me.grdPacientesEncontrados.Visible = True
    ElseIf rsHistorias.RecordCount = 0 Then
        Me.grdPacientesEncontrados.Visible = False
        'LimpiarBusqueda
    End If
    oConexion.Close
    Set oConexion = Nothing
    mo_Apariencia.ConfigurarFilasBiColores Me.grdPacientesEncontrados, SIGHEntidades.GrillaConFilasBicolor
ErrBusq:
End Sub

Sub DeudasPendientesDeAnterioresAtenciones(oConexion As Connection)
        'Deudas
        ms_MensajeError = mo_AdminFacturacion.DevuelveDeudaPacienteDeAntencionesAnteriores(ml_idPaciente, oConexion, mo_CuentasAtencion.idCuentaAtencion)
        If ms_MensajeError <> "" Then
           MsgBox "Tiene Deudas Pendientes por Pagar" & Chr(13) & Chr(13) & ms_MensajeError, vbCritical, Me.Caption
           '
           ucMensajeParpadeando1.Visible = True
           ucMensajeParpadeando1.MensajeDeTexto = "Deudas:  " & ms_MensajeError
        Else
           '
           ucMensajeParpadeando1.Visible = False
           ucMensajeParpadeando1.MensajeDeTexto = ""
        End If
        ms_MensajeError = ""

End Sub

Private Sub btnCancelar_Click()
           Me.Visible = False
End Sub


Private Sub btnImprimePreCta_Click()
   If txtNroCuenta.Text <> "" Then
      ImprimePreCuenta
   End If
End Sub

Sub ImprimePreCuenta()
    Dim oReporte As New RptCaja
    Dim lcPaciente As String
    Dim lcMedico As String
    Dim lcCola As String
    If mi_Opcion <> sghAgregar Then
       Me.UcPacienteDatosAloj1.CargarDatosAlObjetoDatos mo_Pacientes
    End If
    lcPaciente = Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & " " & Trim(mo_Pacientes.PrimerNombre)
    If mo_Pacientes.SegundoNombre <> "" Then
       lcPaciente = lcPaciente & " " & Trim(mo_Pacientes.SegundoNombre)
    End If
    If mo_Pacientes.TercerNombre <> "" Then
      lcPaciente = lcPaciente & " " & Trim(mo_Pacientes.TercerNombre)
    End If
    lcMedico = lblNombreMedico.Text
    lcCola = ""
    oReporte.ImpresionPreCuenta Me.txtFechaIngreso.Text, Me.txtHoraIngreso.Text, lcPaciente, mo_Pacientes.NroHistoriaClinica, Me.cmbServicioIngreso.Text, lcMedico, "PACIENTE EXTERNO", mo_Atenciones.idAtencion, "", mo_Atenciones.idCuentaAtencion, Me.cmbFuenteFinanciamiento.Text, lcCola, ml_idUsuario, txtNroOrdenPago.Text, mo_Pacientes.FichaFamiliar
    Set oReporte = Nothing
    Me.Visible = False
End Sub


Private Sub chkPacienteNuevo_Click()
    If chkPacienteNuevo.Value = 1 Then
       UcPacienteDatosAloj1.ActualizaDatosBasicos Me.txtApellidoPaternoBusqueda.Text, Me.txtApellidoMaternoBusqueda.Text, Me.txtPrimerNombreBusqueda.Text, Me.txtSegundoNombreBusqueda.Text, "00:00", 0
       If Val(lcBuscaParametro.SeleccionaFilaParametro(255)) = sghHistoriaDefinitivaManual Then
          UcPacienteDatosAloj1.SetFocusOnHistoria
       Else
          UcPacienteDatosAloj1.SetFocusOnApellidoPaterno
       End If
    End If
End Sub


Private Sub cmbFuenteFinanciamiento_Click(Area As Integer)
        Set oRsFormaPago = mo_AdminFacturacion.TiposFinanciamientosTarifaSeleccionarPorPlan(Val(cmbFuenteFinanciamiento.BoundText))
        Set cmbFormaPago.RowSource = oRsFormaPago
        cmbFormaPago.ListField = "Descripcion"
        cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
        mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, True
        If oRsFormaPago.RecordCount = 1 Then
           cmbFormaPago.BoundText = oRsFormaPago.Fields!idTipoFinanciamiento
        ElseIf Val(cmbFuenteFinanciamiento.BoundText) = 5 Then
           cmbFormaPago.BoundText = "1"
        End If

End Sub

Private Sub cmbFuenteFinanciamiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbFuenteFinanciamiento
    AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbServicioIngreso_GotFocus()
   cmbServicioIngreso.SetFocus
End Sub

Private Sub cmbServicioIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbServicioIngreso
    AdministrarKeyPreview KeyCode

End Sub

Private Sub Form_Initialize()
    Set mo_cmbServicioIngreso.MiComboBox = cmbServicioIngreso
    Set mo_cmbIdTipoEdad.MiComboBox = cmbIdTipoEdad
End Sub

Private Sub Form_Load()
    mo_Formulario.HabilitarDeshabilitar txtIdMedicoIngreso, False
    mo_Formulario.HabilitarDeshabilitar txtNroCuenta, False
    mo_Formulario.HabilitarDeshabilitar Me.txtEdadEnDias, False
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoEdad, False
    mo_Formulario.HabilitarDeshabilitar txtNroOrdenPago, False
    GenerarRecordsetProductos
    CargaDataCombos
    CargaPreventas
    mo_Apariencia.ConfigurarFilasBiColores Me.grdPreVentaCab, SIGHEntidades.GrillaConFilasBicolor
    mo_Apariencia.ConfigurarFilasBiColores Me.grdPreVentaDet, SIGHEntidades.GrillaConFilasBicolor
    mo_Apariencia.ConfigurarFilasBiColores Me.grdProductos, SIGHEntidades.GrillaConFilasBicolor
    UcPacienteDatosAloj1.IdTipoGenHistoriaClinica = lcBuscaParametro.SeleccionaFilaParametro(255)
    UcPacienteDatosAloj1.Opcion = mi_Opcion
    UcPacienteDatosAloj1.Inicializar
    lnEspecialidadServicio = 0
    Me.txtFechaIngreso = Date
    Me.txtHoraIngreso = Format(Now, SIGHEntidades.DevuelveHoraSoloFormato_HM)
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Paciente Externo"
        If mo_lnIdPuntoCarga <> 99 Then    'Si se Agrega desde Admision y no desde CAJA
           UcPacienteDatosAloj1.HabilitaTipoHistoria True
        End If
    Case sghModificar
        Me.Caption = "Modificar Paciente Externo"
    Case sghConsultar
        Me.Caption = "Consultar Paciente Externo"
    Case sghEliminar
        Me.Caption = "Eliminar Paciente Externo"
    End Select
    CargarDatosAlFormulario
End Sub

Sub CargarDatosAlFormulario()
     Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
     End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   
   oRsPreVentaCab.Close
   oRsPreVentaDet.Close
   mrs_FacturacionProductos.Close
   oRsFuentesFinanciamiento.Close
   Set oRsPreVentaCab = Nothing
   Set oRsPreVentaDet = Nothing
   Set mrs_FacturacionProductos = Nothing
   Set oRsFuentesFinanciamiento = Nothing
End Sub

Private Sub grdPacientesEncontrados_DblClick()
    Dim rsPaciente As Recordset
    Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    On Error Resume Next
    Set rsPaciente = Me.grdPacientesEncontrados.DataSource
    If mo_AdminAdmision.BuscaSiEstaHospitalizado(rsPaciente!idPaciente, oConexion) = True Then
       Exit Sub
    End If
    Me.grdPacientesEncontrados.Visible = False
    chkPacienteNuevo.Value = 0
    ml_idPaciente = rsPaciente!idPaciente
    UcPacienteDatosAloj1.idPaciente = rsPaciente!idPaciente
    UcPacienteDatosAloj1.CargarDatosDePacienteALosControles
    CalculaEdadEnLaAtencion
    DeudasPendientesDeAnterioresAtenciones oConexion
    LimpiarBusqueda
    oConexion.Close
    Set oConexion = Nothing
    cmbServicioIngreso.SetFocus
End Sub

Private Sub grdPacientesEncontrados_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPacientesEncontrados.Bands(0).Columns("IdPaciente").Hidden = True
    grdPacientesEncontrados.Bands(0).Columns("IdTipoNumeracion").Hidden = True
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Hidden = True
    'grdPacientesEncontrados.Bands(0).Columns("FechaNacimiento").Hidden = True
    
    grdPacientesEncontrados.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "NroHistoria"
    grdPacientesEncontrados.Bands(0).Columns("NroHistoriaClinica").Width = 1000
    
    grdPacientesEncontrados.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdPacientesEncontrados.Bands(0).Columns("ApellidoPaterno").Width = 1200
    
    grdPacientesEncontrados.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdPacientesEncontrados.Bands(0).Columns("ApellidoMaterno").Width = 1200
    
    grdPacientesEncontrados.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdPacientesEncontrados.Bands(0).Columns("PrimerNombre").Width = 1200

    grdPacientesEncontrados.Bands(0).Columns("SegundoNombre").Header.Caption = "2do Nombre"
    grdPacientesEncontrados.Bands(0).Columns("SegundoNombre").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("FechaNacimiento").Header.Caption = "Fecha Nac."
    grdPacientesEncontrados.Bands(0).Columns("FechaNacimiento").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Header.Caption = "Tipo Numeración"
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Width = 1500
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").CellAppearance.TextAlign = ssAlignRight

    grdPacientesEncontrados.Bands(0).Columns("TipoServicio").Header.Caption = "Ult. Tipo Serv."
    grdPacientesEncontrados.Bands(0).Columns("TipoServicio").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("FechaIngreso").Header.Caption = "Ult.Fec.Ing"
    grdPacientesEncontrados.Bands(0).Columns("FechaIngreso").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("FechaEgreso").Header.Caption = "Ult.Fec.Egr."
    grdPacientesEncontrados.Bands(0).Columns("FechaEgreso").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("ServicioIngreso").Header.Caption = "Ult. Serv. Ing."
    grdPacientesEncontrados.Bands(0).Columns("ServicioIngreso").Width = 2700

End Sub

Private Sub grdPacientesEncontrados_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
        grdPacientesEncontrados_DblClick
    End If

End Sub














Private Sub txtApellidoMaternoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaternoBusqueda

End Sub

Private Sub txtApellidoMaternoBusqueda_LostFocus()
   txtApellidoMaternoBusqueda.Text = mo_Teclado.CapitalizarNombres(txtApellidoMaternoBusqueda.Text)
   If Len(txtApellidoMaternoBusqueda.Text) > 0 Then
      btnBuscarPaciente_Click
   End If

End Sub

Private Sub txtApellidoPaternoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaternoBusqueda
End Sub



Private Sub txtApellidoPaternoBusqueda_LostFocus()
   txtApellidoPaternoBusqueda.Text = mo_Teclado.CapitalizarNombres(txtApellidoPaternoBusqueda.Text)
   If Len(txtApellidoPaternoBusqueda.Text) > 0 Then
      btnBuscarPaciente_Click
   End If
End Sub



Private Sub txtFechaIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaIngreso
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaIngreso_LostFocus()
    CalculaEdadEnLaAtencion
End Sub

Sub CalculaEdadEnLaAtencion()
    On Error Resume Next
    Me.txtEdadEnDias.Text = ""
    Dim oEdad As Edad
    oEdad = SIGHEntidades.CalcularEdad(CDate(Me.UcPacienteDatosAloj1.FechaNacimiento), CDate(txtFechaIngreso.Text))
    Me.txtEdadEnDias.Text = oEdad.Edad
    mo_cmbIdTipoEdad.BoundText = oEdad.TipoEdad

End Sub



Private Sub txtHoraIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaIngreso
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroDNIBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtNroDNIBusqueda
End Sub


Private Sub txtNroDNIBusqueda_LostFocus()
   If Len(txtNroDNIBusqueda.Text) > 0 Then
      btnBuscarPaciente_Click
   End If

End Sub

Private Sub txtNroHistoriaBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoriaBusqueda

End Sub



Private Sub txtNroHistoriaBusqueda_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
    End If
    If KeyAscii = 13 And Len(txtNroHistoriaBusqueda.Text) > 0 Then
         btnBuscarPaciente_Click
    End If

End Sub

Private Sub txtPrimerNombreBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombreBusqueda

End Sub



Private Sub txtPrimerNombreBusqueda_LostFocus()
   txtPrimerNombreBusqueda.Text = mo_Teclado.CapitalizarNombres(txtPrimerNombreBusqueda.Text)
   'mo_Formulario.MarcarComoVacio txtPrimerNombreBusqueda
   If Len(txtPrimerNombreBusqueda.Text) > 0 Then
      btnBuscarPaciente_Click
   End If

End Sub

Private Sub txtSegundoNombreBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtSegundoNombreBusqueda
End Sub

Private Sub txtSegundoNombreBusqueda_LostFocus()
    txtSegundoNombreBusqueda.Text = mo_Teclado.CapitalizarNombres(txtSegundoNombreBusqueda.Text)
    'mo_Formulario.MarcarComoVacio txtSegundoNombreBusqueda
   If Len(txtSegundoNombreBusqueda.Text) > 0 Then
      btnBuscarPaciente_Click
   End If

End Sub


Sub CargaDataCombos()
    mo_cmbServicioIngreso.BoundColumn = "IdServicio"
    mo_cmbServicioIngreso.ListField = "DservicioHosp"
    Set mo_cmbServicioIngreso.RowSource = mo_AdminAdmision.DevuelveServiciosQueSonPuntosCarga("(1,2,3,4,5,6,7)", sghFiltraSoloActivos, sghPorDescServicio)
    '
    mo_cmbIdTipoEdad.BoundColumn = "IdTipoEdad"
    mo_cmbIdTipoEdad.ListField = "DescripcionLarga"
    Set mo_cmbIdTipoEdad.RowSource = mo_AdminServiciosComunes.TiposEdadSeleccionarTodos
    mo_cmbIdTipoEdad.BoundText = "1"    'Default Años
    '
    lcSql = "select * from FuentesFinanciamiento where esUsadoEnCaja=1"
    oRsFuentesFinanciamiento.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    Set cmbFuenteFinanciamiento.RowSource = oRsFuentesFinanciamiento
    cmbFuenteFinanciamiento.ListField = "Descripcion"
    cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"
    '
    oRsFormaPago.Open "select * from TiposFinanciamiento where esFuenteFinanciamiento=1 order by descripcion", SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    Set cmbFormaPago.RowSource = oRsFormaPago
    cmbFormaPago.ListField = "Descripcion"
    cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
    mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, False
End Sub

Private Sub lblNombreMedico_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, lblNombreMedico
   AdministrarKeyPreview KeyCode

End Sub

Private Sub lblNombreMedico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lbUltimaTeclaPulsoENTER = True
    Else
        lbUltimaTeclaPulsoENTER = False
    End If

End Sub

Private Sub lblNombreMedico_LostFocus()
        If lblNombreMedico.Locked = False And lbUltimaTeclaPulsoENTER = True And lnEspecialidadServicio > 0 Then
           lbUltimaTeclaPulsoENTER = False
           CompletarDatosDeMedico txtIdMedicoIngreso, lblNombreMedico, lnEspecialidadServicio, lblNombreMedico.Text
           On Error Resume Next
           txtFechaIngreso.SetFocus
        End If
End Sub

Sub CompletarDatosDeMedico(txtMedico As TextBox, lblNombreMedico As TextBox, lIdEspecialidad As Long, lcFiltraMedico As String)
Dim oBusqueda As New MedicosBusqueda
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient

    oBusqueda.IdEspecialidad = lIdEspecialidad
    If mi_Opcion = sghAgregar Then
        oBusqueda.NombreMedico = lcFiltraMedico
    End If
    oBusqueda.Show 1
    If oBusqueda.BotonPresionado = sghAceptar Then
       If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
            txtMedico.Text = oDOEmpleado.CodigoPlanilla
            txtMedico.Tag = oDoMedico.idMedico
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
       End If
    End If
    Set oBusqueda = Nothing
    Set oDoMedico = Nothing
    Set oDOEmpleado = Nothing
    Set oDOEspecialidades = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub


Private Sub cmbServicioIngreso_LostFocus()
   If mo_cmbServicioIngreso.BoundText <> "" Then
      Dim oRsTmp As New Recordset
      oRsTmp.Open "select idEspecialidad from Servicios where idServicio=" & mo_cmbServicioIngreso.BoundText, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
      If oRsTmp.RecordCount > 0 Then
         lnEspecialidadServicio = oRsTmp.Fields!IdEspecialidad
      End If
      oRsTmp.Close
      Set oRsTmp = Nothing
   End If
   On Error Resume Next
   lblNombreMedico.SetFocus
End Sub


Private Sub UcPacienteDatosAloj1_SePresionoTeclaEspecial(KeyCode As Integer)
    On Error Resume Next
    Select Case KeyCode
    Case vbKeyReturn
         CalculaEdadEnLaAtencion
         Dim oConexion As New Connection
         oConexion.Open SIGHEntidades.CadenaConexion
         oConexion.CursorLocation = adUseClient
         DeudasPendientesDeAnterioresAtenciones oConexion
         oConexion.Close
         Set oConexion = Nothing
         lbUltimaTeclaPulsoENTER = True
         cmbServicioIngreso.SetFocus
    Case Else
         AdministrarKeyPreview KeyCode
    End Select

End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        btnCancelar_Click
    Case vbKeyF2
        'btnAceptar_Click
    End Select
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------
Function AgregarDatos() As Boolean
    AgregarDatos = mo_AdminAdmision.AdmisionPacienteExternoParticularAgregar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mrs_FacturacionProductos, mo_lnIdPuntoCarga, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    ms_MensajeError = mo_AdminAdmision.MensajeError
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean
    ModificarDatos = mo_AdminAdmision.AdmisionPacienteExternoParticularModificar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mrs_FacturacionProductos, mo_lnIdPuntoCarga, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    ms_MensajeError = mo_AdminAdmision.MensajeError
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------
Function EliminarDatos() As Boolean
    Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    ms_MensajeError = mo_AdminAdmision.VerificaSiTieneMovimientoFarmaciaOservicio(mo_CuentasAtencion.idCuentaAtencion, mo_Atenciones.idTipoServicio, oConexion)
    oConexion.Close
    Set oConexion = Nothing
    If ms_MensajeError = "" Then
        mo_CuentasAtencion.idEstado = 9 'anulado
        mo_Atenciones.IdEstadoAtencion = 0  'anulado
        EliminarDatos = mo_AdminAdmision.AdmisionPacienteExternoParticularAnular(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
        ms_MensajeError = mo_AdminAdmision.MensajeError
    Else
        MsgBox ms_MensajeError & Chr(13) & "La Anulación tendrá que realizarlo FACTURACION ", vbCritical, "Consulta externa"
    End If
End Function


Sub CargarDatosALosControles()
Dim lcEstadoAtencion As String
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oRsTmp As New Recordset
Dim oDOEspecialidades As New Collection
Dim oConexion As New Connection
        oConexion.Open SIGHEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        fraBusqueda.Enabled = False
        '1do:   CARGAR DATOS DE LA ATENCION
        Set mo_Atenciones = mo_AdminAdmision.AtencionesSeleccionarPorId(Me.idAtencion, oConexion)
        If mo_Atenciones.idAtencion = 0 Then
            'El registro ha sido eliminado, pero no se hizo el refresh
             Exit Sub
        End If
        With mo_Atenciones
                mo_cmbServicioIngreso.BoundText = .IdServicioIngreso
                Me.txtIdMedicoIngreso.Tag = .idMedicoIngreso
                Me.txtHoraIngreso.Text = IIf(.HoraIngreso = "", SIGHEntidades.HORA_VACIA_HM, .HoraIngreso)
                Me.txtFechaIngreso.Text = IIf(.FechaIngreso = 0, SIGHEntidades.FECHA_VACIA_DMY, .FechaIngreso)
                Me.txtEdadEnDias.Text = .Edad
                Me.txtEdadEnDias.Tag = .Edad
                mo_cmbIdTipoEdad.BoundText = .IdTipoEdad
                cmbIdTipoEdad.Tag = .IdTipoEdad
'
                If mo_AdminProgramacion.MedicosSeleccionarPorId(.idMedicoIngreso, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
                    Me.txtIdMedicoIngreso = oDOEmpleado.CodigoPlanilla
                    Me.lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                Else
                    Me.lblNombreMedico = ""
                End If
                cmbFuenteFinanciamiento.BoundText = .IdFuenteFinanciamiento
                Me.cmbFormaPago.BoundText = .IdFormaPago
                Select Case .IdEstadoAtencion
                Case 0
                    lcEstadoAtencion = "Anulado"
                    btnAceptar.Enabled = False
                Case 1
                    lcEstadoAtencion = "Registrado"
                Case 2
                    lcEstadoAtencion = "Cerrado"
                    btnAceptar.Enabled = False
                End Select
        End With
        '
        oRsTmp.Open "select idEspecialidad from Servicios where idServicio=" & mo_cmbServicioIngreso.BoundText, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
        If oRsTmp.RecordCount > 0 Then
           lnEspecialidadServicio = oRsTmp.Fields!IdEspecialidad
        End If
        oRsTmp.Close
        '
        Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(mo_Atenciones.idCuentaAtencion, oConexion)
        lblEstadoCta.Caption = mo_ReglasFarmacia.DevuelveEstadoActualDeEstadoCuenta("idEstado=" & mo_CuentasAtencion.idEstado, oConexion)
        If mo_CuentasAtencion.idEstado <> 1 And mo_CuentasAtencion.idEstado <> 12 Then
            btnAceptar.Enabled = False
        End If
        txtNroCuenta.Text = mo_CuentasAtencion.idCuentaAtencion
        '3to:   CARGAR DATOS DEL PACIENTE
        UcPacienteDatosAloj1.idPaciente = mo_Atenciones.idPaciente
        UcPacienteDatosAloj1.CargarDatosDePacienteALosControles
        '
        DeudasPendientesDeAnterioresAtenciones oConexion
        '
        UcPacienteDatosAloj1.CargarDatosAlObjetoDatos mo_Pacientes
        Me.Caption = Trim(Me.Caption) & "                HC: " & Trim(mo_Pacientes.NroHistoriaClinica) & " " & Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & " " & Trim(mo_Pacientes.PrimerNombre) & "     (Estado: " & lcEstadoAtencion & ")"
        'Ya tuvo movimientos(Farmacia/servicios), no podrá cambiar de plan
        If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
            ms_MensajeError = mo_AdminAdmision.VerificaSiTieneMovimientoFarmaciaOservicio(mo_CuentasAtencion.idCuentaAtencion, mo_Atenciones.idTipoServicio, oConexion)
            If ms_MensajeError <> "" Then
               mo_Formulario.HabilitarDeshabilitar Me.cmbFuenteFinanciamiento, False
               Me.ucMensajeParpadeando1.MensajeDeTexto = ms_MensajeError
               Me.ucMensajeParpadeando1.Visible = True
               Me.btnAceptar.Enabled = False
            End If
            ms_MensajeError = ""
        End If
        'Carga Preventas
        lcSql = "SELECT     dbo.FactPreventa.idFactPreventa, dbo.FacturacionPreventa.idProducto, dbo.FactPreventa.idServicio, dbo.Servicios.Nombre AS Servicio, " & _
                "      dbo.FactCatalogoServicios.Codigo, dbo.FactCatalogoServicios.Nombre AS Producto, dbo.FacturacionPreventa.Cantidad," & _
                "      dbo.FacturacionPreventa.Precio, dbo.FacturacionPreventa.Importe, dbo.FactPreventa.idAtencion, dbo.FactPreventa.idOrden," & _
                "      dbo.FactPreventa.idEstadoPreventa" & _
                " FROM         dbo.FactCatalogoServicios RIGHT OUTER JOIN" & _
                "      dbo.FacturacionPreventa ON dbo.FactCatalogoServicios.IdProducto = dbo.FacturacionPreventa.idProducto LEFT OUTER JOIN" & _
                "      dbo.FactPreventa LEFT OUTER JOIN" & _
                "      dbo.Servicios ON dbo.FactPreventa.idServicio = dbo.Servicios.IdServicio ON" & _
                "      dbo.FacturacionPreventa.idFactPreventa = dbo.FactPreventa.idFactPreventa" & _
                " where dbo.FactPreventa.idAtencion=" & mo_Atenciones.idAtencion
        oRsTmp.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
                mrs_FacturacionProductos.AddNew
                mrs_FacturacionProductos.Fields!idPreVenta = oRsTmp.Fields!idFactPreVenta
                mrs_FacturacionProductos.Fields!idProducto = oRsTmp.Fields!idProducto
                mrs_FacturacionProductos.Fields!idServicio = oRsTmp.Fields!idServicio
                mrs_FacturacionProductos.Fields!Servicio = oRsTmp.Fields!Servicio
                mrs_FacturacionProductos.Fields!Codigo = oRsTmp.Fields!Codigo
                mrs_FacturacionProductos.Fields!Producto = oRsTmp.Fields!Producto
                mrs_FacturacionProductos.Fields!Cantidad = oRsTmp.Fields!Cantidad
                mrs_FacturacionProductos.Fields!Precio = oRsTmp.Fields!Precio
                mrs_FacturacionProductos.Fields!Total = oRsTmp.Fields!Importe
                mrs_FacturacionProductos.Update
                oRsTmp.MoveNext
           Loop
           TotalizaProductos
        End If
        oRsTmp.Close
        txtNroOrdenPago.Text = mo_AdminFacturacion.DevuelveOrdenesPagoSegunCuenta(mo_Atenciones.idCuentaAtencion, oConexion)
        Set oRsTmp = Nothing
        Set oDoMedico = Nothing
        Set oDOEmpleado = Nothing
        Set oDOEspecialidades = Nothing
        oConexion.Close
        Set oConexion = Nothing
End Sub


Private Sub grdPreVentaCab_Click()
        On Error GoTo errDet
        lcSql = "SELECT      dbo.FacturacionPreventa.idFactPreventa,dbo.FacturacionPreventa.idProducto, dbo.FactCatalogoServicios.Codigo, dbo.FactCatalogoServicios.Nombre, " & _
                "          dbo.FacturacionPreventa.Cantidad , dbo.FacturacionPreventa.Precio, dbo.FacturacionPreventa.Importe" & _
                " FROM         dbo.FacturacionPreventa LEFT OUTER JOIN" & _
                "          dbo.FactCatalogoServicios ON dbo.FacturacionPreventa.idProducto = dbo.FactCatalogoServicios.IdProducto" & _
                " WHERE     dbo.FacturacionPreventa.idFactPreventa=" & oRsPreVentaCab.Fields!Preventa
        oRsPreVentaDet.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
        Set Me.grdPreVentaDet.DataSource = oRsPreVentaDet
        Exit Sub
errDet:
        If Err.Number = 3705 Then
           oRsPreVentaDet.Close
           Resume
        End If
End Sub



Private Sub grdPreVentaCab_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPreVentaCab.Bands(0).Columns("idServicio").Hidden = True
    grdPreVentaCab.Bands(0).Columns("idTipoFinanciamiento").Hidden = True
    grdPreVentaCab.Bands(0).Columns("PreVenta").Activation = ssActivationActivateNoEdit
    grdPreVentaCab.Bands(0).Columns("PreVenta").Width = 700
    grdPreVentaCab.Bands(0).Columns("FechaCreacion").Width = 800
    grdPreVentaCab.Bands(0).Columns("FechaCreacion").Activation = ssActivationActivateNoEdit
    grdPreVentaCab.Bands(0).Columns("Servicio").Width = 1300
    grdPreVentaCab.Bands(0).Columns("Servicio").Activation = ssActivationActivateNoEdit
    grdPreVentaCab.Bands(0).Columns("Total").Width = 600
    grdPreVentaCab.Bands(0).Columns("Total").Format = "#0.00"
    grdPreVentaCab.Bands(0).Columns("Total").Activation = ssActivationActivateNoEdit
End Sub

Private Sub grdPreVentaCab_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       Dim oRsTmp As New Recordset
       grdPreVentaCab_Click
       'Elimina anterior detalle
       If mrs_FacturacionProductos.RecordCount > 0 Then
          mrs_FacturacionProductos.MoveFirst
          Do While Not mrs_FacturacionProductos.EOF
             If mrs_FacturacionProductos.Fields!idPreVenta = oRsPreVentaCab.Fields!Preventa Then
                mrs_FacturacionProductos.Delete
                mrs_FacturacionProductos.Update
             End If
             mrs_FacturacionProductos.MoveNext
          Loop
       End If
       '
       If mrs_FacturacionProductos.RecordCount = 0 Then
          '*************Actualiza datos de la ficha Atención, siempre y cuando no tenga ITEMS
          oRsPreVentaDet.MoveFirst
          mo_cmbServicioIngreso.BoundText = oRsPreVentaCab.Fields!idServicio
          Set oRsTmp = mo_AdminProgramacion.DevuelveListaMedicosSegunServicio(oRsPreVentaCab.Fields!idServicio)
          If oRsTmp.RecordCount > 0 Then
             Me.txtIdMedicoIngreso = oRsTmp.Fields!CodigoPlanilla
             Me.lblNombreMedico = oRsTmp.Fields!ApellidoPaterno + " " + oRsTmp.Fields!ApellidoMaterno + " " + oRsTmp.Fields!Nombres
             Me.txtIdMedicoIngreso.Tag = oRsTmp.Fields!idMedico
          Else
             Me.lblNombreMedico = ""
             Me.txtIdMedicoIngreso.Tag = 0
          End If
          oRsTmp.Close
          cmbFuenteFinanciamiento.BoundText = "1"
          cmbFormaPago.BoundText = oRsPreVentaCab.Fields!idTipoFinanciamiento
       Else
          '*************Verifica que la Nueva PreVenta tenga el mismo TARIFARIO
          If cmbFormaPago.BoundText <> oRsPreVentaCab.Fields!idTipoFinanciamiento Then
             MsgBox "La PreVenta no puede agregarse porque tiene otro TIPO DE FINANCIAMIENTO", vbCritical, Me.Caption
             Exit Sub
          End If
       End If
       '*************Vuelve a añadir
       oRsPreVentaDet.MoveFirst
       Do While Not oRsPreVentaDet.EOF
            mrs_FacturacionProductos.AddNew
            mrs_FacturacionProductos.Fields!idPreVenta = oRsPreVentaCab.Fields!Preventa
            mrs_FacturacionProductos.Fields!idProducto = oRsPreVentaDet.Fields!idProducto
            mrs_FacturacionProductos.Fields!idServicio = oRsPreVentaCab.Fields!idServicio
            mrs_FacturacionProductos.Fields!Servicio = oRsPreVentaCab.Fields!Servicio
            mrs_FacturacionProductos.Fields!Codigo = oRsPreVentaDet.Fields!Codigo
            mrs_FacturacionProductos.Fields!Producto = oRsPreVentaDet.Fields!Nombre
            mrs_FacturacionProductos.Fields!Cantidad = oRsPreVentaDet.Fields!Cantidad
            mrs_FacturacionProductos.Fields!Precio = oRsPreVentaDet.Fields!Precio
            mrs_FacturacionProductos.Fields!Total = oRsPreVentaDet.Fields!Importe
            mrs_FacturacionProductos.Update
            oRsPreVentaDet.MoveNext
       Loop
       TotalizaProductos
       Set oRsTmp = Nothing
    End If
End Sub

Private Sub grdPreVentaDet_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPreVentaDet.Bands(0).Columns("idFactPreventa").Hidden = True
    grdPreVentaDet.Bands(0).Columns("idProducto").Hidden = True
    grdPreVentaDet.Bands(0).Columns("Codigo").Width = 500
    grdPreVentaDet.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    grdPreVentaDet.Bands(0).Columns("Nombre").Width = 2200
    grdPreVentaDet.Bands(0).Columns("Nombre").Activation = ssActivationActivateNoEdit
    grdPreVentaDet.Bands(0).Columns("Cantidad").Width = 350
    grdPreVentaDet.Bands(0).Columns("Cantidad").Format = "###0"
    grdPreVentaDet.Bands(0).Columns("Cantidad").Activation = ssActivationActivateNoEdit
    grdPreVentaDet.Bands(0).Columns("Precio").Width = 600
    grdPreVentaDet.Bands(0).Columns("Precio").Format = "#0.00"
    grdPreVentaDet.Bands(0).Columns("Precio").Activation = ssActivationActivateNoEdit
    grdPreVentaDet.Bands(0).Columns("Importe").Activation = ssActivationActivateNoEdit
End Sub


Sub GenerarRecordsetProductos()
    With mrs_FacturacionProductos
          .Fields.Append "IdPreventa", adInteger
          .Fields.Append "IdProducto", adInteger
          .Fields.Append "idServicio", adInteger
          .Fields.Append "Servicio", adVarChar, 255, adFldIsNullable
          .Fields.Append "Codigo", adVarChar, 255, adFldIsNullable
          .Fields.Append "Producto", adVarChar, 255, adFldIsNullable
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Precio", adCurrency
          .Fields.Append "Total", adCurrency
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
          .Sort = "idPreventa desc"
    End With
    Set grdProductos.DataSource = mrs_FacturacionProductos
End Sub
Sub CargaPreventas()
    lcSql = "SELECT      dbo.FactPreventa.idFactPreventa AS PreVenta, dbo.FactPreventa.fechaCreacion, dbo.Servicios.Nombre AS Servicio, " & _
            "                      dbo.FactPreventa.Total, dbo.Servicios.IdServicio, dbo.TiposFinanciamiento.Descripcion AS TipoFinanciamiento," & _
            "            dbo.FactPreventa.idTipoFinanciamiento  " & _
            " FROM         dbo.FactPreventa LEFT OUTER JOIN" & _
            "                      dbo.TiposFinanciamiento ON dbo.FactPreventa.idTipoFinanciamiento = dbo.TiposFinanciamiento.IdTipoFinanciamiento LEFT OUTER JOIN" & _
            "                      dbo.Servicios ON dbo.FactPreventa.idServicio = dbo.Servicios.IdServicio" & _
            " WHERE dbo.FactPreventa.idEstadoPreventa=1" & _
            " ORDER BY dbo.FactPreventa.idFactPreventa DESC"
    oRsPreVentaCab.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    Set Me.grdPreVentaCab.DataSource = oRsPreVentaCab
    grdPreVentaCab_Click
End Sub


Sub TotalizaProductos()
        Dim lnTotal As Double
        lnTotal = 0
        mrs_FacturacionProductos.MoveFirst
        Do While Not mrs_FacturacionProductos.EOF
           lnTotal = lnTotal + mrs_FacturacionProductos.Fields!Total
           mrs_FacturacionProductos.MoveNext
        Loop
        Me.lblTotal.Caption = Format(lnTotal, "####,###,##0.00")
End Sub


Private Sub grdProductos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdProductos.Bands(0).Columns("idProducto").Hidden = True
    grdProductos.Bands(0).Columns("idServicio").Hidden = True
    grdProductos.Bands(0).Columns("idPreventa").Width = 1000
    grdProductos.Bands(0).Columns("idPreventa").Activation = ssActivationActivateNoEdit
    grdProductos.Bands(0).Columns("idPreventa").Header.Caption = "Preventa"
    grdProductos.Bands(0).Columns("Servicio").Width = 2000
    grdProductos.Bands(0).Columns("Servicio").Activation = ssActivationActivateNoEdit
    grdProductos.Bands(0).Columns("Codigo").Width = 800
    grdProductos.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    grdProductos.Bands(0).Columns("Producto").Width = 3800
    grdProductos.Bands(0).Columns("Producto").Activation = ssActivationActivateNoEdit
    grdProductos.Bands(0).Columns("Cantidad").Width = 600
    grdProductos.Bands(0).Columns("Cantidad").Activation = ssActivationActivateNoEdit
    grdProductos.Bands(0).Columns("Cantidad").Format = "###0"
    grdProductos.Bands(0).Columns("Precio").Width = 700
    grdProductos.Bands(0).Columns("Precio").Activation = ssActivationActivateNoEdit
    grdProductos.Bands(0).Columns("Precio").Format = "#0.00"
    grdProductos.Bands(0).Columns("Total").Width = 1000
    grdProductos.Bands(0).Columns("Total").Activation = ssActivationActivateNoEdit
    grdProductos.Bands(0).Columns("Total").Format = "#0.00"
End Sub
Private Sub grdProductos_AfterRowsDeleted()
    TotalizaProductos
End Sub
