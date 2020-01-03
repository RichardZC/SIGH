VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form FacReembolsos 
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13875
   Icon            =   "FacReembolsos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   13875
   StartUpPosition =   2  'CenterScreen
   Begin UltraGrid.SSUltraGrid grdPacientesEncontrados 
      Height          =   2265
      Left            =   4140
      TabIndex        =   13
      Top             =   1050
      Visible         =   0   'False
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   3995
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
      Caption         =   "SSUltraGrid1"
   End
   Begin VB.Frame fraDatosHistoria 
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
      Height          =   1905
      Left            =   60
      TabIndex        =   14
      Top             =   1170
      Width           =   13725
      Begin VB.Frame frmFarmacia 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   90
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   10545
         Begin MSDataListLib.DataCombo cmbFormaPago 
            Height          =   360
            Left            =   7230
            TabIndex        =   23
            Top             =   300
            Width           =   3045
            _ExtentX        =   5371
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
         Begin MSDataListLib.DataCombo cmbFarmacia 
            Height          =   360
            Left            =   1020
            TabIndex        =   24
            Top             =   300
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Farmacia"
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
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pago"
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
            Left            =   5880
            TabIndex        =   25
            Top             =   360
            Width           =   1305
         End
      End
      Begin VB.TextBox txtIdOrden 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4230
         MaxLength       =   30
         TabIndex        =   18
         Top             =   510
         Width           =   1365
      End
      Begin VB.ComboBox cmbFechaIngreso 
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
         Left            =   5700
         TabIndex        =   17
         Top             =   540
         Width           =   2610
      End
      Begin VB.TextBox txtPaciente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         MaxLength       =   30
         TabIndex        =   16
         Top             =   510
         Width           =   4125
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
         Left            =   8430
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   540
         Width           =   2805
      End
      Begin MSMask.MaskEdBox txtHentrega 
         Height          =   315
         Left            =   12750
         TabIndex        =   27
         Top             =   540
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
         Left            =   11370
         TabIndex        =   28
         Top             =   540
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   $"FacReembolsos.frx":000C
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
         TabIndex        =   19
         Top             =   210
         Width           =   13290
      End
   End
   Begin GalenHos.ucFacturacionItems ucFacturacionProductos 
      Height          =   4335
      Left            =   90
      TabIndex        =   5
      Top             =   3120
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   7805
   End
   Begin VB.Frame fraDatosAtencion 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   13725
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
         Height          =   315
         Left            =   9180
         TabIndex        =   30
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox txtNroCuenta 
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
         Left            =   10740
         TabIndex        =   12
         Top             =   630
         Width           =   1545
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
         Left            =   6645
         TabIndex        =   11
         Top             =   630
         Width           =   1260
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
         Left            =   5460
         TabIndex        =   10
         Top             =   630
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
         Left            =   7905
         TabIndex        =   9
         Top             =   630
         Width           =   1230
      End
      Begin VB.CommandButton btnBuscarPaciente 
         Height          =   315
         Left            =   12360
         Picture         =   "FacReembolsos.frx":00C2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   630
         Width           =   1305
      End
      Begin VB.TextBox txtNroHistoria 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4080
         MaxLength       =   9
         TabIndex        =   4
         Top             =   630
         Width           =   1365
      End
      Begin VB.ComboBox cmbIdTipoGenHistoriaClinica 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1845
      End
      Begin Threed.SSOption optConHistoriaClinica 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   330
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Con Historia Clínica"
         Value           =   -1
      End
      Begin Threed.SSOption optSinHistoriaClinica 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sin Historia Clínica"
      End
      Begin VB.Label Label50 
         Caption         =   $"FacReembolsos.frx":2D0B
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
         Left            =   2370
         TabIndex        =   8
         Top             =   300
         Width           =   9855
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
         Left            =   7380
         TabIndex        =   6
         Top             =   1290
         Width           =   2205
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   90
      TabIndex        =   0
      Top             =   7590
      Width           =   13680
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FacReembolsos.frx":2D97
         DownPicture     =   "FacReembolsos.frx":31F7
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
         Picture         =   "FacReembolsos.frx":366C
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FacReembolsos.frx":3AE1
         DownPicture     =   "FacReembolsos.frx":3FA5
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
         Picture         =   "FacReembolsos.frx":4491
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FacReembolsos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
