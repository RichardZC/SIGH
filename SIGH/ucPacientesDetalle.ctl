VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ucPacientesDetalle 
   ClientHeight    =   6585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11760
   ScaleHeight     =   6585
   ScaleWidth      =   11760
   Begin VB.CommandButton cmdAcreditaSIS 
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
      Picture         =   "ucPacientesDetalle.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   136
      ToolTipText     =   "Agrega AFILIADO SIS que no se pudo por la WEB SIS pero el AREA SIS lo autorizó"
      Top             =   4635
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txtFichaFamiliar1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7680
      MaxLength       =   4
      TabIndex        =   129
      Top             =   4635
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txtFichaFamiliar2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8130
      MaxLength       =   7
      TabIndex        =   128
      Top             =   4635
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtFichaFamiliar3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8835
      MaxLength       =   2
      TabIndex        =   127
      Top             =   4635
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9810
      TabIndex        =   117
      Top             =   3690
      Width           =   1875
      Begin VB.CommandButton cmdArchivo 
         DisabledPicture =   "ucPacientesDetalle.ctx":058A
         DownPicture     =   "ucPacientesDetalle.ctx":0973
         Height          =   330
         Left            =   1350
         Picture         =   "ucPacientesDetalle.ctx":0D7F
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Registrar Imagen"
         Top             =   840
         Width           =   435
      End
      Begin VB.CommandButton btnQuitaIMG 
         DisabledPicture =   "ucPacientesDetalle.ctx":118B
         DownPicture     =   "ucPacientesDetalle.ctx":1516
         Height          =   330
         Left            =   1350
         Picture         =   "ucPacientesDetalle.ctx":18A9
         Style           =   1  'Graphical
         TabIndex        =   120
         ToolTipText     =   "Eliminar Imagen"
         Top             =   480
         Width           =   435
      End
      Begin VB.CommandButton btnAgreaIMG 
         DisabledPicture =   "ucPacientesDetalle.ctx":1C3A
         DownPicture     =   "ucPacientesDetalle.ctx":2023
         Height          =   330
         Left            =   1350
         Picture         =   "ucPacientesDetalle.ctx":242F
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   "Registrar Imagen"
         Top             =   135
         Width           =   435
      End
      Begin VB.Image pi_ImagSeleccionada 
         BorderStyle     =   1  'Fixed Single
         Height          =   1170
         Left            =   15
         MouseIcon       =   "ucPacientesDetalle.ctx":28B1
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2340
         TabIndex        =   118
         Top             =   600
         Width           =   60
      End
   End
   Begin VB.Frame fraSector 
      Caption         =   "Datos del Sector y Sectorista"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   6330
      TabIndex        =   109
      Top             =   3690
      Width           =   3480
      Begin VB.CommandButton cmdSectorista 
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
         Left            =   1935
         Picture         =   "ucPacientesDetalle.ctx":2BBB
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   570
         Width           =   450
      End
      Begin VB.TextBox txtSector 
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
         Left            =   1350
         MaxLength       =   2
         TabIndex        =   93
         Top             =   240
         Width           =   540
      End
      Begin VB.TextBox txtSectorista 
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
         Left            =   1350
         MaxLength       =   20
         TabIndex        =   94
         Top             =   570
         Width           =   555
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sector"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   112
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sectorista"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   111
         Top             =   660
         Width           =   810
      End
      Begin VB.Label lblSectorista 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2340
         TabIndex        =   110
         Top             =   600
         Width           =   60
      End
   End
   Begin VB.Frame fraMadre 
      Caption         =   "Datos de la MADRE o tutor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   101
      Top             =   3690
      Width           =   6315
      Begin VB.TextBox txtNombreMadre 
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
         MaxLength       =   20
         TabIndex        =   25
         Top             =   870
         Width           =   1725
      End
      Begin VB.TextBox txtMadreApellidoP 
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
         MaxLength       =   20
         TabIndex        =   23
         Top             =   540
         Width           =   1725
      End
      Begin VB.TextBox txtMadreApellidoM 
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
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   24
         Top             =   540
         Width           =   1455
      End
      Begin VB.TextBox txtMadreSnombre 
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
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   26
         Top             =   870
         Width           =   1455
      End
      Begin VB.TextBox txtMadreDocumento 
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
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   22
         Top             =   180
         Width           =   1455
      End
      Begin VB.ComboBox cmbMadreTipoDocumento 
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
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   210
         Width           =   1725
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Primer Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   107
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Apell.Paterno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   106
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Apell.Materno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3660
         TabIndex        =   105
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Segundo Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3360
         TabIndex        =   104
         Top             =   930
         Width           =   1440
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "N° Documento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3570
         TabIndex        =   103
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   102
         Top             =   270
         Width           =   960
      End
   End
   Begin VB.Frame fraNN 
      Height          =   975
      Left            =   9150
      TabIndex        =   76
      Top             =   0
      Width           =   2565
      Begin VB.CheckBox chkNN 
         Caption         =   "No identificado (N.N.)"
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
         Left            =   210
         TabIndex        =   115
         Top             =   240
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label lblFichaFamiliar1 
         Caption         =   "..........FichaFamiliar..........."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   90
         TabIndex        =   116
         Top             =   570
         Visible         =   0   'False
         Width           =   2400
      End
   End
   Begin VB.Frame fraDatosHistoriaClinica 
      Caption         =   "Datos de la Historia Clínica"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   -15
      TabIndex        =   75
      Top             =   15
      Width           =   9135
      Begin VB.TextBox txtGs 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   7095
         MaxLength       =   2
         TabIndex        =   132
         Top             =   600
         Width           =   540
      End
      Begin VB.TextBox txtFRh 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   8175
         MaxLength       =   20
         TabIndex        =   131
         Top             =   600
         Width           =   870
      End
      Begin VB.CommandButton cmdCambiaHC 
         DisabledPicture =   "ucPacientesDetalle.ctx":3145
         DownPicture     =   "ucPacientesDetalle.ctx":352E
         Height          =   330
         Left            =   6000
         Picture         =   "ucPacientesDetalle.ctx":393A
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   "Cambia el NUMERO de HISTORIA"
         Top             =   600
         Width           =   270
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
         Height          =   330
         Left            =   4590
         MaxLength       =   8
         TabIndex        =   0
         Top             =   255
         Width           =   1665
      End
      Begin VB.ComboBox cmbIdDocIdentidad 
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
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   240
         Width           =   2655
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
         Left            =   1530
         TabIndex        =   97
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtIdNroHistoria 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4590
         MaxLength       =   9
         TabIndex        =   1
         Top             =   600
         Width           =   1395
      End
      Begin MSMask.MaskEdBox txtFechaCreacion 
         Height          =   330
         Left            =   7650
         TabIndex        =   99
         Top             =   225
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
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
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Gs"
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
         Left            =   6825
         TabIndex        =   134
         Top             =   630
         Width           =   210
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "F.Rh"
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
         Left            =   7755
         TabIndex        =   133
         Top             =   630
         Width           =   405
      End
      Begin VB.Label Label11 
         Caption         =   "Nº"
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
         Left            =   4350
         TabIndex        =   92
         Top             =   630
         Width           =   195
      End
      Begin VB.Label Label36 
         Caption         =   "Nº"
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
         Left            =   4350
         TabIndex        =   87
         Top             =   270
         Width           =   195
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Doc&umento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   86
         Top             =   300
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Creación:"
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
         Left            =   6180
         TabIndex        =   47
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Nro &Historia:"
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
         TabIndex        =   45
         Top             =   630
         Width           =   1095
      End
   End
   Begin VB.Frame fraDatosPaciente 
      Caption         =   "Datos del Paciente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   0
      TabIndex        =   71
      Top             =   960
      Width           =   11715
      Begin VB.CommandButton cmdSinApellidoMaterno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5955
         Picture         =   "ucPacientesDetalle.ctx":3D46
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   210
         Width           =   315
      End
      Begin VB.CommandButton cmdSinApellidoPaterno 
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
         Left            =   2910
         Picture         =   "ucPacientesDetalle.ctx":42D0
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   240
         Width           =   315
      End
      Begin VB.ComboBox cboTipoEdadPaciente 
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
         ItemData        =   "ucPacientesDetalle.ctx":485A
         Left            =   8085
         List            =   "ucPacientesDetalle.ctx":485C
         TabIndex        =   91
         Top             =   990
         Width           =   1425
      End
      Begin VB.CheckBox chkSinFechaNacimiento 
         Caption         =   "Calcular Fecha Nacimiento"
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
         Left            =   3315
         TabIndex        =   114
         Top             =   1380
         Width           =   2895
      End
      Begin VB.TextBox txtNroHijo 
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
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1650
         Width           =   375
      End
      Begin VB.TextBox txtApellidoMaterno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4455
         MaxLength       =   40
         TabIndex        =   3
         Top             =   210
         Width           =   1485
      End
      Begin VB.TextBox txtSegundoNombre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4455
         MaxLength       =   40
         TabIndex        =   5
         Top             =   630
         Width           =   1785
      End
      Begin VB.TextBox txtEmail 
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
         Left            =   7680
         MaxLength       =   50
         TabIndex        =   19
         Top             =   2010
         Width           =   3975
      End
      Begin VB.TextBox txtIdPaciente 
         BackColor       =   &H00FBF7F4&
         Enabled         =   0   'False
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
         Left            =   10800
         MaxLength       =   9
         TabIndex        =   88
         Top             =   210
         Width           =   825
      End
      Begin VB.TextBox txtEdad 
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
         Left            =   7680
         MaxLength       =   3
         TabIndex        =   89
         Top             =   990
         Width           =   405
      End
      Begin VB.TextBox txtObservacion 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1530
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   2355
         Width           =   4725
      End
      Begin VB.ComboBox cmbIdTipoSexo 
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
         Left            =   1530
         TabIndex        =   9
         Top             =   1650
         Width           =   1695
      End
      Begin VB.TextBox txtNombrePadre 
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
         Left            =   7680
         MaxLength       =   20
         TabIndex        =   20
         Top             =   2340
         Width           =   3975
      End
      Begin VB.TextBox txtTelefono 
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
         Left            =   10320
         MaxLength       =   10
         TabIndex        =   16
         Top             =   960
         Width           =   1305
      End
      Begin VB.TextBox txtApellidoPaterno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1530
         MaxLength       =   40
         TabIndex        =   2
         Top             =   225
         Width           =   1380
      End
      Begin VB.TextBox txtPrimerNombre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1530
         MaxLength       =   40
         TabIndex        =   4
         Top             =   618
         Width           =   1695
      End
      Begin VB.TextBox txtTercerNombre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1530
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1011
         Width           =   1695
      End
      Begin MSMask.MaskEdBox txtFechaNacimiento 
         Height          =   330
         Left            =   4200
         TabIndex        =   7
         Top             =   1035
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
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
      Begin MSMask.MaskEdBox txtHoraNacimiento 
         Height          =   330
         Left            =   5535
         TabIndex        =   8
         Top             =   1035
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   582
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
      Begin PVCOMBOLibCtl.PVComboBox cmbEtnia 
         Height          =   330
         Left            =   1530
         TabIndex        =   11
         Top             =   2010
         Width           =   1695
         _Version        =   524288
         _cx             =   2990
         _cy             =   582
         Appearance      =   1
         Enabled         =   -1  'True
         BackColor       =   16777215
         ForeColor       =   0
         Locked          =   0   'False
         Style           =   0
         Sorted          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowPictures    =   0   'False
         ColumnHeaders   =   -1  'True
         PrimaryColumn   =   1
         VisibleItems    =   10
         ColumnHeaderHeight=   20
         ListMember      =   ""
         ColumnHeaderForeColor=   0
         ColumnHeaderBackColor=   13160660
         SelectedForeColor=   16777215
         SelectedBackColor=   6956042
         AlternateBackColor=   16777215
         ItemLabelStyle  =   1
         ItemLabelType   =   0
         ItemLabelWidth  =   40
         ItemLabelForeColor=   0
         ItemLabelBackColor=   13160660
         ColumnHeaderStyle=   1
         VerticalGridLines=   -1  'True
         HorizontalGridLines=   -1  'True
         ColumnResize    =   0   'False
         ItemLabelResize =   0   'False
         AllowDBAutoConfig=   0   'False
         GridLineColor   =   13421772
         List            =   ""
         NullString      =   "[NULL]"
         DropShadow      =   -1  'True
         Text            =   ""
         SortOnColumnHeaderClick=   0   'False
         DropEffect      =   1
         ColumnCount     =   2
         Column0.Heading =   "Id"
         Column0.Width   =   20
         Column0.Alignment=   0
         Column0.Hidden  =   -1  'True
         Column0.Name    =   "codetni"
         Column0.Format  =   ""
         Column0.Bound   =   -1  'True
         Column0.Locked  =   0   'False
         Column0.HeaderAlignment=   0
         Column1.Heading =   "Etnia"
         Column1.Width   =   50
         Column1.Alignment=   0
         Column1.Hidden  =   0   'False
         Column1.Name    =   "dCorto"
         Column1.Format  =   ""
         Column1.Bound   =   -1  'True
         Column1.Locked  =   0   'False
         Column1.HeaderAlignment=   0
         SortKey1.Column =   -1
         SortKey1.Ascending=   -1  'True
         SortKey1.CaseInsensitive=   -1  'True
         SortKey2.Column =   -1
         SortKey2.Ascending=   -1  'True
         SortKey2.CaseInsensitive=   -1  'True
         SortKey3.Column =   -1
         SortKey3.Ascending=   -1  'True
         SortKey3.CaseInsensitive=   -1  'True
         BoundColumn     =   ""
         Border          =   -1  'True
         VertAlign       =   1
         Format          =   ""
      End
      Begin PVCOMBOLibCtl.PVComboBox cmbIdioma 
         Height          =   330
         Left            =   4650
         TabIndex        =   12
         Top             =   2010
         Width           =   1605
         _Version        =   524288
         _cx             =   2831
         _cy             =   582
         Appearance      =   1
         Enabled         =   -1  'True
         BackColor       =   16777215
         ForeColor       =   0
         Locked          =   0   'False
         Style           =   0
         Sorted          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowPictures    =   0   'False
         ColumnHeaders   =   -1  'True
         PrimaryColumn   =   1
         VisibleItems    =   10
         ColumnHeaderHeight=   20
         ListMember      =   ""
         ColumnHeaderForeColor=   0
         ColumnHeaderBackColor=   13160660
         SelectedForeColor=   16777215
         SelectedBackColor=   6956042
         AlternateBackColor=   16777215
         ItemLabelStyle  =   1
         ItemLabelType   =   0
         ItemLabelWidth  =   40
         ItemLabelForeColor=   0
         ItemLabelBackColor=   13160660
         ColumnHeaderStyle=   1
         VerticalGridLines=   -1  'True
         HorizontalGridLines=   -1  'True
         ColumnResize    =   0   'False
         ItemLabelResize =   0   'False
         AllowDBAutoConfig=   0   'False
         GridLineColor   =   13421772
         List            =   ""
         NullString      =   "[NULL]"
         DropShadow      =   -1  'True
         Text            =   ""
         SortOnColumnHeaderClick=   0   'False
         DropEffect      =   1
         ColumnCount     =   2
         Column0.Heading =   "Id"
         Column0.Width   =   20
         Column0.Alignment=   0
         Column0.Hidden  =   -1  'True
         Column0.Name    =   "IdIdioma"
         Column0.Format  =   ""
         Column0.Bound   =   -1  'True
         Column0.Locked  =   0   'False
         Column0.HeaderAlignment=   0
         Column1.Heading =   "Idioma"
         Column1.Width   =   50
         Column1.Alignment=   0
         Column1.Hidden  =   0   'False
         Column1.Name    =   "dCorto"
         Column1.Format  =   ""
         Column1.Bound   =   -1  'True
         Column1.Locked  =   0   'False
         Column1.HeaderAlignment=   0
         SortKey1.Column =   -1
         SortKey1.Ascending=   -1  'True
         SortKey1.CaseInsensitive=   -1  'True
         SortKey2.Column =   -1
         SortKey2.Ascending=   -1  'True
         SortKey2.CaseInsensitive=   -1  'True
         SortKey3.Column =   -1
         SortKey3.Ascending=   -1  'True
         SortKey3.CaseInsensitive=   -1  'True
         BoundColumn     =   ""
         Border          =   -1  'True
         VertAlign       =   1
         Format          =   ""
      End
      Begin PVCOMBOLibCtl.PVComboBox cmbIdEstadoCivil 
         Height          =   330
         Left            =   7680
         TabIndex        =   14
         Top             =   210
         Width           =   2085
         _Version        =   524288
         _cx             =   3678
         _cy             =   582
         Appearance      =   1
         Enabled         =   -1  'True
         BackColor       =   16777215
         ForeColor       =   0
         Locked          =   0   'False
         Style           =   0
         Sorted          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowPictures    =   0   'False
         ColumnHeaders   =   -1  'True
         PrimaryColumn   =   1
         VisibleItems    =   10
         ColumnHeaderHeight=   20
         ListMember      =   ""
         ColumnHeaderForeColor=   0
         ColumnHeaderBackColor=   13160660
         SelectedForeColor=   16777215
         SelectedBackColor=   6956042
         AlternateBackColor=   16777215
         ItemLabelStyle  =   1
         ItemLabelType   =   0
         ItemLabelWidth  =   40
         ItemLabelForeColor=   0
         ItemLabelBackColor=   13160660
         ColumnHeaderStyle=   1
         VerticalGridLines=   -1  'True
         HorizontalGridLines=   -1  'True
         ColumnResize    =   0   'False
         ItemLabelResize =   0   'False
         AllowDBAutoConfig=   0   'False
         GridLineColor   =   13421772
         List            =   ""
         NullString      =   "[NULL]"
         DropShadow      =   -1  'True
         Text            =   ""
         SortOnColumnHeaderClick=   0   'False
         DropEffect      =   1
         ColumnCount     =   2
         Column0.Heading =   "Id"
         Column0.Width   =   20
         Column0.Alignment=   0
         Column0.Hidden  =   -1  'True
         Column0.Name    =   "IdEstadoCivil"
         Column0.Format  =   ""
         Column0.Bound   =   -1  'True
         Column0.Locked  =   0   'False
         Column0.HeaderAlignment=   0
         Column1.Heading =   "Estado Civil"
         Column1.Width   =   50
         Column1.Alignment=   0
         Column1.Hidden  =   0   'False
         Column1.Name    =   "dCorto"
         Column1.Format  =   ""
         Column1.Bound   =   -1  'True
         Column1.Locked  =   0   'False
         Column1.HeaderAlignment=   0
         SortKey1.Column =   -1
         SortKey1.Ascending=   -1  'True
         SortKey1.CaseInsensitive=   -1  'True
         SortKey2.Column =   -1
         SortKey2.Ascending=   -1  'True
         SortKey2.CaseInsensitive=   -1  'True
         SortKey3.Column =   -1
         SortKey3.Ascending=   -1  'True
         SortKey3.CaseInsensitive=   -1  'True
         BoundColumn     =   ""
         Border          =   -1  'True
         VertAlign       =   1
         Format          =   ""
      End
      Begin PVCOMBOLibCtl.PVComboBox cmbIdTipoOcupacion 
         Height          =   330
         Left            =   7680
         TabIndex        =   18
         Top             =   1680
         Width           =   3975
         _Version        =   524288
         _cx             =   7011
         _cy             =   582
         Appearance      =   1
         Enabled         =   -1  'True
         BackColor       =   16777215
         ForeColor       =   0
         Locked          =   0   'False
         Style           =   0
         Sorted          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowPictures    =   0   'False
         ColumnHeaders   =   -1  'True
         PrimaryColumn   =   1
         VisibleItems    =   10
         ColumnHeaderHeight=   20
         ListMember      =   ""
         ColumnHeaderForeColor=   0
         ColumnHeaderBackColor=   13160660
         SelectedForeColor=   16777215
         SelectedBackColor=   6956042
         AlternateBackColor=   16777215
         ItemLabelStyle  =   1
         ItemLabelType   =   0
         ItemLabelWidth  =   40
         ItemLabelForeColor=   0
         ItemLabelBackColor=   13160660
         ColumnHeaderStyle=   1
         VerticalGridLines=   -1  'True
         HorizontalGridLines=   -1  'True
         ColumnResize    =   0   'False
         ItemLabelResize =   0   'False
         AllowDBAutoConfig=   0   'False
         GridLineColor   =   13421772
         List            =   ""
         NullString      =   "[NULL]"
         DropShadow      =   -1  'True
         Text            =   ""
         SortOnColumnHeaderClick=   0   'False
         DropEffect      =   1
         ColumnCount     =   2
         Column0.Heading =   "Id"
         Column0.Width   =   20
         Column0.Alignment=   0
         Column0.Hidden  =   -1  'True
         Column0.Name    =   "IdTipoOcupacion"
         Column0.Format  =   ""
         Column0.Bound   =   -1  'True
         Column0.Locked  =   0   'False
         Column0.HeaderAlignment=   0
         Column1.Heading =   "Ocupación"
         Column1.Width   =   50
         Column1.Alignment=   0
         Column1.Hidden  =   0   'False
         Column1.Name    =   "dCorto"
         Column1.Format  =   ""
         Column1.Bound   =   -1  'True
         Column1.Locked  =   0   'False
         Column1.HeaderAlignment=   0
         SortKey1.Column =   -1
         SortKey1.Ascending=   -1  'True
         SortKey1.CaseInsensitive=   -1  'True
         SortKey2.Column =   -1
         SortKey2.Ascending=   -1  'True
         SortKey2.CaseInsensitive=   -1  'True
         SortKey3.Column =   -1
         SortKey3.Ascending=   -1  'True
         SortKey3.CaseInsensitive=   -1  'True
         BoundColumn     =   ""
         Border          =   -1  'True
         VertAlign       =   1
         Format          =   ""
      End
      Begin PVCOMBOLibCtl.PVComboBox cmbIdGradoInstruccion 
         Height          =   330
         Left            =   7680
         TabIndex        =   15
         Top             =   600
         Width           =   3975
         _Version        =   524288
         _cx             =   7011
         _cy             =   582
         Appearance      =   1
         Enabled         =   -1  'True
         BackColor       =   16777215
         ForeColor       =   0
         Locked          =   0   'False
         Style           =   0
         Sorted          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowPictures    =   0   'False
         ColumnHeaders   =   -1  'True
         PrimaryColumn   =   1
         VisibleItems    =   10
         ColumnHeaderHeight=   20
         ListMember      =   ""
         ColumnHeaderForeColor=   0
         ColumnHeaderBackColor=   13160660
         SelectedForeColor=   16777215
         SelectedBackColor=   6956042
         AlternateBackColor=   16777215
         ItemLabelStyle  =   1
         ItemLabelType   =   0
         ItemLabelWidth  =   40
         ItemLabelForeColor=   0
         ItemLabelBackColor=   13160660
         ColumnHeaderStyle=   1
         VerticalGridLines=   -1  'True
         HorizontalGridLines=   -1  'True
         ColumnResize    =   0   'False
         ItemLabelResize =   0   'False
         AllowDBAutoConfig=   0   'False
         GridLineColor   =   13421772
         List            =   ""
         NullString      =   "[NULL]"
         DropShadow      =   -1  'True
         Text            =   ""
         SortOnColumnHeaderClick=   0   'False
         DropEffect      =   1
         ColumnCount     =   2
         Column0.Heading =   "Id"
         Column0.Width   =   20
         Column0.Alignment=   0
         Column0.Hidden  =   -1  'True
         Column0.Name    =   "IdGradoInstruccion"
         Column0.Format  =   ""
         Column0.Bound   =   -1  'True
         Column0.Locked  =   0   'False
         Column0.HeaderAlignment=   0
         Column1.Heading =   "Grado Instrucción"
         Column1.Width   =   100
         Column1.Alignment=   0
         Column1.Hidden  =   0   'False
         Column1.Name    =   "dCorto"
         Column1.Format  =   ""
         Column1.Bound   =   -1  'True
         Column1.Locked  =   0   'False
         Column1.HeaderAlignment=   0
         SortKey1.Column =   -1
         SortKey1.Ascending=   -1  'True
         SortKey1.CaseInsensitive=   -1  'True
         SortKey2.Column =   -1
         SortKey2.Ascending=   -1  'True
         SortKey2.CaseInsensitive=   -1  'True
         SortKey3.Column =   -1
         SortKey3.Ascending=   -1  'True
         SortKey3.CaseInsensitive=   -1  'True
         BoundColumn     =   ""
         Border          =   -1  'True
         VertAlign       =   1
         Format          =   ""
      End
      Begin PVCOMBOLibCtl.PVComboBox cmbIdProcedencia 
         Height          =   330
         Left            =   7680
         TabIndex        =   17
         Top             =   1320
         Width           =   3975
         _Version        =   524288
         _cx             =   7011
         _cy             =   582
         Appearance      =   1
         Enabled         =   -1  'True
         BackColor       =   16777215
         ForeColor       =   0
         Locked          =   0   'False
         Style           =   0
         Sorted          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowPictures    =   0   'False
         ColumnHeaders   =   -1  'True
         PrimaryColumn   =   1
         VisibleItems    =   10
         ColumnHeaderHeight=   20
         ListMember      =   ""
         ColumnHeaderForeColor=   0
         ColumnHeaderBackColor=   13160660
         SelectedForeColor=   16777215
         SelectedBackColor=   6956042
         AlternateBackColor=   16777215
         ItemLabelStyle  =   1
         ItemLabelType   =   0
         ItemLabelWidth  =   40
         ItemLabelForeColor=   0
         ItemLabelBackColor=   13160660
         ColumnHeaderStyle=   1
         VerticalGridLines=   -1  'True
         HorizontalGridLines=   -1  'True
         ColumnResize    =   0   'False
         ItemLabelResize =   0   'False
         AllowDBAutoConfig=   0   'False
         GridLineColor   =   13421772
         List            =   ""
         NullString      =   "[NULL]"
         DropShadow      =   -1  'True
         Text            =   ""
         SortOnColumnHeaderClick=   0   'False
         DropEffect      =   1
         ColumnCount     =   2
         Column0.Heading =   "Id"
         Column0.Width   =   20
         Column0.Alignment=   0
         Column0.Hidden  =   -1  'True
         Column0.Name    =   "IdProcedencia"
         Column0.Format  =   ""
         Column0.Bound   =   -1  'True
         Column0.Locked  =   0   'False
         Column0.HeaderAlignment=   0
         Column1.Heading =   "Procedencia"
         Column1.Width   =   100
         Column1.Alignment=   0
         Column1.Hidden  =   0   'False
         Column1.Name    =   "dCorto"
         Column1.Format  =   ""
         Column1.Bound   =   -1  'True
         Column1.Locked  =   0   'False
         Column1.HeaderAlignment=   0
         SortKey1.Column =   -1
         SortKey1.Ascending=   -1  'True
         SortKey1.CaseInsensitive=   -1  'True
         SortKey2.Column =   -1
         SortKey2.Ascending=   -1  'True
         SortKey2.CaseInsensitive=   -1  'True
         SortKey3.Column =   -1
         SortKey3.Ascending=   -1  'True
         SortKey3.CaseInsensitive=   -1  'True
         BoundColumn     =   ""
         Border          =   -1  'True
         VertAlign       =   1
         Format          =   ""
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Height          =   660
         Left            =   3285
         Top             =   990
         Width           =   3030
      End
      Begin VB.Label lblTipoEdad 
         AutoSize        =   -1  'True
         Caption         =   "........"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   8100
         TabIndex        =   108
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "El PACIENTE es el Hijo N°"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3780
         TabIndex        =   100
         Top             =   1680
         Width           =   2100
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   7185
         TabIndex        =   98
         Top             =   2100
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Idioma materno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3330
         TabIndex        =   96
         Top             =   2070
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Etnia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   95
         Top             =   2070
         Width           =   405
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Id Paciente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   9840
         TabIndex        =   90
         Top             =   270
         Width           =   930
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido &Paterno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   84
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "&Primer Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   83
         Top             =   671
         Width           =   1215
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Apell. &Materno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3270
         TabIndex        =   82
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Gr.Instruc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   6780
         TabIndex        =   81
         Top             =   660
         Width           =   810
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   80
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "&Tercer Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   79
         Top             =   1057
         Width           =   1245
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Edad Actual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   6615
         TabIndex        =   78
         Top             =   1035
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   77
         Top             =   2370
         Width           =   1170
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "&Ocupación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   6735
         TabIndex        =   51
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "&Procedencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   6600
         TabIndex        =   53
         Top             =   1380
         Width           =   990
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Pad&re"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   6420
         TabIndex        =   52
         Top             =   2400
         Width           =   1170
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "&Segundo Nom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3300
         TabIndex        =   48
         Top             =   660
         Width           =   1170
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Te&léfono"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   9570
         TabIndex        =   54
         Top             =   990
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado &Civil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   6690
         TabIndex        =   49
         Top             =   255
         Width           =   900
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "F.Nacimien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3330
         TabIndex        =   50
         Top             =   1080
         Width           =   870
      End
   End
   Begin TabDlg.SSTab tabPaciente 
      Height          =   1470
      Left            =   0
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5025
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   2593
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
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
      TabCaption(0)   =   "1.1 Datos de domicilio (F7)"
      TabPicture(0)   =   "ucPacientesDetalle.ctx":485E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FraDomicilio"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "1.2. Datos de procedencia (F8)"
      TabPicture(1)   =   "ucPacientesDetalle.ctx":487A
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraProcedencia"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "1.3 Datos de nacimiento (F9)"
      TabPicture(2)   =   "ucPacientesDetalle.ctx":4896
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraNacimiento"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "1.4 Datos PDF/JGP"
      TabPicture(3)   =   "ucPacientesDetalle.ctx":48B2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ucPacientesPDF1"
      Tab(3).ControlCount=   1
      Begin SISGalenPlus.ucPacientesPDF ucPacientesPDF1 
         Height          =   1005
         Left            =   -74970
         TabIndex        =   135
         Top             =   345
         Width           =   11610
         _ExtentX        =   20479
         _ExtentY        =   1773
      End
      Begin VB.Frame fraProcedencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   945
         Left            =   120
         TabIndex        =   74
         Top             =   360
         Width           =   11415
         Begin VB.ComboBox cmbIdPaisProcedencia 
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
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   570
            Width           =   1425
         End
         Begin VB.ComboBox cmbIdCentroPobladoProcedencia 
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
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   570
            Width           =   4185
         End
         Begin VB.ComboBox cmbIdDistritoProcedencia 
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
            Left            =   8475
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   240
            Width           =   2745
         End
         Begin VB.ComboBox cmbIdProvinciaProcedencia 
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
            Left            =   4785
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   240
            Width           =   2655
         End
         Begin VB.ComboBox cmbIdDepartamentoProcedencia 
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
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   210
            Width           =   2250
         End
         Begin VB.CheckBox chkIgualQueDomicilio 
            Caption         =   "Igual que el domicilio"
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
            Left            =   8490
            TabIndex        =   38
            Top             =   660
            Width           =   2685
         End
         Begin VB.Label lblNotaDeUbicación 
            Caption         =   "..."
            Height          =   195
            Left            =   7740
            TabIndex        =   113
            Top             =   690
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   120
            TabIndex        =   61
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3990
            TabIndex        =   62
            Top             =   270
            Width           =   705
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Distrito"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   7815
            TabIndex        =   63
            Top             =   300
            Width           =   570
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Centro Poblado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   120
            TabIndex        =   64
            Top             =   630
            Width           =   1260
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "País"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5610
            TabIndex        =   65
            Top             =   630
            Width           =   300
         End
      End
      Begin VB.Frame fraNacimiento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   -74880
         TabIndex        =   73
         Top             =   330
         Width           =   11385
         Begin VB.ComboBox cmbIdDepartamentoNacimiento 
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
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   255
            Width           =   2250
         End
         Begin VB.ComboBox cmbIdProvinciaNacimiento 
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
            Left            =   4785
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   255
            Width           =   2655
         End
         Begin VB.ComboBox cmbIdDistritoNacimiento 
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
            Left            =   8475
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   255
            Width           =   2745
         End
         Begin VB.ComboBox cmbIdCentroPobladoNacimiento 
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
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   615
            Width           =   4395
         End
         Begin VB.ComboBox cmbIdPaisNacimiento 
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
            Left            =   6090
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   615
            Width           =   1335
         End
         Begin VB.CheckBox chkIgualUQueDomicilioNac 
            Caption         =   "Igual que el domicilio"
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
            Left            =   8490
            TabIndex        =   46
            Top             =   660
            Width           =   2655
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   120
            TabIndex        =   66
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3990
            TabIndex        =   67
            Top             =   285
            Width           =   705
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Distrito"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   7860
            TabIndex        =   68
            Top             =   300
            Width           =   570
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Centro Poblado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   120
            TabIndex        =   69
            Top             =   630
            Width           =   1260
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "País"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5790
            TabIndex        =   70
            Top             =   660
            Width           =   300
         End
      End
      Begin VB.Frame FraDomicilio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   945
         Left            =   -74940
         TabIndex        =   72
         Top             =   330
         Width           =   11595
         Begin VB.CommandButton cmdBuscaDistrito 
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
            Left            =   11205
            Picture         =   "ucPacientesDetalle.ctx":48CE
            Style           =   1  'Graphical
            TabIndex        =   123
            ToolTipText     =   "Buscar"
            Top             =   210
            Width           =   345
         End
         Begin VB.ComboBox cmbIdPaisDomicilio 
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
            Left            =   4995
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   555
            Width           =   1905
         End
         Begin VB.ComboBox cmbIdCentroPobladoDomicilio 
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
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   570
            Width           =   3195
         End
         Begin VB.ComboBox cmbIdDistritoDomicilio 
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
            Left            =   7725
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   225
            Width           =   3495
         End
         Begin VB.ComboBox cmbIdProvinciaDomicilio 
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
            Left            =   4155
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   225
            Width           =   2745
         End
         Begin VB.ComboBox cmbIdDepartamentoDomicilio 
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
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   210
            Width           =   1770
         End
         Begin VB.TextBox txtDireccionDomicilio 
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
            Left            =   7740
            MaxLength       =   100
            TabIndex        =   32
            Top             =   570
            Width           =   3765
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "País"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   4680
            TabIndex        =   59
            Top             =   600
            Width           =   300
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Dirección"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   6990
            TabIndex        =   60
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Depar&tamento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   135
            TabIndex        =   55
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Pro&vincia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3450
            TabIndex        =   56
            Top             =   270
            Width           =   705
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Distrito"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   7140
            TabIndex        =   57
            Top             =   270
            Width           =   570
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Centro Poblado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   135
            TabIndex        =   58
            Top             =   660
            Width           =   1260
         End
      End
   End
   Begin VB.Label lblFichaFamiliar 
      Alignment       =   1  'Right Justify
      Caption         =   "Ficha Familiar:"
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
      Left            =   6465
      TabIndex        =   130
      Top             =   4665
      Visible         =   0   'False
      Width           =   1155
   End
End
Attribute VB_Name = "ucPacientesDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para registrar los datos personales del Paciente
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'------------------------------------------------------------------------------------
Option Explicit

Private Const cb_setdroppedwidth = &H160
Dim ldHoy As Date
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Cadena As New sighEntidades.cadena
Dim mo_Formulario As New sighEntidades.Formulario
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim mb_ExistenDatos As Boolean
Dim ms_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim ml_TipoServicio As sghTipoServicio
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdminHoteleria As New SIGHNegocios.ReglasHoteleria

'<(Inicio)Comentado Por: WABG el: 16/10/2020-10:29:40 a.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
''Dim mo_Reniec As New ReniecGalenhos
'</(Fin)Comentado por: WABG el: 16/10/2020-10:29:40 a.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>

'<(Inicio) Añadido Por: WABG el: 16/10/2020-10:53:29 a.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
Dim mo_Reniec As New ReniecGalenhosNegocios
Dim lcIdDistrito As String
'</(Fin) Añadido Por: WABG el: 16/10/2020-10:53:29 a.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>

Dim mrs_Diagnosticos As New ADODB.Recordset
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim ml_TipoVistaForm As sghTipoVistaFormAtenciones
Dim mb_PacienteNoIdentificado As Boolean
Public Event SeModificoFechaNacimiento(sFechaNacimiento As String, sHoraNacimiento As String)
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Public Event SeModificoPacienteNoIdentificado(bPacienteNoIdentificado As Boolean)
Public Event SeModificoSexo(lIdTipoSexo As Long)
Dim oRsPaisDomicilio As New Recordset
Dim oRsPaisProcedencia As New Recordset
Dim oRsPaisNacimiento As New Recordset
Dim oRsDptoDomicilio As New Recordset
Dim oRsDptoProcedencia As New Recordset
Dim oRsDptoNacimiento As New Recordset
Dim mo_cmbIdTipoGenHistoriaClinica As New sighEntidades.ListaDespleglable
Dim mo_CmbIdTipoSexo As New sighEntidades.ListaDespleglable
'Dim mo_cmbIdEstadoCivil As New sighentidades.ListaDespleglable
Dim mo_cmbIdDocIdentidad As New sighEntidades.ListaDespleglable
'Dim mo_cmbIdGradoInstruccion As New sighentidades.ListaDespleglable
'Dim mo_cmbIdTipoOcupacion As New sighentidades.ListaDespleglable
'Dim mo_cmbIdProcedencia As New sighentidades.ListaDespleglable
Dim mo_cmbIdDepartamentoDomicilio As New sighEntidades.ListaDespleglable
Dim mo_cmbIdProvinciaDomicilio As New sighEntidades.ListaDespleglable
Dim mo_cmbIdDistritoDomicilio As New sighEntidades.ListaDespleglable
Dim mo_cmbIdPaisDomicilio As New sighEntidades.ListaDespleglable
Dim mo_cmbIdCentroPobladoDomicilio As New sighEntidades.ListaDespleglable
Dim mo_cmbIdDepartamentoProcedencia As New sighEntidades.ListaDespleglable
Dim mo_cmbIdProvinciaProcedencia As New sighEntidades.ListaDespleglable
Dim mo_cmbIdDistritoProcedencia As New sighEntidades.ListaDespleglable
Dim mo_cmbIdCentroPobladoProcedencia As New sighEntidades.ListaDespleglable
Dim mo_cmbIdPaisProcedencia As New sighEntidades.ListaDespleglable
Dim mo_cmbIdDepartamentoNacimiento As New sighEntidades.ListaDespleglable
Dim mo_cmbIdProvinciaNacimiento As New sighEntidades.ListaDespleglable
Dim mo_cmbIdDistritoNacimiento As New sighEntidades.ListaDespleglable
Dim mo_cmbIdCentroPobladoNacimiento As New sighEntidades.ListaDespleglable
Dim mo_cmbIdPaisNacimiento As New sighEntidades.ListaDespleglable
'Combo Etnia-GLCC-10/07/2020
Dim mo_cmbEtnia As New sighEntidades.ListaDespleglable
'Dim mo_cmbIdioma As New sighentidades.ListaDespleglable
Dim mo_cmbMadreTipoDocumento As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoEdad As New sighEntidades.ListaDespleglable
'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA FILIACION
'------------------------------------------------------------------------------------
Dim ml_IdPaciente As Long
Dim ms_Autogenerado As String
Dim mo_Historia As New DOHistoriaClinica
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lc_ArchivoElegido As String
Dim ml_meHwnd As Long
Dim lcFormaQgeneraHistoria As String
Dim lcEtniaDefault As String
Dim lbExigeIngresoDelDNI As Boolean
Dim lbExigeIngresoDeCentroPoblado As Boolean
Dim lbBuscaDNIenReniec As Boolean
Dim mb_UsoWebReniec As Boolean

'<(Inicio) Añadido Por: WABG el: 23/10/2020-07:42:08 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
Dim mb_validacionReniec As Boolean
'</(Fin) Añadido Por: WABG el: 23/10/2020-07:42:08 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>

Dim mb_MarcoCheckPacienteNuevo As Boolean
Dim oCampos() As String
Dim lbUsoWebReniec_SinMostrar As Boolean
Dim lcFactorRh_SinMostrar As String
Dim lcGrupoSanguineo_SinMostrar As String
Dim lblSegundoNombrePacienteSIS As String 'Frank 28 01 2015

Dim lnOpcionQueUsaEsteControl As Long     '1->Pacientes, 2->Admision de Emergencia, 3->Admision de Hospitalizacion
Dim lcIdTipoGenHistoriaClinicaActual As String

Property Let OpcionQueUsaEsteControl(iValue As Long)
    lnOpcionQueUsaEsteControl = iValue
    If iValue = 1 Then
        cmdAcreditaSIS.Visible = True
    End If
End Property

Property Let MarcoCheckPacienteNuevo(iValue As Boolean)
   mb_MarcoCheckPacienteNuevo = iValue
End Property


Property Let UsoWebReniec(iValue As Boolean)
   mb_UsoWebReniec = iValue
End Property
Property Get UsoWebReniec() As Boolean
    UsoWebReniec = mb_UsoWebReniec
End Property


'<(Inicio) Añadido Por: WABG el: 23/10/2020-07:45:00 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
Property Let validacionReniec(vValue As Boolean)
   mb_validacionReniec = vValue
End Property
Property Get validacionReniec() As Boolean
    validacionReniec = mb_validacionReniec
End Property
'</(Fin) Añadido Por: WABG el: 23/10/2020-07:45:00 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>

Property Let meHwnd(lValue As Long)
   ml_meHwnd = lValue
End Property
Property Get ArchivoElegido() As String
    ArchivoElegido = lc_ArchivoElegido
End Property

Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Let idPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_IdPaciente
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let Autogenerado(sValue As String)
   ms_Autogenerado = sValue
End Property
Property Get Autogenerado() As String
   Autogenerado = ms_Autogenerado
End Property
Property Let FechaNacimiento(sValue As String)
   txtFechaNacimiento.Text = sValue
End Property
Property Get FechaNacimiento() As String
   FechaNacimiento = txtFechaNacimiento.Text
End Property
Property Get HoraNacimiento() As String
   HoraNacimiento = txtHoraNacimiento.Text
End Property
Property Let PacienteNoIdentificado(bValue As Boolean)
   UserControl.chkNN.Value = IIf(bValue, 1, 0)
End Property
Property Get PacienteNoIdentificado() As Boolean
   PacienteNoIdentificado = IIf(chkNN.Value, True, False)
End Property
Property Let NroHistoriaClinica(lValue As Long)
   txtIdNroHistoria.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lValue)), False)
   txtIdNroHistoria.Tag = lValue
End Property
Property Get NroHistoriaClinica() As Long
   NroHistoriaClinica = Val(txtIdNroHistoria.Text)
End Property
Property Let TipoServicio(sValue As sghTipoServicio)
   ml_TipoServicio = sValue
End Property
Property Get TipoServicio() As sghTipoServicio
   TipoServicio = ml_TipoServicio
End Property
Property Get FechaCreacionHistoria() As String
   FechaCreacionHistoria = UserControl.txtFechaCreacion.Text
End Property
Property Get NotaSobreUbicacion() As String
    NotaSobreUbicacion = UserControl.lblNotaDeUbicación
End Property
Property Let NotaSobreUbicacion(sValue As String)
     UserControl.lblNotaDeUbicación = sValue
End Property

'Frank 28 01 2015
Property Get SegundoNombrePacienteSIS() As String
    SegundoNombrePacienteSIS = lblSegundoNombrePacienteSIS
End Property
Property Let SegundoNombrePacienteSIS(sValue As String)
     lblSegundoNombrePacienteSIS = sValue
End Property



Property Let TipoNumeracion(lValue As sghTipoNumeracionDeNroHistoria)
    mo_cmbIdTipoGenHistoriaClinica.BoundText = CStr(lValue)
    cmbIdTipoGenHistoriaClinica.Tag = CStr(lValue)
    
    Select Case lValue
    Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
        mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, False
        mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    Case Else
        mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, True
        mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, True
        'mo_Formulario.HabilitarDeshabilitar txtFechaCreacion, True
    End Select
    HabilitaFechaCreacion
End Property
Property Get TipoNumeracionActual() As sghTipoNumeracionDeNroHistoria
   TipoNumeracionActual = Val(mo_cmbIdTipoGenHistoriaClinica.BoundText)
End Property
Property Get TipoNumeracionAnterior() As sghTipoNumeracionDeNroHistoria
   TipoNumeracionAnterior = Val(cmbIdTipoGenHistoriaClinica.Tag)
End Property
Property Get IdHistoriaClinicaAnterior() As Long
   IdHistoriaClinicaAnterior = Val(txtIdNroHistoria.Tag)
End Property

Property Get ExistePaciente() As Boolean
   ExistePaciente = mb_ExistenDatos
End Property

Property Get idTipoSexo() As Long
   idTipoSexo = Val(mo_CmbIdTipoSexo.BoundText)
End Property

Private Sub btnAgreaIMG_Click()

    On Error Resume Next
    Dim oMuestraCAmara As New SIGHNegocios.BuscaArchivo
    oMuestraCAmara.MuestraCamara
    Set oMuestraCAmara = Nothing
    lc_ArchivoElegido = sighEntidades.RutaImagenConPermiso
    If lc_ArchivoElegido <> "" Then
       pi_ImagSeleccionada.Picture = LoadPicture(lc_ArchivoElegido)
    End If
End Sub


Private Sub btnQuitaIMG_Click()
    pi_ImagSeleccionada.Picture = LoadPicture("")
    lc_ArchivoElegido = "DEL"
End Sub

Private Sub cboTipoEdadPaciente_GotFocus()
    cboTipoEdadPaciente.Tag = mo_cmbIdTipoEdad.BoundText
End Sub

Private Sub cboTipoEdadPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cboTipoEdadPaciente
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub cboTipoEdadPaciente_LostFocus()
    If cboTipoEdadPaciente.Enabled = False Or cboTipoEdadPaciente.Locked = True Then
        Exit Sub
    End If
    If cboTipoEdadPaciente.Tag <> mo_cmbIdTipoEdad.BoundText Then
        Call calcularFechaDeNacimiento(txtEdad.Text, mo_cmbIdTipoEdad.BoundText)
    End If
End Sub

Private Sub chkIgualQueDomicilio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkIgualQueDomicilio
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub chkIgualUQueDomicilioNac_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkIgualUQueDomicilioNac
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub
'<(Inicio) Añadido Por: WABG el: 23/10/2020-07:58:47 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
'DESHABILITA CONTROLES PARA MODIFICAR UN PACIENTE VALIDADO POR LA RENIEC
Sub deshabilitarControlesRENIECModificarPacienteValidado()

            cmdSinApellidoPaterno.Enabled = False
            cmdSinApellidoMaterno.Enabled = False
            cmbIdDocIdentidad.Enabled = False
            cmdCambiaHC.Enabled = False
            chkSinFechaNacimiento.Enabled = False
            cmbIdTipoSexo.Enabled = False
            
            txtNroDocumento.Enabled = False
            txtApellidoPaterno.Enabled = False
            txtApellidoMaterno.Enabled = False
            txtPrimerNombre.Enabled = False
            txtSegundoNombre.Enabled = False
            txtTercerNombre.Enabled = False
            txtFechaNacimiento.Enabled = False
            txtEdad.Enabled = False
            
            UserControl.TabPaciente.Tab = 0
End Sub
 Sub deshabilitarControlesDeTextoRENIEC()
 
                    cmdSinApellidoPaterno.Enabled = False
                    cmdSinApellidoMaterno.Enabled = False
                    cmbIdDocIdentidad.Enabled = False
                    cmdCambiaHC.Enabled = False
                    chkSinFechaNacimiento.Enabled = False
                    
                     
'<(Inicio) Añadido Por: WABG el: 27/10/2020-10:03:25 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
                     txtNroDocumento.Enabled = False
                     txtIdNroHistoria.Enabled = False
                     cmbIdTipoGenHistoriaClinica.Enabled = False
'</(Fin) Añadido Por: WABG el: 27/10/2020-10:03:25 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
                                          
                     
                     txtApellidoPaterno.Enabled = False
                     txtApellidoMaterno.Enabled = False
                     txtPrimerNombre.Enabled = False
                     txtSegundoNombre.Enabled = False
                     txtTercerNombre.Enabled = False
                     txtFechaNacimiento.Enabled = False
                     cmbIdTipoSexo.Enabled = False
                     
                      If txtDireccionDomicilio.Text = "" Then
                            txtDireccionDomicilio.Enabled = True
                      Else
                            txtDireccionDomicilio.Enabled = False
                     End If

                     If Len(Trim(cmbIdDepartamentoDomicilio.Text)) > 0 Then
                            cmbIdDepartamentoDomicilio.Enabled = False
                     Else
                            cmbIdDepartamentoDomicilio.Enabled = True
                     End If

                     If Len(Trim(cmbIdProvinciaDomicilio.Text)) > 0 Then
                            cmbIdProvinciaDomicilio.Enabled = False
                     Else
                            cmbIdProvinciaDomicilio.Enabled = True
                     End If

                     If Len(Trim(cmbIdDistritoDomicilio.Text)) > 0 Then
                            cmbIdDistritoDomicilio.Enabled = False
                     Else
                            cmbIdDistritoDomicilio.Enabled = True
                     End If

                     If Len(Trim(cmbIdCentroPobladoDomicilio.Text)) > 0 Then
                            cmbIdCentroPobladoDomicilio.Enabled = False
                     Else
                            cmbIdCentroPobladoDomicilio.Enabled = True
                     End If

                      If Len(Trim(cmbIdPaisDomicilio.Text)) > 0 Then
                            cmbIdPaisDomicilio.Enabled = False
                    Else
                            cmbIdPaisDomicilio.Enabled = True
                     End If
  
End Sub

'CARGAR CONTROLES DE TEXTO DESDE RENIEC
Sub CargarDatosDesdeRENIEC(DNI As String)
                  
                  'LIMPIAR TODOS LOS CONTROLES, METODO ENCONTRADO EN EL CODIGO
                  Call LimpiarDatosDePaciente(0, Format(ldHoy, sighEntidades.DevuelveFechaSoloFormato_DMY))
                
                
'<Agregado por: WABG el: 11/24/2020-09:55:01 en el equipo: SISGALENPLUS-PC><CAMBIO 37>
                
                 mo_Reniec.SeAccesaAlaWebDesdeGalenhos = True
'</Agregado por: WABG el: 11/24/2020-09:55:01 en el equipo: SISGALENPLUS-PC><CAMBIO 37>
                  mo_Reniec.Inicializar
'                  mo_Reniec.ConsultarDNIenReniec Trim(DNI) genaro leonel campos carmen
                  mo_Reniec.ConsultarDNIenReniec DNI

                  If mo_Reniec.ApellidoPaterno <> "" Then
                  If Trim(txtNroDocumento.Text) <> DNI Then
                  txtNroDocumento.Text = DNI
                  'txtNroDocumento.Enabled = False
                  End If
                  txtIdNroHistoria.Text = DNI
                  cmbIdTipoGenHistoriaClinica.ListIndex = 1
                  txtIdNroHistoria.Locked = False
                  txtIdNroHistoria.Enabled = True
                  txtIdNroHistoria.SetFocus
                  
                  txtApellidoPaterno.Text = mo_Reniec.ApellidoPaterno
                  txtApellidoMaterno.Text = mo_Reniec.ApellidoMaterno
                  txtPrimerNombre.Text = mo_Reniec.PrimerNombre
                  txtSegundoNombre.Text = mo_Reniec.SegundoNombre
                  txtTercerNombre.Text = mo_Reniec.TercerNombre
                  txtFechaNacimiento.Text = mo_Reniec.FechaNacimiento
                  mo_CmbIdTipoSexo.BoundText = mo_Reniec.idTipoSexo
                  txtDireccionDomicilio.Text = mo_Reniec.DireccionDomicilio
                  
                
'<(Inicio) Añadido Por: WABG el: 27/10/2020-10:03:00 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
                  txtHoraNacimiento.SetFocus
'</(Fin) Añadido Por: WABG el: 27/10/2020-10:03:00 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
                  
                  'llenando los combobox  etnia e idioma por defecto
                  
                  'Idioma español = 38 en el combobox
                  cmbIdioma.ListIndex = 38
                  
                  'Etnia mestizo = 43 en el combobox
                  cmbEtnia.ListIndex = 43
                                   
                  If mo_Reniec.IdDistritoDomicilio > 0 Then
                     lcIdDistrito = Right("0" & Trim(Str(mo_Reniec.IdDistritoDomicilio)), 6)
                     mo_cmbIdDepartamentoDomicilio.BoundText = Left(lcIdDistrito, 2)
                     mo_cmbIdProvinciaDomicilio.BoundText = Left(lcIdDistrito, 4)
                     mo_cmbIdDistritoDomicilio.BoundText = lcIdDistrito
'                    mo_cmbIdPaisDomicilio.BoundText = 166   'Peru
'                    mo_cmbIdPaisNacimiento.BoundText = 166   'Peru
'                    mo_cmbIdPaisProcedencia.BoundText = 166   'Peru'
                  End If
                     UserControl.TabPaciente.Tab = 0
                     MsgBox "Los Datos del Paciente con  DNI : " & DNI & "  Fueron cargados desde la RENIEC", vbInformation, ""
                    
                    'DESHABILITAR CONTROLES IMPLICADOS EN DATOS DE RENIEC
                     deshabilitarControlesDeTextoRENIEC
                     
                     'SE GRABA EN EL CAMPO validacionReniec de sigh.dbo.pacientes
                     mb_validacionReniec = True
                    
                    
                     
            Else
            
                     MsgBox "El Numero de DNI : " & DNI & " no fue encontrado en la RENIEC o no tiene conexion a INTERNET" & Chr(13) & _
                     "Ingrese los Datos Manualmente", vbInformation, "Error"
                    
                     
                     'LIMPIAR TODOS LOS CONTROLES, METODO ENCONTRADO EN EL CODIGO
                      Call LimpiarDatosDePaciente(0, Format(ldHoy, sighEntidades.DevuelveFechaSoloFormato_DMY))

                     'HABILITAR CONTROLES IMPLICADOS EN DATOS DE RENIEC
                     HabilitarControlesDeTextoRENIEC
                     
                     txtNroDocumento.Text = DNI
                     txtIdNroHistoria.Text = DNI
                     cmbIdTipoGenHistoriaClinica.ListIndex = 1
                     txtApellidoPaterno.SetFocus
                     
'
            End If
            
End Sub
'CARGAR DATOS DEL TUTOR DESDE RENIEC
Sub CargarDatosTutorDesdeRENIEC(DNITUTOR As String)
    mo_Reniec.Inicializar
    mo_Reniec.ConsultarDNIenReniec Trim(DNITUTOR)
    If Len(txtMadreDocumento.Text) = 8 And Val(mo_cmbMadreTipoDocumento.BoundText) = 1 Then

        If mo_Reniec.ApellidoPaterno <> "" Then
            txtMadreApellidoP.Text = mo_Reniec.ApellidoPaterno
            txtMadreApellidoM.Text = mo_Reniec.ApellidoMaterno
            txtNombreMadre.Text = mo_Reniec.PrimerNombre
            txtMadreSnombre.Text = mo_Reniec.SegundoNombre
            txtMadreApellidoP.Enabled = False
            txtMadreApellidoM.Enabled = False
            txtNombreMadre.Enabled = False
            txtMadreSnombre.Enabled = False
            
            MsgBox "Los Datos del Tutor con  DNI : " & DNITUTOR & "  Fueron cargados desde la RENIEC", vbInformation, ""
        Else
            MsgBox "El Numero de DNI : " & DNITUTOR & " no fue encontrado en la RENIEC", vbInformation, "Error"
            txtMadreApellidoP.Text = ""
            txtMadreApellidoM.Text = ""
            txtNombreMadre.Text = ""
            txtMadreSnombre.Text = ""
            txtMadreApellidoP.Enabled = True
            txtMadreApellidoM.Enabled = True
            txtNombreMadre.Enabled = True
            txtMadreSnombre.Enabled = True
            txtMadreDocumento.Text = DNITUTOR
            mo_cmbMadreTipoDocumento.BoundText = 1
            txtMadreApellidoP.SetFocus
     End If

    Else
           MsgBox "Ingreso Incorrecto", vbInformation, "Error"
  End If
End Sub
'BUSCAR EN BD EXISTENCIA DE HISTORIAS CLINICAS
 Sub VerificarExistenciaHistoriaClinica(NroHistoriaClinica As String)
   On Error Resume Next
    If mo_Teclado.TextoEsSoloNumeros(NroHistoriaClinica) Then
        Dim lbContinua000 As Boolean
        
        lbContinua000 = True
        If Len(NroHistoriaClinica) = 8 And wxParametro351 = "S" Then
           If NroHistoriaClinica = txtNroDocumento.Text Then
              lbContinua000 = False
           End If
        End If
        If lbContinua000 = True Then
            NroHistoriaClinica = mo_Teclado.CapitalizarNombres(NroHistoriaClinica)
            txtIdNroHistoria.Tag = NroHistoriaClinica
        End If
        mo_Formulario.MarcarComoVacio txtIdNroHistoria
        If txtIdNroHistoria.Locked = True Then Exit Sub
        If Trim(NroHistoriaClinica) = "" Then txtIdNroHistoria.SetFocus: Exit Sub
        ms_MensajeError = mo_AdminAdmision.ExisteNroHistoria(Trim(Str(txtIdNroHistoria.Tag)))
        If ms_MensajeError <> "" Then
           MsgBox "Existe un paciente con el mismo número de historia clínica: " + Chr(13) + ms_MensajeError
           txtIdNroHistoria.Text = ""
           txtIdNroHistoria.Enabled = True
           txtIdNroHistoria.SetFocus
        End If
    End If
End Sub

Sub HabilitarControlesDeTextoRENIEC()

                     cmdSinApellidoPaterno.Enabled = True
                     cmdSinApellidoMaterno.Enabled = True
                     cmbIdDocIdentidad.Enabled = True
'<(Inicio) Añadido Por: WABG el: 27/10/2020-08:12:48 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
                     cmbIdTipoGenHistoriaClinica.Enabled = True
'</(Fin) Añadido Por: WABG el: 27/10/2020-08:12:48 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
                     cmdCambiaHC.Enabled = True
                     chkSinFechaNacimiento.Enabled = True
                    
                     txtNroDocumento.Enabled = True
                     txtApellidoPaterno.Enabled = True
                     txtApellidoMaterno.Enabled = True
                     txtPrimerNombre.Enabled = True
                     txtSegundoNombre.Enabled = True
                     txtTercerNombre.Enabled = True
                     txtIdNroHistoria.Enabled = True
                     txtIdNroHistoria.Locked = False
                     txtFechaNacimiento.Enabled = True
                     txtDireccionDomicilio.Enabled = True
                     
                     cmbIdTipoSexo.Enabled = True
                     cmbIdDepartamentoDomicilio.Enabled = True
                     cmbIdProvinciaDomicilio.Enabled = True
                     cmbIdDistritoDomicilio.Enabled = True
                     cmbIdCentroPobladoDomicilio.Enabled = True
                     cmbIdPaisDomicilio.Enabled = True
                     UserControl.TabPaciente.Tab = 0
'<(Inicio)Comentado Por: WABG el: 27/10/2020-07:46:37 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
'                     txtNroDocumento.SetFocus
'</(Fin)Comentado por: WABG el: 27/10/2020-07:46:37 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
End Sub
'</(Fin) Añadido Por: WABG el: 23/10/2020-07:58:47 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>

Private Sub chkNN_Click()
    
    RaiseEvent SeModificoPacienteNoIdentificado(chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar txtApellidoPaterno, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar txtApellidoMaterno, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar txtPrimerNombre, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar txtSegundoNombre, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar txtTercerNombre, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar txtFichaFamiliar1, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar txtFichaFamiliar2, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar txtFichaFamiliar3, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar cmbEtnia, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdioma, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar cmbMadreTipoDocumento, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar txtMadreDocumento, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar txtMadreApellidoP, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar txtMadreApellidoM, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar txtNombreMadre, Not (chkNN.Value = 1)
    mo_Formulario.HabilitarDeshabilitar txtMadreSnombre, Not (chkNN.Value = 1)
    If mi_Opcion = sghAgregar Or mi_Opcion = sghModificar Then
        'mo_Formulario.HabilitarDeshabilitar txtFechaCreacion, Not (chkNN.Value = 1)
        
        ConfigurarPacienteNuevoONoIdentificado chkNN.Value
        'mo_cmbIdTipoGenHistoriaClinica.BoundText = lcBuscaParametro.SeleccionaFilaParametro(210)
        
    Else
        mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    End If
    If chkNN.Value = 1 Then
       mb_PacienteNoIdentificado = True
    Else
       mb_PacienteNoIdentificado = False
       UserControl.cmbEtnia.Text = ""
       UserControl.cmbIdioma.Text = ""
       On Error Resume Next
       UserControl.txtNroDocumento.SetFocus
    End If
    
End Sub

Private Sub chkNN_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkNN
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub
Private Sub chkSinFechaNacimiento_Click()
    Call bloquearControlEdad
End Sub

Private Sub chkSinFechaNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSinFechaNacimiento
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub cmbEtnia_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbEtnia
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub



Private Sub cmbIdCentroPobladoNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdCentroPobladoNacimiento
    RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub

Private Sub cmbIdCentroPobladoNacimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdDocIdentidad_Click()
  On Error Resume Next
  txtNroDocumento.Text = ""
  txtNroDocumento.SetFocus
End Sub



Private Sub cmbIdioma_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdioma
   RaiseEvent SePresionoTeclaEspecial(KeyCode)

End Sub

Private Sub cmbIdPaisDomicilio_Click()
  If cmbIdPaisDomicilio.Text = "Peru" Then
    cmbIdDepartamentoDomicilio.Enabled = True
    mo_cmbIdDepartamentoDomicilio.BoundText = Left(lcBuscaParametro.SeleccionaFilaParametro(242), 2)
  Else
    cmbIdDepartamentoDomicilio.Enabled = False
    cmbIdDepartamentoDomicilio.ListIndex = -1
    cmbIdProvinciaDomicilio.Enabled = False
    cmbIdProvinciaDomicilio.ListIndex = -1
    cmbIdDistritoDomicilio.Enabled = False
    cmbIdDistritoDomicilio.ListIndex = -1
    cmbIdCentroPobladoDomicilio.Enabled = False
    cmbIdCentroPobladoDomicilio.ListIndex = -1
  End If
End Sub

Private Sub cmbIdPaisNacimiento_Click()
  If cmbIdPaisNacimiento.Text = "Peru" Then
    cmbIdDepartamentoNacimiento.Enabled = True
    mo_cmbIdDepartamentoNacimiento.BoundText = Left(lcBuscaParametro.SeleccionaFilaParametro(242), 2)
  Else
    cmbIdDepartamentoNacimiento.Enabled = False
    cmbIdDepartamentoNacimiento.ListIndex = -1
    cmbIdProvinciaNacimiento.Enabled = False
    cmbIdProvinciaNacimiento.ListIndex = -1
    cmbIdDistritoNacimiento.Enabled = False
    cmbIdDistritoNacimiento.ListIndex = -1
    cmbIdCentroPobladoNacimiento.Enabled = False
    cmbIdCentroPobladoNacimiento.ListIndex = -1
  End If
End Sub

'-----------------------------------------------------------------------------------------
'*****************************************************************************************
'                               EVENTOS PARA PACIENTES
'*****************************************************************************************
'-----------------------------------------------------------------------------------------
Private Sub cmbIdPaisNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, mo_cmbIdPaisNacimiento
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub cmbIdPaisNacimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdPaisNacimiento_LostFocus()
   'If cmbIdPaisNacimiento.Text <> "" Then
   '    mo_cmbIdPaisNacimiento.BoundText = Val(Split(cmbIdPaisNacimiento.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdPaisNacimiento
End Sub

Private Sub cmbIdPaisProcedencia_Click()
  If cmbIdPaisProcedencia.Text = "Peru" Then
    cmbIdDepartamentoProcedencia.Enabled = True
    mo_cmbIdDepartamentoProcedencia.BoundText = Left(lcBuscaParametro.SeleccionaFilaParametro(242), 2)
  Else
    cmbIdDepartamentoProcedencia.Enabled = False
    cmbIdDepartamentoProcedencia.ListIndex = -1
    cmbIdProvinciaProcedencia.Enabled = False
    cmbIdProvinciaProcedencia.ListIndex = -1
    cmbIdDistritoProcedencia.Enabled = False
    cmbIdDistritoProcedencia.ListIndex = -1
    cmbIdCentroPobladoProcedencia.Enabled = False
    cmbIdCentroPobladoProcedencia.ListIndex = -1
  End If
End Sub

Private Sub cmbIdPaisProcedencia_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdPaisProcedencia
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub cmbIdPaisProcedencia_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdPaisProcedencia_LostFocus()
   'If cmbIdPaisProcedencia.Text <> "" Then
   '    mo_cmbIdPaisProcedencia.BoundText = Val(Split(cmbIdPaisProcedencia.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdPaisProcedencia
End Sub

Private Sub cmbIdTipoSexo_Change()
    RaiseEvent SeModificoSexo(Val(mo_CmbIdTipoSexo.BoundText))
End Sub







Private Sub cmbMadreTipoDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbMadreTipoDocumento
  RaiseEvent SePresionoTeclaEspecial(KeyCode)
  AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbMadreTipoDocumento_LostFocus()
        mo_Formulario.MarcarComoVacio cmbMadreTipoDocumento
        If mo_cmbMadreTipoDocumento.BoundText <> "" Then
           If mo_cmbMadreTipoDocumento.BoundText = "1" Then
              txtMadreDocumento.MaxLength = 8
           Else
              txtMadreDocumento.MaxLength = 12
           End If
        End If
End Sub


Private Sub cmdAcreditaSIS_Click()
  AfiliacionSIS.ApellidoPaterno = UserControl.txtApellidoPaterno.Text
  AfiliacionSIS.ApellidoMaterno = UserControl.txtApellidoMaterno.Text
  AfiliacionSIS.PrimerNombre = UserControl.txtPrimerNombre.Text
  AfiliacionSIS.SegundoNombre = UserControl.txtSegundoNombre.Text
  AfiliacionSIS.idTipoSexo = Val(mo_CmbIdTipoSexo.BoundText)
  AfiliacionSIS.FechaNacimiento = UserControl.txtFechaNacimiento.Text
  AfiliacionSIS.DocumentoTipo = UserControl.cmbIdDocIdentidad.Text
  AfiliacionSIS.DocumentoTipo1 = mo_cmbIdDocIdentidad.BoundText
  AfiliacionSIS.DocumentoNro = UserControl.txtNroDocumento.Text
  AfiliacionSIS.Show 1
End Sub

Private Sub cmdArchivo_Click()
    Dim oBuscaImg As New SIGHNegocios.BuscaArchivo
    oBuscaImg.MuestraImagen = True
    oBuscaImg.PathDefault = lcBuscaParametro.SeleccionaFilaParametro(236)
    oBuscaImg.MostrarFormulario
    lc_ArchivoElegido = oBuscaImg.ArchivoElegido
    If lc_ArchivoElegido <> "" Then
       pi_ImagSeleccionada.Picture = LoadPicture(lc_ArchivoElegido)
    End If
End Sub

Private Sub cmdBuscaDistrito_Click()
        Dim oBusquedaDistrito As New SIGHNegocios.BuscarDistrito
        Dim lnIdDistrito As Long, lnIdProvincia As Long
        oBusquedaDistrito.IdDepartamentoBusqueda = Val(cmbIdDepartamentoDomicilio.ItemData(cmbIdDepartamentoDomicilio.ListIndex))
        If cmbIdProvinciaDomicilio.ListIndex >= 0 Then
           oBusquedaDistrito.IdProvinciaBusqueda = Val(cmbIdProvinciaDomicilio.ItemData(cmbIdProvinciaDomicilio.ListIndex))
        End If
        oBusquedaDistrito.MostrarFormulario
        If oBusquedaDistrito.BotonPresionado = sghAceptar Then
            If oBusquedaDistrito.idRegistroSeleccionado <> 0 Then
                lnIdDistrito = oBusquedaDistrito.idRegistroSeleccionado
                cmbIdDepartamentoDomicilio_Click
                mo_cmbIdProvinciaDomicilio.BoundText = Left(Right("000" & Trim(Str(lnIdDistrito)), 6), 4)
                cmbIdProvinciaDomicilio_Click
                mo_cmbIdDistritoDomicilio.BoundText = Trim(Str(lnIdDistrito))
                
            End If
        End If
        Set oBusquedaDistrito = Nothing

End Sub

Private Sub cmdCambiaHC_Click()
    If lcBuscaParametro.SeleccionaFilaParametro(351) <> "S" Then
        MsgBox "Solo funciona si el Parametro 351=S", vbInformation, ""
    Else
        Dim oPacienteNuevaHistoria As New PacienteNuevaHistoria
        PacienteNuevaHistoria.idPaciente = ml_IdPaciente
        PacienteNuevaHistoria.NroHistoriaClinica = Val(txtIdNroHistoria.Tag)
        If Len(txtNroDocumento.Text) = 8 Then
           PacienteNuevaHistoria.NewHistoria = txtNroDocumento.Text
        End If
        PacienteNuevaHistoria.idTipoNumeracion = Val(mo_cmbIdTipoGenHistoriaClinica.BoundText)
        PacienteNuevaHistoria.Show 1
        If PacienteNuevaHistoria.BotonPresionado = sghAceptar Then
           Dim oConexion As New Connection
           oConexion.CommandTimeout = 900
           oConexion.CursorLocation = adUseClient
           oConexion.Open sighEntidades.CadenaConexion
           CargarDatosDePacienteALosControles oConexion, lcBuscaParametro.SeleccionaFilaParametro(242), _
                                              lcBuscaParametro.SeleccionaFilaParametro(287)
           oConexion.Close
           Set oConexion = Nothing
        End If
        Set PacienteNuevaHistoria = Nothing
    End If
End Sub

Private Sub cmdSectorista_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaEmpleados
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        BuscaEmpleadoYllenaDatosDelSectorista oBusqueda.idRegistroSeleccionado
    End If
    Set oBusqueda = Nothing
End Sub



Sub BuscaEmpleadoYllenaDatosDelSectorista(lnIdEmpleado As Long)
    Dim oDOEmpleado As New dOEmpleado
    Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(lnIdEmpleado)
    lblSectorista.Caption = ""
    txtSectorista.Text = ""
    If Not oDOEmpleado Is Nothing Then
        txtSectorista.Text = oDOEmpleado.IdEmpleado
        lblSectorista.Caption = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
    End If
    Set oDOEmpleado = Nothing
End Sub



Private Sub cmdSinApellidoMaterno_Click()
    If Len(txtNroDocumento.Text) <> 8 And mo_cmbIdDocIdentidad.BoundText = "1" Then
       MsgBox "Debe registrar el DNI para que el Paciente tenga un sólo apellido", vbInformation, ""
       Exit Sub
    End If
    If txtApellidoPaterno.Text = wxSinApellido Then
       MsgBox "El Paciente no tiene apellido PATERNO, no se puede haber un Paciente sin Apellido Paterno y Materno a la vez", vbInformation, ""
       Exit Sub
    End If
    'txtApellidoMaterno.Text = wxSinApellido
    txtApellidoMaterno.Text = sighEntidades.DevuelveSinApellido
End Sub

Private Sub cmdSinApellidoPaterno_Click()
    If Len(txtNroDocumento.Text) <> 8 And mo_cmbIdDocIdentidad.BoundText = "1" Then
       MsgBox "Debe registrar el DNI para que el Paciente tenga un sólo apellido", vbInformation, ""
       Exit Sub
    End If
    If txtApellidoMaterno.Text = wxSinApellido Then
       MsgBox "El Paciente no tiene apellido MATERNO, no se puede haber un Paciente sin Apellido Paterno y Materno a la vez", vbInformation, ""
       Exit Sub
    End If
    'txtApellidoPaterno.Text = wxSinApellido
    txtApellidoPaterno.Text = sighEntidades.DevuelveSinApellido
End Sub



'Private Sub grdEpicrisis_DblClick()
'     If Len(grdEpicrisis.Text) > 0 Then
'
'        FileCopy lcBuscaParametro.SeleccionaFilaParametro(236) & "\" & grdEpicrisis.Text, "c:\dibujo1.jpg"
'        Dim oCargaImg As Long
'        oCargaImg = Shell("rundll32.exe url.dll,FileProtocolHandler " & "c:\dibujo1.jpg", vbMaximizedFocus)
'     End If
'End Sub
'Sub CargaEpicrisisEscaneadas()
'     Dim lcNombreJpg As String, lcNombre As String
'     Dim lnFor As Integer, lcRuta As String
'     lcRuta = lcBuscaParametro.SeleccionaFilaParametro(237)
'     grdEpicrisis.Clear
'     For lnFor = 1 To 30
'         lcNombre = Trim(Str(txtIdNroHistoria.Tag)) & "-" & Trim(Str(lnFor)) & ".jpg"
'         lcNombreJpg = lcRuta & "\" & lcNombre
'         If sighentidades.ArchivoExiste(lcNombreJpg) Then
'            grdEpicrisis.AddItem lcNombre
'         End If
'     Next
'End Sub
'
'Sub CargaPDFgenerados()
'    Dim sArchivo As String
'    grdPDF.Clear
'    If ml_IdPaciente > 0 Then
'        sArchivo = Dir(lcBuscaParametro.SeleccionaFilaParametro(237) & "\" & Trim(Str(ml_IdPaciente)) & "*.pdf")
'        Do While sArchivo <> ""
'            grdPDF.AddItem Mid(sArchivo, InStr(sArchivo, "-") + 1)
'            sArchivo = Dir
'        Loop
'    End If
'End Sub



'Private Sub grdPDF_Click()
'     On Error GoTo ErrPDF
'     ShellExecute ml_meHwnd, vbNullString, lcBuscaParametro.SeleccionaFilaParametro(237) & "\" & _
'                  Trim(Str(ml_IdPaciente)) & "-" & grdPDF.Text, _
'                  vbNullString, "C:\", 1
'ErrPDF:
'End Sub

Private Sub tabPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    'RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub








Private Sub txtEdad_GotFocus()
    txtEdad.Tag = txtEdad.Text
End Sub

Private Sub txtEdad_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtEdad
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtEdad_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtEdad_LostFocus()
    If txtEdad.Enabled = False Or txtEdad.Locked = True Then
        Exit Sub
    End If
    If txtEdad.Text <> txtEdad.Tag Then
        Call calcularFechaDeNacimiento(txtEdad.Text, mo_cmbIdTipoEdad.BoundText)
    End If
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtEmail
End Sub

Private Sub txtFechaNacimiento_Change()
    RaiseEvent SeModificoFechaNacimiento(txtFechaNacimiento.Text, txtHoraNacimiento.Text)
End Sub

Private Sub chkIgualQueDomicilio_Click()
    
    'If MsgBox("¿Desea copiar los datos de domicilio?", vbQuestion + vbYesNo, "Pacientes") = vbYes Then
    
    If chkIgualQueDomicilio.Value = 1 Then
       mo_cmbIdPaisProcedencia.BoundText = mo_cmbIdPaisDomicilio.BoundText
       mo_cmbIdDepartamentoProcedencia.BoundText = mo_cmbIdDepartamentoDomicilio.BoundText
       mo_cmbIdProvinciaProcedencia.BoundText = mo_cmbIdProvinciaDomicilio.BoundText
       mo_cmbIdDistritoProcedencia.BoundText = mo_cmbIdDistritoDomicilio.BoundText
       mo_cmbIdCentroPobladoProcedencia.BoundText = mo_cmbIdCentroPobladoDomicilio.BoundText
       cmbIdDepartamentoProcedencia.Enabled = False
       cmbIdProvinciaProcedencia.Enabled = False
       cmbIdDistritoProcedencia.Enabled = False
       cmbIdCentroPobladoProcedencia.Enabled = False
    Else
       mo_cmbIdPaisProcedencia.BoundText = 166
       mo_cmbIdDepartamentoProcedencia.BoundText = Left(lcBuscaParametro.SeleccionaFilaParametro(242), 2)
       cmbIdDepartamentoProcedencia.Enabled = True
       cmbIdDepartamentoProcedencia_Click
       mo_cmbIdProvinciaProcedencia.BoundText = ""
       mo_cmbIdDistritoProcedencia.BoundText = ""
       mo_cmbIdCentroPobladoProcedencia.BoundText = ""
    End If
    'End If
    
End Sub

Private Sub chkIgualUQueDomicilioNac_Click()
    'If MsgBox("¿Desea copiar los datos de domicilio?", vbQuestion + vbYesNo, "Pacientes") = vbYes Then
       
    If chkIgualUQueDomicilioNac.Value = 1 Then
       mo_cmbIdPaisNacimiento.BoundText = mo_cmbIdPaisDomicilio.BoundText
       mo_cmbIdDepartamentoNacimiento.BoundText = mo_cmbIdDepartamentoDomicilio.BoundText
       mo_cmbIdProvinciaNacimiento.BoundText = mo_cmbIdProvinciaDomicilio.BoundText
       mo_cmbIdDistritoNacimiento.BoundText = mo_cmbIdDistritoDomicilio.BoundText
       mo_cmbIdCentroPobladoNacimiento.BoundText = mo_cmbIdCentroPobladoDomicilio.BoundText
       cmbIdDepartamentoNacimiento.Enabled = False
       cmbIdProvinciaNacimiento.Enabled = False
       cmbIdDistritoNacimiento.Enabled = False
       cmbIdCentroPobladoNacimiento.Enabled = False
    Else
       mo_cmbIdPaisNacimiento.BoundText = 166
       mo_cmbIdDepartamentoNacimiento.BoundText = Left(lcBuscaParametro.SeleccionaFilaParametro(242), 2)
       cmbIdDepartamentoNacimiento.Enabled = True
       cmbIdDepartamentoNacimiento_Click
       mo_cmbIdProvinciaNacimiento.BoundText = ""
       mo_cmbIdDistritoNacimiento.BoundText = ""
       mo_cmbIdCentroPobladoNacimiento.BoundText = ""
    End If
       
    'End If
End Sub

Private Sub cmbIdDepartamentoDomicilio_Click()
  If cmbIdDepartamentoDomicilio.ListIndex = -1 Then Exit Sub
       
  mo_cmbIdProvinciaDomicilio.BoundColumn = "IdProvincia"
  mo_cmbIdProvinciaDomicilio.ListField = "Nombre"
  On Error Resume Next
  Set mo_cmbIdProvinciaDomicilio.RowSource = mo_AdminServiciosGeograficos.ProvinciasSeleccionarPorDepartamento(Val(cmbIdDepartamentoDomicilio.ItemData(cmbIdDepartamentoDomicilio.ListIndex)))
       
  mo_cmbIdProvinciaDomicilio.BoundText = ""
  mo_cmbIdDistritoDomicilio.BoundText = ""
  mo_cmbIdCentroPobladoDomicilio.BoundText = ""
  cmbIdProvinciaDomicilio.Enabled = True
End Sub

Private Sub cmbIdDepartamentoProcedencia_Click()
  If cmbIdDepartamentoProcedencia.ListIndex = -1 Then Exit Sub
       
  mo_cmbIdProvinciaProcedencia.BoundColumn = "IdProvincia"
  mo_cmbIdProvinciaProcedencia.ListField = "Nombre"
  On Error Resume Next
  Set mo_cmbIdProvinciaProcedencia.RowSource = mo_AdminServiciosGeograficos.ProvinciasSeleccionarPorDepartamento(Val(cmbIdDepartamentoProcedencia.ItemData(cmbIdDepartamentoProcedencia.ListIndex)))
       
  mo_cmbIdProvinciaProcedencia.BoundText = ""
  mo_cmbIdDistritoProcedencia.BoundText = ""
  mo_cmbIdCentroPobladoProcedencia.BoundText = ""
  cmbIdProvinciaProcedencia.Enabled = True
End Sub

Private Sub cmbIdDepartamentoProcedencia_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamentoProcedencia
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub cmbIdDepartamentoProcedencia_LostFocus()
   'If cmbIdDepartamentoProcedencia.Text <> "" Then
   '    mo_cmbIdDepartamentoProcedencia.BoundText = Val(Split(cmbIdDepartamentoProcedencia.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdDepartamentoProcedencia
End Sub

Private Sub cmbIdDepartamentoProcedencia_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdProvinciaProcedencia_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdProvinciaProcedencia
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub cmbIdProvinciaProcedencia_LostFocus()
   'If cmbIdProvinciaProcedencia.Text <> "" Then
   '    mo_cmbIdProvinciaProcedencia.BoundText = Val(Split(cmbIdProvinciaProcedencia.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdProvinciaProcedencia
End Sub

Private Sub cmbIdProvinciaProcedencia_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdDistritoProcedencia_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDistritoProcedencia
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub cmbIdDistritoProcedencia_LostFocus()
   'If cmbIdDistritoProcedencia.Text <> "" Then
   '    mo_cmbIdDistritoProcedencia.BoundText = Val(Split(cmbIdDistritoProcedencia.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdDistritoProcedencia
End Sub

Private Sub cmbIdDistritoProcedencia_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdCentroPobladoProcedencia_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdCentroPobladoProcedencia
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub cmbIdCentroPobladoProcedencia_LostFocus()
   'If cmbIdCentroPobladoProcedencia.Text <> "" Then
   '    mo_cmbIdCentroPobladoProcedencia.BoundText = Val(Split(cmbIdCentroPobladoProcedencia.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdCentroPobladoProcedencia
   TabPaciente.Tab = 2
End Sub

Private Sub cmbIdCentroPobladoProcedencia_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub cmbIdDepartamentoNacimiento_Click()
        
       If cmbIdDepartamentoNacimiento.ListIndex = -1 Then Exit Sub
       mo_cmbIdProvinciaNacimiento.BoundColumn = "IdProvincia"
       mo_cmbIdProvinciaNacimiento.ListField = "Nombre"
       Set mo_cmbIdProvinciaNacimiento.RowSource = mo_AdminServiciosGeograficos.ProvinciasSeleccionarPorDepartamento(Val(cmbIdDepartamentoNacimiento.ItemData(cmbIdDepartamentoNacimiento.ListIndex)))
       
       If mo_AdminServiciosGeograficos.MensajeError <> "" Then
            MsgBox mo_AdminServiciosGeograficos.MensajeError, vbInformation, "Datos de paciente"
       End If
       
       mo_cmbIdProvinciaNacimiento.BoundText = ""
       mo_cmbIdDistritoNacimiento.BoundText = ""
       mo_cmbIdCentroPobladoNacimiento.BoundText = ""
       cmbIdProvinciaNacimiento.Enabled = True
End Sub

Private Sub cmbIdDistritoDomicilio_Click()
               
       If cmbIdDistritoDomicilio.ListIndex = -1 Then Exit Sub
       
       mo_cmbIdCentroPobladoDomicilio.BoundColumn = "IdCentroPoblado"
       mo_cmbIdCentroPobladoDomicilio.ListField = "Nombre"
       Set mo_cmbIdCentroPobladoDomicilio.RowSource = mo_AdminServiciosGeograficos.CentroPobladoSeleccionarPorDistrito(Val(cmbIdDistritoDomicilio.ItemData(cmbIdDistritoDomicilio.ListIndex)))

       If mo_AdminServiciosGeograficos.MensajeError <> "" Then
            MsgBox mo_AdminServiciosGeograficos.MensajeError, vbInformation, "Datos de paciente"
       End If
        
        mo_cmbIdCentroPobladoDomicilio.BoundText = ""
        cmbIdCentroPobladoDomicilio.Enabled = True
End Sub
Private Sub cmbIdDistritoProcedencia_Click()
       
       If cmbIdDistritoProcedencia.ListIndex = -1 Then Exit Sub
       mo_cmbIdCentroPobladoProcedencia.BoundColumn = "IdCentroPoblado"
       mo_cmbIdCentroPobladoProcedencia.ListField = "Nombre"
       Set mo_cmbIdCentroPobladoProcedencia.RowSource = mo_AdminServiciosGeograficos.CentroPobladoSeleccionarPorDistrito(Val(cmbIdDistritoProcedencia.ItemData(cmbIdDistritoProcedencia.ListIndex)))

       If mo_AdminServiciosGeograficos.MensajeError <> "" Then
            MsgBox mo_AdminServiciosGeograficos.MensajeError, vbInformation, "Datos de paciente"
       End If
        
        mo_cmbIdCentroPobladoProcedencia.BoundText = ""
        cmbIdCentroPobladoProcedencia.Enabled = True

End Sub
Private Sub cmbIdDistritoNacimiento_Click()
       
       If cmbIdDistritoNacimiento.ListIndex = -1 Then Exit Sub
       
       mo_cmbIdCentroPobladoNacimiento.BoundColumn = "IdCentroPoblado"
       mo_cmbIdCentroPobladoNacimiento.ListField = "Nombre"
       Set mo_cmbIdCentroPobladoNacimiento.RowSource = mo_AdminServiciosGeograficos.CentroPobladoSeleccionarPorDistrito(Val(cmbIdDistritoNacimiento.ItemData(cmbIdDistritoNacimiento.ListIndex)))

       If mo_AdminServiciosGeograficos.MensajeError <> "" Then
            MsgBox mo_AdminServiciosGeograficos.MensajeError, vbInformation, "Datos de paciente"
       End If
        
        mo_cmbIdCentroPobladoNacimiento.BoundText = ""
        cmbIdCentroPobladoNacimiento.Enabled = True
End Sub
Private Sub cmbIdPaisDomicilio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdPaisDomicilio
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub cmbIdPaisDomicilio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdPaisDomicilio_LostFocus()
   'If cmbIdPaisDomicilio.Text <> "" Then
   '    mo_cmbIdPaisDomicilio.BoundText = Val(Split(cmbIdPaisDomicilio.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdPaisDomicilio
End Sub

Private Sub cmbIdProvinciaDomicilio_Click()
       
    If cmbIdProvinciaDomicilio.ListIndex = -1 Then Exit Sub
       
       mo_cmbIdDistritoDomicilio.BoundColumn = "IdDistrito"
       mo_cmbIdDistritoDomicilio.ListField = "Nombre"
       Set mo_cmbIdDistritoDomicilio.RowSource = mo_AdminServiciosGeograficos.DistritoSeleccionarPorProvincia(Val(cmbIdProvinciaDomicilio.ItemData(cmbIdProvinciaDomicilio.ListIndex)))

       If mo_AdminServiciosGeograficos.MensajeError <> "" Then
            MsgBox mo_AdminServiciosGeograficos.MensajeError, vbInformation, "Datos de paciente"
       End If
       
       mo_cmbIdDistritoDomicilio.BoundText = ""
       mo_cmbIdCentroPobladoDomicilio.BoundText = ""
       cmbIdDistritoDomicilio.Enabled = True

End Sub
Private Sub cmbIdProvinciaProcedencia_Click()
       
       If cmbIdProvinciaProcedencia.ListIndex = -1 Then Exit Sub
       mo_cmbIdDistritoProcedencia.BoundColumn = "IdDistrito"
       mo_cmbIdDistritoProcedencia.ListField = "Nombre"
       Set mo_cmbIdDistritoProcedencia.RowSource = mo_AdminServiciosGeograficos.DistritoSeleccionarPorProvincia(Val(cmbIdProvinciaProcedencia.ItemData(cmbIdProvinciaProcedencia.ListIndex)))

       If mo_AdminServiciosGeograficos.MensajeError <> "" Then
            MsgBox mo_AdminServiciosGeograficos.MensajeError, vbInformation, "Datos de paciente"
       End If
       
       mo_cmbIdDistritoProcedencia.BoundText = ""
       mo_cmbIdCentroPobladoProcedencia.BoundText = ""
       cmbIdDistritoProcedencia.Enabled = True
End Sub
Private Sub cmbIdProvinciaNacimiento_Click()
      
      If cmbIdProvinciaNacimiento.ListIndex = -1 Then Exit Sub
      
       mo_cmbIdDistritoNacimiento.BoundColumn = "IdDistrito"
       mo_cmbIdDistritoNacimiento.ListField = "Nombre"
       Set mo_cmbIdDistritoNacimiento.RowSource = mo_AdminServiciosGeograficos.DistritoSeleccionarPorProvincia(Val(cmbIdProvinciaNacimiento.ItemData(cmbIdProvinciaNacimiento.ListIndex)))

       If mo_AdminServiciosGeograficos.MensajeError <> "" Then
            MsgBox mo_AdminServiciosGeograficos.MensajeError, vbInformation, "Datos de paciente"
       End If

       mo_cmbIdDistritoNacimiento.BoundText = ""
       mo_cmbIdCentroPobladoNacimiento.BoundText = ""
       cmbIdDistritoNacimiento.Enabled = True
End Sub








Private Sub txtFichaFamiliar1_GotFocus()
    lblFichaFamiliar1.Caption = "Código del Sector y/o comunidad"
End Sub

Private Sub txtFichaFamiliar1_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFichaFamiliar1
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFichaFamiliar1_LostFocus()
    lblFichaFamiliar1.Caption = ""
End Sub

Private Sub txtFichaFamiliar2_GotFocus()
    lblFichaFamiliar1.Caption = "N° Historia Clínica"
End Sub

Private Sub txtFichaFamiliar2_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFichaFamiliar2
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
    AdministrarKeyPreview KeyCode

End Sub


Private Sub txtFichaFamiliar2_LostFocus()
     lblFichaFamiliar1.Caption = ""
End Sub

Private Sub txtFichaFamiliar3_GotFocus()
    lblFichaFamiliar1.Caption = "Numeración Integrante de Familia"
End Sub

Private Sub txtFichaFamiliar3_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFichaFamiliar3
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFichaFamiliar3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtFichaFamiliar3_LostFocus
    End If
End Sub

Private Sub txtFichaFamiliar3_LostFocus()
    If (mi_Opcion = sghAgregar Or mi_Opcion = sghModificar) And Len(txtFichaFamiliar1.Text) > 0 And Len(txtFichaFamiliar2.Text) > 0 And Len(txtFichaFamiliar3.Text) > 0 Then
       Dim lcFichaFamiliar As String
       txtFichaFamiliar1.Text = UCase(txtFichaFamiliar1.Text)
       txtFichaFamiliar2.Text = UCase(txtFichaFamiliar2.Text)
       txtFichaFamiliar3.Text = UCase(txtFichaFamiliar3.Text)
       Select Case lcFormaQgeneraHistoria
       Case "1"
            mo_cmbIdTipoGenHistoriaClinica.BoundText = "1"
       Case "2"
            mo_cmbIdTipoGenHistoriaClinica.BoundText = "2"
       Case "3"   'Formula 1
            mo_cmbIdTipoGenHistoriaClinica.BoundText = "2"
            txtIdNroHistoria.Text = Val(Trim(txtFichaFamiliar1.Text) & Trim(txtFichaFamiliar2.Text) & Trim(txtFichaFamiliar3.Text))
       End Select
       On Error Resume Next
       lcFichaFamiliar = DevuelveFichaFamiliarUnida 'txtFichaFamiliar1.Text & "-" & txtFichaFamiliar2.Text & "-" & txtFichaFamiliar3.Text
       ms_MensajeError = mo_AdminAdmision.ExisteFichaFamiliar(lcFichaFamiliar, ml_IdPaciente)
       If ms_MensajeError <> "" Then
           MsgBox "Existe un paciente con la misma FICHA FAMILIAR: " + Chr(13) + ms_MensajeError
           txtFichaFamiliar3.Text = ""
           txtFichaFamiliar3.SetFocus
       End If
    End If
    lblFichaFamiliar1.Caption = ""
End Sub

Private Sub txtHoraNacimiento_Change()
    RaiseEvent SeModificoFechaNacimiento(txtFechaNacimiento.Text, txtHoraNacimiento.Text)

End Sub

Private Sub txtHoraNacimiento_GotFocus()
     If txtHoraNacimiento.Text = "00:00" Then
        txtHoraNacimiento.Text = sighEntidades.HORA_VACIA_HM
     End If
End Sub

Private Sub txtHoraNacimiento_LostFocus()
       If txtHoraNacimiento.Text = sighEntidades.HORA_VACIA_HM Then
          txtHoraNacimiento.Text = "00:00"
       End If
       If txtFechaNacimiento <> sighEntidades.FECHA_VACIA_DMY And txtHoraNacimiento.Text <> sighEntidades.HORA_VACIA_HM Then
            If Not EsFecha(txtFechaNacimiento, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, "Datos de paciente"
                 txtFechaNacimiento = sighEntidades.FECHA_VACIA_DMY
            ElseIf Not EsHora(txtHoraNacimiento.Text) Then
                MsgBox "La HORA ingresada no es válida", vbInformation, "Datos de paciente"
                txtHoraNacimiento.Text = sighEntidades.HORA_VACIA_HM
            ElseIf CDate(txtFechaNacimiento.Text & " " & txtHoraNacimiento.Text) > ldHoy Then
                MsgBox "La fecha ingresada debe ser mayor a la Fecha actual", vbInformation, "Datos de paciente"
                'mgaray20141008
                On Error Resume Next
                txtFechaNacimiento = sighEntidades.FECHA_VACIA_DMY
                txtFechaNacimiento.SetFocus
    
            Else
                ActualizaEdad
                'txtEdad.Text = Trim(Str(EdadActual(CDate(txtFechaNacimiento.Text & " " & txtHoraNacimiento.Text), Now)))
            End If
            
        End If

End Sub

Sub ActualizaEdad()
    Dim oEdad As Edad
    oEdad = CalcularEdad(CDate(txtFechaNacimiento.Text & " " & txtHoraNacimiento.Text), Now)
    txtEdad.Text = oEdad.Edad
    lblTipoEdad.Caption = oEdad.NombreEdad
    mo_cmbIdTipoEdad.BoundText = oEdad.TipoEdad
End Sub





Private Sub txtMadreApellidoM_KeyDown(KeyCode As Integer, Shift As Integer)
 mo_Teclado.RealizarNavegacion KeyCode, txtMadreApellidoM
  RaiseEvent SePresionoTeclaEspecial(KeyCode)
  AdministrarKeyPreview KeyCode

End Sub

'A.Yañez 30-10-2014 **************************************************
Private Sub txtMadreApellidoM_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
'********************************************************************


Private Sub txtMadreApellidoM_LostFocus()
    txtMadreApellidoM.Text = UCase(txtMadreApellidoM.Text)
    mo_Formulario.MarcarComoVacio txtMadreApellidoM
End Sub

Private Sub txtMadreApellidoP_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtMadreApellidoP
  RaiseEvent SePresionoTeclaEspecial(KeyCode)
  AdministrarKeyPreview KeyCode

End Sub

'A.Yañez 30-10-2014***************************************************
Private Sub txtMadreApellidoP_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
'*********************************************************************


Private Sub txtMadreApellidoP_LostFocus()
    txtMadreApellidoP.Text = UCase(txtMadreApellidoP.Text)
    mo_Formulario.MarcarComoVacio txtMadreApellidoP

End Sub

Private Sub txtMadreDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtMadreDocumento
  RaiseEvent SePresionoTeclaEspecial(KeyCode)
  AdministrarKeyPreview KeyCode

End Sub

Private Sub txtMadreDocumento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If mo_cmbMadreTipoDocumento.BoundText = "1" Then
        If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
            KeyAscii = 0
        End If
       End If
   End If

End Sub

Private Sub txtMadreDocumento_LostFocus()
       mo_Formulario.MarcarComoVacio txtMadreDocumento
       If txtMadreDocumento.Text <> "" Then
          Dim rspacientes As New Recordset
          Set rspacientes = mo_AdminAdmision.PacientesFiltraPorNroDocumentoYtipo(txtMadreDocumento.Text, Val(mo_cmbMadreTipoDocumento.BoundText))
          If rspacientes.RecordCount > 0 Then
             'If rspacientes.Fields!idTipoSexo = 2 Then
                txtMadreApellidoP.Text = rspacientes.Fields!ApellidoPaterno
                txtMadreApellidoM.Text = rspacientes.Fields!ApellidoMaterno
                txtNombreMadre.Text = rspacientes.Fields!PrimerNombre
                txtMadreSnombre.Text = IIf(IsNull(rspacientes.Fields!SegundoNombre), "", rspacientes.Fields!SegundoNombre)
             'End If
             
             '<(Inicio) Añadido Por: WABG el: 16/10/2020-12:10:08 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
             Else
                CargarDatosTutorDesdeRENIEC (Trim(txtMadreDocumento.Text))
             '</(Fin) Añadido Por: WABG el: 16/10/2020-12:10:08 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
                
          End If
          rspacientes.Close
          Set rspacientes = Nothing
          
          '****buscar a la madre en la RENIEC
'          <(Inicio)Comentado Por: WABG el: 16/10/2020-12:13:01 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
'          If lbBuscaDNIenReniec = True And Len(txtMadreDocumento.Text) = 8 And Val(mo_cmbMadreTipoDocumento.BoundText) = 1 Then
'               Dim lbContinuar As Boolean
'               lbContinuar = True
'               If mi_Opcion <> sghAgregar Then
'                  If txtMadreApellidoP.Text <> "" Then
'                     lbContinuar = False
'                  End If
'               Else
'                  If mb_MarcoCheckPacienteNuevo = False Then
'                     lbContinuar = False
'                  End If
'               End If
'               If lbContinuar = True Then
'                     mo_Reniec.ConsultarDNIenReniec txtMadreDocumento.Text
'                     If mo_Reniec.ApellidoPaterno <> "" Then
'                           txtMadreApellidoP.Text = mo_Reniec.ApellidoPaterno
'                           txtMadreApellidoM.Text = mo_Reniec.ApellidoMaterno
'                           txtNombreMadre.Text = mo_Reniec.PrimerNombre
'                           txtMadreSnombre.Text = mo_Reniec.SegundoNombre
''                           mb_UsoWebReniec = True
''                           MuestraQueUsoWebReniec
'                     End If
'               End If
'          End If
'          </(Fin)Comentado por: WABG el: 16/10/2020-12:13:01 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
        End If
End Sub

Private Sub txtMadreSnombre_KeyDown(KeyCode As Integer, Shift As Integer)
 mo_Teclado.RealizarNavegacion KeyCode, txtMadreSnombre
  RaiseEvent SePresionoTeclaEspecial(KeyCode)
  AdministrarKeyPreview KeyCode

End Sub
'A.Yañez 30-10-2014 **********************************************************
Private Sub txtMadreSnombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
'*****************************************************************************

Private Sub txtMadreSnombre_LostFocus()
    txtMadreSnombre.Text = UCase(txtMadreSnombre.Text)
    mo_Formulario.MarcarComoVacio txtMadreSnombre
End Sub

Private Sub txtNombreMadre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombreMadre
RaiseEvent SePresionoTeclaEspecial(KeyCode)
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNombreMadre_LostFocus()
  'txtNombreMadre.Text = mo_Teclado.CapitalizarNombres(txtNombreMadre.Text)
   txtNombreMadre.Text = UCase(txtNombreMadre.Text)
   mo_Formulario.MarcarComoVacio txtNombreMadre
End Sub

Private Sub txtNombreMadre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtNombrePadre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNombrePadre
RaiseEvent SePresionoTeclaEspecial(KeyCode)
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNombrePadre_LostFocus()
txtNombrePadre.Text = mo_Teclado.CapitalizarNombres(txtNombrePadre.Text)
   mo_Formulario.MarcarComoVacio txtNombrePadre
End Sub

Private Sub txtNombrePadre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdCentroPobladoDomicilio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdCentroPobladoDomicilio
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub cmbIdCentroPobladoDomicilio_LostFocus()
   'If cmbIdCentroPobladoDomicilio.Text <> "" Then
   '    mo_cmbIdCentroPobladoDomicilio.BoundText = Val(Split(cmbIdCentroPobladoDomicilio.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdCentroPobladoDomicilio
End Sub

Private Sub cmbIdCentroPobladoDomicilio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdDistritoDomicilio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDistritoDomicilio
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub cmbIdDistritoDomicilio_LostFocus()
   'If cmbIdDistritoDomicilio.Text <> "" Then
   '    mo_cmbIdDistritoDomicilio.BoundText = Val(Split(cmbIdDistritoDomicilio.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdDistritoDomicilio
End Sub

Private Sub cmbIdDistritoDomicilio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdProvinciaDomicilio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdProvinciaDomicilio
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub cmbIdProvinciaDomicilio_LostFocus()
   
   'If cmbIdProvinciaDomicilio.Text <> "" Then
   '    mo_cmbIdProvinciaDomicilio.BoundText = Val(Split(cmbIdProvinciaDomicilio.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdProvinciaDomicilio
   
End Sub

Private Sub cmbIdProvinciaDomicilio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdDepartamentoDomicilio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamentoDomicilio
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub cmbIdDepartamentoDomicilio_LostFocus()
   'If cmbIdDepartamentoDomicilio.Text <> "" Then
   '    mo_cmbIdDepartamentoDomicilio.BoundText = Val(Split(cmbIdDepartamentoDomicilio.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdDepartamentoDomicilio
End Sub

Private Sub cmbIdDepartamentoDomicilio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdDistritoNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDistritoNacimiento
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub cmbIdDistritoNacimiento_LostFocus()
   'If cmbIdDistritoNacimiento.Text <> "" Then
   '    mo_cmbIdDistritoNacimiento.BoundText = Val(Split(cmbIdDistritoNacimiento.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdDistritoNacimiento
End Sub

Private Sub cmbIdDistritoNacimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdProvinciaNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdProvinciaNacimiento
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub cmbIdProvinciaNacimiento_LostFocus()
   'If cmbIdProvinciaNacimiento.Text <> "" Then
   '    mo_cmbIdProvinciaNacimiento.BoundText = Val(Split(cmbIdProvinciaNacimiento.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdProvinciaNacimiento
End Sub

Private Sub cmbIdProvinciaNacimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdDepartamentoNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamentoNacimiento
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub cmbIdDepartamentoNacimiento_LostFocus()
   'If cmbIdDepartamentoNacimiento.Text <> "" Then
   '    mo_cmbIdDepartamentoNacimiento.BoundText = Val(Split(cmbIdDepartamentoNacimiento.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdDepartamentoNacimiento
End Sub

Private Sub cmbIdDepartamentoNacimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdTipoOcupacion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoOcupacion
RaiseEvent SePresionoTeclaEspecial(KeyCode)
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoOcupacion_LostFocus()
    
'   If cmbIdTipoOcupacion.Text <> "" Then
'        On Error Resume Next
'       mo_cmbIdTipoOcupacion.BoundText = Val(Split(cmbIdTipoOcupacion.Text, " = ")(0))
'       If Err.Number <> 0 Then
'        cmbIdTipoOcupacion.Text = ""
'       End If
'   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoOcupacion
End Sub

Private Sub cmbIdTipoOcupacion_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
End Sub


Private Sub cmbIdDocIdentidad_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, cmbIdDocIdentidad
  RaiseEvent SePresionoTeclaEspecial(KeyCode)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdDocIdentidad_LostFocus()
   If cmbIdDocIdentidad.Text <> "" Then
        On Error Resume Next
        If sighEntidades.Parametro584valorInt = "1" Then
           mo_cmbIdTipoGenHistoriaClinica.BoundText = lcIdTipoGenHistoriaClinicaActual
        End If
       mo_cmbIdDocIdentidad.BoundText = Val(Split(cmbIdDocIdentidad.Text, " = ")(0))
        If mo_cmbIdDocIdentidad.BoundText = "1" Then
            txtNroDocumento.MaxLength = 8
        Else
            txtNroDocumento.MaxLength = 12
            If mo_cmbIdDocIdentidad.BoundText = "2" And sighEntidades.Parametro584valorInt = "1" Then
               mo_cmbIdTipoGenHistoriaClinica.BoundText = "1"
            End If
        End If
        
   End If
   mo_Formulario.MarcarComoVacio cmbIdDocIdentidad
   mo_Formulario.MarcarComoVacio cmbIdDocIdentidad
End Sub

Private Sub cmbIdDocIdentidad_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdEstadoCivil_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdEstadoCivil
RaiseEvent SePresionoTeclaEspecial(KeyCode)
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdEstadoCivil_LostFocus()
    
   'If cmbIdEstadoCivil.Text <> "" Then
   '     On Error Resume Next
   '    mo_cmbIdEstadoCivil.BoundText = Val(Split(cmbIdEstadoCivil.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbIdEstadoCivil
End Sub

Private Sub cmbIdEstadoCivil_KeyPress(KeyAscii As Integer)
  ' If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
  '     If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
  '         KeyAscii = 0
  '     End If
  ' End If
End Sub


Private Sub cmbIdGradoInstruccion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdGradoInstruccion
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdGradoInstruccion_LostFocus()
'   If cmbIdGradoInstruccion.Text <> "" Then
'       mo_cmbIdGradoInstruccion.BoundText = Val(Split(cmbIdGradoInstruccion.Text, " = ")(0))
'   End If
   mo_Formulario.MarcarComoVacio cmbIdGradoInstruccion
End Sub

Private Sub cmbIdGradoInstruccion_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
End Sub


Private Sub cmbIdProcedencia_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdProcedencia
RaiseEvent SePresionoTeclaEspecial(KeyCode)
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdProcedencia_LostFocus()
'   If cmbIdProcedencia.Text <> "" Then
'        On Error Resume Next
'       mo_cmbIdProcedencia.BoundText = Val(Split(cmbIdProcedencia.Text, " = ")(0))
'   End If
   mo_Formulario.MarcarComoVacio cmbIdProcedencia
End Sub

Private Sub cmbIdProcedencia_KeyPress(KeyAscii As Integer)
'   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
'   End If
End Sub

Private Sub cmbIdTipoSexo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoSexo
RaiseEvent SePresionoTeclaEspecial(KeyCode)
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoSexo_LostFocus()
   If cmbIdTipoSexo.Text <> "" Then
        On Error Resume Next
       mo_CmbIdTipoSexo.BoundText = Val(Split(cmbIdTipoSexo.Text, " = ")(0))
       
       If Err.Number <> 0 Then
        cmbIdTipoSexo.Text = ""
       End If
       
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoSexo
End Sub

Private Sub cmbIdTipoSexo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtDireccionDomicilio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDireccionDomicilio
RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtDireccionDomicilio_LostFocus()
    txtDireccionDomicilio.Text = mo_Teclado.CapitalizarNombres(txtDireccionDomicilio.Text)
    mo_Formulario.MarcarComoVacio txtDireccionDomicilio
End Sub

Private Sub txtDireccionDomicilio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
'       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
'           KeyAscii = 0
'       End If
   End If
End Sub

Private Sub txtNroDomicilio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumeroDeDomicilio(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtNroDocumento_KeyUp(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroDocumento
'   RaiseEvent SePresionoTeclaEspecial(KeyCode) 'Actualizado 13102014
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroHijo_KeyDown(KeyCode As Integer, Shift As Integer)
 mo_Teclado.RealizarNavegacion KeyCode, txtNroHijo
  RaiseEvent SePresionoTeclaEspecial(KeyCode)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroHijo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If

End Sub

Private Sub txtNroHijo_LostFocus()
    mo_Formulario.MarcarComoVacio txtNroHijo
End Sub

Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtObservacion
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub



Private Sub txtPisoDomicilio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub




Private Sub txtManzanaDomicilio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub




Private Sub txtLoteDomicilio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtSectorDomicilio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub




Private Sub txtEtapaDomicilio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub





Private Sub txtSector_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, txtSector
  RaiseEvent SePresionoTeclaEspecial(KeyCode)
  AdministrarKeyPreview KeyCode

End Sub

Private Sub txtSector_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtNroDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroDocumento
   RaiseEvent SePresionoTeclaEspecial(KeyCode)
   AdministrarKeyPreview KeyCode
End Sub


Private Sub txtNroDocumento_LostFocus()
   
   If Len(Trim(txtNroDocumento.Text)) > 0 Then
        If mo_cmbIdDocIdentidad.BoundText = "1" And Len(Trim(txtNroDocumento.Text)) <> 8 Then
           MsgBox "Si el Documento es DNI debe tener 8 dígitos", vbInformation, "Mensaje"
           On Error Resume Next
           txtNroDocumento.SetFocus
           Exit Sub
        End If
   Else
        Exit Sub
   End If
   '
   Dim rspacientes As New Recordset
   Set rspacientes = mo_AdminAdmision.PacientesFiltraPorNroDocumentoYtipo(Trim(txtNroDocumento.Text), Val(mo_cmbIdDocIdentidad.BoundText))
   If rspacientes.RecordCount > 0 Then
         If mi_Opcion = sghAgregar Or mi_Opcion = sghModificar Then
         '   rspacientes.MoveFirst
         '   MsgBox "Es N°DOCUMENTO ya existe para el Paciente: " + Trim(Str(rspacientes!NroHistoriaClinica)) + " " + rspacientes!ApellidoPaterno + " " + rspacientes!ApellidoMaterno + " " + rspacientes!PrimerNombre, vbInformation, "Datos de paciente"
         '   rspacientes.Close
         '   Set rspacientes = Nothing
            'On Error Resume Next
            'txtNroDocumento.SetFocus
         '   Exit Sub
         'ElseIf mi_Opcion = sghModificar Then
               rspacientes.MoveFirst
               Do While Not rspacientes.EOF
                  If rspacientes!idPaciente <> ml_IdPaciente Then
                        MsgBox "El N° DOCUMENTO ya existe para el Paciente: " + _
                                 HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rspacientes!NroHistoriaClinica)), False) + _
                                 " " + rspacientes!ApellidoPaterno + " " + rspacientes!ApellidoMaterno + " " + rspacientes!PrimerNombre, vbInformation, "Datos de paciente"
                        rspacientes.Close
                        Set rspacientes = Nothing
                        On Error Resume Next
                        txtNroDocumento.SetFocus
                        Exit Sub
                  End If
                  rspacientes.MoveNext
               Loop
         End If
'<(Inicio) Añadido Por: WABG el: 16/10/2020-11:46:19 a.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
         Else
    
    'CARGAR CONTROLES DE TEXTO DESDE RENIEC
  '  CargarDatosDesdeRENIEC ((Trim(txtNroDocumento.Text)))
    
'</(Fin) Añadido Por: WABG el: 16/10/2020-11:46:19 a.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
   End If
   rspacientes.Close
   '
   mo_Formulario.MarcarComoVacio txtNroDocumento
   'If lbBuscaDNIenReniec = True And Len(txtNroDocumento.Text) = 8 Then
   If lbBuscaDNIenReniec = True And Len(txtNroDocumento.Text) = 8 And mo_cmbIdDocIdentidad.BoundText = "1" Then
      Dim lbContinuar As Boolean
      lbContinuar = True
      If mi_Opcion <> sghAgregar Then
         If txtApellidoPaterno.Text <> "" Then
            lbContinuar = False
         End If
      Else
         If mb_MarcoCheckPacienteNuevo = False Then
            lbContinuar = False
         End If
      End If
      If lbContinuar = True Then
            mo_Reniec.ConsultarDNIenReniec txtNroDocumento.Text
            If mo_Reniec.ApellidoPaterno <> "" Then
'                  <(Inicio)Comentado Por: WABG el: 16/10/2020-12:19:18 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
'                  Dim lcIdDistrito As String
'                  </(Fin)Comentado por: WABG el: 16/10/2020-12:19:18 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
                  txtApellidoPaterno.Text = mo_Reniec.ApellidoPaterno
                  txtApellidoMaterno.Text = mo_Reniec.ApellidoMaterno
                  txtPrimerNombre.Text = mo_Reniec.PrimerNombre
                  txtSegundoNombre.Text = mo_Reniec.SegundoNombre
                  txtTercerNombre.Text = mo_Reniec.TercerNombre
                  txtFechaNacimiento.Text = mo_Reniec.FechaNacimiento
                  mo_CmbIdTipoSexo.BoundText = mo_Reniec.idTipoSexo
                  txtDireccionDomicilio.Text = mo_Reniec.DireccionDomicilio
                  If mo_Reniec.IdDistritoDomicilio > 0 Then
                     lcIdDistrito = Right("0" & Trim(Str(mo_Reniec.IdDistritoDomicilio)), 6)
                     mo_cmbIdDepartamentoDomicilio.BoundText = Left(lcIdDistrito, 2)
                     mo_cmbIdProvinciaDomicilio.BoundText = Left(lcIdDistrito, 4)
                     mo_cmbIdDistritoDomicilio.BoundText = lcIdDistrito
                  End If
                  mb_UsoWebReniec = True
                  MuestraQueUsoWebReniec
'<(Inicio) Añadido Por: WABG el: 26/01/2021-11:57:02 a.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
                    'Idioma español = 38 en el combobox
                  cmbIdioma.ListIndex = 38
                  
                  'Etnia mestizo = 43 en el combobox
                  cmbEtnia.ListIndex = 43
                  
'</(Fin) Añadido Por: WABG el: 26/01/2021-11:57:02 a.m. en el Equipo: SISGALENPLUS-PC<CAMBIO-37>
            End If
      End If
      UserControl.TabPaciente.Tab = 0
   End If
   Set rspacientes = Nothing
   '******** Nº Historia = Nº DNI
   If mb_MarcoCheckPacienteNuevo = True And wxParametro351 = "S" And txtIdNroHistoria.Locked = False Then
      If Val(txtNroDocumento.Text) > 0 Then
        Dim lnHC11 As Long
        lnHC11 = Val(txtNroDocumento.Text)
        'GLCC 02/11/20 CAMBIO36 INICIO
        'Quita el wxNueve &
        'txtIdNroHistoria.Tag = wxNueve & Trim(Str(lnHC11))
        txtIdNroHistoria.Tag = Trim$(Str(lnHC11))
       'GLCC 02/11/20 CAMBIO36 FIN
        txtIdNroHistoria.Text = Trim(Str(lnHC11))
      End If
   End If
End Sub

Private Sub txtNroDocumento_KeyPress(KeyAscii As Integer)
       'Actualizado 20140919
       If Val(mo_cmbIdDocIdentidad.BoundText) = 1 Then
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If
       Else
            If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
                KeyAscii = 0
            End If
       End If
End Sub

Private Sub txtFechaNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaNacimiento
'   AdministrarKeyPreview KeyCode
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtFechaNacimiento_LostFocus()
       If txtFechaNacimiento <> sighEntidades.FECHA_VACIA_DMY Then
            On Error Resume Next
            If txtHoraNacimiento.Text = sighEntidades.HORA_VACIA_HM Then
               txtHoraNacimiento.Text = "00:00"
            End If
            If Not EsFecha(txtFechaNacimiento, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, "Datos de paciente"
                txtFechaNacimiento = sighEntidades.FECHA_VACIA_DMY
                txtFechaNacimiento.SetFocus
            ElseIf CDate(txtFechaNacimiento.Text & " " & txtHoraNacimiento.Text) > ldHoy Then
                MsgBox "La fecha ingresada NO debe ser mayor a la Fecha actual", vbInformation, "Datos de paciente"
                txtFechaNacimiento = sighEntidades.FECHA_VACIA_DMY
                txtFechaNacimiento.SetFocus
            Else
                'txtEdad.Text = Trim(Str(EdadActual(CDate(txtFechaNacimiento.Text), Date)))
                ActualizaEdad
            End If
            
        End If
   mo_Formulario.MarcarComoVacio txtFechaNacimiento
End Sub

Private Sub txtFechaNacimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtTercerNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtTercerNombre
RaiseEvent SePresionoTeclaEspecial(KeyCode)
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtTercerNombre_LostFocus()
  txtTercerNombre.Text = mo_Teclado.CapitalizarNombres(txtTercerNombre.Text)
  txtTercerNombre.Text = mo_Teclado.DevuelveTextoSINtildes(txtTercerNombre.Text)
  'mo_Formulario.MarcarComoVacio txtTercerNombre
  VerificaSiExistePaciente txtApellidoPaterno.Text, txtApellidoMaterno.Text, txtPrimerNombre.Text, txtSegundoNombre.Text, txtTercerNombre.Text
End Sub

Private Sub txtTercerNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtSegundoNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtSegundoNombre
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtSegundoNombre_LostFocus()
    If txtSegundoNombre.Text <> "NN" Then
        txtSegundoNombre.Text = mo_Teclado.CapitalizarNombres(txtSegundoNombre.Text)
        txtSegundoNombre.Text = mo_Teclado.DevuelveTextoSINtildes(txtSegundoNombre.Text)
    End If
    VerificaSiExistePaciente txtApellidoPaterno.Text, txtApellidoMaterno.Text, txtPrimerNombre.Text, txtSegundoNombre.Text, txtTercerNombre.Text
End Sub
'A.Yañez 30-10-2014******************************************
Private Sub txtSegundoNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
'************************************************************


Private Sub txtPrimerNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombre
RaiseEvent SePresionoTeclaEspecial(KeyCode)
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtPrimerNombre_LostFocus()
    
    If txtPrimerNombre.Text <> "NN" Then
       'txtPrimerNombre.Text = mo_Teclado.CapitalizarNombres(txtPrimerNombre.Text)
        txtPrimerNombre.Text = UCase(txtPrimerNombre.Text)
        txtPrimerNombre.Text = mo_Teclado.DevuelveTextoSINtildes(txtPrimerNombre.Text)
    End If
    If UCase(Right(txtPrimerNombre.Text, 1)) = "A" Then
        mo_CmbIdTipoSexo.BoundText = "2"
    ElseIf UCase(Right(txtPrimerNombre.Text, 1)) = "O" Then
        mo_CmbIdTipoSexo.BoundText = "1"
    End If
    mo_Formulario.MarcarComoVacio txtPrimerNombre
  VerificaSiExistePaciente txtApellidoPaterno.Text, txtApellidoMaterno.Text, txtPrimerNombre.Text, txtSegundoNombre.Text, txtTercerNombre.Text
End Sub

Private Sub txtPrimerNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
RaiseEvent SePresionoTeclaEspecial(KeyCode)
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoMaterno_LostFocus()
    If txtApellidoMaterno.Text <> "NN" Then
        'txtApellidoMaterno.Text = mo_Teclado.CapitalizarNombres(txtApellidoMaterno.Text)
        txtApellidoMaterno.Text = UCase(txtApellidoMaterno.Text)
        txtApellidoMaterno.Text = mo_Teclado.DevuelveTextoSINtildes(txtApellidoMaterno.Text)
    End If
   mo_Formulario.MarcarComoVacio txtApellidoMaterno
  VerificaSiExistePaciente txtApellidoPaterno.Text, txtApellidoMaterno.Text, txtPrimerNombre.Text, txtSegundoNombre.Text, txtTercerNombre.Text
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
  RaiseEvent SePresionoTeclaEspecial(KeyCode)
  AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoPaterno_LostFocus()
    If txtApellidoPaterno.Text <> "NN" Then
        'txtApellidoPaterno.Text = mo_Teclado.CapitalizarNombres(txtApellidoPaterno.Text)
        txtApellidoPaterno.Text = UCase(txtApellidoPaterno.Text)
        txtApellidoPaterno.Text = mo_Teclado.DevuelveTextoSINtildes(txtApellidoPaterno.Text)
    End If
    mo_Formulario.MarcarComoVacio txtApellidoPaterno
    VerificaSiExistePaciente txtApellidoPaterno.Text, txtApellidoMaterno.Text, txtPrimerNombre.Text, txtSegundoNombre.Text, txtTercerNombre.Text
End Sub

Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Public Sub ConfigurarComboBoxes()
Dim sMensaje As String
        
        mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
        mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()
        
        
        'CARGA COMBO BOXES DE PACIENTE
        mo_CmbIdTipoSexo.BoundColumn = "IdtipoSexo"
        mo_CmbIdTipoSexo.ListField = "DescripcionLarga"
        Set mo_CmbIdTipoSexo.RowSource = mo_AdminServiciosComunes.TiposSexoSeleccionarTodos()
        sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
       
'        mo_cmbIdProcedencia.BoundColumn = "IdProcedencia"
'        mo_cmbIdProcedencia.ListField = "DescripcionLarga"
'        Set mo_cmbIdProcedencia.RowSource = mo_AdminServiciosComunes.TiposProcedenciaSeleccionarTodos()
'        sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
        Set cmbIdProcedencia.ListSource = mo_AdminServiciosComunes.TiposProcedenciaTodos()
        
'        mo_cmbIdGradoInstruccion.BoundColumn = "IdGradoInstruccion"
'        mo_cmbIdGradoInstruccion.ListField = "DescripcionLarga"
'        Set mo_cmbIdGradoInstruccion.RowSource = mo_AdminServiciosComunes.TiposGradosInstruccionSeleccionarTodos()
'        sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
        Set cmbIdGradoInstruccion.ListSource = mo_AdminServiciosComunes.TiposGradosInstruccionTodos()
        
'        mo_cmbIdEstadoCivil.BoundColumn = "IdEstadoCivil"
'        mo_cmbIdEstadoCivil.ListField = "DescripcionLarga"
'        Set mo_cmbIdEstadoCivil.RowSource = mo_AdminServiciosComunes.TiposEstadoCivilSeleccionarTodos()
'        sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
        Set cmbIdEstadoCivil.ListSource = mo_AdminServiciosComunes.TiposEstadoCivilTodos()
        
        
        mo_cmbIdDocIdentidad.BoundColumn = "IdDocIdentidad"
        mo_cmbIdDocIdentidad.ListField = "DescripcionLarga"
        Set mo_cmbIdDocIdentidad.RowSource = mo_AdminServiciosComunes.TiposDocIdentidadSeleccionarTodosIncSinTipoDoc()
        SendMessage cmbIdDocIdentidad.hwnd, cb_setdroppedwidth, 250, 0 'A.Yañez 11-11-2014
        sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
        
'        mo_cmbIdTipoOcupacion.BoundColumn = "IdTipoOcupacion"
'        mo_cmbIdTipoOcupacion.ListField = "DescripcionLarga"
'        Set mo_cmbIdTipoOcupacion.RowSource = mo_AdminServiciosComunes.TiposOcupacionSeleccionarTodos()
'        sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
        Set cmbIdTipoOcupacion.ListSource = mo_AdminServiciosComunes.TiposOcupacionTodos()
        
        mo_cmbIdPaisDomicilio.BoundColumn = "IdPais"
        mo_cmbIdPaisDomicilio.ListField = "Nombre"
        sMensaje = sMensaje + mo_AdminServiciosGeograficos.MensajeError
        Set oRsPaisDomicilio = mo_AdminServiciosGeograficos.PaisesSeleccionarTodos()
        Set mo_cmbIdPaisDomicilio.RowSource = oRsPaisDomicilio
        
        mo_cmbIdPaisProcedencia.BoundColumn = "IdPais"
        mo_cmbIdPaisProcedencia.ListField = "Nombre"
        sMensaje = sMensaje + mo_AdminServiciosGeograficos.MensajeError
        Set oRsPaisProcedencia = oRsPaisDomicilio.Clone()
        Set mo_cmbIdPaisProcedencia.RowSource = oRsPaisProcedencia
        
        mo_cmbIdPaisNacimiento.BoundColumn = "IdPais"
        mo_cmbIdPaisNacimiento.ListField = "Nombre"
        sMensaje = sMensaje + mo_AdminServiciosGeograficos.MensajeError
        Set oRsPaisNacimiento = oRsPaisDomicilio.Clone()
        Set mo_cmbIdPaisNacimiento.RowSource = oRsPaisNacimiento
        
        mo_cmbIdDepartamentoNacimiento.BoundColumn = "IdDepartamento"
        mo_cmbIdDepartamentoNacimiento.ListField = "Nombre"
        sMensaje = sMensaje + mo_AdminServiciosGeograficos.MensajeError
        Set oRsDptoNacimiento = mo_AdminServiciosGeograficos.DepartamentosSeleccionarTodos()
        Set mo_cmbIdDepartamentoNacimiento.RowSource = oRsDptoNacimiento
        
        mo_cmbIdDepartamentoDomicilio.BoundColumn = "IdDepartamento"
        mo_cmbIdDepartamentoDomicilio.ListField = "Nombre"
        mo_cmbIdDepartamentoDomicilio.BoundText = Left(lcBuscaParametro.SeleccionaFilaParametro(242), 2)
        sMensaje = sMensaje + mo_AdminServiciosGeograficos.MensajeError
        Set oRsDptoDomicilio = oRsDptoNacimiento.Clone()
        Set mo_cmbIdDepartamentoDomicilio.RowSource = oRsDptoDomicilio
    
        mo_cmbIdDepartamentoProcedencia.BoundColumn = "IdDepartamento"
        mo_cmbIdDepartamentoProcedencia.ListField = "Nombre"
        sMensaje = sMensaje + mo_AdminServiciosGeograficos.MensajeError
        Set oRsDptoProcedencia = oRsDptoNacimiento.Clone()
        Set mo_cmbIdDepartamentoProcedencia.RowSource = oRsDptoProcedencia
        
        Set cmbEtnia.ListSource = mo_AdminServiciosComunes.EtniaHISseleccionarTodos()
        
        Set cmbIdioma.ListSource = mo_AdminServiciosComunes.TiposIdiomasSeleccionarTodos
        
        mo_cmbMadreTipoDocumento.BoundColumn = "IdDocIdentidad"
        mo_cmbMadreTipoDocumento.ListField = "DescripcionLarga"
        Set mo_cmbMadreTipoDocumento.RowSource = mo_AdminServiciosComunes.TiposDocIdentidadSeleccionarTodosIncSinTipoDoc()
        mo_cmbMadreTipoDocumento.BoundText = "1"
        SendMessage cmbMadreTipoDocumento.hwnd, cb_setdroppedwidth, 250, 0 'A.Yañez
        sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
        
        mo_cmbIdTipoEdad.BoundColumn = "IdTipoEdad"
        mo_cmbIdTipoEdad.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoEdad.RowSource = mo_AdminServiciosComunes.TiposEdadSeleccionarTodos

        If sMensaje <> "" Then
            MsgBox sMensaje, vbInformation, "Datos de paciente"
        End If


End Sub

Public Function ValidarDatosObligatorios(wxParametro282 As String, wxParametro333 As String) As String
Dim sMensajeLocal As String
Dim ml_ColorError As Long
ml_ColorError = CLng(lcBuscaParametro.SeleccionaFilaParametro(343))
Dim mb_FaltaDato As Boolean

 'mgaray20141008
   On Error Resume Next
   
   'Yamill Palomino 2508 ini
   
   If mo_cmbIdTipoGenHistoriaClinica.BoundText = "" Then
       sMensajeLocal = sMensajeLocal + vbCrLf + "Ingrese el tipo de generacion de historia" '+ Chr(13)
   Else
       Select Case Val(mo_cmbIdTipoGenHistoriaClinica.BoundText)
           Case sghHistoriaTemporalCOnsultaExterna, sghHistoriaTemporalEmergencia, sghSinHistoria
           Case sghHistoriaDefinitivaManual
               If Trim(txtIdNroHistoria.Text) = "" Then
                    sMensajeLocal = sMensajeLocal + vbCrLf + "Ingrese el número de historia clínica"    '+ Chr(13)
                    If mb_FaltaDato = False Then
                       UserControl.txtIdNroHistoria.SetFocus
                       mb_FaltaDato = True
                    End If
                    txtIdNroHistoria.BackColor = ml_ColorError
               End If
           Case Else
       End Select
   End If
      
   If Not mb_PacienteNoIdentificado Then
'   'GLCC-Validad que campo DNI Tenga registros-20/07/2020
'   If Trim(txtNroDocumento.Text) = "" Then
'            sMensajeLocal = sMensajeLocal + vbCrLf + "Ingrese el N° de Documento" '+ Chr(13)
'            If mb_FaltaDato = False Then
'                UserControl.txtNroDocumento.SetFocus
'                mb_FaltaDato = True
'            End If
'            txtNroDocumento.BackColor = ml_ColorError
'        End If
        'Termina Modificación GLCC-20-07-2020
        If Trim(txtApellidoPaterno.Text) = "" Then
            sMensajeLocal = sMensajeLocal + vbCrLf + "Ingrese el Apellido Paterno" '+ Chr(13)
            If mb_FaltaDato = False Then
                UserControl.txtApellidoPaterno.SetFocus
                mb_FaltaDato = True
            End If
            txtApellidoPaterno.BackColor = ml_ColorError
        ElseIf mo_Teclado.TextoAlmenosExisteAlgunaLetra(txtApellidoPaterno.Text) = False And wxSinApellido <> txtApellidoPaterno.Text Then
            sMensajeLocal = sMensajeLocal + vbCrLf + "El Apellido Paterno NO TIENE LETRA" '+ Chr(13)
            UserControl.txtApellidoPaterno.SetFocus
        End If
        If Trim(txtApellidoMaterno.Text) = "" Then
            sMensajeLocal = sMensajeLocal + vbCrLf + "Ingrese el apellido Materno" '+ Chr(13)
            If mb_FaltaDato = False Then
                UserControl.txtApellidoMaterno.SetFocus
                mb_FaltaDato = True
            End If
            txtApellidoMaterno.BackColor = ml_ColorError
        ElseIf mo_Teclado.TextoAlmenosExisteAlgunaLetra(txtApellidoMaterno.Text) = False And wxSinApellido <> txtApellidoMaterno.Text Then
            sMensajeLocal = sMensajeLocal + vbCrLf + "El Apellido Materno NO TIENE LETRA" '+ Chr(13)
            UserControl.txtApellidoMaterno.SetFocus
        End If
        If Trim(txtPrimerNombre.Text) = "" Then
            sMensajeLocal = sMensajeLocal + vbCrLf + "Ingrese el Primer Nombre" '+ Chr(13)
            If mb_FaltaDato = False Then
                UserControl.txtPrimerNombre.SetFocus
                mb_FaltaDato = True
            End If
            txtPrimerNombre.BackColor = ml_ColorError
        ElseIf mo_Teclado.TextoAlmenosExisteAlgunaLetra(txtPrimerNombre.Text) = False Then
            sMensajeLocal = sMensajeLocal + vbCrLf + "El Primer Nombre NO TIENE LETRA" '+ Chr(13)
            UserControl.txtPrimerNombre.SetFocus
        End If
        
        If chkSinFechaNacimiento.Value = 1 Then
            If txtEdad.Text = "" Or mo_cmbIdTipoEdad.BoundText = "" Then
                sMensajeLocal = sMensajeLocal + vbCrLf + "Debe Ingresar una edad y un Tipo de Edad(Edad Actual)" '+ Chr(13)
                
            Else
                If Val(txtEdad.Text) = 0 Then
                    sMensajeLocal = sMensajeLocal + vbCrLf + "Edad ingresada debe ser mayor a cero" '+ Chr(13)
                End If
            End If
            
        End If
        
        If txtFechaNacimiento.Text = sighEntidades.FECHA_VACIA_DMY Then
            sMensajeLocal = sMensajeLocal + vbCrLf + "Debe registrar la FECHA DE NACIMIENTO" '+ Chr(13)
            If mb_FaltaDato = False Then
                If UserControl.txtFechaNacimiento.Enabled = True Then 'Actualizado 16092014
                    UserControl.txtFechaNacimiento.SetFocus
                    mb_FaltaDato = True
                End If
            End If
            txtFechaNacimiento.BackColor = ml_ColorError
        End If
        'Validación de Etnia
         'If Val(mo_cmbEtnia.BoundText) = "" Then
         'GLCC-Validar combo que no se encuntre Vacio -- 20-07/-2020
         If cmbEtnia.Text = "" Then
            sMensajeLocal = sMensajeLocal + vbCrLf + "Debe registrar la Etnia" '+ Chr(13)
            If mb_FaltaDato = False Then
                 UserControl.cmbEtnia.SetFocus
                mb_FaltaDato = False
            End If
            cmbEtnia.BackColor = ml_ColorError
        End If
        If txtHoraNacimiento.Text = sighEntidades.HORA_VACIA_HM Then
           txtHoraNacimiento.Text = "00:00"
        End If
        '
        If Val(mo_CmbIdTipoSexo.BoundText) = 0 Then
            sMensajeLocal = sMensajeLocal + vbCrLf + "Ingrese el sexo" '+ Chr(13)
            If mb_FaltaDato = False Then
                UserControl.cmbIdTipoSexo.SetFocus
                mb_FaltaDato = True
            End If
            cmbIdTipoSexo.BackColor = ml_ColorError
        End If
        '
        If lnOpcionQueUsaEsteControl <> 1 Then
            If cmbEtnia.ListIndex < 0 Then
               cmbEtnia.Text = ""
            End If
            If cmbEtnia.Text = "" Then
                sMensajeLocal = sMensajeLocal + vbCrLf + "Elija la ETNIA" '+ Chr(13)
                If mb_FaltaDato = False Then
                   UserControl.cmbEtnia.SetFocus
                   mb_FaltaDato = True
                End If
                cmbEtnia.BackColor = ml_ColorError
            End If
        End If
        '
        If lnOpcionQueUsaEsteControl <> 1 Then
            If cmbIdioma.ListIndex < 0 Then
               cmbIdioma.Text = ""
            End If
            If cmbIdioma.Text = "" Then
                sMensajeLocal = sMensajeLocal + vbCrLf + "Elija la IDIOMA" '+ Chr(13)
                If mb_FaltaDato = False Then
                   UserControl.cmbIdioma.SetFocus
                   mb_FaltaDato = True
                End If
                cmbIdioma.BackColor = ml_ColorError
            End If
        End If
        '
        If txtEmail.Text <> "" Then
            If mo_Cadena.DevuelveARROBAS(txtEmail.Text) <> 1 Then
               sMensajeLocal = sMensajeLocal + "Debe haber un    @    en el EMAIL" + Chr(13)
            ElseIf Len(txtEmail.Text) < 3 Then
               sMensajeLocal = sMensajeLocal + vbCrLf + "La longitud del Email no es adecuado" '+ Chr(13)
            End If
        End If
        If wxParametro282 = "S" And wxParametro333 = "S" Then  'solo para CS y que se exija el ingreso
                If Trim(txtSector.Text) = "" Then
                   sMensajeLocal = sMensajeLocal + vbCrLf + "Debe registrar el SECTOR (por ser un CS/PS)" '+ Chr(13)
                End If
                If Trim(lblSectorista.Caption) = "" Then
                   sMensajeLocal = sMensajeLocal + vbCrLf + "Elija el SECTORISTA (por ser un CS/PS)" '+ Chr(13)
                End If
        End If
    Else
        If Val(mo_CmbIdTipoSexo.BoundText) = 0 Then
            sMensajeLocal = sMensajeLocal + vbCrLf + "Ingrese el sexo" '+ Chr(13)
            If mb_FaltaDato = False Then
                UserControl.cmbIdTipoSexo.SetFocus
                mb_FaltaDato = True
            End If
            cmbIdTipoSexo.BackColor = ml_ColorError
        End If
    End If

   
   If txtFechaCreacion = sighEntidades.FECHA_VACIA_DMY Then
        sMensajeLocal = sMensajeLocal + vbCrLf + "Por favor ingrese la fecha de creación" '+ Chr(13)
   End If
   
   ValidarDatosObligatorios = sMensajeLocal


'Dim sMensajeLocal As String
'
'
'   If Not mb_PacienteNoIdentificado Then
'        If txtApellidoPaterno.Text = "" Then
'            sMensajeLocal = sMensajeLocal + "Ingrese el Apellido Paterno" + Chr(13)
'        ElseIf mo_Teclado.TextoAlmenosExisteAlgunaLetra(txtApellidoPaterno.Text) = False And wxSinApellido <> txtApellidoPaterno.Text Then
'            sMensajeLocal = sMensajeLocal + "El Apellido Paterno NO TIENE LETRA" + Chr(13)
'        End If
'        If txtApellidoMaterno.Text = "" Then
'            sMensajeLocal = sMensajeLocal + "Ingrese el apellido materno" + Chr(13)
'        ElseIf mo_Teclado.TextoAlmenosExisteAlgunaLetra(txtApellidoMaterno.Text) = False And wxSinApellido <> txtApellidoMaterno.Text Then
'            sMensajeLocal = sMensajeLocal + "El Apellido Materno NO TIENE LETRA" + Chr(13)
'        End If
'        If txtPrimerNombre.Text = "" Then
'            sMensajeLocal = sMensajeLocal + "Ingrese el primer nombre" + Chr(13)
'        ElseIf mo_Teclado.TextoAlmenosExisteAlgunaLetra(txtPrimerNombre.Text) = False Then
'            sMensajeLocal = sMensajeLocal + "El Primer Nombre NO TIENE LETRA" + Chr(13)
'        End If
'
'        If chkSinFechaNacimiento.Value = 1 Then
'            If txtEdad.Text = "" Or mo_cmbIdTipoEdad.BoundText = "" Then
'                sMensajeLocal = sMensajeLocal + "Debe Ingresar una edad y un Tipo de Edad(Edad Actual)" + Chr(13)
'            Else
'                If Val(txtEdad.Text) = 0 Then
'                    sMensajeLocal = sMensajeLocal + "Edad ingresada debe ser mayor a cero" + Chr(13)
'                End If
'            End If
'
'        End If
'
'        If txtFechaNacimiento.Text = sighEntidades.FECHA_VACIA_DMY Then
'           sMensajeLocal = sMensajeLocal + "Debe registrar la FECHA DE NACIMIENTO" + Chr(13)
'        End If
'        If txtHoraNacimiento.Text = sighEntidades.HORA_VACIA_HM Then
'           txtHoraNacimiento.Text = "00:00"
'        End If
'        '
'        If cmbEtnia.ListIndex < 0 Then
'           cmbEtnia.Text = ""
'        End If
'        If cmbEtnia.Text = "" Then
'            sMensajeLocal = sMensajeLocal + "Elija la ETNIA" + Chr(13)
'        End If
'        '
'        If cmbIdioma.ListIndex < 0 Then
'           cmbIdioma.Text = ""
'        End If
'        If cmbIdioma.Text = "" Then
'            sMensajeLocal = sMensajeLocal + "Elija la IDIOMA" + Chr(13)
'        End If
'        '
'        If txtEmail.Text <> "" Then
'            If mo_Cadena.DevuelveARROBAS(txtEmail.Text) <> 1 Then
'               sMensajeLocal = sMensajeLocal + "Debe haber un    @    en el EMAIL" + Chr(13)
'            ElseIf Len(txtEmail.Text) < 3 Then
'               sMensajeLocal = sMensajeLocal + "La longitud del Email no es adecuado" + Chr(13)
'            End If
'        End If
'        If wxParametro282 = "S" And wxParametro333 = "S" Then  'solo para CS y que se exija el ingreso
'                If Trim(txtSector.Text) = "" Then
'                   sMensajeLocal = sMensajeLocal + "Debe registrar el SECTOR (por ser un CS/PS)" + Chr(13)
'                End If
'                If Trim(lblSectorista.Caption) = "" Then
'                   sMensajeLocal = sMensajeLocal + "Elija el SECTORISTA (por ser un CS/PS)" + Chr(13)
'                End If
'        End If
'    End If
'
'   If Val(mo_CmbIdTipoSexo.BoundText) = 0 Then
'       sMensajeLocal = sMensajeLocal + "Ingrese el sexo" + Chr(13)
'   End If
'
'   If mo_cmbIdTipoGenHistoriaClinica.BoundText = "" Then
'       sMensajeLocal = sMensajeLocal + "Ingrese el tipo de generacion de historia" + Chr(13)
'   Else
'        Select Case Val(mo_cmbIdTipoGenHistoriaClinica.BoundText)
'            Case sghHistoriaTemporalCOnsultaExterna, sghHistoriaTemporalEmergencia, sghSinHistoria
'
'            Case sghHistoriaDefinitivaManual
'                If txtIdNroHistoria.Text = "" Then
'                    sMensajeLocal = sMensajeLocal + "Ingrese el número de historia clínica" + Chr(13)
'                End If
'            Case Else
'
'        End Select
'   End If
'
'   If txtFechaCreacion = sighEntidades.FECHA_VACIA_DMY Then
'        sMensajeLocal = sMensajeLocal + "Por favor ingrese la fecha de creación" + Chr(13)
'   End If
'
'   ValidarDatosObligatorios = sMensajeLocal

End Function

Public Function ValidarReglas(oDOPaciente As doPaciente) As Boolean
Dim rspacientes As ADODB.Recordset

    ValidarReglas = False
    
    'Si el paciente aun no existe (IdPaciente = 0) se verifica que no haya duplicados
    If oDOPaciente.idPaciente = 0 Then
         Set rspacientes = mo_AdminAdmision.PacientesObtenerConElAutogenerado(oDOPaciente)
         If Not (rspacientes.EOF And rspacientes.BOF) Then
             rspacientes.MoveFirst
             If UserControl.chkNN.Value = 0 Then
                MsgBox "Existe un paciente con el mismo número autogenerado (HC: " & _
                        HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rspacientes.Fields!NroHistoriaClinica)), False) & _
                         ")", vbExclamation, "Datos de paciente"
                rspacientes.Close
                Set rspacientes = Nothing
                Exit Function
             Else
                If MsgBox("Existe un paciente con el mismo número autogenerado: " + rspacientes!ApellidoPaterno + " " + rspacientes!ApellidoMaterno + " " + rspacientes!PrimerNombre + "  HC: " + _
                        HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rspacientes.Fields!NroHistoriaClinica)), False) + _
                        Chr(13) + "Desea continuar?", vbQuestion + vbYesNo, "Datos de paciente") = vbNo Then
                    rspacientes.Close
                    Set rspacientes = Nothing
                    Exit Function
                End If
             End If
         End If
         rspacientes.Close
        
         Select Case oDOPaciente.idTipoNumeracion
         Case sghHistoriaDefinitivaManual    ', sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
             Set rspacientes = mo_AdminAdmision.PacientesObtenerConElMismoNroHistoriaDefinitiva(oDOPaciente)
             If Not (rspacientes.EOF And rspacientes.BOF) Then
                 rspacientes.MoveFirst
                 MsgBox "Existe un paciente con el mismo número de historia clínica: " + rspacientes!ApellidoPaterno + " " + rspacientes!ApellidoMaterno + " " + rspacientes!PrimerNombre, vbInformation, "Datos de paciente"
                 rspacientes.Close
                 Set rspacientes = Nothing
                 Exit Function
             End If
             rspacientes.Close
         End Select
         
       If mo_cmbIdDocIdentidad.BoundText <> "" And txtNroDocumento.Text <> "" Then
            Set rspacientes = mo_AdminAdmision.PacientesFiltraPorNroDocumentoYtipo(txtNroDocumento.Text, Val(mo_cmbIdDocIdentidad.BoundText))
            If rspacientes.RecordCount > 0 Then
                 rspacientes.MoveFirst
                 'Actualizado 20092014
                 MsgBox "El nro de documento: " & txtNroDocumento.Text & ", ya existe para el Paciente: " + _
                        HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rspacientes!NroHistoriaClinica)), False) + _
                        " " + rspacientes!ApellidoPaterno + " " + rspacientes!ApellidoMaterno + " " + rspacientes!PrimerNombre, vbInformation, "Datos de paciente"
                 rspacientes.Close
                 Set rspacientes = Nothing
                 On Error Resume Next
                 txtNroDocumento.Text = ""
                 Exit Function
            End If
            rspacientes.Close
         End If
         '
         If mi_Opcion = sghAgregar And txtFichaFamiliar3.Visible = True And txtFichaFamiliar1.Text <> "" And txtFichaFamiliar2.Text <> "" And txtFichaFamiliar3.Text <> "" Then
            ms_MensajeError = mo_AdminAdmision.ExisteFichaFamiliar(DevuelveFichaFamiliarUnida, ml_IdPaciente)
            If ms_MensajeError <> "" Then
                MsgBox "Existe un paciente con la misma FICHA FAMILIAR: " + Chr(13) + ms_MensajeError
                'mgaray20141008
                On Error Resume Next
                txtFichaFamiliar3.Text = ""
                txtFichaFamiliar3.SetFocus
            End If
         End If
    Else
         If mo_cmbIdDocIdentidad.BoundText <> "" And txtNroDocumento.Text <> "" Then
            Set rspacientes = mo_AdminAdmision.PacientesFiltraPorNroDocumentoYtipo(txtNroDocumento.Text, Val(mo_cmbIdDocIdentidad.BoundText))
            If rspacientes.RecordCount > 0 Then
                 rspacientes.MoveFirst
                 Do While Not rspacientes.EOF
                    If rspacientes.Fields!idPaciente <> oDOPaciente.idPaciente Then
                        'debb-hra-ya en version Polsalud
                        If (Not IsNull(rspacientes!NroHistoriaClinica)) Then
                           MsgBox "Es N°DOCUMENTO ya existe para el Paciente: " + _
                                    HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rspacientes!NroHistoriaClinica)), False) + " " + rspacientes!ApellidoPaterno + " " + rspacientes!ApellidoMaterno + " " + rspacientes!PrimerNombre, vbInformation, "Datos de paciente"
                        Else
                           MsgBox "Es N°DOCUMENTO ya existe para el Paciente: " + rspacientes!ApellidoPaterno + " " + rspacientes!ApellidoMaterno + " " + rspacientes!PrimerNombre, vbInformation, "Datos de paciente"
                        End If
                        rspacientes.Close
                        Set rspacientes = Nothing
                        Exit Function
                    End If
                    rspacientes.MoveNext
                 Loop
            End If
            rspacientes.Close
         End If
    End If
    '
    'mgaray20141008
    On Error Resume Next
    If txtFechaNacimiento.Text <> sighEntidades.FECHA_VACIA_DMY And mi_Opcion <> sghEliminar Then
        If CDate(txtFechaNacimiento.Text) > CDate(txtFechaCreacion) Then
            MsgBox "La fecha de nacimiento no puede ser mayor que la fecha de creación de la historia", vbExclamation, "Registro de pacientes"
            Exit Function
        End If
        If CDate(txtFechaNacimiento.Text) > Format(ldHoy, sighEntidades.DevuelveFechaSoloFormato_DMY) Then
                MsgBox "La fecha de nacimiento no puede ser mayor a HOY", vbExclamation, "Registro de pacientes"
                Exit Function
        End If
    End If
    '
    If lbExigeIngresoDelDNI = True And mi_Opcion <> sghEliminar Then
       If cmbIdDocIdentidad.Text = "" Or txtNroDocumento.Text = "" Then
            MsgBox "Es obligatorio ingresar el DNI", vbExclamation, "Registro de pacientes"
            Exit Function
       End If
    End If
    '
    If cmbIdDocIdentidad.Locked = False And mo_cmbIdDocIdentidad.BoundText = "1" And Len(txtNroDocumento.Text) <> 8 And mi_Opcion <> sghEliminar Then
        MsgBox "DNI debe tener 8 dígitos", vbExclamation, "Registro de pacientes"
        Exit Function
    End If
    '
    If lbExigeIngresoDeCentroPoblado = True And mi_Opcion <> sghEliminar Then
       If cmbIdCentroPobladoDomicilio.Text = "" Then
            MsgBox "Es obligatorio elegir el CENTRO POBLADO para CS", vbExclamation, "Registro de pacientes"
            Exit Function
       End If
    End If
    '
    If Val(txtEdad.Text) < 18 And mo_cmbIdDocIdentidad.BoundText <> "10" And (txtNroDocumento.Text = "" Or mo_cmbIdDocIdentidad.BoundText = "8" Or mo_cmbIdDocIdentidad.BoundText = "9") And mi_Opcion <> sghEliminar And mb_PacienteNoIdentificado = False Then
       If txtMadreDocumento.Text = "" Then
          If Val(txtNroHijo.Text) = 0 Then
             MsgBox "El Paciente es MENOR DE EDAD, por favor debe registrar el N°HIJO y el DNI DE LA MADRE" & _
                     Chr(13) & "Si no tiene MADRE o TUTOR elegir en TIPO DOCUMENTO del PACIENTE= 10(Sin registro madre/tutor)", vbInformation, "Registro de pacientes"
             txtNroHijo.SetFocus
             Exit Function
          End If
          If txtMadreApellidoP.Text = "" Or txtMadreApellidoM.Text = "" Or txtNombreMadre.Text = "" Then
             MsgBox "El Paciente es MENOR DE EDAD, por favor debe registrar El N° DNI de la MADRE o los APELLIDOS Y NOMBRES DE LA MADRE" & _
                     Chr(13) & "Si no tiene MADRE o TUTOR elegir en TIPO DOCUMENTO del PACIENTE= 10(Sin registro madre/tutor)", vbInformation, "Registro de pacientes"
             txtMadreApellidoP.SetFocus
             Exit Function
          End If
       ElseIf Len(txtMadreDocumento.Text) <> 8 And mo_cmbMadreTipoDocumento.BoundText = "1" Then
          MsgBox "El N° DNI de la MADRE tiene longitud diferente a OCHO", vbInformation, "Registro de pacientes"
          txtMadreDocumento.SetFocus
          Exit Function
       ElseIf Val(txtNroHijo.Text) = 0 Then
             MsgBox "El Paciente es MENOR DE EDAD, por favor debe registrar el N°HIJO" & _
                     Chr(13) & "Si no tiene MADRE o TUTOR elegir en TIPO DOCUMENTO del PACIENTE= 10(Sin registro madre/tutor)", vbInformation, "Registro de pacientes"
             txtNroHijo.SetFocus
             Exit Function
       End If
    End If
    '
    If txtApellidoMaterno.Text = wxSinApellido Or txtApellidoMaterno.Text = wxSinApellido Then
       If Len(txtNroDocumento.Text) <> 8 And mo_cmbIdDocIdentidad.BoundText = "1" Then
            MsgBox "Debe registrar el DNI para que el Paciente tenga un sólo apellido", vbInformation, ""
            Exit Function
       End If
    End If
    'mgaray201503
    Dim sMessageError As String
    
    If ValidarNumeroDeHistoriaClinica(sMessageError) = False Then
        MsgBox sMessageError, vbInformation, ""
        Exit Function
    End If
    '
    Set rspacientes = Nothing
    ValidarReglas = True

End Function

Sub ActualizaTipoYnroDocumentoDelPaciente(doPacientes As doPaciente)
    If txtNroDocumento.Text = "" Or mo_cmbIdDocIdentidad.BoundText = "8" Or mo_cmbIdDocIdentidad.BoundText = "9" Then
       If Val(txtNroHijo.Text) > 0 And Len(txtMadreDocumento.Text) > 0 Then
            mo_cmbIdDocIdentidad.BoundText = "8"
            txtNroDocumento.MaxLength = 12
            txtNroDocumento.Text = Trim(mo_cmbMadreTipoDocumento.BoundText) & Right("0000" & Trim(txtMadreDocumento.Text), 8) & Right("0" & Trim(txtNroHijo.Text), 2)
            
       ElseIf txtMadreApellidoP.Text <> "" And txtMadreApellidoM.Text <> "" And txtNombreMadre.Text <> "" Then
            Dim P1 As String    'Primer digito del apellido paterno
            Dim P4 As String    'Cuarto Digito del apellido paterno
            Dim M1 As String    'Primer digito del apellido materno
            Dim M4 As String    'Cuarto digito del apellido materno
            Dim N11 As String   'Primer digito del primer nombre
            Dim N41 As String   'Cuarto digito del primer materno
            Dim N12 As String   'Primer digito del Ultimo materno
            Dim N42 As String   'Cuarto digito del Ultimo materno
            mo_AdminAdmision.DevuelvePrimeryCuartoCaracter txtMadreApellidoP.Text, P1, P4
            mo_AdminAdmision.DevuelvePrimeryCuartoCaracter txtMadreApellidoM, M1, M4
            mo_AdminAdmision.DevuelvePrimeryCuartoCaracter txtNombreMadre.Text, N11, N41
            mo_AdminAdmision.DevuelvePrimeryCuartoCaracter txtMadreSnombre.Text, N12, N42
            mo_cmbIdDocIdentidad.BoundText = "9"
            txtNroDocumento.MaxLength = 12
            txtNroDocumento.Text = P1 + P4 + M1 + M4 + N11 + N41 + N12 + N42 + Trim(txtNroHijo.Text)
            
       End If
    End If
    doPacientes.nrodocumento = txtNroDocumento.Text
    doPacientes.IdDocIdentidad = Val(mo_cmbIdDocIdentidad.BoundText)
End Sub


Public Function CargarDatosAlObjetoDatos(oDOPaciente As doPaciente, oDOHistoria As DOHistoriaClinica, _
        Optional mo_DoPacientesDatosAdd As DoPacienteDatosAdd = Nothing)
    If mo_DoPacientesDatosAdd Is Nothing Then
        Set mo_DoPacientesDatosAdd = New DoPacienteDatosAdd
    End If
    With mo_DoPacientesDatosAdd
        mo_DoPacientesDatosAdd.idPaciente = Me.idPaciente
        mo_DoPacientesDatosAdd.FNacimientoCalculada = IIf(chkSinFechaNacimiento.Value = 1, True, False)
    End With
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DEL PACIENTE
    '---------------------------------------------------------------------------------
   With oDOPaciente
        .idPaciente = Me.idPaciente
        .ApellidoPaterno = txtApellidoPaterno.Text
        .ApellidoMaterno = txtApellidoMaterno.Text
        .PrimerNombre = txtPrimerNombre.Text
        .SegundoNombre = txtSegundoNombre.Text
        .TercerNombre = txtTercerNombre.Text
        If txtFechaNacimiento.Text = sighEntidades.FECHA_VACIA_DMY Then
           .FechaNacimiento = 0
        Else
           If txtHoraNacimiento.Text = sighEntidades.HORA_VACIA_HM Then
              txtHoraNacimiento.Text = "00:00"
           End If
           .FechaNacimiento = CDate(txtFechaNacimiento.Text & " " & txtHoraNacimiento.Text)
        End If
        .nrodocumento = txtNroDocumento.Text
        .TELEFONO = txtTelefono.Text
        .NroHistoriaClinica = txtIdNroHistoria.Tag
        .DireccionDomicilio = txtDireccionDomicilio.Text
        .idTipoSexo = Val(mo_CmbIdTipoSexo.BoundText)
        '
        '.IdProcedencia = Val(mo_cmbIdProcedencia.BoundText)
        If cmbIdProcedencia.ListIndex < 0 Then
            .IdProcedencia = 0
        Else
            oCampos = Split(cmbIdProcedencia.List(cmbIdProcedencia.ListIndex), "|")
            .IdProcedencia = Val(oCampos(0))
        End If
        '
        '.IdGradoInstruccion = Val(mo_cmbIdGradoInstruccion.BoundText)
        If cmbIdGradoInstruccion.ListIndex < 0 Then
            .IdGradoInstruccion = 0
        Else
            oCampos = Split(cmbIdGradoInstruccion.List(cmbIdGradoInstruccion.ListIndex), "|")
            .IdGradoInstruccion = Val(oCampos(0))
        End If
        '
        '.IdEstadoCivil = Val(mo_cmbIdEstadoCivil.BoundText)
        If cmbIdEstadoCivil.ListIndex < 0 Then
            .IdEstadoCivil = 0
        Else
            oCampos = Split(cmbIdEstadoCivil.List(cmbIdEstadoCivil.ListIndex), "|")
            .IdEstadoCivil = Val(oCampos(0))
        End If
        '
        .IdDocIdentidad = Val(mo_cmbIdDocIdentidad.BoundText)
        '
        '.idTipoOcupacion = Val(mo_cmbIdTipoOcupacion.BoundText)
        If cmbIdTipoOcupacion.ListIndex < 0 Then
            .idTipoOcupacion = 0
        Else
            oCampos = Split(cmbIdTipoOcupacion.List(cmbIdTipoOcupacion.ListIndex), "|")
            .idTipoOcupacion = Val(oCampos(0))
        End If
        '
        .IdPaisNacimiento = Val(mo_cmbIdPaisNacimiento.BoundText)
        .IdDistritoNacimiento = Val(mo_cmbIdDistritoNacimiento.BoundText)
        .IdCentroPobladoNacimiento = Val(mo_cmbIdCentroPobladoNacimiento.BoundText)
        
         .IdPaisDomicilio = Val(mo_cmbIdPaisDomicilio.BoundText)
        .IdDistritoDomicilio = Val(mo_cmbIdDistritoDomicilio.BoundText)
        .IdCentroPobladoDomicilio = Val(mo_cmbIdCentroPobladoDomicilio.BoundText)
        
        .IdPaisProcedencia = Val(mo_cmbIdPaisProcedencia.BoundText)
'        .IdDepartamentoProcedencia = Val(mo_cmbIdDepartamentoProcedencia.BoundText)
'        .IdProvinciaProcedencia = Val(mo_cmbIdProvinciaProcedencia.BoundText)
        .IdDistritoProcedencia = Val(mo_cmbIdDistritoProcedencia.BoundText)
        .IdCentroPobladoProcedencia = Val(mo_cmbIdCentroPobladoProcedencia.BoundText)
        
'        .EtapaDomicilio = txtEtapaDomicilio.Text
'        .SectorDomicilio = txtSectorDomicilio.Text
'        .LoteDomicilio = txtLoteDomicilio.Text
'        .ManzanaDomicilio = txtManzanaDomicilio.Text
'        .PisoDomicilio = txtPisoDomicilio.Text
'        .NroDomicilio = txtNroDomicilio.Text
        
         .NombrePadre = txtNombrePadre.Text
         .Nombremadre = txtNombreMadre.Text
         .idTipoNumeracion = mo_cmbIdTipoGenHistoriaClinica.BoundText
        
        .Autogenerado = mo_AdminAdmision.PacienteCrearNroAutogenerado(oDOPaciente)
         Autogenerado = .Autogenerado
         .IdUsuarioAuditoria = Me.idUsuario
         .Observacion = txtObservacion
         If Len(Trim(txtFichaFamiliar1.Text)) > 0 And Len(Trim(txtFichaFamiliar2.Text)) > 0 And Len(Trim(txtFichaFamiliar3.Text)) > 0 Then
            .FichaFamiliar = txtFichaFamiliar1.Text & "-" & txtFichaFamiliar2.Text & "-" & txtFichaFamiliar3.Text
         Else
            .FichaFamiliar = ""
         End If
         '
         '.IdEtnia = Right("0" & mo_cmbEtnia.BoundText, 2)
         If cmbEtnia.ListIndex < 0 Then
            .IdEtnia = ""
         Else
            oCampos = Split(cmbEtnia.List(cmbEtnia.ListIndex), "|")
            .IdEtnia = Right("0" & oCampos(0), 2)
         End If
         '
         '.IdIdioma = Val(mo_cmbIdioma.BoundText)
         If cmbIdioma.ListIndex < 0 Then
            .IdIdioma = 0
         Else
            oCampos = Split(cmbIdioma.List(cmbIdioma.ListIndex), "|")
            .IdIdioma = Val(oCampos(0))
         End If
         
         '
'<(Inicio) Añadido Por: WABG el: 23/10/2020-07:48:42 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
         .validacionReniec = mb_validacionReniec
'</(Fin) Añadido Por: WABG el: 23/10/2020-07:48:42 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>

         .UsoWebReniec = mb_UsoWebReniec
         .Email = UserControl.txtEmail.Text
         .NroOrdenHijo = Val(txtNroHijo.Text)
         .madreTipoDocumento = Val(mo_cmbMadreTipoDocumento.BoundText)
         .madreDocumento = txtMadreDocumento.Text
         .madreApellidoPaterno = txtMadreApellidoP.Text
         .madreApellidoMaterno = txtMadreApellidoM.Text
         .madrePrimerNombre = txtNombreMadre.Text
         .madreSegundoNombre = txtMadreSnombre.Text
         '
         If mi_Opcion <> sghAgregar Then
            .UsoWebReniec = lbUsoWebReniec_SinMostrar
         End If
        .FactorRh = lcFactorRh_SinMostrar
        .GrupoSanguineo = lcGrupoSanguineo_SinMostrar
         If fraSector.Enabled = True Then
            .sector = txtSector.Text
            .sectorista = Val(txtSectorista.Text)
         End If
   End With
   ActualizaTipoYnroDocumentoDelPaciente oDOPaciente

    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA HISTORIA CLINICA
    '---------------------------------------------------------------------------------
    With oDOHistoria
        .NroHistoriaClinica = txtIdNroHistoria.Tag
        .fechacreacion = IIf(txtFechaCreacion = sighEntidades.FECHA_VACIA_DMY Or txtFechaCreacion = "", 0, txtFechaCreacion)
        .idTipoNumeracion = Val(mo_cmbIdTipoGenHistoriaClinica.BoundText)
        .FechaPasoAPasivo = 0
        .IdEstadoHistoria = 1
        .idPaciente = Me.idPaciente
        .idTipoHistoria = 1
        .IdUsuarioAuditoria = Me.idUsuario
    End With

    Set CargarDatosAlObjetoDatos = oDOPaciente
    
    
    
End Function

Public Sub CargarDatosDePacienteALosControles(oConexion1 As Connection, lcParametro242 As String, lcParametro287 As String)
On Error GoTo ErrrCargaDatos
Dim oPacientes  As New doPaciente
Dim lcRutaImg As String
Dim lcUbigeoDistrito As String
Dim oConexion As New Connection
        If oConexion1 Is Nothing Then
            oConexion.CommandTimeout = 300
            oConexion.CursorLocation = adUseClient
            oConexion.Open sighEntidades.CadenaConexion
        End If
        
        'CARGAR DATOS DEL PACIENTE
        If oConexion1 Is Nothing Then
           Set oPacientes = mo_AdminAdmision.PacientesSeleccionarPorId(ml_IdPaciente, oConexion)
        Else
           Set oPacientes = mo_AdminAdmision.PacientesSeleccionarPorId(ml_IdPaciente, oConexion1)
        End If
        'Frank 28 01 2015
        If oPacientes.SegundoNombre = "" Then
            If lblSegundoNombrePacienteSIS <> "" Then
                oPacientes.SegundoNombre = lblSegundoNombrePacienteSIS
            End If
        End If
        
        If mo_AdminAdmision.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbInformation, "Datos de paciente"
             mb_ExistenDatos = False
             Exit Sub
        End If
        If Not oPacientes Is Nothing Then
            CargaDatosPersonales oPacientes, lcParametro242, wxParametro237
        Else
            mb_ExistenDatos = False
            Exit Sub
        End If
        If oConexion1 Is Nothing Then
           oConexion.Close
        End If
        
'<(Inicio) Añadido Por: WABG el: 23/10/2020-08:01:40 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
        If oPacientes.validacionReniec = True Then
        
        deshabilitarControlesRENIECModificarPacienteValidado
        
        End If
'</(Fin) Añadido Por: WABG el: 23/10/2020-08:01:40 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
        
        Set oConexion = Nothing
ErrrCargaDatos:
End Sub
Public Sub ReemplazarDatosDeUbicacion(oUbicacionPaciente As doPaciente)
                
        With oUbicacionPaciente
            mo_cmbIdPaisDomicilio.BoundText = .IdPaisDomicilio
'            mo_cmbIdDepartamentoDomicilio.BoundText = .IdDepartamentoDomicilio
'            mo_cmbIdProvinciaDomicilio.BoundText = .IdProvinciaDomicilio
'            mo_cmbIdDistritoDomicilio.BoundText = .IdDistritoDomicilio
            mo_cmbIdCentroPobladoDomicilio.BoundText = .IdCentroPobladoDomicilio
            
            mo_cmbIdPaisProcedencia.BoundText = .IdPaisProcedencia
'            mo_cmbIdDepartamentoProcedencia.BoundText = .IdDepartamentoProcedencia
'            mo_cmbIdProvinciaProcedencia.BoundText = .IdProvinciaProcedencia
'            mo_cmbIdDistritoProcedencia.BoundText = .IdDistritoProcedencia
            mo_cmbIdCentroPobladoProcedencia.BoundText = .IdCentroPobladoProcedencia
            
            txtDireccionDomicilio.Text = Trim(.DireccionDomicilio)
'            txtEtapaDomicilio.Text = Trim(.EtapaDomicilio)
'            txtSectorDomicilio.Text = Trim(.SectorDomicilio)
'            txtLoteDomicilio.Text = Trim(.LoteDomicilio)
'            txtManzanaDomicilio.Text = Trim(.ManzanaDomicilio)
'            txtPisoDomicilio.Text = Trim(.PisoDomicilio)
'            txtNroDomicilio.Text = Trim(.NroDomicilio)
            
        End With
End Sub
Public Sub ConfigurarPacienteNuevoONoIdentificado(bNoIdentificado As Integer)
    
    txtApellidoPaterno = IIf(bNoIdentificado = 1, "NN", "")
    txtApellidoMaterno = IIf(bNoIdentificado = 1, "NN", "")
    txtPrimerNombre = IIf(bNoIdentificado = 1, "NN", "")
    txtSegundoNombre = IIf(bNoIdentificado = 1, "NN", "")
   
    mo_Formulario.HabilitarDeshabilitar txtApellidoPaterno, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar txtApellidoMaterno, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar txtPrimerNombre, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar txtSegundoNombre, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar txtTercerNombre, Not (bNoIdentificado = 1)
    'mo_Formulario.HabilitarDeshabilitar txtFechaNacimiento, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar txtNroDocumento, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar txtTelefono, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar txtDireccionDomicilio, Not (bNoIdentificado = 1)
    'mo_Formulario.HabilitarDeshabilitar cmbIdTipoSexo, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdProcedencia, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdGradoInstruccion, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdEstadoCivil, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdDocIdentidad, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoOcupacion, Not (bNoIdentificado = 1)
    
    mo_Formulario.HabilitarDeshabilitar cmbIdPaisNacimiento, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdDepartamentoNacimiento, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdProvinciaNacimiento, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdDistritoNacimiento, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdCentroPobladoNacimiento, Not (bNoIdentificado = 1)
    
    mo_Formulario.HabilitarDeshabilitar cmbIdPaisDomicilio, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdDepartamentoDomicilio, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdProvinciaDomicilio, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdDistritoDomicilio, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdCentroPobladoDomicilio, Not (bNoIdentificado = 1)
    
    mo_Formulario.HabilitarDeshabilitar cmbIdPaisProcedencia, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdDepartamentoProcedencia, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdProvinciaProcedencia, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdDistritoProcedencia, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdCentroPobladoProcedencia, Not (bNoIdentificado = 1)
    
    'mo_Formulario.HabilitarDeshabilitar txtNombrePadre, Not (bNoIdentificado = 1)
    'mo_Formulario.HabilitarDeshabilitar txtNombreMadre, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbEtnia, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar cmbIdioma, Not (bNoIdentificado = 1)
    mo_Formulario.HabilitarDeshabilitar txtEmail, Not (bNoIdentificado = 1)
'    mo_Formulario.HabilitarDeshabilitar txtNroDomicilio, Not (bNoIdentificado = 1)
'    mo_Formulario.HabilitarDeshabilitar txtManzanaDomicilio, Not (bNoIdentificado = 1)
'    mo_Formulario.HabilitarDeshabilitar txtLoteDomicilio, Not (bNoIdentificado = 1)
'    mo_Formulario.HabilitarDeshabilitar txtPisoDomicilio, Not (bNoIdentificado = 1)
'    mo_Formulario.HabilitarDeshabilitar txtSectorDomicilio, Not (bNoIdentificado = 1)
'    mo_Formulario.HabilitarDeshabilitar txtEtapaDomicilio, Not (bNoIdentificado = 1)
    

End Sub
Public Sub ConfigurarValoresPorDefecto()

    mo_cmbIdPaisDomicilio.BoundText = 166   'Peru
    mo_cmbIdPaisNacimiento.BoundText = 166   'Peru
    mo_cmbIdPaisProcedencia.BoundText = 166   'Peru
    mo_cmbIdDocIdentidad.BoundText = 1   'dNI
    
    mo_Formulario.HabilitarDeshabilitar txtIdPaciente, False
    
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, True
    mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, True
    
    mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
    mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
    
    Select Case ml_TipoServicio
    Case sghConsultaExterna
        Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarDeConsultaExterna()
        If sighEntidades.GenerarHistoriaClinicaSiempre Then
            mo_cmbIdTipoGenHistoriaClinica.BoundText = sghHistoriaDefinitivaAutomatica
        Else
            mo_cmbIdTipoGenHistoriaClinica.BoundText = sghHistoriaTemporalCOnsultaExterna
        End If
        mo_cmbIdTipoGenHistoriaClinica.BoundText = lcBuscaParametro.SeleccionaFilaParametro(211)
    Case sghHospitalizacion
        Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarDeHospitalizacion()
        mo_cmbIdTipoGenHistoriaClinica.BoundText = lcBuscaParametro.SeleccionaFilaParametro(212)
    Case sghEmergenciaConsultorios, sghEmergenciaObservacion
        Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarDeEmergencia()
        mo_cmbIdTipoGenHistoriaClinica.BoundText = lcBuscaParametro.SeleccionaFilaParametro(210)
    End Select
    

    Dim rsTiposGeneracion As Recordset
    Set rsTiposGeneracion = mo_cmbIdTipoGenHistoriaClinica.RowSource
    
    If rsTiposGeneracion.RecordCount = 1 Then
        rsTiposGeneracion.MoveFirst
        mo_cmbIdTipoGenHistoriaClinica.BoundText = rsTiposGeneracion!idTipoNumeracion
    End If
    HabilitaFechaCreacion
    lcIdTipoGenHistoriaClinicaActual = mo_cmbIdTipoGenHistoriaClinica.BoundText
End Sub

Public Sub LimpiarDatosDePaciente(lcParametro211 As String, ldFechaActual As Date)
           
           'LIMPIAR DATOS DEL PACIENTE
           chkNN.Value = 0
           idPaciente = 0
           Autogenerado = 0
           txtIdPaciente = ""
           txtApellidoPaterno.Text = ""
           txtApellidoMaterno.Text = ""
           txtPrimerNombre.Text = ""
           txtSegundoNombre.Text = ""
           txtTercerNombre.Text = ""
           txtFechaNacimiento.Text = sighEntidades.FECHA_VACIA_DMY
           txtNroDocumento.Text = ""
           txtTelefono.Text = ""
           txtEdad.Text = ""
           lblTipoEdad.Caption = ""
           txtIdNroHistoria.Text = ""
           
           txtDireccionDomicilio.Text = ""
           mo_CmbIdTipoSexo.BoundText = ""
           cmbIdProcedencia.Text = ""      ' mo_cmbIdProcedencia.BoundText = ""
           cmbIdGradoInstruccion.Text = "" '  mo_cmbIdGradoInstruccion.BoundText = ""
           cmbIdEstadoCivil.Text = ""      'mo_cmbIdEstadoCivil.BoundText = ""
           mo_cmbIdDocIdentidad.BoundText = "1"
           txtNroDocumento.MaxLength = 8
           cmbIdTipoOcupacion.Text = ""    ' mo_cmbIdTipoOcupacion.BoundText = ""
           
            mo_cmbIdPaisNacimiento.BoundText = ""
           mo_cmbIdDepartamentoNacimiento.BoundText = ""
           mo_cmbIdProvinciaNacimiento.BoundText = ""
           mo_cmbIdDistritoNacimiento.BoundText = ""
           mo_cmbIdCentroPobladoNacimiento.BoundText = ""
           
           mo_cmbIdPaisDomicilio.BoundText = ""
           mo_cmbIdDepartamentoDomicilio.BoundText = ""
           mo_cmbIdProvinciaDomicilio.BoundText = ""
           mo_cmbIdDistritoDomicilio.BoundText = ""
           mo_cmbIdCentroPobladoDomicilio.BoundText = ""
           
           mo_cmbIdPaisProcedencia.BoundText = ""
           mo_cmbIdDepartamentoProcedencia.BoundText = ""
           mo_cmbIdProvinciaProcedencia.BoundText = ""
           mo_cmbIdDistritoProcedencia.BoundText = ""
           mo_cmbIdCentroPobladoProcedencia.BoundText = ""
           
           txtNombrePadre.Text = ""
           txtNombreMadre.Text = ""

'            txtNroDomicilio.Text = ""
'            txtManzanaDomicilio.Text = ""
'            txtLoteDomicilio.Text = ""
'            txtPisoDomicilio.Text = ""
'            txtSectorDomicilio.Text = ""
'            txtEtapaDomicilio.Text = ""

            mo_cmbIdTipoGenHistoriaClinica.BoundText = lcParametro211
            txtIdNroHistoria.Text = ""
            txtIdNroHistoria.Tag = 0
           txtFechaCreacion.Text = ldFechaActual

            mo_cmbIdPaisDomicilio.BoundText = 166   'Peru
            mo_cmbIdPaisNacimiento.BoundText = 166   'Peru
            mo_cmbIdPaisProcedencia.BoundText = 166   'Peru

            txtObservacion.Text = ""
            '
            txtFichaFamiliar1.Text = ""
            txtFichaFamiliar2.Text = ""
            txtFichaFamiliar3.Text = ""
            lblFichaFamiliar1.Caption = ""
            '
            chkIgualQueDomicilio.Value = 0
            chkIgualUQueDomicilioNac.Value = 0
            '
            If Val(lcEtniaDefault) > 0 Then
               cmbEtnia_UbicaPosicion lcEtniaDefault   'debb-03/07/2016
            Else
               cmbEtnia.Text = ""
            End If
            '
            'mo_cmbIdioma.BoundText = ""
            cmbIdioma.Text = ""
            '
            cmbIdioma.Text = ""
            
'<(Inicio) Añadido Por: WABG el: 23/10/2020-07:50:15 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
            mb_validacionReniec = False
'</(Fin) Añadido Por: WABG el: 23/10/2020-07:50:15 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
            '
            mb_UsoWebReniec = False
            txtEmail.Text = ""
            '
            mo_cmbMadreTipoDocumento.BoundText = "1"
            txtMadreDocumento.MaxLength = 8
            txtNroHijo.Text = ""
            txtMadreDocumento.Text = ""
            txtMadreApellidoP.Text = ""
            txtMadreApellidoM.Text = ""
            txtNombreMadre.Text = ""
            txtMadreSnombre.Text = ""
            
            lbUsoWebReniec_SinMostrar = False
            lcFactorRh_SinMostrar = ""
            lcGrupoSanguineo_SinMostrar = ""
            txtHoraNacimiento.Text = "00:00"
            
            txtSector.Text = ""
            txtSectorista.Text = ""
            lblSectorista.Caption = ""
            
            chkSinFechaNacimiento.Value = 0
            cboTipoEdadPaciente.Text = ""
            pi_ImagSeleccionada.Picture = LoadPicture("")
            
            txtGs.Text = ""
            txtFRh.Text = ""
            
'<(Inicio) Añadido Por: WABG el: 16/11/2020-09:44:22 a.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
            txtNroDocumento.Text = ""
'</(Fin) Añadido Por: WABG el: 16/11/2020-09:44:22 a.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
            
            Call bloquearControlEdad
            
End Sub



Public Sub DeshabilitarFrames()
    
    fraDatosHistoriaClinica.Enabled = False
    fraDatosPaciente.Enabled = False
    FraDomicilio.Enabled = False
    fraProcedencia.Enabled = False
    fraNacimiento.Enabled = False
    fraNN.Enabled = False
    fraMadre.Enabled = False
End Sub
Public Sub HabilitarFrames()
    
    fraDatosHistoriaClinica.Enabled = True
    fraDatosPaciente.Enabled = True
    FraDomicilio.Enabled = True
    fraProcedencia.Enabled = True
    fraNacimiento.Enabled = True
    fraNN.Enabled = True
    fraMadre.Enabled = True
End Sub
Public Sub SetFocusEnDNI()
   On Error Resume Next
   txtNroDocumento.SetFocus
End Sub
Public Sub SetFocusOnApellidoPaterno()
         On Error Resume Next
         txtApellidoPaterno.SetFocus
End Sub
Public Sub SetFocusOnDepartamentoDomicilio()
    On Error Resume Next
    cmbIdDepartamentoDomicilio.SetFocus
End Sub
Public Sub SetFocusOnDepartamentoProcedencia()
    'cmbIdDepartamentoProcedencia.SetFocus
End Sub
Public Sub SetFocusOnDepartamentoNacimiento()
    'cmbIdDepartamentoNacimiento.SetFocus
End Sub
Public Sub SetPestaniaTabPaciente(iPestania As Integer)
    TabPaciente.Tab = iPestania
End Sub
Private Sub txtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtTelefono
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoGenHistoriaClinica_click()

    txtFechaCreacion.Text = Format(ldHoy, sighEntidades.DevuelveFechaSoloFormato_DMY)
    
    txtIdNroHistoria.Text = ""
    If cmbIdTipoGenHistoriaClinica.Tag = mo_cmbIdTipoGenHistoriaClinica.BoundText Then
        txtIdNroHistoria.Text = txtIdNroHistoria.Tag
    End If
    
    Select Case mo_cmbIdTipoGenHistoriaClinica.BoundText
    Case sghHistoriaDefinitivaManual
        mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, True
    Case Else
        mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, False
        txtIdNroHistoria.Text = ""
        txtIdNroHistoria.Tag = 0
    End Select
    HabilitaFechaCreacion
End Sub

Sub HabilitaFechaCreacion()
    mo_Formulario.HabilitarDeshabilitar txtFechaCreacion, False
    If Val(mo_cmbIdTipoGenHistoriaClinica.BoundText) = sghHistoriaDefinitivaManual And _
       (mi_Opcion = sghAgregar Or mi_Opcion = sghModificar) Then
       mo_Formulario.HabilitarDeshabilitar txtFechaCreacion, True
    End If
    
End Sub

Private Sub cmbIdTipoGenHistoriaClinica_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, mo_cmbIdTipoGenHistoriaClinica
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub cmbIdTipoGenHistoriaClinica_LostFocus()
   If cmbIdTipoGenHistoriaClinica.Text <> "" Then
       mo_cmbIdTipoGenHistoriaClinica.BoundText = Val(Split(cmbIdTipoGenHistoriaClinica.Text, " = ")(0))
   End If

   mo_Formulario.MarcarComoVacio cmbIdTipoGenHistoriaClinica
End Sub

Private Sub cmbIdTipoGenHistoriaClinica_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtFechaCreacion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaCreacion
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub


Private Sub txtFechaCreacion_LostFocus()
       If txtFechaCreacion <> sighEntidades.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaCreacion, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, "Datos de paciente"
                 txtFechaCreacion = sighEntidades.FECHA_VACIA_DMY
            End If
        End If
   mo_Formulario.MarcarComoVacio txtFechaCreacion
End Sub

Private Sub txtFechaCreacion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdNroHistoria
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
    AdministrarKeyPreview KeyCode
End Sub



Private Sub txtIdNroHistoria_LostFocus()

'<(Inicio) Añadido Por: WABG el: 16/10/2020-12:02:05 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
VerificarExistenciaHistoriaClinica (txtIdNroHistoria.Text)
'</(Fin) Añadido Por: WABG el: 16/10/2020-12:02:05 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>


'<(Inicio)Comentado Por: WABG el: 16/10/2020-12:02:23 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
'    On Error Resume Next
'    If mo_Teclado.TextoEsSoloNumeros(txtIdNroHistoria.Text) Then
'        Dim lbContinua000 As Boolean
'
'        lbContinua000 = True
'        If Len(txtNroDocumento.Text) = 8 And wxParametro351 = "S" Then
'           If txtIdNroHistoria.Text = txtNroDocumento.Text Then
'              lbContinua000 = False
'           End If
'        End If
'        If lbContinua000 = True Then
'            txtIdNroHistoria.Text = mo_Teclado.CapitalizarNombres(txtIdNroHistoria.Text)
'            txtIdNroHistoria.Tag = txtIdNroHistoria.Text
'        End If
'        mo_Formulario.MarcarComoVacio txtIdNroHistoria
'        If txtIdNroHistoria.Locked = True Then Exit Sub
'        If Trim(txtIdNroHistoria.Text) = "" Then txtIdNroHistoria.SetFocus: Exit Sub
'        ms_MensajeError = mo_AdminAdmision.ExisteNroHistoria(Trim(Str(txtIdNroHistoria.Tag)))
'        If ms_MensajeError <> "" Then
'           MsgBox "Existe un paciente con el mismo número de historia clínica: " + Chr(13) + ms_MensajeError
'        End If
'    End If
'</(Fin)Comentado por: WABG el: 16/10/2020-12:02:23 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
End Sub

Private Sub txtIdNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Public Sub HacerVisibleCheckPacienteNoIdentificado(bVisible As Boolean)
    
    If bVisible Then
        'fraNN.Visible = True
        chkNN.Visible = True
        fraDatosHistoriaClinica.Width = 9135
    Else
        'fraNN.Visible = False
        chkNN.Visible = False
        fraDatosHistoriaClinica.Width = fraDatosPaciente.Width
    End If
    
End Sub

Public Function Inicializar()
    
    
    Dim oRsPermisos As New Recordset
    Set oRsPermisos = ms_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(sighEntidades.Usuario)
    oRsPermisos.Filter = "idPermiso=409"
    If oRsPermisos.RecordCount > 0 Then
       cmdCambiaHC.Enabled = True
    Else
       cmdCambiaHC.Enabled = False
    End If
    oRsPermisos.Close
    Set oRsPermisos = Nothing
   

    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica
    HabilitaFechaCreacion
    
    Set mo_CmbIdTipoSexo.MiComboBox = cmbIdTipoSexo
    'Set mo_cmbIdEstadoCivil.MiComboBox = cmbIdEstadoCivil
    Set mo_cmbIdDocIdentidad.MiComboBox = cmbIdDocIdentidad
    'Set mo_cmbIdGradoInstruccion.MiComboBox = cmbIdGradoInstruccion
   ' Set mo_cmbIdTipoOcupacion.MiComboBox = cmbIdTipoOcupacion
    'Set mo_cmbIdProcedencia.MiComboBox = cmbIdProcedencia
    Set mo_cmbIdDepartamentoDomicilio.MiComboBox = cmbIdDepartamentoDomicilio
    Set mo_cmbIdProvinciaDomicilio.MiComboBox = cmbIdProvinciaDomicilio
    Set mo_cmbIdDistritoDomicilio.MiComboBox = cmbIdDistritoDomicilio
    Set mo_cmbIdPaisDomicilio.MiComboBox = cmbIdPaisDomicilio
    Set mo_cmbIdCentroPobladoDomicilio.MiComboBox = cmbIdCentroPobladoDomicilio
    
    Set mo_cmbIdDepartamentoProcedencia.MiComboBox = cmbIdDepartamentoProcedencia
    Set mo_cmbIdProvinciaProcedencia.MiComboBox = cmbIdProvinciaProcedencia
    Set mo_cmbIdDistritoProcedencia.MiComboBox = cmbIdDistritoProcedencia
    Set mo_cmbIdCentroPobladoProcedencia.MiComboBox = cmbIdCentroPobladoProcedencia
    Set mo_cmbIdPaisProcedencia.MiComboBox = cmbIdPaisProcedencia
    
    Set mo_cmbIdDepartamentoNacimiento.MiComboBox = cmbIdDepartamentoNacimiento
    Set mo_cmbIdProvinciaNacimiento.MiComboBox = cmbIdProvinciaNacimiento
    Set mo_cmbIdDistritoNacimiento.MiComboBox = cmbIdDistritoNacimiento
    Set mo_cmbIdCentroPobladoNacimiento.MiComboBox = cmbIdCentroPobladoNacimiento
    Set mo_cmbIdPaisNacimiento.MiComboBox = cmbIdPaisNacimiento
    'Set mo_cmbEtnia.MiComboBox = cmbEtnia
    'Set mo_cmbIdioma.MiComboBox = cmbIdioma
    Set mo_cmbMadreTipoDocumento.MiComboBox = cmbMadreTipoDocumento
    '
    ldHoy = lcBuscaParametro.RetornaFechaHoraServidorSQL
    mo_Formulario.HabilitarDeshabilitar txtEdad, False
    txtHoraNacimiento.Text = "00:00"
   
    'grdEpicrisis.Clear
    If ml_meHwnd = 0 Then
       TabPaciente.TabVisible(3) = False
    End If
    'Ficha Familiar
    lcFormaQgeneraHistoria = "0"
    lcFormaQgeneraHistoria = Trim(lcBuscaParametro.SeleccionaFilaParametro(278))
    If lcBuscaParametro.SeleccionaFilaParametro(277) = "S" Then
       txtFichaFamiliar1.Visible = True: txtFichaFamiliar2.Visible = True
       txtFichaFamiliar3.Visible = True: lblFichaFamiliar.Visible = True
       lblFichaFamiliar1.Visible = True: lblFichaFamiliar1.Caption = ""
    End If
    lcEtniaDefault = lcBuscaParametro.SeleccionaFilaParametro(283)
    lbExigeIngresoDelDNI = IIf(lcBuscaParametro.SeleccionaFilaParametro(287) = "S", True, False)
    lbExigeIngresoDeCentroPoblado = IIf(lcBuscaParametro.SeleccionaFilaParametro(282) = "S", True, False)
    lbBuscaDNIenReniec = IIf(lcBuscaParametro.SeleccionaFilaParametro(296) = "S", True, False)
    If lbBuscaDNIenReniec = True Then
       mo_Reniec.SeAccesaAlaWebDesdeGalenhos = True
       mo_Reniec.Inicializar
    End If
    '
    fraSector.Enabled = False
    If lcBuscaParametro.SeleccionaFilaParametro(282) = "S" Then
       fraSector.Enabled = True
    End If
    mo_Formulario.HabilitarDeshabilitar txtSectorista, False
    '
    Set mo_cmbIdTipoEdad.MiComboBox = cboTipoEdadPaciente
    mo_Formulario.HabilitarDeshabilitar cboTipoEdadPaciente, False
    
    mo_Formulario.HabilitarDeshabilitar txtGs, False
    mo_Formulario.HabilitarDeshabilitar txtFRh, False
    
End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
    On Error Resume Next
    Select Case KeyCode
    Case vbKeyEscape
    Case vbKeyF2
    Case vbKeyF3
     Case vbKeyF4
     Case vbKeyF5
     Case vbKeyF6
     Case vbKeyF7
        UserControl.TabPaciente.Tab = 0
        'On Error Resume Next
        cmbIdProvinciaDomicilio.SetFocus
     Case vbKeyF8
        UserControl.TabPaciente.Tab = 1
    Case vbKeyF9
        UserControl.TabPaciente.Tab = 2
    End Select
       
End Sub


Sub CargaDatosBasicosPacienteNuevo(lcApellidoPaterno As String, lcApellidoMaterno As String, _
                                   lcPrimerNombre As String, IdTipoGenHistoriaClinica As String, _
                                   lcSegundoNombre As String, ldFechaNacimiento As Date, _
                                   lnIdTipoSexo As Long, lcDireccionDomicilio As String, _
                                   lbUsoWebReniec As Boolean, lcNroDni As String, _
                                   lcSegundoNombreSIS As String, lnIdDistritoSIS As Long, _
                                   lnIdSexoSIS As Long, ldFechaNacimientoSIS As Date)
                                   
                                   
    mo_cmbIdTipoGenHistoriaClinica.BoundText = IdTipoGenHistoriaClinica
    txtApellidoPaterno.Text = Trim(lcApellidoPaterno)
    txtApellidoMaterno.Text = Trim(lcApellidoMaterno)
    txtPrimerNombre.Text = Trim(lcPrimerNombre)
    If lcSegundoNombre <> "" Then
       txtSegundoNombre.Text = lcSegundoNombre
       'txtSegundoNombre.Text = ""
    End If
    If ldFechaNacimiento <> 0 Then
       txtFechaNacimiento.Text = ldFechaNacimiento
       ActualizaEdad
       'txtEdad.Text = Trim(Str(EdadActual(CDate(txtFechaNacimiento.Text), Date)))
    End If
    If lnIdTipoSexo > 0 Then
       mo_CmbIdTipoSexo.BoundText = lnIdTipoSexo
    End If
    If Len(lcDireccionDomicilio) > 0 Then
       txtDireccionDomicilio.Text = lcDireccionDomicilio
    End If
    mb_UsoWebReniec = lbUsoWebReniec: MuestraQueUsoWebReniec
    If cmbIdDepartamentoDomicilio.Text = "" Then
       mo_cmbIdDepartamentoDomicilio.BoundText = Left(lcBuscaParametro.SeleccionaFilaParametro(242), 2)
    End If
    If lcNroDni <> "" Then
       txtNroDocumento.Text = lcNroDni 'Frank 29 01 2015
       'txtNroDocumento.Text = ""
    End If
    If lcSegundoNombreSIS <> "" And txtSegundoNombre.Text = "" Then
       txtSegundoNombre.Text = lcSegundoNombreSIS 'Frank 29 01 2015
       'txtSegundoNombre.Text = ""
    End If
    If lnIdDistritoSIS > 0 And Val(mo_cmbIdDistritoDomicilio.BoundText) = 0 Then
       Dim lcIdDistritoSIS As String
       lcIdDistritoSIS = Right("0" & Trim(Str(lnIdDistritoSIS)), 6)
       mo_cmbIdDepartamentoDomicilio.BoundText = Val(Left(lcIdDistritoSIS, 2))
       mo_cmbIdProvinciaDomicilio.BoundText = Val(Left(lcIdDistritoSIS, 4))
       mo_cmbIdDistritoDomicilio.BoundText = Val(lcIdDistritoSIS)
    End If
    If lnIdSexoSIS > 0 And Val(mo_CmbIdTipoSexo.BoundText) = 0 Then
       mo_CmbIdTipoSexo.BoundText = lnIdSexoSIS
    End If
    If ldFechaNacimientoSIS <> 0 And txtFechaNacimiento.Text = sighEntidades.FECHA_VACIA_DMY Then
       txtFechaNacimiento.Text = ldFechaNacimientoSIS
       'txtEdad.Text = Trim(Str(EdadActual(CDate(txtFechaNacimiento.Text), Date)))
       ActualizaEdad
    End If
'    txtSegundoNombre.SetFocus
    If Val(lcEtniaDefault) > 0 Then
       cmbEtnia_UbicaPosicion lcEtniaDefault   'debb-03/07/2016
    Else
       cmbEtnia.Text = ""
    End If
    HabilitaFechaCreacion
End Sub
Sub MuestraQueUsoWebReniec()
    lblFichaFamiliar1.Visible = False
    lblFichaFamiliar1.Caption = ""
    If mb_UsoWebReniec = True Then
       lblFichaFamiliar1.Visible = True
       lblFichaFamiliar1.Caption = "Usó la WEB RENIEC"
    End If
End Sub

Sub CargaDireccionDomicilioDeLaAtencion(lcDireccionDomicilio As String)
    txtDireccionDomicilio.Text = lcDireccionDomicilio
    Label38.Caption = "Direc(Fech.Atenc)"
End Sub
Private Sub txtHoraNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraNacimiento
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
    AdministrarKeyPreview KeyCode
End Sub



Private Function VerificaSiExistePaciente(ap As String, am As String, PN As String, SN As String, TN As String)
  If mi_Opcion <> sghAgregar Then Exit Function
  If Trim(ap) = "" Or Trim(am) = "" Or Trim(PN) = "" Then Exit Function
  If UserControl.chkNN.Value = 1 Then Exit Function
  Dim rsTmp As ADODB.Recordset
  Set rsTmp = mo_AdminAdmision.ExisteAlgunPaciente(Trim(ap), Trim(am), Trim(PN), Trim(SN), Trim(TN))
  If Not (rsTmp.EOF = True And rsTmp.BOF = True) Then
     rsTmp.MoveFirst
     Do While Not rsTmp.EOF
        If rsTmp.Fields!idPaciente <> ml_IdPaciente Then
           MsgBox "Existe un paciente con los mismos nombres y apellidos (HC: " & _
           HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsTmp.Fields!NroHistoriaClinica)), False) & _
           ")" & Chr(13) & "Complete los apellidos y nombres ó verifique los ya ingresados", vbInformation, "Datos de paciente"
           Exit Do
        End If
        rsTmp.MoveNext
     Loop
  End If
  rsTmp.Close
End Function

Public Sub TabEnNroHistoria()
    On Error Resume Next
    txtIdNroHistoria.SetFocus
End Sub


Sub CargaDatosPersonales(oPacientes As doPaciente, lcParametro242 As String, lcParametro237 As String)
Dim lcRutaImg As String
Dim lcUbigeoDistrito As String
           With oPacientes
                
                txtGs.Text = .GrupoSanguineo
                txtFRh.Text = .FactorRh
                
                chkNN.Value = 0
                If UCase(.ApellidoPaterno) = "NN" And UCase(.ApellidoMaterno) = "NN" And UCase(.PrimerNombre) = "NN" And UCase(.SegundoNombre) = "NN" Then
                    chkNN.Value = 1
                End If
                
                Me.idPaciente = .idPaciente
                txtIdPaciente = .idPaciente
                Autogenerado = .Autogenerado
                txtApellidoPaterno.Text = Trim(.ApellidoPaterno)
                txtApellidoMaterno.Text = Trim(.ApellidoMaterno)
                txtPrimerNombre.Text = Trim(.PrimerNombre)
                txtSegundoNombre.Text = Trim(.SegundoNombre)
                txtTercerNombre.Text = Trim(.TercerNombre)
                If .FechaNacimiento <> 0 Then
                    txtFechaNacimiento.Text = Format(.FechaNacimiento, sighEntidades.DevuelveFechaSoloFormato_DMY)
                    txtHoraNacimiento.Text = Format(.FechaNacimiento, sighEntidades.DevuelveHoraSoloFormato_HM)
                End If
                RaiseEvent SeModificoFechaNacimiento(txtFechaNacimiento.Text, txtHoraNacimiento.Text)
                
                txtTelefono.Text = Trim(.TELEFONO)
                txtDireccionDomicilio.Text = Trim(.DireccionDomicilio)
                mo_CmbIdTipoSexo.BoundText = .idTipoSexo
                RaiseEvent SeModificoSexo(.idTipoSexo)
                '
                'mo_cmbIdProcedencia.BoundText = .IdProcedencia
                If .IdProcedencia > 0 Then
                   cmbIdProcedencia_UbicaPosicion (.IdProcedencia)
                Else
                   cmbIdProcedencia.Text = ""
                End If
                '
                'mo_cmbIdGradoInstruccion.BoundText = .IdGradoInstruccion
                If .IdGradoInstruccion > 0 Then
                   cmbIdGradoInstruccion_UbicaPosicion (.IdGradoInstruccion)
                Else
                   cmbIdGradoInstruccion.Text = ""
                End If
                
                '
                'mo_cmbIdEstadoCivil.BoundText = .IdEstadoCivil
                If .IdEstadoCivil > 0 Then
                   cmbIdEstadoCivil_UbicaPosicion (.IdEstadoCivil)
                Else
                   cmbIdEstadoCivil.Text = ""
                End If
                '
                mo_cmbIdDocIdentidad.BoundText = .IdDocIdentidad
                If .IdDocIdentidad <> 1 Then
                   txtNroDocumento.MaxLength = 12
                End If
                txtNroDocumento.Text = Trim(.nrodocumento)
                '
                'mo_cmbIdTipoOcupacion.BoundText = .idTipoOcupacion
                If .idTipoOcupacion > 0 Then
                   cmbIdTipoOcupacion_UbicaPosicion (.idTipoOcupacion)
                Else
                   cmbIdTipoOcupacion.Text = ""
                End If
                '
                Dim oRsBuscaUbigeo As New Recordset
                '
                Set oRsBuscaUbigeo = mo_AdminAdmision.CentrosPobladosDevuelveDptoProvDistritoSegunIdCentroPoblado(.IdCentroPobladoNacimiento)
                mo_cmbIdPaisNacimiento.BoundText = .IdPaisNacimiento
                If oRsBuscaUbigeo.RecordCount > 0 Then
                    mo_cmbIdDepartamentoNacimiento.BoundText = oRsBuscaUbigeo.Fields!IdDepartamento    '.IdDepartamentoNacimiento
                    mo_cmbIdProvinciaNacimiento.BoundText = oRsBuscaUbigeo.Fields!IdProvincia      '.IdProvinciaNacimiento
                    mo_cmbIdDistritoNacimiento.BoundText = oRsBuscaUbigeo.Fields!IdDistrito      '.IdDistritoNacimiento
                Else
                   If .IdDistritoNacimiento > 0 Then
                        lcUbigeoDistrito = Right("0" & Trim(Str(.IdDistritoNacimiento)), 6)
                        mo_cmbIdDepartamentoNacimiento.BoundText = Val(Left(lcUbigeoDistrito, 2))
                        mo_cmbIdProvinciaNacimiento.BoundText = Val(Left(lcUbigeoDistrito, 4))
                        mo_cmbIdDistritoNacimiento.BoundText = Val(lcUbigeoDistrito)
                   End If
                End If
                mo_cmbIdCentroPobladoNacimiento.BoundText = .IdCentroPobladoNacimiento
                '
                Set oRsBuscaUbigeo = mo_AdminAdmision.CentrosPobladosDevuelveDptoProvDistritoSegunIdCentroPoblado(.IdCentroPobladoDomicilio)
                mo_cmbIdPaisDomicilio.BoundText = .IdPaisDomicilio
                If oRsBuscaUbigeo.RecordCount > 0 Then
                    mo_cmbIdDepartamentoDomicilio.BoundText = oRsBuscaUbigeo.Fields!IdDepartamento      '.IdDepartamentoDomicilio
                    mo_cmbIdProvinciaDomicilio.BoundText = oRsBuscaUbigeo.Fields!IdProvincia      '.IdProvinciaDomicilio
                    mo_cmbIdDistritoDomicilio.BoundText = oRsBuscaUbigeo.Fields!IdDistrito      '.IdDistritoDomicilio
                Else
                   If .IdDistritoDomicilio > 0 Then
                        lcUbigeoDistrito = Right("0" & Trim(Str(.IdDistritoDomicilio)), 6)
                        mo_cmbIdDepartamentoDomicilio.BoundText = Val(Left(lcUbigeoDistrito, 2))
                        mo_cmbIdProvinciaDomicilio.BoundText = Val(Left(lcUbigeoDistrito, 4))
                        mo_cmbIdDistritoDomicilio.BoundText = Val(lcUbigeoDistrito)
                   End If
                End If
                mo_cmbIdCentroPobladoDomicilio.BoundText = .IdCentroPobladoDomicilio
                If cmbIdDepartamentoDomicilio.Text = "" Then
                   mo_cmbIdDepartamentoDomicilio.BoundText = Left(lcParametro242, 2)
                End If
                '
                Set oRsBuscaUbigeo = mo_AdminAdmision.CentrosPobladosDevuelveDptoProvDistritoSegunIdCentroPoblado(.IdCentroPobladoProcedencia)
                mo_cmbIdPaisProcedencia.BoundText = .IdPaisProcedencia
                If oRsBuscaUbigeo.RecordCount > 0 Then
                    mo_cmbIdDepartamentoProcedencia.BoundText = oRsBuscaUbigeo.Fields!IdDepartamento      '.IdDepartamentoProcedencia
                    mo_cmbIdProvinciaProcedencia.BoundText = oRsBuscaUbigeo.Fields!IdProvincia      '.IdProvinciaProcedencia
                    mo_cmbIdDistritoProcedencia.BoundText = oRsBuscaUbigeo.Fields!IdDistrito      '.IdDistritoProcedencia
                Else
                   If .IdDistritoProcedencia > 0 Then
                        lcUbigeoDistrito = Right("0" & Trim(Str(.IdDistritoProcedencia)), 6)
                        mo_cmbIdDepartamentoProcedencia.BoundText = Val(Left(lcUbigeoDistrito, 2))
                        mo_cmbIdProvinciaProcedencia.BoundText = Val(Left(lcUbigeoDistrito, 4))
                        mo_cmbIdDistritoProcedencia.BoundText = Val(lcUbigeoDistrito)
                   End If
                End If
                mo_cmbIdCentroPobladoProcedencia.BoundText = .IdCentroPobladoProcedencia
                Set oRsBuscaUbigeo = Nothing
                txtNombrePadre.Text = Trim(.NombrePadre)
                txtNombreMadre.Text = Trim(.Nombremadre)
                
                Select Case .idTipoNumeracion
                Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
                      
                    Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()
                    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
                    mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, False
                Case Else 'sghHistoriaTemporalCOnsultaExterna, sghHistoriaTemporalEmergencia, sghSinHistoria
                    Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriaSeleccionarDefinitivos(.idTipoNumeracion)
                    mo_cmbIdTipoGenHistoriaClinica.BoundText = .idTipoNumeracion
                End Select
                mo_cmbIdTipoGenHistoriaClinica.BoundText = .idTipoNumeracion
                HabilitaFechaCreacion
                
                cmbIdTipoGenHistoriaClinica.Tag = .idTipoNumeracion         'lo guarda para luego comparar
                txtIdNroHistoria.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(.NroHistoriaClinica)), False)          'esto tiene que ir luego del tipo de generacion, por que sino se borra con el change del combo box
                txtIdNroHistoria.Tag = .NroHistoriaClinica
                
                'Esto debe ir aqui despues del setear el tipo de generacion
                Select Case .idTipoNumeracion
                Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
                    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
                    mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, False
                End Select
                txtFechaNacimiento_LostFocus
                
'                txtEtapaDomicilio.Text = Trim(.EtapaDomicilio)
'                txtSectorDomicilio.Text = Trim(.SectorDomicilio)
'                txtLoteDomicilio.Text = Trim(.LoteDomicilio)
'                txtManzanaDomicilio.Text = Trim(.ManzanaDomicilio)
'                txtPisoDomicilio.Text = Trim(.PisoDomicilio)
'                txtNroDomicilio.Text = Trim(.NroDomicilio)
                
                If txtIdNroHistoria.Text <> 0 Then
                    Select Case .idTipoNumeracion
                    Case sghHistoriaDefinitivaManual, sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
                        Dim oDOHistoria As DOHistoriaClinica
                        Set oDOHistoria = mo_AdminArchivoClinico.HistoriaClinicaSeleccionarPorId(txtIdNroHistoria.Tag)
                        If Not oDOHistoria Is Nothing Then
                            txtFechaCreacion = Format(oDOHistoria.fechacreacion, sighEntidades.DevuelveFechaSoloFormato_DMY)
                        Else
                            MsgBox "Por algún motivo el paciente no tiene el registro asociado en la tabla de historias clinicas, consulte al administrador de sistemas", vbInformation, "Datos de paciente"
                        End If
                    End Select
                Else
                    MsgBox "El paciente no tiene historia clinica", vbInformation, "Datos de paciente"
                End If
                txtObservacion.Text = Trim(.Observacion)
                If Not IsNull(.FichaFamiliar) Then
                    txtFichaFamiliar1.Text = DevuelveParteFichaFamiliar(.FichaFamiliar, 1)
                    txtFichaFamiliar2.Text = DevuelveParteFichaFamiliar(.FichaFamiliar, 2)
                    txtFichaFamiliar3.Text = DevuelveParteFichaFamiliar(.FichaFamiliar, 3)
                Else
                    txtFichaFamiliar1.Text = ""
                    txtFichaFamiliar2.Text = ""
                    txtFichaFamiliar3.Text = ""
                End If
                '

                If IsNull(.IdEtnia) Or .IdEtnia = "" Then
                   If lcEtniaDefault <> "" Then
                      cmbEtnia_UbicaPosicion lcEtniaDefault   'debb-03/07/2016
                   End If
                Else
                   cmbEtnia_UbicaPosicion (.IdEtnia)
                End If
                '
                If .IdIdioma > 0 Then
                   'mo_cmbIdioma.BoundText = .IdIdioma
                   cmbIdioma_UbicaPosicion (.IdIdioma)
                End If
                '
                mb_UsoWebReniec = .UsoWebReniec
                txtEmail.Text = .Email
                txtNroHijo.Text = .NroOrdenHijo
                mo_cmbMadreTipoDocumento.BoundText = .madreTipoDocumento
                txtMadreDocumento.Text = .madreDocumento
                txtMadreApellidoP.Text = .madreApellidoPaterno
                txtMadreApellidoM.Text = .madreApellidoMaterno
                txtNombreMadre.Text = .madrePrimerNombre
                txtMadreSnombre.Text = .madreSegundoNombre
                '
                lbUsoWebReniec_SinMostrar = .UsoWebReniec
                lcFactorRh_SinMostrar = .FactorRh
                lcGrupoSanguineo_SinMostrar = .GrupoSanguineo
                '
                txtSector.Text = .sector
                txtSectorista.Text = .sectorista
                BuscaEmpleadoYllenaDatosDelSectorista .sectorista
                '
                Call cargarDatosPersonalesAdicionales(oPacientes)

                mb_ExistenDatos = True
            End With
            
            'Excepciones
            If mo_cmbIdPaisDomicilio.BoundText = "" Then
                mo_cmbIdPaisDomicilio.BoundText = "166" 'Peru
            End If
            If mo_cmbIdPaisNacimiento.BoundText = "" Then
                mo_cmbIdPaisNacimiento.BoundText = "166" 'Peru
            End If
            If mo_cmbIdPaisProcedencia.BoundText = "" Then
                mo_cmbIdPaisProcedencia.BoundText = "166" 'Peru
            End If
            'carga Imagen..........si demora mucho al cargar, cambiar en parametros la ruta
            lcRutaImg = lcParametro237 & "\" & Trim(Str(txtIdNroHistoria.Tag)) & ".jpg"
            If sighEntidades.ArchivoExiste(lcRutaImg) Then
               pi_ImagSeleccionada.Picture = LoadPicture(lcRutaImg)
            Else
               pi_ImagSeleccionada.Picture = LoadPicture("")
            End If
            '
            If ml_meHwnd > 0 Then
'               CargaEpicrisisEscaneadas
'               CargaPDFgenerados
               ucPacientesPDF1.Inicializar ml_IdPaciente, txtIdNroHistoria.Tag
            End If

End Sub



Public Sub CargarDatosDePacienteALosControlesSinBuscar(oPacientes As doPaciente, wxParametro242 As String, wxParametro237 As String)
    CargaDatosPersonales oPacientes, wxParametro242, wxParametro237
End Sub

Function DevuelveParteFichaFamiliar(lcFichaFamiliarJunta As String, lnParte As Integer) As String
    Dim lnFor As Integer, lnPos1 As Integer, lnPos2 As Integer
    lnPos1 = InStr(lcFichaFamiliarJunta, "-")
    If lcFichaFamiliarJunta = "" Or lnPos1 = 0 Then
        DevuelveParteFichaFamiliar = ""
    Else
        Select Case lnParte
        Case 1
             DevuelveParteFichaFamiliar = Left(lcFichaFamiliarJunta, InStr(lcFichaFamiliarJunta, "-") - 1)
        Case 2
             lnPos1 = InStr(lcFichaFamiliarJunta, "-")
             lnPos2 = InStr(lnPos1 + 1, lcFichaFamiliarJunta, "-")
             DevuelveParteFichaFamiliar = Mid(lcFichaFamiliarJunta, lnPos1 + 1, lnPos2 - lnPos1 - 1)
        Case 3
             lnPos1 = InStr(lcFichaFamiliarJunta, "-")
             lnPos2 = InStr(lnPos1 + 1, lcFichaFamiliarJunta, "-")
             DevuelveParteFichaFamiliar = Mid(lcFichaFamiliarJunta, lnPos2 + 1, Len(lcFichaFamiliarJunta))
        End Select
    End If
End Function

Public Function DevuelvePaciente() As String
    DevuelvePaciente = Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & " " & Trim(txtPrimerNombre.Text) & " " & IIf(IsNull(txtSegundoNombre.Text), "", Trim(txtSegundoNombre.Text))
End Function

Public Function DevuelveDocumento() As String
    DevuelveDocumento = cmbIdDocIdentidad.Text
End Function

Public Function DevuelveNroDocumento() As String
    DevuelveNroDocumento = txtNroDocumento.Text
End Function

Public Function DevuelveUbigeoDomicilio() As String
    DevuelveUbigeoDomicilio = txtDireccionDomicilio.Text
End Function
Public Function DevuelvePaisDomicilio() As String
    DevuelvePaisDomicilio = cmbIdPaisDomicilio.Text
End Function

Public Function DevuelveSiElPacienteEsNN() As Boolean
    DevuelveSiElPacienteEsNN = IIf(chkNN.Value = 1, True, False)
End Function

Function DevuelveFichaFamiliarUnida() As String
    DevuelveFichaFamiliarUnida = Trim(txtFichaFamiliar1.Text) & "-" & Trim(txtFichaFamiliar2.Text) & "-" & Trim(txtFichaFamiliar3.Text)
End Function
Public Sub SetFocusEnIdioma()
    On Error Resume Next
    cmbIdioma.SetFocus
End Sub
Public Sub SetFocusEnEtnia()
    On Error Resume Next
    cmbEtnia.SetFocus
End Sub

Function DevuelveEtnia() As String
    DevuelveEtnia = Trim(cmbEtnia.Text)
End Function
Function DevuelveIdioma() As String
    DevuelveIdioma = Trim(cmbIdioma.Text)
End Function

Public Function DevuelveApaterno() As String
    DevuelveApaterno = Trim(txtApellidoPaterno.Text)
End Function

Public Function DevuelveAmaterno() As String
    DevuelveAmaterno = Trim(txtApellidoMaterno.Text)
End Function

Public Function DevuelvePnombre() As String
    DevuelvePnombre = Trim(txtPrimerNombre.Text)
End Function

Public Function DevuelveSnombre() As String
    DevuelveSnombre = Trim(txtSegundoNombre.Text)
End Function
Public Function DevuelveDNI() As String
    If mo_cmbIdDocIdentidad.BoundText = "1" Then
       DevuelveDNI = UserControl.txtNroDocumento.Text
    Else
       DevuelveDNI = ""
    End If
End Function
Public Function DevuelveFechaNacimiento() As String
    DevuelveFechaNacimiento = txtFechaNacimiento.Text
End Function

Public Function DevuelveHoraNacimiento() As String
    DevuelveHoraNacimiento = txtHoraNacimiento.Text
End Function

Public Function DevuelveSexo() As String
    DevuelveSexo = cmbIdTipoSexo.Text
End Function

Sub cmbEtnia_UbicaPosicion(lcCodigoEtnia As String)
    Dim lnFor As Integer
    For lnFor = 0 To (cmbEtnia.ListCount - 1)
        cmbEtnia.ListIndex = lnFor
        If cmbEtnia.SubItem(cmbEtnia.ListIndex, 0) = Val(lcCodigoEtnia) Then
           Exit For
        End If
    Next
End Sub

Sub cmbIdioma_UbicaPosicion(lnIdIdioma As Long)
    Dim lnFor As Integer
    For lnFor = 0 To (cmbIdioma.ListCount - 1)
        cmbIdioma.ListIndex = lnFor
        If cmbIdioma.SubItem(cmbIdioma.ListIndex, 0) = Val(lnIdIdioma) Then
           Exit For
        End If
    Next
End Sub

Sub cmbIdEstadoCivil_UbicaPosicion(lnIdEstadoCivil As Long)
    Dim lnFor As Integer
    For lnFor = 0 To (cmbIdEstadoCivil.ListCount - 1)
        cmbIdEstadoCivil.ListIndex = lnFor
        If cmbIdEstadoCivil.SubItem(cmbIdEstadoCivil.ListIndex, 0) = Val(lnIdEstadoCivil) Then
           Exit For
        End If
    Next
End Sub

Sub cmbIdTipoOcupacion_UbicaPosicion(lnIdTipoOcupacion As Long)
    Dim lnFor As Integer
    For lnFor = 0 To (cmbIdTipoOcupacion.ListCount - 1)
        cmbIdTipoOcupacion.ListIndex = lnFor
        If cmbIdTipoOcupacion.SubItem(cmbIdTipoOcupacion.ListIndex, 0) = Val(lnIdTipoOcupacion) Then
           Exit For
        End If
    Next
End Sub

Sub cmbIdGradoInstruccion_UbicaPosicion(lnIdGradoInstruccion As Long)
    Dim lnFor As Integer
    For lnFor = 0 To (cmbIdGradoInstruccion.ListCount - 1)
        cmbIdGradoInstruccion.ListIndex = lnFor
        If cmbIdGradoInstruccion.SubItem(cmbIdGradoInstruccion.ListIndex, 0) = Val(lnIdGradoInstruccion) Then
           Exit For
        End If
    Next
End Sub

Sub cmbIdProcedencia_UbicaPosicion(lnIdProcedencia As Long)
    Dim lnFor As Integer
    For lnFor = 0 To (cmbIdProcedencia.ListCount - 1)
        cmbIdProcedencia.ListIndex = lnFor
        If cmbIdProcedencia.SubItem(cmbIdProcedencia.ListIndex, 0) = Val(lnIdProcedencia) Then
           Exit For
        End If
    Next
End Sub
Public Function DevuelveFechaCreacionHistoria() As Date
    DevuelveFechaCreacionHistoria = CDate(txtFechaCreacion.Text)
End Function

Public Sub SetFocusEnHistoria()
   On Error Resume Next
   txtIdNroHistoria.SetFocus
End Sub

Private Sub bloquearControlEdad()
    mo_Formulario.HabilitarDeshabilitar txtEdad, False
    mo_Formulario.HabilitarDeshabilitar cboTipoEdadPaciente, False
    mo_Formulario.HabilitarDeshabilitar txtFechaNacimiento, True
    mo_Formulario.HabilitarDeshabilitar txtHoraNacimiento, True
    
    If chkSinFechaNacimiento.Value = 1 Then
        mo_Formulario.HabilitarDeshabilitar txtEdad, True
        mo_Formulario.HabilitarDeshabilitar cboTipoEdadPaciente, True
        mo_Formulario.HabilitarDeshabilitar txtFechaNacimiento, False
        mo_Formulario.HabilitarDeshabilitar txtHoraNacimiento, False
        On Error Resume Next
        txtEdad.SetFocus
    End If
    
End Sub

Private Sub cargarDatosPersonalesAdicionales(oDOPaciente As doPaciente)
On Error GoTo miError
    Dim oConexion As ADODB.Connection
    Set oConexion = New ADODB.Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
            
    Dim oDoPacienteDatosAdd As New DoPacienteDatosAdd
    Set oDoPacienteDatosAdd = mo_AdminAdmision.PacientesDatosAdicionalesSeleccionarPorId(oDOPaciente.idPaciente, _
                                oConexion)
    If Not (oDoPacienteDatosAdd Is Nothing) Then
        chkSinFechaNacimiento.Value = IIf(oDoPacienteDatosAdd.FNacimientoCalculada = True, 1, 0)
    End If
    oConexion.Close
    Set oConexion = Nothing
    Call bloquearControlEdad
    
miError:
    If Err Then
        Exit Sub
    End If
End Sub

Private Function calcularFechaDeNacimiento(edadEnAnios As String, ms_tipoEdad As String) As Date
    If txtEdad.Enabled = False Or txtEdad.Locked = True Then
        Exit Function
    End If
    If edadEnAnios <> "" And ms_tipoEdad <> "" Then
        Dim oFechaHOra As New FechaHora
        Dim md_fechaNacimiento As Date
        Dim horaServidor As String
        Dim FechaHoraServidor As Date
        
        FechaHoraServidor = lcBuscaParametro.RetornaFechaServidorSQLserver
        
        If Val(ms_tipoEdad) = sghTipoEdades.sghHoras Then
            horaServidor = Format(FechaHoraServidor, oFechaHOra.DevuelveHoraSoloFormato_HM)
        Else
            horaServidor = Format("00:00", oFechaHOra.DevuelveHoraSoloFormato_HM)
        End If
        
        md_fechaNacimiento = oFechaHOra.DevuelveFechaNacimiento(Format(FechaHoraServidor, oFechaHOra.DevuelveFechaSoloFormato_DMY), _
                                    horaServidor, CInt(edadEnAnios), Val(ms_tipoEdad))
        txtFechaNacimiento.Text = Format(md_fechaNacimiento, oFechaHOra.DevuelveFechaSoloFormato_DMY)
        UserControl.txtHoraNacimiento.Text = Format(md_fechaNacimiento, oFechaHOra.DevuelveHoraSoloFormato_HM)
    Else
        txtFechaNacimiento.Text = oFechaHOra.FECHA_VACIA_DMY
        txtHoraNacimiento.Text = oFechaHOra.HORA_VACIA_HM
    End If
End Function

'mgary201503
Private Function ValidarNumeroDeHistoriaClinica(ByRef messageError As String) As Boolean
    Dim sMensajeLocal As String
    ValidarNumeroDeHistoriaClinica = False
    
    Select Case Val(mo_cmbIdTipoGenHistoriaClinica.BoundText)
        Case sghHistoriaTemporalCOnsultaExterna, sghHistoriaTemporalEmergencia, sghSinHistoria
        Case sghHistoriaDefinitivaManual
            If Trim(txtIdNroHistoria.Text) <> "" Then
                Dim lUltimoNumeroHistoria As Long
                lUltimoNumeroHistoria = mo_AdminAdmision.UltimoNroHistoriaGenerado()
                If Val(txtIdNroHistoria.Text) > lUltimoNumeroHistoria + 1 Then
'                    sMensajeLocal = "Número de Historia Ingresado no puede ser mayor que " & CStr(lUltimoNumeroHistoria + 1)
'                    messageError = sMensajeLocal
'                    Exit Function
                End If
            End If
        Case Else
    End Select
    ValidarNumeroDeHistoriaClinica = True
End Function
