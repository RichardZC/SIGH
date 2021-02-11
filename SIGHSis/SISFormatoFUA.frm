VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form SisFuaVersion2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SISFormatoFUA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInstitucionEducativa 
      Height          =   1400
      Left            =   6360
      TabIndex        =   86
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox cmbColegioTurno 
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
         Left            =   4680
         TabIndex        =   152
         Top             =   960
         Width           =   1410
      End
      Begin VB.ComboBox cmbColegioGrado 
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
         Left            =   1800
         TabIndex        =   151
         Top             =   960
         Width           =   1650
      End
      Begin VB.ComboBox cmbColegioNivel 
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
         Left            =   120
         TabIndex        =   149
         Top             =   960
         Width           =   1650
      End
      Begin VB.TextBox txtColegioCodigo 
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
         TabIndex        =   92
         ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
         Top             =   330
         Width           =   1155
      End
      Begin VB.TextBox txtColegio 
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
         TabIndex        =   91
         ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
         Top             =   330
         Width           =   4425
      End
      Begin VB.CommandButton btnBuscaColegio 
         Caption         =   "..."
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
         Left            =   1320
         TabIndex        =   90
         Top             =   330
         Width           =   315
      End
      Begin VB.TextBox txtColegioSeccion 
         Height          =   315
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   87
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label94 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Grado"
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
         Left            =   1800
         TabIndex        =   150
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label93 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel"
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
         TabIndex        =   148
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "Código                 Institución Educativa"
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
         TabIndex        =   93
         Top             =   120
         Width           =   3270
      End
      Begin VB.Label Label90 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Turno"
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
         Left            =   4800
         TabIndex        =   89
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label89 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sección"
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
         Left            =   3480
         TabIndex        =   88
         Top             =   720
         Width           =   630
      End
   End
   Begin SIGHSis.ucSISfuaCodPrestacion ucSISfuaCodPrestacion1 
      Height          =   405
      Left            =   120
      TabIndex        =   84
      Top             =   550
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   714
   End
   Begin VB.TextBox txtFua1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   2970
      TabIndex        =   70
      ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
      Top             =   90
      Width           =   645
   End
   Begin VB.TextBox txtFua3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   4110
      MaxLength       =   8
      TabIndex        =   0
      Top             =   90
      Width           =   2175
   End
   Begin VB.TextBox txtFua2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   3630
      TabIndex        =   69
      Top             =   90
      Width           =   465
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
      Height          =   915
      Left            =   0
      TabIndex        =   9
      Top             =   9390
      Width           =   12510
      Begin VB.CommandButton btnguardafua 
         Caption         =   "Command1"
         Height          =   480
         Left            =   8520
         TabIndex        =   258
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprime FUA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   90
         Picture         =   "SISFormatoFUA.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "SISFormatoFUA.frx":11A3
         DownPicture     =   "SISFormatoFUA.frx":1667
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   10680
         Picture         =   "SISFormatoFUA.frx":1B53
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "SISFormatoFUA.frx":203F
         DownPicture     =   "SISFormatoFUA.frx":249F
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4845
         Picture         =   "SISFormatoFUA.frx":2914
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   1365
      End
   End
   Begin TabDlg.SSTab TabFua 
      Height          =   7635
      Left            =   0
      TabIndex        =   10
      Top             =   1710
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   13467
      _Version        =   393216
      Tab             =   2
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
      TabCaption(0)   =   "Cabecera  (F3)"
      TabPicture(0)   =   "SISFormatoFUA.frx":2D89
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame17"
      Tab(0).Control(1)=   "Frame16"
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "N° Hoja Refer/Cont"
      TabPicture(1)   =   "SISFormatoFUA.frx":2DA5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtMedicoRNE"
      Tab(1).Control(1)=   "chkMedicoEgresado"
      Tab(1).Control(2)=   "txtMedicoEspecialidad"
      Tab(1).Control(3)=   "txtMedicoDni"
      Tab(1).Control(4)=   "txtMedicoColegiatura"
      Tab(1).Control(5)=   "txtMedico"
      Tab(1).Control(6)=   "FraDx"
      Tab(1).Control(7)=   "Frame4"
      Tab(1).Control(8)=   "Frame5"
      Tab(1).Control(9)=   "btnRefrescar"
      Tab(1).Control(10)=   "Label53"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Medicamentos/Cpt  (F5)"
      TabPicture(2)   =   "SISFormatoFUA.frx":2DC1
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label15"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "grdDiag"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame19"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "FraPatologia"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "FraFarmacia"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.TextBox txtMedicoRNE 
         Height          =   315
         Left            =   -65160
         TabIndex        =   250
         Top             =   7200
         Width           =   1305
      End
      Begin VB.CheckBox chkMedicoEgresado 
         Caption         =   "EGRESADO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -63840
         TabIndex        =   249
         Top             =   7200
         Width           =   1215
      End
      Begin VB.Frame FraFarmacia 
         Caption         =   "Farmacia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   120
         TabIndex        =   65
         Top             =   360
         Width           =   12135
         Begin VB.CommandButton btnAddFarmacia 
            DisabledPicture =   "SISFormatoFUA.frx":2DDD
            DownPicture     =   "SISFormatoFUA.frx":31C6
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   11580
            Picture         =   "SISFormatoFUA.frx":35D2
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   300
            Width           =   405
         End
         Begin UltraGrid.SSUltraGrid grdFarmacia 
            Height          =   2655
            Left            =   120
            TabIndex        =   67
            Top             =   225
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   4683
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "SSUltraGrid1"
         End
      End
      Begin VB.Frame FraPatologia 
         Caption         =   "Procedimientos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   120
         TabIndex        =   62
         Top             =   3345
         Width           =   12165
         Begin VB.CommandButton btnAddPatologia 
            DisabledPicture =   "SISFormatoFUA.frx":39DE
            DownPicture     =   "SISFormatoFUA.frx":3DC7
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   11670
            Picture         =   "SISFormatoFUA.frx":41D3
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   240
            Width           =   405
         End
         Begin UltraGrid.SSUltraGrid grdPatologia 
            Height          =   1740
            Left            =   60
            TabIndex        =   63
            Top             =   240
            Width           =   11565
            _ExtentX        =   20399
            _ExtentY        =   3069
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "SSUltraGrid1"
         End
      End
      Begin VB.TextBox txtMedicoEspecialidad 
         Height          =   315
         Left            =   -66480
         TabIndex        =   61
         Top             =   7200
         Width           =   1305
      End
      Begin VB.TextBox txtMedicoDni 
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
         Left            =   -74820
         TabIndex        =   59
         ToolTipText     =   "DNI del médico"
         Top             =   7200
         Width           =   1845
      End
      Begin VB.TextBox txtMedicoColegiatura 
         Height          =   315
         Left            =   -67800
         TabIndex        =   58
         Top             =   7200
         Width           =   1275
      End
      Begin VB.TextBox txtMedico 
         Height          =   315
         Left            =   -72960
         TabIndex        =   57
         Top             =   7200
         Width           =   5160
      End
      Begin VB.Frame Frame19 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   120
         TabIndex        =   55
         Top             =   6120
         Width           =   8100
         Begin VB.TextBox txtObservaciones 
            Height          =   1125
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   56
            Top             =   240
            Width           =   7845
         End
      End
      Begin VB.Frame FraDx 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   -74880
         TabIndex        =   31
         Top             =   4620
         Width           =   12255
         Begin VB.TextBox lblDiag3 
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
            Height          =   315
            Left            =   8880
            TabIndex        =   83
            Text            =   "Dx Egreso"
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   250
            Width           =   3315
         End
         Begin VB.TextBox lblDiag2 
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
            Height          =   315
            Left            =   4860
            TabIndex        =   82
            Text            =   "Dx Ingreso"
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   250
            Width           =   4005
         End
         Begin VB.TextBox lblDiag1 
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
            Height          =   315
            Left            =   60
            TabIndex        =   81
            Text            =   "Diagnósticos"
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   250
            Width           =   4785
         End
         Begin UltraGrid.SSUltraGrid grdDx 
            Height          =   1290
            Left            =   60
            TabIndex        =   48
            Top             =   615
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   2275
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "SSUltraGrid1"
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Servicios Preventivos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4305
         Left            =   -74850
         TabIndex        =   30
         Top             =   330
         Width           =   7880
         Begin VB.Frame fraTamizajeSaludM 
            Caption         =   "Tamizaje Salud M."
            Height          =   480
            Left            =   4365
            TabIndex        =   251
            Top             =   3765
            Width           =   3465
            Begin Threed.SSCheck chkSPTamizajeSalMPAT 
               Height          =   225
               Left            =   135
               TabIndex        =   252
               Top             =   195
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   397
               _Version        =   262144
               CaptionStyle    =   1
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "PATOLOGICO"
               Alignment       =   1
            End
            Begin Threed.SSCheck chkSPTamizajeSalMNOR 
               Height          =   225
               Left            =   1815
               TabIndex        =   253
               Top             =   165
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   397
               _Version        =   262144
               CaptionStyle    =   1
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "NORMAL"
               Alignment       =   1
            End
            Begin VB.Label Label72 
               AutoSize        =   -1  'True
               Caption         =   "(407)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3045
               TabIndex        =   254
               Top             =   165
               Width           =   315
            End
         End
         Begin VB.Frame Frame 
            Caption         =   "Gestante / RN / Niño / Adolesc, Etc..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3540
            Index           =   6
            Left            =   4350
            TabIndex        =   186
            Top             =   240
            Width           =   3495
            Begin VB.TextBox txtSPVacam 
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
               Left            =   2400
               TabIndex        =   225
               Top             =   3105
               Width           =   630
            End
            Begin VB.Frame fraEvalIntegral 
               Caption         =   "Evaluación Integral"
               Height          =   525
               Left            =   105
               TabIndex        =   221
               Top             =   2970
               Width           =   1605
               Begin Threed.SSCheck chkSPEvalIntegralSI 
                  Height          =   225
                  Left            =   90
                  TabIndex        =   222
                  Top             =   210
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Si"
                  Alignment       =   1
               End
               Begin Threed.SSCheck chkSPEvalIntegralNO 
                  Height          =   225
                  Left            =   600
                  TabIndex        =   223
                  Top             =   210
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "No"
                  Alignment       =   1
               End
               Begin VB.Label Label68 
                  AutoSize        =   -1  'True
                  Caption         =   "(401)"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   225
                  Left            =   1200
                  TabIndex        =   224
                  Top             =   240
                  Width           =   315
               End
            End
            Begin VB.TextBox txtSPIMC 
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
               Left            =   1575
               TabIndex        =   219
               Top             =   2625
               Width           =   885
            End
            Begin VB.TextBox txtSPNFamGestante 
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
               TabIndex        =   216
               Top             =   2625
               Width           =   885
            End
            Begin VB.Frame fraConIntegral 
               Caption         =   "Consejeria Integral"
               Height          =   525
               Left            =   1710
               TabIndex        =   212
               Top             =   1830
               Width           =   1710
               Begin Threed.SSCheck chkSPConIntegralSI 
                  Height          =   225
                  Left            =   120
                  TabIndex        =   213
                  Top             =   210
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Si"
                  Alignment       =   1
               End
               Begin Threed.SSCheck chkSPConIntegralNO 
                  Height          =   225
                  Left            =   600
                  TabIndex        =   214
                  Top             =   210
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "No"
                  Alignment       =   1
               End
               Begin VB.Label Label71 
                  AutoSize        =   -1  'True
                  Caption         =   "(013)"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   225
                  Left            =   1200
                  TabIndex        =   215
                  Top             =   240
                  Width           =   315
               End
            End
            Begin VB.Frame fraSecuelaNacer 
               Caption         =   "Secuela al Nacer"
               Height          =   525
               Left            =   120
               TabIndex        =   208
               ToolTipText     =   "Enfer.Congenita/Secuela al nacer"
               Top             =   1830
               Width           =   1575
               Begin Threed.SSCheck chkSPSecuelaNaceSI 
                  Height          =   225
                  Left            =   120
                  TabIndex        =   209
                  Top             =   210
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Si"
                  Alignment       =   1
               End
               Begin Threed.SSCheck chkSPSecuelaNaceNO 
                  Height          =   225
                  Left            =   600
                  TabIndex        =   210
                  Top             =   210
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "No"
                  Alignment       =   1
               End
               Begin VB.Label Label62 
                  AutoSize        =   -1  'True
                  Caption         =   "(021)"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   225
                  Left            =   1200
                  TabIndex        =   211
                  Top             =   240
                  Width           =   315
               End
            End
            Begin VB.Frame fraConsejNutricional 
               Caption         =   "Consej. Nutricional"
               Height          =   525
               Left            =   1710
               TabIndex        =   204
               ToolTipText     =   "Consejería Nutricional"
               Top             =   1320
               Width           =   1695
               Begin Threed.SSCheck chkSPconsejeriaNsi 
                  Height          =   225
                  Left            =   90
                  TabIndex        =   205
                  Top             =   210
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Si"
                  Alignment       =   1
               End
               Begin Threed.SSCheck chkSPconsejeriaNno 
                  Height          =   225
                  Left            =   600
                  TabIndex        =   206
                  Top             =   210
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "No"
                  Alignment       =   1
               End
               Begin VB.Label Label61 
                  AutoSize        =   -1  'True
                  Caption         =   "(307)"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   225
                  Left            =   1200
                  TabIndex        =   207
                  Top             =   240
                  Width           =   315
               End
            End
            Begin VB.Frame fraEEDP 
               Caption         =   "EEDP/TEPSI"
               Height          =   525
               Left            =   1710
               TabIndex        =   200
               Top             =   780
               Width           =   1695
               Begin Threed.SSCheck chkSPeedpSI 
                  Height          =   225
                  Left            =   90
                  TabIndex        =   201
                  Top             =   210
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Si"
                  Alignment       =   1
               End
               Begin Threed.SSCheck chkSPeedpNO 
                  Height          =   225
                  Left            =   600
                  TabIndex        =   202
                  Top             =   210
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "No"
                  Alignment       =   1
               End
               Begin VB.Label Label70 
                  AutoSize        =   -1  'True
                  Caption         =   "(312)"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   225
                  Left            =   1200
                  TabIndex        =   203
                  Top             =   240
                  Width           =   315
               End
            End
            Begin VB.Frame fraBajoPesoNacer 
               Caption         =   "Bajo Peso al Nacer"
               Height          =   525
               Left            =   120
               TabIndex        =   196
               Top             =   1320
               Width           =   1575
               Begin Threed.SSCheck chkSBajoPesoSI 
                  Height          =   225
                  Left            =   90
                  TabIndex        =   197
                  Top             =   210
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Si"
                  Alignment       =   1
               End
               Begin Threed.SSCheck chkSBajoPesoNO 
                  Height          =   225
                  Left            =   600
                  TabIndex        =   198
                  Top             =   210
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "No"
                  Alignment       =   1
               End
               Begin VB.Label Label64 
                  AutoSize        =   -1  'True
                  Caption         =   "(020)"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   225
                  Left            =   1200
                  TabIndex        =   199
                  Top             =   240
                  Width           =   315
               End
            End
            Begin VB.Frame fraRnPrematuro 
               Caption         =   "RN Prematuro"
               Height          =   525
               Left            =   120
               TabIndex        =   192
               Top             =   780
               Width           =   1575
               Begin Threed.SSCheck chkSPRNPrematuroSI 
                  Height          =   225
                  Left            =   120
                  TabIndex        =   193
                  Top             =   240
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Si"
                  Alignment       =   1
               End
               Begin Threed.SSCheck chkSPRNPrematuroNO 
                  Height          =   225
                  Left            =   600
                  TabIndex        =   194
                  Top             =   240
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "No"
                  Alignment       =   1
               End
               Begin VB.Label Label65 
                  AutoSize        =   -1  'True
                  Caption         =   "(019)"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   225
                  Left            =   1200
                  TabIndex        =   195
                  Top             =   240
                  Width           =   315
               End
            End
            Begin VB.TextBox txtSPPAB 
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
               TabIndex        =   190
               Top             =   480
               Width           =   825
            End
            Begin VB.TextBox txtSPcred 
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
               TabIndex        =   187
               Top             =   480
               Width           =   885
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "VACAM"
               Height          =   195
               Left            =   1830
               TabIndex        =   227
               Top             =   3150
               Width           =   525
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               Caption         =   "(018)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3105
               TabIndex        =   226
               Top             =   3105
               Width           =   315
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "(014)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   2520
               TabIndex        =   220
               Top             =   2640
               Width           =   315
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "NºFam.Gest/Puerp   IMC (Kg/m2)"
               Height          =   195
               Left            =   135
               TabIndex        =   218
               ToolTipText     =   "N° Familiares Gest/Puerp Casa Mat"
               Top             =   2385
               Width           =   2385
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "(404)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   1095
               TabIndex        =   217
               Top             =   2625
               Width           =   315
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "(015)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   2520
               TabIndex        =   191
               Top             =   480
               Width           =   315
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               Caption         =   "Cred (N°)                    PAB (cm)"
               Height          =   195
               Left            =   120
               TabIndex        =   189
               Top             =   240
               Width           =   2355
            End
            Begin VB.Label Label66 
               AutoSize        =   -1  'True
               Caption         =   "(120)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   1080
               TabIndex        =   188
               Top             =   480
               Width           =   315
            End
         End
         Begin VB.Frame Frame 
            Caption         =   "Del Recien Nacido"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1365
            Index           =   5
            Left            =   120
            TabIndex        =   176
            Top             =   2880
            Width           =   4215
            Begin VB.Frame fraCorTardio 
               Caption         =   "Corte Tardio de cordón (2a3 min)"
               Height          =   525
               Left            =   1680
               TabIndex        =   240
               Top             =   240
               Width           =   2415
               Begin Threed.SSCheck chkSPCorTarCordonSI 
                  Height          =   225
                  Left            =   120
                  TabIndex        =   241
                  Top             =   240
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Si"
                  Alignment       =   1
               End
               Begin Threed.SSCheck chkSPCorteTarCordonNO 
                  Height          =   225
                  Left            =   840
                  TabIndex        =   242
                  Top             =   240
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "No"
                  Alignment       =   1
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "(409)"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   225
                  Left            =   1560
                  TabIndex        =   243
                  Top             =   240
                  Width           =   315
               End
            End
            Begin VB.TextBox txtSPapgar5 
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
               Left            =   2400
               TabIndex        =   184
               Top             =   960
               Width           =   435
            End
            Begin VB.TextBox Text4 
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2280
               MultiLine       =   -1  'True
               TabIndex        =   182
               Text            =   "SISFormatoFUA.frx":45DF
               ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
               Top             =   960
               Width           =   135
            End
            Begin VB.TextBox txtSPapgar1 
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
               Left            =   1320
               TabIndex        =   181
               Top             =   960
               Width           =   435
            End
            Begin VB.TextBox txtSPedadGrn 
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
               TabIndex        =   177
               Top             =   480
               Width           =   765
            End
            Begin VB.Label Label3 
               Caption         =   "Edad Gest RN (Sem)  "
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   185
               Top             =   280
               Width           =   1575
            End
            Begin VB.Label Label60 
               AutoSize        =   -1  'True
               Caption         =   "(306)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   2880
               TabIndex        =   183
               Top             =   960
               Width           =   315
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               Caption         =   "(305)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   1800
               TabIndex        =   180
               Top             =   960
               Width           =   315
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               Caption         =   "A  P  G  A  R     1"
               Height          =   195
               Left            =   120
               TabIndex        =   179
               Top             =   960
               Width           =   1185
            End
            Begin VB.Label Label58 
               AutoSize        =   -1  'True
               Caption         =   "(304)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   960
               TabIndex        =   178
               Top             =   480
               Width           =   315
            End
         End
         Begin VB.Frame Frame 
            Caption         =   "De la Gestante"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Index           =   3
            Left            =   120
            TabIndex        =   165
            Top             =   1200
            Width           =   4215
            Begin VB.Frame frPartoVertical 
               Caption         =   "Parto Vertical"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   120
               TabIndex        =   244
               Top             =   960
               Width           =   1815
               Begin Threed.SSCheck chkSPPartoVertSI 
                  Height          =   225
                  Left            =   120
                  TabIndex        =   245
                  Top             =   285
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Si"
                  Alignment       =   1
               End
               Begin Threed.SSCheck chkSPPartoVertNO 
                  Height          =   225
                  Left            =   720
                  TabIndex        =   246
                  Top             =   280
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   397
                  _Version        =   262144
                  CaptionStyle    =   1
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "No"
                  Alignment       =   1
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "(408)"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   225
                  Left            =   1320
                  TabIndex        =   247
                  Top             =   285
                  Width           =   315
               End
            End
            Begin VB.TextBox txtSPpuerperio 
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
               Left            =   2040
               TabIndex        =   173
               Top             =   1200
               Width           =   825
            End
            Begin VB.TextBox txtSPalturaU 
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
               Left            =   2880
               TabIndex        =   171
               Top             =   600
               Width           =   825
            End
            Begin VB.TextBox txtSPedadG 
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
               Left            =   1440
               TabIndex        =   169
               Top             =   600
               Width           =   825
            End
            Begin VB.TextBox txtSPcpn 
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
               TabIndex        =   166
               Top             =   600
               Width           =   825
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   " Control de puerperio (N°)"
               Height          =   195
               Left            =   2040
               TabIndex        =   175
               Top             =   960
               Width           =   1875
            End
            Begin VB.Label Label67 
               AutoSize        =   -1  'True
               Caption         =   "(209)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   2880
               TabIndex        =   174
               Top             =   1200
               Width           =   315
            End
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
               Caption         =   "(010)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3720
               TabIndex        =   172
               Top             =   600
               Width           =   315
            End
            Begin VB.Label Label56 
               AutoSize        =   -1  'True
               Caption         =   "(005)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   2280
               TabIndex        =   170
               Top             =   600
               Width           =   315
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "CPN (N°)                Edad Gest (Sem)     Altu.Uterina (cm)"
               Height          =   195
               Left            =   120
               TabIndex        =   168
               Top             =   360
               Width           =   4020
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "(300)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   960
               TabIndex        =   167
               Top             =   600
               Width           =   315
            End
         End
         Begin VB.Frame Frame 
            Caption         =   "Otras Actividades"
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
            Index           =   0
            Left            =   120
            TabIndex        =   159
            Top             =   240
            Width           =   4215
            Begin VB.TextBox txtSPpeso 
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
               TabIndex        =   161
               Top             =   510
               Width           =   615
            End
            Begin VB.TextBox txtSPtalla 
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
               Left            =   1200
               TabIndex        =   160
               ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
               Top             =   510
               Width           =   645
            End
            Begin MSMask.MaskEdBox txtSPpa 
               Height          =   315
               Left            =   2400
               TabIndex        =   162
               Top             =   510
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "###/###"
               PromptChar      =   "_"
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               Caption         =   "(004)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   1900
               TabIndex        =   228
               Top             =   480
               Width           =   315
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Peso  (Kg)        Talla (cm)           P.A.(mmHg)"
               Height          =   195
               Left            =   120
               TabIndex        =   164
               Top             =   240
               Width           =   3150
            End
            Begin VB.Label Label63 
               AutoSize        =   -1  'True
               Caption         =   "(901) / (301)"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3360
               TabIndex        =   163
               Top             =   240
               Width           =   765
            End
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Vacunas N° Dosis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4305
         Left            =   -66955
         TabIndex        =   29
         Top             =   330
         Width           =   4380
         Begin VB.TextBox txtVacOtraVacuna 
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
            Left            =   3090
            TabIndex        =   238
            Top             =   3255
            Width           =   645
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   3495
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   237
            Text            =   "SISFormatoFUA.frx":45E3
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtVacSR 
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
            TabIndex        =   232
            Top             =   3255
            Width           =   645
         End
         Begin VB.TextBox txtVacIPV 
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
            TabIndex        =   231
            Top             =   3240
            Width           =   645
         End
         Begin VB.TextBox txtVacVPH 
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
            Left            =   3090
            TabIndex        =   229
            Top             =   2715
            Width           =   645
         End
         Begin VB.TextBox txtVacRiesgoHVB 
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
            Left            =   3090
            MaxLength       =   1
            TabIndex        =   43
            Top             =   3795
            Width           =   645
         End
         Begin VB.TextBox txtVacPentaval 
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
            TabIndex        =   47
            Top             =   3795
            Width           =   645
         End
         Begin VB.TextBox txtVacHVB 
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
            Left            =   180
            TabIndex        =   46
            Top             =   3795
            Width           =   645
         End
         Begin VB.TextBox txtVacDt 
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
            Left            =   1530
            TabIndex        =   45
            Top             =   2715
            Width           =   645
         End
         Begin VB.TextBox txtVacSpr 
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
            TabIndex        =   44
            Top             =   2715
            Width           =   645
         End
         Begin VB.TextBox txtVacRotavirus 
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
            Left            =   1530
            TabIndex        =   42
            Top             =   2175
            Width           =   645
         End
         Begin VB.TextBox txtVacAsa 
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
            TabIndex        =   41
            Top             =   2175
            Width           =   645
         End
         Begin VB.TextBox txtVacAntitetanica 
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
            Left            =   3090
            TabIndex        =   40
            Top             =   1635
            Width           =   645
         End
         Begin VB.TextBox txtVacRubeola 
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
            Left            =   1530
            TabIndex        =   39
            Top             =   1635
            Width           =   645
         End
         Begin VB.TextBox txtVacApo 
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
            TabIndex        =   38
            Top             =   1635
            Width           =   645
         End
         Begin VB.TextBox txtVacAntineumoc 
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
            Left            =   3090
            TabIndex        =   37
            Top             =   1095
            Width           =   645
         End
         Begin VB.TextBox txtVacParotid 
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
            Left            =   1530
            TabIndex        =   36
            Top             =   1095
            Width           =   645
         End
         Begin VB.TextBox txtVacDpt 
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
            TabIndex        =   35
            Top             =   1095
            Width           =   645
         End
         Begin VB.TextBox txtVacAntiamarilica 
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
            TabIndex        =   34
            Top             =   555
            Width           =   645
         End
         Begin VB.TextBox txtVacInfluenz 
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
            Left            =   1530
            TabIndex        =   33
            Top             =   555
            Width           =   645
         End
         Begin VB.TextBox txtVacBcg 
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
            TabIndex        =   32
            Top             =   555
            Width           =   645
         End
         Begin Threed.SSCheck chkVacCompEdSI 
            Height          =   225
            Left            =   2880
            TabIndex        =   235
            Top             =   2205
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   397
            _Version        =   262144
            CaptionStyle    =   1
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Si"
            Alignment       =   1
         End
         Begin Threed.SSCheck chkVacCompEdNo 
            Height          =   225
            Left            =   3360
            TabIndex        =   236
            Top             =   2205
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   397
            _Version        =   262144
            CaptionStyle    =   1
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "No"
            Alignment       =   1
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "(XXX)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   3960
            TabIndex        =   248
            Top             =   2160
            Width           =   360
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "(XXX)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   3840
            TabIndex        =   239
            Top             =   3240
            Width           =   360
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "SR                           IPV                             OTRA VACUNA"
            Height          =   195
            Left            =   120
            TabIndex        =   234
            Top             =   3060
            Width           =   4035
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "(316)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2280
            TabIndex        =   233
            Top             =   3240
            Width           =   315
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "(319)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   3810
            TabIndex        =   230
            Top             =   2760
            Width           =   315
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "(406)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   3840
            TabIndex        =   85
            Top             =   3840
            Width           =   315
         End
         Begin VB.Label Label87 
            AutoSize        =   -1  'True
            Caption         =   "(124)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2310
            TabIndex        =   80
            Top             =   3840
            Width           =   315
         End
         Begin VB.Label Label85 
            AutoSize        =   -1  'True
            Caption         =   "(007)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2280
            TabIndex        =   79
            Top             =   2760
            Width           =   315
         End
         Begin VB.Label Label84 
            AutoSize        =   -1  'True
            Caption         =   "(127)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2280
            TabIndex        =   78
            Top             =   2160
            Width           =   315
         End
         Begin VB.Label Label81 
            AutoSize        =   -1  'True
            Caption         =   "(208)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   3840
            TabIndex        =   77
            Top             =   1680
            Width           =   315
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "(122)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2280
            TabIndex        =   76
            Top             =   1680
            Width           =   315
         End
         Begin VB.Label Label78 
            AutoSize        =   -1  'True
            Caption         =   "(126)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   3840
            TabIndex        =   75
            Top             =   1080
            Width           =   315
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            Caption         =   "(121)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2280
            TabIndex        =   74
            Top             =   1080
            Width           =   315
         End
         Begin VB.Label Label75 
            AutoSize        =   -1  'True
            Caption         =   "(211)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   3840
            TabIndex        =   73
            Top             =   600
            Width           =   315
         End
         Begin VB.Label Label74 
            AutoSize        =   -1  'True
            Caption         =   "(318)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2280
            TabIndex        =   72
            Top             =   600
            Width           =   315
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "HVB                        PENTAVAL           GRUPO RIESGO HVB"
            Height          =   195
            Left            =   180
            TabIndex        =   54
            Top             =   3600
            Width           =   4080
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "SPR                      DT ADULTO(N°Dosis)    VPH"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   2520
            Width           =   3255
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "ASA                        ROTAVIRUS        COMPLETAS P. EDAD"
            Height          =   195
            Left            =   150
            TabIndex        =   52
            Top             =   1980
            Width           =   4140
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "APO                        RUBEOLA                   ANTITETANICA"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   1440
            Width           =   4050
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "DPT                         PAROTID                   ANTINEUMOC"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   900
            Width           =   3945
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "BCG                         INFLUENZ                  ANTIAMARILICA"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   4155
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Se Refiere / Contrarefiere A:"
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
         Left            =   -71040
         TabIndex        =   26
         Top             =   6480
         Width           =   8475
         Begin VB.CommandButton btnBuscarEstablecimientoD 
            Caption         =   "..."
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
            Left            =   1320
            TabIndex        =   2
            Top             =   510
            Width           =   315
         End
         Begin VB.TextBox txtRDcodigo 
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
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   510
            Width           =   1185
         End
         Begin VB.TextBox txtRD 
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
            TabIndex        =   27
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   510
            Width           =   4305
         End
         Begin VB.TextBox txtRDnumero 
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
            Left            =   6000
            MaxLength       =   20
            TabIndex        =   3
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   510
            Width           =   1665
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Cod.ES/Eq            IPRESS/AISPED al que se refiere al Paciente            N° Hoja Refer/Cont"
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
            TabIndex        =   28
            Top             =   270
            Width           =   7500
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Otros datos de Salida"
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
         Left            =   -74940
         TabIndex        =   24
         Top             =   6480
         Width           =   3825
         Begin VB.ComboBox cmbIdDestinoAtencion 
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
            Left            =   120
            TabIndex        =   1
            Top             =   510
            Width           =   3615
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Destino del asegurado"
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
            Top             =   270
            Width           =   1815
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "De la Institución Prestadora de Servicios de Salud (IPRESS)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   -74940
         TabIndex        =   14
         Top             =   390
         Width           =   12345
         Begin VB.Frame Frame23 
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
            Height          =   900
            Left            =   7920
            TabIndex        =   127
            Top             =   1150
            Width           =   4305
            Begin Threed.SSCheck chkAtencionAmbulatoria 
               Height          =   255
               Left            =   120
               TabIndex        =   128
               Top             =   360
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   262144
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Ambulatoria"
               Value           =   1
            End
            Begin Threed.SSCheck chkAtencionReferencia 
               Height          =   255
               Left            =   1440
               TabIndex        =   129
               Top             =   360
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   450
               _Version        =   262144
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Referencia"
            End
            Begin Threed.SSCheck chkAtencionEmergencia 
               Height          =   285
               Left            =   2640
               TabIndex        =   130
               Top             =   360
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               _Version        =   262144
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Emergencia"
            End
         End
         Begin VB.Frame FarLugarAtencion 
            Caption         =   "Lugar de Atención"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   7920
            TabIndex        =   124
            Top             =   600
            Width           =   4305
            Begin Threed.SSCheck chkIntramural 
               Height          =   255
               Left            =   120
               TabIndex        =   125
               Top             =   240
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   450
               _Version        =   262144
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Intramural"
               Value           =   1
            End
            Begin Threed.SSCheck chkExtramural 
               Height          =   255
               Left            =   1560
               TabIndex        =   126
               Top             =   240
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   450
               _Version        =   262144
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Extramural"
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Referencia Realizada Por:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   120
            TabIndex        =   118
            Top             =   1150
            Width           =   7695
            Begin VB.CommandButton btnBuscarEstablecimientoO 
               Caption         =   "..."
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
               Left            =   1320
               TabIndex        =   122
               Top             =   480
               Width           =   315
            End
            Begin VB.TextBox txtRONumero 
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
               Left            =   5880
               MaxLength       =   20
               TabIndex        =   121
               ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtRO 
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
               TabIndex        =   120
               ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
               Top             =   480
               Width           =   4185
            End
            Begin VB.TextBox txtROcodigo 
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
               TabIndex        =   119
               ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
               Top             =   480
               Width           =   1155
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Cod.ES/Eq            IPRESS/AISPED que Refirió al Paciente                  N° Hoja Referencia "
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
               TabIndex        =   123
               Top             =   240
               Width           =   7425
            End
         End
         Begin VB.Frame FraPersonal 
            Caption         =   "Personal que atiende"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            TabIndex        =   114
            Top             =   600
            Width           =   7695
            Begin VB.TextBox txtPACodOfFlexible 
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
               Left            =   5160
               MaxLength       =   20
               TabIndex        =   146
               Top             =   180
               Width           =   2415
            End
            Begin Threed.SSCheck chkPAestablecimiento 
               Height          =   255
               Left            =   120
               TabIndex        =   115
               Top             =   240
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   450
               _Version        =   262144
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "De la IPRESS"
               Value           =   1
            End
            Begin Threed.SSCheck chkPAaisped 
               Height          =   315
               Left            =   1560
               TabIndex        =   116
               Top             =   240
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               _Version        =   262144
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Itinerante"
            End
            Begin Threed.SSCheck chkPAOfeFlexible 
               Height          =   315
               Left            =   2880
               TabIndex        =   117
               Top             =   240
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               _Version        =   262144
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Oferta Flex."
            End
            Begin VB.Label lblPACodigoOfFlexible 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Código"
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
               Left            =   4440
               TabIndex        =   147
               Top             =   270
               Width           =   555
            End
         End
         Begin VB.TextBox txtCS 
            Height          =   315
            Left            =   4020
            TabIndex        =   17
            Top             =   240
            Width           =   8175
         End
         Begin VB.TextBox txtCScodigo 
            Height          =   315
            Left            =   2370
            TabIndex        =   15
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "IPRESS/Equipo AISPED"
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
            Left            =   150
            TabIndex        =   16
            Top             =   270
            Width           =   1890
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos de la atención"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   -74940
         TabIndex        =   13
         Top             =   4680
         Width           =   12375
         Begin VB.Frame Frame 
            Caption         =   "Reporte Vinculado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   4
            Left            =   120
            TabIndex        =   154
            Top             =   840
            Width           =   4770
            Begin VB.TextBox txtFuaVincular 
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
               Left            =   2280
               MaxLength       =   50
               TabIndex        =   156
               ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
               Top             =   480
               Width           =   2400
            End
            Begin VB.TextBox txtCodAutorizacion 
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
               MaxLength       =   50
               TabIndex        =   155
               ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Cod. Autorización                     Nº Fua Vincular"
               Height          =   195
               Left            =   120
               TabIndex        =   157
               Top             =   240
               Width           =   3315
            End
         End
         Begin VB.TextBox lblUpsFua 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
            Height          =   615
            Left            =   3480
            MultiLine       =   -1  'True
            TabIndex        =   153
            Text            =   "SISFormatoFUA.frx":4681
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtCodPrestAdicional 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   4920
            MultiLine       =   -1  'True
            TabIndex        =   143
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   480
            Width           =   1515
         End
         Begin VB.Frame Frame18 
            Caption         =   "Hospitalización"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   9600
            TabIndex        =   136
            Top             =   240
            Width           =   2685
            Begin MSMask.MaskEdBox txtHfingreso 
               Height          =   315
               Left            =   1280
               TabIndex        =   137
               Top             =   240
               Width           =   1365
               _ExtentX        =   2408
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
            Begin MSMask.MaskEdBox txtHfalta 
               Height          =   315
               Left            =   1280
               TabIndex        =   138
               Top             =   600
               Width           =   1365
               _ExtentX        =   2408
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
            Begin MSMask.MaskEdBox txtHFCortAdmin 
               Height          =   315
               Left            =   1280
               TabIndex        =   141
               Top             =   960
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   13
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
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "F. Corte Adm."
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
               TabIndex        =   142
               Top             =   960
               Width           =   1155
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "F. Alta"
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
               TabIndex        =   140
               Top             =   600
               Width           =   885
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "F.Ingreso"
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
               TabIndex        =   139
               Top             =   270
               Width           =   885
            End
         End
         Begin VB.Frame fraConcPrestacional 
            Caption         =   "Concepto Prestacional"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   6480
            TabIndex        =   131
            Top             =   240
            Width           =   3105
            Begin VB.TextBox txtNautorizacion 
               Height          =   315
               Left            =   120
               MaxLength       =   15
               TabIndex        =   134
               Top             =   960
               Width           =   1905
            End
            Begin VB.TextBox txtMonto 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2040
               TabIndex        =   133
               Top             =   960
               Width           =   1005
            End
            Begin VB.ComboBox cmbConceptoP 
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
               Left            =   120
               TabIndex        =   132
               Top             =   270
               Width           =   2955
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "N° Autorización           Monto"
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
               TabIndex        =   135
               Top             =   690
               Width           =   2445
            End
         End
         Begin MSMask.MaskEdBox txtFantencion 
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Top             =   480
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
         Begin MSMask.MaskEdBox txtHatencion 
            Height          =   315
            Left            =   1440
            TabIndex        =   23
            Top             =   480
            Width           =   705
            _ExtentX        =   1244
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
         Begin PVCOMBOLibCtl.PVComboBox cmbUPSfua 
            Height          =   330
            Left            =   2160
            TabIndex        =   145
            Top             =   480
            Width           =   1185
            _Version        =   524288
            _cx             =   2090
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
               Size            =   9
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
            Column0.Heading =   "Descripción"
            Column0.Width   =   200
            Column0.Alignment=   0
            Column0.Hidden  =   0   'False
            Column0.Name    =   "descripcion"
            Column0.Format  =   ""
            Column0.Bound   =   -1  'True
            Column0.Locked  =   0   'False
            Column0.HeaderAlignment=   0
            Column1.Heading =   "Código"
            Column1.Width   =   60
            Column1.Alignment=   0
            Column1.Hidden  =   0   'False
            Column1.Name    =   "ups"
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
         Begin VB.Label lblCodPrestAdicional 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Prestac. Adicional"
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
            Left            =   4920
            TabIndex        =   158
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de la Atención      UPS"
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
            TabIndex        =   22
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos del Asegurado / Usuario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   -74940
         TabIndex        =   12
         Top             =   2520
         Width           =   12315
         Begin VB.Frame Frame 
            Caption         =   "DNI / CNV / AFILIACION DEL RN"
            Height          =   1215
            Index           =   2
            Left            =   8160
            TabIndex        =   112
            Top             =   840
            Width           =   4095
            Begin UltraGrid.SSUltraGrid grdRN 
               Height          =   930
               Left            =   0
               TabIndex        =   113
               Top             =   240
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   1640
               _Version        =   131072
               GridFlags       =   17040384
               LayoutFlags     =   67108884
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "grdRN"
            End
         End
         Begin VB.Frame Frame 
            Caption         =   "Fecha Parto        F:Nacimiento      F.Fallecimiento    "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   3720
            TabIndex        =   108
            Top             =   1440
            Width           =   4455
            Begin MSMask.MaskEdBox txtFnacimiento 
               Height          =   315
               Left            =   1560
               TabIndex        =   109
               Top             =   240
               Width           =   1455
               _ExtentX        =   2566
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
            Begin MSMask.MaskEdBox txtFFallecimiento 
               Height          =   315
               Left            =   3000
               TabIndex        =   110
               Top             =   240
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
            Begin MSMask.MaskEdBox txtFparto 
               Height          =   315
               Left            =   120
               TabIndex        =   111
               Top             =   240
               Width           =   1455
               _ExtentX        =   2566
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
         End
         Begin VB.TextBox txtNhistoriaClinica 
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
            Left            =   5040
            TabIndex        =   106
            Top             =   510
            Width           =   1455
         End
         Begin VB.TextBox txtSexo 
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
            Left            =   5040
            TabIndex        =   105
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   1050
            Width           =   1365
         End
         Begin VB.CommandButton btnBuscarPaciente 
            Caption         =   "..."
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
            Left            =   4680
            TabIndex        =   104
            ToolTipText     =   "Busca por Apellidos y Nombres"
            Top             =   1050
            Width           =   315
         End
         Begin VB.TextBox txtPaciente 
            Height          =   315
            Left            =   120
            TabIndex        =   103
            Top             =   1050
            Width           =   4515
         End
         Begin VB.TextBox txtNroAfiliacion3 
            Height          =   315
            Left            =   3720
            TabIndex        =   101
            Top             =   510
            Width           =   1305
         End
         Begin VB.TextBox txtNroAfiliacion2 
            Height          =   315
            Left            =   3330
            TabIndex        =   100
            Top             =   510
            Width           =   375
         End
         Begin VB.TextBox txtNroAfiliacion1 
            Height          =   315
            Left            =   2880
            TabIndex        =   99
            Top             =   510
            Width           =   435
         End
         Begin VB.TextBox txtNdocumento 
            Height          =   315
            Left            =   1200
            TabIndex        =   98
            Top             =   510
            Width           =   1635
         End
         Begin VB.ComboBox cmbTipoDocumento 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   97
            Top             =   510
            Width           =   1125
         End
         Begin VB.Frame fraGestantePuerpera 
            Caption         =   "Salud Materna"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   94
            Top             =   1440
            Width           =   3615
            Begin Threed.SSCheck chkGestante 
               Height          =   255
               Left            =   120
               TabIndex        =   95
               Top             =   240
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   450
               _Version        =   262144
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Gestante"
            End
            Begin Threed.SSCheck chkPuerpera 
               Height          =   255
               Left            =   1680
               TabIndex        =   96
               Top             =   240
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   450
               _Version        =   262144
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Puerpera"
            End
         End
         Begin VB.Frame fraCodAfiliacionSeguro 
            Caption         =   "Código de Afiliación de Seguro"
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
            Height          =   735
            Left            =   6600
            TabIndex        =   18
            Top             =   120
            Width           =   5685
            Begin VB.TextBox txtCodSeguro 
               Height          =   315
               Left            =   4320
               TabIndex        =   5
               Top             =   360
               Width           =   1305
            End
            Begin VB.TextBox txtInstitucion 
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1080
               TabIndex        =   4
               Text            =   "0"
               Top             =   360
               Width           =   2115
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Cód.Seguro"
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
               Left            =   3240
               TabIndex        =   20
               Top             =   360
               Width           =   1080
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Institución"
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
               Left            =   120
               TabIndex        =   19
               Top             =   360
               Width           =   975
            End
         End
         Begin PVCOMBOLibCtl.PVComboBox cmbEtnia 
            Height          =   330
            Left            =   6420
            TabIndex        =   144
            Top             =   1050
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Apellidos y Nombres del Asegurado                                  Sexo                  Etnia"
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
            TabIndex        =   107
            Top             =   840
            Width           =   6795
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Identificación                            Cod.Afiliación/Inscripción     N° Historia      "
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
            TabIndex        =   102
            Top             =   240
            Width           =   6270
         End
      End
      Begin VB.CommandButton btnRefrescar 
         Caption         =   "Refrescar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74850
         TabIndex        =   11
         Top             =   8970
         Width           =   1515
      End
      Begin UltraGrid.SSUltraGrid grdDiag 
         Height          =   2145
         Left            =   8295
         TabIndex        =   257
         Top             =   5385
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   3784
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   71303188
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   "SISFormatoFUA.frx":46AA
         Caption         =   "grdDiag"
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "F7 = Busca MEDICAMENTO/INSUMO                                            F8 = Busca procedimiento CPT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   150
         TabIndex        =   256
         Top             =   5550
         Width           =   7485
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   $"SISFormatoFUA.frx":46E6
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74760
         TabIndex        =   60
         Top             =   6960
         Width           =   9990
      End
   End
   Begin VB.Label lblCtaEmergencia 
      Caption         =   "...."
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
      Height          =   195
      Left            =   120
      TabIndex        =   255
      Top             =   1470
      Width           =   12405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número de Formato FUA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   120
      TabIndex        =   71
      Top             =   90
      Width           =   2715
   End
End
Attribute VB_Name = "SisFuaVersion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: MINSA - Oficina De Informatica y Telecomunicaciones
'        Aplicativo: SisGalenPlus v.3
'        Programa: Registra y emite formato FUA - Actualizado para los nuevos fuas
'        Programado por: franklin cachay
'        Fecha: agosto 2015
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim oRsRayosX As New Recordset, lnIdRecetaRayosX As Long
Dim oRsPatologia As New Recordset, lnIdRecetaPatología As Long
Dim oRsFarmacia As New Recordset, lnIdRecetaFarmacia As Long
Dim oRsDx As New Recordset, oRsVacunasSp As New Recordset
Dim oRsNacimientos As New Recordset
Dim mo_cmbIdDestinoAtencion As New sighentidades.ListaDespleglable
Dim mo_cmbConceptoP As New sighentidades.ListaDespleglable
Dim mo_cmbTipoDocumento As New sighentidades.ListaDespleglable
Dim mo_cmbColegioNivel As New sighentidades.ListaDespleglable
Dim mo_cmbColegioGrado As New sighentidades.ListaDespleglable
Dim mo_cmbColegioTurno As New sighentidades.ListaDespleglable
Dim mo_cmbFuaUps As New sighentidades.ListaDespleglable
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim oDoSisFuaAtencion As New DoSisFuaAtencion
Dim oDoServicio As New doServicio
Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim lcSql As String
Dim ml_idUsuario As Long
Dim mi_opcion As sghOpciones, mi_opcion_fua As sghOpciones
Dim ml_IdCuentaAtencion As Long
Dim mo_lnIdTablaLISTBARITEMS As sghOpcionGalenHos
Dim mo_lcNombrePc As String
Dim ml_Paciente As String
Dim ml_NroDocumento As String
Dim ml_TipoDocumentoGalenhos As Long
Dim md_FechaNacimiento As Date
Dim ml_Sexo As String
Dim ml_Etnia As String
Dim ml_NroHistoriaClinica As Long
Dim ml_ApellidoPaterno As String
Dim ml_ApellidoMaterno As String
Dim ml_PrimerNombre As String
Dim ml_SegundoNombre As String
Dim ml_edad_En_Dias As Long, ml_edad_En_YYYYMMDD As String
Dim ml_IdMedico As String
Dim md_FechaAtencion As Date
Dim ml_HoraAtencion As String
Dim ml_IdTipoServicio As sghTipoServicio
Dim ml_IdDestinoPaciente As String
Dim ml_IdConceptoPrestacional As String
Dim lcCodigoDxBuscado As String
Dim lcDxPrincipal As String, lcDxPrincipalNro As Long
Dim lcEquix As String, mo_lbCargaTablasUnaVez As Boolean
Dim lcAfiliacionNroIntegrante As String, lcAfiliacionCodigo As String, lcAfiliacionIdSiaSis As String
Dim ml_idAtencion As Long, lcOpcion As String, lcNivelEstablecimiento As String
Dim lbEsIgualQueArSIS As Boolean, lcElServicioUsaGalenHos As String
Dim lcCodigoEstablecimientoAdscripcionSIS As String
Const lcInsumo As String = "Insumo": Const lcMedicamento As String = "Medicamento"
Const lcOtros As String = "Otros": Const lcLaboratorio As String = "Laboratorio"
Const lcImagenes As String = "Imágenes"
Const lcGalenHosVersion As String = "v.3": Const lcGalenHosNombre As String = "1000"
Const lcVacio As String = "(VACIO)": Const lcFarmacia As String = "Farmacia"
Const lcMasculino As String = "Masculino": Const lcFemenino As String = "Femenino"
Const lcMinimoCuentaARFSIS As String = "999999999"
Dim wxParametro205 As String, wxParametro242 As String, wxParametro280 As String
Dim wxParametro303 As String, wxParametro304 As String, wxParametro305 As String
Dim wxParametro306 As String, wxParametro310 As String, wxParametro320 As String
Dim wxParametro327 As String, wxParametro328 As String, wxParametroJAMO As String
Dim wxParametroSIS  As String, wxParametro302 As String, wxParametro338 As String
Dim wxParametro359 As String, wxParametro553 As String, wxParametro554 As String
Dim lnNroFuaRepetido As Boolean
Dim mo_lbEsAltaMedica As Boolean
Dim mc_FuaVersionFormato As String
Dim mi_FuaTipoAnexo2015 As Integer
Dim oCampos() As String
Const lcFuaAnexo1 As Integer = 1: Const lcFuaAnexo2 As Integer = 2
Dim lnIdCuentaAtencionEmergenciaOce As Long, lnIdAtencionEmergenciaOce As Long
Dim ml_idPaciente As Long, ml_fechaIngreso As Date, ml_IdOrigenAtencion As Long
Dim ml_EsPacienteExterno As Boolean
Dim lnIdDiagnosticoPacExtSeguro As Long
Dim ldHoy As Date
Dim mo_SoloImprimeFUAyaGrabado As Boolean
Dim ml_CodigoPrestacion As String

'HRA 10/12/2020 Cambio Inicio
Dim mc_GuardarFua As String
Property Let GuardarFua(lValue As String)
   mc_GuardarFua = lValue
End Property
'HRA 10/12/2020 Cambio Fin

Property Let CodigoPrestacion(lValue As String)
    ml_CodigoPrestacion = lValue
End Property

Property Let SoloImprimeFUAyaGrabado(lValue As Boolean)
   mo_SoloImprimeFUAyaGrabado = lValue
End Property

Property Let EsAltaMedica(lValue As Boolean)
   mo_lbEsAltaMedica = lValue
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As sghOpcionGalenHos)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Let IdCuentaAtencion(lValue As Long)
   ml_IdCuentaAtencion = lValue
End Property

Property Let Opcion(lValue As sghOpciones)
    mi_opcion = lValue
    mi_opcion_fua = mi_opcion
End Property

Property Let FuaVersionFormato(lValue As String)
   mc_FuaVersionFormato = lValue
End Property

Property Let FuaTipoAnexo2015(lValue As Integer)
   mi_FuaTipoAnexo2015 = lValue
End Property



Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
            btnCancelar_Click
       Case vbKeyF1
            If Me.grdDx.Enabled = False Then
               Exit Sub
            End If
            If lcCodigoDxBuscado <> "" Then
                Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
                Dim oDODiagnostico As DODiagnostico
                oBusqueda.SoloMuestraDxGalenHos = False
                oBusqueda.USAcodigoCIEsinPto = True
                oBusqueda.MostrarFormulario
                If oBusqueda.BotonPresionado = sghAceptar Then
                    Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
                    If Not oDODiagnostico Is Nothing Then
                        If lcCodigoDxBuscado = "DxEgreso" Then
                           oRsDx.Fields!DxEgreso = oDODiagnostico.codigoCIEsinPto
                        Else
                           oRsDx.Fields!dxIngreso = oDODiagnostico.codigoCIEsinPto
                        End If
                        oRsDx.Fields!Descripcion = oDODiagnostico.Descripcion
                    End If
                End If
                Set oBusqueda = Nothing
            End If
       Case vbKeyF2
            btnAceptar_Click
       Case vbKeyF3
            Me.TabFua.Tab = 0
            On Error Resume Next
            chkPAestablecimiento.SetFocus
       Case vbKeyF4
            Me.TabFua.Tab = 1
            On Error Resume Next
            txtSPcpn.SetFocus
       Case vbKeyF5
            Me.TabFua.Tab = 2
            On Error Resume Next
            grdFarmacia.SetFocus
       Case vbKeyF7
            Me.TabFua.Tab = 2
            On Error Resume Next
            btnAddFarmacia_Click
       Case vbKeyF8
            Me.TabFua.Tab = 2
            On Error Resume Next
            btnAddPatologia_Click
       End Select
End Sub

Sub ArsSisCargaDatosAntesDeGrabar()
    If lbEsIgualQueArSIS = True Then
       ml_IdCuentaAtencion = Left("9" & Trim(txtFua1.Text) & Trim(txtFua2.Text) & Trim(Val(txtFua3.Text)) & lcMinimoCuentaARFSIS, 9)
       'ml_idCuentaAtencion = 999999999
    End If
   
End Sub

Public Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_opcion
   Case sghAgregar
       ArsSisCargaDatosAntesDeGrabar
       If ValidacionesOK = True Then
            If AgregarDatos() Then
                 If MsgBox("Los datos se agregaron correctamente" & Chr(13) & "Desea Imprimir el Formato", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                    ImpresionFua
                 End If
                 Me.Visible = False
            Else
                MsgBox "No se pudo agregar los datos" + Chr(13) + mo_ReglasSISgalenhos.MensajeError, vbExclamation, Me.Caption
                If lnNroFuaRepetido = True Then
'                    If Val(wxParametro320) = sghFuaTipo.sghFuaTipoAutomatico Then
'                       txtFua3.Text = Right("00000000" & Trim(Str(Val(mo_ReglasSISgalenhos.sisFuaAtencionUltimoCorrelativo()) + 1)), 8)
'                    Else
                    Dim oRsTmp1 As New Recordset
                    Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaSeleccionarTodos
                    If oRsTmp1.RecordCount > 0 Then
                       txtFua3.Text = Right("00000000" & Trim(Str(Val(oRsTmp1.Fields!FuaUltimoGenerado) + 1)), 8)
                    End If
                    oRsTmp1.Close
                    Set oRsTmp1 = Nothing
'                    End If
                End If
            End If
       End If
   Case sghModificar
       If ValidacionesOK = True Then
            If ModificarDatos() Then
                 If MsgBox("Los datos se modificaron correctamente" & Chr(13) & "Desea Imprimir el Formato", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                    ImpresionFua
                 End If
                 Me.Visible = False
            Else
                 MsgBox "No se pudo modificar los datos" + Chr(13) + mo_ReglasSISgalenhos.MensajeError, vbExclamation, Me.Caption
            End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Function ValidarDatosObligatorios() As Boolean
    Dim lbEstaOk As Boolean, lcTexto As String
    
    ValidarDatosObligatorios = False
    If Val(txtFua3.Text) = 0 Then 'And Val(wxParametro320) = sghFuaTipo.sghFuaTipoManual Then
       MsgBox "Debe registrar el NUMERO FUA", vbInformation, Me.Caption
       If txtFua3.Enabled Then
          txtFua3.SetFocus
       End If
       Exit Function
    End If
    'mgaray20140926
    If Me.ucSISfuaCodPrestacion1.Prestacion = "" And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghRegistroCitaCE _
                And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghAdmisionEmergencia _
                And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghPacienteExternoConSeguro Then
       MsgBox "Debe elegir CODIGO DE PRESTACION", vbInformation, Me.Caption
       On Error Resume Next
       ucSISfuaCodPrestacion1.SetFocus
       Exit Function
    End If
    If mi_FuaTipoAnexo2015 = lcFuaAnexo1 And Trim(txtColegioCodigo.Text) <> "" Then
        If cmbColegioNivel.Text = "" Then
            MsgBox "Debe elegir el NIVEL del estudiante en la Institución Educativa", vbInformation, Me.Caption
            Me.TabFua.Tab = 0
            On Error Resume Next
            cmbColegioNivel.SetFocus
            Exit Function
        End If
        If cmbColegioGrado.Text = "" Then
            MsgBox "Debe elegir el GRADO del estudiante en la Institución Educativa", vbInformation, Me.Caption
            Me.TabFua.Tab = 0
            On Error Resume Next
            cmbColegioGrado.SetFocus
            Exit Function
        End If
        If cmbColegioTurno.Text = "" Then
            MsgBox "Debe elegir el TURNO del estudiante en la Institución Educativa", vbInformation, Me.Caption
            Me.TabFua.Tab = 0
            On Error Resume Next
            cmbColegioTurno.SetFocus
            Exit Function
        End If
    End If
    
    If txtNroAfiliacion1.Text = "" Or txtNroAfiliacion2.Text = "" Or txtNroAfiliacion3.Text = "" Then
       MsgBox "Debe registrar el COD.AFILIACION/INSCRIPCION", vbInformation, Me.Caption
       Me.TabFua.Tab = 0
       On Error Resume Next
       txtNroAfiliacion1.SetFocus
       Exit Function
    Else
       Dim oRsTmp1 As New Recordset
       Set oRsTmp1 = mo_ReglasSISgalenhos.SisFiltraPacientesAfiliados("Where afiliacionDisa='" & txtNroAfiliacion1.Text & _
                                                        "' and afiliacionTipoFormato='" & txtNroAfiliacion2.Text & _
                                                        "' and afiliacionNroFormato='" & txtNroAfiliacion3.Text & _
                                                        "'", wxParametroJAMO)
       If oRsTmp1.RecordCount = 0 Then
          oRsTmp1.Close
          Set oRsTmp1 = Nothing
          MsgBox "El COD.AFILIACION/INSCRIPCION no existe en la tabla FILIACIONES SIS", vbInformation, Me.Caption
          Me.TabFua.Tab = 0
          On Error Resume Next
          txtNroAfiliacion1.SetFocus
          Exit Function
       Else
          oRsTmp1.Close
          Set oRsTmp1 = Nothing
       End If
    End If
    If Me.txtNhistoriaClinica.Text = "" Then
       MsgBox "Debe registrar el Nro HISTORIA", vbInformation, Me.Caption
       Me.TabFua.Tab = 0
       On Error Resume Next
       txtNhistoriaClinica.SetFocus
       Exit Function
    End If
    'mgaray20140926
    If (chkAtencionAmbulatoria.Value = 0 And chkAtencionReferencia.Value = 0 And chkAtencionEmergencia.Value = 0) _
                And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghRegistroCitaCE _
                And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghAdmisionEmergencia _
                And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghPacienteExternoConSeguro Then
       MsgBox "Debe marcar alguna ATENCION", vbInformation, Me.Caption
       Me.TabFua.Tab = 0
       Exit Function
    End If
    'mgaray20140926
    If Me.cmbConceptoP.Text = "" And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghRegistroCitaCE _
                And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghAdmisionEmergencia _
                And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghPacienteExternoConSeguro Then
       MsgBox "Debe elegir CONCEPTO PRESTACIONAL", vbInformation, Me.Caption
       Me.TabFua.Tab = 0
       If cmbConceptoP.Enabled = True Then cmbConceptoP.SetFocus
       Exit Function
    End If
    If chkIntramural.Value = 0 And chkExtramural.Value = 0 Then
       MsgBox "Debe marcar LUGAR DE ATENCION", vbInformation, Me.Caption
       Me.TabFua.Tab = 0
       Exit Function
    End If
    
    If mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghAdmisionHospitalizacion Then        'debb-02/05/2016
        If chkAtencionReferencia.Value = -1 And mo_lnIdTablaLISTBARITEMS = sghFormatoFUA And txtRO.Text = "" Then
           MsgBox "Debe elegir el Establecimiento de la REFERENCIA ORIGEN", vbInformation, Me.Caption
           Me.TabFua.Tab = 0
           Exit Function
           
        End If
        If txtRO.Text <> "" And txtRONumero.Enabled = True And txtRONumero.Text = "" Then
           MsgBox "Debe ingresar el N° HOJA REFERENCIA - ORIGEN", vbInformation, Me.Caption
           Me.TabFua.Tab = 0
           Exit Function
        End If
    End If                                                                                  'debb-02/05/2016
    
    If chkPAestablecimiento.Value = 0 And chkPAaisped.Value = 0 Then
       MsgBox "Debe marcar PERSONAL QUE ATIENDE", vbInformation, Me.Caption
       Me.TabFua.Tab = 0
       Exit Function
    End If
    'debb-22/09/2015
    lcTexto = "S01/900/901/906//"
    If InStr(lcTexto, Me.ucSISfuaCodPrestacion1.CodigoPrestacion) = 0 Then
        If Val(mo_cmbIdDestinoAtencion.BoundText) = 0 And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghRegistroCitaCE _
                         And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghAdmisionEmergencia And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghPacienteExternoConSeguro Then
           MsgBox "Debe elegir DESTINO DEL ASEGURADO", vbInformation, Me.Caption
           Me.TabFua.Tab = 0
           If cmbIdDestinoAtencion.Enabled = True Then cmbIdDestinoAtencion.SetFocus
           Exit Function
        End If
    End If
    '
    If Val(mo_cmbIdDestinoAtencion.BoundText) = 6 And txtRDnumero.Text = "" Then
       MsgBox "Falta REFERENCIA DESTINO: Establecimiento y Número de Hoja de Referencia", vbInformation, Me.Caption
       Me.TabFua.Tab = 0
       Exit Function
    End If
    '
    If sighentidades.PresionVerificaSiTieneDatosYsiEstaOK(txtSPpa.Text) = False Then
       Me.TabFua.Tab = 1
    End If
    'mgaray20140926
    If mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghRegistroCitaCE And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghAdmisionEmergencia And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghPacienteExternoConSeguro Then
        lbEstaOk = False
        oRsDx.MoveFirst
        Do While Not oRsDx.EOF
           If oRsDx.Fields!DxIngresoDefinitivo = True Or oRsDx.Fields!DxEgresoDefinitivo = True Then
              lcDxPrincipal = IIf(oRsDx.Fields!DxIngresoDefinitivo = True, Trim(oRsDx.Fields("DxIngreso").Value), IIf(IsNull(oRsDx.Fields("DxEgreso").Value), "", oRsDx.Fields("DxEgreso").Value))
              lcDxPrincipalNro = oRsDx.Fields!dxNro
              lbEstaOk = True
              Exit Do
           End If
           oRsDx.MoveNext
        Loop
        If lbEstaOk = False And ml_IdTipoServicio <> sghConsultaExterna Then
           MsgBox "No existe ningún DIAGNOSTICO DEFINITIVO", vbInformation, Me.Caption
           Me.TabFua.Tab = 1
           Exit Function
        End If
    End If
    If Me.txtMedico.Text = "" Then
       MsgBox "Debe registrar el DNI del Médico", vbInformation, Me.Caption
       Me.TabFua.Tab = 1
       On Error Resume Next
       Me.txtMedicoDni.SetFocus
       Exit Function
    End If
    If txtNautorizacion.Locked = False And txtNautorizacion.Text = "" Then
       MsgBox "Debe registrar el N° AUTORIZACION (Concepto Prestacional)", vbInformation, Me.Caption
       Me.TabFua.Tab = 0
       On Error Resume Next
       Me.txtNautorizacion.SetFocus
       Exit Function
    End If
    If txtMonto.Locked = False And txtMonto.Text = "" Then
       MsgBox "Debe registrar el MONTO (Concepto Prestacional)", vbInformation, Me.Caption
       Me.TabFua.Tab = 0
       On Error Resume Next
       Me.txtMonto.SetFocus
       Exit Function
    End If

    ActualizaDxEnServiciosIntermediosParaLosVacios
    ValidarDatosObligatorios = True
End Function

Sub ActualizaDxEnServiciosIntermediosParaLosVacios()
    If oRsFarmacia.RecordCount > 0 Then
       oRsFarmacia.MoveFirst
       Do While Not oRsFarmacia.EOF
          If oRsFarmacia.Fields!dx = "" Then
                oRsFarmacia.Fields!dx = lcDxPrincipal
                oRsFarmacia.Fields!dxNro = lcDxPrincipalNro
                oRsFarmacia.Update
          End If
          oRsFarmacia.MoveNext
       Loop
    End If
    If oRsPatologia.RecordCount > 0 Then
       oRsPatologia.MoveFirst
       Do While Not oRsPatologia.EOF
          If oRsPatologia.Fields!dx = "" Then
                oRsPatologia.Fields!dx = lcDxPrincipal
                oRsPatologia.Fields!dxNro = lcDxPrincipalNro
                oRsPatologia.Update
          End If
          oRsPatologia.MoveNext
       Loop
    End If
End Sub

Function ValidarReglas() As Boolean
    ValidarReglas = False
    Dim lcMensaje As String, lbProseguir As Boolean
    lcMensaje = ""
    If mi_opcion = sghAgregar Then
    
        If lcCodigoEstablecimientoAdscripcionSIS <> "" And _
           mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghAdmisionHospitalizacion Then    'debb-02/05/2016
           lbProseguir = True
           If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia Then
              If Not (ml_IdOrigenAtencion = 21 Or ml_IdOrigenAtencion = 22) Then
                 lbProseguir = False
              End If
              If lbProseguir = True Then
                 lcMensaje = mo_ReglasSISgalenhos.ChequeaCodigoEstablecimientoAdscripcion(lcCodigoEstablecimientoAdscripcionSIS, _
                                                ml_IdTipoServicio, _
                                                IIf(txtRONumero.Text <> "", 4, 0), _
                                                ucSISfuaCodPrestacion1.CodigoPrestacion)
              End If
           End If
        End If
        
        
        If lcMensaje <> "" Then
              MsgBox lcMensaje, vbInformation, Me.Caption
              Exit Function
        End If
    End If
    ValidarReglas = True

End Function

Private Sub btnAceptar_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAddFarmacia_Click()
    If Me.btnAddFarmacia.Enabled = False Then
       Exit Sub
    End If
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    Dim lcCabecera As String
    Dim oRsTmp4 As New Recordset
    oPaquetesBuscar.idPuntoCarga = 2501   '= sghPtoCargaFarmacia
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
       If oRsFarmacia.RecordCount > 0 Then
          oRsFarmacia.MoveFirst
          oRsFarmacia.Find "id=" & oPaquetesBuscar.IdProducto
          If Not oRsFarmacia.EOF Then
             MsgBox "Ese Medicamento/Insumo ya se registró", vbInformation, "Mensaje"
             Set oPaquetesBuscar = Nothing
             Exit Sub
          End If
       End If
       Set oRsTmp4 = mo_AdminServiciosComunes.MedicamentosInsumosSeleccionarPorCodigo(oPaquetesBuscar.codigo)
       oRsFarmacia.AddNew
       oRsFarmacia.Fields!id = oPaquetesBuscar.IdProducto
       oRsFarmacia.Fields!tipo = IIf(oRsTmp4!TipoProducto = 1, lcInsumo, lcMedicamento)
       oRsFarmacia.Fields!MedicInsumo = oPaquetesBuscar.Descripcion
       oRsFarmacia.Fields!recetado = 1
       oRsFarmacia.Fields!cantidad = 1
       oRsFarmacia.Fields!dx = lcDxPrincipal
       oRsFarmacia.Fields!Precio = oPaquetesBuscar.Precio  'DevuelvePrecioItem(oPaquetesBuscar.idProducto, sghPtoCargaFarmacia)
       oRsFarmacia.Fields!codigo = oPaquetesBuscar.codigo
       oRsFarmacia.Fields!dxNro = lcDxPrincipalNro
       oRsFarmacia.Fields!formaF = IIf(IsNull(oRsTmp4.Fields!FormaFarmaceutica), "", oRsTmp4.Fields!FormaFarmaceutica)
       oRsFarmacia.Update
       oRsFarmacia.Sort = "tipo,MedicInsumo"
       btnAddFarmacia.SetFocus
    End If
    Set oPaquetesBuscar = Nothing
    Set oRsTmp4 = Nothing
End Sub

Private Sub btnAddFarmacia_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lnKeyCode As Integer
    lnKeyCode = KeyCode
    AdministrarKeyPreview lnKeyCode

End Sub

Private Sub btnAddPatologia_Click()
    If Me.btnAddPatologia.Enabled = False Then
       Exit Sub
    End If
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    Dim lcCabecera As String, lnPrecio As Double
    Dim oRsTmp1 As New Recordset, lnIdPuntoCarga As Long, lcPuntoCarga As String
    oPaquetesBuscar.idPuntoCarga = 2500
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
       If oRsPatologia.RecordCount > 0 Then
          oRsPatologia.MoveFirst
          oRsPatologia.Find "id=" & oPaquetesBuscar.IdProducto
          If Not oRsPatologia.EOF Then
             MsgBox "Ese Procedimiento ya se registró", vbInformation, "Mensaje"
             Set oPaquetesBuscar = Nothing
             Exit Sub
          End If
       End If
       'lnPrecio = DevuelvePrecioItem(oPaquetesBuscar.idProducto, sghPtoCargaPatologiaClinica)
       '
       lnIdPuntoCarga = 0
       lcPuntoCarga = lcOtros
       Set oRsTmp1 = mo_AdminServiciosComunes.FactCatalogoServiciosPtosSeleccionar(" Where IdProducto=" & oPaquetesBuscar.IdProducto)
       If oRsTmp1.RecordCount > 0 Then
          lnIdPuntoCarga = oRsTmp1.Fields!idPuntoCarga
          Select Case lnIdPuntoCarga
          Case sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, sghPuntosCargaBasicos.sghPtoCargaBancoSangre1, sghPuntosCargaBasicos.sghPtoCargaBancoSangre2, sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica   'Laboratorio
               lcPuntoCarga = lcLaboratorio
          Case sghPuntosCargaBasicos.sghPtoCargaEcogGeneral, sghPuntosCargaBasicos.sghPtoCargaEcogObstetrica, sghPuntosCargaBasicos.sghPtoCargaRayosX, sghPuntosCargaBasicos.sghPtoCargaTomografia  'Imágenes
               lcPuntoCarga = lcImagenes
          End Select
       End If
       oRsTmp1.Close
       '
       oRsPatologia.AddNew
       oRsPatologia.Fields!id = oPaquetesBuscar.IdProducto
       oRsPatologia.Fields!tipo = lcPuntoCarga
       oRsPatologia.Fields!procedimiento = oPaquetesBuscar.Descripcion
       oRsPatologia.Fields!indicado = 1
       oRsPatologia.Fields!ejecutado = 1
       oRsPatologia.Fields!dx = lcDxPrincipal
       oRsPatologia.Fields!Precio = oPaquetesBuscar.Precio
       oRsPatologia.Fields!idPuntoCarga = lnIdPuntoCarga
       oRsPatologia.Fields!codigo = oPaquetesBuscar.codigo
       oRsPatologia.Fields!dxNro = lcDxPrincipalNro
       oRsPatologia.Update
       oRsPatologia.Sort = "tipo,procedimiento"
       btnAddPatologia.SetFocus
    End If
    Set oPaquetesBuscar = Nothing
    Set oRsTmp1 = Nothing
End Sub


Private Sub btnAddPatologia_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lnKeyCode As Integer
    lnKeyCode = KeyCode
    AdministrarKeyPreview lnKeyCode
End Sub

Private Sub btnBuscaColegio_Click()
    Dim mo_DetalleColegios As New FuaDetalleColegios
    mo_DetalleColegios.CodigoColegio = Trim(Me.txtColegioCodigo.Text)
    mo_DetalleColegios.DescColegio = Trim(Me.txtColegio.Text)
    mo_DetalleColegios.MostrarFormulario
    If mo_DetalleColegios.BotonPresionado = sghAceptar Then
        txtColegioCodigo.Text = mo_DetalleColegios.CodigoColegio
        txtColegio.Text = mo_DetalleColegios.DescColegio
        Set mo_DetalleColegios = Nothing
    Else
        Set mo_DetalleColegios = Nothing
    End If
End Sub

Private Sub btnBuscarEstablecimientoD_Click()
        Dim oBusqueda As New SIGHNegocios.BuscaEstablecimientos
        Dim oDOEstablecimiento As New DOEstablecimiento
        oBusqueda.MostrarFormulario
        If oBusqueda.BotonPresionado = sghAceptar Then
            Set oDOEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
            If Not oDOEstablecimiento Is Nothing Then
                txtRDcodigo.Text = Right("0000000000" & oDOEstablecimiento.codigo, 10)
                txtRD.Text = oDOEstablecimiento.nombre
                txtRDnumero.SetFocus
            End If
        End If
        Set oBusqueda = Nothing
        Set oDOEstablecimiento = Nothing
End Sub

Private Sub btnBuscarEstablecimientoO_Click()
        Dim oBusqueda As New SIGHNegocios.BuscaEstablecimientos
        Dim oDOEstablecimiento As New DOEstablecimiento
        oBusqueda.MostrarFormulario
        If oBusqueda.BotonPresionado = sghAceptar Then
            Set oDOEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
            If Not oDOEstablecimiento Is Nothing Then
                txtROcodigo.Text = Right("0000000000" & oDOEstablecimiento.codigo, 10)
                txtRO.Text = oDOEstablecimiento.nombre
                txtRONumero.SetFocus
            End If
        End If
        Set oBusqueda = Nothing
        Set oDOEstablecimiento = Nothing
End Sub

Private Sub btnBuscarPaciente_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New doPaciente
    Dim oConexion As New Connection
    Dim oRsTmp1 As New Recordset
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
           txtNhistoriaClinica.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oDOPaciente.NroHistoriaClinica)), False)
           ml_NroHistoriaClinica = txtNhistoriaClinica.Text
           lcSql = "Where Paterno='" & oDOPaciente.ApellidoPaterno & "' and Materno='" & oDOPaciente.ApellidoMaterno & _
                         "' and pNombre='" & oDOPaciente.PrimerNombre & _
                         IIf(oDOPaciente.SegundoNombre = "", "", "' and oNombres='" & oDOPaciente.SegundoNombre) & _
                         "' and fNacimiento=CONVERT(DATETIME,'" & oDOPaciente.FechaNacimiento & "',103)" & _
                         " and Genero='" & IIf(oDOPaciente.idTipoSexo = 1, "1", "0") & "'"
           Set oRsTmp1 = mo_ReglasSISgalenhos.SisFiltraPacientesAfiliados(lcSql, wxParametroJAMO)
           If oRsTmp1.RecordCount > 0 Then
               LlenaDatosPersonalesDesdeFiliacionesSIS oRsTmp1, True
               If mo_lnIdTablaLISTBARITEMS = sghFormatoFUA Then
                  ReglasDeConsistenciasAntesDeCargarFormulario
               End If
           End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oDOPaciente = Nothing
    Set oBusqueda = Nothing
    Set oRsTmp1 = Nothing
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Public Sub btnguardafua_Click()
Select Case mi_opcion
   Case sghAgregar
       ArsSisCargaDatosAntesDeGrabar
       If ValidacionesOK = True Then
            If AgregarDatos() Then
                 'If MsgBox("Los datos se agregaron correctamente" & Chr(13) & "Desea Imprimir el Formato", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                   ' ImpresionFua
                 'End If
                 'Me.Visible = False
                 Unload Me
            Else
               ' MsgBox "No se pudo agregar los datos" + Chr(13) + mo_ReglasSISgalenhos.MensajeError, vbExclamation, Me.Caption
                If lnNroFuaRepetido = True Then
'                    If Val(wxParametro320) = sghFuaTipo.sghFuaTipoAutomatico Then
'                       txtFua3.Text = Right("00000000" & Trim(Str(Val(mo_ReglasSISgalenhos.sisFuaAtencionUltimoCorrelativo()) + 1)), 8)
'                    Else
                    Dim oRsTmp1 As New Recordset
                    Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaSeleccionarTodos
                    If oRsTmp1.RecordCount > 0 Then
                       txtFua3.Text = Right("00000000" & Trim(Str(Val(oRsTmp1.Fields!FuaUltimoGenerado) + 1)), 8)
                    End If
                    oRsTmp1.Close
                    Set oRsTmp1 = Nothing
'                    End If
                End If
            End If
       End If

   End Select
End Sub

Private Sub btnImprimir_Click()
    If ValidacionesOK = True Then
        Dim lbOk As Boolean
        If mi_opcion = sghAgregar Then
            MsgBox "No se puede Imprimir el Formato FUA, tiene que guardarlos previamente", vbInformation, Me.Caption
            Exit Sub
        End If
        ImpresionFua
        Unload Me
    End If
End Sub

Function ValidacionesOK() As Boolean
    'debb-16/11/2015 (inicio)
    If mi_opcion = sghConsultar And mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghFormatoFUA Then
       ValidacionesOK = True
       Exit Function
    End If
    'debb-16/11/2015 (fin)
    ValidacionesOK = False
    If (mi_opcion = sghModificar And EsUnFuaEmitidoEnVentanillaCitas = True And cmbConceptoP.Text = "" And _
                    ucSISfuaCodPrestacion1.CodigoPrestacion <> "") Or (cmbConceptoP.Text = "" And _
                    mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghPacienteExternoConSeguro) Then
        CargaDatosAlObjetosDeDatos
        CargaValoresVacunasSp
        ValidacionesOK = True
    Else
        If ValidarDatosObligatorios() Then
            If ValidarReglas() Then
               CargaDatosAlObjetosDeDatos
               CargaValoresVacunasSp
               If ReglasDeConsistenciasAntesDeGrabarFUA = True Then
                  ValidacionesOK = True
               End If
            End If
        End If
    End If
End Function

Private Sub chkAtencionAmbulatoria_Click(Value As Integer)
    If chkAtencionAmbulatoria.Value = -1 Then
       chkAtencionReferencia.Value = 0
       chkAtencionEmergencia.Value = 0
       If ml_IdTipoServicio = 0 Then
          mo_Formulario.HabilitarDeshabilitar txtRONumero, False
          btnBuscarEstablecimientoO.Enabled = False
          txtROcodigo.Text = "": txtRO.Text = "": txtRONumero.Text = ""
       End If
       On Error Resume Next
       chkGestante.SetFocus
    End If

End Sub

Private Sub chkAtencionAmbulatoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkAtencionAmbulatoria
    AdministrarKeyPreview KeyCode
   
End Sub

Private Sub chkAtencionAmbulatoria_LostFocus()
    chkAtencionAmbulatoria_Click 1
End Sub

Private Sub chkAtencionEmergencia_Click(Value As Integer)
    If chkAtencionEmergencia.Value = -1 Then
       chkAtencionAmbulatoria.Value = 0
       chkAtencionReferencia.Value = 0
       If ml_IdTipoServicio = 0 Then
          mo_Formulario.HabilitarDeshabilitar txtRONumero, False
          btnBuscarEstablecimientoO.Enabled = False
          txtROcodigo.Text = "": txtRO.Text = "": txtRONumero.Text = ""
       End If
       On Error Resume Next
       chkGestante.SetFocus
    End If

End Sub

Private Sub chkAtencionEmergencia_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkAtencionEmergencia
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkAtencionEmergencia_LostFocus()
   chkAtencionEmergencia_Click 1
End Sub


Private Sub chkAtencionReferencia_Click(Value As Integer)
    If chkAtencionReferencia.Value = -1 Then
       chkAtencionAmbulatoria.Value = 0
       chkAtencionEmergencia.Value = 0
       If ml_IdTipoServicio = 0 Then
          mo_Formulario.HabilitarDeshabilitar txtRONumero, True
          btnBuscarEstablecimientoO.Enabled = True
       End If
       On Error Resume Next
       chkGestante.SetFocus
    End If
    

End Sub

Private Sub chkAtencionReferencia_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkAtencionReferencia
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkAtencionReferencia_LostFocus()
    chkAtencionReferencia_Click 1
End Sub


Private Sub chkExtramural_Click(Value As Integer)
    If chkExtramural.Value = -1 Then
       chkIntramural.Value = 0
       On Error Resume Next
       cmbConceptoP.SetFocus
    End If

End Sub

Private Sub chkExtramural_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkExtramural
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkExtramural_LostFocus()
    chkExtramural_Click 1
End Sub


Private Sub chkGestante_Click(Value As Integer)
    If chkGestante.Value = 1 Then
       On Error Resume Next
       cmbConceptoP.SetFocus
    End If
End Sub

Private Sub chkGestante_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkGestante
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkGestante_LostFocus()
    chkGestante_Click 1
End Sub


Private Sub chkIntramural_Click(Value As Integer)
    If chkIntramural.Value = -1 Then
       chkExtramural.Value = 0
       On Error Resume Next
       cmbConceptoP.SetFocus
    End If

End Sub

Private Sub chkIntramural_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkIntramural
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkIntramural_LostFocus()
    chkIntramural_Click 1
End Sub


Private Sub chkPAaisped_Click(Value As Integer)
    If chkPAaisped.Value = -1 Then
       chkPAestablecimiento.Value = 0
       chkPAOfeFlexible.Value = 0
       txtPACodOfFlexible.Text = ""
       mo_Formulario.HabilitarDeshabilitar txtPACodOfFlexible, False
       On Error Resume Next
       ucSISfuaCodPrestacion1.SetFocus
    End If
End Sub

Private Sub chkPAaisped_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkPAaisped
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkPAaisped_LostFocus()
    chkPAaisped_Click 1
End Sub

  
Private Sub chkPAestablecimiento_Click(Value As Integer)
    If chkPAestablecimiento.Value = -1 Then
       chkPAaisped.Value = 0
       chkPAOfeFlexible.Value = 0
       txtPACodOfFlexible.Text = ""
       mo_Formulario.HabilitarDeshabilitar txtPACodOfFlexible, False
       On Error Resume Next
       ucSISfuaCodPrestacion1.SetFocus
    End If
End Sub

Private Sub chkPAestablecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkPAestablecimiento
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkPAestablecimiento_LostFocus()
    chkPAestablecimiento_Click 1
End Sub


Private Sub chkPAOfeFlexible_Click(Value As Integer)
    If chkPAOfeFlexible.Value = -1 Then
       chkPAestablecimiento.Value = 0
       chkPAaisped.Value = 0
       mo_Formulario.HabilitarDeshabilitar txtPACodOfFlexible, True
       On Error Resume Next
       ucSISfuaCodPrestacion1.SetFocus
    End If
End Sub

Private Sub chkPAOfeFlexible_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkPAOfeFlexible
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkPAOfeFlexible_LostFocus()
    chkPAOfeFlexible_Click 1
End Sub

Private Sub chkPuerpera_Click(Value As Integer)
    If chkPuerpera.Value = 1 Then
       On Error Resume Next
       cmbConceptoP.SetFocus
    End If
End Sub

Private Sub chkPuerpera_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkPuerpera
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkPuerpera_LostFocus()
    chkPuerpera_Click 1
End Sub

Private Sub chkSPCorTarCordonSI_Click(Value As Integer)
    If Me.chkSPCorTarCordonSI.Value = -1 Then
       Me.chkSPCorteTarCordonNO.Value = 0
    End If
End Sub

Private Sub chkSPCorTarCordonSI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPCorTarCordonSI
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkSPCorTarCordonSI_LostFocus()
    chkSPCorTarCordonSI_Click 1
End Sub

Private Sub chkSPCorteTarCordonNO_Click(Value As Integer)
    If Me.chkSPCorteTarCordonNO.Value = -1 Then
       Me.chkSPCorTarCordonSI.Value = 0
    End If
End Sub

Private Sub chkSPCorteTarCordonNO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPCorteTarCordonNO
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkSPCorteTarCordonNO_LostFocus()
    chkSPCorteTarCordonNO_Click 1
End Sub

Private Sub chkSPPartoVertNO_Click(Value As Integer)
    If Me.chkSPPartoVertNO.Value = -1 Then
       Me.chkSPPartoVertSI.Value = 0
    End If
End Sub

Private Sub chkSPPartoVertNO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPPartoVertNO
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkSPPartoVertNO_LostFocus()
    chkSPPartoVertNO_Click 1
End Sub

Private Sub chkSPPartoVertSI_Click(Value As Integer)
    If Me.chkSPPartoVertSI.Value = -1 Then
       Me.chkSPPartoVertNO.Value = 0
    End If
End Sub

Private Sub chkSPPartoVertSI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPPartoVertSI
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkSPPartoVertSI_LostFocus()
    chkSPPartoVertSI_Click 1
End Sub

Private Sub chkSPRNPrematuroNO_Click(Value As Integer)
    If Me.chkSPRNPrematuroNO.Value = -1 Then
       Me.chkSPRNPrematuroSI.Value = 0
    End If
End Sub

Private Sub chkSPRNPrematuroNO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPRNPrematuroNO
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkSPRNPrematuroNO_LostFocus()
    chkSPRNPrematuroNO_Click 1
End Sub

Private Sub chkSPRNPrematuroSI_Click(Value As Integer)
    If Me.chkSPRNPrematuroSI.Value = -1 Then
       Me.chkSPRNPrematuroNO.Value = 0
    End If
End Sub

Private Sub chkSPRNPrematuroSI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPRNPrematuroSI
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkSPRNPrematuroSI_LostFocus()
    chkSPRNPrematuroSI_Click 1
End Sub

Private Sub chkSPconsejeriaNno_Click(Value As Integer)
    If Me.chkSPconsejeriaNno.Value = -1 Then
       Me.chkSPconsejeriaNsi.Value = 0
    End If
End Sub

Private Sub chkSPconsejeriaNno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPconsejeriaNno
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkSPconsejeriaNno_LostFocus()
    chkSPconsejeriaNno_Click 1
End Sub

Private Sub chkSPconsejeriaNsi_Click(Value As Integer)
    If Me.chkSPconsejeriaNsi.Value = -1 Then
       Me.chkSPconsejeriaNno.Value = 0
    End If
End Sub

Private Sub chkSPconsejeriaNsi_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPconsejeriaNsi
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPconsejeriaNsi_LostFocus()
    chkSPconsejeriaNsi_Click 1
End Sub


Private Sub chkSPSecuelaNaceNO_Click(Value As Integer)
    If Me.chkSPSecuelaNaceNO.Value = -1 Then
       Me.chkSPSecuelaNaceSI.Value = 0
    End If
End Sub

Private Sub chkSPSecuelaNaceNO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPSecuelaNaceNO
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkSPSecuelaNaceNO_LostFocus()
    chkSPSecuelaNaceNO_Click 1
End Sub

Private Sub chkSPSecuelaNaceSI_Click(Value As Integer)
    If Me.chkSPSecuelaNaceSI.Value = -1 Then
       Me.chkSPSecuelaNaceNO.Value = 0
    End If
End Sub

Private Sub chkSPSecuelaNaceSI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPSecuelaNaceSI
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPSecuelaNaceSI_LostFocus()
    chkSPSecuelaNaceSI_Click 1
End Sub

Private Sub chkSPeedpNO_Click(Value As Integer)
    If Me.chkSPeedpNO.Value = -1 Then
       Me.chkSPeedpSI.Value = 0
    End If

End Sub

Private Sub chkSPeedpNO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPeedpNO
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPeedpNO_LostFocus()
    chkSPeedpNO_Click 1
End Sub

Private Sub chkSPeedpSI_Click(Value As Integer)
    If Me.chkSPeedpSI.Value = -1 Then
       Me.chkSPeedpNO.Value = 0
    End If

End Sub

Private Sub chkSPeedpSI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPeedpSI
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPeedpSI_LostFocus()
    chkSPeedpSI_Click 1
End Sub

Private Sub chkSPEvalIntegralNO_Click(Value As Integer)
    If Me.chkSPEvalIntegralNO.Value = -1 Then
       Me.chkSPEvalIntegralSI.Value = 0
    End If

End Sub

Private Sub chkSPEvalIntegralNO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPEvalIntegralNO
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPEvalIntegralNO_LostFocus()
    chkSPEvalIntegralNO_Click 1
End Sub

Private Sub chkSPEvalIntegralSI_Click(Value As Integer)
    If Me.chkSPEvalIntegralSI.Value = -1 Then
       Me.chkSPEvalIntegralNO.Value = 0
    End If

End Sub

Private Sub chkSPEvalIntegralSI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPEvalIntegralSI
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPEvalIntegralSI_LostFocus()
    chkSPEvalIntegralSI_Click 1
End Sub

Private Sub chkSPTamizajeSalMNOR_Click(Value As Integer)
    If Me.chkSPTamizajeSalMNOR.Value = -1 Then
       Me.chkSPTamizajeSalMPAT.Value = 0
    End If

End Sub

Private Sub chkSPTamizajeSalMNOR_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPTamizajeSalMNOR
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPTamizajeSalMNOR_LostFocus()
    chkSPTamizajeSalMNOR_Click 1
End Sub

Private Sub chkSPTamizajeSalMPAT_Click(Value As Integer)
    If Me.chkSPTamizajeSalMPAT.Value = -1 Then
       Me.chkSPTamizajeSalMNOR.Value = 0
    End If
End Sub

Private Sub chkSPTamizajeSalMPAT_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPTamizajeSalMPAT
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkSPTamizajeSalMPAT_LostFocus()
     chkSPTamizajeSalMPAT_Click 1
End Sub

Private Sub chkSBajoPesoNO_Click(Value As Integer)
    If Me.chkSBajoPesoNO.Value = -1 Then
       Me.chkSBajoPesoSI.Value = 0
    End If
End Sub

Private Sub chkSBajoPesoNO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSBajoPesoNO
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSBajoPesoNO_LostFocus()
    chkSBajoPesoNO_Click 1
End Sub

Private Sub chkSBajoPesoSI_Click(Value As Integer)
    If Me.chkSBajoPesoSI.Value = -1 Then
       Me.chkSBajoPesoNO.Value = 0
    End If
End Sub

Private Sub chkSBajoPesoSI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSBajoPesoSI
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkSBajoPesoSI_LostFocus()
    chkSBajoPesoSI_Click 1
End Sub



Private Sub chkSPConIntegralNO_Click(Value As Integer)
    If Me.chkSPConIntegralNO.Value = -1 Then
       Me.chkSPConIntegralSI.Value = 0
    End If

End Sub

Private Sub chkSPConIntegralNO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPConIntegralNO
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPConIntegralNO_LostFocus()
    chkSPConIntegralNO_Click 1
End Sub

Private Sub chkSPConIntegralSI_Click(Value As Integer)
    If Me.chkSPConIntegralSI.Value = -1 Then
       Me.chkSPConIntegralNO.Value = 0
    End If

End Sub

Private Sub chkSPConIntegralSI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPConIntegralSI
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPConIntegralSI_LostFocus()
   chkSPConIntegralSI_Click 1
End Sub

Private Sub cmbColegioNivel_Click()
  If cmbColegioNivel.ListIndex = -1 Then Exit Sub
       
  mo_cmbColegioGrado.BoundColumn = "IdGrado"
  mo_cmbColegioGrado.ListField = "Grado"
  On Error Resume Next
  Set mo_cmbColegioGrado.RowSource = mo_ReglasSISgalenhos.SisFuaColegioGradoSeleccionarPorNivel(Val(cmbColegioNivel.ItemData(cmbColegioNivel.ListIndex)))
  mo_cmbColegioGrado.BoundText = ""
'  mo_cmbColegioGrado.Enabled = True
End Sub

Private Sub cmbColegioNivel_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbColegioNivel
End Sub


Private Sub cmbColegioNivel_LostFocus()
   'If cmbIdDepartamentoDomicilio.Text <> "" Then
   '    mo_cmbIdDepartamentoDomicilio.BoundText = Val(Split(cmbIdDepartamentoDomicilio.Text, " = ")(0))
   'End If
   mo_Formulario.MarcarComoVacio cmbColegioNivel
End Sub

Private Sub cmbColegioNivel_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbConceptoP_Click()
    If Val(mo_cmbConceptoP.BoundText) = 2 Or Val(mo_cmbConceptoP.BoundText) = 3 Or Val(mo_cmbConceptoP.BoundText) = 6 Then
        mo_Formulario.HabilitarDeshabilitar txtNautorizacion, True
        mo_Formulario.HabilitarDeshabilitar txtMonto, True
    Else
        txtNautorizacion.Text = ""
        txtMonto.Text = ""
        mo_Formulario.HabilitarDeshabilitar txtNautorizacion, False
        mo_Formulario.HabilitarDeshabilitar txtMonto, False
    End If
End Sub

Private Sub cmbConceptoP_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbConceptoP
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtRONumero
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdDestinoAtencion_Click()
    If Val(mo_cmbIdDestinoAtencion.BoundText) = 6 Then
        mo_Formulario.HabilitarDeshabilitar Me.txtRDnumero, True
    Else
        mo_Formulario.HabilitarDeshabilitar Me.txtRDnumero, False
        Me.txtRD.Text = ""
        Me.txtRDcodigo.Text = ""
        Me.txtRDnumero.Text = ""
    End If
End Sub

Private Sub cmbIdDestinoAtencion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDestinoAtencion
    AdministrarKeyPreview KeyCode
End Sub

'Private Sub Command1_Click()
'Text1.Text = ucSISfuaCodPrestacion1.CodigoPrestacion
'End Sub

Private Sub Form_Activate()
  If lbEsIgualQueArSIS = False Then
        If mo_lbCargaTablasUnaVez = True Then
            mo_lbCargaTablasUnaVez = False
            If mi_opcion = sghAgregar And ml_CodigoPrestacion <> "" Then
               ucSISfuaCodPrestacion1.CodigoPrestacion = ml_CodigoPrestacion
            End If
            ReglasDeConsistenciasAntesDeCargarFormulario
            PermitirManipularDatosJaladosDesdeGalenHos
        End If
        On Error Resume Next
        mo_Formulario.HabilitarDeshabilitar txtFua3, False
        If Val(wxParametro320) = sghFuaTipo.sghFuaTipoAutomatico And mi_opcion = sghAgregar Then
           ucSISfuaCodPrestacion1.SetFocus
        ElseIf Val(wxParametro320) = sghFuaTipo.sghFuaTipoManual And mi_opcion = sghAgregar Then
           mo_Formulario.HabilitarDeshabilitar txtFua3, True
           txtFua3.SetFocus
        Else
           Me.btnAceptar.SetFocus
        End If
  Else
        If mo_lbCargaTablasUnaVez = True Then
           mo_lbCargaTablasUnaVez = False
           ReglasDeConsistenciasAntesDeCargarFormulario
           If mi_opcion = sghModificar Then
              ArsSisHabilitaAgregarYmodificar
              PermitirManipularDatosSegunSexo
              mo_Formulario.HabilitarDeshabilitar txtFua3, False
                          
              
           End If
        End If
        If Val(wxParametro320) = sghFuaTipo.sghFuaTipoAutomatico And mi_opcion = sghAgregar Then
           mo_Formulario.HabilitarDeshabilitar txtFua3, False
        ElseIf Val(wxParametro320) = sghFuaTipo.sghFuaTipoManual And mi_opcion = sghAgregar Then
           mo_Formulario.HabilitarDeshabilitar txtFua3, True
           txtFua3.SetFocus
        Else
           mo_Formulario.HabilitarDeshabilitar txtFua3, True
           txtFua3.SetFocus
        End If
  End If
  If mo_SoloImprimeFUAyaGrabado = True Then
     btnImprimir_Click
     Unload Me
  End If
  
  'HRA 10/12/2020 Cambio 47 Inicio
    If mc_GuardarFua = "S" Then
        Me.btnguardafua_Click
    End If
  'HRA 10/12/2020  Cambio 47 Fin
  
End Sub


Sub ReglasDeConsistenciasAntesDeCargarFormulario()
     Me.ucSISfuaCodPrestacion1.ReglasDeConsistenciasAntesDeCargarFormulario ml_IdTipoServicio, Left(txtSexo.Text, 1), ml_edad_En_YYYYMMDD
        
End Sub

Private Sub Form_Load()    '
    '
    ldHoy = lcBuscaParametro.RetornaFechaServidorSQL
    '
    lblDiag1.Width = 4800
    lblDiag2.Width = 3995
    lblDiag3.Width = 3025
    lblDiag1.Top = 250
    lblDiag2.Top = 250
    lblDiag3.Top = 250
    Me.grdDx.Top = 550
    '
    mo_lbCargaTablasUnaVez = True
    lcEquix = "X"
    InicilizarParametros
    lcNivelEstablecimiento = mo_ReglasSISgalenhos.Sis_a_categoriaeessDevuelveNivel(wxParametro303)
    '
    mo_Formulario.HabilitarDeshabilitar lblDiag1, False
    mo_Formulario.HabilitarDeshabilitar lblDiag2, False
    mo_Formulario.HabilitarDeshabilitar lblDiag3, False
    mo_Formulario.HabilitarDeshabilitar txtFua1, False
    mo_Formulario.HabilitarDeshabilitar txtFua2, False
    If Val(wxParametro320) = sghFuaTipo.sghFuaTipoAutomatico Then
       mo_Formulario.HabilitarDeshabilitar txtFua3, False
    Else
       mo_Formulario.HabilitarDeshabilitar txtFua3, True
    End If
    mo_Formulario.HabilitarDeshabilitar txtCScodigo, False
    mo_Formulario.HabilitarDeshabilitar txtCS, False
    mo_Formulario.HabilitarDeshabilitar txtNroAfiliacion1, False
    mo_Formulario.HabilitarDeshabilitar txtNroAfiliacion2, False
    mo_Formulario.HabilitarDeshabilitar txtNroAfiliacion3, False
    mo_Formulario.HabilitarDeshabilitar cmbTipoDocumento, False
    mo_Formulario.HabilitarDeshabilitar txtNdocumento, False
    mo_Formulario.HabilitarDeshabilitar txtFnacimiento, False
    mo_Formulario.HabilitarDeshabilitar txtFFallecimiento, False
    mo_Formulario.HabilitarDeshabilitar txtSexo, False
    mo_Formulario.HabilitarDeshabilitar txtNhistoriaClinica, False
    mo_Formulario.HabilitarDeshabilitar txtFantencion, False
    mo_Formulario.HabilitarDeshabilitar txtHatencion, False
    mo_Formulario.HabilitarDeshabilitar txtMedicoDni, False
    mo_Formulario.HabilitarDeshabilitar txtMedico, False
    mo_Formulario.HabilitarDeshabilitar txtMedicoColegiatura, False
    mo_Formulario.HabilitarDeshabilitar txtMedicoEspecialidad, False
    mo_Formulario.HabilitarDeshabilitar txtMedicoRNE, False
    chkMedicoEgresado.Enabled = False
    mo_Formulario.HabilitarDeshabilitar txtPaciente, False
    mo_Formulario.HabilitarDeshabilitar txtRO, False
    mo_Formulario.HabilitarDeshabilitar txtRD, False
    mo_Formulario.HabilitarDeshabilitar txtROcodigo, False
    mo_Formulario.HabilitarDeshabilitar txtRDcodigo, False
    mo_Formulario.HabilitarDeshabilitar Me.txtRONumero, False
    mo_Formulario.HabilitarDeshabilitar Me.txtRDnumero, False
    mo_Formulario.HabilitarDeshabilitar txtNautorizacion, False
    mo_Formulario.HabilitarDeshabilitar txtMonto, False
    mo_Formulario.HabilitarDeshabilitar txtInstitucion, False
    mo_Formulario.HabilitarDeshabilitar txtCodSeguro, False
    mo_Formulario.HabilitarDeshabilitar fraCodAfiliacionSeguro, False
    mo_Formulario.HabilitarDeshabilitar txtPACodOfFlexible, False
    cmbEtnia.Locked = True: cmbEtnia.BackColor = &HF9EADF: cmbEtnia.ForeColor = &H808080
    mo_Formulario.HabilitarDeshabilitar txtColegioCodigo, False
    mo_Formulario.HabilitarDeshabilitar txtColegio, False
    cmbUPSfua.Locked = True: cmbUPSfua.BackColor = &HF9EADF: cmbUPSfua.ForeColor = &H808080
    mo_Formulario.HabilitarDeshabilitar txtCodPrestAdicional, True

    
    mo_Formulario.HabilitarDeshabilitar txtFuaVincular, False 'FRANK por mientras

    
    
    btnBuscarPaciente.Enabled = False
    CreaTemporales
    CargaComboBoxes
    CargarDatosAlFormulario
    
    CargaGrdDiagAyuda
    
  
    
 
End Sub

Sub CargaGrdDiagAyuda()
    Set grdDiag.DataSource = oRsDx
    If oRsDx.RecordCount > 0 Then
       oRsDx.MoveFirst
    End If
End Sub

Sub CargaFormatoFUA()
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaSeleccionarTodos
    If oRsTmp1.RecordCount = 0 Then
       MsgBox "No se ha configurado el FORMATO FUA" & Chr(13) & "use la opción 'Herramientas --> Exporta/importa SIS'", vbInformation, Me.Caption
       Me.Visible = False
    Else
        If Val(wxParametro320) = sghFuaTipo.sghFuaTipoAutomatico Then 'Frank 2508
'            MsgBox "El FORMATO FUA tiene una configuración Automatica" & Chr(13) & "[Tabla Parametro:320]", vbExclamation, Me.Caption
            If Val(oRsTmp1.Fields!FuaDisa) = 0 Or Val(oRsTmp1.Fields!FuaLote) = 0 Then
                 MsgBox "No se ha configurado el FORMATO FUA, alguno de los datos tiene valor CERO" & Chr(13) & "use la opción 'Herramientas --> Exporta/importa SIS'", vbInformation, Me.Caption
                 Me.Visible = False
            ElseIf oRsTmp1.Fields!FuaUltimoGenerado < 0 And Val(wxParametro320) = sghFuaTipo.sghFuaTipoAutomatico Then
                 MsgBox "No se ha configurado el FORMATO FUA, el ULTIMO NUMERO FUA GENERADO tiene que ser un número mayor o igual a CERO" & Chr(13) & "use la opción 'Herramientas --> Exporta/importa SIS'", vbInformation, Me.Caption
                 Me.Visible = False
            Else
                 txtFua1.Text = oRsTmp1.Fields!FuaDisa
                 txtFua2.Text = oRsTmp1.Fields!FuaLote
                 txtFua3.Text = ""
            End If
        Else
            If IsNull(oRsTmp1.Fields!FuaNumeroInicial) Or IsNull(oRsTmp1.Fields!FuaNumeroFinal) Then 'Frank 2508
                MsgBox "No se ha configurado el FORMATO FUA, alguno de los datos tiene valor CERO" & Chr(13) & "use la opción 'Herramientas --> Exporta/importa SIS'", vbInformation, Me.Caption
                Me.Visible = False
            Else
                If Val(oRsTmp1.Fields!FuaDisa) = 0 Or Val(oRsTmp1.Fields!FuaLote) = 0 Or Val(oRsTmp1.Fields!FuaNumeroInicial) = 0 Or Val(oRsTmp1.Fields!FuaNumeroFinal) = 0 Then
                     MsgBox "No se ha configurado el FORMATO FUA, alguno de los datos tiene valor CERO" & Chr(13) & "use la opción 'Herramientas --> Exporta/importa SIS'", vbInformation, Me.Caption
                     Me.Visible = False
                ElseIf oRsTmp1.Fields!FuaUltimoGenerado < 0 And Val(wxParametro320) = sghFuaTipo.sghFuaTipoAutomatico Then
                     MsgBox "No se ha configurado el FORMATO FUA, el ULTIMO NUMERO FUA GENERADO tiene que ser un número mayor o igual a CERO" & Chr(13) & "use la opción 'Herramientas --> Exporta/importa SIS'", vbInformation, Me.Caption
                     Me.Visible = False
                Else
                     txtFua1.Text = oRsTmp1.Fields!FuaDisa
                     txtFua2.Text = oRsTmp1.Fields!FuaLote
                     txtFua3.Text = ""
                End If
            End If
        End If

    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Sub

Sub CargaDatosDelFuaDesdeTablasGalenHos()
    Dim oConexion As New Connection
    Dim oConexionExterna As New Connection
    Dim oRsTmp187 As New Recordset
    Dim oDOEstablecimiento As New DOEstablecimiento
    Dim oDOEstablecimientoNoMinsa As New DOEstablecimientoNoMinsa
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.Open wxParametroJAMO
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ml_IdConceptoPrestacional = ""
    If CargarDatosDelPaciente(oConexion) = True Then
        'Paciente Hospitalizado proviene de EMERGENCIA
        lnIdCuentaAtencionEmergenciaOce = 0:  lblCtaEmergencia.Caption = ""
        If mo_ReglasComunes.HospitalizacionConOrigenEmergenciaOconsultorios(mo_lnIdTablaLISTBARITEMS, ml_IdOrigenAtencion) = True Then
           lnIdCuentaAtencionEmergenciaOce = mo_ReglasComunes.DevuelveCuentaEmergenciaOceDelPacienteHospitalizado(ml_idPaciente, _
                                                           ml_fechaIngreso, lnIdAtencionEmergenciaOce, ml_IdOrigenAtencion)
           lblCtaEmergencia.Caption = IIf(ml_IdOrigenAtencion = 30, "Cta CE: ", "Cta Emerg: ") & Trim(Str(lnIdCuentaAtencionEmergenciaOce))
           Set oRsTmp187 = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(lnIdCuentaAtencionEmergenciaOce, oConexion)
           If oRsTmp187.RecordCount > 0 Then
                txtHfingreso.Text = Format(oRsTmp187!fechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY)
                If Not IsNull(oRsTmp187!idEstablecimientoOrigen) Or Not IsNull(oRsTmp187!IdEstablecimientoNoMinsaOrigen) Then
                     If Not IsNull(oRsTmp187!idEstablecimientoOrigen) Then
                         Set oDOEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(oRsTmp187!idEstablecimientoOrigen)
                         If Not oDOEstablecimiento Is Nothing Then
                             txtRO.Text = oDOEstablecimiento.nombre
                             Me.txtROcodigo.Text = Right("0000000000" & oDOEstablecimiento.codigo, 10)
                         End If
                         Me.txtRONumero.Text = IIf(IsNull(oRsTmp187!nroReferenciaOrigen), "", oRsTmp187!nroReferenciaOrigen)
                     Else
                         Set oDOEstablecimientoNoMinsa = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(oRsTmp187!IdEstablecimientoNoMinsaOrigen)
                         If Not oDOEstablecimiento Is Nothing Then
                             txtRO.Text = oDOEstablecimientoNoMinsa.nombre
                             Me.txtROcodigo.Text = Right("0000000000" & oDOEstablecimientoNoMinsa.codigo, 10)
                         End If
                     End If
                     Me.txtRONumero.Text = IIf(IsNull(oRsTmp187!nroReferenciaOrigen), "", oRsTmp187!nroReferenciaOrigen)
                     Me.btnBuscarEstablecimientoO.Enabled = False
                     mo_Formulario.HabilitarDeshabilitar Me.txtRONumero, False
                     chkAtencionReferencia.Value = ssCBChecked
                End If
           End If
           oRsTmp187.Close
        Else
           lnIdAtencionEmergenciaOce = 0
        End If
        '
        Select Case mo_lnIdTablaLISTBARITEMS
        Case sghRegistroCitaCE, sghRegistroAtencionCE
             If CDate(ml_fechaIngreso) < ldHoy Then
                lblCtaEmergencia.Caption = lblCtaEmergencia.Caption & " <> El FUA no fué registrado en la CITA " & ml_fechaIngreso
                MsgBox lblCtaEmergencia, vbInformation, Me.Caption
             End If
        Case sghAdmisionEmergencia, sghAdmisionHospitalizacion
             If Me.txtHfalta.Text <> sighentidades.FECHA_VACIA_DMY Then
                If CDate(Me.txtHfalta.Text) < ldHoy Then
                   lblCtaEmergencia.Caption = lblCtaEmergencia.Caption & " <> El FUA no fué registrado el  " & Me.txtHfalta.Text
                   MsgBox lblCtaEmergencia, vbInformation, Me.Caption
                End If
             End If
        End Select
        '
        CargaFormatoFUA
        txtCScodigo.Text = wxParametro280           '?
        txtCS.Text = wxParametro205
        Me.txtPaciente = ml_Paciente
        txtNdocumento.Text = ml_NroDocumento
        txtFnacimiento.Text = Format(md_FechaNacimiento, sighentidades.DevuelveFechaSoloFormato_DMY)
        txtSexo.Text = ml_Sexo
        If ml_Etnia = "" Then
           cmbEtnia.Text = lcBuscaParametro.SeleccionaFilaParametro(283)
        Else
           cmbEtnia_UbicaPosicion (ml_Etnia)
        End If
        txtNhistoriaClinica.Text = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(ml_NroHistoriaClinica)), False)
        If ml_HoraAtencion <> "" Then
           txtHatencion.Text = ml_HoraAtencion
           txtFantencion.Text = Format(md_FechaAtencion, sighentidades.DevuelveFechaSoloFormato_DMY)
        End If
        '
        CargaDatosDeDx oConexion, True
        CargaDatosMedico oConexion, False
        CargaConsumosEnServiciosIntermedios oConexion, True
        CargaDatosDeAfiliacion True
        CargaDatosDeTriajeVacunas oConexionExterna, True, oConexion
        CargaDatosDeNacimiento oConexion, True
        CPTesPAQUETEdisminuyeMedicamentosInsumos
        ChequeaQueRecetadoNoSeaMenorAdespachado
        '
        If lnIdCuentaAtencionEmergenciaOce > 0 Then
           AsignaDx_SegunCptYfarmaciaDeRecetas oConexion, lnIdCuentaAtencionEmergenciaOce
        End If
        AsignaDx_SegunCptYfarmaciaDeRecetas oConexion, ml_IdCuentaAtencion
        '
    Else
        Me.btnAceptar.Enabled = False
        Me.btnImprimir.Enabled = False
    End If
    oConexion.Close
    oConexionExterna.Close
    Set oConexion = Nothing
    Set oConexionExterna = Nothing
    Set oRsTmp187 = Nothing
    Set oDOEstablecimiento = Nothing
    Set oDOEstablecimientoNoMinsa = Nothing
End Sub

Sub HabilitaTextosParaCRED()
    If cmbUPSfua.Text = "301202" Then
       Me.btnAddFarmacia.Visible = True
       btnAddPatologia.Visible = True
       mo_Formulario.HabilitarDeshabilitar Frame(0), False
       
       mo_Formulario.HabilitarDeshabilitar Frame(3), True
       mo_Formulario.HabilitarDeshabilitar Me.frPartoVertical, True
       mo_Formulario.HabilitarDeshabilitar txtSPcpn, True
       mo_Formulario.HabilitarDeshabilitar txtSPedadG, True
       mo_Formulario.HabilitarDeshabilitar txtSPalturaU, True
       mo_Formulario.HabilitarDeshabilitar chkSPPartoVertSI, True
       mo_Formulario.HabilitarDeshabilitar chkSPPartoVertNO, True
       mo_Formulario.HabilitarDeshabilitar txtSPpuerperio, True
       Frame(5).Enabled = True
       mo_Formulario.HabilitarDeshabilitar txtSPedadGrn, True
       mo_Formulario.HabilitarDeshabilitar chkSPCorTarCordonSI, True
       mo_Formulario.HabilitarDeshabilitar chkSPCorteTarCordonNO, True
       mo_Formulario.HabilitarDeshabilitar fraCorTardio, True
       mo_Formulario.HabilitarDeshabilitar chkSPCorTarCordonSI, True
       mo_Formulario.HabilitarDeshabilitar chkSPCorteTarCordonNO, True
       mo_Formulario.HabilitarDeshabilitar txtSPapgar1, True
       mo_Formulario.HabilitarDeshabilitar txtSPapgar5, True
       Frame(6).Enabled = True
       mo_Formulario.HabilitarDeshabilitar txtSPcred, True
       mo_Formulario.HabilitarDeshabilitar txtSPPAB, True
       mo_Formulario.HabilitarDeshabilitar fraRnPrematuro, True
       mo_Formulario.HabilitarDeshabilitar chkSPRNPrematuroSI, True
       mo_Formulario.HabilitarDeshabilitar chkSPRNPrematuroNO, True
       mo_Formulario.HabilitarDeshabilitar fraEEDP, True
       mo_Formulario.HabilitarDeshabilitar chkSPeedpSI, True
       mo_Formulario.HabilitarDeshabilitar chkSPeedpNO, True
       mo_Formulario.HabilitarDeshabilitar fraBajoPesoNacer, True
       mo_Formulario.HabilitarDeshabilitar chkSBajoPesoSI, True
       mo_Formulario.HabilitarDeshabilitar chkSBajoPesoNO, True
       mo_Formulario.HabilitarDeshabilitar fraConsejNutricional, True
       mo_Formulario.HabilitarDeshabilitar chkSPconsejeriaNsi, True
       mo_Formulario.HabilitarDeshabilitar chkSPconsejeriaNno, True
       mo_Formulario.HabilitarDeshabilitar fraSecuelaNacer, True
       mo_Formulario.HabilitarDeshabilitar chkSPSecuelaNaceSI, True
       mo_Formulario.HabilitarDeshabilitar chkSPSecuelaNaceNO, True
       mo_Formulario.HabilitarDeshabilitar fraConIntegral, True
       mo_Formulario.HabilitarDeshabilitar chkSPConIntegralSI, True
       mo_Formulario.HabilitarDeshabilitar chkSPConIntegralNO, True
       mo_Formulario.HabilitarDeshabilitar txtSPNFamGestante, True
       mo_Formulario.HabilitarDeshabilitar fraEvalIntegral, True
       mo_Formulario.HabilitarDeshabilitar chkSPEvalIntegralSI, True
       mo_Formulario.HabilitarDeshabilitar chkSPEvalIntegralNO, True
       mo_Formulario.HabilitarDeshabilitar txtSPVacam, True
       mo_Formulario.HabilitarDeshabilitar fraTamizajeSaludM, True
       mo_Formulario.HabilitarDeshabilitar chkSPTamizajeSalMPAT, True
       mo_Formulario.HabilitarDeshabilitar chkSPTamizajeSalMNOR, True
       mo_Formulario.HabilitarDeshabilitar Frame5, True
       mo_Formulario.HabilitarDeshabilitar txtVacBcg, True
       mo_Formulario.HabilitarDeshabilitar txtVacInfluenz, True
       mo_Formulario.HabilitarDeshabilitar txtVacAntiamarilica, True
       mo_Formulario.HabilitarDeshabilitar txtVacDpt, True
       mo_Formulario.HabilitarDeshabilitar txtVacParotid, True
       mo_Formulario.HabilitarDeshabilitar txtVacAntineumoc, True
       mo_Formulario.HabilitarDeshabilitar txtVacApo, True
       mo_Formulario.HabilitarDeshabilitar txtVacRubeola, True
       mo_Formulario.HabilitarDeshabilitar txtVacAntitetanica, True
       mo_Formulario.HabilitarDeshabilitar txtVacAsa, True
       mo_Formulario.HabilitarDeshabilitar txtVacRotavirus, True
       mo_Formulario.HabilitarDeshabilitar chkVacCompEdSI, True
       mo_Formulario.HabilitarDeshabilitar chkVacCompEdNo, True
       mo_Formulario.HabilitarDeshabilitar txtVacSpr, True
       mo_Formulario.HabilitarDeshabilitar txtVacDt, True
       mo_Formulario.HabilitarDeshabilitar txtVacVPH, True
       mo_Formulario.HabilitarDeshabilitar txtVacSR, True
       mo_Formulario.HabilitarDeshabilitar txtVacIPV, True
       mo_Formulario.HabilitarDeshabilitar txtVacOtraVacuna, True
       mo_Formulario.HabilitarDeshabilitar txtVacHVB, True
       mo_Formulario.HabilitarDeshabilitar txtVacPentaval, True
       mo_Formulario.HabilitarDeshabilitar txtVacRiesgoHVB, True
    End If
End Sub

Sub cmbEtnia_UbicaPosicion(lcCodigoEtnia As String)
    Dim lnFor As Integer
    For lnFor = 0 To (cmbEtnia.ListCount - 1)
        cmbEtnia.ListIndex = lnFor
        If cmbEtnia.SubItem(cmbEtnia.ListIndex, 0) = Val(lcCodigoEtnia) Then
           Exit For
        End If
    Next
End Sub

Sub CargaDatosDeDx(oConexion As Connection, lbDesdeGalenHos As Boolean)
    Dim oRsTmp1 As New Recordset
    Dim lnDxNro As Integer, lnUno As Integer
    If lbDesdeGalenHos = True Then
        lnUno = 1
        mo_ReglasSISgalenhos.FuaCargaDxDesdeGAlenHos oRsDx, oConexion, ml_idAtencion, ml_IdTipoServicio, lcDxPrincipal, _
                                                     lcDxPrincipalNro, lnUno, True
        'Paciente Hospitalizado proviene de EMERGENCIA
        If lnIdAtencionEmergenciaOce > 0 Then
           mo_ReglasSISgalenhos.FuaCargaDxDesdeGAlenHos oRsDx, oConexion, lnIdAtencionEmergenciaOce, _
                                                        sghTipoServicio.sghEmergenciaConsultorios, lcDxPrincipal, _
                                                        lcDxPrincipalNro, lnUno, False
        End If
        'Carga Dx de ingreso (paciente Referido Origen)
        If mo_lnIdTablaLISTBARITEMS = sghPacienteExternoConSeguro And _
                                                                    lnIdDiagnosticoPacExtSeguro > 0 Then
            Dim lcDxCodigo As String, lcDxDescripcion As String
            mo_AdminServiciosComunes.DiagnosticosSeleccionarPorIdDevuelveDescripcion lnIdDiagnosticoPacExtSeguro, _
                                                        oConexion, lcDxCodigo, lcDxDescripcion
            If lcDxCodigo <> "" Then
                lcDxCodigo = sighentidades.DevuelveCodigoDxSinPUNTO(lcDxCodigo)
                lcDxPrincipal = lcDxCodigo
                lcDxPrincipalNro = 1
                oRsDx.MoveFirst
                oRsDx.Fields!DxIngresoDefinitivo = True
                oRsDx.Fields!dxIngreso = lcDxCodigo
                oRsDx.Fields!Descripcion = lcDxDescripcion
                oRsDx.Update
            End If
        End If
    Else
        Dim lcDxPrincipal2 As String, lcDxPrincipalNro2 As Long
        Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaAtencionDIAxIdCuentaAtencion(ml_IdCuentaAtencion)
        lcDxPrincipal2 = "": lcDxPrincipalNro2 = 0
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              oRsDx.MoveFirst
              oRsDx.Find "DxNro=" & oRsTmp1.Fields!DxNumero
              If oRsDx.EOF Then
                MsgBox "Cargar Dx: No se encontro Dx Nro: " & oRsTmp1.Fields!DxNumero
              Else
                oRsDx.Fields!id = oRsTmp1.Fields!id
                oRsDx.Fields!Descripcion = mo_AdminServiciosComunes.DiagnosticosSeleccionarXcodigoCIEsinPtoDevuelveDescripcion(oRsTmp1.Fields!DxCodigo)
                If oRsTmp1.Fields!DxTipoIE = "I" Then
                   oRsDx.Fields!dxIngreso = oRsTmp1.Fields!DxCodigo
                   oRsDx.Fields!DxIngresoPresuntivo = IIf(oRsTmp1.Fields!DxTipoDPR = "P", True, False)
                   oRsDx.Fields!DxIngresoDefinitivo = IIf(oRsTmp1.Fields!DxTipoDPR = "D", True, False)
                   oRsDx.Fields!DxIngresoRepetido = IIf(oRsTmp1.Fields!DxTipoDPR = "R", True, False)
                Else
                   oRsDx.Fields!DxEgreso = oRsTmp1.Fields!DxCodigo
                   oRsDx.Fields!DxEgresoRepetido = IIf(oRsTmp1.Fields!DxTipoDPR <> "D", True, False)
                   oRsDx.Fields!DxEgresoDefinitivo = IIf(oRsTmp1.Fields!DxTipoDPR = "D", True, False)
                End If
                If oRsTmp1.Fields!DxTipoDPR = "D" Then
                   lcDxPrincipal = oRsTmp1.Fields!DxCodigo
                   lcDxPrincipalNro = oRsTmp1.Fields!DxNumero
                End If
                oRsDx.Update
              End If
              lcDxPrincipal2 = oRsTmp1.Fields!DxCodigo
              lcDxPrincipalNro2 = oRsTmp1.Fields!DxNumero
              oRsTmp1.MoveNext
           Loop
           oRsDx.Sort = "dxNro"
           If lcDxPrincipal = "" And ml_IdTipoServicio = sghConsultaExterna Then
              lcDxPrincipal = lcDxPrincipal2
              lcDxPrincipalNro = lcDxPrincipalNro2
           End If
        End If
        oRsTmp1.Close
    End If
    Set oRsTmp1 = Nothing
    'debb-20/07/2016
    On Error Resume Next
    With grdFarmacia.ValueLists.Add("DxPrincipal1").ValueListItems
       lnDxNro = 1
       oRsDx.MoveFirst
       Do While Not oRsDx.EOF
          If Not IsNull(oRsDx.Fields!DxEgreso) Then
             ' .Add Trim(Str(oRsDx.Fields!dxNro)), oRsDx.Fields!DxEgreso
              .Add Trim(oRsDx.Fields!DxEgreso), oRsDx.Fields!DxEgreso
          ElseIf Not IsNull(oRsDx.Fields!dxIngreso) Then
              .Add Trim(oRsDx.Fields!dxIngreso), oRsDx.Fields!dxIngreso
          End If
          oRsDx.MoveNext
       Loop
    End With
    grdFarmacia.Bands(0).Columns("dx").ValueList = "DxPrincipal1"
    With grdPatologia.ValueLists.Add("DxPrincipal2").ValueListItems
       lnDxNro = 1
       oRsDx.MoveFirst
       Do While Not oRsDx.EOF
          If Not IsNull(oRsDx.Fields!DxEgreso) Then
             .Add Trim(oRsDx.Fields!DxEgreso), oRsDx.Fields!DxEgreso
          ElseIf Not IsNull(oRsDx.Fields!dxIngreso) Then
              .Add Trim(oRsDx.Fields!dxIngreso), oRsDx.Fields!dxIngreso
          End If
          oRsDx.MoveNext
       Loop
    End With
    grdPatologia.Bands(0).Columns("dx").ValueList = "DxPrincipal2"
    
    '
End Sub

'Frank 19092014
Function ColocarFormatoPresAtmosferica(lcPresion As String) As String
    Dim lnUbicaSeparador As Integer
    Dim lcPrimerValor As String
    Dim lcSegundoValor As String
    If lcPresion = "" Then
        ColocarFormatoPresAtmosferica = "___/___"
    Else
        If Len(lcPresion) <= 7 Then
            lnUbicaSeparador = InStr(lcPresion, "/")
            If lnUbicaSeparador = -1 Then
                ColocarFormatoPresAtmosferica = "___/___"
            Else
                lcPrimerValor = Mid(lcPresion, 1, lnUbicaSeparador - 1)
                lcSegundoValor = Mid(lcPresion, lnUbicaSeparador + 1, Len(lcPresion) - lnUbicaSeparador)
                ColocarFormatoPresAtmosferica = Right("___" & lcPrimerValor, 3) & "/" & Left(lcSegundoValor & "___", 3)
            End If
        End If
    End If
End Function

'debb-06/08/2015
Sub CargaDatosDeTriajeVacunas(oConexionExterna As Connection, lbDesdeGalenHos As Boolean, oConexion As Connection)
    If lbDesdeGalenHos = True Then
        Dim lcTxtSpPeso As String, lcTxtSPtalla As String, lcTxtSPpa As String, lcTxtObservaciones As String
        mo_ReglasSISgalenhos.FuaCargaTriajeVacunasDesdeGAlenHos lcTxtSpPeso, lcTxtSPtalla, lcTxtSPpa, lcTxtObservaciones, _
                                                                ml_idAtencion, oConexionExterna
        txtSPpeso.Text = lcTxtSpPeso
        txtSPtalla.Text = lcTxtSPtalla
        txtSPpa.Text = ColocarFormatoPresAtmosferica(lcTxtSPpa)
        If lcTxtObservaciones <> "" Then
           mo_Formulario.HabilitarDeshabilitar txtObservaciones, False
        End If
        Me.txtObservaciones.Text = lcTxtObservaciones
        '
        oRsPatologia.Filter = "tipo='" & lcOtros & "'"
        If oRsPatologia.RecordCount > 0 Then
            Dim oRsTmp9 As New Recordset
            Dim lcSistolica9 As String
            Set oRsTmp9 = mo_ReglasFacturacion.EquivalenciaCPT_SMIseleccionarTodos
            If oRsTmp9.RecordCount > 0 Then
               oRsTmp9.MoveFirst
               Do While Not oRsTmp9.EOF
                  oRsPatologia.MoveFirst
                  oRsPatologia.Find "codigo='" & oRsTmp9!codigoCPT & "'"
                  If Not oRsPatologia.EOF Then
                     If CargaVacunaYsp(oRsTmp9!codigoSMI, Trim(Str(oRsPatologia!indicado)), lcSistolica9) = True Then
                        oRsPatologia.Delete
                        oRsPatologia.Update
                        If oRsPatologia.RecordCount = 0 Then
                           Exit Do
                        End If
                     End If
                  End If
                  oRsTmp9.MoveNext
               Loop
            End If
            oRsTmp9.Close
            Set oRsTmp9 = Nothing
       End If
       oRsPatologia.Filter = ""
       '
    Else
       Dim oRsTmp1 As New Recordset, lcSistolica As String, lcPresion As String
       Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaAtencionSMIxIdCuentaAtencion(ml_IdCuentaAtencion)
       If oRsTmp1.RecordCount > 0 Then
           txtSPpa.Text = sighentidades.PresionDevuelveVacia
           lcSistolica = "___"
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              CargaVacunaYsp oRsTmp1.Fields!IntervencionesPreventivas, oRsTmp1.Fields!Valor, lcSistolica
              oRsTmp1.MoveNext
           Loop
        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
    End If
    'carga LAB si UPS=CRED
    If cmbUPSfua.Text = "301202" Then
       Dim oRstmp99 As New Recordset
       Dim oRsTmp91 As New Recordset
       Dim oRstmp92 As New Recordset
       Dim lnCuenta99 As Long, lnIdAtencion99 As Long, lnCorrelativo As Long
       lnCorrelativo = mo_AdminAdmision.ServiciosAtenSimultaneaMovXidatencion(ml_idAtencion, oConexion)
       Set oRstmp92 = mo_AdminAdmision.ServiciosAtenSimultaneaMovXcorrelativo(lnCorrelativo, True)
       If oRstmp92.RecordCount > 0 Then
          oRstmp92.MoveFirst
          Do While Not oRstmp92.EOF
             lnCuenta99 = oRstmp92!IdCuentaAtencion
             lnIdAtencion99 = oRstmp92!idAtencion
             '
             Set oRstmp99 = mo_AdminAdmision.BuscaAtencionesDxCEparaFormatoHIS(lnIdAtencion99)
             If oRstmp99.RecordCount > 0 Then
                oRstmp99.MoveFirst
                Do While Not oRstmp99.EOF
                   Select Case UCase(Trim(oRstmp99!codigoCIE10))
                   Case "Z00.1"
                        If Not IsNull(oRstmp99!labConfHIS) Then
                            txtSPcred.Text = oRstmp99!labConfHIS
                        End If
                   End Select
                   oRstmp99.MoveNext
                Loop
             End If
             oRstmp99.Close
             '
             Set oRstmp99 = mo_AdminAdmision.BuscaAtencionesCptCEparaFormatoHIS(lnCuenta99, sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
             If oRstmp99.RecordCount > 0 Then
                   Set oRsTmp91 = mo_ReglasFacturacion.EquivalenciaCPT_SMIseleccionarTodos
                   oRsTmp91.MoveFirst
                   Do While Not oRsTmp91.EOF
                      oRstmp99.MoveFirst
                      oRstmp99.Find "codigo='" & oRsTmp91!codigoCPT & "'"
                      If Not oRstmp99.EOF Then
                         If Not IsNull(oRstmp99!labConfHIS) Then
                            If CargaVacunaYsp(oRsTmp91!codigoSMI, Trim(oRstmp99!labConfHIS), "") = True Then
                               If oRsPatologia.RecordCount > 0 Then
                                  oRsPatologia.MoveFirst
                                  oRsPatologia.Find "codigo='" & oRstmp99!codigo & "'"
                                  If Not oRsPatologia.EOF Then
                                     oRsPatologia.Delete
                                     oRsPatologia.Update
                                  End If
                               End If
                            End If
                         End If
                      End If
                      oRsTmp91.MoveNext
                   Loop
                   oRsTmp91.Close
             End If
             oRstmp99.Close
             '
             oRstmp92.MoveNext
       Loop
       End If
       oRstmp92.Close
       Set oRstmp92 = Nothing
       Set oRsTmp91 = Nothing
       Set oRstmp99 = Nothing
    End If
End Sub

Sub CargaDatosDeAfiliacion(lbDesdeGalenHos As Boolean)
    If lbDesdeGalenHos = True Then
        Dim oRsAfiliadosSIS As New Recordset
        lcSql = " where idSiaSis=" & lcAfiliacionIdSiaSis & " and codigo='" & lcAfiliacionCodigo & "'"
        Set oRsAfiliadosSIS = mo_ReglasSISgalenhos.SisFiltraPacientesAfiliados(lcSql, wxParametroJAMO)
        If oRsAfiliadosSIS.RecordCount > 0 Then
           lcAfiliacionNroIntegrante = IIf(IsNull(oRsAfiliadosSIS.Fields!AfiliacionNroIntegrante), "", oRsAfiliadosSIS.Fields!AfiliacionNroIntegrante)
           lcAfiliacionCodigo = oRsAfiliadosSIS.Fields!codigo
           lcAfiliacionIdSiaSis = Trim(Str(oRsAfiliadosSIS.Fields!idSiasis))
           txtNroAfiliacion1.Text = oRsAfiliadosSIS.Fields!cDisa
           txtNroAfiliacion2.Text = oRsAfiliadosSIS.Fields!cFormato
           txtNroAfiliacion3.Text = oRsAfiliadosSIS.Fields!cNumero
           lcCodigoEstablecimientoAdscripcionSIS = IIf(IsNull(oRsAfiliadosSIS.Fields!CodigoEstablAdscripcion), "", oRsAfiliadosSIS.Fields!CodigoEstablAdscripcion)
           Set oRsAfiliadosSIS = mo_ReglasSISgalenhos.SisA_LotesSeleccionarXCodigo(oRsAfiliadosSIS.Fields!codigo)
           If oRsAfiliadosSIS.RecordCount > 0 Then
              If Not IsNull(oRsAfiliadosSIS.Fields!lot_IdTipoFormato) Then
                 AsignaTipoAfiliacion oRsAfiliadosSIS.Fields!lot_IdTipoFormato
              End If
           End If
        End If
        oRsAfiliadosSIS.Close
        Set oRsAfiliadosSIS = Nothing
    Else
    End If
End Sub

Sub Mantenimiento()
    Select Case mi_opcion
    Case sghAgregar
        Me.Caption = "Agregar FUA " & lcOpcion & " " & sighentidades.TipoServicioDevuelveNombreCorto(ml_IdTipoServicio)
        'IIf(ml_IdTipoServicio = sghConsultaExterna, " (CE)", IIf(ml_IdTipoServicio = sghHospitalizacion, " (Hosp)", IIf(ml_IdTipoServicio = sghEmergenciaConsultorios, " (Emerg)", " (ArfSIS)")))
    Case sghModificar
        Me.Caption = "Modificar FUA " & lcOpcion & " " & sighentidades.TipoServicioDevuelveNombreCorto(ml_IdTipoServicio)
        If mo_lnIdTablaLISTBARITEMS = sghFormatoFUA And (ml_IdTipoServicio = sghEmergenciaConsultorios Or ml_IdTipoServicio = sghHospitalizacion) Then     'DEBB-23/02/2017
           MsgBox "No se puede MODIFICAR desde aquí, hagalo desde el MODULO DE " & _
                   sighentidades.TipoServicioDevuelveNombreCorto(ml_IdTipoServicio), vbInformation, Me.Caption
                   
           Me.btnAceptar.Enabled = False
           Me.btnImprimir.Visible = False
        End If
    Case sghConsultar
        Me.Caption = "Consultar FUA " & lcOpcion & " " & sighentidades.TipoServicioDevuelveNombreCorto(ml_IdTipoServicio)
        Frame7.Enabled = False
        Frame1.Enabled = False
        Frame3.Enabled = False
        Frame16.Enabled = False
        Frame17.Enabled = False
        Frame4.Enabled = False
        Frame5.Enabled = False
        FraDx.Enabled = False
        FraFarmacia.Enabled = False
        FraPatologia.Enabled = False
        Frame19.Enabled = False
        fraInstitucionEducativa.Enabled = False
    Case sghEliminar
        Me.Caption = "Eliminar FUA " & lcOpcion & " " & sighentidades.TipoServicioDevuelveNombreCorto(ml_IdTipoServicio)
        Me.btnImprimir.Visible = False
    End Select
End Sub

Sub CargaComboBoxes()
        Dim lcFiltro As String
        Set mo_cmbIdDestinoAtencion.MiComboBox = cmbIdDestinoAtencion
        mo_cmbIdDestinoAtencion.BoundColumn = "des_idDestinoAsegurado"
        mo_cmbIdDestinoAtencion.ListField = "des_descripcion"
        Set mo_cmbIdDestinoAtencion.RowSource = mo_ReglasSISgalenhos.SisDestinoAtencionSeleccionarTodos
        '
        Set mo_cmbConceptoP.MiComboBox = cmbConceptoP
        mo_cmbConceptoP.BoundColumn = "mod_idModalidad"
        mo_cmbConceptoP.ListField = "mod_descripcion"
        Set mo_cmbConceptoP.RowSource = mo_ReglasSISgalenhos.SisConceptoPrestacionalSeleccionarTodos(True)
        '
        Set mo_cmbTipoDocumento.MiComboBox = cmbTipoDocumento
        mo_cmbTipoDocumento.BoundColumn = "ide_idTipoDocumento"
        mo_cmbTipoDocumento.ListField = "ide_descripcion"
        Set mo_cmbTipoDocumento.RowSource = mo_ReglasSISgalenhos.SisTiposDocumentosSeleccionarTodos
        '
        Set cmbEtnia.ListSource = mo_AdminServiciosComunes.EtniaHISseleccionarTodos()
        Set cmbUPSfua.ListSource = mo_ReglasComunes.SisFuaUPServiciosSeleccionarTodos
        
        Set mo_cmbColegioNivel.MiComboBox = cmbColegioNivel
        mo_cmbColegioNivel.BoundColumn = "IdNivel"
        mo_cmbColegioNivel.ListField = "Nivel"
        Set mo_cmbColegioNivel.RowSource = mo_ReglasSISgalenhos.SisFuaColegioSeccionSeleccionarTodos
        
        Set mo_cmbColegioGrado.MiComboBox = cmbColegioGrado
        
        Set mo_cmbColegioTurno.MiComboBox = cmbColegioTurno
        mo_cmbColegioTurno.BoundColumn = "IdTurno"
        mo_cmbColegioTurno.ListField = "Turno"
        Set mo_cmbColegioTurno.RowSource = mo_ReglasSISgalenhos.SisFuaColegioTurnoSeleccionarTodos
End Sub

Sub CargaDatosMedico(oConexion As Connection, lbPorDNI As Boolean)
        Dim oRsTmp1 As New Recordset
        If lbPorDNI = True Then
           Set oRsTmp1 = mo_ReglasDeProgMedica.MedicosSeleccionarPorDNI(Me.txtMedicoDni.Text)
        Else
           Set oRsTmp1 = mo_ReglasDeProgMedica.MedicosSeleccionarXIdMedico(Val(ml_IdMedico))
        End If
        
        txtMedicoDni.Text = ""
        txtMedico.Text = ""
        txtMedicoColegiatura.Text = ""
        txtMedicoEspecialidad.Text = ""
        If oRsTmp1.RecordCount > 0 Then
            txtMedicoDni.Text = oRsTmp1.Fields!DNI
            txtMedico.Text = Trim(oRsTmp1.Fields!ApellidoPaterno) & " " & Trim(oRsTmp1.Fields!ApellidoMaterno) & " " & oRsTmp1.Fields!Nombres
            txtMedicoColegiatura.Text = oRsTmp1.Fields!Colegiatura
            txtMedicoEspecialidad.Text = IIf(IsNull(oRsTmp1.Fields!TipoEmpleadoSIS), "", oRsTmp1.Fields!TipoEmpleadoSIS)  'debb2014b
            txtMedicoRNE.Text = IIf(IsNull(oRsTmp1.Fields!rne), "", oRsTmp1.Fields!rne)
            chkMedicoEgresado.Value = IIf(IsNull(oRsTmp1.Fields!Egresado), 0, IIf(oRsTmp1.Fields!Egresado = True, 1, 0))
        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
End Sub



Private Sub grdDiag_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     grdDiag.Bands(0).Columns("Id").Hidden = True
     grdDiag.Bands(0).Columns("DxIngresoPresuntivo").Hidden = True
     grdDiag.Bands(0).Columns("DxIngresoDefinitivo").Hidden = True
     grdDiag.Bands(0).Columns("DxIngresoDefinitivo").Hidden = True
     grdDiag.Bands(0).Columns("DxIngresoRepetido").Hidden = True
     grdDiag.Bands(0).Columns("DxEgresoRepetido").Hidden = True
     grdDiag.Bands(0).Columns("DxEgresoDefinitivo").Hidden = True
     grdDiag.Bands(0).Columns("DxNro").Width = 200
     grdDiag.Bands(0).Columns("Descripcion").Width = 1800
     grdDiag.Bands(0).Columns("dxIngreso").Width = 700
     grdDiag.Bands(0).Columns("DxEgreso").Width = 700
     grdDiag.Bands(0).Columns("Descripcion").Activation = ssActivationActivateNoEdit
     grdDiag.Bands(0).Columns("dxIngreso").Activation = ssActivationActivateNoEdit
     grdDiag.Bands(0).Columns("DxEgreso").Activation = ssActivationActivateNoEdit
End Sub

Private Sub grdDx_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
    Dim lcDescripcion As String
    Dim oRsTmp1 As New Recordset
    Dim oRow As SSRow
    Set oRow = grdDx.ActiveCell.Row
    Select Case grdDx.ActiveCell.Column.Key
    Case "DxIngresoPresuntivo"
         oRsDx.Fields("DxIngresoDefinitivo").Value = False
         oRsDx.Fields("DxIngresoRepetido").Value = False
    Case "DxIngresoDefinitivo"
         oRsDx.Fields("DxIngresoPresuntivo").Value = False
         oRsDx.Fields("DxIngresoRepetido").Value = False
         ActualizaDxParaFarmaciaServicios
    Case "DxIngresoRepetido"
         oRsDx.Fields("DxIngresoDefinitivo").Value = False
         oRsDx.Fields("DxIngresoPresuntivo").Value = False
    Case "DxIngreso"
         If Trim(oRsDx.Fields("DxIngreso").Value) <> "" And ((ml_IdTipoServicio = sghConsultaExterna And _
                                                        mo_lnIdTablaLISTBARITEMS = sghFormatoFUA) Or _
                                                        (mi_opcion = sghAgregar And _
                                                        mo_lnIdTablaLISTBARITEMS = sghFormatoFUA)) Then
            lcDescripcion = ""
            lcDxPrincipal = ""
            lcDxPrincipalNro = 0
            Set oRsTmp1 = mo_AdminServiciosComunes.DiagnosticosSeleccionarXcodigoCIEsinPto(oRsDx.Fields("DxIngreso").Value)
            If oRsTmp1.RecordCount > 0 Then
               lcDescripcion = oRsTmp1.Fields!Descripcion
               If oRsDx.Fields("DxIngresoDefinitivo").Value = True Or oRsDx.Fields("DxEgresoDefinitivo").Value = True Then
                  lcDxPrincipal = oRsDx.Fields("DxIngreso").Value
                  lcDxPrincipalNro = oRsDx.Fields!dxNro
               End If
            End If
            oRsTmp1.Close
            oRsDx.Fields("Descripcion").Value = lcDescripcion
            ActualizaDxParaFarmaciaServicios
         End If
    Case "DxEgreso"
         If Trim(oRsDx.Fields("DxEgreso").Value) <> "" And (mi_opcion = sghAgregar And _
                                                            mo_lnIdTablaLISTBARITEMS = sghFormatoFUA) Then
            lcDescripcion = ""
            lcDxPrincipal = ""
            lcDxPrincipalNro = 0
            Set oRsTmp1 = mo_AdminServiciosComunes.DiagnosticosSeleccionarXcodigoCIEsinPto(oRsDx.Fields("DxEgreso").Value)
            If oRsTmp1.RecordCount > 0 Then
               lcDescripcion = oRsTmp1.Fields!Descripcion
               If oRsDx.Fields("DxIngresoDefinitivo").Value = True Or oRsDx.Fields("DxEgresoDefinitivo").Value = True Then
                  lcDxPrincipal = oRsDx.Fields("DxEgreso").Value
                  lcDxPrincipalNro = oRsDx.Fields!dxNro
               End If
            End If
            oRsTmp1.Close
            oRsDx.Fields("Descripcion").Value = lcDescripcion
            ActualizaDxParaFarmaciaServicios
         End If
    Case "DxEgresoDefinitivo"
         oRsDx.Fields("DxEgresoRepetido").Value = False
         ActualizaDxParaFarmaciaServicios
    Case "DxEgresoRepetido"
         oRsDx.Fields("DxEgresoDefinitivo").Value = False
    End Select
    Set oRsTmp1 = Nothing

End Sub

Private Sub grdDx_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
         grdDx.Bands(0).Columns("Id").Hidden = True
         grdDx.Bands(0).Columns("DxNro").Activation = ssActivationActivateNoEdit
         grdDx.Bands(0).Columns("DxNro").Width = 500
         grdDx.Bands(0).Columns("Descripcion").Activation = ssActivationActivateNoEdit
         grdDx.Bands(0).Columns("Descripcion").Width = 4000
         grdDx.Bands(0).Columns("DxIngresoPresuntivo").Width = 1000
         grdDx.Bands(0).Columns("DxIngresoDefinitivo").Width = 1000
         grdDx.Bands(0).Columns("DxIngresoRepetido").Width = 1000
         grdDx.Bands(0).Columns("DxIngreso").Width = 1000
         grdDx.Bands(0).Columns("DxEgreso").Width = 1000
         grdDx.Bands(0).Columns("DxEgresoDefinitivo").Width = 1000
         grdDx.Bands(0).Columns("DxEgresoRepetido").Width = 1000
         'grdDx.Bands(0).Columns("DxEgreso").Activation = ssActivationActivateNoEdit
         'grdDx.Bands(0).Columns("DxEgresoDefinitivo").Activation = ssActivationActivateNoEdit
         'grdDx.Bands(0).Columns("DxEgresoRepetido").Activation = ssActivationActivateNoEdit
End Sub

Private Sub grdDx_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    Dim lnKeyCode As Integer
    lnKeyCode = KeyCode
    Select Case lnKeyCode
    Case vbKeyReturn
       SendKeys "{Tab}"
    Case Else
        On Error Resume Next
        lcCodigoDxBuscado = ""
        If grdDx.ActiveCell.Column.Key = "DxIngreso" Then
           lcCodigoDxBuscado = "DxIngreso"
        ElseIf grdDx.ActiveCell.Column.Key = "DxEgreso" Then
           lcCodigoDxBuscado = "DxEgreso"
        Else
           lcCodigoDxBuscado = "DxIngreso"
        End If
        AdministrarKeyPreview lnKeyCode
    End Select
End Sub
'debb-20/07/2016
Private Sub grdFarmacia_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    On Error Resume Next
    If Cell.Column.Key = "Dx" And Not IsNull(Cell.Row.Cells("Dx").Value) Then
        oRsDx.MoveFirst
        If ml_IdTipoServicio = sghConsultaExterna Then
           oRsDx.Find "DxIngreso='" & Cell.Row.Cells("Dx").Value & "'"
        Else
           oRsDx.Find "DxEgreso='" & Cell.Row.Cells("Dx").Value & "'"
           If oRsDx.EOF Then
              oRsDx.MoveFirst
              oRsDx.Find "DxIngreso='" & Cell.Row.Cells("Dx").Value & "'"
           End If
        End If
        Cell.Row.Cells("DxNro").Value = Trim(Str(oRsDx!dxNro))
    End If
End Sub

Private Sub grdFarmacia_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If btnAddFarmacia.Visible = True Then
    Else
        Cancel = True
    End If

End Sub

Private Sub grdFarmacia_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
         grdFarmacia.Bands(0).Columns("Id").Hidden = True
         grdFarmacia.Bands(0).Columns("Precio").Hidden = True
         grdFarmacia.Bands(0).Columns("DxNro").Hidden = True
         grdFarmacia.Bands(0).Columns("formaF").Hidden = True
         grdFarmacia.Bands(0).Columns("esPaquete").Hidden = True
         grdFarmacia.Bands(0).Columns("Tipo").Width = 1000
         grdFarmacia.Bands(0).Columns("codigo").Width = 1000
         grdFarmacia.Bands(0).Columns("MedicInsumo").Activation = ssActivationActivateNoEdit
         grdFarmacia.Bands(0).Columns("MedicInsumo").Width = 6800
         grdFarmacia.Bands(0).Columns("Recetado").Width = 500
         grdFarmacia.Bands(0).Columns("Cantidad").Width = 500
         grdFarmacia.Bands(0).Columns("Cantidad").Header.Caption = "Entregado"
         grdFarmacia.Bands(0).Columns("Dx").Width = 1000
         
End Sub

Private Sub grdFarmacia_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    Dim lnKeyCode As Integer
    lnKeyCode = KeyCode
    AdministrarKeyPreview lnKeyCode
End Sub
'debb-20/07/2016
Private Sub grdPatologia_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    On Error Resume Next
    If Cell.Column.Key = "Dx" And Not IsNull(Cell.Row.Cells("Dx").Value) Then
        oRsDx.MoveFirst
        If ml_IdTipoServicio = sghConsultaExterna Then
           oRsDx.Find "DxIngreso='" & Cell.Row.Cells("Dx").Value & "'"
        Else
           oRsDx.Find "DxEgreso='" & Cell.Row.Cells("Dx").Value & "'"
           If oRsDx.EOF Then
              oRsDx.MoveFirst
              oRsDx.Find "DxIngreso='" & Cell.Row.Cells("Dx").Value & "'"
           End If
        End If
        Cell.Row.Cells("DxNro").Value = Trim(Str(oRsDx!dxNro))
    End If

End Sub

Private Sub grdPatologia_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If btnAddPatologia.Visible = True Then
    Else
        Cancel = True
    End If
End Sub

Private Sub grdPatologia_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
         grdPatologia.Bands(0).Columns("Id").Hidden = True
         grdPatologia.Bands(0).Columns("Precio").Hidden = True
         grdPatologia.Bands(0).Columns("IdPuntoCarga").Hidden = True
         grdPatologia.Bands(0).Columns("DxNro").Hidden = True
         grdPatologia.Bands(0).Columns("Tipo").Width = 1000
         grdPatologia.Bands(0).Columns("codigo").Width = 800
         grdPatologia.Bands(0).Columns("Procedimiento").Activation = ssActivationActivateNoEdit
         grdPatologia.Bands(0).Columns("Procedimiento").Width = 7000
         grdPatologia.Bands(0).Columns("Indicado").Width = 500
         grdPatologia.Bands(0).Columns("Ejecutado").Width = 500
         grdPatologia.Bands(0).Columns("Dx").Width = 1000
End Sub

Private Sub grdPatologia_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    Dim lnKeyCode As Integer
    lnKeyCode = KeyCode
    AdministrarKeyPreview lnKeyCode

End Sub

Private Sub txtCodSeguro_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodSeguro
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtCodSeguro_LostFocus()
    chkAtencionAmbulatoria.SetFocus
End Sub

Private Sub txtFantencion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFantencion
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFantencion_LostFocus()
    If Not EsFecha(txtFantencion.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
        txtFantencion.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
    md_FechaAtencion = txtFantencion.Text
    ml_edad_En_Dias = sighentidades.EdadActualEnDias(CDate(Me.txtFnacimiento.Text), CDate(txtFantencion.Text & " " & txtHatencion.Text))
    ml_edad_En_YYYYMMDD = sighentidades.EdadActualEnFormatoYYYYMMDD(CDate(Me.txtFnacimiento.Text), CDate(txtFantencion.Text & " " & txtHatencion.Text))
    
End Sub

Private Sub txtFparto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFparto
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFFallecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFFallecimiento
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFparto_LostFocus()
    If Not EsFecha(txtFparto.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
        txtFparto.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFFallecimiento_LostFocus()
    If Not EsFecha(txtFFallecimiento.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
        txtFFallecimiento.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFua3_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFua3
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFua3_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If

End Sub

Private Sub txtFua3_LostFocus()
    'Me.ucSISfuaCodPrestacion1.FocusEnCodigoPrestacion
End Sub

Private Sub txtHatencion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHatencion
    AdministrarKeyPreview KeyCode

End Sub





Private Sub txtHatencion_LostFocus()
   If Not sighentidades.ValidaHora(txtHatencion.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            txtHatencion.Text = sighentidades.HORA_VACIA_HM
    Else
            ml_edad_En_Dias = sighentidades.EdadActualEnDias(CDate(Me.txtFnacimiento.Text), CDate(txtFantencion.Text & " " & txtHatencion.Text))
            ml_edad_En_YYYYMMDD = sighentidades.EdadActualEnFormatoYYYYMMDD(CDate(Me.txtFnacimiento.Text), CDate(txtFantencion.Text & " " & txtHatencion.Text))
            ml_HoraAtencion = txtHatencion.Text
    End If
End Sub

Private Sub txtHfalta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHfalta
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtHfalta_LostFocus()
    If Not EsFecha(txtHfalta.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
        txtHfalta.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
    AdministrarKeyPreview vbKeyF4
End Sub

Private Sub txtHfingreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHfingreso
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtHfingreso_LostFocus()
    If Not EsFecha(txtFantencion.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
        txtHfingreso.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If

End Sub

Private Sub txtInstitucion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtInstitucion
    AdministrarKeyPreview KeyCode

End Sub





Private Sub txtMedicoDni_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtMedicoDni
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtMedicoDni_LostFocus()
     If txtMedicoDni.Locked = False And Len(txtMedicoDni.Text) = 8 Then
        Dim oRsTmp1 As New Recordset
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set oRsTmp1 = mo_ReglasDeProgMedica.MedicosSeleccionarPorDNI(txtMedicoDni.Text)
        If oRsTmp1.RecordCount > 0 Then
           ml_IdMedico = oRsTmp1.Fields!idMedico
           CargaDatosMedico oConexion, False
        Else
           MsgBox "El DNI del Médico no existe", vbInformation, Me.Caption
           txtMedicoDni.SetFocus
        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
        oConexion.Close
        Set oConexion = Nothing
     End If
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtMonto
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtNautorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNautorizacion
    AdministrarKeyPreview KeyCode

End Sub






Private Sub txtNhistoriaClinica_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNhistoriaClinica
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNhistoriaClinica_LostFocus()
       ml_NroHistoriaClinica = Val(txtNhistoriaClinica.Text)
'      If txtNhistoriaClinica.Locked = False And mo_Teclado.TextoEsSoloNumeros(txtNhistoria.Text) And txtDatosDeCuenta.Text = "" Then
'        Dim oRsTmp1 As New ADODB.Recordset
'        Dim oDOPaciente As New sighComun.doPaciente
'        oDOPaciente.NroHistoriaClinica = txtNhistoria.Text
'        Set oRsTmp1 = mo_AdminAdmision.PacientesFiltrar(oDOPaciente)
'        If oRsTmp1.RecordCount > 0 Then
'           lcSql = "Where apPaterno='" & oDOPaciente.ApellidoPaterno & "' and apMaterno='" & oDOPaciente.ApellidoMaterno & _
'                         "' and pNombre='" & oDOPaciente.PrimerNombre & _
'                         IIf(oDOPaciente.SegundoNombre = "", "", "' and sNombre='" & oDOPaciente.SegundoNombre) & _
'                         "' and fNacimiento=CONVERT(DATETIME,'" & oDOPaciente.FechaNacimiento & "',103)" & _
'                         " and Sexo=" & IIf(oDOPaciente.idTipoSexo = 1, "1", "0")
'           Set oRsTmp1 = mo_ReglasSISgalenhos.SisFiltraPacientesAfiliados(lcSql, wxParametroJAMO)
'           If oRsTmp1.RecordCount > 0 Then
'               LlenaDatosPersonalesDesdeFiliacionesSIS oRsTmp1, True
'               oRsTmp1.Close
'               Set oRsTmp1 = Nothing
'           End If
'        Else
'            txtNombrePaciente.Text = ""
'        End If
'        Set oRsTmp1 = Nothing
'        Set oDOPaciente = Nothing
'      End If
         

End Sub




Private Sub txtNroAfiliacion1_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroAfiliacion1
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtNroAfiliacion2_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroAfiliacion2
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNroAfiliacion3_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroAfiliacion3
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNroAfiliacion3_LostFocus()
    If txtNroAfiliacion3.Locked = False And txtNroAfiliacion1.Text <> "" And txtNroAfiliacion2.Text <> "" And txtNroAfiliacion3.Text <> "" Then
       Dim oRsTmp1 As New Recordset
       Set oRsTmp1 = mo_ReglasSISgalenhos.SisFiltraPacientesAfiliados("Where afiliacionDisa='" & txtNroAfiliacion1.Text & _
                                                        "' and afiliacionTipoFormato='" & txtNroAfiliacion2.Text & _
                                                        "' and afiliacionNroFormato='" & txtNroAfiliacion3.Text & _
                                                        "'", wxParametroJAMO)
       If oRsTmp1.RecordCount = 0 Then
          oRsTmp1.Close
          Set oRsTmp1 = Nothing
          MsgBox "El COD.AFILIACION/INSCRIPCION no existe en la tabla FILIACIONES SIS", vbInformation, Me.Caption
          On Error Resume Next
          txtNroAfiliacion1.Text = ""
          txtNroAfiliacion1.SetFocus
          Exit Sub
       Else
          LlenaDatosPersonalesDesdeFiliacionesSIS oRsTmp1, False
          oRsTmp1.Close
          Set oRsTmp1 = Nothing
          If mo_lnIdTablaLISTBARITEMS = sghFormatoFUA Then
             ReglasDeConsistenciasAntesDeCargarFormulario
          End If
       End If
    End If
End Sub

Sub LlenaDatosPersonalesDesdeFiliacionesSIS(oRsTmpFiliacionesSIS As Recordset, lbLlenaTambienInscripcion As Boolean)
          txtPaciente.Text = Trim(oRsTmpFiliacionesSIS.Fields!apPaterno) & " " & Trim(oRsTmpFiliacionesSIS.Fields!apMaterno) & _
                             " " & oRsTmpFiliacionesSIS.Fields!Pnombre & " " & _
                             IIf(IsNull(oRsTmpFiliacionesSIS.Fields!sNombre), "", oRsTmpFiliacionesSIS.Fields!sNombre)
          txtFnacimiento.Text = oRsTmpFiliacionesSIS.Fields!Fnacimiento
          mo_cmbTipoDocumento.BoundText = IIf(IsNull(oRsTmpFiliacionesSIS.Fields!DNI), 2, 1)
          txtNdocumento.Text = IIf(IsNull(oRsTmpFiliacionesSIS.Fields!DNI), "", oRsTmpFiliacionesSIS.Fields!DNI)
          txtSexo.Text = IIf(oRsTmpFiliacionesSIS.Fields!Sexo = "1", lcMasculino, lcFemenino)
          md_FechaNacimiento = txtFnacimiento.Text
          ml_Paciente = txtPaciente.Text
          ml_NroDocumento = txtNdocumento.Text
          ml_Sexo = txtSexo.Text
          ml_ApellidoPaterno = oRsTmpFiliacionesSIS.Fields!apPaterno
          ml_ApellidoMaterno = oRsTmpFiliacionesSIS.Fields!apMaterno
          ml_PrimerNombre = oRsTmpFiliacionesSIS.Fields!Pnombre
          ml_SegundoNombre = IIf(IsNull(oRsTmpFiliacionesSIS.Fields!sNombre), "", oRsTmpFiliacionesSIS.Fields!sNombre)
          If lbLlenaTambienInscripcion = True Then
               txtNroAfiliacion1.Text = oRsTmpFiliacionesSIS.Fields!cDisa
               txtNroAfiliacion2.Text = oRsTmpFiliacionesSIS.Fields!cFormato
               txtNroAfiliacion3.Text = oRsTmpFiliacionesSIS.Fields!cNumero
          End If
          If txtHatencion.Text = sighentidades.HORA_VACIA_HM Then
             txtHatencion.Text = "00:01"
          End If
          ml_edad_En_Dias = sighentidades.EdadActualEnDias(oRsTmpFiliacionesSIS.Fields!Fnacimiento, CDate(txtFantencion.Text & " " & txtHatencion.Text))
          ml_edad_En_YYYYMMDD = sighentidades.EdadActualEnFormatoYYYYMMDD(oRsTmpFiliacionesSIS.Fields!Fnacimiento, CDate(txtFantencion.Text & " " & txtHatencion.Text))
          lcAfiliacionCodigo = oRsTmpFiliacionesSIS.Fields!codigo
          lcAfiliacionIdSiaSis = oRsTmpFiliacionesSIS.Fields!idSiasis
          lcCodigoEstablecimientoAdscripcionSIS = oRsTmpFiliacionesSIS.Fields!CodigoEstablAdscripcion
          PermitirManipularDatosSegunSexo
End Sub

Private Sub txtObservaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtObservaciones
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtRDcodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtRDcodigo
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtRDnumero_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtRDnumero
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtROcodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtROcodigo
    AdministrarKeyPreview KeyCode
End Sub



Private Sub txtRONumero_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtRONumero
    AdministrarKeyPreview KeyCode

End Sub









Private Sub txtSPalturaU_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSPalturaU
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtSPalturaU_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtSPapgar1_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSPapgar1
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtSPapgar1_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtSPapgar5_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSPapgar5
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtSPapgar5_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtSPcpn_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSPcpn
    AdministrarKeyPreview KeyCode

End Sub







Private Sub txtSPcpn_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtSPcred_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSPcred
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtSPcred_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtSPedadG_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSPedadG
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtSPedadG_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtSPedadGrn_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSPedadGrn
    AdministrarKeyPreview KeyCode

End Sub





Private Sub txtSPedadGrn_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub



Private Sub txtSPpa_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSPpa
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtSPpeso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSPpeso
    AdministrarKeyPreview KeyCode

End Sub







Private Sub txtSPpeso_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtSPpuerperio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSPpuerperio
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtSPpuerperio_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub



Private Sub txtSPtalla_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSPtalla
    AdministrarKeyPreview KeyCode

End Sub





Private Sub txtSPtalla_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacAntiamarilica_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacAntiamarilica
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtVacAntiamarilica_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacAntineumoc_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacAntineumoc
    AdministrarKeyPreview KeyCode

End Sub





Private Sub txtVacAntineumoc_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacAntitetanica_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacAntitetanica
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtVacAntitetanica_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacApo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacApo
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtVacApo_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacAsa_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacAsa
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtVacAsa_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacBcg_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacBcg
    AdministrarKeyPreview KeyCode

End Sub





Private Sub txtVacBcg_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacDpt_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacDpt
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtVacDpt_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacDt_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacDt
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtVacDt_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub


Private Sub txtVacIPV_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacIPV
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtVacIPV_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
End Sub

Private Sub txtVacRiesgoHVB_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacRiesgoHVB
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtVacRiesgoHVB_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacHVB_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacHVB
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtVacHVB_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacInfluenz_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacInfluenz
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtVacInfluenz_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacParotid_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacParotid
    AdministrarKeyPreview KeyCode

End Sub







Private Sub txtVacParotid_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacPentaval_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacPentaval
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtVacPentaval_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacRotavirus_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacRotavirus
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtVacRotavirus_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacRubeola_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacRubeola
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtVacRubeola_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtVacSpr_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacSpr
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtVacSpr_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub


Sub CreaTemporales()
    mo_ReglasSISgalenhos.FUAcreaTemporales oRsDx, oRsPatologia, oRsFarmacia, oRsVacunasSp
    Set grdDx.DataSource = oRsDx
    mo_Apariencia.ConfigurarFilasBiColores grdDx, sighentidades.GrillaConFilasBicolor
    grdDx.Caption = ""
    FraDx.Enabled = True
    oRsDx.MoveFirst
    '
    Set grdPatologia.DataSource = oRsPatologia
    mo_Apariencia.ConfigurarFilasBiColores grdPatologia, sighentidades.GrillaConFilasBicolor
    grdPatologia.Caption = ""
    FraPatologia.Enabled = True
    '
    Set grdFarmacia.DataSource = oRsFarmacia
    mo_Apariencia.ConfigurarFilasBiColores grdFarmacia, sighentidades.GrillaConFilasBicolor
    grdFarmacia.Caption = ""
    FraFarmacia.Enabled = True
End Sub

Sub AsignaDx_SegunCptYfarmaciaDeRecetas(oConexion As Connection, lnCuenta11 As Long)
    Dim oRsTmp33 As New Recordset
    Dim lcDx33 As String
    Set oRsTmp33 = mo_AdminServiciosComunes.RecetaDetallePorCuenta(lnCuenta11, oConexion)
    If oRsTmp33.RecordCount > 0 Then
       oRsTmp33.MoveFirst
       Do While Not oRsTmp33.EOF
          If Not IsNull(oRsTmp33!dx) Then
                lcDx33 = sighentidades.DevuelveCodigoDxSinPUNTO(Trim(oRsTmp33!dx))
                oRsDx.MoveFirst
                oRsDx.Find "dxIngreso='" & lcDx33 & "'"
                If Not oRsDx.EOF Then
                   If oRsTmp33!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaFarmacia Then
                      If oRsFarmacia.RecordCount > 0 Then
                         oRsFarmacia.MoveFirst
                         oRsFarmacia.Find "id=" & oRsTmp33!idItem
                         If Not oRsFarmacia.EOF Then
                            oRsFarmacia!dx = lcDx33
                            oRsFarmacia.Update
                         End If
                      End If
                   Else
                      If oRsPatologia.RecordCount > 0 Then
                         oRsPatologia.MoveFirst
                         oRsPatologia.Find "id=" & oRsTmp33!idItem
                         If Not oRsPatologia.EOF Then
                            oRsPatologia!dx = lcDx33
                            oRsPatologia.Update
                         End If
                      End If
                   End If
                End If
          End If
          oRsTmp33.MoveNext
       Loop
    End If
    oRsTmp33.Close
    Set oRsTmp33 = Nothing
End Sub

'debb-20/07/2016
Sub CargaConsumosEnServiciosIntermedios(oConexion As Connection, lbDesdeGalenHos As Boolean)
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oRsTmp4 As New Recordset
    Dim lnRecetado As Long, lnIdPuntoCarga As Long, lcPuntoCarga As String, lbSeGraboUnCptOFarmaciaEnTablasSIS As Boolean
    If lbDesdeGalenHos = True Then
        mo_ReglasSISgalenhos.FuaCargaSIDesdeGAlenHos oRsFarmacia, oRsPatologia, mo_lnIdTablaLISTBARITEMS, ml_IdCuentaAtencion, _
                             lcInsumo, lcMedicamento, lcDxPrincipal, lcDxPrincipalNro, lcOtros, lcLaboratorio, lcImagenes
        'Paciente Hospitalizado proviene de EMERGENCIA
        If lnIdAtencionEmergenciaOce > 0 Then
           mo_ReglasSISgalenhos.FuaCargaSIDesdeGAlenHos oRsFarmacia, oRsPatologia, mo_lnIdTablaLISTBARITEMS, lnIdCuentaAtencionEmergenciaOce, _
                             lcInsumo, lcMedicamento, lcDxPrincipal, lcDxPrincipalNro, lcOtros, lcLaboratorio, lcImagenes
        End If
        '
        mo_ReglasSISgalenhos.FuaPaquetesFarmaciaDesagregaEnMedicInsumos oRsFarmacia, mo_lnIdTablaLISTBARITEMS
        '
        
    Else
        lbSeGraboUnCptOFarmaciaEnTablasSIS = False
        mo_ReglasSISgalenhos.FuaActualizaDespachosEnFarmacia ml_IdCuentaAtencion, wxParametro302, ml_IdTipoServicio, sghFuenteFinanciamiento.sghFFSIS
        mo_ReglasSISgalenhos.FuaActualizaDespachosEnServicios ml_IdCuentaAtencion, wxParametro302, ml_IdTipoServicio, sghFuenteFinanciamiento.sghFFSIS
        '*********************Farmacia - Medicamentos - Desde el SIS
        Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaAtencionMEDxIdCuentaAtencion(ml_IdCuentaAtencion)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
                 Set oRsTmp4 = mo_AdminServiciosComunes.MedicamentosInsumosSeleccionarPorCodigo(oRsTmp1.Fields!codigo)
                 If oRsTmp4.RecordCount > 0 Then 'debb2014b
                    lbSeGraboUnCptOFarmaciaEnTablasSIS = True
                    oRsDx.MoveFirst
                    oRsDx.Find "dxNro=" & oRsTmp1!DxNumero
                    oRsFarmacia.AddNew
                    oRsFarmacia.Fields!id = oRsTmp1.Fields!id
                    oRsFarmacia.Fields!tipo = lcMedicamento
                    oRsFarmacia.Fields!MedicInsumo = IIf(IsNull(oRsTmp4.Fields!nombre), "", oRsTmp4.Fields!nombre)
                    oRsFarmacia.Fields!recetado = oRsTmp1.Fields!CantidadPrescrita
                    oRsFarmacia.Fields!cantidad = oRsTmp1.Fields!CantidadEntregada
                    oRsFarmacia.Fields!dx = IIf(oRsDx.EOF, lcDxPrincipal, IIf(IsNull(oRsDx!DxEgreso), oRsDx!dxIngreso, oRsDx!DxEgreso))   'debb-23/02/2017
                    oRsFarmacia.Fields!Precio = oRsTmp1.Fields!PrecioUnitario
                    oRsFarmacia.Fields!codigo = oRsTmp1.Fields!codigo
                    oRsFarmacia.Fields!dxNro = IIf(oRsDx.EOF, lcDxPrincipalNro, oRsTmp1!DxNumero)
                    oRsFarmacia.Fields!formaF = IIf(IsNull(oRsTmp4.Fields!FormaFarmaceutica), "", oRsTmp4.Fields!FormaFarmaceutica)
                    oRsFarmacia.Update
                 End If
                 oRsTmp1.MoveNext
            Loop
            oRsFarmacia.Sort = "tipo,medicInsumo"
        End If
        oRsTmp1.Close
        'Farmacia - Insumos - Desde el SIS
        Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaAtencionINSxIdCuentaAtencion(ml_IdCuentaAtencion)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
                 lbSeGraboUnCptOFarmaciaEnTablasSIS = True
                 oRsDx.MoveFirst
                 oRsDx.Find "dxNro=" & oRsTmp1!DxNumero
                 Set oRsTmp4 = mo_AdminServiciosComunes.MedicamentosInsumosSeleccionarPorCodigo(oRsTmp1.Fields!codigo)
                 oRsFarmacia.AddNew
                 oRsFarmacia.Fields!id = oRsTmp1.Fields!id
                 oRsFarmacia.Fields!tipo = lcInsumo
                 oRsFarmacia.Fields!MedicInsumo = IIf(IsNull(oRsTmp4.Fields!nombre), "", oRsTmp4.Fields!nombre)
                 oRsFarmacia.Fields!recetado = oRsTmp1.Fields!CantidadPrescrita
                 oRsFarmacia.Fields!cantidad = oRsTmp1.Fields!CantidadEntregada
                 oRsFarmacia.Fields!dx = IIf(oRsDx.EOF, lcDxPrincipal, IIf(IsNull(oRsDx!DxEgreso), oRsDx!dxIngreso, oRsDx!DxEgreso))   'debb-23/02/2017
                 oRsFarmacia.Fields!Precio = oRsTmp1.Fields!PrecioUnitario
                 oRsFarmacia.Fields!codigo = oRsTmp1.Fields!codigo
                 oRsFarmacia.Fields!dxNro = IIf(oRsDx.EOF, lcDxPrincipalNro, oRsTmp1!DxNumero)
                 oRsFarmacia.Fields!formaF = IIf(IsNull(oRsTmp4.Fields!FormaFarmaceutica), "", oRsTmp4.Fields!FormaFarmaceutica)
                 oRsFarmacia.Update
                 oRsTmp1.MoveNext
            Loop
            oRsFarmacia.Sort = "tipo,medicInsumo"
        End If
        oRsTmp1.Close
        'Cpt - Desde el SIS
        Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaAtencionPROxIdCuentaAtencion(ml_IdCuentaAtencion)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
                    lbSeGraboUnCptOFarmaciaEnTablasSIS = True
                    oRsDx.MoveFirst
                    oRsDx.Find "dxNro=" & oRsTmp1!DxNumero
                    Set oRsTmp2 = mo_AdminServiciosComunes.ProcedimientosSeleccionarPorCodigo(oRsTmp1.Fields!codigo)
                    If oRsTmp2.RecordCount = 0 Then
                           lnIdPuntoCarga = 1
                           lcSql = "Consumo en el Servicio"
                    Else
                           lnIdPuntoCarga = oRsTmp2.Fields!idPuntoCarga
                           lcSql = oRsTmp2.Fields!nombre
                    End If
                    lcPuntoCarga = lcOtros
                    Select Case lnIdPuntoCarga
                    Case sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, sghPuntosCargaBasicos.sghPtoCargaBancoSangre1, sghPuntosCargaBasicos.sghPtoCargaBancoSangre2, sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica   'Laboratorio
                       lcPuntoCarga = lcLaboratorio
                    Case sghPuntosCargaBasicos.sghPtoCargaEcogGeneral, sghPuntosCargaBasicos.sghPtoCargaEcogObstetrica, sghPuntosCargaBasicos.sghPtoCargaRayosX, sghPuntosCargaBasicos.sghPtoCargaTomografia  'Imágenes
                       lcPuntoCarga = lcImagenes
                    End Select
                    oRsPatologia.AddNew
                    oRsPatologia.Fields!id = oRsTmp1.Fields!id
                    oRsPatologia.Fields!tipo = lcPuntoCarga
                    oRsPatologia.Fields!procedimiento = lcSql
                    oRsPatologia.Fields!indicado = oRsTmp1.Fields!CantidadPrescrita
                    oRsPatologia.Fields!ejecutado = oRsTmp1.Fields!CantidadEjecutada
                    oRsPatologia.Fields!dx = IIf(oRsDx.EOF, lcDxPrincipal, IIf(IsNull(oRsDx!DxEgreso), oRsDx!dxIngreso, oRsDx!DxEgreso))  'debb-23/02/2017
                    oRsPatologia.Fields!Precio = oRsTmp1.Fields!PrecioUnitario
                    oRsPatologia.Fields!idPuntoCarga = lnIdPuntoCarga
                    oRsPatologia.Fields!codigo = oRsTmp1.Fields!codigo
                    oRsPatologia.Fields!dxNro = IIf(oRsDx.EOF, lcDxPrincipalNro, oRsTmp1!DxNumero)
                    oRsPatologia.Update
                    oRsTmp1.MoveNext
            Loop
            oRsPatologia.Sort = "tipo,procedimiento"
        End If
        oRsTmp1.Close
        'No se registró aun consumos en tablas del SIS -Farmacia
        'Se jalan datos desde tablas GalenHos
        If oRsFarmacia.RecordCount = 0 Then
            Set oRsTmp1 = mo_ReglasFarmacia.farmMovimientoVentasDetalleSeleccionarXidCuenta(ml_IdCuentaAtencion)
            If oRsTmp1.RecordCount > 0 Then
               oRsTmp1.MoveFirst
               Do While Not oRsTmp1.EOF
                    If oRsTmp1.Fields!Precio > 0 Then
                        lnRecetado = 0
                        Set oRsTmp2 = mo_AdminServiciosComunes.RecetaDetalleSeleccionarXidCuentaIdItemDocumento(ml_IdCuentaAtencion, oRsTmp1.Fields!IdProducto, oRsTmp1.Fields!DocumentoNumero)
                        If oRsTmp2.RecordCount > 0 Then
                           lnRecetado = oRsTmp2.Fields!CantidadPedida
                        End If
                        oRsTmp2.Close
                        oRsFarmacia.AddNew
                        oRsFarmacia.Fields!id = oRsTmp1.Fields!IdProducto
                        oRsFarmacia.Fields!tipo = IIf(oRsTmp1.Fields!TipoProducto = 1, lcInsumo, lcMedicamento)
                        oRsFarmacia.Fields!MedicInsumo = oRsTmp1.Fields!nombre
                        oRsFarmacia.Fields!recetado = lnRecetado   'IIf(lnRecetado = 0, oRsTmp1.Fields!Cantidad, lnRecetado)
                        oRsFarmacia.Fields!cantidad = oRsTmp1.Fields!cantidad
                        oRsFarmacia.Fields!dx = lcDxPrincipal
                        oRsFarmacia.Fields!Precio = oRsTmp1.Fields!Precio
                        oRsFarmacia.Fields!codigo = oRsTmp1.Fields!codigo
                        oRsFarmacia.Fields!dxNro = lcDxPrincipalNro
                        oRsFarmacia.Fields!formaF = oRsTmp1.Fields!FormaFarmaceutica
                        oRsFarmacia.Fields!esPaquete = IIf(IsNull(oRsTmp1!esPaquete), False, oRsTmp1!esPaquete)
                        oRsFarmacia.Update
                    End If
                    oRsTmp1.MoveNext
               Loop
               mo_ReglasSISgalenhos.FuaPaquetesFarmaciaDesagregaEnMedicInsumos oRsFarmacia, mo_lnIdTablaLISTBARITEMS
               oRsFarmacia.Sort = "tipo,medicInsumo"
            End If
            oRsTmp1.Close
        End If
        'No se registró aun consumos en tablas del SIS -Cpt
        'Se jalan datos desde tablas GalenHos
        If oRsPatologia.RecordCount = 0 Then
            Set oRsTmp1 = mo_ReglasFacturacion.FacturacionServicioDespachoSeleccionarXidCuenta(ml_IdCuentaAtencion)
            If oRsTmp1.RecordCount > 0 Then
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                     If oRsTmp1.Fields!Precio > 0 Then
                         lnIdPuntoCarga = oRsTmp1.Fields!idPuntoCarga
                         lcPuntoCarga = lcOtros
                         Select Case lnIdPuntoCarga
                         Case sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, sghPuntosCargaBasicos.sghPtoCargaBancoSangre1, sghPuntosCargaBasicos.sghPtoCargaBancoSangre2, sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica   'Laboratorio
                              lcPuntoCarga = lcLaboratorio
                         Case sghPuntosCargaBasicos.sghPtoCargaEcogGeneral, sghPuntosCargaBasicos.sghPtoCargaEcogObstetrica, sghPuntosCargaBasicos.sghPtoCargaRayosX, sghPuntosCargaBasicos.sghPtoCargaTomografia  'Imágenes
                              lcPuntoCarga = lcImagenes
                         End Select
                         oRsPatologia.AddNew
                         oRsPatologia.Fields!id = oRsTmp1.Fields!IdProducto
                         oRsPatologia.Fields!tipo = lcPuntoCarga
                         oRsPatologia.Fields!procedimiento = oRsTmp1.Fields!nombre
                         oRsPatologia.Fields!indicado = 0   'oRsTmp1.Fields!Cantidad
                         oRsPatologia.Fields!ejecutado = oRsTmp1.Fields!cantidad
                         oRsPatologia.Fields!dx = lcDxPrincipal
                         oRsPatologia.Fields!Precio = oRsTmp1.Fields!Precio
                         oRsPatologia.Fields!idPuntoCarga = lnIdPuntoCarga
                         oRsPatologia.Fields!codigo = oRsTmp1.Fields!codigo
                         oRsPatologia.Fields!dxNro = lcDxPrincipalNro
                         oRsPatologia.Update
                     End If
                     oRsTmp1.MoveNext
                Loop
            End If
            oRsTmp1.Close
        End If
        'Cuenta que se emitió el FUA desde CITAS
        'ahora se quiere que se reimprima con el mismos numero en ATENCION MEDICO, mostrando Recetas,Dx,Cpt
        If mo_lnIdTablaLISTBARITEMS = sghRegistroAtencionCE And lbSeGraboUnCptOFarmaciaEnTablasSIS = False Then
            CargaDatosDeDx oConexion, True
            'Solo Receta
            Set oRsTmp1 = mo_AdminServiciosComunes.RecetaDetalleSoloFarmaciaSeleccionarXidCuentaAtencion(ml_IdCuentaAtencion, False)
            If oRsTmp1.RecordCount > 0 Then
               oRsTmp1.MoveFirst
               Do While Not oRsTmp1.EOF
                  If oRsTmp1.Fields!Precio > 0 Then
                     oRsFarmacia.AddNew
                     oRsFarmacia.Fields!id = oRsTmp1.Fields!IdProducto
                     oRsFarmacia.Fields!tipo = IIf(oRsTmp1.Fields!TipoProducto = 1, lcInsumo, lcMedicamento)
                     oRsFarmacia.Fields!MedicInsumo = oRsTmp1.Fields!nombre
                     oRsFarmacia.Fields!recetado = oRsTmp1.Fields!CantidadPedida
                     oRsFarmacia.Fields!cantidad = 0
                     oRsFarmacia.Fields!dx = lcDxPrincipal
                     oRsFarmacia.Fields!Precio = oRsTmp1.Fields!Precio
                     oRsFarmacia.Fields!codigo = oRsTmp1.Fields!codigo
                     oRsFarmacia.Fields!dxNro = lcDxPrincipalNro
                     oRsFarmacia.Fields!formaF = oRsTmp1.Fields!FormaFarmaceutica
                     oRsFarmacia.Fields!esPaquete = IIf(IsNull(oRsTmp1!esPaquete), False, oRsTmp1!esPaquete)
                     oRsFarmacia.Update
                  End If
                  oRsTmp1.MoveNext
               Loop
               mo_ReglasSISgalenhos.FuaPaquetesFarmaciaDesagregaEnMedicInsumos oRsFarmacia, mo_lnIdTablaLISTBARITEMS
               oRsFarmacia.Sort = "tipo,medicInsumo"
            End If
            oRsTmp1.Close
            'cpt-Solo Receta
            If oRsPatologia.RecordCount > 0 Then
               oRsPatologia.MoveFirst
               Do While Not oRsPatologia.EOF
                  If oRsPatologia.Fields!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAdmisionCE Then
                        oRsPatologia.Delete
                  Else
                        oRsPatologia!indicado = oRsPatologia!ejecutado
                        oRsPatologia.Fields!dx = lcDxPrincipal
                        oRsPatologia.Fields!dxNro = lcDxPrincipalNro
                  End If
                  oRsPatologia.Update
                  oRsPatologia.MoveNext
               Loop
            End If
            Set oRsTmp1 = mo_AdminServiciosComunes.RecetaDetalleSoloServiciosSeleccionarXidCuentaAtencion(ml_IdCuentaAtencion, False)
            If oRsTmp1.RecordCount > 0 Then
               oRsTmp1.MoveFirst
               Do While Not oRsTmp1.EOF
                  If oRsTmp1.Fields!Precio > 0 Then
                     lnIdPuntoCarga = oRsTmp1.Fields!idPuntoCarga
                     lcPuntoCarga = lcOtros
                     Select Case lnIdPuntoCarga
                     Case sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, sghPuntosCargaBasicos.sghPtoCargaBancoSangre1, sghPuntosCargaBasicos.sghPtoCargaBancoSangre2, sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica   'Laboratorio
                        lcPuntoCarga = lcLaboratorio
                     Case sghPuntosCargaBasicos.sghPtoCargaEcogGeneral, sghPuntosCargaBasicos.sghPtoCargaEcogObstetrica, sghPuntosCargaBasicos.sghPtoCargaRayosX, sghPuntosCargaBasicos.sghPtoCargaTomografia  'Imágenes
                        lcPuntoCarga = lcImagenes
                     End Select
                     oRsPatologia.AddNew
                     oRsPatologia.Fields!id = oRsTmp1.Fields!IdProducto
                     oRsPatologia.Fields!tipo = lcPuntoCarga
                     oRsPatologia.Fields!procedimiento = oRsTmp1.Fields!nombre
                     oRsPatologia.Fields!indicado = oRsTmp1.Fields!CantidadPedida
                     oRsPatologia.Fields!ejecutado = 0
                     oRsPatologia.Fields!dx = lcDxPrincipal
                     oRsPatologia.Fields!Precio = oRsTmp1.Fields!Precio
                     oRsPatologia.Fields!idPuntoCarga = lnIdPuntoCarga
                     oRsPatologia.Fields!codigo = oRsTmp1.Fields!codigo
                     oRsPatologia.Fields!dxNro = lcDxPrincipalNro
                     oRsPatologia.Update
                  End If
                  
                  
                  oRsTmp1.MoveNext
               Loop
               oRsPatologia.Sort = "tipo,procedimiento"
            End If
            oRsTmp1.Close
            
            
        End If
    End If
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oRsTmp4 = Nothing
End Sub

Sub ActualizaDxParaFarmaciaServicios()
    If lcDxPrincipal <> "" Then
       If oRsPatologia.RecordCount > 0 Then
          oRsPatologia.MoveFirst
          Do While Not oRsPatologia.EOF
             oRsPatologia.Fields!dx = lcDxPrincipal
             oRsPatologia.Fields!dxNro = lcDxPrincipalNro
             oRsPatologia.Update
             oRsPatologia.MoveNext
          Loop
       End If
       If oRsFarmacia.RecordCount > 0 Then
          oRsFarmacia.MoveFirst
          Do While Not oRsFarmacia.EOF
             oRsFarmacia.Fields!dx = lcDxPrincipal
             oRsFarmacia.Fields!dxNro = lcDxPrincipalNro
             oRsFarmacia.Update
             oRsFarmacia.MoveNext
          Loop
       End If
    End If
End Sub


Sub ImpresionFua()
       Dim rsReporte As New Recordset
        Dim rsReporte1 As New Recordset
        Dim lnImpSubTot As Double: Dim lntImpSubTot As Double
        Dim lnImpAnul As Double: Dim lntImpAnul As Double
        Dim lnImpExo As Double: Dim lntImpExo As Double
        Dim lnImpDevol As Double: Dim lntImpDevol As Double
        Dim lnImpPagCta As Double: Dim lntImpPagCta As Double
        Dim lnImpTot As Double: Dim lntImpTot As Double
        Dim lnIdPartida As Long: Dim lcPartida As String: Dim lcDpartida As String
        Dim iFila As Long, iColumna As Integer
        Dim lRecordCount As Long
        Dim lbNuevo As Boolean, lbMuestraCPTdefaults As Boolean
        Dim lbSeguir As Boolean: Dim lcDocum As String: Dim lcBuscar As String: Dim lnDctos As Double
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim lbDerecha As Boolean
        Dim lnTotalLineas As Long, lnTotalPaginas As Integer, lnMaximaLineaPorPagina As Integer, lnFilasAinsertar As Integer
        Dim lnFor As Integer
        Dim lbEsOpenOffice As Boolean, lcTipoMI As String
        Dim lnHwnd As Long
        Dim mo_ReporteUtil As New sighentidades.ReporteUtil

        lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)

        If lbEsOpenOffice = True Then
            Dim ServiceManager As Object
            Dim Desktop As Object
            Dim Document As Object
            Dim Feuille As Object
            Dim Plage As Object
            Dim args()
            Dim Chemin As String
            Dim Fichier As String
            Dim lcArchivoExcel As String
            Dim PrintArea(0)
            Dim Style As Object
            Dim Border As Object
            'encabezado
            Dim PageStyles As Object
            Dim Sheet As Object
            Dim StyleFamilies As Object
            Dim DefPage As Object
            Dim Htext As Object
            Dim Hcontent As Object
            Dim ret As Long
            Dim PrintArgs(2)
            Dim PathFileOpenOffice As String
        Else
            Dim oExcel As Excel.Application
            Dim oWorkBookPlantilla As Workbook
            Dim oWorkBook As Workbook
            Dim oWorkSheet As Worksheet
        End If
        On Error GoTo ManejadorError
        If Val(txtFua3.Text) = 0 And oDoSisFuaAtencion.FuaNumero = "" Then
            MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
        Else
        
            If Val(txtFua3.Text) = 0 Then
               txtFua3.Text = oDoSisFuaAtencion.FuaNumero
               ml_IdCuentaAtencion = oDoSisFuaAtencion.IdCuentaAtencion
            End If

            If lbEsOpenOffice = True Then
                If mi_FuaTipoAnexo2015 = lcFuaAnexo1 Then lcArchivoExcel = App.Path + "\Plantillas\SisFua2Anex1.ods"
                If mi_FuaTipoAnexo2015 = lcFuaAnexo2 Then lcArchivoExcel = App.Path + "\Plantillas\SisFua2Anex2.ods"
                Fichier = Format(Time, "hhmmss") & ".ods"
                PathFileOpenOffice = App.Path + "\Plantillas\" & Fichier
                FileCopy lcArchivoExcel, PathFileOpenOffice
                lcArchivoExcel = Fichier
                Chemin = "file:///" & App.Path & "\Plantillas\"
                Chemin = Replace(Chemin, "\", "/")
                Fichier = Chemin & "/" & lcArchivoExcel
                '
                Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
                Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
                Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
                Set Feuille = Document.getSheets().getByIndex(0)
                ' Pone la ventana en primer plano, pasándole el Hwnd
                ret = SetForegroundWindow(lnHwnd)
            Else
                'Crea nueva hoja
                Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                Set oWorkBook = oExcel.Workbooks.Add
                'Abre, copia y cierra la plantilla
                If mi_FuaTipoAnexo2015 = lcFuaAnexo1 Then Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\SisFua2Anex1.xls")
                If mi_FuaTipoAnexo2015 = lcFuaAnexo2 Then Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\SisFua2Anex2.xls")
                oWorkBookPlantilla.Worksheets("SisFua").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
            End If
            If lcBuscaParametro.SeleccionaFilaParametro(582) = "S" Then
               oWorkSheet.PageSetup.LeftHeaderPicture.FileName = ""
            End If


            'DE LA INSTITUCIÓN PRESTADORA DE SERVICIOS DE SALUD '''''''''''''''''''''''''''''''''''''''''''''''
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(57, 1).setFormula("F.Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL & " " & lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos)
                Call Feuille.getcellbyposition(57, 2).setFormula("Cta: " & ml_IdCuentaAtencion & " " & " " & sighentidades.TipoServicioDevuelveNombreCorto(ml_IdTipoServicio))

                If mi_FuaTipoAnexo2015 = lcFuaAnexo1 Then
                    Call Feuille.getcellbyposition(43, 6).setFormula(Trim(txtColegio.Text))
                    Call Feuille.getcellbyposition(60, 6).setFormula(Trim(txtColegioCodigo.Text))
                    Call Feuille.getcellbyposition(43, 7).setFormula(Trim(cmbColegioNivel.Text))
                    Call Feuille.getcellbyposition(50, 7).setFormula(Trim(cmbColegioGrado.Text))
                    Call Feuille.getcellbyposition(58, 7).setFormula(Trim(txtColegioSeccion.Text))
                    Call Feuille.getcellbyposition(62, 7).setFormula(Trim(cmbColegioTurno.Text))
                End If

                Call Feuille.getcellbyposition(17, 6).setFormula(Trim(txtFua1.Text))
                Call Feuille.getcellbyposition(23, 6).setFormula(Trim(txtFua2.Text))
                Call Feuille.getcellbyposition(27, 6).setFormula(Trim(txtFua3.Text))
                Call Feuille.getcellbyposition(0, 12).setFormula(Trim(txtCScodigo.Text))
                Call Feuille.getcellbyposition(19, 12).setFormula(Trim(txtCS.Text))

                Call Feuille.getcellbyposition(43, 25).setFormula(IIf(chkIntramural.Value <> ssCBUnchecked, "X", ""))
                Call Feuille.getcellbyposition(43, 26).setFormula(IIf(chkExtramural.Value <> ssCBUnchecked, "X", ""))

                Call Feuille.getcellbyposition(35, 15).setFormula(IIf(chkAtencionAmbulatoria.Value <> ssCBUnchecked, "X", ""))
                Call Feuille.getcellbyposition(35, 17).setFormula(IIf(chkAtencionReferencia.Value <> ssCBUnchecked, "X", ""))
                Call Feuille.getcellbyposition(35, 18).setFormula(IIf(chkAtencionEmergencia.Value <> ssCBUnchecked, "X", ""))

                Call Feuille.getcellbyposition(5, 15).setFormula(IIf(chkPAestablecimiento.Value <> ssCBUnchecked, lcEquix, ""))
                Call Feuille.getcellbyposition(5, 17).setFormula(IIf(chkPAaisped.Value <> ssCBUnchecked, lcEquix, ""))
                If mi_FuaTipoAnexo2015 = lcFuaAnexo1 Then
                    Call Feuille.getcellbyposition(5, 18).setFormula(IIf(chkPAOfeFlexible.Value <> ssCBUnchecked, lcEquix, ""))
                    Call Feuille.getcellbyposition(8, 17).setFormula(txtPACodOfFlexible.Text)
                End If
                Call Feuille.getcellbyposition(38, 17).setFormula(Trim(txtROcodigo.Text))
                Call Feuille.getcellbyposition(46, 17).setFormula(txtRO.Text)
                Call Feuille.getcellbyposition(61, 17).setFormula(txtRONumero.Text)
            Else
                oWorkSheet.Cells(2, 58).Value = "F.Emisión: " & lcBuscaParametro.RetornaFechaServidorSQL & " " & lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos
                oWorkSheet.Cells(3, 58).Value = "Cta: " & ml_IdCuentaAtencion & " " & " " & IIf(ml_EsPacienteExterno = True, " (Cta con Seguros)", sighentidades.TipoServicioDevuelveNombreCorto(ml_IdTipoServicio))

                If mi_FuaTipoAnexo2015 = lcFuaAnexo1 Then
                    oWorkSheet.Cells(7, 44).Value = Trim(txtColegio.Text)
                    oWorkSheet.Cells(7, 61).Value = Trim(txtColegioCodigo.Text)
                    oWorkSheet.Cells(8, 44).Value = Trim(cmbColegioNivel.Text)
                    oWorkSheet.Cells(8, 51).Value = Trim(cmbColegioGrado.Text)
                    oWorkSheet.Cells(8, 59).Value = Trim(txtColegioSeccion.Text)
                    oWorkSheet.Cells(8, 63).Value = Trim(cmbColegioTurno.Text)
                End If

                oWorkSheet.Cells(7, 18).Value = Trim(txtFua1.Text)
                oWorkSheet.Cells(7, 24).Value = Trim(txtFua2.Text)
                oWorkSheet.Cells(7, 28).Value = Trim(txtFua3.Text)
                oWorkSheet.Cells(13, 1).Value = Trim(txtCScodigo.Text)
                oWorkSheet.Cells(13, 20).Value = Trim(txtCS.Text)

                oWorkSheet.Cells(16, 27).Value = IIf(chkIntramural.Value <> ssCBUnchecked, "X", "")
                oWorkSheet.Cells(18, 27).Value = IIf(chkExtramural.Value <> ssCBUnchecked, "X", "")

                oWorkSheet.Cells(16, 36).Value = IIf(chkAtencionAmbulatoria.Value <> ssCBUnchecked, "X", "")
                oWorkSheet.Cells(18, 36).Value = IIf(chkAtencionReferencia.Value <> ssCBUnchecked, "X", "")
                oWorkSheet.Cells(19, 36).Value = IIf(chkAtencionEmergencia.Value <> ssCBUnchecked, "X", "")

                oWorkSheet.Cells(16, 6).Value = IIf(chkPAestablecimiento.Value <> ssCBUnchecked, lcEquix, "")
                oWorkSheet.Cells(18, 6).Value = IIf(chkPAaisped.Value <> ssCBUnchecked, lcEquix, "")
                If mi_FuaTipoAnexo2015 = lcFuaAnexo1 Then
                    oWorkSheet.Cells(19, 6).Value = IIf(chkPAOfeFlexible.Value <> ssCBUnchecked, lcEquix, "")
                    oWorkSheet.Cells(18, 9).Value = txtPACodOfFlexible.Text
                End If
                oWorkSheet.Cells(18, 39).Value = Trim(txtROcodigo.Text)
                oWorkSheet.Cells(18, 47).Value = txtRO.Text
                oWorkSheet.Cells(18, 62).Value = txtRONumero.Text
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'DEL ASEGURADO / USUARIO ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(0, 24).setFormula(Trim(Left(Me.cmbTipoDocumento.Text, 8)))
                Call Feuille.getcellbyposition(2, 24).setFormula(Trim(txtNdocumento.Text))

                Call Feuille.getcellbyposition(16, 24).setFormula(Trim(txtNroAfiliacion1.Text))
                Call Feuille.getcellbyposition(22, 24).setFormula(Trim(txtNroAfiliacion2.Text))
                Call Feuille.getcellbyposition(25, 24).setFormula(Trim(txtNroAfiliacion3.Text))

                Call Feuille.getcellbyposition(42, 22).setFormula(txtInstitucion.Text)
                Call Feuille.getcellbyposition(42, 24).setFormula(txtCodSeguro.Text)

                Call Feuille.getcellbyposition(0, 27).setFormula(ml_ApellidoPaterno)
                Call Feuille.getcellbyposition(36, 27).setFormula(ml_ApellidoMaterno)
                Call Feuille.getcellbyposition(0, 30).setFormula(ml_PrimerNombre)
                Call Feuille.getcellbyposition(36, 30).setFormula(ml_SegundoNombre)

                Call Feuille.getcellbyposition(4, 34).setFormula(IIf(Trim(txtSexo.Text) = "Masculino", "X", ""))
                Call Feuille.getcellbyposition(4, 35).setFormula(IIf(Trim(txtSexo.Text) = "Femenino", "X", ""))

                Call Feuille.getcellbyposition(4, 41).setFormula(IIf(chkGestante.Value <> ssCBUnchecked, lcEquix, ""))
                Call Feuille.getcellbyposition(4, 46).setFormula(IIf(chkPuerpera.Value <> ssCBUnchecked, lcEquix, ""))

                If Trim(txtFparto.Text) <> "" Then
                    If sighentidades.EsFecha(txtFparto.Text, "DD/MM/AAAA") = True Then
                        Call Feuille.getcellbyposition(17, 34).setFormula(Day(txtFparto.Text))
                        Call Feuille.getcellbyposition(24, 34).setFormula(Month(txtFparto.Text))
                        Call Feuille.getcellbyposition(30, 34).setFormula(Year(txtFparto.Text))
                    End If
                End If

                If Trim(txtFnacimiento.Text) <> "" Then
                    If sighentidades.EsFecha(txtFnacimiento.Text, "DD/MM/AAAA") = True Then
                        Call Feuille.getcellbyposition(17, 37).setFormula(Day(txtFparto.Text))
                        Call Feuille.getcellbyposition(24, 37).setFormula(Month(txtFparto.Text))
                        Call Feuille.getcellbyposition(30, 37).setFormula(Year(txtFparto.Text))
                    End If
                End If

                If Trim(txtFFallecimiento.Text) <> "" Then
                    If sighentidades.EsFecha(txtFFallecimiento.Text, "DD/MM/AAAA") = True Then
                        Call Feuille.getcellbyposition(17, 44).setFormula(Day(txtFparto.Text))
                        Call Feuille.getcellbyposition(24, 44).setFormula(Month(txtFparto.Text))
                        Call Feuille.getcellbyposition(30, 44).setFormula(Year(txtFparto.Text))
                    End If
                End If

                Call Feuille.getcellbyposition(43, 34).setFormula(txtNhistoriaClinica.Text)
                Call Feuille.getcellbyposition(57, 34).setFormula(cmbEtnia.Text)

                'RN
                If oRsNacimientos.RecordCount > 0 Then
                   oRsNacimientos.MoveFirst
                   Call Feuille.getcellbyposition(59, 37).setFormula(oRsNacimientos!documento)
                   oRsNacimientos.MoveNext
                   If Not oRsNacimientos.EOF Then
                        Call Feuille.getcellbyposition(59, 41).setFormula(oRsNacimientos!documento)
                        oRsNacimientos.MoveNext
                        If Not oRsNacimientos.EOF Then
                           Call Feuille.getcellbyposition(59, 46).setFormula(oRsNacimientos!documento)
                        End If
                   End If
                End If

            Else
                oWorkSheet.Cells(25, 1).Value = Trim(Left(Me.cmbTipoDocumento.Text, 8))
                oWorkSheet.Cells(25, 3).Value = txtNdocumento.Text

                oWorkSheet.Cells(25, 17).Value = Trim(txtNroAfiliacion1.Text)
                oWorkSheet.Cells(25, 23).Value = Trim(txtNroAfiliacion2.Text)
                oWorkSheet.Cells(25, 26).Value = Trim(txtNroAfiliacion3.Text)

                oWorkSheet.Cells(23, 43).Value = txtInstitucion.Text
                oWorkSheet.Cells(25, 43).Value = txtCodSeguro.Text

                oWorkSheet.Cells(28, 1).Value = ml_ApellidoPaterno
                oWorkSheet.Cells(28, 37).Value = ml_ApellidoMaterno
                oWorkSheet.Cells(31, 1).Value = ml_PrimerNombre
                oWorkSheet.Cells(31, 37).Value = ml_SegundoNombre

                oWorkSheet.Cells(35, 5).Value = IIf(Trim(txtSexo.Text) = "Masculino", "X", "")
                oWorkSheet.Cells(36, 5).Value = IIf(Trim(txtSexo.Text) = "Femenino", "X", "")

                oWorkSheet.Cells(42, 5).Value = IIf(chkGestante.Value <> ssCBUnchecked, lcEquix, "")
                oWorkSheet.Cells(47, 5).Value = IIf(chkPuerpera.Value <> ssCBUnchecked, lcEquix, "")

                If Trim(txtFparto.Text) <> "" Then
                    If sighentidades.EsFecha(txtFparto.Text, "DD/MM/AAAA") = True Then
                        oWorkSheet.Cells(35, 18).Value = Day(txtFparto.Text)
                        oWorkSheet.Cells(35, 25).Value = Month(txtFparto.Text)
                        oWorkSheet.Cells(35, 31).Value = Year(txtFparto.Text)
                    End If
                End If

                If Trim(txtFnacimiento.Text) <> "" Then
                    If sighentidades.EsFecha(txtFnacimiento.Text, "DD/MM/AAAA") = True Then
                        oWorkSheet.Cells(38, 18).Value = Day(txtFnacimiento.Text)
                        oWorkSheet.Cells(38, 25).Value = Month(txtFnacimiento.Text)
                        oWorkSheet.Cells(38, 31).Value = Year(txtFnacimiento.Text)
                    End If
                End If

                If Trim(txtFFallecimiento.Text) <> "" Then
                    If sighentidades.EsFecha(txtFFallecimiento.Text, "DD/MM/AAAA") = True Then
                        oWorkSheet.Cells(45, 18).Value = Day(txtFFallecimiento.Text)
                        oWorkSheet.Cells(45, 25).Value = Month(txtFFallecimiento.Text)
                        oWorkSheet.Cells(45, 31).Value = Year(txtFFallecimiento.Text)
                    End If
                End If

                oWorkSheet.Cells(35, 44).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(txtNhistoriaClinica.Text, True)
                oWorkSheet.Cells(35, 58).Value = cmbEtnia.Text

                'RN
                If oRsNacimientos.RecordCount > 0 Then
                   oRsNacimientos.MoveFirst
                   oWorkSheet.Cells(38, 60).Value = oRsNacimientos!documento
                   oRsNacimientos.MoveNext
                   If Not oRsNacimientos.EOF Then
                        oWorkSheet.Cells(42, 60).Value = oRsNacimientos!documento
                        oRsNacimientos.MoveNext
                        If Not oRsNacimientos.EOF Then
                           oWorkSheet.Cells(47, 60).Value = oRsNacimientos!documento
                        End If
                   End If
                End If

            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'DE LA ATENCIÓN '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If lbEsOpenOffice = True Then
                If Trim(txtFantencion.Text) <> "" Then
                    If sighentidades.EsFecha(txtFantencion.Text, "DD/MM/AAAA") = True Then
                        Call Feuille.getcellbyposition(0, 53).setFormula(Day(txtFantencion.Text))
                        Call Feuille.getcellbyposition(2, 53).setFormula(Month(txtFantencion.Text))
                        Call Feuille.getcellbyposition(6, 53).setFormula(Year(txtFantencion.Text))
                    End If
                End If
                Call Feuille.getcellbyposition(15, 52).setFormula(CStr("'" & Format(txtHatencion.Text, "hh:mm")))
                Call Feuille.getcellbyposition(22, 52).setFormula(cmbUPSfua.Text)

                Call Feuille.getcellbyposition(28, 52).setFormula("'" & Right("000" & ucSISfuaCodPrestacion1.CodigoPrestacion, 3))
                Call Feuille.getcellbyposition(28, 55).setFormula(ucSISfuaCodPrestacion1.Prestacion)

                Call Feuille.getcellbyposition(36, 52).setFormula(txtCodPrestAdicional.Text)

                Call Feuille.getcellbyposition(11, 61).setFormula(txtCodAutorizacion.Text)
                Call Feuille.getcellbyposition(26, 61).setFormula(txtFuaVincular.Text)

                If Trim(txtHfingreso.Text) <> "" Then
                    If sighentidades.EsFecha(txtHfingreso.Text, "DD/MM/AAAA") = True Then
                        Call Feuille.getcellbyposition(57, 51).setFormula(Day(txtHfingreso.Text))
                        Call Feuille.getcellbyposition(59, 51).setFormula(Month(txtHfingreso.Text))
                        Call Feuille.getcellbyposition(61, 51).setFormula(Year(txtHfingreso.Text))
                    End If
                End If

                If Trim(txtHfalta.Text) <> "" Then
                    If sighentidades.EsFecha(txtHfalta.Text, "DD/MM/AAAA") = True Then
                        Call Feuille.getcellbyposition(57, 55).setFormula(Day(txtHfalta.Text))
                        Call Feuille.getcellbyposition(59, 55).setFormula(Month(txtHfalta.Text))
                        Call Feuille.getcellbyposition(61, 55).setFormula(Year(txtHfalta.Text))
                    End If
                End If

                If Trim(txtHFCortAdmin.Text) <> "" Then
                    If sighentidades.EsFecha(txtHFCortAdmin.Text, "DD/MM/AAAA") = True Then
                        Call Feuille.getcellbyposition(57, 60).setFormula(Day(txtHfalta.Text))
                        Call Feuille.getcellbyposition(59, 60).setFormula(Month(txtHfalta.Text))
                        Call Feuille.getcellbyposition(61, 60).setFormula(Year(txtHfalta.Text))
                    End If
                End If

            Else
                If Trim(txtFantencion.Text) <> "" Then
                    If sighentidades.EsFecha(txtFantencion.Text, "DD/MM/AAAA") = True Then
                        oWorkSheet.Cells(54, 1).Value = Day(txtFantencion.Text)
                        oWorkSheet.Cells(54, 3).Value = Month(txtFantencion.Text)
                        oWorkSheet.Cells(54, 7).Value = Year(txtFantencion.Text)
                    End If
                End If
                oWorkSheet.Cells(53, 16).Value = CStr("'" & Format(txtHatencion.Text, "hh:mm"))
                oWorkSheet.Cells(53, 23).Value = cmbUPSfua.Text

                oWorkSheet.Cells(53, 29).Value = "'" & Right("000" & ucSISfuaCodPrestacion1.CodigoPrestacion, 3)
                oWorkSheet.Cells(56, 29).Value = ucSISfuaCodPrestacion1.Prestacion

                oWorkSheet.Cells(53, 37).Value = txtCodPrestAdicional.Text

                oWorkSheet.Cells(62, 12).Value = txtCodAutorizacion.Text
                oWorkSheet.Cells(62, 27).Value = txtFuaVincular.Text

                If Trim(txtHfingreso.Text) <> "" Then
                    If sighentidades.EsFecha(txtHfingreso.Text, "DD/MM/AAAA") = True Then
                        oWorkSheet.Cells(52, 58).Value = Day(txtHfingreso.Text)
                        oWorkSheet.Cells(52, 60).Value = Month(txtHfingreso.Text)
                        oWorkSheet.Cells(52, 62).Value = Year(txtHfingreso.Text)
                    End If
                End If

                If Trim(txtHfalta.Text) <> "" Then
                    If sighentidades.EsFecha(txtHfalta.Text, "DD/MM/AAAA") = True Then
                        oWorkSheet.Cells(56, 58).Value = Day(txtHfalta.Text)
                        oWorkSheet.Cells(56, 60).Value = Month(txtHfalta.Text)
                        oWorkSheet.Cells(56, 62).Value = Year(txtHfalta.Text)
                    End If
                End If

                If Trim(txtHFCortAdmin.Text) <> "" Then
                    If sighentidades.EsFecha(txtHFCortAdmin.Text, "DD/MM/AAAA") = True Then
                        oWorkSheet.Cells(61, 58).Value = Day(txtHFCortAdmin.Text)
                        oWorkSheet.Cells(61, 60).Value = Month(txtHFCortAdmin.Text)
                        oWorkSheet.Cells(61, 62).Value = Year(txtHFCortAdmin.Text)
                    End If
                End If

            End If
             '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'CONCEPTO PRESTACIONAL''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            lcSql = lcEquix
            Select Case mo_cmbConceptoP.BoundText
            Case 1    'Atención Directa
                 iFila = 65
                 iColumna = 4
            Case 2    'Enfermedad Alto Costo (No LPIS)
                 iFila = 65
                 iColumna = 4
            Case 3    'Caso Especial / Cob. Extraordinaria
                 iFila = 66
                 iColumna = 16
            Case 4    'Sepelio
                 iFila = 66
                 iColumna = 54
            Case 5    'Traslado
                 iFila = 65
                 iColumna = 44
            Case 6   'Carta de Garantia
                 iFila = 66
                 iColumna = 31
            Case 7   'Sepelio Natimuerto
                 iFila = 66
                 iColumna = 54
            Case 8   'Sepelio Obito
                 iFila = 66
                 iColumna = 60
            Case 9   'Sepelio Otro
                 iFila = 66
                 iColumna = 64
            Case Else
                 iFila = 65
                 iColumna = 4
                 lcSql = ""
            End Select
            If mo_cmbConceptoP.BoundText <> "" Then
                If Not (mo_cmbConceptoP.BoundText = 3 Or mo_cmbConceptoP.BoundText = 6) Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(iColumna - 1, iFila - 1).setFormula(lcSql)
                    Else
                        oWorkSheet.Cells(iFila, iColumna).Value = lcSql
                    End If
                Else
                    If txtNautorizacion.Text <> "" Then
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iColumna - 1, iFila - 1).setFormula(txtNautorizacion.Text)
                        Else
                            oWorkSheet.Cells(iFila, iColumna).Value = txtNautorizacion.Text
                        End If
                    End If
                    If Val(txtMonto.Text) > 0 Then
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(iColumna - 1, iFila).setFormula(txtMonto.Text)
                        Else
                            oWorkSheet.Cells(iFila + 1, iColumna).Value = txtMonto.Text
                        End If
                    End If
                End If
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'DEL DESTINO DEL ASEGURADO/USUARIO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            lcSql = lcEquix
            Select Case mo_cmbIdDestinoAtencion.BoundText
            Case "1"     'alta
                 iFila = 70
                 iColumna = 2
            Case "2"     'citado
                 iFila = 70
                 iColumna = 7
            Case "3"     'Ref. Emergencia
                 iFila = 71
                 iColumna = 26
            Case "4"     'Ref. Consulta Externa
                 iFila = 71
                 iColumna = 36
            Case "5"     'Ref. Apoyo al Dx.
                 iFila = 71
                 iColumna = 48
            Case "6"     'Contrarreferido
                 iFila = 70
                 iColumna = 56
            Case "7"     'Fallecido
                 iFila = 70
                 iColumna = 61
            Case "8"     'Hospitalizado
                 iFila = 70
                 iColumna = 18
            Case "9"     'Corte Administrativo
                 iFila = 70
                 iColumna = 64
            Case Else
                 iFila = 70
                 iColumna = 2
                 lcSql = ""
            End Select

            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(iColumna - 1, iFila - 1).setFormula(lcSql)
            Else
                oWorkSheet.Cells(iFila, iColumna).Value = lcSql
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'SE REFIERE / CONTRARREFIERE A:'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(2, 37).setFormula(Trim(txtRDcodigo.Text))
                Call Feuille.getcellbyposition(19, 37).setFormula(txtRD.Text)
                Call Feuille.getcellbyposition(68, 37).setFormula(txtRDnumero.Text)
            Else
                oWorkSheet.Cells(75, 1).Value = Trim(txtRDcodigo.Text)
                oWorkSheet.Cells(75, 14).Value = txtRD.Text
                oWorkSheet.Cells(75, 56).Value = txtRDnumero.Text
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'ACTIVIDADES PREVENTIVAS Y OTROS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(9, 77).setFormula(txtSPpeso.Text)
                Call Feuille.getcellbyposition(24, 77).setFormula(txtSPtalla.Text)
                Call Feuille.getcellbyposition(39, 77).setFormula(txtSPpa.Text)
                If mi_FuaTipoAnexo2015 = lcFuaAnexo1 Then
                    Call Feuille.getcellbyposition(42, 86).setFormula(IIf(chkSPTamizajeSalMPAT.Value <> ssCBUnchecked, "PAT.", IIf(chkSPTamizajeSalMNOR.Value <> ssCBUnchecked, "NOR.", "")))
                Else
                    Call Feuille.getcellbyposition(55, 77).setFormula(IIf(chkSPTamizajeSalMPAT.Value <> ssCBUnchecked, "PAT.", IIf(chkSPTamizajeSalMNOR.Value <> ssCBUnchecked, "NOR.", "")))
                End If
            Else
                oWorkSheet.Cells(78, 10).Value = txtSPpeso.Text
                oWorkSheet.Cells(78, 25).Value = txtSPtalla.Text
                oWorkSheet.Cells(78, 40).Value = txtSPpa.Text
                If mi_FuaTipoAnexo2015 = lcFuaAnexo1 Then
                    oWorkSheet.Cells(87, 43).Value = IIf(chkSPTamizajeSalMPAT.Value <> ssCBUnchecked, "PAT.", IIf(chkSPTamizajeSalMNOR.Value <> ssCBUnchecked, "NOR.", ""))
                Else
                    oWorkSheet.Cells(78, 56).Value = IIf(chkSPTamizajeSalMPAT.Value <> ssCBUnchecked, "PAT.", IIf(chkSPTamizajeSalMNOR.Value <> ssCBUnchecked, "NOR.", ""))
                End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'DE LA GESTANTE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(2, 79).setFormula(txtSPcpn.Text)
                Call Feuille.getcellbyposition(2, 82).setFormula(txtSPedadG.Text)
                Call Feuille.getcellbyposition(2, 84).setFormula(txtSPalturaU.Text)
                Call Feuille.getcellbyposition(2, 86).setFormula(IIf(chkSPPartoVertSI.Value <> ssCBUnchecked, "Si", IIf(chkSPPartoVertNO.Value <> ssCBUnchecked, "No", "")))
                Call Feuille.getcellbyposition(2, 88).setFormula(txtSPpuerperio.Text)
            Else
                oWorkSheet.Cells(80, 3).Value = txtSPcpn.Text
                oWorkSheet.Cells(83, 3).Value = txtSPedadG.Text
                oWorkSheet.Cells(85, 3).Value = txtSPalturaU.Text
                oWorkSheet.Cells(87, 3).Value = IIf(chkSPPartoVertSI.Value <> ssCBUnchecked, "Si", IIf(chkSPPartoVertNO.Value <> ssCBUnchecked, "No", ""))
                oWorkSheet.Cells(89, 3).Value = txtSPpuerperio.Text
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'DEL RECIEN NACIDO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(14, 79).setFormula(txtSPedadGrn.Text)
                Call Feuille.getcellbyposition(12, 82).setFormula(txtSPapgar1.Text)
                Call Feuille.getcellbyposition(14, 82).setFormula(txtSPapgar5.Text)
                Call Feuille.getcellbyposition(14, 86).setFormula(IIf(chkSPCorTarCordonSI.Value <> ssCBUnchecked, "Si", IIf(chkSPCorteTarCordonNO.Value <> ssCBUnchecked, "No", "")))
            Else
                oWorkSheet.Cells(80, 15).Value = txtSPedadGrn.Text
                oWorkSheet.Cells(83, 13).Value = txtSPapgar1.Text
                oWorkSheet.Cells(83, 15).Value = txtSPapgar5.Text
                oWorkSheet.Cells(87, 15).Value = IIf(chkSPCorTarCordonSI.Value <> ssCBUnchecked, "Si", IIf(chkSPCorteTarCordonNO.Value <> ssCBUnchecked, "No", ""))
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'GESTANTE / RN /  NIÑO / ADOLESCENTE / JOVEN Y ADULTO / ADULTO MAYOR''''''''''''''''''''''''''''''''
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(26, 79).setFormula(txtSPcred.Text)
                Call Feuille.getcellbyposition(33, 79).setFormula(txtSPPAB.Text)
                Call Feuille.getcellbyposition(26, 82).setFormula(IIf(chkSPRNPrematuroSI.Value <> ssCBUnchecked, "Si", IIf(chkSPRNPrematuroNO.Value <> ssCBUnchecked, "No", "")))
                Call Feuille.getcellbyposition(34, 82).setFormula(IIf(chkSPeedpSI.Value <> ssCBUnchecked, "Si", IIf(chkSPeedpNO.Value <> ssCBUnchecked, "No", "")))
                Call Feuille.getcellbyposition(26, 84).setFormula(IIf(chkSBajoPesoSI.Value <> ssCBUnchecked, "Si", IIf(chkSBajoPesoNO.Value <> ssCBUnchecked, "No", "")))
                Call Feuille.getcellbyposition(34, 84).setFormula(IIf(chkSPconsejeriaNsi.Value <> ssCBUnchecked, "Si", IIf(chkSPconsejeriaNno.Value <> ssCBUnchecked, "No", "")))
                Call Feuille.getcellbyposition(26, 86).setFormula(IIf(chkSPSecuelaNaceSI.Value <> ssCBUnchecked, "Si", IIf(chkSPSecuelaNaceNO.Value <> ssCBUnchecked, "No", "")))
                Call Feuille.getcellbyposition(34, 86).setFormula(IIf(chkSPConIntegralSI.Value <> ssCBUnchecked, "Si", IIf(chkSPConIntegralNO.Value <> ssCBUnchecked, "No", "")))
                Call Feuille.getcellbyposition(26, 89).setFormula(txtSPNFamGestante.Text)
                Call Feuille.getcellbyposition(34, 89).setFormula(txtSPIMC.Text)
            Else
                oWorkSheet.Cells(80, 27).Value = txtSPcred.Text
                oWorkSheet.Cells(80, 34).Value = txtSPPAB.Text
                oWorkSheet.Cells(83, 27).Value = IIf(chkSPRNPrematuroSI.Value <> ssCBUnchecked, "Si", IIf(chkSPRNPrematuroNO.Value <> ssCBUnchecked, "No", ""))
                oWorkSheet.Cells(83, 35).Value = IIf(chkSPeedpSI.Value <> ssCBUnchecked, "Si", IIf(chkSPeedpNO.Value <> ssCBUnchecked, "No", ""))
                oWorkSheet.Cells(85, 27).Value = IIf(chkSBajoPesoSI.Value <> ssCBUnchecked, "Si", IIf(chkSBajoPesoNO.Value <> ssCBUnchecked, "No", ""))
                oWorkSheet.Cells(85, 35).Value = IIf(chkSPconsejeriaNsi.Value <> ssCBUnchecked, "Si", IIf(chkSPconsejeriaNno.Value <> ssCBUnchecked, "No", ""))
                oWorkSheet.Cells(87, 27).Value = IIf(chkSPSecuelaNaceSI.Value <> ssCBUnchecked, "Si", IIf(chkSPSecuelaNaceNO.Value <> ssCBUnchecked, "No", ""))
                oWorkSheet.Cells(87, 35).Value = IIf(chkSPConIntegralSI.Value <> ssCBUnchecked, "Si", IIf(chkSPConIntegralNO.Value <> ssCBUnchecked, "No", ""))
                oWorkSheet.Cells(90, 27).Value = txtSPNFamGestante.Text
                oWorkSheet.Cells(90, 35).Value = txtSPIMC.Text
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'JOVEN Y ADULTO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(42, 79).setFormula(IIf(chkSPEvalIntegralSI.Value <> ssCBUnchecked, "Si", IIf(chkSPEvalIntegralNO.Value <> ssCBUnchecked, "No", "")))
                Call Feuille.getcellbyposition(42, 83).setFormula(txtSPVacam.Text)
            Else
                oWorkSheet.Cells(80, 43).Value = IIf(chkSPEvalIntegralSI.Value <> ssCBUnchecked, "Si", IIf(chkSPEvalIntegralNO.Value <> ssCBUnchecked, "No", ""))
                oWorkSheet.Cells(84, 43).Value = txtSPVacam.Text
            End If


            If chkSPTamizajeSalMPAT.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(79, 45).setFormula("Si")
                    Call Feuille.getcellbyposition(79, 46).setFormula("")
                Else
                    oWorkSheet.Cells(46, 80).Value = "Si"
                    oWorkSheet.Cells(47, 80).Value = ""
                End If
            ElseIf chkSPTamizajeSalMNOR.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(79, 45).setFormula("")
                    Call Feuille.getcellbyposition(79, 46).setFormula("No")
                Else
                    oWorkSheet.Cells(46, 80).Value = ""
                    oWorkSheet.Cells(47, 80).Value = "No"
                End If
            End If

            'Vacunas
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(51, 77).setFormula(txtVacBcg.Text)
                Call Feuille.getcellbyposition(58, 77).setFormula(txtVacInfluenz.Text)
                Call Feuille.getcellbyposition(63, 77).setFormula(txtVacAntiamarilica.Text)
                Call Feuille.getcellbyposition(51, 78).setFormula(txtVacDpt.Text)
                Call Feuille.getcellbyposition(58, 78).setFormula(txtVacParotid.Text)
                Call Feuille.getcellbyposition(63, 78).setFormula(txtVacAntineumoc.Text)
                Call Feuille.getcellbyposition(51, 80).setFormula(txtVacApo.Text)
                Call Feuille.getcellbyposition(58, 80).setFormula(txtVacRubeola.Text)
                Call Feuille.getcellbyposition(63, 80).setFormula(txtVacAntitetanica.Text)
                Call Feuille.getcellbyposition(51, 82).setFormula(txtVacAsa.Text)
                Call Feuille.getcellbyposition(58, 82).setFormula(txtVacRotavirus.Text)
                Call Feuille.getcellbyposition(63, 82).setFormula(IIf(chkVacCompEdSI.Value <> ssCBUnchecked, "Si", IIf(chkVacCompEdNo.Value <> ssCBUnchecked, "No", "")))
                Call Feuille.getcellbyposition(51, 84).setFormula(txtVacSpr.Text)
                Call Feuille.getcellbyposition(58, 84).setFormula(txtVacDt.Text)
                Call Feuille.getcellbyposition(63, 84).setFormula(txtVacVPH.Text)
                Call Feuille.getcellbyposition(51, 86).setFormula(txtVacSR.Text)
                Call Feuille.getcellbyposition(58, 86).setFormula(txtVacIPV.Text)
                Call Feuille.getcellbyposition(63, 86).setFormula(txtVacOtraVacuna.Text)
                Call Feuille.getcellbyposition(51, 87).setFormula(txtVacHVB.Text)
                Call Feuille.getcellbyposition(58, 87).setFormula(txtVacPentaval.Text)
                Call Feuille.getcellbyposition(52, 89).setFormula(txtVacRiesgoHVB.Text)
            Else
                oWorkSheet.Cells(78, 52).Value = txtVacBcg.Text
                oWorkSheet.Cells(78, 59).Value = txtVacInfluenz.Text
                oWorkSheet.Cells(78, 64).Value = txtVacAntiamarilica.Text
                oWorkSheet.Cells(79, 52).Value = txtVacDpt.Text
                oWorkSheet.Cells(79, 59).Value = txtVacParotid.Text
                oWorkSheet.Cells(79, 64).Value = txtVacAntineumoc.Text
                oWorkSheet.Cells(81, 52).Value = txtVacApo.Text
                oWorkSheet.Cells(81, 59).Value = txtVacRubeola.Text
                oWorkSheet.Cells(81, 64).Value = txtVacAntitetanica.Text
                oWorkSheet.Cells(83, 52).Value = txtVacAsa.Text
                oWorkSheet.Cells(83, 59).Value = txtVacRotavirus.Text
                oWorkSheet.Cells(83, 64).Value = IIf(chkVacCompEdSI.Value <> ssCBUnchecked, "Si", IIf(chkVacCompEdNo.Value <> ssCBUnchecked, "No", ""))
                oWorkSheet.Cells(85, 52).Value = txtVacSpr.Text
                oWorkSheet.Cells(85, 59).Value = txtVacDt.Text
                oWorkSheet.Cells(85, 64).Value = txtVacVPH.Text
                oWorkSheet.Cells(87, 52).Value = txtVacSR.Text
                oWorkSheet.Cells(87, 59).Value = txtVacIPV.Text
                oWorkSheet.Cells(87, 64).Value = txtVacOtraVacuna.Text
                oWorkSheet.Cells(88, 52).Value = txtVacHVB.Text
                oWorkSheet.Cells(88, 59).Value = txtVacPentaval.Text
                oWorkSheet.Cells(90, 53).Value = txtVacRiesgoHVB.Text
            End If

            'Dx
            iFila = 96
            oRsDx.MoveFirst
            Do While Not oRsDx.EOF
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula(IIf(IsNull(oRsDx.Fields!Descripcion), "", oRsDx.Fields!Descripcion))
                    Call Feuille.getcellbyposition(40, iFila - 1).setFormula(IIf(oRsDx.Fields!DxIngresoPresuntivo = True, lcEquix, ""))
                    Call Feuille.getcellbyposition(42, iFila - 1).setFormula(IIf(oRsDx.Fields!DxIngresoDefinitivo = True, lcEquix, ""))
                    Call Feuille.getcellbyposition(45, iFila - 1).setFormula(IIf(oRsDx.Fields!DxIngresoRepetido = True, lcEquix, ""))
                    Call Feuille.getcellbyposition(49, iFila - 1).setFormula(IIf(IsNull(oRsDx.Fields!dxIngreso), "", oRsDx.Fields!dxIngreso))
                    Call Feuille.getcellbyposition(61, iFila - 1).setFormula(IIf(IsNull(oRsDx.Fields!DxEgreso), "", oRsDx.Fields!DxEgreso))
                    Call Feuille.getcellbyposition(56, iFila - 1).setFormula(IIf(oRsDx.Fields!DxEgresoDefinitivo = True, lcEquix, ""))
                    Call Feuille.getcellbyposition(59, iFila - 1).setFormula(IIf(oRsDx.Fields!DxEgresoRepetido = True, lcEquix, ""))
                Else
                    oWorkSheet.Cells(iFila, 2).Value = IIf(IsNull(oRsDx.Fields!Descripcion), "", oRsDx.Fields!Descripcion)
                    oWorkSheet.Cells(iFila, 41).Value = IIf(oRsDx.Fields!DxIngresoPresuntivo = True, lcEquix, "")
                    oWorkSheet.Cells(iFila, 43).Value = IIf(oRsDx.Fields!DxIngresoDefinitivo = True, lcEquix, "")
                    oWorkSheet.Cells(iFila, 46).Value = IIf(oRsDx.Fields!DxIngresoRepetido = True, lcEquix, "")
                    oWorkSheet.Cells(iFila, 50).Value = IIf(IsNull(oRsDx.Fields!dxIngreso), "", oRsDx.Fields!dxIngreso)
                    oWorkSheet.Cells(iFila, 62).Value = IIf(IsNull(oRsDx.Fields!DxEgreso), "", oRsDx.Fields!DxEgreso)
                    oWorkSheet.Cells(iFila, 57).Value = IIf(oRsDx.Fields!DxEgresoDefinitivo = True, lcEquix, "")
                    oWorkSheet.Cells(iFila, 60).Value = IIf(oRsDx.Fields!DxEgresoRepetido = True, lcEquix, "")
                End If
                oRsDx.MoveNext
                iFila = iFila + 1
            Loop

            'Medico
            Dim oRsTmp5 As New Recordset
            Set oRsTmp5 = mo_ReglasComunes.EmpleadosSeleccionarPorDNI(txtMedicoDni.Text)
            iFila = 104
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(0, 103).setFormula(txtMedicoDni.Text)
                Call Feuille.getcellbyposition(14, 103).setFormula(txtMedico.Text)
                Call Feuille.getcellbyposition(57, 103).setFormula(txtMedicoColegiatura.Text)
                Call Feuille.getcellbyposition(25, 105).setFormula(oRsTmp5.Fields!TipoEmpleado)
                Call Feuille.getcellbyposition(14, 105).setFormula(txtMedicoEspecialidad.Text)
                Call Feuille.getcellbyposition(52, 105).setFormula(txtMedicoRNE.Text)
                Call Feuille.getcellbyposition(63, 105).setFormula(IIf(chkMedicoEgresado.Value = 1, "X", ""))
            Else
                oWorkSheet.Cells(104, 1).Value = txtMedicoDni.Text
                oWorkSheet.Cells(104, 15).Value = txtMedico.Text
                oWorkSheet.Cells(104, 58).Value = txtMedicoColegiatura.Text
                oWorkSheet.Cells(106, 26).Value = oRsTmp5.Fields!TipoEmpleado
                oWorkSheet.Cells(106, 15).Value = txtMedicoEspecialidad.Text
                oWorkSheet.Cells(106, 53).Value = txtMedicoRNE.Text
                oWorkSheet.Cells(106, 64).Value = IIf(chkMedicoEgresado.Value = 1, "X", "")
            End If
            Set oRsTmp5 = Nothing

            'Se emite desde CITA, falta llenar ANEXO de Cpt y Farmacia
            lbMuestraCPTdefaults = False
            Select Case mo_lnIdTablaLISTBARITEMS
            Case sghOpcionGalenHos.sghRegistroCitaCE
               lbMuestraCPTdefaults = True
               If mo_SoloImprimeFUAyaGrabado = True Then
                  lbMuestraCPTdefaults = False
               End If
            Case sghOpcionGalenHos.sghAdmisionEmergencia
               If mo_lbEsAltaMedica = False Then
                  lbMuestraCPTdefaults = True
               End If
            End Select
            If lbMuestraCPTdefaults = True Then
               Dim orstmp As New Recordset, oRsTmp12 As New Recordset
               Dim oConexion As New Connection, oConexionExterna As New Connection
               Dim lcDescripcion As String, lcPuntoCarga As String, lnIdPuntoCarga As Long
               oConexionExterna.CommandTimeout = 300
               oConexionExterna.Open wxParametroJAMO
               oConexionExterna.CursorLocation = adUseClient
               oConexion.CommandTimeout = 300
               oConexion.Open sighentidades.CadenaConexion
               oConexion.CursorLocation = adUseClient
               Set orstmp = mo_ReglasSISgalenhos.FuaDefaultsCptFarmaciaSeleccionarTodos(oConexionExterna)
               If orstmp.RecordCount > 0 Then
                   orstmp.MoveFirst
                   Do While Not orstmp.EOF
                      If UCase(Trim(orstmp.Fields!tipo)) = "CPT" Then
                         Set oRsTmp12 = mo_AdminServiciosComunes.CatalogoServiciosSeleccionarPorCodigo(orstmp.Fields!codigo, oConexion)
                      Else
                         Set oRsTmp12 = mo_AdminServiciosComunes.CatalogoBienesInsumosSeleccionarPorCodigo(orstmp.Fields!codigo, oConexion)
                      End If
                      lcDescripcion = lcVacio
                      lcTipoMI = ""
                      If UCase(Trim(orstmp.Fields!tipo)) <> "CPT" Then
                         lcTipoMI = IIf(orstmp.Fields!esMedicamento = 1, lcMedicamento, lcInsumo)
                      End If
                      If oRsTmp12.RecordCount > 0 Then
                         lcDescripcion = Trim(oRsTmp12.Fields!nombre)
                         If UCase(Trim(orstmp.Fields!tipo)) <> "CPT" Then
                            lcTipoMI = IIf(oRsTmp12!TipoProducto = 1, lcInsumo, lcMedicamento)
                         End If
                      End If
                      oRsTmp12.Close
                      If UCase(Trim(orstmp.Fields!tipo)) = "CPT" Then
                         lcPuntoCarga = lcOtros
                         lnIdPuntoCarga = orstmp.Fields!idPuntoCarga
                         Select Case lnIdPuntoCarga
                         Case sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, sghPuntosCargaBasicos.sghPtoCargaBancoSangre1, sghPuntosCargaBasicos.sghPtoCargaBancoSangre2, sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica   'Laboratorio
                             lcPuntoCarga = lcLaboratorio
                         Case sghPuntosCargaBasicos.sghPtoCargaEcogGeneral, sghPuntosCargaBasicos.sghPtoCargaEcogObstetrica, sghPuntosCargaBasicos.sghPtoCargaRayosX, sghPuntosCargaBasicos.sghPtoCargaTomografia  'Imágenes
                             lcPuntoCarga = lcImagenes
                         Case sghPuntosCargaBasicos.sghPtoCargaFarmacia
                             lcPuntoCarga = lcFarmacia
                         End Select
                         oRsPatologia.AddNew
                         oRsPatologia.Fields!codigo = IIf(lcDescripcion = lcVacio, Space(2), orstmp.Fields!codigo)                                  'debb-09/06/2016
                         oRsPatologia.Fields!procedimiento = Left(IIf(lcDescripcion = lcVacio, Space(2), lcDescripcion & String(255, "_")), 255)    'debb-09/06/2016
                         oRsPatologia.Fields!dx = " "
                         oRsPatologia.Fields!tipo = lcPuntoCarga
                         oRsPatologia.Fields!idPuntoCarga = lnIdPuntoCarga
                         oRsPatologia.Update
                      Else
                         oRsFarmacia.AddNew
                         oRsFarmacia.Fields!codigo = IIf(lcDescripcion = lcVacio, Space(2), orstmp.Fields!codigo)                                   'debb-09/06/2016
                         oRsFarmacia.Fields!MedicInsumo = Left(IIf(lcDescripcion = lcVacio, Space(2), lcDescripcion & String(255, "_")), 255)       'debb-09/06/2016
                         oRsFarmacia.Fields!tipo = lcTipoMI
                         oRsFarmacia.Fields!dx = " "
                         oRsFarmacia.Update
                      End If
                      orstmp.MoveNext
                   Loop
               End If
               orstmp.Close
               Set orstmp = Nothing
               Set oRsTmp12 = Nothing
               oConexion.Close
               oConexionExterna.Close
               Set oConexion = Nothing
               Set oConexionExterna = Nothing
            Else
            End If
            'calcula las Paginas totales y añade lineas vacias
            oRsFarmacia.Filter = ""
            oRsPatologia.Filter = ""
            lnTotalLineas = Round((oRsFarmacia.RecordCount + oRsPatologia.RecordCount) / 2, 0)
            lnMaximaLineaPorPagina = 175
            lnTotalPaginas = 2
            If lnTotalLineas > lnMaximaLineaPorPagina Then
                lnTotalPaginas = Round(lnTotalLineas / lnMaximaLineaPorPagina, 0) + 1
                lnFilasAinsertar = lnTotalLineas - lnMaximaLineaPorPagina
                If lbEsOpenOffice = True Then
                    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                    PrintArea(0).Sheet = 0
                    PrintArea(0).startcolumn = 0
                    PrintArea(0).StartRow = 0
                    PrintArea(0).EndColumn = 67
                    PrintArea(0).EndRow = lnTotalPaginas * lnMaximaLineaPorPagina '73   'Trim(Str(lnTotalPaginas * 55))
                    Call Feuille.SetPrintAreas(PrintArea())
                    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                Else
                    oWorkSheet.PageSetup.PrintArea = "$A$1:$BL$" & Trim(Str(lnTotalPaginas * lnMaximaLineaPorPagina))
                End If
            Else
                If lbEsOpenOffice = True Then
                    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                    PrintArea(0).Sheet = 0
                    PrintArea(0).startcolumn = 0
                    PrintArea(0).StartRow = 0
                    PrintArea(0).EndColumn = 67
                    PrintArea(0).EndRow = IIf(lbMuestraCPTdefaults = True, 320, 2 * lnMaximaLineaPorPagina)
                    Call Feuille.SetPrintAreas(PrintArea())
                    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                Else
                    If mo_SoloImprimeFUAyaGrabado = True Or lcBuscaParametro.SeleccionaFilaParametro(535) = "S" Then
                       oWorkSheet.PageSetup.PrintArea = "$A$1:$BL$125"
                    Else
                    oWorkSheet.PageSetup.PrintArea = "$A$1:$BL$" & IIf(lbMuestraCPTdefaults = True, _
                                                                       "320", Trim(Str(2 * lnMaximaLineaPorPagina)))
                    End If
                End If
            End If
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(43, 126).setFormula(Trim(txtFua1.Text))
                Call Feuille.getcellbyposition(49, 126).setFormula(Trim(txtFua2.Text))
                Call Feuille.getcellbyposition(55, 126).setFormula(Trim(txtFua3.Text))
            Else
                oWorkSheet.Cells(127, 44).Value = Trim(txtFua1.Text)
                oWorkSheet.Cells(127, 50).Value = Trim(txtFua2.Text)
                oWorkSheet.Cells(127, 56).Value = Trim(txtFua3.Text)
            End If
            'Medicamentos
            iFila = 132
            oRsFarmacia.Filter = "tipo='" & lcMedicamento & "'"
            If oRsFarmacia.RecordCount > 0 Then
                oRsFarmacia.MoveFirst
                lbDerecha = True
                Do While Not oRsFarmacia.EOF
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 0, 37), iFila - 1).setFormula(Trim(oRsFarmacia.Fields!codigo))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 41), iFila - 1).setFormula(Left(oRsFarmacia.Fields!MedicInsumo, 65))
                        If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or IIf(mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia And mo_lbEsAltaMedica = False, True, False) Then
                        Else
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 27, 60), iFila - 1).setFormula("'" & IIf(oRsFarmacia.Fields!recetado = 0, "", oRsFarmacia.Fields!recetado))
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 31, 62), iFila - 1).setFormula("'" & IIf(oRsFarmacia.Fields!cantidad = 0, "", oRsFarmacia.Fields!cantidad))
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 34, 63), iFila - 1).setFormula("'" & oRsFarmacia.Fields!dx)
                        End If
                        'Falta metodo grosor de fila en OpenOffice
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 1, 38)).Value = Trim(oRsFarmacia.Fields!codigo)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 42)).Value = Left(oRsFarmacia.Fields!MedicInsumo, 65)
                        If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or IIf(mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia And mo_lbEsAltaMedica = False, True, False) Then
                        Else
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 28, 61)).Value = "'" & IIf(oRsFarmacia.Fields!recetado = 0, "", oRsFarmacia.Fields!recetado)
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 32, 63)).Value = "'" & IIf(oRsFarmacia.Fields!cantidad = 0, "", oRsFarmacia.Fields!cantidad)
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 35, 64)).Value = "'" & oRsFarmacia.Fields!dx
                        End If
                        oWorkSheet.Range(oWorkSheet.Cells(iFila, 1), oWorkSheet.Cells(iFila, 64)).RowHeight = 21.75
                    End If
                  oRsFarmacia.MoveNext
                  If lbDerecha = True Then
                     lbDerecha = False
                  Else
                     lbDerecha = True
                     iFila = iFila + 1
                  End If
               Loop
               iFila = iFila + 1
            End If

            'Insumos
            iFila = 160
            oRsFarmacia.Filter = "tipo='" & lcInsumo & "'"
            If oRsFarmacia.RecordCount > 0 Then
                oRsFarmacia.MoveFirst
                lbDerecha = True
                Do While Not oRsFarmacia.EOF
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 0, 37), iFila - 1).setFormula(Trim(oRsFarmacia.Fields!codigo))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 41), iFila - 1).setFormula(Left(oRsFarmacia.Fields!MedicInsumo, 65))
                        If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or IIf(mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia And mo_lbEsAltaMedica = False, True, False) Then
                        Else
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 27, 60), iFila - 1).setFormula("'" & IIf(oRsFarmacia.Fields!recetado = 0, "", oRsFarmacia.Fields!recetado))
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 31, 62), iFila - 1).setFormula("'" & IIf(oRsFarmacia.Fields!cantidad = 0, "", oRsFarmacia.Fields!cantidad))
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 34, 63), iFila - 1).setFormula("'" & oRsFarmacia.Fields!dx)
                        End If
                        'Falta metodo grosor de fila en OpenOffice
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 1, 38)).Value = Trim(oRsFarmacia.Fields!codigo)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 42)).Value = Left(oRsFarmacia.Fields!MedicInsumo, 65)
                        If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or IIf(mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia And mo_lbEsAltaMedica = False, True, False) Then
                        Else
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 28, 61)).Value = "'" & IIf(oRsFarmacia.Fields!recetado = 0, "", oRsFarmacia.Fields!recetado)
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 32, 63)).Value = "'" & IIf(oRsFarmacia.Fields!cantidad = 0, "", oRsFarmacia.Fields!cantidad)
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 35, 64)).Value = "'" & oRsFarmacia.Fields!dx
                        End If
                        oWorkSheet.Range(oWorkSheet.Cells(iFila, 1), oWorkSheet.Cells(iFila, 64)).RowHeight = 21.75
                    End If
                  oRsFarmacia.MoveNext
                  If lbDerecha = True Then
                     lbDerecha = False
                  Else
                     lbDerecha = True
                     iFila = iFila + 1
                  End If
               Loop
               iFila = iFila + 1
            End If

            'Laboratorio
            iFila = 188
            oRsPatologia.Filter = "tipo='" & lcLaboratorio & "'"
            If oRsPatologia.RecordCount > 0 Then
                oRsPatologia.MoveFirst
                lbDerecha = True
                Do While Not oRsPatologia.EOF
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 0, 37), iFila - 1).setFormula(Trim(oRsPatologia.Fields!codigo))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 41), iFila - 1).setFormula(Left(oRsPatologia.Fields!procedimiento, 65))
                        If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or IIf(mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia And mo_lbEsAltaMedica = False, True, False) Then
                        Else
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 21, 56), iFila - 1).setFormula("'" & IIf(oRsPatologia.Fields!indicado = 0, "", oRsPatologia.Fields!indicado))
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 27, 60), iFila - 1).setFormula("'" & IIf(oRsPatologia.Fields!ejecutado = 0, "", oRsPatologia.Fields!ejecutado))
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 31, 62), iFila - 1).setFormula("'" & oRsPatologia.Fields!dx)
                        End If
                        'Falta metodo grosor de fila en OpenOffice
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 1, 38)).Value = Trim(oRsPatologia.Fields!codigo)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 42)).Value = Left(oRsPatologia.Fields!procedimiento, 65)
                        If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or IIf(mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia And mo_lbEsAltaMedica = False, True, False) Then
                        Else
                            
                            'HRA 23/11/2020 Cambio 48 Inicio
                            If ucSISfuaCodPrestacion1.CodigoPrestacion = "071" Then
                           ' oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 22, 57)).Value = "'" & IIf(oRsPatologia.Fields!indicado = 0, "", oRsPatologia.Fields!indicado)
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 22, 57)).Value = "'" & IIf(oRsPatologia.Fields!ejecutado = 0, "", oRsPatologia.Fields!ejecutado) 'raul
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 28, 61)).Value = "'" & IIf(oRsPatologia.Fields!ejecutado = 0, "", oRsPatologia.Fields!ejecutado)
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 32, 63)).Value = "'" & oRsPatologia.Fields!dx
                            Else
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 22, 57)).Value = "'" & IIf(oRsPatologia.Fields!indicado = 0, "", oRsPatologia.Fields!indicado)
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 28, 61)).Value = "'" & IIf(oRsPatologia.Fields!ejecutado = 0, "", oRsPatologia.Fields!ejecutado)
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 32, 63)).Value = "'" & oRsPatologia.Fields!dx
                            End If
                            'HRA 23/11/2020 Cambio 48 Fin
                            
                            
                        End If
                        oWorkSheet.Range(oWorkSheet.Cells(iFila, 1), oWorkSheet.Cells(iFila, 64)).RowHeight = 21.75
                    End If
                  oRsPatologia.MoveNext
                  If lbDerecha = True Then
                     lbDerecha = False
                  Else
                     lbDerecha = True
                     iFila = iFila + 1
                  End If
               Loop
               iFila = iFila + 1
            End If

            'Imágenes
            oRsPatologia.Filter = "tipo='" & lcImagenes & "'"
            If oRsPatologia.RecordCount > 0 Then
                oRsPatologia.MoveFirst
                lbDerecha = True
                Do While Not oRsPatologia.EOF
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 0, 37), iFila - 1).setFormula(Trim(oRsPatologia.Fields!codigo))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 41), iFila - 1).setFormula(Left(oRsPatologia.Fields!procedimiento, 65))
                        If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or IIf(mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia And mo_lbEsAltaMedica = False, True, False) Then
                        Else
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 21, 56), iFila - 1).setFormula("'" & IIf(oRsPatologia.Fields!indicado = 0, "", oRsPatologia.Fields!indicado))
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 27, 60), iFila - 1).setFormula("'" & IIf(oRsPatologia.Fields!ejecutado = 0, "", oRsPatologia.Fields!ejecutado))
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 31, 62), iFila - 1).setFormula("'" & oRsPatologia.Fields!dx)
                        End If
                        'Falta metodo grosor de fila en OpenOffice
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 1, 38)).Value = Trim(oRsPatologia.Fields!codigo)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 42)).Value = Left(oRsPatologia.Fields!procedimiento, 65)
                        If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or IIf(mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia And mo_lbEsAltaMedica = False, True, False) Then
                        Else
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 22, 57)).Value = "'" & IIf(oRsPatologia.Fields!indicado = 0, "", oRsPatologia.Fields!indicado)
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 28, 61)).Value = "'" & IIf(oRsPatologia.Fields!ejecutado = 0, "", oRsPatologia.Fields!ejecutado)
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 32, 63)).Value = "'" & oRsPatologia.Fields!dx
                        End If
                        oWorkSheet.Range(oWorkSheet.Cells(iFila, 1), oWorkSheet.Cells(iFila, 64)).RowHeight = 21.75
                    End If
                  oRsPatologia.MoveNext
                  If lbDerecha = True Then
                     lbDerecha = False
                  Else
                     lbDerecha = True
                     iFila = iFila + 1
                  End If
               Loop
               iFila = iFila + 1
            End If

            'Otros CPT
            oRsPatologia.Filter = "tipo='" & lcOtros & "'"
            If oRsPatologia.RecordCount > 0 Then
                oRsPatologia.MoveFirst
                lbDerecha = True
                Do While Not oRsPatologia.EOF
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 0, 37), iFila - 1).setFormula(Trim(oRsPatologia.Fields!codigo))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 41), iFila - 1).setFormula(Left(oRsPatologia.Fields!procedimiento, 65))
                        If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or IIf(mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia And mo_lbEsAltaMedica = False, True, False) Then
                        Else
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 21, 56), iFila - 1).setFormula("'" & IIf(oRsPatologia.Fields!indicado = 0, "", oRsPatologia.Fields!indicado))
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 27, 60), iFila - 1).setFormula("'" & IIf(oRsPatologia.Fields!ejecutado = 0, "", oRsPatologia.Fields!ejecutado))
                            Call Feuille.getcellbyposition(IIf(lbDerecha = True, 31, 62), iFila - 1).setFormula("'" & oRsPatologia.Fields!dx)
                        End If
                        'Falta metodo grosor de fila en OpenOffice
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 1, 38)).Value = Trim(oRsPatologia.Fields!codigo)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 42)).Value = Left(oRsPatologia.Fields!procedimiento, 65)
                        If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or IIf(mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia And mo_lbEsAltaMedica = False, True, False) Then
                        Else
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 22, 57)).Value = "'" & IIf(oRsPatologia.Fields!indicado = 0, "", oRsPatologia.Fields!indicado)
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 28, 61)).Value = "'" & IIf(oRsPatologia.Fields!ejecutado = 0, "", oRsPatologia.Fields!ejecutado)
                            oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 32, 63)).Value = "'" & oRsPatologia.Fields!dx
                        End If
                        oWorkSheet.Range(oWorkSheet.Cells(iFila, 1), oWorkSheet.Cells(iFila, 64)).RowHeight = 21.75
                    End If
                  oRsPatologia.MoveNext
                  If lbDerecha = True Then
                     lbDerecha = False
                  Else
                     lbDerecha = True
                     iFila = iFila + 1
                  End If
               Loop
               iFila = iFila + 1
            End If

            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(0, 241).setFormula(txtObservaciones.Text)
            Else
                oWorkSheet.Cells(242, 1).Value = txtObservaciones.Text
            End If
            '
            oRsPatologia.Filter = ""
            oRsFarmacia.Filter = ""
            'Eliminar Cpts, Farmacia solo si se emitió el FUA desde CITAS
            If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia Then
               If oRsFarmacia.RecordCount > 0 Then
                    oRsFarmacia.MoveFirst
                    Do While Not oRsFarmacia.EOF
                       oRsFarmacia.Delete
                       oRsFarmacia.Update
                       oRsFarmacia.MoveNext
                    Loop
               End If
               If oRsPatologia.RecordCount > 0 Then
                    oRsPatologia.MoveFirst
                    Do While Not oRsPatologia.EOF
                       oRsPatologia.Delete
                       oRsPatologia.Update
                       oRsPatologia.MoveNext
                    Loop
               End If
            End If

            If lbEsOpenOffice = True Then
                If lcBuscaParametro.SeleccionaFilaParametro(338) = "S" Then
                    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                    CallByName Document, "print", VbMethod, PrintArgs()
                    Sleep (1000)
                    Call Document.Close(True)
                    If Dir$(PathFileOpenOffice, vbArchive) <> "" Then
                        Kill PathFileOpenOffice
                    End If
                Else
                    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                    MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
                End If
            Else
                    If mo_SoloImprimeFUAyaGrabado = True And wxParametro338 = "S" Then
                        oWorkSheet.PrintOut
                        oWorkBook.Close SaveChanges:=False
                    ElseIf wxParametro338 = "S" Then
                        oWorkSheet.PrintOut
                        oWorkBook.Close SaveChanges:=False
                    Else
                        oExcel.Visible = True
                        oWorkSheet.PrintPreview
                    End If
            End If
        End If
        If lbEsOpenOffice = True Then
            'Liberar Memoria
            Set Plage = Nothing
            Set Feuille = Nothing
            Set Document = Nothing
            Set Desktop = Nothing
            Set ServiceManager = Nothing
            Set Style = Nothing
            Set Border = Nothing
            'encabezado de pagina
            Set PageStyles = Nothing
            Set Sheet = Nothing
            Set StyleFamilies = Nothing
            Set DefPage = Nothing
            Set Htext = Nothing
            Set Hcontent = Nothing
        Else
            'Liberar memoria
            Set oExcel = Nothing
            Set oWorkBookPlantilla = Nothing
            Set oWorkBook = Nothing
            Set oWorkSheet = Nothing
        End If
    
        SeteaOtraImpresoraDefault sighentidades.ImpresoraDefaultDeEstaPC

        
        Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        Resume
        MsgBox Err.Description
    End Select

End Sub

Sub ChequeaSiHaySaltoDePagina(ByRef lnFila As Long, oWorkSheet As Worksheet)
       lnFila = lnFila + 1
End Sub
Sub ChequeaSiHaySaltoDePaginaOpenOffice(ByRef lnFila As Long)
       lnFila = lnFila + 1
End Sub

Sub ExcelCuadricularRango(oExcelApp As Excel.Application, oWorkSheet As Worksheet, lFilaIni As Long, lColumnaIni As Integer, lFilaFin As Long, lColumnaFin As Integer)
On Error Resume Next

    oWorkSheet.Range(oWorkSheet.Cells(lFilaIni, lColumnaIni), oWorkSheet.Cells(lFilaFin, lColumnaFin)).Select
    
    With oExcelApp.Selection.Borders(xlEdgeBottom)
        .Weight = xlMedium
        .ColorIndex = vbBlack    '41
    End With
    With oExcelApp.Selection.Borders(xlEdgeTop)
        .Weight = xlMedium
        .ColorIndex = vbBlack
    End With
    With oExcelApp.Selection.Borders(xlEdgeRight)
        .Weight = xlMedium
        .ColorIndex = vbBlack
    End With
    With oExcelApp.Selection.Borders(xlEdgeLeft)
        .Weight = xlMedium
        .ColorIndex = vbBlack
    End With
'    With oExcelApp.Selection.Borders(xlInsideVertical)
'        .Weight = xlThin
'        .ColorIndex = vbBlack
'    End With
'    With oExcelApp.Selection.Borders(xlInsideHorizontal)
'        .Weight = xlThin
'        .ColorIndex = vbBlack
'    End With

End Sub

Sub CargaDatosAlObjetosDeDatos()
    Dim lcEstablecimiento As String, lcCodigoSis As String
    txtFua3.Text = Right("00000000" & Trim(txtFua3.Text), 8)
    With oDoSisFuaAtencion
        .IdCuentaAtencion = ml_IdCuentaAtencion
        .FuaDisa = txtFua1.Text
        .FuaLote = txtFua2.Text
        .FuaNumero = Right("00000000" & txtFua3.Text, 8)
        .EstablecimientoCodigoRenaes = txtCScodigo.Text
        .Reconsideracion = "N"
        .ReconsideracionCodigoDisa = ""
        .ReconsideracionLote = "" 'IIf(Me.chkReconsideracion.Value = 1, txtFua2.Text, "")
        .ReconsideracionNroFormato = "" 'IIf(Me.chkReconsideracion.Value = 1, Right("00000000" & txtFua3.Text, 8), "")
        .FuaComponente = "4" 'CargaComponente(chkCsubsidiado.Value, chkCSemiS.Value)
        .Situacion = 2
        .AfiliacionDisa = txtNroAfiliacion1.Text
        .AfiliacionTipoFormato = txtNroAfiliacion2.Text
        .AfiliacionNroFormato = txtNroAfiliacion3.Text
        '.CodigoTipoFormato                                                      'no va en galenhos
        .OrigenAseguradoInstitucion = "0"
        '.OrigenAseguradoCodigo                                                  'no va en galenhos
        '.Edad                                                                   'no va en galenhos
        .GrupoEtareo = "0"
        .Genero = IIf(UCase(Left(txtSexo.Text, 1)) = "M", 1, 0)
        .FuaAtencion = CargaAtencion(chkAtencionAmbulatoria.Value, chkAtencionReferencia.Value, chkAtencionEmergencia.Value)
        .FuaCondicionMaterna = CargaCondicionMaterna(chkGestante.Value, chkPuerpera.Value)
        .FuaNrohistoria = Left(txtNhistoriaClinica.Text, 20)
        .FuaConceptoPr = IIf(Val(mo_cmbConceptoP.BoundText) = 0, 1, Val(mo_cmbConceptoP.BoundText))
        .FuaConceptoPrAutoriz = txtNautorizacion.Text
        .FuaConceptoPrMonto = Val(txtMonto.Text)
        .FuaAtencionFecha = Format(CDate(txtFantencion.Text), sighentidades.DevuelveFechaSoloFormato_DMY)
        .FuaAtencionHora = txtHatencion.Text        '
        mo_ReglasSISgalenhos.Sis_m_ee_ssDevuelveCodigoSIS txtRDcodigo.Text, lcCodigoSis, lcEstablecimiento
        .FuaReferidoDestinoCodigoRENAES = lcCodigoSis
        .FuaReferidoDestinoNreferencia = txtRDnumero.Text
        .FuaCodigoPrestacion = ucSISfuaCodPrestacion1.CodigoPrestacion
        .FuaPersonalQatiende = CargaOrigenPersonal(chkPAestablecimiento.Value, chkPAaisped.Value, chkPAOfeFlexible.Value)
        .FuaAtencionLugar = CargaLugarAtencion(chkIntramural.Value, chkExtramural.Value)
        .FuaDestino = IIf(Val(mo_cmbIdDestinoAtencion.BoundText) = 0, 1, Val(mo_cmbIdDestinoAtencion.BoundText)) 'Frank
        
        If .FuaDestino = 8 Then 'Coordinado Entre Rosa Celio (SIS) y Esteban Juarez 28/05/2015 - correo
            .FuaHospitalizadoFingreso = .FuaAtencionFecha
        Else
            If sighentidades.EsFecha(txtHfingreso.Text, "DD/MM/AAAA") = True Then
                .FuaHospitalizadoFingreso = Format(CDate(txtHfingreso.Text), sighentidades.DevuelveFechaSoloFormato_DMY)
            End If
        End If
            
        If .FuaDestino = 8 Then 'Coordinado Entre Rosa Celio (SIS) y Esteban Juarez 28/05/2015 - correo
            .FuaHospitalizadoFalta = .FuaAtencionFecha
        Else
            If sighentidades.EsFecha(txtHfalta.Text, "DD/MM/AAAA") = True Then
                .FuaHospitalizadoFalta = Format(CDate(txtHfalta.Text), sighentidades.DevuelveFechaSoloFormato_DMY)
            End If
        End If
        
        .FuaMedicoDNI = txtMedicoDni.Text
        .FuaMedico = Left(txtMedico.Text, 120)
        .FuaMedicoTipo = txtMedicoEspecialidad.Text
        If .FuaComponente = 2 Then
            .AfiliacionNroIntegrante = lcAfiliacionNroIntegrante
        Else
            .AfiliacionNroIntegrante = ""
        End If
        .codigo = lcAfiliacionCodigo
        .idSiasis = lcAfiliacionIdSiaSis
        .FuaObservaciones = Left(txtObservaciones.Text, 200)
        .CabDniUsuarioRegistra = CargaDNIusuarioRegistra
        .UltimaFechaAddMod = lcBuscaParametro.RetornaFechaServidorSQL
        .CabEstado = IIf(mi_opcion = sghEliminar, "1", "0")
        If sighentidades.EsFecha(txtFparto.Text, "DD/MM/AAAA") Then
           .FuaFechaParto = Format(CDate(txtFparto.Text), sighentidades.DevuelveFechaSoloFormato_DMY)
        Else
           .FuaFechaParto = ""
        End If
        .EstablecimientoDistrito = wxParametro242
        .Anio = Trim(Str(Year(CDate(txtFantencion.Text))))
        .Mes = Trim(Str(Month(CDate(txtFantencion.Text))))
        .CostoTotal = 0
        .Apaterno = ml_ApellidoPaterno
        .Amaterno = ml_ApellidoMaterno
        .Pnombre = ml_PrimerNombre
        .Onombre = ml_SegundoNombre
        .Fnacimiento = Format(md_FechaNacimiento, sighentidades.DevuelveFechaSoloFormato_DMY)
        '.Autogenerado                                                                       'no va en galenhos
        .DocumentoTipo = mo_cmbTipoDocumento.BoundText
        .DocumentoNumero = Me.txtNdocumento.Text
        .EstablecimientoCategoria = wxParametro303
        .CostoServicio = 0
        .CostoMedicamento = 0
        .CostoProcedimiento = 0
        .CostoInsumo = 0
        .MedicoDocumentoTipo = "1"    'dni
        '.ate_grupoRiesgo                                                                    'no va en galenhos
        .CabCodigoPuntoDigitacion = wxParametro304
        .CabCodigoUDR = wxParametro305
        '.CabNroEnvioAlSIS                                     'Se actualiza al EXPORTAR DESDE GALENHOS
        .CabOrigenDelRegistro = lcGalenHosNombre
        .CabVersionAplicativo = lcGalenHosVersion
        .CabIdentificacionPaquete = 0
        '.IdentificacionArfsis                                                                'no va en galenhos
        If mi_opcion = sghAgregar Then
           .CabFechaFuaPrimeraVez = lcBuscaParametro.RetornaFechaServidorSQL
        End If
        '.PeriodoOrigen                                                                        'no va en galenhos
        .IdUsuarioAuditoria = ml_idUsuario
        'mo_ReglasSISgalenhos.Sis_m_ee_ssDevuelveCodigoSIS txtROcodigo.Text, lcCodigoSis, lcEstablecimiento
        .FuaReferidoOrigenCodigoRENAES = txtROcodigo.Text        '
        .FuaReferidoOrigenNreferencia = txtRONumero.Text
        
        If mi_FuaTipoAnexo2015 = lcFuaAnexo1 Then
            .FuacolegioCodigo = Trim(txtColegioCodigo.Text)
            .FuacolegioNivel = mo_cmbColegioNivel.BoundText
            .FuacolegioGrado = mo_cmbColegioGrado.BoundText
            .FuacolegioSeccion = Trim(txtColegioSeccion.Text)
            .FuacolegioTurno = mo_cmbColegioTurno.BoundText
        End If
        
        If cmbEtnia.ListIndex < 0 Then
           .Fuaetnia = ""
        Else
           oCampos = Split(cmbEtnia.List(cmbEtnia.ListIndex), "|")
           .Fuaetnia = Right("0" & oCampos(0), 2)
        End If
        .FuaVersionFormato = mc_FuaVersionFormato
        .FuaTipoAnexo2015 = mi_FuaTipoAnexo2015
        If sighentidades.EsFecha(txtHFCortAdmin.Text, "DD/MM/AAAA") = True Then
            .FuaFechaCorteAdm = Format(CDate(txtHFCortAdmin.Text), sighentidades.DevuelveFechaSoloFormato_DMY)
        End If
        If cmbUPSfua.Text = "" Then
            .FuaUPS = ""
        Else
             oCampos = Split(cmbUPSfua.List(cmbUPSfua.ListIndex), "|")
            .FuaUPS = oCampos(1)
        End If
        .FuaCodAutorizacion = Trim(txtCodAutorizacion.Text)
        
        If sighentidades.EsFecha(txtFFallecimiento.Text, "DD/MM/AAAA") Then
           .FuafechaFallecimiento = Format(CDate(txtFFallecimiento.Text), sighentidades.DevuelveFechaSoloFormato_DMY)
        Else
           .FuafechaFallecimiento = 0
        End If
        If .FuaPersonalQatiende = 3 Then
            .FuaCodOferFlexible = txtPACodOfFlexible.Text
        Else
            .FuaCodOferFlexible = ""
        End If
    End With
End Sub

Function CargaComponente(lnSubsidiado As Integer, lnCSemiS As Integer) As String
    CargaComponente = "4"
    If lnSubsidiado <> 0 And lnCSemiS <> 0 Then
       CargaComponente = "3"
    ElseIf lnSubsidiado <> 0 Then
       CargaComponente = "1"
    ElseIf lnCSemiS <> 0 Then
       CargaComponente = "2"
    End If
End Function

Function CargaAtencion(lnAtencionAmbulatoria As Integer, lnAtencionReferencia As Integer, lnAtencionEmergencia As Integer) As Long
    If lnAtencionAmbulatoria <> 0 Then
       CargaAtencion = 1
    ElseIf lnAtencionReferencia <> 0 Then
       CargaAtencion = 2
    Else
       CargaAtencion = 3
    End If
End Function

Function CargaSepelioTipo(lnSepelioNatimuerto As Integer, lnSepelioObito As Integer, lnSepelioOtro As Integer) As Long
    If lnSepelioNatimuerto <> 0 Then CargaSepelioTipo = 1
    If lnSepelioObito <> 0 Then CargaSepelioTipo = 2
    If lnSepelioOtro <> 0 Then CargaSepelioTipo = 3
End Function

Function CargaCondicionMaterna(lnGestante As Integer, lnPuerpera As Integer) As String
    CargaCondicionMaterna = "0"
    If lnGestante <> 0 And lnPuerpera <> 0 Then
       CargaCondicionMaterna = "3"
    ElseIf lnGestante <> 0 Then
       CargaCondicionMaterna = "1"
    ElseIf lnPuerpera <> 0 Then
       CargaCondicionMaterna = "2"
    End If
End Function

Function CargaOrigenPersonal(lnPAestablecimiento As Integer, lnPAaisped As Integer, lnPAOfertaFlexible As Integer) As String
    If lnPAestablecimiento <> 0 Then CargaOrigenPersonal = 1
    If lnPAaisped <> 0 Then CargaOrigenPersonal = 2
    If lnPAOfertaFlexible <> 0 Then CargaOrigenPersonal = 3
End Function

Function CargaLugarAtencion(lnIntramural As Integer, lnExtramural As Integer) As String
    If lnIntramural <> 0 Then
       CargaLugarAtencion = "1"
    Else
       CargaLugarAtencion = "2"
    End If
End Function

Function CargaColegioGrado(lnInicial As Integer, lnPrimaria As Integer, lnSecundaria As Integer) As Integer
    If lnInicial <> 0 Then CargaColegioGrado = 1
    If lnPrimaria <> 0 Then CargaColegioGrado = 2
    If lnSecundaria <> 0 Then CargaColegioGrado = 3
End Function

Function CargaDNIusuarioRegistra() As String
    CargaDNIusuarioRegistra = ""
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = mo_AdminServiciosComunes.EmpleadosSeleccionarPorIdEmpleado(ml_idUsuario)
    If oRsTmp1.RecordCount > 0 Then
       CargaDNIusuarioRegistra = oRsTmp1.Fields!DNI
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Sub CargarDatosAlFormulario()
    If mi_opcion = sghAgregar And mo_lnIdTablaLISTBARITEMS = sghFormatoFUA Then
        ArsSisCargaDatosAlIniciarFormulario
    Else
        CargarDatosALosControles
        ConfiguraDiseñoPorAnexoNuevoFUA
        Select Case mi_opcion
             Case sghAgregar
             Case sghModificar
             Case sghConsultar
                Me.btnAceptar.Enabled = False
             Case sghEliminar
        End Select
        'lblAnexoFUA.Caption = "FUA Versión 2: Anexo " & CStr(mi_FuaTipoAnexo2015)
    End If
End Sub

Sub ConfiguraDiseñoPorAnexoNuevoFUA()
    If mi_FuaTipoAnexo2015 = lcFuaAnexo1 Then
        fraInstitucionEducativa.Visible = True
        chkPAOfeFlexible.Visible = True
        lblPACodigoOfFlexible.Visible = True
        txtPACodOfFlexible.Visible = True
        lblCodPrestAdicional.Visible = True
        txtCodPrestAdicional.Visible = True
        fraConcPrestacional.Left = 6480
        fraConcPrestacional.Width = 3105
        'debb-21/09/2015 (inicio)
        Frame5.Visible = True
        Frame(3).Visible = True
        Frame(5).Visible = True
        Frame(6).Visible = True
        fraTamizajeSaludM.Top = 3765
        'debb-21/09/2015 (fin)
    ElseIf mi_FuaTipoAnexo2015 = lcFuaAnexo2 Then
        fraInstitucionEducativa.Visible = False
        chkPAOfeFlexible.Visible = False
        lblPACodigoOfFlexible.Visible = False
        txtPACodOfFlexible.Visible = False
        lblCodPrestAdicional.Visible = False
        txtCodPrestAdicional.Visible = False
        fraConcPrestacional.Left = 4800
        fraConcPrestacional.Width = 4785
        'debb-21/09/2015 (inicio)
        Frame5.Visible = False
        Frame(3).Visible = False
        Frame(5).Visible = False
        Frame(6).Visible = False
        fraTamizajeSaludM.Top = 240
        'debb-21/09/2015 (fin)
    End If
End Sub

Sub ArsSisHabilitaAgregarYmodificar()
    'mo_Formulario.HabilitarDeshabilitar FraComponente, True
    'mo_Formulario.HabilitarDeshabilitar FraTipoAfiliacion, True
    mo_Formulario.HabilitarDeshabilitar cmbIdDestinoAtencion, True
    mo_Formulario.HabilitarDeshabilitar txtFantencion, True
    mo_Formulario.HabilitarDeshabilitar txtHatencion, True
    mo_Formulario.HabilitarDeshabilitar txtHfingreso, True
    mo_Formulario.HabilitarDeshabilitar txtHfalta, True
    mo_Formulario.HabilitarDeshabilitar txtHFCortAdmin, True
    mo_Formulario.HabilitarDeshabilitar txtSPpeso, True
    mo_Formulario.HabilitarDeshabilitar txtSPtalla, True
    mo_Formulario.HabilitarDeshabilitar txtSPpa, True
    FraDx.Enabled = True
    mo_Formulario.HabilitarDeshabilitar txtMedicoDni, True
    mo_Formulario.HabilitarDeshabilitar txtObservaciones, True
    
End Sub

Sub ArsSisCargaDatosAlIniciarFormulario()
    lcOpcion = "(todos los Datos se registran como el ArSis)"
    mi_opcion = sghAgregar
    ArsSisHabilitaAgregarYmodificar
    mo_Formulario.HabilitarDeshabilitar txtNroAfiliacion1, True
    mo_Formulario.HabilitarDeshabilitar txtNroAfiliacion2, True
    mo_Formulario.HabilitarDeshabilitar txtNroAfiliacion3, True
    btnBuscarPaciente.Enabled = True
    mo_Formulario.HabilitarDeshabilitar txtNhistoriaClinica, True
    lbEsIgualQueArSIS = True
    ml_IdCuentaAtencion = 0   'Antes de grabar = txtFua1.text &  txtFua2.text &  txtFua3.text
    txtFantencion.Text = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY)
    md_FechaAtencion = txtFantencion.Text
    txtHatencion.Text = Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
    ml_HoraAtencion = txtHatencion.Text
    txtCScodigo.Text = wxParametro280
    txtCS.Text = wxParametro205
    CargaFormatoFUA
    Mantenimiento
    '
'    If Val(wxParametro320) = sghFuaTipo.sghFuaTipoAutomatico Then
'       txtFua3.Text = Right("00000000" & Trim(Str(Val(mo_ReglasSISgalenhos.sisFuaAtencionUltimoCorrelativo()) + 1)), 8)
'    Else
        Dim oRsTmp1 As New Recordset
        Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaSeleccionarTodos
        If oRsTmp1.RecordCount > 0 Then
           txtFua3.Text = Right("00000000" & Trim(Str(Val(oRsTmp1.Fields!FuaUltimoGenerado) + 1)), 8)
        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
'        txtFua3.Text = ""
'    End If
End Sub

Sub CargarDatosALosControles()
       Dim oSisFuaAtencion As New SisFuaAtencion
       Dim oConexionExterna As New Connection
       oConexionExterna.CommandTimeout = 300
       oConexionExterna.CursorLocation = adUseClient
       oConexionExterna.Open wxParametroJAMO
       '
       oDoSisFuaAtencion.IdCuentaAtencion = ml_IdCuentaAtencion
       Set oSisFuaAtencion.Conexion = oConexionExterna
       If oSisFuaAtencion.SeleccionarPorId(oDoSisFuaAtencion) = True Then
            mi_opcion = sghModificar
            lcOpcion = "(jala datos grabados en tablas del SIS)"
            CargaDatosDelFuaDesdeTablasSIS oConexionExterna
            If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghFormatoFUA Then
               mi_opcion = mi_opcion_fua
            End If
       Else
            lcOpcion = "(jala datos grabados en tablas de SIGH)"
            mi_opcion = sghAgregar
            CargaDatosDelFuaDesdeTablasGalenHos
            
            If Val(wxParametro320) = sghFuaTipo.sghFuaTipoAutomatico Then
                If Val(mo_ReglasSISgalenhos.sisFuaAtencionUltimoCorrelativo()) > 0 Then
                    txtFua3.Text = Right("00000000" & Trim(Str(Val(mo_ReglasSISgalenhos.sisFuaAtencionUltimoCorrelativo()) + 1)), 8)
                Else
                    Dim oRsTmp2 As New Recordset
                    Set oRsTmp2 = mo_ReglasSISgalenhos.SisFuaSeleccionarTodos
                    If oRsTmp2.RecordCount > 0 Then
                       txtFua3.Text = Right("00000000" & Trim(Str(Val(IIf(IsNull(oRsTmp2.Fields!FuaUltimoGenerado), 0, oRsTmp2.Fields!FuaUltimoGenerado)) + 1)), 8) 'Actualizado 30092014
                    End If
                    oRsTmp2.Close
                    Set oRsTmp2 = Nothing
                End If
            Else
                Dim oRsTmp1 As New Recordset
                Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaSeleccionarTodos
                If oRsTmp1.RecordCount > 0 Then
                   txtFua3.Text = Right("00000000" & Trim(Str(Val(IIf(IsNull(oRsTmp1.Fields!FuaUltimoGenerado), 0, oRsTmp1.Fields!FuaUltimoGenerado)) + 1)), 8) 'Actualizado 30092014
                End If
                oRsTmp1.Close
                Set oRsTmp1 = Nothing
            '        txtFua3.Text = ""
            End If
          '
       End If
       Set oSisFuaAtencion = Nothing
       oConexionExterna.Close
       Set oConexionExterna = Nothing
       '
       ucSISfuaCodPrestacion1_LostFocus
       '
       Mantenimiento
End Sub

Public Function DevolverNombreColegio(ByVal mc_codigocolegio As String) As String
    Dim orstmp As New ADODB.Recordset
    DevolverNombreColegio = ""
    Set orstmp = mo_ReglasSISgalenhos.SisFuaColegiosSeleccionarPorCodigoNombre(mc_codigocolegio, "")
    If orstmp.RecordCount > 0 Then
        DevolverNombreColegio = orstmp.Fields!colegio
    End If
    Set orstmp = Nothing
End Function


Sub CargaDatosDelFuaDesdeTablasSIS(oConexionExterna As Connection)
    Dim oSisFuaAtencion As New SisFuaAtencion
    Dim oConexion As New Connection
    Dim lcCodigoRenaes As String, lcDescripcionRenaes As String
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    '
    If CargarDatosDelPaciente(oConexion) = True Then
        CargaDatosMedico oConexion, False
'        oConexion.Close
'        Set oConexion = Nothing
        Set oSisFuaAtencion.Conexion = oConexionExterna
        oDoSisFuaAtencion.IdCuentaAtencion = ml_IdCuentaAtencion
        If oSisFuaAtencion.SeleccionarPorId(oDoSisFuaAtencion) = False Then
           MsgBox "No se pudo Cargar FUA, desde tablas SIS", vbInformation, Me.Caption
           Me.Visible = False
           Exit Sub
        End If
        '
        With oDoSisFuaAtencion
               txtFua1.Text = .FuaDisa
               txtFua2.Text = .FuaLote
               txtFua3.Text = Right("00000000" & CStr(.FuaNumero), 8)
               txtCScodigo.Text = .EstablecimientoCodigoRenaes
               txtCS.Text = wxParametro205
               txtNroAfiliacion1.Text = .AfiliacionDisa
               txtNroAfiliacion2.Text = .AfiliacionTipoFormato
               txtNroAfiliacion3.Text = .AfiliacionNroFormato
               mi_FuaTipoAnexo2015 = .FuaTipoAnexo2015  'FUA2015
               mo_cmbTipoDocumento.BoundText = .DocumentoTipo
               txtNdocumento.Text = .DocumentoNumero
               txtPaciente.Text = Trim(.Apaterno) & " " & Trim(.Amaterno) & " " & Trim(.Pnombre) & " " & IIf(IsNull(.Onombre), "", .Onombre)
               ml_ApellidoPaterno = .Apaterno
               ml_ApellidoMaterno = .Amaterno
               ml_PrimerNombre = .Pnombre
               ml_SegundoNombre = IIf(IsNull(.Onombre), "", .Onombre)
               txtFnacimiento.Text = DevuelveFechaSegunFormato_YMD_SIS(.Fnacimiento)
               md_FechaNacimiento = DevuelveFechaSegunFormato_YMD_SIS(.Fnacimiento)
               txtSexo.Text = IIf(.Genero = 1, lcMasculino, lcFemenino)
               txtNhistoriaClinica.Text = .FuaNrohistoria
               'FUA2015
               ml_Etnia = .Fuaetnia
               If ml_Etnia = "" Then
                  cmbEtnia.Text = lcBuscaParametro.SeleccionaFilaParametro(283)
               Else
                  cmbEtnia_UbicaPosicion (ml_Etnia)
               End If
               cmbUPSfua.Text = .FuaUPS
               If cmbUPSfua.Text <> "" Then
                    oCampos = Split(cmbUPSfua.List(cmbUPSfua.ListIndex), "|")
                    lblUpsFua.Text = oCampos(0)
               End If
               
               If .FuafechaFallecimiento <> 0 Then txtFFallecimiento.Text = DevuelveFechaSegunFormato_YMD_SIS(.FuafechaFallecimiento)
               ''''''''
               ''''''''
               'FUA2015
               If mi_FuaTipoAnexo2015 = lcFuaAnexo1 Then
                    txtColegioCodigo.Text = IIf(IsNull(.FuacolegioCodigo), "", .FuacolegioCodigo)
                    mo_cmbColegioNivel.BoundText = IIf(IsNull(.FuacolegioNivel), "", .FuacolegioNivel)
                    mo_cmbColegioGrado.BoundText = IIf(IsNull(.FuacolegioGrado), "", .FuacolegioGrado)
                    txtColegioSeccion.Text = IIf(IsNull(.FuacolegioSeccion), "", .FuacolegioSeccion)
                    mo_cmbColegioTurno.BoundText = IIf(IsNull(.FuacolegioTurno), "", .FuacolegioTurno)
                    txtColegio.Text = DevolverNombreColegio(.FuacolegioCodigo)
               End If
               txtCodAutorizacion.Text = .FuaCodAutorizacion
               mo_Formulario.HabilitarDeshabilitar txtPACodOfFlexible, False
               If .FuaPersonalQatiende = 3 Then
                    mo_Formulario.HabilitarDeshabilitar txtPACodOfFlexible, True
                    txtPACodOfFlexible.Text = IIf(IsNull(.FuaCodOferFlexible), "", .FuaCodOferFlexible)
               End If
               '''''''
               txtInstitucion.Text = .OrigenAseguradoInstitucion
               chkAtencionAmbulatoria.Value = IIf(.FuaAtencion = 1, 1, 0)
               chkAtencionReferencia.Value = IIf(.FuaAtencion = 2, 1, 0)
               chkAtencionEmergencia.Value = IIf(.FuaAtencion = 3, 1, 0)
               chkGestante.Value = IIf(.FuaCondicionMaterna = 1, 1, 0)
               chkPuerpera.Value = IIf(.FuaCondicionMaterna = 2, 1, 0)
               chkIntramural.Value = IIf(.FuaAtencionLugar = 1, 1, 0)
               chkExtramural.Value = IIf(.FuaAtencionLugar = 2, 1, 0)
               mo_cmbConceptoP.BoundText = .FuaConceptoPrAutoriz
               
               txtNautorizacion.Text = .FuaConceptoPrAutoriz
               txtMonto.Text = .FuaConceptoPrMonto
               txtFparto.Text = DevuelveFechaSegunFormato_YMD_SIS(.FuaFechaParto)
               chkPAestablecimiento.Value = IIf(.FuaPersonalQatiende = 1, 1, 0)
               chkPAaisped.Value = IIf(.FuaPersonalQatiende = 2, 1, 0)
               chkPAOfeFlexible.Value = IIf(.FuaPersonalQatiende = 3, 1, 0)
               ucSISfuaCodPrestacion1.CodigoPrestacion = .FuaCodigoPrestacion
               'ucSISfuaCodPrestacion1.AsignaDescripcionSegunCodigoPrestacion
               '
               lcCodigoRenaes = "": lcDescripcionRenaes = ""
               mo_ReglasSISgalenhos.Sis_m_ee_ssDevuelveCodigoRENAES .FuaReferidoOrigenCodigoRENAES, lcCodigoRenaes, lcDescripcionRenaes
'               txtROcodigo.Text = lcCodigoRenaes
               txtROcodigo.Text = .FuaReferidoOrigenCodigoRENAES
               Dim orstemp2 As New Recordset
               Dim oDOEstablecimiento As New DOEstablecimiento
               txtRO.Text = lcDescripcionRenaes
               If .FuaReferidoOrigenCodigoRENAES <> "" Then
                    If mo_AdminServiciosComunes.EstablecimientosSeleccionarPorCodigo(Right(.FuaReferidoOrigenCodigoRENAES, 5), oDOEstablecimiento) Then
                         txtRO.Text = oDOEstablecimiento.nombre
                    Else
                         Set orstemp2 = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorCodigo(.FuaReferidoOrigenCodigoRENAES)
                         If orstemp2.RecordCount > 0 Then
                             orstemp2.MoveFirst
                             txtRO.Text = orstemp2.Fields!nombre
                         End If
                    End If
               End If
               txtRONumero.Text = .FuaReferidoOrigenNreferencia
               mo_cmbIdDestinoAtencion.BoundText = .FuaDestino
               ml_IdDestinoPaciente = .FuaDestino
               lcCodigoRenaes = "": lcDescripcionRenaes = ""
               mo_ReglasSISgalenhos.Sis_m_ee_ssDevuelveCodigoRENAES .FuaReferidoDestinoCodigoRENAES, lcCodigoRenaes, lcDescripcionRenaes
               txtRDcodigo.Text = lcCodigoRenaes
               txtRD.Text = lcDescripcionRenaes
               txtRDnumero.Text = .FuaReferidoDestinoNreferencia
               ''''''''''''''''''''''''''''''''''''''''''''''''''''
               txtHfingreso.Text = DevuelveFechaSegunFormato_YMD_SIS(.FuaHospitalizadoFingreso)
               txtHfalta.Text = DevuelveFechaSegunFormato_YMD_SIS(.FuaHospitalizadoFalta)
               txtMedicoDni.Text = .FuaMedicoDNI
               txtMedico.Text = .FuaMedico
               If txtMedicoColegiatura.Text = "" Then
                  CargaDatosMedico oConexion, True
               End If
               txtMedicoEspecialidad.Text = .FuaMedicoTipo
               txtObservaciones.Text = .FuaObservaciones
               txtFantencion.Text = DevuelveFechaSegunFormato_YMD_SIS(.FuaAtencionFecha)
               txtHatencion.Text = .FuaAtencionHora
               mo_cmbConceptoP.BoundText = .FuaConceptoPr
               ml_IdConceptoPrestacional = .FuaConceptoPr
               txtNautorizacion.Text = .FuaConceptoPrAutoriz
               txtMonto.Text = .FuaConceptoPrMonto
        End With
        If txtHatencion.Text = sighentidades.HORA_VACIA_HM Then
             txtHatencion.Text = "00:01"
        End If

        ml_edad_En_Dias = sighentidades.EdadActualEnDias(CDate(txtFnacimiento.Text), CDate(txtFantencion.Text & " " & txtHatencion.Text))
        ml_edad_En_YYYYMMDD = sighentidades.EdadActualEnFormatoYYYYMMDD(CDate(txtFnacimiento.Text), CDate(txtFantencion.Text & " " & txtHatencion.Text))
        '
        CargaDatosDeDx oConexionExterna, False
        CargaConsumosEnServiciosIntermedios oConexionExterna, False
        CargaDatosDeTriajeVacunas oConexionExterna, False, oConexion
        '
        If Val(oDoSisFuaAtencion.CabNroEnvioAlSIS) > 0 Then
           Me.btnAceptar.Enabled = False
           lcOpcion = lcOpcion & " (Ya fué enviado al SIS CENTRAL)"
        End If
        CargaDatosDeNacimiento oConexion, False
        oConexion.Close
        Set oConexion = Nothing
    Else
        Me.btnAceptar.Enabled = False
        Me.btnImprimir.Enabled = False
        oConexion.Close
        Set oConexion = Nothing
    End If
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------
Function AgregarDatos() As Boolean
    lnNroFuaRepetido = False
    AgregarDatos = mo_ReglasSISgalenhos.FuaAgregar(oDoSisFuaAtencion, oRsVacunasSp, oRsPatologia, oRsFarmacia, oRsDx, _
                                        lcInsumo, lcMedicamento, ml_idUsuario, 1345, mo_lcNombrePc, _
                                        "Fua: " & Trim(txtFua1.Text) & "-" & Trim(txtFua2.Text) & Trim(txtFua3.Text) & _
                                        " - Cta: " & Trim(Str(ml_IdCuentaAtencion)) & " - " & Trim(Me.txtPaciente.Text), wxParametro320, ml_idAtencion, lnNroFuaRepetido)
End Function


'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean
    ModificarDatos = mo_ReglasSISgalenhos.FuaModificar(oDoSisFuaAtencion, oRsVacunasSp, oRsPatologia, oRsFarmacia, oRsDx, _
                                          lcInsumo, lcMedicamento, ml_idUsuario, 1345, mo_lcNombrePc, _
                                          "Fua: " & Trim(txtFua1.Text) & "-" & Trim(txtFua2.Text) & "-" & Trim(txtFua3.Text) & _
                                          "  Cta: " & Trim(Str(ml_IdCuentaAtencion)) & " - " & Trim(Me.txtPaciente.Text), ml_idAtencion)
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    EliminarDatos = mo_ReglasSISgalenhos.FuaEliminar(oDoSisFuaAtencion, ml_idUsuario, 1345, _
                                         mo_lcNombrePc, "Fua: " & Trim(txtFua1.Text) & "-" & Trim(txtFua2.Text) & "-" & Trim(txtFua3.Text) & _
                                         "  Cta: " & Trim(Str(ml_IdCuentaAtencion)) & " - " & Trim(Me.txtPaciente.Text), ml_idAtencion)
End Function

Sub CargarUPSFuaEnControl(lcUpsFua As String)
    If lcUpsFua <> "" Then
        cmbUPSfua.Text = lcUpsFua
        If cmbUPSfua.Text <> "" Then
             oCampos = Split(cmbUPSfua.List(cmbUPSfua.ListIndex), "|")
             lblUpsFua.Text = oCampos(0)
        End If
    End If
End Sub

Function CargarDatosDelPaciente(oConexion As Connection) As Boolean
        CargarDatosDelPaciente = True
        Dim oRsTmp1 As New Recordset
        Dim mo_DoAtencionDatosAdicionales As New DoAtencionDatosAdicionales
        Dim oDOEstablecimiento As New DOEstablecimiento
        Dim oDOEstablecimientoNoMinsa As New DOEstablecimientoNoMinsa
        Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
        Dim lcFiltro As String
        Dim lbEsAmbulatoria As Boolean, lbEsEmergencia As Boolean, lbEsReferencia As Boolean
        '
        ml_EsPacienteExterno = False
        Set oRsTmp1 = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(ml_IdCuentaAtencion, oConexion)
        If oRsTmp1.RecordCount > 0 Then
           If oRsTmp1!esPacienteExterno = True Then
              ml_EsPacienteExterno = True
           End If
           ml_idPaciente = oRsTmp1!idPaciente
           ml_fechaIngreso = Format(oRsTmp1!fechaIngreso & " " & oRsTmp1!horaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
           ml_IdOrigenAtencion = IIf(IsNull(oRsTmp1!idOrigenAtencion), 0, oRsTmp1!idOrigenAtencion)
           lcElServicioUsaGalenHos = mo_ReglasArchivoClinico.ServicioUsaGalenHos(oRsTmp1.Fields!IdServicioIngreso)
           ml_IdTipoServicio = oRsTmp1.Fields!IdTipoServicio
           ml_Paciente = Trim(oRsTmp1.Fields!ApellidoPaterno) & " " & Trim(oRsTmp1.Fields!ApellidoMaterno) & " " & Trim(oRsTmp1.Fields!PrimerNombre) & IIf(IsNull(oRsTmp1.Fields!SegundoNombre), "", " " & oRsTmp1.Fields!SegundoNombre)
           md_FechaNacimiento = oRsTmp1.Fields!FechaNacimiento
           ml_Sexo = IIf(oRsTmp1.Fields!idTipoSexo = 2, lcFemenino, lcMasculino)
           ml_Etnia = oRsTmp1.Fields!IdEtnia
           ml_NroHistoriaClinica = oRsTmp1.Fields!NroHistoriaClinica
           ml_ApellidoPaterno = oRsTmp1.Fields!ApellidoPaterno
           ml_ApellidoMaterno = oRsTmp1.Fields!ApellidoMaterno
           ml_PrimerNombre = oRsTmp1.Fields!PrimerNombre
           ml_SegundoNombre = IIf(IsNull(oRsTmp1.Fields!SegundoNombre), "", " " & oRsTmp1.Fields!SegundoNombre)
           ml_NroDocumento = IIf(IsNull(oRsTmp1.Fields!NroDocumento), "", oRsTmp1.Fields!NroDocumento)
           ml_edad_En_Dias = sighentidades.EdadActualEnDias(oRsTmp1.Fields!FechaNacimiento, oRsTmp1.Fields!fechaIngreso)
           ml_edad_En_YYYYMMDD = sighentidades.EdadActualEnFormatoYYYYMMDD(oRsTmp1.Fields!FechaNacimiento, CDate(Format(oRsTmp1.Fields!fechaIngreso & " " & oRsTmp1.Fields!horaIngreso, "dd/mm/yyyy hh:mm")))
           '
           ml_TipoDocumentoGalenhos = IIf(IsNull(oRsTmp1.Fields!IdDocIdentidad), 0, oRsTmp1.Fields!IdDocIdentidad)
           mo_cmbTipoDocumento.BoundText = mo_AdminAdmision.TiposDocIdentidadDevuelveIdSis(ml_TipoDocumentoGalenhos)
           '
           ml_HoraAtencion = IIf(ml_IdTipoServicio = sghConsultaExterna, oRsTmp1.Fields!horaIngreso, IIf(IsNull(oRsTmp1!HoraEgreso), oRsTmp1.Fields!horaIngreso, oRsTmp1!HoraEgreso))
           md_FechaAtencion = IIf(ml_IdTipoServicio = sghConsultaExterna, oRsTmp1.Fields!fechaIngreso, IIf(IsNull(oRsTmp1!fechaEgreso), oRsTmp1.Fields!fechaIngreso, oRsTmp1!fechaEgreso))
           
           'mgaray20140926
           If ml_IdTipoServicio = sghConsultaExterna Or ml_IdTipoServicio = 5 Or mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia Or mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghPacienteExternoConSeguro Then
                ml_IdMedico = oRsTmp1.Fields!IdMedicoIngreso
                mo_Formulario.HabilitarDeshabilitar txtHfingreso, False
                mo_Formulario.HabilitarDeshabilitar txtHfalta, False
                mo_Formulario.HabilitarDeshabilitar txtHFCortAdmin, False
                Set oDoServicio = mo_ReglasServiciosHosp.ServiciosSeleccionarPorId(oRsTmp1.Fields!IdServicioIngreso, oConexion)
                CargarUPSFuaEnControl (oDoServicio.codigoServicioFUA)
                Set oDoServicio = Nothing
                If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia Then
                    ml_IdMedico = IIf(IsNull(oRsTmp1.Fields!IdMedicoEgreso), oRsTmp1.Fields!IdMedicoIngreso, oRsTmp1.Fields!IdMedicoEgreso)    'debb-27/12/2016
                    Me.txtHfingreso.Text = Format(oRsTmp1.Fields!fechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY)
                    Me.txtHfalta.Text = IIf(IsNull(oRsTmp1.Fields!fechaEgreso), _
                                        sighentidades.FECHA_VACIA_DMY, _
                                        Format(oRsTmp1.Fields!fechaEgreso, sighentidades.DevuelveFechaSoloFormato_DMY)) 'Frank 2508
                    Set oDoServicio = mo_ReglasServiciosHosp.ServiciosSeleccionarPorId(IIf(IsNull(oRsTmp1.Fields!IdServicioEgreso), oRsTmp1.Fields!IdServicioIngreso, oRsTmp1.Fields!IdServicioEgreso), oConexion)
                    CargarUPSFuaEnControl (oDoServicio.codigoServicioFUA)
                    Set oDoServicio = Nothing
                End If
           Else
                ml_IdMedico = IIf(IsNull(oRsTmp1.Fields!IdMedicoEgreso), 0, oRsTmp1.Fields!IdMedicoEgreso)
                Me.txtHfingreso.Text = Format(oRsTmp1.Fields!fechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY)
                Me.txtHfalta.Text = IIf(IsNull(oRsTmp1.Fields!fechaEgreso), _
                                        sighentidades.FECHA_VACIA_DMY, _
                                        Format(oRsTmp1.Fields!fechaEgreso, sighentidades.DevuelveFechaSoloFormato_DMY)) 'Frank 2508
                mo_Formulario.HabilitarDeshabilitar Me.txtHfingreso, False
                mo_Formulario.HabilitarDeshabilitar Me.txtHfalta, False
                mo_Formulario.HabilitarDeshabilitar txtHFCortAdmin, False
                Set oDoServicio = mo_ReglasServiciosHosp.ServiciosSeleccionarPorId(IIf(IsNull(oRsTmp1.Fields!IdServicioEgreso), oRsTmp1.Fields!IdServicioIngreso, oRsTmp1.Fields!IdServicioEgreso), oConexion)
                CargarUPSFuaEnControl (oDoServicio.codigoServicioFUA)
                Set oDoServicio = Nothing
           End If
           ml_idAtencion = oRsTmp1.Fields!idAtencion
           If Not IsNull(oRsTmp1.Fields!idEstablecimientoOrigen) Or Not IsNull(oRsTmp1.Fields!IdEstablecimientoNoMinsaOrigen) Then
                If Not IsNull(oRsTmp1.Fields!idEstablecimientoOrigen) Then
                    Set oDOEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(oRsTmp1.Fields!idEstablecimientoOrigen)
                    If Not oDOEstablecimiento Is Nothing Then
                        txtRO.Text = oDOEstablecimiento.nombre
                        Me.txtROcodigo.Text = Right("0000000000" & oDOEstablecimiento.codigo, 10)
                    End If
                    Me.txtRONumero.Text = IIf(IsNull(oRsTmp1.Fields!nroReferenciaOrigen), "", oRsTmp1.Fields!nroReferenciaOrigen)
                Else
                    Set oDOEstablecimientoNoMinsa = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(oRsTmp1.Fields!IdEstablecimientoNoMinsaOrigen)
                    If Not oDOEstablecimiento Is Nothing Then
                        txtRO.Text = oDOEstablecimientoNoMinsa.nombre
                        Me.txtROcodigo.Text = Right("0000000000" & oDOEstablecimientoNoMinsa.codigo, 10)
                    End If
                End If
                Me.txtRONumero.Text = IIf(IsNull(oRsTmp1.Fields!nroReferenciaOrigen), "", oRsTmp1.Fields!nroReferenciaOrigen)
                Me.btnBuscarEstablecimientoO.Enabled = False
                mo_Formulario.HabilitarDeshabilitar Me.txtRONumero, False
                chkAtencionReferencia.Value = ssCBChecked
           End If
           'mgaray20140926
           If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia Or mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghPacienteExternoConSeguro Then
              Me.btnBuscarEstablecimientoO.Enabled = True
              mo_Formulario.HabilitarDeshabilitar Me.txtRONumero, True
           End If
           If Not IsNull(oRsTmp1.Fields!idEstablecimientoDestino) Or Not IsNull(oRsTmp1.Fields!IdEstablecimientoNoMinsaDestino) Then
                If Not IsNull(oRsTmp1.Fields!idEstablecimientoDestino) Then
                    Set oDOEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(oRsTmp1.Fields!idEstablecimientoDestino)
                    If Not oDOEstablecimiento Is Nothing Then
                        txtRD.Text = oDOEstablecimiento.nombre
                        Me.txtRDcodigo.Text = Right("0000000000" & oDOEstablecimiento.codigo, 10)
                    End If
                Else
                    Set oDOEstablecimientoNoMinsa = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(oRsTmp1.Fields!IdEstablecimientoNoMinsaDestino)
                    If Not oDOEstablecimiento Is Nothing Then
                        txtRD.Text = oDOEstablecimientoNoMinsa.nombre
                        Me.txtRDcodigo.Text = Right("0000000000" & oDOEstablecimientoNoMinsa.codigo, 10)
                    End If
                
                End If
                Me.txtRDnumero.Text = IIf(IsNull(oRsTmp1.Fields!NroReferenciaDestino), "", oRsTmp1.Fields!NroReferenciaDestino)
                Me.btnBuscarEstablecimientoD.Enabled = False
                mo_Formulario.HabilitarDeshabilitar Me.txtRDnumero, False
           End If
           '
           ml_IdDestinoPaciente = mo_AdminAdmision.TiposDestinoAtencionDevuelveIdSis(IIf(IsNull(oRsTmp1.Fields!IdDestinoAtencion), 0, oRsTmp1.Fields!IdDestinoAtencion))
           If Val(ml_IdDestinoPaciente) > 0 Then
               mo_cmbIdDestinoAtencion.BoundText = ml_IdDestinoPaciente
           End If
           '
           mo_ReglasSISgalenhos.Sis_FuaAtencionMarcaAutomaticamente lbEsAmbulatoria, lbEsEmergencia, lbEsReferencia, _
                                                                    ml_IdTipoServicio, IIf(txtROcodigo.Text = "", False, True)
           If lbEsAmbulatoria = True Then
              chkAtencionAmbulatoria.Value = ssCBChecked
           ElseIf lbEsReferencia = True Then
              chkAtencionReferencia.Value = ssCBChecked
           ElseIf lbEsEmergencia = True Then
              chkAtencionEmergencia.Value = ssCBChecked
           End If
           '
        Else
           If mo_lnIdTablaLISTBARITEMS = sghFormatoFUA And mi_opcion <> sghAgregar Then
                
                lbEsIgualQueArSIS = True
           Else
                MsgBox "No se encontró Nro Cuenta en tabla ATENCIONES", vbInformation, Me.Caption
                CargarDatosDelPaciente = False
           End If
        End If
        oRsTmp1.Close
        '
        Set mo_DoAtencionDatosAdicionales = mo_AdminAdmision.AtencionesDatosAdicionalesSeleccionarPorId(ml_idAtencion, oConexion)
        lnIdAtencionEmergenciaOce = mo_DoAtencionDatosAdicionales.idAtencionEmeg_CE
        If mo_DoAtencionDatosAdicionales.idSiasis = 0 Then
           Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaAtencionSeleccionarPorCuenta(ml_IdCuentaAtencion)
           If oRsTmp1.RecordCount = 0 Then
                MsgBox "No se encontró Id en tabla 'SisFiliaciones'", vbInformation, Me.Caption
                Set mo_DoAtencionDatosAdicionales = Nothing
                CargarDatosDelPaciente = False
           Else
                If IsNull(oRsTmp1.Fields!idSiasis) Then
                    MsgBox "No se encontró Id en tabla 'SisFiliaciones'", vbInformation, Me.Caption
                    Set mo_DoAtencionDatosAdicionales = Nothing
                    CargarDatosDelPaciente = False
                Else
                    lcAfiliacionIdSiaSis = oRsTmp1.Fields!idSiasis
                    lcAfiliacionCodigo = oRsTmp1.Fields!codigo
                    If Not IsNull(oRsTmp1.Fields!FuaCodigoPrestacion) Then
                       ucSISfuaCodPrestacion1.CodigoPrestacion = oRsTmp1.Fields!FuaCodigoPrestacion
                    End If
                End If
           End If
           Exit Function
        End If
        lcAfiliacionIdSiaSis = mo_DoAtencionDatosAdicionales.idSiasis
        lcAfiliacionCodigo = mo_DoAtencionDatosAdicionales.SisCodigo
        If mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion <> "" Then
           ucSISfuaCodPrestacion1.CodigoPrestacion = mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion
           'If Val(mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion) > 0 Then
           '   ucSISfuaCodPrestacion1.AsignaDescripcionSegunCodigoPrestacion
           'End If
        End If
        lnIdDiagnosticoPacExtSeguro = mo_DoAtencionDatosAdicionales.referenciaOidDiagnostico
        '
        Set mo_DoAtencionDatosAdicionales = Nothing
        Set oDOEstablecimientoNoMinsa = Nothing
        Set oDOEstablecimiento = Nothing
        Set oRsTmp1 = Nothing
        Set mo_ReglasArchivoClinico = Nothing
End Function

Sub CargaValoresVacunasSp()
    If oRsVacunasSp.RecordCount > 0 Then
       oRsVacunasSp.MoveFirst
       Do While Not oRsVacunasSp.EOF
          oRsVacunasSp.Delete
          oRsVacunasSp.Update
          oRsVacunasSp.MoveNext
       Loop
    End If
    If Val(txtSPcpn.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "300"
        oRsVacunasSp.Fields!Valor = txtSPcpn.Text
        oRsVacunasSp.Update
    End If
    If Val(txtSPpeso.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "003"
        oRsVacunasSp.Fields!Valor = txtSPpeso.Text
        oRsVacunasSp.Update
    End If
    If Val(txtSPtalla.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "004"
        oRsVacunasSp.Fields!Valor = txtSPtalla.Text
        oRsVacunasSp.Update
    End If
    If Val(txtSPedadG.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "005"
        oRsVacunasSp.Fields!Valor = txtSPedadG.Text
        oRsVacunasSp.Update
    End If
    If Val(txtSPedadGrn.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "304"
        oRsVacunasSp.Fields!Valor = txtSPedadGrn.Text
        oRsVacunasSp.Update
    End If
    If Val(txtSPapgar1.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "305"
        oRsVacunasSp.Fields!Valor = txtSPapgar1.Text
        oRsVacunasSp.Update
    End If
    If Val(txtSPapgar5.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "306"
        oRsVacunasSp.Fields!Valor = txtSPapgar5.Text
        oRsVacunasSp.Update
    End If
    If Val(txtSPalturaU.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "010"
        oRsVacunasSp.Fields!Valor = txtSPalturaU.Text
        oRsVacunasSp.Update
    End If
    If chkSPPartoVertSI.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "408"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPPartoVertNO.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "408"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If chkSPCorTarCordonSI.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "409"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPCorteTarCordonNO.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "409"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If txtSPpa.Text <> sighentidades.PresionDevuelveVacia Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "901"
        oRsVacunasSp.Fields!Valor = Left(txtSPpa.Text, InStr(txtSPpa.Text, "/") - 1)
        oRsVacunasSp.Update
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "301"
        oRsVacunasSp.Fields!Valor = Trim(Mid(txtSPpa.Text, InStr(txtSPpa.Text, "/") + 1, 10))
        oRsVacunasSp.Update
    End If
    If Val(txtSPcred.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "120"
        oRsVacunasSp.Fields!Valor = txtSPcred.Text
        oRsVacunasSp.Update
    End If
    If Val(txtSPPAB.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "015"
        oRsVacunasSp.Fields!Valor = txtSPPAB.Text
        oRsVacunasSp.Update
    End If
    If Val(txtSPNFamGestante.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "404"
        oRsVacunasSp.Fields!Valor = txtSPNFamGestante.Text
        oRsVacunasSp.Update
    End If
    If Val(txtSPIMC.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "014"
        oRsVacunasSp.Fields!Valor = txtSPIMC.Text
        oRsVacunasSp.Update
    End If
    If Val(txtSPVacam.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "018"
        oRsVacunasSp.Fields!Valor = txtSPVacam.Text
        oRsVacunasSp.Update
    End If
    If Val(txtSPpuerperio.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "209"
        oRsVacunasSp.Fields!Valor = txtSPpuerperio.Text
        oRsVacunasSp.Update
    End If
    If chkSPconsejeriaNsi.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "307"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPconsejeriaNno.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "307"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If chkSPSecuelaNaceSI.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "021"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPSecuelaNaceNO.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "021"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If chkSPTamizajeSalMPAT.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "407"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPTamizajeSalMNOR.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "407"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If chkSPeedpSI.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "312"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPeedpNO.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "312"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If chkSBajoPesoSI.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "020"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSBajoPesoNO.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "020"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If chkSPRNPrematuroSI.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "019"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPRNPrematuroNO.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "019"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If chkSPEvalIntegralSI.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "401"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPEvalIntegralNO.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "401"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If chkSPConIntegralSI.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "013"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPConIntegralNO.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "013"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If Val(txtVacBcg.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "102"
        oRsVacunasSp.Fields!Valor = txtVacBcg.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacInfluenz.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "318"
        oRsVacunasSp.Fields!Valor = txtVacInfluenz.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacAntiamarilica.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "211"
        oRsVacunasSp.Fields!Valor = txtVacAntiamarilica.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacDpt.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "117"
        oRsVacunasSp.Fields!Valor = txtVacDpt.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacParotid.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "121"
        oRsVacunasSp.Fields!Valor = txtVacParotid.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacAntineumoc.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "126"
        oRsVacunasSp.Fields!Valor = txtVacAntineumoc.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacApo.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "313"
        oRsVacunasSp.Fields!Valor = txtVacApo.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacRubeola.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "122"
        oRsVacunasSp.Fields!Valor = txtVacRubeola.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacAntitetanica.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "208"
        oRsVacunasSp.Fields!Valor = txtVacAntitetanica.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacAsa.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "314"
        oRsVacunasSp.Fields!Valor = txtVacAsa.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacRotavirus.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "127"
        oRsVacunasSp.Fields!Valor = txtVacRotavirus.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacRiesgoHVB.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "406"
        oRsVacunasSp.Fields!Valor = txtVacRiesgoHVB.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacSpr.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "125"
        oRsVacunasSp.Fields!Valor = txtVacSpr.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacDt.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "007"
        oRsVacunasSp.Fields!Valor = txtVacDt.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacHVB.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "315"
        oRsVacunasSp.Fields!Valor = txtVacHVB.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacPentaval.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "124"
        oRsVacunasSp.Fields!Valor = txtVacPentaval.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacSR.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "317"
        oRsVacunasSp.Fields!Valor = txtVacSR.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacIPV.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "316"
        oRsVacunasSp.Fields!Valor = txtVacIPV.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacRiesgoHVB.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "406"
        oRsVacunasSp.Fields!Valor = txtVacRiesgoHVB.Text
        oRsVacunasSp.Update
    End If
    If Val(txtVacVPH.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "319"
        oRsVacunasSp.Fields!Valor = txtVacVPH.Text
        oRsVacunasSp.Update
    End If
End Sub

Sub InicilizarParametros()
    wxParametro205 = lcBuscaParametro.SeleccionaFilaParametro(205)
    wxParametro242 = lcBuscaParametro.SeleccionaFilaParametro(242)
    wxParametro280 = lcBuscaParametro.SeleccionaFilaParametro(280)
    wxParametro303 = lcBuscaParametro.SeleccionaFilaParametro(303)
    wxParametro304 = lcBuscaParametro.SeleccionaFilaParametro(304)
    wxParametro305 = lcBuscaParametro.SeleccionaFilaParametro(305)
    wxParametro306 = lcBuscaParametro.SeleccionaFilaParametro(306)
    wxParametro310 = lcBuscaParametro.SeleccionaFilaParametro(310)
    wxParametro320 = lcBuscaParametro.SeleccionaFilaParametro(320)
    wxParametro327 = lcBuscaParametro.SeleccionaFilaParametro(327)
    wxParametro328 = lcBuscaParametro.SeleccionaFilaParametro(328)
    wxParametro338 = lcBuscaParametro.SeleccionaFilaParametro(338)
    wxParametroJAMO = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    wxParametroSIS = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    wxParametro359 = lcBuscaParametro.SeleccionaFilaParametro(359)
    wxParametro553 = lcBuscaParametro.SeleccionaFilaParametro(553)
    wxParametro554 = lcBuscaParametro.SeleccionaFilaParametro(554)
End Sub

Function ReglasDeConsistenciasAntesDeGrabarFUA() As Boolean
    'mgaray20140926
    If mo_lnIdTablaLISTBARITEMS = sghRegistroCitaCE Or mo_lnIdTablaLISTBARITEMS = sghAdmisionEmergencia Or mo_lnIdTablaLISTBARITEMS = sghPacienteExternoConSeguro Then
       ReglasDeConsistenciasAntesDeGrabarFUA = True
       Exit Function
    End If

    ReglasDeConsistenciasAntesDeGrabarFUA = False
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oConexionExterna As New Connection
    Dim lcMensaje As String, lcCodigosCpts As String, lcCpt As String, lcCodigosFarmacia As String, lcFarmacia As String
    Dim lbAmbulatoria As Boolean, lbReferencia As Boolean, lbEmergencia As Boolean, lnPesoKg As Double
    Dim lnFor As Integer, lnPosActual As Integer, lcFiltro As String, lbEsNuevo As Boolean, lbItemsNoTienenDespacho As Boolean
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.Open wxParametroSIS
    oConexionExterna.CursorLocation = adUseClient
    lcMensaje = ""
    If Len(Trim(ucSISfuaCodPrestacion1.CodigoPrestacion)) > 0 Then
       Dim lcCeroItemsMedicamentos As Boolean, lcCeroItemsCpt As Boolean
       lcFiltro = " and rc12_idServicio='" & ucSISfuaCodPrestacion1.CodigoPrestacion & "'"
       Set oRsTmp1 = mo_ReglasSISgalenhos.Rc12SeleccionarPorFiltro(lcFiltro, oConexionExterna)
       If oRsTmp1.RecordCount > 0 Then
          lcCeroItemsMedicamentos = False
          lcCeroItemsCpt = False
          'Mínimo 1 item en FArmacia - rc12
          If oRsTmp1.Fields!rc12_MedIns = 1 Then
              If oRsFarmacia.RecordCount = 0 Then
                 lcCeroItemsMedicamentos = True
              ElseIf mo_lnIdTablaLISTBARITEMS = sghFormatoFUA Or mo_lnIdTablaLISTBARITEMS = sghAdmisionEmergencia Or mo_lnIdTablaLISTBARITEMS = sghAdmisionHospitalizacion Then
                 lbItemsNoTienenDespacho = True
                 oRsFarmacia.MoveFirst
                 Do While Not oRsFarmacia.EOF
                    If oRsFarmacia.Fields!cantidad > 0 Then
                       lbItemsNoTienenDespacho = False
                       Exit Do
                    End If
                    oRsFarmacia.MoveNext
                 Loop
                 If lbItemsNoTienenDespacho = True Then
                    lcCeroItemsMedicamentos = True
                 End If
              End If
          End If
          'Mínimo 1 item en CPT - rc12
          If oRsTmp1.Fields!rc12_apoDiag = 1 Then
              If oRsPatologia.RecordCount = 0 Then
                 lcCeroItemsCpt = True
              ElseIf mo_lnIdTablaLISTBARITEMS = sghFormatoFUA Or mo_lnIdTablaLISTBARITEMS = sghAdmisionEmergencia Or mo_lnIdTablaLISTBARITEMS = sghAdmisionHospitalizacion Then
                 lbItemsNoTienenDespacho = True
                 oRsPatologia.MoveFirst
                 Do While Not oRsPatologia.EOF
                    If oRsPatologia.Fields!ejecutado > 0 Then
                       lbItemsNoTienenDespacho = False
                       Exit Do
                    End If
                    oRsPatologia.MoveNext
                 Loop
                 If lbItemsNoTienenDespacho = True Then
                    lcCeroItemsCpt = True
                 End If
              End If
          End If
          Select Case oRsTmp1.Fields!rc12_ControlCantidad
          Case 1    'No grabar si falta Item
               'debb-21/09/2015 (inicio)
               Dim lcDx56 As String, lbEncontroDx As Boolean
               If Val(ucSISfuaCodPrestacion1.CodigoPrestacion) = 56 Then
                    'c.prestacion=consulta externa
                    lbEncontroDx = False
                    lcDx56 = "B15/J00/A09/Z35/Z10"
                    If oRsDx.RecordCount > 0 Then
                       oRsDx.MoveFirst
                       Do While Not oRsDx.EOF
                          If InStr(lcDx56, Left(UCase(Trim(oRsDx!dxIngreso)), 3)) > 0 Then
                             lbEncontroDx = True
                             Exit Do
                          End If
                          oRsDx.MoveNext
                       Loop
                    End If
                    If lbEncontroDx = False Then
                       If lcCeroItemsMedicamentos = True And lcCeroItemsCpt = True Then
                          lcMensaje = lcMensaje & "Para la PRESTACION elegida, debe registrar al menos 1 MEDICAMENTO/INSUMO o un CPT (rc12)    o   Dx= " & lcDx56 & Chr(13)
                       End If
                    End If
               Else
          
                    If lcCeroItemsMedicamentos = True Then
                       lcMensaje = lcMensaje & "Para la PRESTACION elegida, debe registrar al menos 1 MEDICAMENTO/INSUMO (rc12)" & Chr(13)
                    End If
                    If lcCeroItemsCpt = True Then
                       lcMensaje = lcMensaje & "Para la PRESTACION elegida, debe registrar al menos 1 PROCEDIMIENTO CPT (rc12)" & Chr(13)
                    End If
               End If
          Case 2    'No grabar si falta registrar Farmacia y Cie10
               If lcCeroItemsMedicamentos = True Or lcCeroItemsCpt = True Then
                    If lcCeroItemsMedicamentos = True Then
                       lcMensaje = lcMensaje & "Para el COD.PRESTACION elegido, debe registrar al menos 1 MEDICAMENTO/INSUMO (rc12)" & Chr(13)
                    End If
                    If lcCeroItemsCpt = True Then
                       lcMensaje = lcMensaje & "Para el COD.PRESTACION elegido, debe registrar al menos 1 PROCEDIMIENTO CPT (rc12)" & Chr(13)
                    End If
               End If
          Case 3    'Grabar si hay al menos un item de Farmacia o un item de Cie10
               If lcCeroItemsMedicamentos = True And lcCeroItemsCpt = True Then
                    If lcCeroItemsMedicamentos = True Then
                       lcMensaje = lcMensaje & "Para el COD.PRESTACION elegido, debe registrar al menos 1 MEDICAMENTO/INSUMO (rc12)" & Chr(13)
                    End If
                    If lcCeroItemsCpt = True Then
                       lcMensaje = lcMensaje & "Para el COD.PRESTACION elegido, debe registrar al menos 1 PROCEDIMIENTO CPT (rc12)" & Chr(13)
                    End If
               End If
          End Select
       End If
       oRsTmp1.Close
    End If
    If Len(Trim(ucSISfuaCodPrestacion1.CodigoPrestacion)) > 0 Then
       'Cpts LABORATORIO que deben tener - rc15
       lcFiltro = " and rc15_idServicio='" & ucSISfuaCodPrestacion1.CodigoPrestacion & "'"
       Set oRsTmp1 = mo_ReglasSISgalenhos.Rc15SeleccionarPorFiltro(lcFiltro, oConexionExterna)
       If oRsTmp1.RecordCount > 0 Then
          If oRsPatologia.RecordCount = 0 Then
             lcMensaje = lcMensaje & "Para el COD.PRESTACION elegido, debe registrar al menos 1 PROCEDIMIENTO CPT (rc15)" & Chr(13)
          Else
             oRsTmp1.MoveFirst
             Do While Not oRsTmp1.EOF
                oRsPatologia.MoveFirst
                oRsPatologia.Find "codigo='" & oRsTmp1.Fields!rc15_idProcedimiento & "'"
                If oRsPatologia.EOF Then
                    lcMensaje = lcMensaje & "Para el COD.PRESTACION elegido, debe registrar el PROCEDIMIENTO CPT: " & oRsTmp1.Fields!rc15_idProcedimiento & " (RC15)" & Chr(13)
                ElseIf mo_lnIdTablaLISTBARITEMS = sghFormatoFUA Or mo_lnIdTablaLISTBARITEMS = sghAdmisionEmergencia Or mo_lnIdTablaLISTBARITEMS = sghAdmisionHospitalizacion Then
                    If oRsPatologia.Fields!ejecutado = 0 Then
                       lcMensaje = lcMensaje & "Para el COD.PRESTACION elegido, debe registrar el PROCEDIMIENTO CPT: " & oRsTmp1.Fields!rc15_idProcedimiento & " (RC15)" & Chr(13)
                    End If
                End If
                oRsTmp1.MoveNext
             Loop
           End If
           
       End If
       oRsTmp1.Close
       '
       If oRsFarmacia.RecordCount > 0 Then
          lcMensaje = lcMensaje & mo_ReglasSISgalenhos.ReglasDeConsistenciaSISsoloFarmacia(oRsFarmacia, ml_IdCuentaAtencion, _
                                       ucSISfuaCodPrestacion1.CodigoPrestacion, oConexionExterna, txtFantencion.Text, _
                                       ml_IdTipoServicio, mo_lnIdTablaLISTBARITEMS, mi_opcion_fua)
       End If
       'Servicio Preventivos - Rc14
       If cmbUPSfua.Text <> "301202" Then    'cred
            lcFiltro = " and rc14_idServicio='" & ucSISfuaCodPrestacion1.CodigoPrestacion & "'"
            Set oRsTmp1 = mo_ReglasSISgalenhos.Rc14SeleccionarPorFiltro(lcFiltro, oConexionExterna)
            If oRsTmp1.RecordCount > 0 Then
               If oRsVacunasSp.RecordCount = 0 Then
                  lcMensaje = lcMensaje & "Para el COD.PRESTACION elegido, " & _
                              "se necesita registrar SERVICIOS PREVENTIVOS Y/O VACUNAS (RC14)" & Chr(13)
               Else
                     oRsTmp1.MoveFirst
                     Do While Not oRsTmp1.EOF
                        oRsVacunasSp.MoveFirst
                        oRsVacunasSp.Find "intervencionP='" & oRsTmp1.Fields!rc14_idSmi & "'"
                        If oRsVacunasSp.EOF Then
                           If IsNull(oRsTmp1.Fields!rc14_dosis) Then
                              lcMensaje = lcMensaje & "Para el COD.PRESTACION elegido, " & _
                                          "se necesita registrar: (" & oRsTmp1.Fields!rc14_idSmi & "-" & Trim(oRsTmp1.Fields!dSmi) & ") (RC14)" & Chr(13)
                           End If
                        ElseIf oRsTmp1.Fields!rc14_dosis > 0 Then
                           If oRsTmp1.Fields!rc14_dosis < oRsVacunasSp.Fields!Valor Then
                              lcMensaje = lcMensaje & "Para el COD.PRESTACION elegido, " & _
                                          "se necesita registrar: (" & oRsTmp1.Fields!rc14_idSmi & "-" & Trim(oRsTmp1.Fields!dSmi) & ") máximo: " & oRsTmp1.Fields!rc14_dosis & " (RC14)" & Chr(13)
                           End If
                        End If
                        oRsTmp1.MoveNext
                     Loop
               End If
            End If
       End If
       'Rangos Numeros de Servicio Materno Infantil - rc05
       lcFiltro = " and rc05_idServicio='" & ucSISfuaCodPrestacion1.CodigoPrestacion & "'"
       Set oRsTmp1 = mo_ReglasSISgalenhos.Rc05SeleccionarPorFiltro(lcFiltro, oConexionExterna)
       If oRsTmp1.RecordCount > 0 Then
          oRsTmp1.MoveFirst
          Do While Not oRsTmp1.EOF
             If oRsVacunasSp.RecordCount > 0 Then
                oRsVacunasSp.MoveFirst
                oRsVacunasSp.Find "intervencionP='" & oRsTmp1.Fields!rc05_idSMI & "'"
             End If
'             If oRsVacunasSp.EOF Then
'                   If oRsTmp1.Fields!rc05_sino = 1 Then
'                      lcMensaje = lcMensaje & "Para el COD.PRESTACION elegido, " & _
'                                  "debe 'Marcar SI o NO' en : (" & oRsTmp1.Fields!rc05_idSMI & "-" & Trim(oRsTmp1.Fields!dSmi) & ") (RC05)" & Chr(13)
'                   Else
'                      lcMensaje = lcMensaje & "Para el COD.PRESTACION elegido, " & _
'                                  " (" & oRsTmp1.Fields!rc05_idSMI & "-" & Trim(oRsTmp1.Fields!dSmi) & ") el Rango es : " & oRsTmp1.Fields!rc05_minimo & " , " & oRsTmp1.Fields!rc05_maximo & " (RC05)" & Chr(13)
'                   End If
'             Else
'                   If oRsTmp1.Fields!rc05_sino = 0 Then
'                        If Not (Val(oRsVacunasSp.Fields!valor) >= oRsTmp1.Fields!rc05_minimo And Val(oRsVacunasSp.Fields!valor) <= oRsTmp1.Fields!rc05_maximo) Then
'                           lcMensaje = lcMensaje & "Para el COD.PRESTACION elegido, " & _
'                                       " (" & oRsTmp1.Fields!rc05_idSMI & "-" & Trim(oRsTmp1.Fields!dSmi) & ") el Rango es : " & oRsTmp1.Fields!rc05_minimo & " , " & oRsTmp1.Fields!rc05_maximo & " (RC05)" & Chr(13)
'                        End If
'                   End If
'             End If
             If Not oRsVacunasSp.EOF Then
                   If oRsTmp1.Fields!rc05_sino = 0 Then
                        If Not (Val(oRsVacunasSp.Fields!Valor) <= oRsTmp1.Fields!rc05_maximo) Then
                           lcMensaje = lcMensaje & "Para el COD.PRESTACION elegido, " & _
                                       " (" & oRsTmp1.Fields!rc05_idSMI & "-" & Trim(oRsTmp1.Fields!dSmi) & ") el Rango es : " & oRsTmp1.Fields!rc05_minimo & " , " & oRsTmp1.Fields!rc05_maximo & " (RC05)" & Chr(13)
                        End If
                   End If
             End If
             oRsTmp1.MoveNext
          Loop
       End If
       
    End If
    'Chequea Si el Medico existe en tablas SIS
    Set oRsTmp1 = mo_ReglasSISgalenhos.a_resatencionSeleccionarPorDNI(txtMedicoDni.Text, oConexionExterna)
    If oRsTmp1.RecordCount = 0 Then
        MsgBox "El MEDICO del FUA no se encuentra en la BD del SIS, debe regularizar su inscripción en el SIS-LIMA" & Chr(10) & _
               "sino cuando el FUA llegue a SIS-LIMA será rechazado", vbInformation, Me.Caption
    End If
    '
    oConexionExterna.Close
    Set oConexionExterna = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    If lcMensaje = "" Then
        ReglasDeConsistenciasAntesDeGrabarFUA = True
    Else
        MsgBox lcMensaje, vbInformation, "Reglas de Consistencia"
    End If
End Function

Private Sub txtVacSR_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacSR
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtVacSR_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
End Sub

Private Sub txtVacVPH_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacVPH
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtVacVPH_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
End Sub

Private Sub ucSISfuaCodPrestacion1_LostFocus()
     If mo_lnIdTablaLISTBARITEMS = sghFormatoFUA And txtPaciente.Text = "" Then
        MsgBox "Debe elegir al Paciente SIS, antes del CODIGO DE PRESTACION", vbInformation, Me.Caption
        ucSISfuaCodPrestacion1.CodigoPrestacion = ""
        Exit Sub
     End If
     'debb-02/05/2016 (inicio)
     If mo_AdminServiciosComunes.FUAvalidaCodigoPrestacionSegunAdmision(mo_lnIdTablaLISTBARITEMS, ucSISfuaCodPrestacion1.CodigoPrestacion) = False Then
        ucSISfuaCodPrestacion1.CodigoPrestacion = ""
        Exit Sub
     End If
     'debb-02/05/2016 (fin)
     ReglasDeConsistenciasDespuesDeElegirCodigoPrestacion ucSISfuaCodPrestacion1.CodigoPrestacion
     ReglasDeConsistenciasDespuesDeElegirCodigoPrestacion1 ucSISfuaCodPrestacion1.CodigoPrestacion
     PermitirManipularDatosSegunSexo
     'debb-18/05/2016 (inicio)
     If txtSPIMC.Enabled = True And Val(txtSPpeso.Text) > 0 And Val(txtSPtalla.Text) > 0 Then
        txtSPIMC.Text = Round(Val(txtSPpeso.Text) / (Val(txtSPtalla.Text) * Val(txtSPtalla.Text) * 0.0001), 2)
     End If
     'debb-18/05/2016 (fin)
     HabilitaTextosParaCRED
End Sub

Sub PermitirManipularDatosJaladosDesdeGalenHos()
    btnBuscarEstablecimientoO.Enabled = False
    btnBuscarEstablecimientoD.Enabled = False
    mo_Formulario.HabilitarDeshabilitar cmbIdDestinoAtencion, False
    mo_Formulario.HabilitarDeshabilitar txtSPpeso, False
    mo_Formulario.HabilitarDeshabilitar txtSPtalla, False
    mo_Formulario.HabilitarDeshabilitar txtSPpa, False
    mo_Formulario.HabilitarDeshabilitar txtObservaciones, False
    Me.ucSISfuaCodPrestacion1.HabilitaCodigoPrestacion (False)
    FraDx.Enabled = False
    grdFarmacia.Bands(0).Columns("Recetado").Activation = ssActivationActivateNoEdit
    grdFarmacia.Bands(0).Columns("Cantidad").Activation = ssActivationActivateNoEdit
    Me.btnAddFarmacia.Visible = False
    grdPatologia.Bands(0).Columns("Indicado").Activation = ssActivationActivateNoEdit
    grdPatologia.Bands(0).Columns("Ejecutado").Activation = ssActivationActivateNoEdit
    Me.btnAddPatologia.Visible = False
    grdDx.Bands(0).Columns("DxIngresoPresuntivo").Activation = ssActivationActivateNoEdit
    grdDx.Bands(0).Columns("DxIngresoDefinitivo").Activation = ssActivationActivateNoEdit
    grdDx.Bands(0).Columns("DxIngresoRepetido").Activation = ssActivationActivateNoEdit
    grdDx.Bands(0).Columns("DxIngreso").Activation = ssActivationActivateNoEdit
    
'    fraFarmacia.Enabled = False: Me.btnAddFarmacia.ToolTipText = "Los DATOS que se jalan desde SIGH no se pueden modificar en el formato FUA"
'    FraPatologia.Enabled = False: Me.btnAddPatologia.ToolTipText = "Los DATOS que se jalan desde SIGH no se pueden modificar en el formato FUA"
    Select Case mo_lnIdTablaLISTBARITEMS
    Case sghOpcionGalenHos.sghFormatoFUA
         Select Case ml_IdTipoServicio
         Case sghConsultaExterna
              'S (Tiene PC en Consultorios externos donde usan GalenHos, se usa en el registro del FUA)
              If lcElServicioUsaGalenHos = "S" Then
                     
              Else
                    btnBuscarEstablecimientoO.Enabled = True
                    btnBuscarEstablecimientoD.Enabled = True
                    mo_Formulario.HabilitarDeshabilitar cmbIdDestinoAtencion, True
                    mo_Formulario.HabilitarDeshabilitar txtSPpeso, True
                    mo_Formulario.HabilitarDeshabilitar txtSPtalla, True
                    mo_Formulario.HabilitarDeshabilitar txtSPpa, True
                    mo_Formulario.HabilitarDeshabilitar txtObservaciones, True
                    grdDx.Bands(0).Columns("DxIngresoPresuntivo").Activation = ssActivationAllowEdit
                    grdDx.Bands(0).Columns("DxIngresoDefinitivo").Activation = ssActivationAllowEdit
                    grdDx.Bands(0).Columns("DxIngresoRepetido").Activation = ssActivationAllowEdit
                    grdDx.Bands(0).Columns("DxIngreso").Activation = ssActivationAllowEdit
                    FraDx.Enabled = True
                    Me.ucSISfuaCodPrestacion1.HabilitaCodigoPrestacion (True)
              End If
         End Select
    Case sghOpcionGalenHos.sghRegistroAtencionCE
         Select Case ml_IdTipoServicio
         Case sghConsultaExterna
              If lcElServicioUsaGalenHos = "S" Then
                 Me.ucSISfuaCodPrestacion1.HabilitaCodigoPrestacion (True)
              End If
             ' If txtRO.Text <> "" Then
             '    mo_Formulario.HabilitarDeshabilitar txtRONumero, True
             ' End If
         End Select
    End Select
    'S (Tiene PC en Laboratorio/Imágenes donde usan GalenHos, se usa en el registro del FUA)
    If wxParametro327 <> "S" And mo_lnIdTablaLISTBARITEMS = sghFormatoFUA Then
        grdPatologia.Bands(0).Columns("Indicado").Activation = ssActivationAllowEdit
        grdPatologia.Bands(0).Columns("Ejecutado").Activation = ssActivationAllowEdit
        Me.btnAddPatologia.Visible = True
'       FraPatologia.Enabled = True: Me.btnAddPatologia.ToolTipText = ""
    End If
    'S (Tiene PC en Farmacia donde usan GalenHos, se usa en el registro del FUA)
    If wxParametro328 <> "S" And mo_lnIdTablaLISTBARITEMS = sghFormatoFUA Then
        grdFarmacia.Bands(0).Columns("Recetado").Activation = ssActivationAllowEdit
        grdFarmacia.Bands(0).Columns("Cantidad").Activation = ssActivationAllowEdit
        Me.btnAddFarmacia.Visible = True
       'Me.fraFarmacia.Enabled = True: fraFarmacia.Enabled = False: Me.btnAddFarmacia.ToolTipText = ""
    End If
    'Sexo
    PermitirManipularDatosSegunSexo
    'Hospitalizacion/Emergencia
    If ml_IdTipoServicio <> sghConsultaExterna Then
       mo_Formulario.HabilitarDeshabilitar FarLugarAtencion, False
       mo_Formulario.HabilitarDeshabilitar FraPersonal, False
       mo_Formulario.HabilitarDeshabilitar txtObservaciones, True
    End If
    '
    If ucSISfuaCodPrestacion1.CodigoPrestacion = "" Then
       ucSISfuaCodPrestacion1.HabilitaCodigoPrestacion True
    End If
End Sub

Sub PermitirManipularDatosSegunSexo()
    If txtSexo.Text = lcMasculino Then
       mo_Formulario.HabilitarDeshabilitar fraGestantePuerpera, False
       mo_Formulario.HabilitarDeshabilitar txtFparto, False
       mo_Formulario.HabilitarDeshabilitar txtSPedadG, False
       mo_Formulario.HabilitarDeshabilitar txtSPalturaU, False
    End If
End Sub



Private Sub ucSISfuaCodPrestacion1_SePresionoTeclaEspecial(KeyCode As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub ReglasDeConsistenciasDespuesDeElegirCodigoPrestacion1(lcCodigoPrestacion As String)
    Dim oRsTmp1 As New Recordset
    Dim oConexionExterna As New Connection
    Dim lcFiltro As String
    mo_Formulario.HabilitarDeshabilitar txtSPcpn, False
    mo_Formulario.HabilitarDeshabilitar fraConsejNutricional, False
    mo_Formulario.HabilitarDeshabilitar txtSPedadG, False
    mo_Formulario.HabilitarDeshabilitar txtSPalturaU, False
    mo_Formulario.HabilitarDeshabilitar txtSPedadGrn, False
    mo_Formulario.HabilitarDeshabilitar txtSPapgar1, False
    mo_Formulario.HabilitarDeshabilitar txtSPapgar5, False
    mo_Formulario.HabilitarDeshabilitar fraSecuelaNacer, False
    mo_Formulario.HabilitarDeshabilitar fraBajoPesoNacer, False
    mo_Formulario.HabilitarDeshabilitar txtSPcred, False
    mo_Formulario.HabilitarDeshabilitar fraEEDP, False
    mo_Formulario.HabilitarDeshabilitar frPartoVertical, False
    mo_Formulario.HabilitarDeshabilitar fraCorTardio, False
    mo_Formulario.HabilitarDeshabilitar fraTamizajeSaludM, False
    mo_Formulario.HabilitarDeshabilitar fraRnPrematuro, False
    mo_Formulario.HabilitarDeshabilitar fraEvalIntegral, False
    mo_Formulario.HabilitarDeshabilitar txtSPpuerperio, False
    mo_Formulario.HabilitarDeshabilitar fraConIntegral, False
    mo_Formulario.HabilitarDeshabilitar txtSPPAB, False
    mo_Formulario.HabilitarDeshabilitar txtSPNFamGestante, False
    mo_Formulario.HabilitarDeshabilitar txtSPIMC, False
    mo_Formulario.HabilitarDeshabilitar txtVacBcg, False
    mo_Formulario.HabilitarDeshabilitar txtVacInfluenz, False
    mo_Formulario.HabilitarDeshabilitar txtVacAntiamarilica, False
    mo_Formulario.HabilitarDeshabilitar txtVacDpt, False
    mo_Formulario.HabilitarDeshabilitar txtVacParotid, False
    mo_Formulario.HabilitarDeshabilitar txtVacAntineumoc, False
    mo_Formulario.HabilitarDeshabilitar txtVacApo, False
    mo_Formulario.HabilitarDeshabilitar txtVacRubeola, False
    mo_Formulario.HabilitarDeshabilitar txtVacAntitetanica, False
    mo_Formulario.HabilitarDeshabilitar txtVacAsa, False
    mo_Formulario.HabilitarDeshabilitar txtVacRotavirus, False
    mo_Formulario.HabilitarDeshabilitar txtVacSpr, False
    mo_Formulario.HabilitarDeshabilitar txtVacDt, False
    mo_Formulario.HabilitarDeshabilitar txtVacHVB, False
    mo_Formulario.HabilitarDeshabilitar txtVacPentaval, False
    mo_Formulario.HabilitarDeshabilitar txtVacSR, False
    mo_Formulario.HabilitarDeshabilitar txtVacIPV, False
    mo_Formulario.HabilitarDeshabilitar txtVacRiesgoHVB, False
    mo_Formulario.HabilitarDeshabilitar txtVacVPH, False
                  
    If Len(Trim(lcCodigoPrestacion)) > 0 Then
       'Habilitar solo los que sigan la RC05
       oConexionExterna.CommandTimeout = 300
       oConexionExterna.Open wxParametroSIS
       oConexionExterna.CursorLocation = adUseClient
       lcFiltro = " and rc05_idServicio='" & lcCodigoPrestacion & "'"
       Set oRsTmp1 = mo_ReglasSISgalenhos.Rc05SeleccionarPorFiltro(lcFiltro, oConexionExterna)
       If oRsTmp1.RecordCount > 0 Then
          oRsTmp1.MoveFirst
          Do While Not oRsTmp1.EOF
             Select Case oRsTmp1.Fields!rc05_idSMI
             Case "300"
                  mo_Formulario.HabilitarDeshabilitar txtSPcpn, True
             Case "005"
                  mo_Formulario.HabilitarDeshabilitar txtSPedadG, True
             Case "304"
                  mo_Formulario.HabilitarDeshabilitar txtSPedadGrn, True
             Case "305"
                  mo_Formulario.HabilitarDeshabilitar txtSPapgar1, True
             Case "306"
                  mo_Formulario.HabilitarDeshabilitar txtSPapgar5, True
             Case "010"
                  mo_Formulario.HabilitarDeshabilitar txtSPalturaU, True
             Case "120"
                  mo_Formulario.HabilitarDeshabilitar txtSPcred, True
             Case "209"
                  mo_Formulario.HabilitarDeshabilitar txtSPpuerperio, True
             Case "307"
                  mo_Formulario.HabilitarDeshabilitar fraConsejNutricional, True
             Case "408"
                  mo_Formulario.HabilitarDeshabilitar frPartoVertical, True
             Case "409"
                  mo_Formulario.HabilitarDeshabilitar fraCorTardio, True
             Case "021"
                  mo_Formulario.HabilitarDeshabilitar fraSecuelaNacer, True
             Case "407"
                  mo_Formulario.HabilitarDeshabilitar fraTamizajeSaludM, True
             Case "312"
                  mo_Formulario.HabilitarDeshabilitar fraEEDP, True
             Case "020"
                  mo_Formulario.HabilitarDeshabilitar fraBajoPesoNacer, True
             Case "019"
                  mo_Formulario.HabilitarDeshabilitar fraRnPrematuro, True
             Case "401"
                  mo_Formulario.HabilitarDeshabilitar fraEvalIntegral, True
             Case "013"
                  mo_Formulario.HabilitarDeshabilitar fraConIntegral, True
             Case "015"
                  mo_Formulario.HabilitarDeshabilitar txtSPPAB, True
             Case "404"
                  mo_Formulario.HabilitarDeshabilitar txtSPNFamGestante, True
             Case "014"
                  mo_Formulario.HabilitarDeshabilitar txtSPIMC, True
             Case "102"
                  mo_Formulario.HabilitarDeshabilitar txtVacBcg, True
             Case "318"
                  mo_Formulario.HabilitarDeshabilitar txtVacInfluenz, True
             Case "211"
                  mo_Formulario.HabilitarDeshabilitar txtVacAntiamarilica, True
             Case "117"
                  mo_Formulario.HabilitarDeshabilitar txtVacDpt, True
             Case "121"
                  mo_Formulario.HabilitarDeshabilitar txtVacParotid, True
             Case "126"
                  mo_Formulario.HabilitarDeshabilitar txtVacAntineumoc, True
             Case "313"
                  mo_Formulario.HabilitarDeshabilitar txtVacApo, True
             Case "122"
                  mo_Formulario.HabilitarDeshabilitar txtVacRubeola, True
             Case "208"
                  mo_Formulario.HabilitarDeshabilitar txtVacAntitetanica, True
             Case "314"
                  mo_Formulario.HabilitarDeshabilitar txtVacAsa, True
             Case "127"
                  mo_Formulario.HabilitarDeshabilitar txtVacRotavirus, True
             Case "406"
                  mo_Formulario.HabilitarDeshabilitar txtVacRiesgoHVB, True
             Case "125"
                  mo_Formulario.HabilitarDeshabilitar txtVacSpr, True
             Case "007"
                  mo_Formulario.HabilitarDeshabilitar txtVacDt, True
             Case "315"
                  mo_Formulario.HabilitarDeshabilitar txtVacHVB, True
             Case "124"
                  mo_Formulario.HabilitarDeshabilitar txtVacPentaval, True
             Case "317"
                  mo_Formulario.HabilitarDeshabilitar txtVacSR, True
             Case "316"
                  mo_Formulario.HabilitarDeshabilitar txtVacIPV, True
             Case "319"
                  mo_Formulario.HabilitarDeshabilitar txtVacVPH, True
             End Select
             oRsTmp1.MoveNext
          Loop
       End If
       oRsTmp1.Close
       oConexionExterna.Close
    End If
    Set oRsTmp1 = Nothing
    Set oConexionExterna = Nothing
End Sub


Sub ReglasDeConsistenciasDespuesDeElegirCodigoPrestacion(lcCodigoPrestacion As String)
    If Len(Trim(lcCodigoPrestacion)) = 0 Then
       Exit Sub
    End If
    Dim oRsTmp1 As New Recordset
    Dim oConexionExterna As New Connection
    Dim lcMensaje As String, lcCodigosCpts As String, lcCpt As String, lcCodigosFarmacia As String, lcFarmacia As String
    Dim lbAmbulatoria As Boolean, lbReferencia As Boolean, lbEmergencia As Boolean, lnPesoKg As Double
    Dim lnFor As Integer, lnPosActual As Integer, lbContinuar As Boolean, lcFiltro As String
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.Open wxParametroSIS
    oConexionExterna.CursorLocation = adUseClient
    lbContinuar = True
    'Maximos FUA por dia,mes,año - rc13
    lcFiltro = " and rc13_idServicio='" & lcCodigoPrestacion & "'"
    Set oRsTmp1 = mo_ReglasSISgalenhos.Rc13SeleccionarPorFiltro(lcFiltro, oConexionExterna)
    If oRsTmp1.RecordCount > 0 Then
       
       If Not (IsNull(oRsTmp1.Fields!rc13_TopeDia) And IsNull(oRsTmp1.Fields!rc13_topeMes) And IsNull(oRsTmp1.Fields!rc13_topeAnio)) Then
            Dim lnFuasDelDia As Integer, lnFuasDelMes As Integer, lnFuasDelAnio As Integer
            mo_ReglasSISgalenhos.Sis_FuasRegistradasXpacienteSegunCodigoPrestacion CDate(txtFantencion.Text), ml_IdCuentaAtencion, _
                                 lcCodigoPrestacion, txtNhistoriaClinica.Text, lnFuasDelDia, lnFuasDelMes, lnFuasDelAnio
            If Not (oRsTmp1.Fields!rc13_TopeDia >= lnFuasDelDia) And (oRsTmp1.Fields!rc13_topeMes >= lnFuasDelMes) And _
                   (oRsTmp1.Fields!rc13_topeAnio >= lnFuasDelAnio) Then
               MsgBox "Para la PRESTACION elegida, excede en N°Fuas (rc13), (Fuas registradas Día: " & lnFuasDelDia & _
                      ", Mes:" & lnFuasDelMes & ", Año: " & lnFuasDelAnio & ") (Topes: " & oRsTmp1.Fields!rc13_TopeDia & _
                      "," & oRsTmp1.Fields!rc13_topeMes & "," & oRsTmp1.Fields!rc13_topeAnio & ")", vbInformation, Me.Caption
               ucSISfuaCodPrestacion1.CodigoPrestacion = ""
               Me.ucSISfuaCodPrestacion1.Prestacion = ""
               lbContinuar = False
            End If
       End If
    End If
    oRsTmp1.Close
    'Destino del Asegurado - rc4
    If lbContinuar = True Then
       lcFiltro = " and rc04_idServicio='" & lcCodigoPrestacion & "'"
       Set mo_cmbIdDestinoAtencion.RowSource = mo_ReglasSISgalenhos.Rc04SeleccionarPorFiltro(lcFiltro, oConexionExterna)
    End If
    If Val(ml_IdDestinoPaciente) > 0 Then
       mo_cmbIdDestinoAtencion.BoundText = ml_IdDestinoPaciente
    Else
       Set oRsTmp1 = mo_cmbIdDestinoAtencion.RowSource
       If oRsTmp1.RecordCount = 1 Then
          oRsTmp1.MoveFirst
          mo_cmbIdDestinoAtencion.BoundText = oRsTmp1.Fields!des_IdDestinoAsegurado
       End If
    End If
    'Atencion - rc16
    If lbContinuar = True Then
       lbAmbulatoria = False
       lbReferencia = False
       lbEmergencia = False
       lcFiltro = " where rc16_idServicio='" & lcCodigoPrestacion & "' and rc16_nivel='" & lcNivelEstablecimiento & "'"
       Set oRsTmp1 = mo_ReglasSISgalenhos.Rc16SeleccionarPorFiltro(lcFiltro, oConexionExterna)
       If oRsTmp1.RecordCount > 0 Then
          oRsTmp1.MoveFirst
          Do While Not oRsTmp1.EOF
             Select Case oRsTmp1.Fields!rc16_idTipoAtencion
             Case 1
                  lbAmbulatoria = True
             Case 2
                  lbReferencia = True
             Case 3
                  lbEmergencia = True
             End Select
             oRsTmp1.MoveNext
          Loop
          chkAtencionAmbulatoria.Enabled = lbAmbulatoria: If lbAmbulatoria = False Then chkAtencionAmbulatoria.Value = ssCBUnchecked
          chkAtencionReferencia.Enabled = lbReferencia: If lbReferencia = False Then chkAtencionReferencia.Value = ssCBUnchecked
          chkAtencionEmergencia.Enabled = lbEmergencia: If lbEmergencia = False Then chkAtencionEmergencia.Value = ssCBUnchecked
       End If
       oRsTmp1.Close
    End If
    'concepto Prestacional - rc3
    If lbContinuar = True Then
       lcFiltro = " AND rc03_IdServicio = '" & lcCodigoPrestacion & "'"
       Set mo_cmbConceptoP.RowSource = mo_ReglasSISgalenhos.Rc03SeleccionarPorFiltro(lcFiltro, oConexionExterna)
       If Val(ml_IdConceptoPrestacional) > 0 Then
          mo_cmbConceptoP.BoundText = ml_IdConceptoPrestacional
       Else
          Set oRsTmp1 = mo_cmbConceptoP.RowSource
          If oRsTmp1.RecordCount = 1 Then
             oRsTmp1.MoveFirst
             mo_cmbConceptoP.BoundText = oRsTmp1.Fields!mod_idModalidad
          End If
       End If
    End If
    'rc1
    If lbContinuar = True Then
       lcFiltro = " and rc01_idServicio='" & lcCodigoPrestacion & "' and ('" & ml_edad_En_YYYYMMDD & _
                  "'>=rc01_edadMin  and '" & ml_edad_En_YYYYMMDD & "'<= rc01_edadMax)"
       Set oRsTmp1 = mo_ReglasSISgalenhos.Rc01SeleccionarPorFiltro(lcFiltro, oConexionExterna)
       If oRsTmp1.RecordCount > 0 Then
          'Gestante/puerpera/ninguno - rc1
          Select Case oRsTmp1.Fields!rc01_idCondicion
          Case 0     'ninguna
               chkGestante.Enabled = False: chkGestante.Value = ssCBUnchecked
               chkPuerpera.Enabled = False: chkPuerpera.Value = ssCBUnchecked
          Case 1     'solo Gestante
               chkGestante.Enabled = True
               chkPuerpera.Enabled = False: chkPuerpera.Value = ssCBUnchecked
          Case 2     'solo puerpera
               chkGestante.Enabled = False: chkGestante.Value = ssCBUnchecked
               chkPuerpera.Enabled = True
          End Select
       End If
       
       oRsTmp1.Close
    End If
    '
    oConexionExterna.Close
    Set oConexionExterna = Nothing
    Set oRsTmp1 = Nothing
End Sub

Sub AsignaTipoAfiliacion(lnAfiliacion As Integer)
    Select Case lnAfiliacion
    Case 0
    End Select
End Sub

Function EsUnFuaEmitidoEnVentanillaCitas() As Boolean
    EsUnFuaEmitidoEnVentanillaCitas = False
    If mo_lnIdTablaLISTBARITEMS = sghFormatoFUA And ml_IdTipoServicio = sghConsultaExterna And _
                                                             lcElServicioUsaGalenHos <> "S" Then
       EsUnFuaEmitidoEnVentanillaCitas = True
    End If
End Function

'debb-06/08/2015
Function CargaVacunaYsp(lcIntervencionesPreventivas As String, lcValor As String, ByRef lcSistolica As String) As Boolean
              CargaVacunaYsp = True
              Select Case lcIntervencionesPreventivas
              Case "300"
                 txtSPcpn.Text = lcValor
              Case "003"
                 txtSPpeso.Text = lcValor
              Case "004"
                 txtSPtalla.Text = lcValor
              Case "005"
                 txtSPedadG.Text = lcValor
              Case "304"
                 txtSPedadGrn.Text = lcValor
              Case "305"
                 txtSPapgar1.Text = lcValor
              Case "306"
                 txtSPapgar5.Text = lcValor
              Case "010"
                 txtSPalturaU.Text = lcValor
              Case "901"
                 lcSistolica = Trim(lcValor)
              Case "301"
                 txtSPpa.Text = sighentidades.PresionJuntaSistolicaDiastolica(lcSistolica, lcValor)
              Case "120"
                 txtSPcred.Text = lcValor
              Case "015"
                 txtSPPAB.Text = lcValor
              Case "209"
                 txtSPpuerperio.Text = lcValor
              Case "018"
                 txtSPVacam.Text = lcValor
              Case "408"
                 If Val(lcValor) = 1 Then
                    chkSPPartoVertSI.Value = 1
                 Else
                    chkSPPartoVertNO.Value = 1
                 End If
              Case "409"
                 If Val(lcValor) = 1 Then
                    chkSPCorTarCordonSI.Value = 1
                 Else
                    chkSPCorteTarCordonNO.Value = 1
                 End If
              Case "307"
                 If Val(lcValor) = 1 Then
                    chkSPconsejeriaNsi.Value = 1
                 Else
                    chkSPconsejeriaNno.Value = 1
                 End If
              Case "021"
                 If Val(lcValor) = 1 Then
                    chkSPSecuelaNaceSI.Value = 1
                 Else
                    chkSPSecuelaNaceNO.Value = 1
                 End If
              Case "407"
                 If Val(lcValor) = 1 Then
                    chkSPTamizajeSalMPAT.Value = 1
                 Else
                    chkSPTamizajeSalMNOR.Value = 1
                 End If
              Case "312"
                 If Val(lcValor) = 1 Then
                    chkSPeedpSI.Value = 1
                 Else
                    chkSPeedpNO.Value = 1
                 End If
              Case "020"
                 If Val(lcValor) = 1 Then
                    chkSBajoPesoSI.Value = 1
                 Else
                    chkSBajoPesoNO.Value = 1
                 End If
              Case "019"
                 If Val(lcValor) = 1 Then
                    chkSPRNPrematuroSI.Value = 1
                 Else
                    chkSPRNPrematuroNO.Value = 1
                 End If
              Case "401"
                 If Val(lcValor) = 1 Then
                    chkSPEvalIntegralSI.Value = 1
                 Else
                    chkSPEvalIntegralNO.Value = 1
                 End If
              Case "013"
                 If Val(lcValor) = 1 Then
                    chkSPConIntegralSI.Value = 1
                 Else
                    chkSPConIntegralNO.Value = 1
                 End If
              Case "404"
                 txtSPNFamGestante.Text = lcValor
              Case "014"
                 txtSPIMC.Text = lcValor
              Case "102"
                 txtVacBcg.Text = lcValor
              Case "318"
                 txtVacInfluenz.Text = lcValor
              Case "211"
                 txtVacAntiamarilica.Text = lcValor
              Case "117"
                 txtVacDpt.Text = lcValor
              Case "121"
                 txtVacParotid.Text = lcValor
              Case "126"
                 txtVacAntineumoc.Text = lcValor
              Case "313"
                 txtVacApo.Text = lcValor
              Case "122"
                 txtVacRubeola.Text = lcValor
              Case "208"
                 txtVacAntitetanica.Text = lcValor
              Case "314"
                 txtVacAsa.Text = lcValor
              Case "127"
                 txtVacRotavirus.Text = lcValor
              Case "406"
                 txtVacRiesgoHVB.Text = lcValor
              Case "125"
                 txtVacSpr.Text = lcValor
              Case "007"
                 txtVacDt.Text = lcValor
              Case "315"
                 txtVacHVB.Text = lcValor
              Case "124"
                 txtVacPentaval.Text = lcValor
              Case "317"
                 txtVacSR.Text = lcValor
              Case "316"
                 txtVacIPV.Text = lcValor
              Case "319"
                 txtVacVPH.Text = lcValor
              Case Else
                 CargaVacunaYsp = False
              End Select
End Function



Sub CargaDatosDeNacimiento(oConexion As Connection, lbDesdeGalenHos As Boolean)
    On Error GoTo ErrN
    Dim oRsTmp1 As New Recordset
    With oRsNacimientos
          .Fields.Append "documento", adVarChar, 22, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With

   ' If lbDesdeGalenHos = True Then
        Set oRsTmp1 = mo_AdminAdmision.AtencionesNacimientosXidAtencion(ml_idAtencion, oConexion)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
                If Not (IsNull(oRsTmp1!IdDocIdentidad) And IsNull(oRsTmp1!docIdentidad)) Then
                   oRsNacimientos.AddNew
                   oRsNacimientos.Fields!documento = Trim(Str(oRsTmp1!IdDocIdentidad)) & "-" & Trim(oRsTmp1!docIdentidad)
                   oRsNacimientos.Update
                End If
                oRsTmp1.MoveNext
           Loop
        End If
        oRsTmp1.Close
  '  Else
         'falta desarrollar
   ' End If
    Set grdRN.DataSource = oRsNacimientos
    mo_Apariencia.ConfigurarFilasBiColores grdRN, sighentidades.GrillaConFilasBicolor
ErrN:
End Sub

Sub CPTesPAQUETEdisminuyeMedicamentosInsumos()


    On Error GoTo ErrCptEsPte
    Dim oRsTmpCab As New Recordset
    Dim oRsTmpDet As New Recordset
    Dim lbTieneUnCaso As Boolean
    If wxParametro554 = "S" And oRsPatologia.RecordCount > 0 And oRsFarmacia.RecordCount > 0 Then
       lbTieneUnCaso = False
       Set oRsTmpCab = mo_ReglasFacturacion.FactCatalogoPaqueteSeleccionarPorFiltro("not cpt is null")
       If oRsTmpCab.RecordCount > 0 Then
          oRsPatologia.MoveFirst
          Do While Not oRsPatologia.EOF
             oRsTmpCab.MoveFirst
             oRsTmpCab.Find "cpt='" & Trim(oRsPatologia!codigo) & "'"
             If Not oRsTmpCab.EOF Then
                Set oRsTmpDet = mo_ReglasFacturacion.FacturacionCatalogoPaqueteFarmSeleccionarXid(oRsTmpCab!idFactPaquete)
                If oRsTmpDet.RecordCount > 0 Then
                   oRsTmpDet.MoveFirst
                   Do While Not oRsTmpDet.EOF
                      oRsFarmacia.MoveFirst
                      oRsFarmacia.Find "codigo='" & Trim(oRsTmpDet!codigo) & "'"
                      If Not oRsFarmacia.EOF Then
                         If (oRsFarmacia!cantidad - oRsTmpDet!cantidad) <= 0 Then
                            oRsFarmacia.Delete
                         Else
                            oRsFarmacia!cantidad = oRsFarmacia!cantidad - oRsTmpDet!cantidad
                            oRsFarmacia!recetado = oRsFarmacia!recetado - oRsTmpDet!cantidad
                         End If
                         oRsFarmacia.Update
                         lbTieneUnCaso = True
                      End If
                      oRsTmpDet.MoveNext
                   Loop
                End If
                oRsTmpDet.Close
             End If
             oRsPatologia.MoveNext
          Loop
       End If
       oRsTmpCab.Close
       If lbTieneUnCaso = True Then
          lblCtaEmergencia.Caption = lblCtaEmergencia.Caption & _
                                     " <>  existen CPT que disminuyen cantidades en Medicamentos/Insumos"
       End If
    End If
ErrCptEsPte:
    Set oRsTmpCab = Nothing
    Set oRsTmpDet = Nothing



End Sub

Sub ChequeaQueRecetadoNoSeaMenorAdespachado()
       If oRsPatologia.RecordCount > 0 Then
          oRsPatologia.MoveFirst
          Do While Not oRsPatologia.EOF
             If oRsPatologia!indicado < oRsPatologia!ejecutado Then
                oRsPatologia!indicado = oRsPatologia!ejecutado
                oRsPatologia.Update
             End If
             oRsPatologia.MoveNext
          Loop
       End If
       If oRsFarmacia.RecordCount > 0 Then
          oRsFarmacia.MoveFirst
          Do While Not oRsFarmacia.EOF
             If oRsFarmacia!recetado < oRsFarmacia!cantidad Then
                oRsFarmacia!recetado = oRsFarmacia!cantidad
                oRsFarmacia.Update
             End If
             oRsFarmacia.MoveNext
          Loop
       End If
       
End Sub


'- faltan "reglas de consistencia" como procedimiento: rc09,rc12,rc16
'- no se usarán las siguientes "Reglas de Consistencia": rc02, rc07, rc08, rc10, rc11, rc17
'- Algunos Codigos diferentes: Hemograma (85027 en GalenHos)(85031 en SIS) , Examen de orina (81000 en galenhos)(81005 en SIS)
'- Para cod.prestacion: 056-consulta externa (vs rc015), pide que se ingrese la vacuna: 123-vacuna anti
'  haemuphilus influenzae,  sin embargo en el FORMATO FUA FISICO no existe, pero si en las reglas
'  de consistencia SIS actual. Lo mismo para vacuna: 316(vacuna antipolio inyectable IPV),
'  318 (vacuna SR), 319(vph-virus de papiloma humano)
'- Para cod.prestacion: 056- consulta externa (vs rc015), es muy generico, por ejm un varon que tenga DOLOR EN EL ESTOMAGO
'  le obligará a registrar "altura uterina (010), edad gestacional (005)". Debería haber filtros de "Cod.Prestacion"
'  vs "Especialidad del consultorio" para que filtre solo "codigos de prestaciones" del Consutorio Externo.
'- Si se registró exámen en Laboratorio o Imágen lo jalará al SIS. Sin comprobar que se registró RESULTADOS.
'- Caso: se RECETO 1 Item para Farmacia y 1 item para Servicio, se graba e imprime FUA desde Atencion CE ,solo se despacho
'  Servicio, cuando se quiere actualizar algunos datos de FUA no deja grabar porque exige 1 medicamento en Farmacia, para
'  cod.prestacion=056-consulta externa
'





