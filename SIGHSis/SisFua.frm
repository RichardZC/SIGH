VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form SisFua 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12585
   Icon            =   "SisFua.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SIGHSis.ucSISfuaCodPrestacion ucSISfuaCodPrestacion1 
      Height          =   285
      Left            =   5760
      TabIndex        =   233
      Top             =   60
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   503
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
      Left            =   1890
      TabIndex        =   164
      ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
      Top             =   0
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
      Left            =   3030
      MaxLength       =   8
      TabIndex        =   0
      Top             =   0
      Width           =   2535
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
      Left            =   2550
      TabIndex        =   163
      Top             =   0
      Width           =   465
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   30
      TabIndex        =   36
      Top             =   8130
      Width           =   12510
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
         Height          =   700
         Left            =   90
         Picture         =   "SisFua.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "SisFua.frx":11A3
         DownPicture     =   "SisFua.frx":1667
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
         Left            =   6397
         Picture         =   "SisFua.frx":1B53
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "SisFua.frx":203F
         DownPicture     =   "SisFua.frx":249F
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
         Left            =   4852
         Picture         =   "SisFua.frx":2914
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   225
         Width           =   1365
      End
   End
   Begin TabDlg.SSTab TabFua 
      Height          =   7695
      Left            =   0
      TabIndex        =   37
      Top             =   390
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   13573
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
      TabPicture(0)   =   "SisFua.frx":2D89
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "btnBuscarPaciente"
      Tab(0).Control(1)=   "Frame18"
      Tab(0).Control(2)=   "Frame17"
      Tab(0).Control(3)=   "Frame16"
      Tab(0).Control(4)=   "Frame15"
      Tab(0).Control(5)=   "Frame13"
      Tab(0).Control(6)=   "fraReconsideracion"
      Tab(0).Control(7)=   "Frame7"
      Tab(0).Control(8)=   "Frame3"
      Tab(0).Control(9)=   "Frame1"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Serv.Prev/Vacunas/Dx  (F4)"
      TabPicture(1)   =   "SisFua.frx":2DA5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtMedicoEspecialidad"
      Tab(1).Control(1)=   "txtMedicoDni"
      Tab(1).Control(2)=   "txtMedicoColegiatura"
      Tab(1).Control(3)=   "txtMedico"
      Tab(1).Control(4)=   "FraDx"
      Tab(1).Control(5)=   "Frame4"
      Tab(1).Control(6)=   "Frame5"
      Tab(1).Control(7)=   "btnRefrescar"
      Tab(1).Control(8)=   "Label54"
      Tab(1).Control(9)=   "Label53"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Medicamentos/Cpt  (F5)"
      TabPicture(2)   =   "SisFua.frx":2DC1
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame19"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "FraPatologia"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "FraFarmacia"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.CommandButton btnBuscarPaciente 
         Caption         =   "..."
         Height          =   315
         Left            =   -66120
         TabIndex        =   232
         ToolTipText     =   "Busca por Apellidos y Nombres"
         Top             =   2040
         Width           =   315
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
         Height          =   3405
         Left            =   150
         TabIndex        =   153
         Top             =   360
         Width           =   12135
         Begin VB.CommandButton btnAddFarmacia 
            DisabledPicture =   "SisFua.frx":2DDD
            DownPicture     =   "SisFua.frx":31C6
            Height          =   345
            Left            =   11580
            Picture         =   "SisFua.frx":35D2
            Style           =   1  'Graphical
            TabIndex        =   154
            Top             =   300
            Width           =   405
         End
         Begin UltraGrid.SSUltraGrid grdFarmacia 
            Height          =   2835
            Left            =   120
            TabIndex        =   155
            Top             =   300
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   5001
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "F7 = Busca MEDICAMENTO/INSUMO "
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
            TabIndex        =   160
            Top             =   3150
            Width           =   3000
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
         Height          =   2805
         Left            =   150
         TabIndex        =   150
         Top             =   3810
         Width           =   12165
         Begin VB.CommandButton btnAddPatologia 
            DisabledPicture =   "SisFua.frx":39DE
            DownPicture     =   "SisFua.frx":3DC7
            Height          =   345
            Left            =   11670
            Picture         =   "SisFua.frx":41D3
            Style           =   1  'Graphical
            TabIndex        =   152
            Top             =   240
            Width           =   405
         End
         Begin UltraGrid.SSUltraGrid grdPatologia 
            Height          =   2265
            Left            =   60
            TabIndex        =   151
            Top             =   240
            Width           =   11565
            _ExtentX        =   20399
            _ExtentY        =   3995
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
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "F8 = Busca procedimiento CPT"
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
            Left            =   90
            TabIndex        =   161
            Top             =   2550
            Width           =   2550
         End
      End
      Begin VB.TextBox txtMedicoEspecialidad 
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
         Left            =   -63120
         TabIndex        =   149
         Top             =   7320
         Width           =   465
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
         Left            =   -72420
         TabIndex        =   146
         ToolTipText     =   "DNI del médico"
         Top             =   7320
         Width           =   1455
      End
      Begin VB.TextBox txtMedicoColegiatura 
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
         Left            =   -65730
         TabIndex        =   145
         Top             =   7320
         Width           =   1035
      End
      Begin VB.TextBox txtMedico 
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
         Left            =   -70950
         TabIndex        =   144
         Top             =   7320
         Width           =   5205
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
         Height          =   975
         Left            =   150
         TabIndex        =   142
         Top             =   6630
         Width           =   12315
         Begin VB.TextBox txtObservaciones 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   143
            Top             =   270
            Width           =   11535
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
         Height          =   3075
         Left            =   -74850
         TabIndex        =   86
         Top             =   4230
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
            TabIndex        =   195
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
            TabIndex        =   194
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
            TabIndex        =   193
            Text            =   "Diagnósticos"
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   250
            Width           =   4785
         End
         Begin UltraGrid.SSUltraGrid grdDx 
            Height          =   2205
            Left            =   60
            TabIndex        =   114
            Top             =   615
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   3889
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "F1 = Busca Dx por descripción"
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
            Left            =   120
            TabIndex        =   159
            Top             =   2820
            Width           =   2520
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Hospitalizados"
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
         Left            =   -65070
         TabIndex        =   82
         Top             =   5640
         Width           =   2445
         Begin MSMask.MaskEdBox txtHfingreso 
            Height          =   315
            Left            =   960
            TabIndex        =   32
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
            Left            =   960
            TabIndex        =   33
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
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
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
            Left            =   150
            TabIndex        =   84
            Top             =   270
            Width           =   765
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
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
            Left            =   150
            TabIndex        =   83
            Top             =   630
            Width           =   525
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
         Height          =   3915
         Left            =   -74850
         TabIndex        =   81
         Top             =   330
         Width           =   7695
         Begin VB.Frame fraAdmVitaminaK 
            Caption         =   "Adm. Vitamina K"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   4200
            TabIndex        =   228
            Top             =   2610
            Width           =   1725
            Begin Threed.SSCheck chkSPvitaminaKsi 
               Height          =   225
               Left            =   90
               TabIndex        =   229
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
            Begin Threed.SSCheck chkSPvitaminaKno 
               Height          =   225
               Left            =   1110
               TabIndex        =   230
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
            Begin VB.Label Label69 
               AutoSize        =   -1  'True
               Caption         =   "(311)"
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
               Left            =   1350
               TabIndex        =   231
               Top             =   0
               Width           =   315
            End
         End
         Begin VB.Frame fraProfilaxisO 
            Caption         =   "Profilaxis Ocular"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   5970
            TabIndex        =   224
            Top             =   2610
            Width           =   1665
            Begin Threed.SSCheck chkSPprofilaxisOsi 
               Height          =   225
               Left            =   90
               TabIndex        =   225
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
            Begin Threed.SSCheck chkSPprofilaxisOno 
               Height          =   225
               Left            =   1020
               TabIndex        =   226
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
            Begin VB.Label Label72 
               AutoSize        =   -1  'True
               Caption         =   "(309)"
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
               TabIndex        =   227
               Top             =   0
               Width           =   315
            End
         End
         Begin VB.Frame fraEEDP 
            Caption         =   "EEDP/TEPSI"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   2640
            TabIndex        =   220
            Top             =   2610
            Width           =   1485
            Begin Threed.SSCheck chkSPeedpSI 
               Height          =   225
               Left            =   90
               TabIndex        =   221
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
               Left            =   900
               TabIndex        =   222
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
               Left            =   1080
               TabIndex        =   223
               Top             =   0
               Width           =   315
            End
         End
         Begin VB.Frame fraConsejPPFF 
            Caption         =   "Consejería PP.FF."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   5790
            TabIndex        =   216
            Top             =   1440
            Width           =   1845
            Begin Threed.SSCheck chkSPconsejeriaPPffSI 
               Height          =   225
               Left            =   90
               TabIndex        =   217
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
            Begin Threed.SSCheck chkSPconsejeriaPPffNO 
               Height          =   225
               Left            =   1230
               TabIndex        =   218
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
               Caption         =   "(308)"
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
               Left            =   1500
               TabIndex        =   219
               Top             =   0
               Width           =   315
            End
         End
         Begin VB.Frame fraConsejNutricional 
            Caption         =   "Consejería Nutricional"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   5520
            TabIndex        =   212
            Top             =   240
            Width           =   2115
            Begin Threed.SSCheck chkSPconsejeriaNsi 
               Height          =   225
               Left            =   90
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
            Begin Threed.SSCheck chkSPconsejeriaNno 
               Height          =   225
               Left            =   1500
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
               Left            =   1770
               TabIndex        =   215
               Top             =   0
               Width           =   315
            End
         End
         Begin VB.Frame fraAdmSuplNutr 
            Caption         =   "Adm.Supl.Nutr"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   6030
            TabIndex        =   208
            Top             =   3300
            Width           =   1605
            Begin Threed.SSCheck chkSPsuplNsi 
               Height          =   225
               Left            =   90
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
            Begin Threed.SSCheck chkSPsuplNno 
               Height          =   225
               Left            =   1020
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
            Begin VB.Label Label71 
               AutoSize        =   -1  'True
               Caption         =   "(310)"
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
               Left            =   1260
               TabIndex        =   211
               Top             =   0
               Width           =   315
            End
         End
         Begin VB.Frame fraLactanciaM 
            Caption         =   "Lactancia Mat.Excl"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   1830
            TabIndex        =   204
            Top             =   3330
            Width           =   1905
            Begin Threed.SSCheck chkSPlactanciaMsi 
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
            Begin Threed.SSCheck chkSPlactanciaMno 
               Height          =   225
               Left            =   1290
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
            Begin VB.Label Label68 
               AutoSize        =   -1  'True
               Caption         =   "(002)"
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
               Left            =   1530
               TabIndex        =   207
               Top             =   0
               Width           =   315
            End
         End
         Begin VB.Frame fraAdmOxitocina 
            Caption         =   "Adm.Oxitocina"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   120
            TabIndex        =   200
            Top             =   3330
            Width           =   1635
            Begin Threed.SSCheck chkSPadmOxitocinaSI 
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
            Begin Threed.SSCheck chkSPadmOxitocinaNO 
               Height          =   225
               Left            =   990
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
            Begin VB.Label Label65 
               AutoSize        =   -1  'True
               Caption         =   "(303)"
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
               Left            =   1230
               TabIndex        =   203
               Top             =   0
               Width           =   315
            End
         End
         Begin VB.Frame fraPsicoprofilaxis 
            Caption         =   "Psicoprofilaxis "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   120
            TabIndex        =   196
            Top             =   2610
            Width           =   1635
            Begin Threed.SSCheck chkSPsicoprofilaxisSI 
               Height          =   225
               Left            =   90
               TabIndex        =   198
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
            Begin Threed.SSCheck chkSPsicoprofilaxisNO 
               Height          =   225
               Left            =   990
               TabIndex        =   199
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
               Caption         =   "(302)"
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
               Left            =   1170
               TabIndex        =   197
               Top             =   0
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
            Left            =   4410
            TabIndex        =   97
            Top             =   3510
            Width           =   645
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
            Left            =   1920
            TabIndex        =   96
            Top             =   2790
            Width           =   645
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
            Left            =   4950
            TabIndex        =   94
            Top             =   1560
            Width           =   315
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
            Left            =   4350
            TabIndex        =   93
            Top             =   1560
            Width           =   315
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
            Left            =   2610
            TabIndex        =   92
            Top             =   1590
            Width           =   645
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
            Left            =   4380
            TabIndex        =   89
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   420
            Width           =   645
         End
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
            Left            =   2640
            TabIndex        =   88
            Top             =   420
            Width           =   645
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
            Left            =   1290
            TabIndex        =   91
            Top             =   1575
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
            Left            =   1320
            TabIndex        =   90
            Top             =   930
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
            Left            =   1320
            TabIndex        =   87
            Top             =   420
            Width           =   825
         End
         Begin MSMask.MaskEdBox txtSPpa 
            Height          =   315
            Left            =   1290
            TabIndex        =   95
            Top             =   2040
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
            Left            =   3870
            TabIndex        =   177
            Top             =   3450
            Width           =   315
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
            Left            =   2100
            TabIndex        =   176
            Top             =   2400
            Width           =   315
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
            Left            =   2370
            TabIndex        =   175
            Top             =   2070
            Width           =   765
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
            Left            =   4740
            TabIndex        =   174
            Top             =   1890
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
            Left            =   4080
            TabIndex        =   173
            Top             =   1890
            Width           =   315
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
            Left            =   3270
            TabIndex        =   172
            Top             =   1620
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
            Left            =   90
            TabIndex        =   171
            Top             =   1440
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
            Left            =   2190
            TabIndex        =   170
            Top             =   960
            Width           =   315
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
            Left            =   5070
            TabIndex        =   169
            Top             =   210
            Width           =   315
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "(003)"
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
            TabIndex        =   168
            Top             =   180
            Width           =   315
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
            Left            =   90
            TabIndex        =   167
            Top             =   630
            Width           =   315
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4830
            TabIndex        =   141
            Top             =   1590
            Width           =   90
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4230
            TabIndex        =   140
            Top             =   1590
            Width           =   90
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Control de puerperio (N°)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   127
            Top             =   3270
            Width           =   1920
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Cred (N°)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1920
            TabIndex        =   126
            Top             =   2580
            Width           =   690
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "A  P  G  A  R"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4350
            TabIndex        =   125
            Top             =   1350
            Width           =   870
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Edad Gest RN (Sem)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2310
            TabIndex        =   124
            Top             =   1380
            Width           =   1455
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Talla (cm)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4380
            TabIndex        =   120
            Top             =   210
            Width           =   690
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Peso  (Kg)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2655
            TabIndex        =   119
            Top             =   210
            Width           =   735
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "P.A.(mmHg)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   118
            Top             =   2070
            Width           =   870
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Altu.Uterina (cm)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   117
            Top             =   1620
            Width           =   1230
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Edad Gest (Sem)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   116
            Top             =   975
            Width           =   1200
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "CPN (N°)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   115
            Top             =   450
            Width           =   645
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
         Height          =   3915
         Left            =   -66990
         TabIndex        =   80
         Top             =   330
         Width           =   4395
         Begin VB.TextBox txtVacEpatB 
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
            TabIndex        =   109
            Top             =   2310
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
            Left            =   1530
            TabIndex        =   113
            Top             =   3480
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
            Left            =   150
            TabIndex        =   112
            Top             =   3480
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
            TabIndex        =   111
            Top             =   2910
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
            TabIndex        =   110
            Top             =   2910
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
            TabIndex        =   108
            Top             =   2310
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
            TabIndex        =   107
            Top             =   2310
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
            TabIndex        =   106
            Top             =   1710
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
            TabIndex        =   105
            Top             =   1710
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
            TabIndex        =   104
            Top             =   1710
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
            TabIndex        =   103
            Top             =   1110
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
            TabIndex        =   102
            Top             =   1110
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
            Left            =   150
            TabIndex        =   101
            Top             =   1110
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
            Left            =   3090
            TabIndex        =   100
            Top             =   510
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
            TabIndex        =   99
            Top             =   510
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
            Left            =   150
            TabIndex        =   98
            Top             =   510
            Width           =   645
         End
         Begin VB.Label Label88 
            AutoSize        =   -1  'True
            Caption         =   "ANTI.EPATIT-B"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3120
            TabIndex        =   235
            Top             =   2100
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "(119)"
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
            Left            =   2760
            TabIndex        =   234
            Top             =   2100
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
            Left            =   1200
            TabIndex        =   192
            Top             =   3270
            Width           =   315
         End
         Begin VB.Label Label86 
            AutoSize        =   -1  'True
            Caption         =   "(315)"
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
            Left            =   450
            TabIndex        =   191
            Top             =   3270
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
            Left            =   1170
            TabIndex        =   190
            Top             =   2700
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
            Left            =   1170
            TabIndex        =   189
            Top             =   2100
            Width           =   315
         End
         Begin VB.Label Label83 
            AutoSize        =   -1  'True
            Caption         =   "(125)"
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
            Left            =   480
            TabIndex        =   188
            Top             =   2700
            Width           =   315
         End
         Begin VB.Label Label82 
            AutoSize        =   -1  'True
            Caption         =   "(314)"
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
            Left            =   450
            TabIndex        =   187
            Top             =   2100
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
            Left            =   2730
            TabIndex        =   186
            Top             =   1500
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
            Left            =   1200
            TabIndex        =   185
            Top             =   1500
            Width           =   315
         End
         Begin VB.Label Label79 
            AutoSize        =   -1  'True
            Caption         =   "(313)"
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
            Left            =   480
            TabIndex        =   184
            Top             =   1500
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
            Left            =   2730
            TabIndex        =   183
            Top             =   900
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
            Left            =   1200
            TabIndex        =   182
            Top             =   900
            Width           =   315
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "(117)"
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
            Left            =   450
            TabIndex        =   181
            Top             =   900
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
            Left            =   2760
            TabIndex        =   180
            Top             =   300
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
            Left            =   1200
            TabIndex        =   179
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            Caption         =   "(102)"
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
            Left            =   450
            TabIndex        =   178
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "PENTAVAL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1530
            TabIndex        =   139
            Top             =   3270
            Width           =   750
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "HVB"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   138
            Top             =   3300
            Width           =   285
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "DT ADULTO (N° Dosis)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1530
            TabIndex        =   137
            Top             =   2700
            Width           =   1605
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "SPR"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   136
            Top             =   2730
            Width           =   285
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "ROTAVIRUS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1530
            TabIndex        =   135
            Top             =   2100
            Width           =   870
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "ASA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   134
            Top             =   2130
            Width           =   300
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "ANTITETANICA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3060
            TabIndex        =   133
            Top             =   1500
            Width           =   1110
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "RUBEOLA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1530
            TabIndex        =   132
            Top             =   1500
            Width           =   690
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "APO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   131
            Top             =   1530
            Width           =   315
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "ANTINEUMOC"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3060
            TabIndex        =   130
            Top             =   900
            Width           =   1005
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "PAROTID"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1530
            TabIndex        =   129
            Top             =   900
            Width           =   675
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "DPT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   128
            Top             =   930
            Width           =   285
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "ANTIAMARILICA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3060
            TabIndex        =   123
            Top             =   300
            Width           =   1200
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "INFLUENZ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1530
            TabIndex        =   122
            Top             =   300
            Width           =   720
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "BCG"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   121
            Top             =   330
            Width           =   300
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Referencia Destino"
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
         Left            =   -72150
         TabIndex        =   76
         Top             =   5640
         Width           =   7035
         Begin VB.CommandButton btnBuscarEstablecimientoD 
            Caption         =   "..."
            Height          =   315
            Left            =   720
            TabIndex        =   26
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
            TabIndex        =   31
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   510
            Width           =   585
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
            Left            =   1020
            TabIndex        =   77
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   510
            Width           =   4095
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
            Left            =   5160
            MaxLength       =   20
            TabIndex        =   27
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   510
            Width           =   1665
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Cod.ES/Eq AISPED al que se refiere al Paciente"
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
            TabIndex        =   79
            Top             =   270
            Width           =   3870
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "N° Hoja Refer/Cont"
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
            Left            =   5220
            TabIndex        =   78
            Top             =   270
            Width           =   1590
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
         TabIndex        =   74
         Top             =   5640
         Width           =   2745
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
            TabIndex        =   25
            Top             =   510
            Width           =   2535
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
            TabIndex        =   75
            Top             =   270
            Width           =   1815
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Referencia Origen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -69840
         TabIndex        =   70
         Top             =   4470
         Width           =   7215
         Begin VB.CommandButton btnBuscarEstablecimientoO 
            Caption         =   "..."
            Height          =   315
            Left            =   810
            TabIndex        =   23
            Top             =   540
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
            Left            =   5460
            MaxLength       =   20
            TabIndex        =   24
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   540
            Width           =   1635
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
            Left            =   1140
            TabIndex        =   72
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   540
            Width           =   4305
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
            TabIndex        =   30
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   540
            Width           =   675
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "N° Hoja Referencia"
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
            Left            =   5490
            TabIndex        =   73
            Top             =   330
            Width           =   1545
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Cod.ES/Eq   AISPED que Refirió al Paciente"
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
            TabIndex        =   71
            Top             =   330
            Width           =   3540
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Otros datos de Ingreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74940
         TabIndex        =   68
         Top             =   4470
         Width           =   5025
         Begin VB.Frame FraPersonal 
            Caption         =   "Personal que atiende"
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
            TabIndex        =   69
            Top             =   270
            Width           =   4785
            Begin Threed.SSCheck chkPAestablecimiento 
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   270
               Width           =   1875
               _ExtentX        =   3307
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
               Caption         =   "Del Establecimiento"
               Value           =   1
            End
            Begin Threed.SSCheck chkPAaisped 
               Height          =   315
               Left            =   2400
               TabIndex        =   22
               Top             =   270
               Width           =   2145
               _ExtentX        =   3784
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
               Caption         =   "Itinerante/Eq.AISPED"
            End
         End
      End
      Begin VB.Frame fraReconsideracion 
         Height          =   675
         Left            =   -65190
         TabIndex        =   46
         Top             =   399
         Width           =   2565
         Begin VB.TextBox txtReconsideracion 
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
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   47
            Top             =   240
            Width           =   765
         End
         Begin Threed.SSCheck chkReconsideracion 
            Height          =   255
            Left            =   120
            TabIndex        =   166
            Top             =   270
            Width           =   1635
            _ExtentX        =   2884
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
            Caption         =   "Reconsideración"
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Datos del Establecimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -74940
         TabIndex        =   42
         Top             =   390
         Width           =   9705
         Begin VB.TextBox txtCS 
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
            Left            =   3780
            TabIndex        =   45
            Top             =   240
            Width           =   5775
         End
         Begin VB.TextBox txtCScodigo 
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
            Left            =   2370
            TabIndex        =   43
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo ES/Equipo AISPED"
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
            TabIndex        =   44
            Top             =   270
            Width           =   2130
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
         Height          =   1485
         Left            =   -74940
         TabIndex        =   41
         Top             =   2940
         Width           =   12315
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
            Height          =   585
            Left            =   90
            TabIndex        =   158
            Top             =   210
            Width           =   3795
            Begin Threed.SSCheck chkAtencionAmbulatoria 
               Height          =   255
               Left            =   90
               TabIndex        =   10
               Top             =   210
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
            End
            Begin Threed.SSCheck chkAtencionReferencia 
               Height          =   255
               Left            =   1380
               TabIndex        =   11
               Top             =   210
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
               TabIndex        =   12
               Top             =   210
               Width           =   1035
               _ExtentX        =   1826
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
               Caption         =   "Emergen"
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
            Height          =   585
            Left            =   3930
            TabIndex        =   67
            Top             =   780
            Width           =   2745
            Begin Threed.SSCheck chkIntramural 
               Height          =   255
               Left            =   90
               TabIndex        =   15
               Top             =   210
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
               Left            =   1410
               TabIndex        =   16
               Top             =   210
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
         Begin VB.Frame Frame9 
            Caption         =   "Concepto Prestacional"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1155
            Left            =   6720
            TabIndex        =   61
            Top             =   230
            Width           =   4185
            Begin VB.TextBox txtNautorizacion 
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
               Left            =   1395
               MaxLength       =   15
               TabIndex        =   18
               Top             =   750
               Width           =   1185
            End
            Begin VB.TextBox txtMonto 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   19
               Top             =   750
               Width           =   885
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
               TabIndex        =   17
               Top             =   270
               Width           =   4035
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "N° Autorización"
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
               TabIndex        =   63
               Top             =   810
               Width           =   1260
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Monto"
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
               Left            =   2700
               TabIndex        =   62
               Top             =   780
               Width           =   525
            End
         End
         Begin VB.Frame fraGestantePuerpera 
            Height          =   555
            Left            =   3930
            TabIndex        =   60
            Top             =   240
            Width           =   2745
            Begin Threed.SSCheck chkGestante 
               Height          =   255
               Left            =   90
               TabIndex        =   13
               Top             =   180
               Width           =   1035
               _ExtentX        =   1826
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
               Left            =   1410
               TabIndex        =   14
               Top             =   180
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
         Begin MSMask.MaskEdBox txtFparto 
            Height          =   315
            Left            =   10890
            TabIndex        =   20
            Top             =   510
            Width           =   1335
            _ExtentX        =   2355
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
         Begin MSMask.MaskEdBox txtFantencion 
            Height          =   315
            Left            =   1440
            TabIndex        =   64
            Top             =   990
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
         Begin MSMask.MaskEdBox txtHatencion 
            Height          =   315
            Left            =   2850
            TabIndex        =   66
            Top             =   990
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
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Atención"
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
            TabIndex        =   65
            Top             =   1020
            Width           =   1275
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Parto"
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
            Left            =   11040
            TabIndex        =   59
            Top             =   270
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   1755
         Left            =   -74940
         TabIndex        =   39
         Top             =   1140
         Width           =   12315
         Begin VB.Frame FraTipoAfiliacion 
            Caption         =   "Tipo F.Afiliación"
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
            Left            =   2370
            TabIndex        =   157
            Top             =   240
            Width           =   1755
            Begin Threed.SSCheck chkTAnuevo 
               Height          =   255
               Left            =   120
               TabIndex        =   3
               Top             =   300
               Width           =   1035
               _ExtentX        =   1826
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
               Caption         =   "Nuevo"
            End
            Begin Threed.SSCheck chkTAAntiguoI 
               Height          =   285
               Left            =   120
               TabIndex        =   4
               Top             =   660
               Width           =   1515
               _ExtentX        =   2672
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
               Caption         =   "Antiguo - Inscr"
            End
            Begin Threed.SSCheck chkTAantiguoA 
               Height          =   285
               Left            =   120
               TabIndex        =   5
               Top             =   1020
               Width           =   1455
               _ExtentX        =   2566
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
               Caption         =   "Antiguo - Afil"
            End
         End
         Begin VB.Frame FraComponente 
            Caption         =   "Componente/Régimen"
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
            Left            =   120
            TabIndex        =   156
            Top             =   240
            Width           =   2205
            Begin Threed.SSCheck chkCsubsidiado 
               Height          =   255
               Left            =   120
               TabIndex        =   1
               Top             =   300
               Width           =   1215
               _ExtentX        =   2143
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
               Caption         =   "Subsidiado"
            End
            Begin Threed.SSCheck chkCSemiS 
               Height          =   555
               Left            =   120
               TabIndex        =   2
               Top             =   750
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   979
               _Version        =   262144
               CaptionStyle    =   1
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
               Caption         =   "SemiSubsidiado/ Semicontributivo"
            End
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
            Left            =   6390
            TabIndex        =   85
            Top             =   510
            Width           =   1125
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
            Left            =   10290
            TabIndex        =   9
            Top             =   1290
            Width           =   1875
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
            Left            =   7530
            TabIndex        =   56
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   1290
            Width           =   1605
         End
         Begin VB.TextBox txtPaciente 
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
            Left            =   4170
            TabIndex        =   53
            Top             =   900
            Width           =   4635
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
            Height          =   975
            Left            =   9180
            TabIndex        =   50
            Top             =   210
            Width           =   3045
            Begin VB.TextBox txtCodSeguro 
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
               Left            =   1080
               TabIndex        =   29
               Top             =   600
               Width           =   1905
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
               TabIndex        =   28
               Text            =   "0"
               Top             =   240
               Width           =   1905
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
               Height          =   210
               Left            =   90
               TabIndex        =   52
               Top             =   630
               Width           =   960
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
               Height          =   210
               Left            =   90
               TabIndex        =   51
               Top             =   270
               Width           =   855
            End
         End
         Begin VB.TextBox txtNdocumento 
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
            Left            =   7500
            TabIndex        =   48
            Top             =   510
            Width           =   1635
         End
         Begin VB.TextBox txtNroAfiliacion1 
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
            Left            =   4170
            TabIndex        =   6
            Top             =   510
            Width           =   435
         End
         Begin VB.TextBox txtNroAfiliacion2 
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
            Left            =   4620
            TabIndex        =   7
            Top             =   510
            Width           =   375
         End
         Begin VB.TextBox txtNroAfiliacion3 
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
            Left            =   5010
            TabIndex        =   8
            Top             =   510
            Width           =   1185
         End
         Begin MSMask.MaskEdBox txtFnacimiento 
            Height          =   315
            Left            =   5280
            TabIndex        =   54
            Top             =   1290
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
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "N° Historia"
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
            Left            =   9300
            TabIndex        =   58
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Sexo"
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
            Left            =   7110
            TabIndex        =   57
            Top             =   1320
            Width           =   405
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "F.Nacimiento"
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
            Left            =   4170
            TabIndex        =   55
            Top             =   1320
            Width           =   1050
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Identificación"
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
            Left            =   6390
            TabIndex        =   49
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lblAfiliacionSIS 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Cod.Afiliación/Inscripción"
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
            Left            =   4170
            TabIndex        =   40
            Top             =   270
            Width           =   1995
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
         TabIndex        =   38
         Top             =   8970
         Width           =   1515
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Especialidad"
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
         Left            =   -64110
         TabIndex        =   148
         Top             =   7380
         Width           =   960
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "Responsable de la atención"
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
         Left            =   -74790
         TabIndex        =   147
         Top             =   7350
         Width           =   2220
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "N° Formato FUA"
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
      Left            =   60
      TabIndex        =   165
      Top             =   60
      Width           =   1800
   End
End
Attribute VB_Name = "SisFua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Registra y emite formato FUA
'        Programado por: Barrantes D
'        Fecha: Enero 2013
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
Dim mo_cmbIdDestinoAtencion As New sighentidades.ListaDespleglable
Dim mo_cmbConceptoP As New sighentidades.ListaDespleglable
Dim mo_cmbTipoDocumento As New sighentidades.ListaDespleglable
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim oDoSisFuaAtencion As New DoSisFuaAtencion
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
Dim wxParametro358 As String
Dim lnNroFuaRepetido As Boolean
Dim mo_lbEsAltaMedica As Boolean

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
'    If Not (mi_opcion = sghAgregar Or mi_opcion = sghModificar Or mi_opcion = sghEliminar) Then
'        Me.btnAceptar.Visible = False
'    End If
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
                'oBusqueda.CodigoDx = lcCodigoDx
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
            chkCsubsidiado.SetFocus
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

Private Sub btnAceptar_Click()
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
    Dim lbEstaOk As Boolean
    
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
                And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghAdmisionEmergencia Then
       MsgBox "Debe elegir CODIGO DE PRESTACION", vbInformation, Me.Caption
       On Error Resume Next
       ucSISfuaCodPrestacion1.SetFocus
       Exit Function
    End If
    If chkCsubsidiado.Value = 0 And chkCSemiS.Value = 0 Then
       MsgBox "Debe marcar algún COMPONENTE", vbInformation, Me.Caption
       Me.TabFua.Tab = 0
       Exit Function
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
                And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghAdmisionEmergencia Then
       MsgBox "Debe marcar alguna ATENCION", vbInformation, Me.Caption
       Me.TabFua.Tab = 0
       Exit Function
    End If
    'mgaray20140926
    If Me.cmbConceptoP.Text = "" And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghRegistroCitaCE _
                And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghAdmisionEmergencia Then
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
    If chkPAestablecimiento.Value = 0 And chkPAaisped.Value = 0 Then
       MsgBox "Debe marcar PERSONAL QUE ATIENDE", vbInformation, Me.Caption
       Me.TabFua.Tab = 0
       Exit Function
    End If
    'mgaray20140926
    If Val(mo_cmbIdDestinoAtencion.BoundText) = 0 And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghRegistroCitaCE _
                     And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghAdmisionEmergencia Then
       MsgBox "Debe elegir DESTINO DEL ASEGURADO", vbInformation, Me.Caption
       Me.TabFua.Tab = 0
       If cmbIdDestinoAtencion.Enabled = True Then cmbIdDestinoAtencion.SetFocus
       Exit Function
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
    If mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghRegistroCitaCE And mo_lnIdTablaLISTBARITEMS <> sghOpcionGalenHos.sghAdmisionEmergencia Then
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
    If txtReconsideracion.Locked = False And txtReconsideracion.Text = "" Then
       MsgBox "Debe registrar la RECONSIDERACION", vbInformation, Me.Caption
       Me.TabFua.Tab = 0
       On Error Resume Next
       txtReconsideracion.SetFocus
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
    Dim lcMensaje As String
    lcMensaje = ""
    If mi_opcion = sghAgregar Then
        If lcCodigoEstablecimientoAdscripcionSIS <> "" Then
            lcMensaje = mo_ReglasSISgalenhos.ChequeaCodigoEstablecimientoAdscripcion(lcCodigoEstablecimientoAdscripcionSIS, _
                                                ml_IdTipoServicio, _
                                                IIf(txtRONumero.Text <> "", 4, 0), _
                                                ucSISfuaCodPrestacion1.CodigoPrestacion)
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
           txtNhistoriaClinica.Text = oDOPaciente.NroHistoriaClinica
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


Private Sub chkCSemiS_Click(Value As Integer)
    If chkCSemiS.Value = -1 Then
       chkCsubsidiado.Value = 0
    End If

End Sub

Private Sub chkCSemiS_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkCSemiS
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkCSemiS_LostFocus()
    chkCSemiS_Click 1
End Sub


Private Sub chkCsubsidiado_Click(Value As Integer)
    If chkCsubsidiado.Value = -1 Then
       chkCSemiS.Value = 0
    End If
End Sub

Private Sub chkCsubsidiado_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkCsubsidiado
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkCsubsidiado_LostFocus()
    chkCsubsidiado_Click 1
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






Private Sub chkReconsideracion_Click(Value As Integer)
    If chkReconsideracion.Value = -1 Then
       mo_Formulario.HabilitarDeshabilitar txtReconsideracion, True
       txtReconsideracion.SetFocus
    Else
       mo_Formulario.HabilitarDeshabilitar txtReconsideracion, False
    End If

End Sub



Private Sub chkSPadmOxitocinaNO_Click(Value As Integer)
    If Me.chkSPadmOxitocinaNO.Value = -1 Then
       Me.chkSPadmOxitocinaSI.Value = 0
    End If
   
End Sub

Private Sub chkSPadmOxitocinaNO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPadmOxitocinaNO
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPadmOxitocinaNO_LostFocus()
    chkSPadmOxitocinaNO_Click 1
End Sub

Private Sub chkSPadmOxitocinaSI_Click(Value As Integer)
    If Me.chkSPadmOxitocinaSI.Value = -1 Then
       Me.chkSPadmOxitocinaNO.Value = 0
    End If

End Sub

Private Sub chkSPadmOxitocinaSI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPadmOxitocinaSI
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPadmOxitocinaSI_LostFocus()
    chkSPadmOxitocinaSI_Click 1
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





Private Sub chkSPconsejeriaPPffNO_Click(Value As Integer)
    If Me.chkSPconsejeriaPPffNO.Value = -1 Then
       Me.chkSPconsejeriaPPffSI.Value = 0
    End If

End Sub

Private Sub chkSPconsejeriaPPffNO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPconsejeriaPPffNO
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkSPconsejeriaPPffNO_LostFocus()
    chkSPconsejeriaPPffNO_Click 1
End Sub

Private Sub chkSPconsejeriaPPffSI_Click(Value As Integer)
    If Me.chkSPconsejeriaPPffSI.Value = -1 Then
       Me.chkSPconsejeriaPPffNO.Value = 0
    End If


End Sub

Private Sub chkSPconsejeriaPPffSI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPconsejeriaPPffSI
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPconsejeriaPPffSI_LostFocus()
    chkSPconsejeriaPPffSI_Click 1
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

Private Sub chkSPlactanciaMno_Click(Value As Integer)
    If Me.chkSPlactanciaMno.Value = -1 Then
       Me.chkSPlactanciaMsi.Value = 0
    End If

End Sub

Private Sub chkSPlactanciaMno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPlactanciaMno
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPlactanciaMno_LostFocus()
    chkSPlactanciaMno_Click 1
End Sub

Private Sub chkSPlactanciaMsi_Click(Value As Integer)
    If Me.chkSPlactanciaMsi.Value = -1 Then
       Me.chkSPlactanciaMno.Value = 0
    End If

End Sub

Private Sub chkSPlactanciaMsi_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPlactanciaMsi
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPlactanciaMsi_LostFocus()
    chkSPlactanciaMsi_Click 1
End Sub








Private Sub chkSPprofilaxisOno_Click(Value As Integer)
    If Me.chkSPprofilaxisOno.Value = -1 Then
       Me.chkSPprofilaxisOsi.Value = 0
    End If

End Sub

Private Sub chkSPprofilaxisOno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPprofilaxisOno
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPprofilaxisOno_LostFocus()
    chkSPprofilaxisOno_Click 1
End Sub

Private Sub chkSPprofilaxisOsi_Click(Value As Integer)
    If Me.chkSPprofilaxisOsi.Value = -1 Then
       Me.chkSPprofilaxisOno.Value = 0
    End If

End Sub

Private Sub chkSPprofilaxisOsi_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPprofilaxisOsi
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPprofilaxisOsi_LostFocus()
     chkSPprofilaxisOsi_Click 1
End Sub

Private Sub chkSPsicoprofilaxisNO_Click(Value As Integer)
    If Me.chkSPsicoprofilaxisNO.Value = -1 Then
       Me.chkSPsicoprofilaxisSI.Value = 0
    End If
End Sub

Private Sub chkSPsicoprofilaxisNO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPsicoprofilaxisNO
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPsicoprofilaxisNO_LostFocus()
    chkSPsicoprofilaxisNO_Click 1
End Sub

Private Sub chkSPsicoprofilaxisSI_Click(Value As Integer)
    If Me.chkSPsicoprofilaxisSI.Value = -1 Then
       Me.chkSPsicoprofilaxisNO.Value = 0
    End If
End Sub

Private Sub chkSPsicoprofilaxisSI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPsicoprofilaxisSI
    AdministrarKeyPreview KeyCode
End Sub

Private Sub chkSPsicoprofilaxisSI_LostFocus()
    chkSPsicoprofilaxisSI_Click 1
End Sub



Private Sub chkSPsuplNno_Click(Value As Integer)
    If Me.chkSPsuplNno.Value = -1 Then
       Me.chkSPsuplNsi.Value = 0
    End If

End Sub

Private Sub chkSPsuplNno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPsuplNno
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPsuplNno_LostFocus()
    chkSPsuplNno_Click 1
End Sub

Private Sub chkSPsuplNsi_Click(Value As Integer)
    If Me.chkSPsuplNsi.Value = -1 Then
       Me.chkSPsuplNno.Value = 0
    End If

End Sub

Private Sub chkSPsuplNsi_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPsuplNsi
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPsuplNsi_LostFocus()
   chkSPsuplNsi_Click 1
End Sub



Private Sub chkSPvitaminaKno_Click(Value As Integer)
    If Me.chkSPvitaminaKno.Value = -1 Then
       Me.chkSPvitaminaKsi.Value = 0
    End If

End Sub

Private Sub chkSPvitaminaKno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPvitaminaKno
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPvitaminaKno_LostFocus()
    chkSPvitaminaKno_Click 1
End Sub

Private Sub chkSPvitaminaKsi_Click(Value As Integer)
    If Me.chkSPvitaminaKsi.Value = -1 Then
       Me.chkSPvitaminaKno.Value = 0
    End If

End Sub

Private Sub chkSPvitaminaKsi_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkSPvitaminaKsi
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkSPvitaminaKsi_LostFocus()
    chkSPvitaminaKsi_Click 1
End Sub

Private Sub chkTAantiguoA_Click(Value As Integer)
    If chkTAantiguoA.Value = -1 Then
       chkTAnuevo.Value = 0
       chkTAAntiguoI.Value = 0
       On Error Resume Next
       txtInstitucion.SetFocus
    End If

End Sub

Private Sub chkTAantiguoA_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkTAantiguoA
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkTAantiguoA_LostFocus()
    chkTAantiguoA_Click 1
End Sub


Private Sub chkTAAntiguoI_Click(Value As Integer)
    If chkTAAntiguoI.Value = -1 Then
       chkTAnuevo.Value = 0
       chkTAantiguoA.Value = 0
       On Error Resume Next
       txtInstitucion.SetFocus
    End If

End Sub

Private Sub chkTAAntiguoI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkTAAntiguoI
    AdministrarKeyPreview KeyCode

End Sub

Private Sub chkTAAntiguoI_LostFocus()
    chkTAAntiguoI_Click 1
End Sub


Private Sub chkTAnuevo_Click(Value As Integer)
    If chkTAnuevo.Value = -1 Then
       chkTAAntiguoI.Value = 0
       chkTAantiguoA.Value = 0
       On Error Resume Next
       txtInstitucion.SetFocus
    End If

End Sub

Private Sub chkTAnuevo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkTAnuevo
    AdministrarKeyPreview KeyCode

End Sub






Private Sub chkTAnuevo_LostFocus()
    chkTAnuevo_Click 1
End Sub

Private Sub cmbConceptoP_Click()
    If Val(mo_cmbConceptoP.BoundText) = 2 Or Val(mo_cmbConceptoP.BoundText) = 3 Then
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
            'Me.btnBuscarEstablecimientoD.Enabled = True
            mo_Formulario.HabilitarDeshabilitar Me.txtRDnumero, True
        Else
            'Me.btnBuscarEstablecimientoD.Enabled = False
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




Private Sub Form_Activate()
  If lbEsIgualQueArSIS = False Then
        If mo_lbCargaTablasUnaVez = True Then
            mo_lbCargaTablasUnaVez = False
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
'        If mi_opcion = sghAgregar Then
'           mo_Formulario.HabilitarDeshabilitar txtFua3, True
'        End If
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
End Sub



Sub ReglasDeConsistenciasAntesDeCargarFormulario()
     Me.ucSISfuaCodPrestacion1.ReglasDeConsistenciasAntesDeCargarFormulario ml_IdTipoServicio, Left(txtSexo.Text, 1), ml_edad_En_YYYYMMDD
        
End Sub

Private Sub Form_Load()
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
    mo_Formulario.HabilitarDeshabilitar txtSexo, False
    mo_Formulario.HabilitarDeshabilitar txtNhistoriaClinica, False
    mo_Formulario.HabilitarDeshabilitar txtFantencion, False
    mo_Formulario.HabilitarDeshabilitar txtHatencion, False
    mo_Formulario.HabilitarDeshabilitar txtMedicoDni, False
    mo_Formulario.HabilitarDeshabilitar txtMedico, False
    mo_Formulario.HabilitarDeshabilitar txtMedicoColegiatura, False
    mo_Formulario.HabilitarDeshabilitar txtMedicoEspecialidad, False
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
    mo_Formulario.HabilitarDeshabilitar txtReconsideracion, False
    'mo_Formulario.HabilitarDeshabilitar FraTipoAfiliacion, False
    mo_Formulario.HabilitarDeshabilitar fraCodAfiliacionSeguro, False
    mo_Formulario.HabilitarDeshabilitar fraReconsideracion, False
    'mo_Formulario.HabilitarDeshabilitar FraComponente, False
    btnBuscarPaciente.Enabled = False
    '
    CreaTemporales
    '
    CargaComboBoxes
    '
    CargarDatosAlFormulario
    '
    
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
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.Open wxParametroJAMO
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ml_IdConceptoPrestacional = ""
    If CargarDatosDelPaciente(oConexion) = True Then
        CargaFormatoFUA
        
        txtCScodigo.Text = wxParametro280           '?
        txtCS.Text = wxParametro205
        Me.txtPaciente = ml_Paciente
        txtNdocumento.Text = ml_NroDocumento
        txtFnacimiento.Text = Format(md_FechaNacimiento, sighentidades.DevuelveFechaSoloFormato_DMY)
        txtSexo.Text = ml_Sexo
        txtNhistoriaClinica.Text = ml_NroHistoriaClinica
        txtFantencion.Text = Format(md_FechaAtencion, sighentidades.DevuelveFechaSoloFormato_DMY)
        txtHatencion.Text = ml_HoraAtencion
        '
        CargaDatosDeDx oConexion, True
        CargaDatosMedico oConexion, False
        CargaConsumosEnServiciosIntermedios oConexion, True
        CargaDatosDeAfiliacion True
        CargaDatosDeTriajeVacunas oConexionExterna, True
        
    Else
        Me.btnAceptar.Enabled = False
        Me.btnImprimir.Enabled = False
    End If
    oConexion.Close
    oConexionExterna.Close
    Set oConexion = Nothing
    Set oConexionExterna = Nothing
End Sub

Sub CargaDatosDeDx(oConexion As Connection, lbDesdeGalenHos As Boolean)
    Dim oRsTmp1 As New Recordset
    Dim lnDxNro As Integer, lnUno As Integer
    If lbDesdeGalenHos = True Then
        lnUno = 1
        mo_ReglasSISgalenhos.FuaCargaDxDesdeGAlenHos oRsDx, oConexion, ml_idAtencion, ml_IdTipoServicio, lcDxPrincipal, _
                                                     lcDxPrincipalNro, lnUno, True
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
                'oRsDx.Fields!DxNro = oRsTmp1.Fields!dxNumero
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
Sub CargaDatosDeTriajeVacunas(oConexionExterna As Connection, lbDesdeGalenHos As Boolean)
    If lbDesdeGalenHos = True Then
        Dim lcTxtSpPeso As String, lcTxtSPtalla As String, lcTxtSPpa As String, lcTxtObservaciones As String
        mo_ReglasSISgalenhos.FuaCargaTriajeVacunasDesdeGAlenHos lcTxtSpPeso, lcTxtSPtalla, lcTxtSPpa, lcTxtObservaciones, _
                                                                ml_idAtencion, oConexionExterna
        txtSPpeso.Text = lcTxtSpPeso
        txtSPtalla.Text = lcTxtSPtalla
'        txtSPpa.Text = lcTxtSPpa
        txtSPpa.Text = ColocarFormatoPresAtmosferica(lcTxtSPpa) 'Modificado 19092014
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
    Else
       Dim oRsTmp1 As New Recordset, lcSistolica As String, lcPresion As String
       Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaAtencionSMIxIdCuentaAtencion(ml_IdCuentaAtencion)
       If oRsTmp1.RecordCount > 0 Then
           txtSPpa.Text = sighentidades.PresionDevuelveVacia
           lcSistolica = "___"
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              CargaVacunaYsp oRsTmp1.Fields!IntervencionesPreventivas, oRsTmp1.Fields!Valor, lcSistolica
'              If oRsTmp1.Fields!IntervencionesPreventivas = "300" Then
'                 txtSPcpn.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "003" Then
'                 txtSPpeso.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "004" Then
'                 txtSPtalla.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "005" Then
'                 txtSPedadG.Text = oRsTmp1.Fields!Valor
'              End If
'
'              If oRsTmp1.Fields!IntervencionesPreventivas = "304" Then
'                 txtSPedadGrn.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "305" Then
'                 txtSPapgar1.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "306" Then
'                 txtSPapgar5.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "010" Then
'                 txtSPalturaU.Text = oRsTmp1.Fields!Valor
'              End If
'              '
'              If oRsTmp1.Fields!IntervencionesPreventivas = "901" Then
'                 lcSistolica = Trim(oRsTmp1.Fields!Valor)
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "301" Then
'                 txtSPpa.Text = sighentidades.PresionJuntaSistolicaDiastolica(lcSistolica, oRsTmp1.Fields!Valor)
'              End If
'              '
'              If oRsTmp1.Fields!IntervencionesPreventivas = "120" Then
'                 txtSPcred.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "209" Then
'                 txtSPpuerperio.Text = oRsTmp1.Fields!Valor
'              End If
'
'              If oRsTmp1.Fields!IntervencionesPreventivas = "307" Then
'                 If Val(oRsTmp1.Fields!Valor) = 1 Then
'                    chkSPconsejeriaNsi.Value = 1
'                 Else
'                    chkSPconsejeriaNno.Value = 1
'                 End If
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "308" Then
'                 If Val(oRsTmp1.Fields!Valor) = 1 Then
'                    chkSPconsejeriaPPffSI.Value = 1
'                 Else
'                    chkSPconsejeriaPPffNO.Value = 1
'                 End If
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "309" Then
'                 If Val(oRsTmp1.Fields!Valor) = 1 Then
'                    chkSPprofilaxisOsi.Value = 1
'                 Else
'                    chkSPprofilaxisOno.Value = 1
'                 End If
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "311" Then
'                 If Val(oRsTmp1.Fields!Valor) = 1 Then
'                    chkSPvitaminaKsi.Value = 1
'                 Else
'                    chkSPvitaminaKno.Value = 1
'                 End If
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "312" Then
'                 If Val(oRsTmp1.Fields!Valor) = 1 Then
'                    chkSPeedpSI.Value = 1
'                 Else
'                    chkSPeedpNO.Value = 1
'                 End If
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "302" Then
'                 If Val(oRsTmp1.Fields!Valor) = 1 Then
'                    chkSPsicoprofilaxisSI.Value = 1
'                 Else
'                    chkSPsicoprofilaxisNO.Value = 1
'                 End If
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "303" Then
'                 If Val(oRsTmp1.Fields!Valor) = 1 Then
'                    chkSPadmOxitocinaSI.Value = 1
'                 Else
'                    chkSPadmOxitocinaNO.Value = 1
'                 End If
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "002" Then
'                 If Val(oRsTmp1.Fields!Valor) = 1 Then
'                    chkSPlactanciaMsi.Value = 1
'                 Else
'                    chkSPlactanciaMno.Value = 1
'                 End If
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "310" Then
'                 If Val(oRsTmp1.Fields!Valor) = 1 Then
'                    chkSPsuplNsi.Value = 1
'                 Else
'                    chkSPsuplNno.Value = 1
'                 End If
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "102" Then
'                 txtVacBcg.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "318" Then
'                 txtVacInfluenz.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "211" Then
'                 txtVacAntiamarilica.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "117" Then
'                 txtVacDpt.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "121" Then
'                 txtVacParotid.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "126" Then
'                 txtVacAntineumoc.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "313" Then
'                 txtVacApo.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "122" Then
'                 txtVacRubeola.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "208" Then
'                 txtVacAntitetanica.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "314" Then
'                 txtVacAsa.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "127" Then
'                 txtVacRotavirus.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "119" Then
'                 txtVacEpatB.Text = oRsTmp1.Fields!Valor
'              End If
'
'
'              If oRsTmp1.Fields!IntervencionesPreventivas = "125" Then
'                 txtVacSpr.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "007" Then
'                 txtVacDt.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "315" Then
'                 txtVacHVB.Text = oRsTmp1.Fields!Valor
'              End If
'              If oRsTmp1.Fields!IntervencionesPreventivas = "124" Then
'                 txtVacPentaval.Text = oRsTmp1.Fields!Valor
'              End If
              oRsTmp1.MoveNext
              
           Loop
        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
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
              If Not IsNull(oRsAfiliadosSIS.Fields!lot_idComponente) Then
                 AsignaComponente oRsAfiliadosSIS.Fields!lot_idComponente
              End If
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
    Case sghConsultar
        Me.Caption = "Consultar FUA " & lcOpcion & " " & sighentidades.TipoServicioDevuelveNombreCorto(ml_IdTipoServicio)
    Case sghEliminar
        Me.Caption = "Eliminar FUA " & lcOpcion & " " & sighentidades.TipoServicioDevuelveNombreCorto(ml_IdTipoServicio)
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
        Set mo_cmbConceptoP.RowSource = mo_ReglasSISgalenhos.SisConceptoPrestacionalSeleccionarTodos(False)
        '
        Set mo_cmbTipoDocumento.MiComboBox = cmbTipoDocumento
        mo_cmbTipoDocumento.BoundColumn = "ide_idTipoDocumento"
        mo_cmbTipoDocumento.ListField = "ide_descripcion"
        Set mo_cmbTipoDocumento.RowSource = mo_ReglasSISgalenhos.SisTiposDocumentosSeleccionarTodos
        '
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

        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
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





Private Sub txtFparto_LostFocus()
    If Not EsFecha(txtFparto.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
        txtFparto.Text = sighentidades.FECHA_VACIA_DMY
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





Private Sub txtReconsideracion_LostFocus()
   On Error Resume Next
   chkCsubsidiado.SetFocus
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



Private Sub txtVacEpatB_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVacEpatB
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtVacEpatB_KeyPress(KeyAscii As Integer)
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

Sub CargaConsumosEnServiciosIntermedios(oConexion As Connection, lbDesdeGalenHos As Boolean)
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oRsTmp4 As New Recordset
    Dim lnRecetado As Long, lnIdPuntoCarga As Long, lcPuntoCarga As String
    If lbDesdeGalenHos = True Then
        mo_ReglasSISgalenhos.FuaCargaSIDesdeGAlenHos oRsFarmacia, oRsPatologia, mo_lnIdTablaLISTBARITEMS, ml_IdCuentaAtencion, _
                             lcInsumo, lcMedicamento, lcDxPrincipal, lcDxPrincipalNro, lcOtros, lcLaboratorio, lcImagenes
    Else
        mo_ReglasSISgalenhos.FuaActualizaDespachosEnFarmacia ml_IdCuentaAtencion, wxParametro302, ml_IdTipoServicio, sghFuenteFinanciamiento.sghFFSIS
        mo_ReglasSISgalenhos.FuaActualizaDespachosEnServicios ml_IdCuentaAtencion, wxParametro302, ml_IdTipoServicio, sghFuenteFinanciamiento.sghFFSIS
        '*********************Farmacia - Medicamentos - Desde el SIS
        Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaAtencionMEDxIdCuentaAtencion(ml_IdCuentaAtencion)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
                 Set oRsTmp4 = mo_AdminServiciosComunes.MedicamentosInsumosSeleccionarPorCodigo(oRsTmp1.Fields!codigo)
                 If oRsTmp4.RecordCount > 0 Then 'debb2014b
                    oRsFarmacia.AddNew
                    oRsFarmacia.Fields!id = oRsTmp1.Fields!id
                    oRsFarmacia.Fields!tipo = lcMedicamento
                    oRsFarmacia.Fields!MedicInsumo = IIf(IsNull(oRsTmp4.Fields!nombre), "", oRsTmp4.Fields!nombre)
                    oRsFarmacia.Fields!recetado = oRsTmp1.Fields!CantidadPrescrita
                    oRsFarmacia.Fields!cantidad = oRsTmp1.Fields!CantidadEntregada
                    oRsFarmacia.Fields!dx = lcDxPrincipal
                    oRsFarmacia.Fields!Precio = oRsTmp1.Fields!PrecioUnitario
                    oRsFarmacia.Fields!codigo = oRsTmp1.Fields!codigo
                    oRsFarmacia.Fields!dxNro = lcDxPrincipalNro
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
                 Set oRsTmp4 = mo_AdminServiciosComunes.MedicamentosInsumosSeleccionarPorCodigo(oRsTmp1.Fields!codigo)
                 oRsFarmacia.AddNew
                 oRsFarmacia.Fields!id = oRsTmp1.Fields!id
                 oRsFarmacia.Fields!tipo = lcInsumo
                 oRsFarmacia.Fields!MedicInsumo = IIf(IsNull(oRsTmp4.Fields!nombre), "", oRsTmp4.Fields!nombre)
                 oRsFarmacia.Fields!recetado = oRsTmp1.Fields!CantidadPrescrita
                 oRsFarmacia.Fields!cantidad = oRsTmp1.Fields!CantidadEntregada
                 oRsFarmacia.Fields!dx = lcDxPrincipal
                 oRsFarmacia.Fields!Precio = oRsTmp1.Fields!PrecioUnitario
                 oRsFarmacia.Fields!codigo = oRsTmp1.Fields!codigo
                 oRsFarmacia.Fields!dxNro = lcDxPrincipalNro
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
                    oRsPatologia.Fields!dx = lcDxPrincipal
                    oRsPatologia.Fields!Precio = oRsTmp1.Fields!PrecioUnitario
                    oRsPatologia.Fields!idPuntoCarga = lnIdPuntoCarga
                    oRsPatologia.Fields!codigo = oRsTmp1.Fields!codigo
                    oRsPatologia.Fields!dxNro = lcDxPrincipalNro
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
                        oRsFarmacia.Update
                    End If
                    oRsTmp1.MoveNext
               Loop
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
            'mgaray
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
                'Abre el archivo ExcelOpenOffice
                lcArchivoExcel = App.Path + "\Plantillas\SisFua.ods"
        '        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
        '        Chemin = "file:///" & App.Path & "\Plantillas\"
        '        Chemin = Replace(Chemin, "\", "/")
        '        Fichier = Chemin & "/OpenOffice.ods"
                Fichier = Format(Time, "hhmmss") & ".ods"
                PathFileOpenOffice = App.Path + "\Plantillas\" & Fichier
                'FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
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
                'Encabezado de Pagina
'                mo_CabeceraReportes.CabeceraReportes Document, True
                ' Pone la ventana en primer plano, pasándole el Hwnd
                ret = SetForegroundWindow(lnHwnd)
            Else
                'Crea nueva hoja
                Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                Set oWorkBook = oExcel.Workbooks.Add
                'Abre, copia y cierra la plantilla
                Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\SisFua.xls")
                oWorkBookPlantilla.Worksheets("SisFua").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
'                mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            End If
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(92, 0).setFormula("F.Emisión: " & lcBuscaParametro.RetornaFechaHoraServidorSQL)
                Call Feuille.getcellbyposition(97, 1).setFormula("Cta: " & ml_IdCuentaAtencion & " " & " " & sighentidades.TipoServicioDevuelveNombreCorto(ml_IdTipoServicio))
                Call Feuille.getcellbyposition(36, 3).setFormula(Trim(txtFua1.Text))
                Call Feuille.getcellbyposition(50, 3).setFormula(Trim(txtFua2.Text))
                Call Feuille.getcellbyposition(64, 3).setFormula(Trim(txtFua3.Text))
                Call Feuille.getcellbyposition(2, 6).setFormula(Trim(txtCScodigo.Text))
                Call Feuille.getcellbyposition(25, 6).setFormula(Trim(txtCS.Text))
                Call Feuille.getcellbyposition(98, 7).setFormula(Trim(txtReconsideracion.Text))
                Call Feuille.getcellbyposition(2, 15).setFormula(ml_ApellidoPaterno)
                Call Feuille.getcellbyposition(64, 15).setFormula(ml_ApellidoMaterno)
                Call Feuille.getcellbyposition(2, 18).setFormula(ml_PrimerNombre)
                Call Feuille.getcellbyposition(64, 18).setFormula(ml_SegundoNombre)
                Call Feuille.getcellbyposition(17, 10).setFormula(IIf(chkCsubsidiado.Value <> ssCBUnchecked, "X", ""))
                Call Feuille.getcellbyposition(17, 11).setFormula(IIf(chkCSemiS.Value <> ssCBUnchecked, "X", ""))
                Call Feuille.getcellbyposition(37, 10).setFormula(IIf(chkTAnuevo.Value <> ssCBUnchecked, "X", ""))
                Call Feuille.getcellbyposition(37, 11).setFormula(IIf(chkTAAntiguoI.Value <> ssCBUnchecked, "X", ""))
                Call Feuille.getcellbyposition(37, 12).setFormula(IIf(chkTAantiguoA.Value <> ssCBUnchecked, "X", ""))
                Call Feuille.getcellbyposition(40, 11).setFormula(Trim(txtNroAfiliacion1.Text))
                Call Feuille.getcellbyposition(47, 11).setFormula(Trim(txtNroAfiliacion2.Text))
                Call Feuille.getcellbyposition(51, 11).setFormula(Trim(txtNroAfiliacion3.Text))
                Call Feuille.getcellbyposition(64, 11).setFormula(Trim(Left(Me.cmbTipoDocumento.Text, 8)))
                Call Feuille.getcellbyposition(72, 11).setFormula(Trim(txtNdocumento.Text))
                Call Feuille.getcellbyposition(100, 10).setFormula(txtInstitucion.Text)
                Call Feuille.getcellbyposition(100, 11).setFormula(txtCodSeguro.Text)
                
                Call Feuille.getcellbyposition(2, 22).setFormula("'" & txtFnacimiento.Text)
                Call Feuille.getcellbyposition(27, 21).setFormula(IIf(Trim(txtSexo.Text) = "Masculino", "X", ""))
                Call Feuille.getcellbyposition(27, 22).setFormula(IIf(Trim(txtSexo.Text) = "Femenino", "X", ""))
                Call Feuille.getcellbyposition(31, 21).setFormula(IIf(chkAtencionAmbulatoria.Value <> ssCBUnchecked, chkAtencionAmbulatoria.Caption, IIf(chkAtencionReferencia.Value <> ssCBUnchecked, chkAtencionReferencia.Caption, IIf(chkAtencionEmergencia.Value <> ssCBUnchecked, chkAtencionEmergencia.Caption, ""))))
                Call Feuille.getcellbyposition(57, 21).setFormula(IIf(chkGestante.Value <> ssCBUnchecked, lcEquix, ""))
                Call Feuille.getcellbyposition(57, 22).setFormula(IIf(chkPuerpera.Value <> ssCBUnchecked, lcEquix, ""))
                
                Call Feuille.getcellbyposition(2, 26).setFormula(CStr("'" & Format(txtFantencion.Text, "dd/mm/yyyy")))
                Call Feuille.getcellbyposition(18, 25).setFormula(CStr("'" & Format(txtHatencion.Text, "hh:mm")))
                Call Feuille.getcellbyposition(43, 25).setFormula(IIf(chkIntramural.Value <> ssCBUnchecked, "X", ""))
                Call Feuille.getcellbyposition(43, 26).setFormula(IIf(chkExtramural.Value <> ssCBUnchecked, "X", ""))
                Call Feuille.getcellbyposition(47, 25).setFormula(txtNhistoriaClinica.Text)
                
                Call Feuille.getcellbyposition(112, 26).setFormula("'" & txtFparto.Text)
            Else
                oWorkSheet.Cells(1, 93).Value = "F.Emisión: " & lcBuscaParametro.RetornaFechaHoraServidorSQL
                oWorkSheet.Cells(2, 98).Value = "Cta: " & ml_IdCuentaAtencion & " " & " " & sighentidades.TipoServicioDevuelveNombreCorto(ml_IdTipoServicio)
                oWorkSheet.Cells(4, 37).Value = Trim(txtFua1.Text)
                oWorkSheet.Cells(4, 51).Value = Trim(txtFua2.Text)
                oWorkSheet.Cells(4, 65).Value = Trim(txtFua3.Text)
                oWorkSheet.Cells(7, 3).Value = Trim(txtCScodigo.Text)
                oWorkSheet.Cells(7, 26).Value = Trim(txtCS.Text)
                oWorkSheet.Cells(8, 99).Value = txtReconsideracion.Text
                oWorkSheet.Cells(16, 3).Value = ml_ApellidoPaterno
                oWorkSheet.Cells(16, 65).Value = ml_ApellidoMaterno
                oWorkSheet.Cells(19, 3).Value = ml_PrimerNombre
                oWorkSheet.Cells(19, 65).Value = ml_SegundoNombre
                oWorkSheet.Cells(11, 18).Value = IIf(chkCsubsidiado.Value <> ssCBUnchecked, "X", "")
                oWorkSheet.Cells(12, 18).Value = IIf(chkCSemiS.Value <> ssCBUnchecked, "X", "")
                oWorkSheet.Cells(11, 38).Value = IIf(chkTAnuevo.Value <> ssCBUnchecked, "X", "")
                oWorkSheet.Cells(12, 38).Value = IIf(chkTAAntiguoI.Value <> ssCBUnchecked, "X", "")
                oWorkSheet.Cells(13, 38).Value = IIf(chkTAantiguoA.Value <> ssCBUnchecked, "X", "")
                oWorkSheet.Cells(12, 41).Value = Trim(txtNroAfiliacion1.Text)
                oWorkSheet.Cells(12, 48).Value = Trim(txtNroAfiliacion2.Text)
                oWorkSheet.Cells(12, 52).Value = Trim(txtNroAfiliacion3.Text)
                oWorkSheet.Cells(12, 65).Value = Trim(Left(Me.cmbTipoDocumento.Text, 8))
                oWorkSheet.Cells(12, 73).Value = txtNdocumento.Text
                oWorkSheet.Cells(11, 101).Value = txtInstitucion.Text
                oWorkSheet.Cells(12, 101).Value = txtCodSeguro.Text
                
                oWorkSheet.Cells(23, 3).Value = "'" & txtFnacimiento.Text
                oWorkSheet.Cells(22, 28).Value = IIf(Trim(txtSexo.Text) = "Masculino", "X", "")
                oWorkSheet.Cells(23, 28).Value = IIf(Trim(txtSexo.Text) = "Femenino", "X", "")
                oWorkSheet.Cells(22, 32).Value = IIf(chkAtencionAmbulatoria.Value <> ssCBUnchecked, chkAtencionAmbulatoria.Caption, IIf(chkAtencionReferencia.Value <> ssCBUnchecked, chkAtencionReferencia.Caption, IIf(chkAtencionEmergencia.Value <> ssCBUnchecked, chkAtencionEmergencia.Caption, "")))
                oWorkSheet.Cells(22, 58).Value = IIf(chkGestante.Value <> ssCBUnchecked, lcEquix, "")
                oWorkSheet.Cells(23, 58).Value = IIf(chkPuerpera.Value <> ssCBUnchecked, lcEquix, "")
                
                oWorkSheet.Cells(27, 3).Value = CStr("'" & Format(txtFantencion.Text, "dd/mm/yyyy"))
                oWorkSheet.Cells(26, 19).Value = CStr("'" & Format(txtHatencion.Text, "hh:mm"))
                oWorkSheet.Cells(26, 44).Value = IIf(chkIntramural.Value <> ssCBUnchecked, "X", "")
                oWorkSheet.Cells(27, 44).Value = IIf(chkExtramural.Value <> ssCBUnchecked, "X", "")
                oWorkSheet.Cells(26, 48).Value = txtNhistoriaClinica.Text
                
                oWorkSheet.Cells(27, 113).Value = "'" & txtFparto.Text
            End If
            lcSql = lcEquix
            Select Case mo_cmbConceptoP.BoundText
            Case 1    'Atención Directa
                 iFila = 22
            Case 2    'Enfermedad Alto Costo (No LPIS)
                 iFila = 23
            Case 3    'Caso Especial
                 iFila = 24
            Case 4    'Sepelio
                 iFila = 25
            Case 5    'Traslado
                 iFila = 26
            Case Else
                 iFila = 22
                 lcSql = ""
            End Select
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(92, iFila - 1).setFormula(lcSql)
            Else
                oWorkSheet.Cells(iFila, 93).Value = lcSql
            End If
            If txtNautorizacion.Text <> "" Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(95, iFila - 1).setFormula(txtNautorizacion.Text)
                Else
                    oWorkSheet.Cells(iFila, 96).Value = txtNautorizacion.Text
                End If
            End If
            If Val(txtMonto.Text) > 0 Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(117, iFila - 1).setFormula(txtMonto.Text)
                Else
                    oWorkSheet.Cells(iFila, 118).Value = txtMonto.Text
                End If
            End If
            If lbEsOpenOffice = True Then
                ''''
                Call Feuille.getcellbyposition(24, 29).setFormula(IIf(chkPAestablecimiento.Value <> ssCBUnchecked, lcEquix, ""))
                Call Feuille.getcellbyposition(24, 30).setFormula(IIf(chkPAaisped.Value <> ssCBUnchecked, lcEquix, ""))
                Call Feuille.getcellbyposition(29, 29).setFormula("'" & Right("000" & ucSISfuaCodPrestacion1.CodigoPrestacion, 3))
                Call Feuille.getcellbyposition(48, 29).setFormula(Trim(txtROcodigo.Text))
                Call Feuille.getcellbyposition(64, 29).setFormula(txtRO.Text)
                Call Feuille.getcellbyposition(109, 29).setFormula(txtRONumero.Text)
                Call Feuille.getcellbyposition(29, 30).setFormula(ucSISfuaCodPrestacion1.Prestacion)
            Else
                '''
                oWorkSheet.Cells(30, 25).Value = IIf(chkPAestablecimiento.Value <> ssCBUnchecked, lcEquix, "") 'ojo
                oWorkSheet.Cells(31, 25).Value = IIf(chkPAaisped.Value <> ssCBUnchecked, lcEquix, "")          'ojo
                oWorkSheet.Cells(30, 30).Value = "'" & Right("000" & ucSISfuaCodPrestacion1.CodigoPrestacion, 3)                     'ojo
                oWorkSheet.Cells(30, 49).Value = Trim(txtROcodigo.Text)
                oWorkSheet.Cells(30, 65).Value = txtRO.Text
                oWorkSheet.Cells(30, 110).Value = txtRONumero.Text
                oWorkSheet.Cells(31, 30).Value = ucSISfuaCodPrestacion1.Prestacion                            'ojo
            End If
            lcSql = lcEquix
            Select Case mo_cmbIdDestinoAtencion.BoundText
            Case "1"     'alta
                 iFila = 34
                 iColumna = 9
            Case "2"     'citado
                 iFila = 34
                 iColumna = 19
            Case "3"     'Ref. Emergencia
                 iFila = 35
                 iColumna = 38
            Case "4"     'Ref. Consulta Externa
                 iFila = 35
                 iColumna = 52
            Case "5"     'Ref. Apoyo al Dx.
                 iFila = 35
                 iColumna = 74
            Case "6"     'Contrarreferido
                 iFila = 34
                 iColumna = 87
            Case "7"     'Fallecido
                 iFila = 34
                 iColumna = 100
            Case "8"     'Hospitalizado
                 iFila = 34
                 iColumna = 28
            Case Else
                 iFila = 34
                 iColumna = 9
                 lcSql = ""
            End Select
            
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(iColumna - 1, iFila - 1).setFormula(lcSql)
            Else
                oWorkSheet.Cells(iFila, iColumna).Value = lcSql
            End If
            
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(2, 37).setFormula(Trim(txtRDcodigo.Text))
                Call Feuille.getcellbyposition(19, 37).setFormula(txtRD.Text)
                Call Feuille.getcellbyposition(68, 37).setFormula(txtRDnumero.Text)
                Call Feuille.getcellbyposition(107, 34).setFormula(txtHfingreso.Text)
                Call Feuille.getcellbyposition(107, 38).setFormula("'" & txtHfalta.Text)
            Else
                oWorkSheet.Cells(38, 3).Value = Trim(txtRDcodigo.Text)
                oWorkSheet.Cells(38, 20).Value = txtRD.Text
                oWorkSheet.Cells(38, 69).Value = txtRDnumero.Text
                oWorkSheet.Cells(35, 108).Value = "'" & txtHfingreso.Text
                oWorkSheet.Cells(39, 108).Value = "'" & txtHfalta.Text
            End If
            
            If lbEsOpenOffice = True Then
'                Call Feuille.getcellbyposition(iFila + 22, 35).setFormula(txtSPpeso.Text)
                Call Feuille.getcellbyposition(16, 41).setFormula(txtSPcpn.Text)
                Call Feuille.getcellbyposition(36, 41).setFormula(txtSPpeso.Text)
                Call Feuille.getcellbyposition(54, 41).setFormula(txtSPtalla.Text)
            Else
'                oWorkSheet.Cells(36, iFila + 23).Value = txtSPpeso.Text
                oWorkSheet.Cells(42, 17).Value = txtSPcpn.Text
                oWorkSheet.Cells(42, 37).Value = txtSPpeso.Text
                oWorkSheet.Cells(42, 55).Value = txtSPtalla.Text
            End If
            If chkSPconsejeriaNsi.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(79, 41).setFormula("Si")
                    Call Feuille.getcellbyposition(79, 42).setFormula("")
                Else
                    oWorkSheet.Cells(42, 80).Value = "Si"
                    oWorkSheet.Cells(43, 80).Value = ""
                End If
            ElseIf chkSPconsejeriaNno.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(79, 41).setFormula("")
                    Call Feuille.getcellbyposition(79, 42).setFormula("No")
                Else
                    oWorkSheet.Cells(42, 80).Value = ""
                    oWorkSheet.Cells(43, 80).Value = "No"
                End If
            End If
            If chkSPconsejeriaPPffSI.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(79, 43).setFormula("Si")
                    Call Feuille.getcellbyposition(79, 44).setFormula("")
                Else
                    oWorkSheet.Cells(44, 80).Value = "Si"
                    oWorkSheet.Cells(45, 80).Value = ""
                End If
            ElseIf chkSPconsejeriaPPffNO.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(79, 43).setFormula("")
                    Call Feuille.getcellbyposition(79, 44).setFormula("No")
                Else
                    oWorkSheet.Cells(44, 80).Value = ""
                    oWorkSheet.Cells(45, 80).Value = "No"
                End If
            End If
            If chkSPprofilaxisOsi.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(79, 45).setFormula("Si")
                    Call Feuille.getcellbyposition(79, 46).setFormula("")
                Else
                    oWorkSheet.Cells(46, 80).Value = "Si"
                    oWorkSheet.Cells(47, 80).Value = ""
                End If
            ElseIf chkSPprofilaxisOno.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(79, 45).setFormula("")
                    Call Feuille.getcellbyposition(79, 46).setFormula("No")
                Else
                    oWorkSheet.Cells(46, 80).Value = ""
                    oWorkSheet.Cells(47, 80).Value = "No"
                End If
            End If
            If chkSPsuplNsi.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(79, 43).setFormula("Si")
                    Call Feuille.getcellbyposition(79, 45).setFormula("")
                Else
                    oWorkSheet.Cells(44, 80).Value = "Si"
                    oWorkSheet.Cells(46, 80).Value = ""
                End If
            ElseIf Me.chkSPsuplNno.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(79, 43).setFormula("")
                    Call Feuille.getcellbyposition(79, 45).setFormula("No")
                Else
                    oWorkSheet.Cells(44, 80).Value = ""
                    oWorkSheet.Cells(46, 80).Value = "No"
                End If
            End If
            
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(16, 42).setFormula(txtSPedadG.Text)
                Call Feuille.getcellbyposition(16, 44).setFormula(txtSPalturaU.Text)
                Call Feuille.getcellbyposition(57, 43).setFormula(txtSPapgar1.Text)
                Call Feuille.getcellbyposition(63, 43).setFormula(txtSPapgar5.Text)
            Else
                oWorkSheet.Cells(43, 17).Value = txtSPedadG.Text
                oWorkSheet.Cells(45, 17).Value = txtSPalturaU.Text
                oWorkSheet.Cells(44, 58).Value = txtSPapgar1.Text
                oWorkSheet.Cells(44, 64).Value = txtSPapgar5.Text
            End If
            'Vacunas
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(89, 41).setFormula(txtVacBcg.Text)
                Call Feuille.getcellbyposition(104, 41).setFormula(txtVacInfluenz.Text)
                Call Feuille.getcellbyposition(122, 41).setFormula(txtVacAntiamarilica.Text)
                Call Feuille.getcellbyposition(89, 42).setFormula(txtVacDpt.Text)
                Call Feuille.getcellbyposition(104, 42).setFormula(txtVacParotid.Text)
                Call Feuille.getcellbyposition(122, 42).setFormula(txtVacAntineumoc.Text)
                Call Feuille.getcellbyposition(89, 43).setFormula(txtVacApo.Text)
                Call Feuille.getcellbyposition(104, 43).setFormula(txtVacRubeola.Text)
                Call Feuille.getcellbyposition(122, 43).setFormula(txtVacAntitetanica.Text)
                Call Feuille.getcellbyposition(89, 44).setFormula(txtVacAsa.Text)
                Call Feuille.getcellbyposition(104, 44).setFormula(txtVacRotavirus.Text)
                Call Feuille.getcellbyposition(122, 44).setFormula(txtVacEpatB.Text)
                Call Feuille.getcellbyposition(89, 45).setFormula(txtVacSpr.Text)
                Call Feuille.getcellbyposition(104, 45).setFormula(txtVacDt.Text)
                Call Feuille.getcellbyposition(89, 46).setFormula(txtVacHVB.Text)
                Call Feuille.getcellbyposition(104, 46).setFormula(txtVacPentaval.Text)
            Else
                oWorkSheet.Cells(42, 90).Value = txtVacBcg.Text
                oWorkSheet.Cells(42, 105).Value = txtVacInfluenz.Text
                oWorkSheet.Cells(42, 123).Value = txtVacAntiamarilica.Text
                oWorkSheet.Cells(43, 90).Value = txtVacDpt.Text
                oWorkSheet.Cells(43, 105).Value = txtVacParotid.Text
                oWorkSheet.Cells(43, 123).Value = txtVacAntineumoc.Text
                oWorkSheet.Cells(44, 90).Value = txtVacApo.Text
                oWorkSheet.Cells(44, 105).Value = txtVacRubeola.Text
                oWorkSheet.Cells(44, 123).Value = txtVacAntitetanica.Text
                oWorkSheet.Cells(45, 90).Value = txtVacAsa.Text
                oWorkSheet.Cells(45, 105).Value = txtVacRotavirus.Text
                oWorkSheet.Cells(45, 123).Value = txtVacEpatB.Text
                oWorkSheet.Cells(46, 90).Value = txtVacSpr.Text
                oWorkSheet.Cells(46, 105).Value = txtVacDt.Text
                oWorkSheet.Cells(47, 90).Value = txtVacHVB.Text
                oWorkSheet.Cells(47, 105).Value = txtVacPentaval.Text
            End If
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(36, 43).setFormula(txtSPedadGrn.Text)
            Else
                oWorkSheet.Cells(44, 37).Value = txtSPedadGrn.Text
            End If
            If chkSPeedpSI.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(50, 46).setFormula("Si")
                    Call Feuille.getcellbyposition(50, 47).setFormula("")
                Else
                    oWorkSheet.Cells(47, 51).Value = "Si"
                    oWorkSheet.Cells(48, 51).Value = ""
                End If
            ElseIf chkSPeedpNO.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(50, 46).setFormula("")
                    Call Feuille.getcellbyposition(50, 47).setFormula("NO")
                Else
                    oWorkSheet.Cells(47, 51).Value = ""
                    oWorkSheet.Cells(48, 51).Value = "No"
                End If
            End If

            If chkSPvitaminaKsi.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(65, 46).setFormula("Si")
                    Call Feuille.getcellbyposition(65, 47).setFormula("")
                Else
                    oWorkSheet.Cells(47, 66).Value = "Si"
                    oWorkSheet.Cells(48, 66).Value = ""
                End If
            ElseIf chkSPvitaminaKno.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(65, 46).setFormula("")
                    Call Feuille.getcellbyposition(65, 47).setFormula("No")
                Else
                    oWorkSheet.Cells(47, 66).Value = ""
                    oWorkSheet.Cells(48, 66).Value = "No"
                End If
            End If
            
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(16, 46).setFormula(txtSPpa.Text)
                Call Feuille.getcellbyposition(36, 46).setFormula(txtSPcred.Text)
            Else
                oWorkSheet.Cells(47, 17).Value = txtSPpa.Text
                oWorkSheet.Cells(47, 37).Value = txtSPcred.Text
            End If
            
            If chkSPlactanciaMsi.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(36, 49).setFormula("Si")
                    Call Feuille.getcellbyposition(36, 50).setFormula("")
                Else
                    oWorkSheet.Cells(50, 37).Value = "Si"
                    oWorkSheet.Cells(51, 37).Value = ""
                End If
            ElseIf chkSPlactanciaMno.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(36, 49).setFormula("")
                    Call Feuille.getcellbyposition(36, 50).setFormula("No")
                Else
                    oWorkSheet.Cells(50, 37).Value = ""
                    oWorkSheet.Cells(51, 37).Value = "No"
                End If
            End If
            
            If chkSPsicoprofilaxisSI.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(16, 47).setFormula("Si")
                    Call Feuille.getcellbyposition(16, 48).setFormula("")
                Else
                    oWorkSheet.Cells(48, 17).Value = "Si"
                    oWorkSheet.Cells(49, 17).Value = ""
                End If
            ElseIf Me.chkSPsicoprofilaxisNO.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(16, 47).setFormula("")
                    Call Feuille.getcellbyposition(16, 48).setFormula("No")
                Else
                    oWorkSheet.Cells(48, 17).Value = ""
                    oWorkSheet.Cells(49, 17).Value = "No"
                End If
            End If
            If chkSPadmOxitocinaSI.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(16, 49).setFormula("Si")
                    Call Feuille.getcellbyposition(16, 50).setFormula("")
                Else
                    oWorkSheet.Cells(50, 17).Value = "Si"
                    oWorkSheet.Cells(51, 17).Value = ""
                End If
            ElseIf chkSPadmOxitocinaNO.Value <> ssCBUnchecked Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(16, 49).setFormula("")
                    Call Feuille.getcellbyposition(16, 50).setFormula("No")
                Else
                    oWorkSheet.Cells(50, 17).Value = ""
                    oWorkSheet.Cells(51, 17).Value = "No"
                End If
            End If
            If lbEsOpenOffice = True Then
               Call Feuille.getcellbyposition(53, 49).setFormula(txtSPpuerperio.Text)
            Else
                oWorkSheet.Cells(50, 54).Value = txtSPpuerperio.Text
            End If
            'Dx
            iFila = 55
            oRsDx.MoveFirst
            Do While Not oRsDx.EOF
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula(IIf(IsNull(oRsDx.Fields!Descripcion), "", oRsDx.Fields!Descripcion))
                    Call Feuille.getcellbyposition(89, iFila - 1).setFormula(IIf(oRsDx.Fields!DxIngresoPresuntivo = True, lcEquix, ""))
                    Call Feuille.getcellbyposition(92, iFila - 1).setFormula(IIf(oRsDx.Fields!DxIngresoDefinitivo = True, lcEquix, ""))
                    Call Feuille.getcellbyposition(95, iFila - 1).setFormula(IIf(oRsDx.Fields!DxIngresoRepetido = True, lcEquix, ""))
                    Call Feuille.getcellbyposition(98, iFila - 1).setFormula(IIf(IsNull(oRsDx.Fields!dxIngreso), "", oRsDx.Fields!dxIngreso))
                    Call Feuille.getcellbyposition(109, iFila - 1).setFormula(IIf(IsNull(oRsDx.Fields!DxEgreso), "", oRsDx.Fields!DxEgreso))
                    Call Feuille.getcellbyposition(118, iFila - 1).setFormula(IIf(oRsDx.Fields!DxEgresoDefinitivo = True, lcEquix, ""))
                    Call Feuille.getcellbyposition(122, iFila - 1).setFormula(IIf(oRsDx.Fields!DxEgresoRepetido = True, lcEquix, ""))
                Else
                    oWorkSheet.Cells(iFila, 6).Value = IIf(IsNull(oRsDx.Fields!Descripcion), "", oRsDx.Fields!Descripcion)
                    oWorkSheet.Cells(iFila, 90).Value = IIf(oRsDx.Fields!DxIngresoPresuntivo = True, lcEquix, "")
                    oWorkSheet.Cells(iFila, 93).Value = IIf(oRsDx.Fields!DxIngresoDefinitivo = True, lcEquix, "")
                    oWorkSheet.Cells(iFila, 96).Value = IIf(oRsDx.Fields!DxIngresoRepetido = True, lcEquix, "")
                    oWorkSheet.Cells(iFila, 99).Value = IIf(IsNull(oRsDx.Fields!dxIngreso), "", oRsDx.Fields!dxIngreso)
                    oWorkSheet.Cells(iFila, 110).Value = IIf(IsNull(oRsDx.Fields!DxEgreso), "", oRsDx.Fields!DxEgreso)
                    oWorkSheet.Cells(iFila, 119).Value = IIf(oRsDx.Fields!DxEgresoDefinitivo = True, lcEquix, "")
                    oWorkSheet.Cells(iFila, 123).Value = IIf(oRsDx.Fields!DxEgresoRepetido = True, lcEquix, "")
                End If
                oRsDx.MoveNext
                iFila = iFila + 1
            Loop
            'Medico
            Dim oRsTmp5 As New Recordset
            Set oRsTmp5 = mo_ReglasComunes.EmpleadosSeleccionarPorDNI(txtMedicoDni.Text)
            iFila = 63
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula(txtMedicoDni.Text)
                Call Feuille.getcellbyposition(20, iFila - 1).setFormula(txtMedico.Text)
                Call Feuille.getcellbyposition(89, iFila - 1).setFormula(txtMedicoColegiatura.Text)
                Call Feuille.getcellbyposition(109, iFila - 1).setFormula(txtMedicoEspecialidad.Text & " - " & oRsTmp5.Fields!TipoEmpleado)
            Else
                oWorkSheet.Cells(iFila, 3).Value = txtMedicoDni.Text
                oWorkSheet.Cells(iFila, 21).Value = txtMedico.Text
                oWorkSheet.Cells(iFila, 90).Value = txtMedicoColegiatura.Text
                oWorkSheet.Cells(iFila, 110).Value = txtMedicoEspecialidad.Text & " - " & oRsTmp5.Fields!TipoEmpleado
            End If
            Set oRsTmp5 = Nothing
'            iFila = 50
'            Dim oDOEspecialidades As New DOEspecialidades
'            Set oDOEspecialidades = mo_ReglasServiciosHosp.EspecialidadesSeleccionarPorId(CLng(txtMedicoEspecialidad.Text))
'            If lbEsOpenOffice = True Then
'                Call Feuille.getcellbyposition(2, iFila - 1).setFormula(txtMedicoDni.Text)
'                Call Feuille.getcellbyposition(20, iFila - 1).setFormula(txtMedico.Text)
'                Call Feuille.getcellbyposition(89, iFila - 1).setFormula(txtMedicoColegiatura.Text)
'
'                Call Feuille.getcellbyposition(109, iFila - 1).setFormula(oDOEspecialidades.IdEspecialidad & " - " & oDOEspecialidades.nombre)
'
'            Else
'                oWorkSheet.Cells(iFila, 3).Value = txtMedicoDni.Text
'                oWorkSheet.Cells(iFila, 21).Value = txtMedico.Text
'                oWorkSheet.Cells(iFila, 90).Value = txtMedicoColegiatura.Text
'                oWorkSheet.Cells(iFila, 110).Value = oDOEspecialidades.IdEspecialidad & " - " & oDOEspecialidades.nombre
'            End If

            'Se emite desde CITA, falta llenar ANEXO de Cpt y Farmacia
            lbMuestraCPTdefaults = False
            Select Case mo_lnIdTablaLISTBARITEMS
            Case sghOpcionGalenHos.sghRegistroCitaCE
               lbMuestraCPTdefaults = True
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
                         oRsPatologia.Fields!codigo = IIf(lcDescripcion = lcVacio, ".", orstmp.Fields!codigo)
                         oRsPatologia.Fields!procedimiento = Left(IIf(lcDescripcion = lcVacio, String(255, "_"), lcDescripcion & String(255, "_")), 255)
                         oRsPatologia.Fields!dx = " "
                         oRsPatologia.Fields!tipo = lcPuntoCarga
                         oRsPatologia.Fields!idPuntoCarga = lnIdPuntoCarga
                         oRsPatologia.Update
                      Else
                         oRsFarmacia.AddNew
                         oRsFarmacia.Fields!codigo = IIf(lcDescripcion = lcVacio, ".", orstmp.Fields!codigo)
                         oRsFarmacia.Fields!MedicInsumo = Left(IIf(lcDescripcion = lcVacio, String(255, "_"), lcDescripcion & String(255, "_")), 255)
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
            End If
            'calcula las Paginas totales y añade lineas vacias
            oRsFarmacia.Filter = ""
            oRsPatologia.Filter = ""
            lnTotalLineas = Round((oRsFarmacia.RecordCount + oRsPatologia.RecordCount) / 2, 0)
            lnMaximaLineaPorPagina = 73
            lnTotalPaginas = 2
            If lnTotalLineas > lnMaximaLineaPorPagina Then
                lnTotalPaginas = Round(lnTotalLineas / lnMaximaLineaPorPagina, 0) + 1
                lnFilasAinsertar = lnTotalLineas - lnMaximaLineaPorPagina
                If lbEsOpenOffice = True Then
                    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                    PrintArea(0).Sheet = 0
                    PrintArea(0).startcolumn = 0
                    PrintArea(0).StartRow = 0
                    PrintArea(0).EndColumn = 126
                    PrintArea(0).EndRow = lnTotalPaginas * lnMaximaLineaPorPagina '73   'Trim(Str(lnTotalPaginas * 55))
                    Call Feuille.SetPrintAreas(PrintArea())
                    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                Else
                    oWorkSheet.PageSetup.PrintArea = "$A$1:$DX$" & Trim(Str(lnTotalPaginas * lnMaximaLineaPorPagina))
                End If
            Else
                If lbEsOpenOffice = True Then
                    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                    PrintArea(0).Sheet = 0
                    PrintArea(0).startcolumn = 0
                    PrintArea(0).StartRow = 0
                    PrintArea(0).EndColumn = 126
                    PrintArea(0).EndRow = 2 * lnMaximaLineaPorPagina '73 'Trim(Str(2 * 55))
                    Call Feuille.SetPrintAreas(PrintArea())
                    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                Else
                    oWorkSheet.PageSetup.PrintArea = "$A$1:$DX$" & Trim(Str(2 * lnMaximaLineaPorPagina))
                End If
            End If
            'Medicamentos
            iFila = 82
            If lbEsOpenOffice = True Then
               iFila = 83
            End If
            ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = 62
            oRsFarmacia.Filter = "tipo='" & lcMedicamento & "'"
            If oRsFarmacia.RecordCount > 0 Then
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(lcMedicamento)
                    Call Feuille.getcellbyposition(43, iFila - 1).setFormula("Recetado")
                    Call Feuille.getcellbyposition(47, iFila - 1).setFormula("Entregado")
                    Call Feuille.getcellbyposition(51, iFila - 1).setFormula("Dx")
                    Call Feuille.getcellbyposition(64, iFila - 1).setFormula(lcMedicamento)
                    Call Feuille.getcellbyposition(113, iFila - 1).setFormula("Recetado")
                    Call Feuille.getcellbyposition(117, iFila - 1).setFormula("Entregado")
                    Call Feuille.getcellbyposition(121, iFila - 1).setFormula("Dx")
                Else
                    oWorkSheet.Cells(iFila, 3).Value = lcMedicamento
                    oWorkSheet.Cells(iFila, 44).Value = "Recetado"
                    oWorkSheet.Cells(iFila, 48).Value = "Entregado"
                    oWorkSheet.Cells(iFila, 52).Value = "Dx"
                    oWorkSheet.Cells(iFila, 65).Value = lcMedicamento
                    oWorkSheet.Cells(iFila, 114).Value = "Recetado"
                    oWorkSheet.Cells(iFila, 118).Value = "Entregado"
                    oWorkSheet.Cells(iFila, 122).Value = "Dx"
                End If
                If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName("C" & CStr(iFila) & ":DU" & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Else
                    ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, 125
                End If
                If lbEsOpenOffice = True Then
                    ChequeaSiHaySaltoDePaginaOpenOffice iFila
                Else
                    ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
                End If
               oRsFarmacia.MoveFirst
               lbDerecha = True
               Do While Not oRsFarmacia.EOF
                    'mgaray20140926
                  If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 64), iFila - 1).setFormula(Left(Trim(oRsFarmacia.Fields!codigo) & " " & oRsFarmacia.Fields!MedicInsumo, 65))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 43, 114), iFila - 1).setFormula("'" & String(7, "_"))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 47, 117), iFila - 1).setFormula("'" & String(7, "_"))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 51, 121), iFila - 1).setFormula("'" & String(7, "_"))
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 65)).Value = Left(Trim(oRsFarmacia.Fields!codigo) & " " & oRsFarmacia.Fields!MedicInsumo, 65)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 44, 114)).Value = "'" & String(7, "_")
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 48, 118)).Value = "'" & String(7, "_")
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 52, 122)).Value = "'" & String(7, "_")
                    End If
                  Else
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 64), iFila - 1).setFormula(Left(Trim(oRsFarmacia.Fields!codigo) & " " & oRsFarmacia.Fields!MedicInsumo, 65))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 43, 113), iFila - 1).setFormula("'" & IIf(oRsFarmacia.Fields!recetado = 0, "", oRsFarmacia.Fields!recetado))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 47, 117), iFila - 1).setFormula("'" & IIf(oRsFarmacia.Fields!cantidad = 0, "", oRsFarmacia.Fields!cantidad))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 51, 121), iFila - 1).setFormula("'" & oRsFarmacia.Fields!dx)
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 65)).Value = Left(Trim(oRsFarmacia.Fields!codigo) & " " & oRsFarmacia.Fields!MedicInsumo, 65)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 44, 114)).Value = "'" & IIf(oRsFarmacia.Fields!recetado = 0, "", oRsFarmacia.Fields!recetado)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 48, 118)).Value = "'" & IIf(oRsFarmacia.Fields!cantidad = 0, "", oRsFarmacia.Fields!cantidad)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 52, 122)).Value = "'" & oRsFarmacia.Fields!dx
                    End If
                  End If
                  oRsFarmacia.MoveNext
                  If lbDerecha = True Then
                     lbDerecha = False
                  Else
                     lbDerecha = True
                        If lbEsOpenOffice = True Then
                            ChequeaSiHaySaltoDePaginaOpenOffice iFila
                        Else
                            ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
                        End If
                  End If
               Loop
               If lbEsOpenOffice = True Then
                    ChequeaSiHaySaltoDePaginaOpenOffice iFila
               Else
                    ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
               End If
            End If
            'Insumos
            oRsFarmacia.Filter = "tipo='" & lcInsumo & "'"
            If oRsFarmacia.RecordCount > 0 Then
                iFila = iFila + 1 'FCV
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(lcInsumo)
                    Call Feuille.getcellbyposition(43, iFila - 1).setFormula("Recetado")
                    Call Feuille.getcellbyposition(47, iFila - 1).setFormula("Entregado")
                    Call Feuille.getcellbyposition(51, iFila - 1).setFormula("Dx")
                    Call Feuille.getcellbyposition(64, iFila - 1).setFormula(lcInsumo)
                    Call Feuille.getcellbyposition(113, iFila - 1).setFormula("Recetado")
                    Call Feuille.getcellbyposition(117, iFila - 1).setFormula("Entregado")
                    Call Feuille.getcellbyposition(121, iFila - 1).setFormula("Dx")
                Else
                    oWorkSheet.Cells(iFila, 3).Value = lcInsumo
                    oWorkSheet.Cells(iFila, 44).Value = "Recetado"
                    oWorkSheet.Cells(iFila, 48).Value = "Entregado"
                    oWorkSheet.Cells(iFila, 52).Value = "Dx"
                    oWorkSheet.Cells(iFila, 65).Value = lcInsumo
                    oWorkSheet.Cells(iFila, 114).Value = "Recetado"
                    oWorkSheet.Cells(iFila, 118).Value = "Entregado"
                    oWorkSheet.Cells(iFila, 122).Value = "Dx"
                End If
                If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName("C" & CStr(iFila) & ":DU" & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Else
                    ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, 125
                End If
                If lbEsOpenOffice = True Then
                    ChequeaSiHaySaltoDePaginaOpenOffice iFila
                Else
                    ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
                End If
               oRsFarmacia.MoveFirst
               lbDerecha = True
               Do While Not oRsFarmacia.EOF
                    'mgaray20140926
                  If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 64), iFila - 1).setFormula(Left(Trim(oRsFarmacia.Fields!codigo) & " " & oRsFarmacia.Fields!MedicInsumo, 65))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 43, 113), iFila - 1).setFormula("'" & String(7, "_"))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 47, 117), iFila - 1).setFormula("'" & String(7, "_"))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 51, 121), iFila - 1).setFormula("'" & String(7, "_"))
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 65)).Value = Left(Trim(oRsFarmacia.Fields!codigo) & " " & oRsFarmacia.Fields!MedicInsumo, 65)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 44, 114)).Value = "'" & String(7, "_")
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 48, 118)).Value = "'" & String(7, "_")
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 52, 122)).Value = "'" & String(7, "_")
                    End If
                  Else
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 65), iFila - 1).setFormula(Left(Trim(oRsFarmacia.Fields!codigo) & " " & oRsFarmacia.Fields!MedicInsumo, 65))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 43, 113), iFila - 1).setFormula("'" & IIf(oRsFarmacia.Fields!recetado = 0, "", oRsFarmacia.Fields!recetado))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 47, 117), iFila - 1).setFormula("'" & IIf(oRsFarmacia.Fields!cantidad = 0, "", oRsFarmacia.Fields!cantidad))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 51, 121), iFila - 1).setFormula("'" & oRsFarmacia.Fields!dx)
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 65)).Value = Left(Trim(oRsFarmacia.Fields!codigo) & " " & oRsFarmacia.Fields!MedicInsumo, 65)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 44, 114)).Value = "'" & IIf(oRsFarmacia.Fields!recetado = 0, "", oRsFarmacia.Fields!recetado)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 48, 118)).Value = "'" & IIf(oRsFarmacia.Fields!cantidad = 0, "", oRsFarmacia.Fields!cantidad)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 52, 122)).Value = "'" & oRsFarmacia.Fields!dx
                    End If
                  End If
                  oRsFarmacia.MoveNext
                  If lbDerecha = True Then
                     lbDerecha = False
                  Else
                     lbDerecha = True
                        If lbEsOpenOffice = True Then
                            ChequeaSiHaySaltoDePaginaOpenOffice iFila
                        Else
                            ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
                        End If
                  End If
               Loop
                If lbEsOpenOffice = True Then
                    ChequeaSiHaySaltoDePaginaOpenOffice iFila
                Else
                    ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
                End If
            End If
            
            'Laboratorio
            oRsPatologia.Filter = "tipo='" & lcLaboratorio & "'"
            If oRsPatologia.RecordCount > 0 Then
                iFila = iFila + 1 'FCV
                If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName("C" & CStr(iFila) & ":DU" & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(lcLaboratorio)
                    Call Feuille.getcellbyposition(43, iFila - 1).setFormula("Indicado")
                    Call Feuille.getcellbyposition(47, iFila - 1).setFormula("Ejecutado")
                    Call Feuille.getcellbyposition(51, iFila - 1).setFormula("Dx")
                    Call Feuille.getcellbyposition(64, iFila - 1).setFormula(lcLaboratorio)
                    Call Feuille.getcellbyposition(113, iFila - 1).setFormula("Indicado")
                    Call Feuille.getcellbyposition(117, iFila - 1).setFormula("Ejecutado")
                    Call Feuille.getcellbyposition(121, iFila - 1).setFormula("Dx")
                Else
                    ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, 125
                    oWorkSheet.Cells(iFila, 3).Value = lcLaboratorio
                    oWorkSheet.Cells(iFila, 44).Value = "Indicado"
                    oWorkSheet.Cells(iFila, 48).Value = "Ejecutado"
                    oWorkSheet.Cells(iFila, 52).Value = "Dx"
                    oWorkSheet.Cells(iFila, 65).Value = lcLaboratorio
                    oWorkSheet.Cells(iFila, 114).Value = "Indicado"
                    oWorkSheet.Cells(iFila, 118).Value = "Ejecutado"
                    oWorkSheet.Cells(iFila, 122).Value = "Dx"
                End If
                If lbEsOpenOffice = True Then
                    ChequeaSiHaySaltoDePaginaOpenOffice iFila
                Else
                    ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
                End If
               oRsPatologia.MoveFirst
               lbDerecha = True
               Do While Not oRsPatologia.EOF
                    'mgaray20140926
                  If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 64), iFila - 1).setFormula(Left(Trim(oRsPatologia.Fields!codigo) & " " & oRsPatologia.Fields!procedimiento, 65))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 43, 113), iFila - 1).setFormula("'" & String(7, "_"))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 47, 117), iFila - 1).setFormula("'" & String(7, "_"))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 51, 121), iFila - 1).setFormula("'" & String(7, "_"))
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 65)).Value = Left(Trim(oRsPatologia.Fields!codigo) & " " & oRsPatologia.Fields!procedimiento, 65)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 44, 114)).Value = "'" & String(7, "_")
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 48, 118)).Value = "'" & String(7, "_")
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 52, 122)).Value = "'" & String(7, "_")
                    End If
                  Else
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 64), iFila - 1).setFormula(Left(Trim(oRsPatologia.Fields!codigo) & " " & oRsPatologia.Fields!procedimiento, 65))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 43, 113), iFila - 1).setFormula("'" & IIf(oRsPatologia.Fields!indicado = 0, "", oRsPatologia.Fields!indicado))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 47, 117), iFila - 1).setFormula("'" & IIf(oRsPatologia.Fields!ejecutado = 0, "", oRsPatologia.Fields!ejecutado))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 51, 121), iFila - 1).setFormula("'" & oRsPatologia.Fields!dx)
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 65)).Value = Left(Trim(oRsPatologia.Fields!codigo) & " " & oRsPatologia.Fields!procedimiento, 65)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 44, 114)).Value = "'" & IIf(oRsPatologia.Fields!indicado = 0, "", oRsPatologia.Fields!indicado)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 48, 118)).Value = "'" & IIf(oRsPatologia.Fields!ejecutado = 0, "", oRsPatologia.Fields!ejecutado)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 52, 122)).Value = "'" & oRsPatologia.Fields!dx
                    End If
                  End If
                  oRsPatologia.MoveNext
                  If lbDerecha = True Then
                     lbDerecha = False
                  Else
                     lbDerecha = True
                        If lbEsOpenOffice = True Then
                            ChequeaSiHaySaltoDePaginaOpenOffice iFila
                        Else
                            ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
                        End If
                  End If
               Loop
                If lbEsOpenOffice = True Then
                    ChequeaSiHaySaltoDePaginaOpenOffice iFila
                Else
                    ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
                End If
            End If
            
            'Imágenes
            oRsPatologia.Filter = "tipo='" & lcImagenes & "'"
            If oRsPatologia.RecordCount > 0 Then
                iFila = iFila + 1 'FCV
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(lcImagenes)
                    Call Feuille.getcellbyposition(43, iFila - 1).setFormula("Indicado")
                    Call Feuille.getcellbyposition(47, iFila - 1).setFormula("Ejecutado")
                    Call Feuille.getcellbyposition(51, iFila - 1).setFormula("Dx")
                    Call Feuille.getcellbyposition(64, iFila - 1).setFormula(lcImagenes)
                    Call Feuille.getcellbyposition(113, iFila - 1).setFormula("Indicado")
                    Call Feuille.getcellbyposition(117, iFila - 1).setFormula("Ejecutado")
                    Call Feuille.getcellbyposition(121, iFila - 1).setFormula("Dx")
                Else
                    oWorkSheet.Cells(iFila, 3).Value = lcImagenes
                    oWorkSheet.Cells(iFila, 44).Value = "Indicado"
                    oWorkSheet.Cells(iFila, 48).Value = "Ejecutado"
                    oWorkSheet.Cells(iFila, 52).Value = "Dx"
                    oWorkSheet.Cells(iFila, 65).Value = lcImagenes
                    oWorkSheet.Cells(iFila, 114).Value = "Indicado"
                    oWorkSheet.Cells(iFila, 118).Value = "Ejecutado"
                    oWorkSheet.Cells(iFila, 122).Value = "Dx"
                End If
                 If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName("C" & CStr(iFila) & ":DU" & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Else
                    ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, 125
                End If
                If lbEsOpenOffice = True Then
                    ChequeaSiHaySaltoDePaginaOpenOffice iFila
                Else
                    ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
                End If
               oRsPatologia.MoveFirst
               lbDerecha = True
               Do While Not oRsPatologia.EOF
                    'mgaray20140926
                  If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 64), iFila - 1).setFormula(Left(Trim(oRsPatologia.Fields!codigo) & " " & oRsPatologia.Fields!procedimiento, 65))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 43, 113), iFila - 1).setFormula("'" & String(7, "_"))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 47, 117), iFila - 1).setFormula("'" & String(7, "_"))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 51, 121), iFila - 1).setFormula("'" & String(7, "_"))
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 65)).Value = Left(Trim(oRsPatologia.Fields!codigo) & " " & oRsPatologia.Fields!procedimiento, 65)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 44, 114)).Value = "'" & String(7, "_")
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 48, 118)).Value = "'" & String(7, "_")
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 52, 122)).Value = "'" & String(7, "_")
                    End If
                  Else
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 64), iFila - 1).setFormula(Left(Trim(oRsPatologia.Fields!codigo) & " " & oRsPatologia.Fields!procedimiento, 65))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 43, 113), iFila - 1).setFormula("'" & IIf(oRsPatologia.Fields!indicado = 0, "", oRsPatologia.Fields!indicado))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 47, 117), iFila - 1).setFormula("'" & IIf(oRsPatologia.Fields!ejecutado = 0, "", oRsPatologia.Fields!ejecutado))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 51, 121), iFila - 1).setFormula("'" & oRsPatologia.Fields!dx)
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 65)).Value = Left(Trim(oRsPatologia.Fields!codigo) & " " & oRsPatologia.Fields!procedimiento, 65)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 44, 114)).Value = "'" & IIf(oRsPatologia.Fields!indicado = 0, "", oRsPatologia.Fields!indicado)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 48, 118)).Value = "'" & IIf(oRsPatologia.Fields!ejecutado = 0, "", oRsPatologia.Fields!ejecutado)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 52, 122)).Value = "'" & oRsPatologia.Fields!dx
                    End If
                  End If
                  oRsPatologia.MoveNext
                  If lbDerecha = True Then
                     lbDerecha = False
                  Else
                     lbDerecha = True
                        If lbEsOpenOffice = True Then
                            ChequeaSiHaySaltoDePaginaOpenOffice iFila
                        Else
                            ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
                        End If
                  End If
               Loop
                If lbEsOpenOffice = True Then
                    ChequeaSiHaySaltoDePaginaOpenOffice iFila
                Else
                    ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
                End If
            End If
            
            'Otros CPT
            oRsPatologia.Filter = "tipo='" & lcOtros & "'"
            If oRsPatologia.RecordCount > 0 Then
                iFila = iFila + 1 'FCV
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(lcOtros)
                    Call Feuille.getcellbyposition(43, iFila - 1).setFormula("Indicado")
                    Call Feuille.getcellbyposition(47, iFila - 1).setFormula("Ejecutado")
                    Call Feuille.getcellbyposition(51, iFila - 1).setFormula("Dx")
                    Call Feuille.getcellbyposition(64, iFila - 1).setFormula(lcOtros)
                    Call Feuille.getcellbyposition(113, iFila - 1).setFormula("Indicado")
                    Call Feuille.getcellbyposition(117, iFila - 1).setFormula("Ejecutado")
                    Call Feuille.getcellbyposition(121, iFila - 1).setFormula("Dx")
                Else
                    oWorkSheet.Cells(iFila, 3).Value = lcOtros
                    oWorkSheet.Cells(iFila, 44).Value = "Indicado"
                    oWorkSheet.Cells(iFila, 48).Value = "Ejecutado"
                    oWorkSheet.Cells(iFila, 52).Value = "Dx"
                    oWorkSheet.Cells(iFila, 65).Value = lcOtros
                    oWorkSheet.Cells(iFila, 114).Value = "Indicado"
                    oWorkSheet.Cells(iFila, 118).Value = "Ejecutado"
                    oWorkSheet.Cells(iFila, 122).Value = "Dx"
                End If
                If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName("C" & CStr(iFila) & ":DU" & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Else
                    ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, 125
                End If
                If lbEsOpenOffice = True Then
                    ChequeaSiHaySaltoDePaginaOpenOffice iFila
                Else
                    ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
                End If
               oRsPatologia.MoveFirst
               lbDerecha = True
               Do While Not oRsPatologia.EOF
                    'mgaray20140926
                  If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 64), iFila - 1).setFormula(Left(Trim(oRsPatologia.Fields!codigo) & " " & oRsPatologia.Fields!procedimiento, 65))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 43, 113), iFila - 1).setFormula("'" & String(7, "_"))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 47, 117), iFila - 1).setFormula("'" & String(7, "_"))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 51, 121), iFila - 1).setFormula("'" & String(7, "_"))
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 65)).Value = Left(Trim(oRsPatologia.Fields!codigo) & " " & oRsPatologia.Fields!procedimiento, 65)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 44, 114)).Value = "'" & String(7, "_")
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 48, 118)).Value = "'" & String(7, "_")
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 52, 122)).Value = "'" & String(7, "_")
                    End If
                  Else
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 2, 64), iFila - 1).setFormula(Left(Trim(oRsPatologia.Fields!codigo) & " " & oRsPatologia.Fields!procedimiento, 65))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 43, 113), iFila - 1).setFormula("'" & IIf(oRsPatologia.Fields!indicado = "0", "", oRsPatologia.Fields!indicado))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 47, 117), iFila - 1).setFormula("'" & IIf(oRsPatologia.Fields!ejecutado = 0, "", oRsPatologia.Fields!ejecutado))
                        Call Feuille.getcellbyposition(IIf(lbDerecha = True, 51, 121), iFila - 1).setFormula("'" & oRsPatologia.Fields!dx)
                    Else
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 3, 65)).Value = Left(Trim(oRsPatologia.Fields!codigo) & " " & oRsPatologia.Fields!procedimiento, 65)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 44, 114)).Value = "'" & IIf(oRsPatologia.Fields!indicado = "0", "", oRsPatologia.Fields!indicado)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 48, 118)).Value = "'" & IIf(oRsPatologia.Fields!ejecutado = 0, "", oRsPatologia.Fields!ejecutado)
                        oWorkSheet.Cells(iFila, IIf(lbDerecha = True, 52, 122)).Value = "'" & oRsPatologia.Fields!dx
                    End If
                  End If
                  oRsPatologia.MoveNext
                  If lbDerecha = True Then
                     lbDerecha = False
                  Else
                     lbDerecha = True
                        If lbEsOpenOffice = True Then
                            ChequeaSiHaySaltoDePaginaOpenOffice iFila
                        Else
                            ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
                        End If
                  End If
               Loop
               If lbEsOpenOffice = True Then
                    ChequeaSiHaySaltoDePaginaOpenOffice iFila
               Else
                    ChequeaSiHaySaltoDePagina iFila, oWorkSheet   'iFila = iFila + 1
               End If
            End If
            '
            oRsPatologia.Filter = ""
            oRsFarmacia.Filter = ""
            'Eliminar Cpts, Farmacia solo si se emitió el FUA desde CITAS
            'mgaray20140926
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
            
            'Observaciones
            If lnTotalPaginas = 2 Then
                If iFila > 98 Then
                    iFila = iFila + 2
                Else
                    iFila = 99
                End If
            Else
               iFila = iFila + 2
            End If
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("C" & CStr(iFila) & ":DU" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula("OBSERVACIONES")
            Else
                ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, 125
                oWorkSheet.Cells(iFila, 3).Value = "OBSERVACIONES"
            End If
            iFila = iFila + 1
            'Huella
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("DE" & CStr(iFila + 1) & ":DU" & CStr(iFila + 7))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula(txtObservaciones.Text)
            Else
                oWorkSheet.Range(oWorkSheet.Cells(iFila, 3), oWorkSheet.Cells(iFila + 2, 100)).Select
                With oExcel.Selection
                     .HorizontalAlignment = xlGeneral
                     .VerticalAlignment = xlBottom
                     .WrapText = True
                     .MergeCells = True
                     .Value = txtObservaciones.Text
                End With
             If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("DE" & CStr(iFila + 3) & ":DU" & CStr(iFila + 3))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
             Else
                ExcelCuadricularRango oExcel, oWorkSheet, iFila + 2, 109, iFila + 8, 125
             End If
                'Huella
            End If
            iFila = iFila + 2
            'Firmas
            iFila = iFila + 6
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula("Firma y Sello del Responsable de Farmacia y/o Laboratorio")
                Else
                    oWorkSheet.Cells(iFila, 3).Value = "Firma y Sello del Responsable de Farmacia y/o Laboratorio"
                End If
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(56, iFila - 2).setFormula("_____________________________________________")
            Else
                oWorkSheet.Cells(iFila, 3).Font.Bold = True
                oWorkSheet.Range(oWorkSheet.Cells(iFila, 3), oWorkSheet.Cells(iFila, 50)).Select
                With oExcel.Selection.Borders(xlEdgeTop)
                    .Weight = xlMedium
                    .ColorIndex = vbBlack
                End With
            End If
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(59, iFila - 1).setFormula("Firma del Afiliado o apoderado")
            Else
                oWorkSheet.Cells(iFila, 60).Value = "Firma del Afiliado o apoderado"
            End If
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 2).setFormula("_____________________________________________")
            Else
                oWorkSheet.Cells(iFila, 60).Font.Bold = True
                oWorkSheet.Range(oWorkSheet.Cells(iFila, 60), oWorkSheet.Cells(iFila, 90)).Select
            With oExcel.Selection.Borders(xlEdgeTop)
                .Weight = xlMedium
                .ColorIndex = vbBlack
            End With
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
                'Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                'If wxParametro338 = "S" Then
                   'Dim dummy()
                   'Document.Printer (dummy())
                'Else
                '   MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
                'End If
            Else
                If wxParametro338 = "S" Then
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
        Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select

End Sub

Sub ChequeaSiHaySaltoDePagina(ByRef lnFila As Long, oWorkSheet As Worksheet)
'    Dim lnDiv As Double, lnDiv1 As Double
'    lnDiv = lnFila / 55
'    lnDiv1 = Round(lnFila / 55, 0)
'    If lnDiv1 = lnDiv Then
       lnFila = lnFila + 1
       'oWorkSheet.Cells(lnFila, 93).Value = "F.Emisión: " & lcBuscaParametro.RetornaFechaHoraServidorSQL
       'oWorkSheet.Cells(lnFila + 1, 98).Value = "Cta: " & ml_idCuentaAtencion
       'lnFila = lnFila + 2
'    Else
'       lnFila = lnFila + 1
'    End If
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
        .Reconsideracion = IIf(Me.chkReconsideracion.Value = 1, "S", "N")
        .ReconsideracionCodigoDisa = Me.txtReconsideracion.Text
        .ReconsideracionLote = IIf(Me.chkReconsideracion.Value = 1, txtFua2.Text, "")
        .ReconsideracionNroFormato = IIf(Me.chkReconsideracion.Value = 1, Right("00000000" & txtFua3.Text, 8), "")
        .FuaComponente = CargaComponente(chkCsubsidiado.Value, chkCSemiS.Value)
        .Situacion = 2
        .AfiliacionDisa = txtNroAfiliacion1.Text
        .AfiliacionTipoFormato = txtNroAfiliacion2.Text
        .AfiliacionNroFormato = txtNroAfiliacion3.Text
        '.CodigoTipoFormato                                                      'no va en galenhos
        .OrigenAseguradoInstitucion = "0"
        '.OrigenAseguradoCodigo                                                  'no va en galenhos
        '.Edad                                                                   'no va en galenhos
        '.GrupoEtareo                                                            'no va en galenhos
        .GrupoEtareo = "0"
        .Genero = IIf(UCase(Left(txtSexo.Text, 1)) = "M", 1, 0)
        .FuaAtencion = CargaAtencion(chkAtencionAmbulatoria.Value, chkAtencionReferencia.Value, chkAtencionEmergencia.Value)
        .FuaCondicionMaterna = CargaCondicionMaterna(chkGestante.Value, chkPuerpera.Value)
        .FuaNrohistoria = Left(txtNhistoriaClinica.Text, 20)
        .FuaConceptoPr = IIf(Val(mo_cmbConceptoP.BoundText) = 0, 1, Val(mo_cmbConceptoP.BoundText))
        .FuaConceptoPrAutoriz = txtNautorizacion.Text
        .FuaConceptoPrMonto = Val(txtMonto.Text)
        .FuaAtencionFecha = Format(CDate(txtFantencion.Text), sighentidades.DevuelveFechaSoloFormato_DMY)
        .FuaAtencionHora = txtHatencion.Text
        '
        mo_ReglasSISgalenhos.Sis_m_ee_ssDevuelveCodigoSIS txtRDcodigo.Text, lcCodigoSis, lcEstablecimiento
        .FuaReferidoDestinoCodigoRENAES = lcCodigoSis
        .FuaReferidoDestinoNreferencia = txtRDnumero.Text
        .FuaCodigoPrestacion = ucSISfuaCodPrestacion1.CodigoPrestacion
        .FuaPersonalQatiende = CargaOrigenPersonal(chkPAestablecimiento.Value, chkPAaisped.Value)
        .FuaAtencionLugar = CargaLugarAtencion(chkIntramural.Value, chkExtramural.Value)
        .FuaDestino = IIf(Val(mo_cmbIdDestinoAtencion.BoundText) = 0, 1, Val(mo_cmbIdDestinoAtencion.BoundText)) 'Frank
        
        If .FuaDestino = 8 Then 'Coordinado Entre Rosa Celio y Esteban Juarez 28/05/2015 - correo
            .FuaHospitalizadoFingreso = .FuaAtencionFecha
        Else
            If sighentidades.EsFecha(txtHfingreso.Text, "DD/MM/AAAA") = True Then
                .FuaHospitalizadoFingreso = Format(CDate(txtHfingreso.Text), sighentidades.DevuelveFechaSoloFormato_DMY)
            End If
        End If
            
        If .FuaDestino = 8 Then 'Coordinado Entre Rosa Celio y Esteban Juarez 28/05/2015 - correo
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
'        .Fnacimiento = Format(md_FechaNacimiento, SIGHEntidades.DevuelveFechaSoloFormato_YMD_SIS)
        .Fnacimiento = Format(md_FechaNacimiento, sighentidades.DevuelveFechaSoloFormato_DMY) 'Frank
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
        '
        'mo_ReglasSISgalenhos.Sis_m_ee_ssDevuelveCodigoSIS txtROcodigo.Text, lcCodigoSis, lcEstablecimiento
'        .FuaReferidoOrigenCodigoRENAES = lcCodigoSis
        .FuaReferidoOrigenCodigoRENAES = txtROcodigo.Text        '
        .FuaReferidoOrigenNreferencia = txtRONumero.Text
        
        .FuaVersionFormato = wxParametro358
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

Function CargaOrigenPersonal(lnPAestablecimiento As Integer, lnPAaisped As Integer) As String
    If lnPAestablecimiento <> 0 Then
       CargaOrigenPersonal = 1
    Else
       CargaOrigenPersonal = 2
    End If
End Function

Function CargaLugarAtencion(lnIntramural As Integer, lnExtramural As Integer) As String
    If lnIntramural <> 0 Then
       CargaLugarAtencion = "1"
    Else
       CargaLugarAtencion = "2"
    End If
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
        Select Case mi_opcion
             Case sghAgregar
             Case sghModificar
             Case sghConsultar
                Me.btnAceptar.Enabled = False
             Case sghEliminar
        End Select
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
          '
'          If Val(wxParametro320) = sghFuaTipo.sghFuaTipoManual Then
'             txtFua3.Text = Trim(Str(Val(mo_ReglasSISgalenhos.sisFuaAtencionUltimoCorrelativo()) + 1))
'          End If
          
            If Val(wxParametro320) = sghFuaTipo.sghFuaTipoAutomatico Then
               txtFua3.Text = Right("00000000" & Trim(Str(Val(mo_ReglasSISgalenhos.sisFuaAtencionUltimoCorrelativo()) + 1)), 8)
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
        oConexion.Close
        Set oConexion = Nothing
        '
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
               chkReconsideracion.Value = IIf(.Reconsideracion = "S", 1, 0)
               txtReconsideracion.Text = .ReconsideracionCodigoDisa
               chkCsubsidiado.Value = IIf(.FuaComponente = "1", 1, 0)
               chkCSemiS.Value = IIf(.FuaComponente = 2, 1, 0)
              ' chkTAnuevo.Value
              ' chkTAAntiguoI.Value
              ' chkTAantiguoA.Value
               txtNroAfiliacion1.Text = .AfiliacionDisa
               txtNroAfiliacion2.Text = .AfiliacionTipoFormato
               txtNroAfiliacion3.Text = .AfiliacionNroFormato
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
               txtInstitucion.Text = .OrigenAseguradoInstitucion
               'txtCodSeguro.Text
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
               
               '
               txtRONumero.Text = .FuaReferidoOrigenNreferencia
               mo_cmbIdDestinoAtencion.BoundText = .FuaDestino
               ml_IdDestinoPaciente = .FuaDestino
               '
               lcCodigoRenaes = "": lcDescripcionRenaes = ""
               mo_ReglasSISgalenhos.Sis_m_ee_ssDevuelveCodigoRENAES .FuaReferidoDestinoCodigoRENAES, lcCodigoRenaes, lcDescripcionRenaes
               txtRDcodigo.Text = lcCodigoRenaes
               txtRD.Text = lcDescripcionRenaes
               '
               txtRDnumero.Text = .FuaReferidoDestinoNreferencia
               txtHfingreso.Text = DevuelveFechaSegunFormato_YMD_SIS(.FuaHospitalizadoFingreso)
               txtHfalta.Text = DevuelveFechaSegunFormato_YMD_SIS(.FuaHospitalizadoFalta)
               txtMedicoDni.Text = .FuaMedicoDNI
               txtMedico.Text = .FuaMedico
               If txtMedicoColegiatura.Text = "" Then
                  CargaDatosMedico oConexion, True
               End If
               'txtMedicoColegiatura.Text
               txtMedicoEspecialidad.Text = .FuaMedicoTipo
               txtObservaciones.Text = .FuaObservaciones
               txtFantencion.Text = DevuelveFechaSegunFormato_YMD_SIS(.FuaAtencionFecha)
               txtHatencion.Text = .FuaAtencionHora
               mo_cmbConceptoP.BoundText = .FuaConceptoPr
               ml_IdConceptoPrestacional = .FuaConceptoPr
               txtNautorizacion.Text = .FuaConceptoPrAutoriz
               txtMonto.Text = .FuaConceptoPrMonto
        End With
        ml_edad_En_Dias = sighentidades.EdadActualEnDias(CDate(txtFnacimiento.Text), CDate(txtFantencion.Text & " " & txtHatencion.Text))
        ml_edad_En_YYYYMMDD = sighentidades.EdadActualEnFormatoYYYYMMDD(CDate(txtFnacimiento.Text), CDate(txtFantencion.Text & " " & txtHatencion.Text))
        '
        CargaDatosDeDx oConexionExterna, False
        CargaConsumosEnServiciosIntermedios oConexionExterna, False
        CargaDatosDeTriajeVacunas oConexionExterna, False
        
        '
        If Val(oDoSisFuaAtencion.CabNroEnvioAlSIS) > 0 Then
           Me.btnAceptar.Enabled = False
           lcOpcion = lcOpcion & " (Ya fué enviado al SIS CENTRAL)"
        End If
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
                                        lcInsumo, lcMedicamento, ml_idUsuario, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                        "Fua: " & Trim(txtFua1.Text) & "-" & Trim(txtFua2.Text) & Trim(txtFua3.Text) & _
                                        " - Cta: " & Trim(Str(ml_IdCuentaAtencion)) & " - " & Trim(Me.txtPaciente.Text), wxParametro320, ml_idAtencion, lnNroFuaRepetido)
End Function


'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean
    ModificarDatos = mo_ReglasSISgalenhos.FuaModificar(oDoSisFuaAtencion, oRsVacunasSp, oRsPatologia, oRsFarmacia, oRsDx, _
                                          lcInsumo, lcMedicamento, ml_idUsuario, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                          "Fua: " & Trim(txtFua1.Text) & "-" & Trim(txtFua2.Text) & "-" & Trim(txtFua3.Text) & _
                                          "  Cta: " & Trim(Str(ml_IdCuentaAtencion)) & " - " & Trim(Me.txtPaciente.Text), ml_idAtencion)
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    EliminarDatos = mo_ReglasSISgalenhos.FuaEliminar(oDoSisFuaAtencion, ml_idUsuario, mo_lnIdTablaLISTBARITEMS, _
                                         mo_lcNombrePc, "Fua: " & Trim(txtFua1.Text) & "-" & Trim(txtFua2.Text) & "-" & Trim(txtFua3.Text) & _
                                         "  Cta: " & Trim(Str(ml_IdCuentaAtencion)) & " - " & Trim(Me.txtPaciente.Text), ml_idAtencion)
End Function


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
        Set oRsTmp1 = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(ml_IdCuentaAtencion, oConexion)
        If oRsTmp1.RecordCount > 0 Then
           lcElServicioUsaGalenHos = mo_ReglasArchivoClinico.ServicioUsaGalenHos(oRsTmp1.Fields!IdServicioIngreso)
           ml_IdTipoServicio = oRsTmp1.Fields!IdTipoServicio
           ml_Paciente = Trim(oRsTmp1.Fields!ApellidoPaterno) & " " & Trim(oRsTmp1.Fields!ApellidoMaterno) & " " & Trim(oRsTmp1.Fields!PrimerNombre) & IIf(IsNull(oRsTmp1.Fields!SegundoNombre), "", " " & oRsTmp1.Fields!SegundoNombre)
           md_FechaNacimiento = oRsTmp1.Fields!FechaNacimiento
           ml_Sexo = IIf(oRsTmp1.Fields!idTipoSexo = 2, lcFemenino, lcMasculino)
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
           ml_HoraAtencion = oRsTmp1.Fields!horaIngreso
           md_FechaAtencion = oRsTmp1.Fields!fechaIngreso
           'mgaray20140926
           If ml_IdTipoServicio = sghConsultaExterna Or ml_IdTipoServicio = 5 Or mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia Then
              ml_IdMedico = oRsTmp1.Fields!IdMedicoIngreso
              mo_Formulario.HabilitarDeshabilitar txtHfingreso, False
              mo_Formulario.HabilitarDeshabilitar txtHfalta, False
           Else
              ml_IdMedico = IIf(IsNull(oRsTmp1.Fields!IdMedicoEgreso), 0, oRsTmp1.Fields!IdMedicoEgreso)
              Me.txtHfingreso.Text = Format(oRsTmp1.Fields!fechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY)
              Me.txtHfalta.Text = IIf(IsNull(oRsTmp1.Fields!fechaEgreso), _
                                      sighentidades.FECHA_VACIA_DMY, _
                                      Format(oRsTmp1.Fields!fechaEgreso, sighentidades.DevuelveFechaSoloFormato_DMY)) 'Frank 2508
              mo_Formulario.HabilitarDeshabilitar Me.txtHfingreso, False
              mo_Formulario.HabilitarDeshabilitar Me.txtHfalta, False
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
           If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroCitaCE Or mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia Then
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
    If chkSPconsejeriaPPffSI.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "308"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPconsejeriaPPffNO.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "308"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If chkSPprofilaxisOsi.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "309"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPprofilaxisOno.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "309"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If chkSPvitaminaKsi.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "311"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPvitaminaKno.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "311"
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
    If chkSPsicoprofilaxisSI.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "302"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPsicoprofilaxisNO.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "302"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If chkSPadmOxitocinaSI.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "303"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPadmOxitocinaNO.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "303"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If chkSPlactanciaMsi.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "002"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPlactanciaMno.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "002"
        oRsVacunasSp.Fields!Valor = "0"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    End If
    If chkSPsuplNsi.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "310"
        oRsVacunasSp.Fields!Valor = "1"
        oRsVacunasSp.Fields!esCheck = True
        oRsVacunasSp.Update
    ElseIf chkSPsuplNno.Value <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "310"
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
    If Val(txtVacEpatB.Text) <> 0 Then
        oRsVacunasSp.AddNew
        oRsVacunasSp.Fields!intervencionP = "119"
        oRsVacunasSp.Fields!Valor = txtVacEpatB.Text
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
    wxParametro358 = lcBuscaParametro.SeleccionaFilaParametro(358)
End Sub

Function ReglasDeConsistenciasAntesDeGrabarFUA() As Boolean
    'mgaray20140926
    If mo_lnIdTablaLISTBARITEMS = sghRegistroCitaCE Or mo_lnIdTablaLISTBARITEMS = sghAdmisionEmergencia Then
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
                    lcDx56 = "B15/J00/A09/Z35/"
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
                          lcMensaje = lcMensaje & "Para la PRESTACION elegida, debe registrar al menos 1 MEDICAMENTO/INSUMO o un CPT (rc12)" & Chr(13)
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

Private Sub ucSISfuaCodPrestacion1_LostFocus()
     If mo_lnIdTablaLISTBARITEMS = sghFormatoFUA And txtPaciente.Text = "" Then
        MsgBox "Debe elegir al Paciente SIS, antes del CODIGO DE PRESTACION", vbInformation, Me.Caption
        ucSISfuaCodPrestacion1.CodigoPrestacion = ""
        Exit Sub
     End If
     ReglasDeConsistenciasDespuesDeElegirCodigoPrestacion ucSISfuaCodPrestacion1.CodigoPrestacion
     ReglasDeConsistenciasDespuesDeElegirCodigoPrestacion1 ucSISfuaCodPrestacion1.CodigoPrestacion
     PermitirManipularDatosSegunSexo
     '18/05/2016
     
     '
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
    mo_Formulario.HabilitarDeshabilitar fraConsejPPFF, False
    mo_Formulario.HabilitarDeshabilitar fraPsicoprofilaxis, False
    mo_Formulario.HabilitarDeshabilitar txtSPcred, False
    mo_Formulario.HabilitarDeshabilitar fraEEDP, False
    mo_Formulario.HabilitarDeshabilitar fraAdmVitaminaK, False
    mo_Formulario.HabilitarDeshabilitar fraProfilaxisO, False
    mo_Formulario.HabilitarDeshabilitar fraAdmOxitocina, False
    mo_Formulario.HabilitarDeshabilitar fraLactanciaM, False
    mo_Formulario.HabilitarDeshabilitar txtSPpuerperio, False
    mo_Formulario.HabilitarDeshabilitar fraAdmSuplNutr, False
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
    mo_Formulario.HabilitarDeshabilitar txtVacEpatB, False
    mo_Formulario.HabilitarDeshabilitar txtVacSpr, False
    mo_Formulario.HabilitarDeshabilitar txtVacDt, False
    mo_Formulario.HabilitarDeshabilitar txtVacHVB, False
    mo_Formulario.HabilitarDeshabilitar txtVacPentaval, False
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
             Case "308"
                  mo_Formulario.HabilitarDeshabilitar fraConsejPPFF, True
             Case "309"
                  mo_Formulario.HabilitarDeshabilitar fraProfilaxisO, True
             Case "311"
                  mo_Formulario.HabilitarDeshabilitar fraAdmVitaminaK, True
             Case "312"
                  mo_Formulario.HabilitarDeshabilitar fraEEDP, True
             Case "302"
                  mo_Formulario.HabilitarDeshabilitar fraPsicoprofilaxis, True
             Case "303"
                  mo_Formulario.HabilitarDeshabilitar fraAdmOxitocina, True
             Case "002"
                  mo_Formulario.HabilitarDeshabilitar fraLactanciaM, True
             Case "310"
                  mo_Formulario.HabilitarDeshabilitar fraAdmSuplNutr, True
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
             Case "119"
                  mo_Formulario.HabilitarDeshabilitar txtVacEpatB, True
             Case "125"
                  mo_Formulario.HabilitarDeshabilitar txtVacSpr, True
             Case "007"
                  mo_Formulario.HabilitarDeshabilitar txtVacDt, True
             Case "315"
                  mo_Formulario.HabilitarDeshabilitar txtVacHVB, True
             Case "124"
                  mo_Formulario.HabilitarDeshabilitar txtVacPentaval, True
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
          'Componente - rc1
          Select Case oRsTmp1.Fields!rc01_idComponente
          Case 1  'subsidiado
                chkCsubsidiado.Enabled = True
                chkCSemiS.Enabled = False: chkCSemiS.Value = ssCBUnchecked
          Case 2  'semiSubsidiado
                chkCsubsidiado.Enabled = False: chkCsubsidiado.Value = ssCBUnchecked
                chkCSemiS.Enabled = True
          Case 3  'ambos
                chkCsubsidiado.Enabled = True
                chkCSemiS.Enabled = True
          End Select
       End If
       
       oRsTmp1.Close
    End If
    '
    oConexionExterna.Close
    Set oConexionExterna = Nothing
    Set oRsTmp1 = Nothing
End Sub

Sub AsignaComponente(lcComponente As String)
    Select Case lcComponente
    Case "1"
         chkCsubsidiado.Value = ssCBChecked
    Case "2"
         chkCSemiS.Value = ssCBChecked
    Case "3"
    Case "4"
    Case "5"
    End Select
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
              Case "209"
                 txtSPpuerperio.Text = lcValor
              Case "307"
                 If Val(lcValor) = 1 Then
                    chkSPconsejeriaNsi.Value = 1
                 Else
                    chkSPconsejeriaNno.Value = 1
                 End If
              Case "308"
                 If Val(lcValor) = 1 Then
                    chkSPconsejeriaPPffSI.Value = 1
                 Else
                    chkSPconsejeriaPPffNO.Value = 1
                 End If
              Case "309"
                 If Val(lcValor) = 1 Then
                    chkSPprofilaxisOsi.Value = 1
                 Else
                    chkSPprofilaxisOno.Value = 1
                 End If
              Case "311"
                 If Val(lcValor) = 1 Then
                    chkSPvitaminaKsi.Value = 1
                 Else
                    chkSPvitaminaKno.Value = 1
                 End If
              Case "312"
                 If Val(lcValor) = 1 Then
                    chkSPeedpSI.Value = 1
                 Else
                    chkSPeedpNO.Value = 1
                 End If
              Case "302"
                 If Val(lcValor) = 1 Then
                    chkSPsicoprofilaxisSI.Value = 1
                 Else
                    chkSPsicoprofilaxisNO.Value = 1
                 End If
              Case "303"
                 If Val(lcValor) = 1 Then
                    chkSPadmOxitocinaSI.Value = 1
                 Else
                    chkSPadmOxitocinaNO.Value = 1
                 End If
              Case "002"
                 If Val(lcValor) = 1 Then
                    chkSPlactanciaMsi.Value = 1
                 Else
                    chkSPlactanciaMno.Value = 1
                 End If
              Case "310"
                 If Val(lcValor) = 1 Then
                    chkSPsuplNsi.Value = 1
                 Else
                    chkSPsuplNno.Value = 1
                 End If
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
              Case "119"
                 txtVacEpatB.Text = lcValor
              Case "125"
                 txtVacSpr.Text = lcValor
              Case "007"
                 txtVacDt.Text = lcValor
              Case "315"
                 txtVacHVB.Text = lcValor
              Case "124"
                 txtVacPentaval.Text = lcValor
              Case Else
                 CargaVacunaYsp = False
              End Select
End Function
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





