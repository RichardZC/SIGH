VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.UserControl ucBS_Seleccion 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13155
   ScaleHeight     =   9015
   ScaleWidth      =   13155
   Begin VB.Frame fraBoton 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   870
      Left            =   3000
      TabIndex        =   220
      Top             =   8040
      Width           =   7095
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ucBS_Seleccion.ctx":0000
         DownPicture     =   "ucBS_Seleccion.ctx":04C4
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
         Left            =   4140
         Picture         =   "ucBS_Seleccion.ctx":09B0
         Style           =   1  'Graphical
         TabIndex        =   221
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Imprime"
         Height          =   615
         Left            =   2910
         Picture         =   "ucBS_Seleccion.ctx":0E9C
         Style           =   1  'Graphical
         TabIndex        =   222
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ucBS_Seleccion.ctx":1375
         DownPicture     =   "ucBS_Seleccion.ctx":17D5
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
         Left            =   1560
         Picture         =   "ucBS_Seleccion.ctx":1C4A
         Style           =   1  'Graphical
         TabIndex        =   223
         Top             =   180
         Width           =   1365
      End
   End
   Begin TabDlg.SSTab BS000_04 
      Height          =   6735
      Left            =   60
      TabIndex        =   6
      Top             =   1320
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Personales"
      TabPicture(0)   =   "ucBS_Seleccion.ctx":20BF
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "BS000_06(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "BS000_05(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "BS000_05(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "BS000_07"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Protocolo de Selección"
      TabPicture(1)   =   "ucBS_Seleccion.ctx":20DB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "BS000_06(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "BS000_05(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Exámenes"
      TabPicture(2)   =   "ucBS_Seleccion.ctx":20F7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "BS000_06(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "BS000_05(7)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "BS000_05(8)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Calificación del Donante"
      TabPicture(3)   =   "ucBS_Seleccion.ctx":2113
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "BS000_06(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "BS000_05(9)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.Frame BS000_05 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Index           =   9
         Left            =   -70560
         TabIndex        =   210
         Top             =   2520
         Width           =   3615
         Begin Threed.SSOption optApto 
            Height          =   195
            Index           =   2
            Left            =   465
            TabIndex        =   213
            Top             =   840
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "3.- NO APTO PERMANENTEMENTE"
         End
         Begin Threed.SSOption optApto 
            Height          =   195
            Index           =   1
            Left            =   465
            TabIndex        =   212
            Top             =   480
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "2.- NO APTO TEMPORALMENTE"
         End
         Begin Threed.SSOption optApto 
            Height          =   195
            Index           =   0
            Left            =   465
            TabIndex        =   211
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "1.- APTO"
         End
      End
      Begin VB.Frame BS000_05 
         BackColor       =   &H00C0C0C0&
         Caption         =   "4.- Exámenes Complementarios"
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
         Height          =   2930
         Index           =   8
         Left            =   -72840
         TabIndex        =   174
         Top             =   3060
         Width           =   8535
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   29
            Left            =   2160
            MaxLength       =   35
            TabIndex        =   189
            Top             =   900
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   25
            Left            =   2160
            MaxLength       =   35
            TabIndex        =   188
            Top             =   240
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   27
            Left            =   2160
            MaxLength       =   35
            TabIndex        =   187
            Top             =   570
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   31
            Left            =   2160
            MaxLength       =   35
            TabIndex        =   186
            Top             =   1230
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   37
            Left            =   2160
            MaxLength       =   35
            TabIndex        =   185
            Top             =   2220
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   33
            Left            =   2160
            MaxLength       =   35
            TabIndex        =   184
            Top             =   1560
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   35
            Left            =   2160
            MaxLength       =   35
            TabIndex        =   183
            Top             =   1890
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   39
            Left            =   2160
            MaxLength       =   35
            TabIndex        =   182
            Top             =   2550
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   30
            Left            =   6360
            MaxLength       =   35
            TabIndex        =   181
            Top             =   900
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   26
            Left            =   6360
            MaxLength       =   35
            TabIndex        =   180
            Top             =   240
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   28
            Left            =   6360
            MaxLength       =   35
            TabIndex        =   179
            Top             =   570
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   32
            Left            =   6360
            MaxLength       =   35
            TabIndex        =   178
            Top             =   1230
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   38
            Left            =   6360
            MaxLength       =   35
            TabIndex        =   177
            Top             =   2220
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   34
            Left            =   6360
            MaxLength       =   35
            TabIndex        =   176
            Top             =   1560
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   36
            Left            =   6360
            MaxLength       =   35
            TabIndex        =   175
            Top             =   1890
            Width           =   1875
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Hematocrito"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   40
            Left            =   1245
            TabIndex        =   204
            Top             =   270
            Width           =   870
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "HbsAg"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   44
            Left            =   1650
            TabIndex        =   203
            Top             =   930
            Width           =   465
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Hb"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   42
            Left            =   1920
            TabIndex        =   202
            Top             =   600
            Width           =   195
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "VDRL/RPR"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   46
            Left            =   1380
            TabIndex        =   201
            Top             =   1260
            Width           =   735
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo Sanguíneo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   48
            Left            =   840
            TabIndex        =   200
            Top             =   1590
            Width           =   1230
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fenotipo RH"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   52
            Left            =   1185
            TabIndex        =   199
            Top             =   2250
            Width           =   885
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Factor RH"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   50
            Left            =   1350
            TabIndex        =   198
            Top             =   1920
            Width           =   720
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Anti HTLV"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   54
            Left            =   1380
            TabIndex        =   197
            Top             =   2580
            Width           =   690
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Anti Core VHB"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   41
            Left            =   5310
            TabIndex        =   196
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Anti VIH"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   45
            Left            =   5730
            TabIndex        =   195
            Top             =   930
            Width           =   585
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Anti Chagas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   43
            Left            =   5445
            TabIndex        =   194
            Top             =   600
            Width           =   870
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Anti VHC"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   47
            Left            =   5685
            TabIndex        =   193
            Top             =   1260
            Width           =   630
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Malaria"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   49
            Left            =   5760
            TabIndex        =   192
            Top             =   1590
            Width           =   510
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Variante Du"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   53
            Left            =   5430
            TabIndex        =   191
            Top             =   2250
            Width           =   840
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Bartonella"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   51
            Left            =   5550
            TabIndex        =   190
            Top             =   1920
            Width           =   720
         End
      End
      Begin VB.Frame BS000_05 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2.- Protocolo de Selección al Donante de Sangre"
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
         Height          =   6060
         Index           =   2
         Left            =   -74880
         TabIndex        =   78
         Top             =   480
         Width           =   12735
         Begin VB.Frame BS000_05 
            BackColor       =   &H00C0C0C0&
            Caption         =   "16.- ¿Ha tenido contacto sexual con algún grupo de riesgo?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1215
            Index           =   6
            Left            =   6960
            TabIndex        =   205
            Top             =   4080
            Width           =   4575
            Begin Threed.SSOption optNo 
               Height          =   195
               Index           =   15
               Left            =   3570
               TabIndex        =   219
               Top             =   600
               Visible         =   0   'False
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   344
               _Version        =   262144
               BackColor       =   12632256
               Enabled         =   0   'False
               Caption         =   "No"
            End
            Begin Threed.SSOption optSi 
               Height          =   195
               Index           =   15
               Left            =   3000
               TabIndex        =   218
               Top             =   600
               Visible         =   0   'False
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   344
               _Version        =   262144
               BackColor       =   12632256
               Enabled         =   0   'False
               Caption         =   "Si"
            End
            Begin VB.CheckBox chkCGR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Homosexual"
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
               Index           =   0
               Left            =   60
               TabIndex        =   209
               Top             =   240
               Width           =   1455
            End
            Begin VB.CheckBox chkCGR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Bisexual"
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
               Index           =   1
               Left            =   60
               TabIndex        =   208
               Top             =   480
               Width           =   1455
            End
            Begin VB.CheckBox chkCGR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Promiscuo"
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
               Index           =   2
               Left            =   60
               TabIndex        =   207
               Top             =   720
               Width           =   1455
            End
            Begin VB.CheckBox chkCGR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Prostituta"
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
               Index           =   3
               Left            =   60
               TabIndex        =   206
               Top             =   960
               Width           =   1455
            End
         End
         Begin VB.Frame BS000_05 
            BackColor       =   &H00C0C0C0&
            Caption         =   "12.- ¿Ha tenido o tiene alguna(s) de las siguientes enfermedades?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   2415
            Index           =   3
            Left            =   6960
            TabIndex        =   124
            Top             =   240
            Width           =   5655
            Begin Threed.SSOption optNo 
               Height          =   195
               Index           =   11
               Left            =   5010
               TabIndex        =   215
               Top             =   1800
               Visible         =   0   'False
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   344
               _Version        =   262144
               BackColor       =   12632256
               Enabled         =   0   'False
               Caption         =   "No"
            End
            Begin Threed.SSOption optSi 
               Height          =   195
               Index           =   11
               Left            =   4440
               TabIndex        =   214
               Top             =   1800
               Visible         =   0   'False
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   344
               _Version        =   262144
               BackColor       =   12632256
               Enabled         =   0   'False
               Caption         =   "Si"
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Glomerulonefritis"
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
               Index           =   23
               Left            =   3960
               TabIndex        =   148
               Top             =   1440
               Width           =   1575
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Osteomielitis (5a)"
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
               Index           =   22
               Left            =   3960
               TabIndex        =   147
               Top             =   1200
               Width           =   1575
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Mononucleosis"
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
               Index           =   21
               Left            =   3960
               TabIndex        =   146
               Top             =   960
               Width           =   1575
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Amebiasis (1a)"
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
               Index           =   20
               Left            =   3960
               TabIndex        =   145
               Top             =   720
               Width           =   1575
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Hipertiroidismo"
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
               Index           =   19
               Left            =   3960
               TabIndex        =   144
               Top             =   480
               Width           =   1575
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Dengue (1a)"
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
               Index           =   18
               Left            =   3960
               TabIndex        =   143
               Top             =   240
               Width           =   1575
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Transtornos de Coagulación"
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
               Index           =   17
               Left            =   1920
               TabIndex        =   142
               Top             =   2160
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Enfermedades Venéreas (3a)"
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
               Index           =   16
               Left            =   1920
               TabIndex        =   141
               Top             =   1920
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Fiebre Reumática (Rp)"
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
               Index           =   15
               Left            =   1920
               TabIndex        =   140
               Top             =   1680
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Asma"
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
               Index           =   14
               Left            =   1920
               TabIndex        =   139
               Top             =   1440
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Diabetes (Rp)"
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
               Index           =   13
               Left            =   1920
               TabIndex        =   138
               Top             =   1200
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Cáncer (Rp)"
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
               Index           =   12
               Left            =   1920
               TabIndex        =   137
               Top             =   960
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Hemorragias"
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
               Index           =   11
               Left            =   1920
               TabIndex        =   136
               Top             =   720
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Convulsiones (Rp)"
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
               Index           =   10
               Left            =   1920
               TabIndex        =   135
               Top             =   480
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Hipertensión Arterial"
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
               Index           =   9
               Left            =   1920
               TabIndex        =   134
               Top             =   240
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Cardiopatías (Rp)"
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
               Index           =   8
               Left            =   90
               TabIndex        =   133
               Top             =   2160
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Bartolenosis"
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
               Index           =   7
               Left            =   90
               TabIndex        =   132
               Top             =   1920
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Chagas (Rp)"
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
               Index           =   6
               Left            =   90
               TabIndex        =   131
               Top             =   1680
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Paludismo"
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
               Index           =   5
               Left            =   90
               TabIndex        =   130
               Top             =   1440
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Fiebre Amarilla (1a)"
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
               Index           =   4
               Left            =   90
               TabIndex        =   129
               Top             =   1200
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Fiebre Malta (3a)"
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
               Index           =   3
               Left            =   90
               TabIndex        =   128
               Top             =   960
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Fiebre Tifoidea (2a)"
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
               Index           =   2
               Left            =   90
               TabIndex        =   127
               Top             =   720
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Tuberculosis (5a)"
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
               Index           =   1
               Left            =   90
               TabIndex        =   126
               Top             =   480
               Width           =   2415
            End
            Begin VB.CheckBox chkEnf 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Hepatitis"
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
               Index           =   0
               Left            =   90
               TabIndex        =   125
               Top             =   240
               Width           =   2415
            End
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   13
            Left            =   6060
            TabIndex        =   169
            Top             =   3120
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   13
            Left            =   5490
            TabIndex        =   168
            Top             =   3120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin VB.Frame BS000_05 
            BackColor       =   &H00C0C0C0&
            Caption         =   "15.- ¿Pertenece a algún grupo de riesgo?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1215
            Index           =   4
            Left            =   6960
            TabIndex        =   149
            Top             =   2760
            Width           =   3255
            Begin Threed.SSOption optSi 
               Height          =   195
               Index           =   14
               Left            =   1800
               TabIndex        =   217
               Top             =   840
               Visible         =   0   'False
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   344
               _Version        =   262144
               BackColor       =   12632256
               Enabled         =   0   'False
               Caption         =   "Si"
            End
            Begin Threed.SSOption optNo 
               Height          =   195
               Index           =   14
               Left            =   2370
               TabIndex        =   216
               Top             =   840
               Visible         =   0   'False
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   344
               _Version        =   262144
               BackColor       =   12632256
               Enabled         =   0   'False
               Caption         =   "No"
            End
            Begin VB.CheckBox chkPGR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Prostituta"
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
               Index           =   3
               Left            =   60
               TabIndex        =   153
               Top             =   960
               Width           =   1455
            End
            Begin VB.CheckBox chkPGR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Promiscuo"
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
               Index           =   2
               Left            =   60
               TabIndex        =   152
               Top             =   720
               Width           =   1455
            End
            Begin VB.CheckBox chkPGR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Bisexual"
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
               Index           =   1
               Left            =   60
               TabIndex        =   151
               Top             =   480
               Width           =   1455
            End
            Begin VB.CheckBox chkPGR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Homosexual"
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
               Index           =   0
               Left            =   60
               TabIndex        =   150
               Top             =   240
               Width           =   1455
            End
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   0
            Left            =   5490
            TabIndex        =   108
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   0
            Left            =   6060
            TabIndex        =   107
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   1
            Left            =   5490
            TabIndex        =   106
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   1
            Left            =   6060
            TabIndex        =   105
            Top             =   480
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   2
            Left            =   5490
            TabIndex        =   104
            Top             =   720
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   2
            Left            =   6060
            TabIndex        =   103
            Top             =   720
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   3
            Left            =   5490
            TabIndex        =   102
            Top             =   960
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   3
            Left            =   6060
            TabIndex        =   101
            Top             =   960
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   4
            Left            =   5490
            TabIndex        =   100
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   4
            Left            =   6060
            TabIndex        =   99
            Top             =   1200
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   9
            Left            =   5490
            TabIndex        =   98
            Top             =   2400
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   9
            Left            =   6060
            TabIndex        =   97
            Top             =   2400
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   10
            Left            =   5490
            TabIndex        =   96
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   10
            Left            =   6060
            TabIndex        =   95
            Top             =   2640
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   7
            Left            =   5490
            TabIndex        =   94
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   7
            Left            =   6060
            TabIndex        =   93
            Top             =   1920
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   8
            Left            =   5490
            TabIndex        =   92
            Top             =   2160
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   8
            Left            =   6060
            TabIndex        =   91
            Top             =   2160
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   12
            Left            =   5490
            TabIndex        =   90
            Top             =   2880
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   12
            Left            =   6060
            TabIndex        =   89
            Top             =   2880
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   5
            Left            =   5490
            TabIndex        =   88
            Top             =   1440
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   5
            Left            =   6060
            TabIndex        =   87
            Top             =   1440
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   6
            Left            =   5490
            TabIndex        =   86
            Top             =   1680
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   6
            Left            =   6060
            TabIndex        =   85
            Top             =   1680
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   16
            Left            =   5490
            TabIndex        =   84
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   16
            Left            =   6060
            TabIndex        =   83
            Top             =   3360
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   17
            Left            =   5490
            TabIndex        =   82
            Top             =   3600
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   17
            Left            =   6060
            TabIndex        =   81
            Top             =   3600
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin Threed.SSOption optSi 
            Height          =   195
            Index           =   18
            Left            =   5490
            TabIndex        =   80
            Top             =   3840
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "Si"
         End
         Begin Threed.SSOption optNo 
            Height          =   195
            Index           =   18
            Left            =   6060
            TabIndex        =   79
            Top             =   3840
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   344
            _Version        =   262144
            BackColor       =   12632256
            Caption         =   "No"
         End
         Begin VB.Frame BS000_05 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Mujeres"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1935
            Index           =   5
            Left            =   240
            TabIndex        =   154
            Top             =   4080
            Visible         =   0   'False
            Width           =   6495
            Begin VB.TextBox BS000_01 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   17
               Left            =   3135
               TabIndex        =   173
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox BS000_01 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   16
               Left            =   3135
               TabIndex        =   172
               Top             =   150
               Width           =   1335
            End
            Begin Threed.SSOption optMens 
               Height          =   195
               Index           =   2
               Left            =   5520
               TabIndex        =   171
               Top             =   840
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   344
               _Version        =   262144
               BackColor       =   12632256
               Caption         =   "Escaso"
            End
            Begin VB.TextBox BS000_01 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   18
               Left            =   3135
               TabIndex        =   167
               Top             =   1335
               Width           =   975
            End
            Begin Threed.SSOption optNo 
               Height          =   195
               Index           =   20
               Left            =   3735
               TabIndex        =   160
               Top             =   1680
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   344
               _Version        =   262144
               BackColor       =   12632256
               Caption         =   "No"
            End
            Begin Threed.SSOption optSi 
               Height          =   195
               Index           =   20
               Left            =   3150
               TabIndex        =   159
               Top             =   1680
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   344
               _Version        =   262144
               BackColor       =   12632256
               Caption         =   "Si"
            End
            Begin Threed.SSOption optNo 
               Height          =   195
               Index           =   19
               Left            =   3735
               TabIndex        =   158
               Top             =   1080
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   344
               _Version        =   262144
               BackColor       =   12632256
               Caption         =   "No"
            End
            Begin Threed.SSOption optSi 
               Height          =   195
               Index           =   19
               Left            =   3135
               TabIndex        =   157
               Top             =   1080
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   344
               _Version        =   262144
               BackColor       =   12632256
               Caption         =   "Si"
            End
            Begin Threed.SSOption optMens 
               Height          =   195
               Index           =   1
               Left            =   4335
               TabIndex        =   156
               Top             =   840
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   344
               _Version        =   262144
               BackColor       =   12632256
               Caption         =   "Moderado"
            End
            Begin Threed.SSOption optMens 
               Height          =   195
               Index           =   0
               Left            =   3135
               TabIndex        =   155
               Top             =   840
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   344
               _Version        =   262144
               BackColor       =   12632256
               Caption         =   "Abundante"
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "25.- ¿Está dando de lactar?"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   29
               Left            =   1125
               TabIndex        =   166
               Top             =   1680
               Width           =   1980
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "22.- En su menstruación, el sangrado es "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   26
               Left            =   60
               TabIndex        =   165
               Top             =   840
               Width           =   2940
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "20.- ¿Cuando fué su última regla?"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   24
               Left            =   660
               TabIndex        =   164
               Top             =   180
               Width           =   2415
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "23.- ¿Está gestando?"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   27
               Left            =   1560
               TabIndex        =   163
               Top             =   1080
               Width           =   1530
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "21.- ¿Cuántos días menstrúa?"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   25
               Left            =   945
               TabIndex        =   162
               Top             =   510
               Width           =   2145
            End
            Begin VB.Label BS000_00 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "24.- Fecha del último parto"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   28
               Left            =   1155
               TabIndex        =   161
               Top             =   1365
               Width           =   1935
            End
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "19.- ¿Ha sido excluido como donante anteriormente?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   1650
            TabIndex        =   110
            Top             =   3840
            Width           =   3780
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "14.- ¿Ha tenido contacto directo con personas que tengan ictericia?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   555
            TabIndex        =   170
            Top             =   3120
            Width           =   4860
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "1.- ¿Ha donado sangre alguna vez?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   2835
            TabIndex        =   123
            Top             =   240
            Width           =   2550
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "3.- ¿Se puso nervioso cuando donó sangre?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   2235
            TabIndex        =   122
            Top             =   720
            Width           =   3150
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "2.- ¿Donó sangre los últimos 3 meses?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   2655
            TabIndex        =   121
            Top             =   480
            Width           =   2730
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "10.- ¿Ha sido tatuado?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   3765
            TabIndex        =   120
            Top             =   2400
            Width           =   1635
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "5.- ¿Ha recibido sangre, tranplante de órganos ó tejidos?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   1290
            TabIndex        =   119
            Top             =   1200
            Width           =   4110
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "4.- ¿Ha sido operado en los últimos 6 meses?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   2190
            TabIndex        =   118
            Top             =   960
            Width           =   3210
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "11.- ¿Ha sido sometido a punción de piel (aretes, acupunturas)?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   825
            TabIndex        =   117
            Top             =   2640
            Width           =   4590
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "9.- ¿Está tomando alguna medicina?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   2835
            TabIndex        =   116
            Top             =   2160
            Width           =   2580
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "8.- ¿Ha usado drogas ilegales?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   3225
            TabIndex        =   115
            Top             =   1920
            Width           =   2190
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "13.- ¿Ha tenido contacto directo con personas que tengan hepatitis?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   480
            TabIndex        =   114
            Top             =   2880
            Width           =   4935
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "7.- ¿Viajó fuera del país en los últimos años?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   2250
            TabIndex        =   113
            Top             =   1680
            Width           =   3165
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "6.- ¿Ha viajado a zona endémica de paludismo?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   2025
            TabIndex        =   112
            Top             =   1440
            Width           =   3390
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "17.- ¿Tuvo contacto sexual con más de una persona en los últimos 3 años?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   60
            TabIndex        =   111
            Top             =   3360
            Width           =   5370
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "18.- ¿Tiene SIDA o ha tenido alguna prueba positiva de SIDA?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   975
            TabIndex        =   109
            Top             =   3600
            Width           =   4455
         End
      End
      Begin VB.Frame BS000_05 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3.- Examen Clínico"
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
         Height          =   2535
         Index           =   7
         Left            =   -74280
         TabIndex        =   61
         Top             =   420
         Width           =   11295
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   23
            Left            =   2160
            MaxLength       =   35
            TabIndex        =   67
            Top             =   1560
            Width           =   8955
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   21
            Left            =   2160
            MaxLength       =   35
            TabIndex        =   66
            Top             =   900
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   19
            Left            =   2160
            MaxLength       =   35
            TabIndex        =   65
            Top             =   240
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   20
            Left            =   2160
            MaxLength       =   35
            TabIndex        =   64
            Top             =   570
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   555
            Index           =   24
            Left            =   2160
            MaxLength       =   35
            MultiLine       =   -1  'True
            TabIndex        =   63
            Top             =   1890
            Width           =   8955
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   22
            Left            =   2160
            MaxLength       =   35
            TabIndex        =   62
            Top             =   1230
            Width           =   1875
         End
         Begin VB.Label BS000_00 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Kg."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   31
            Left            =   4110
            TabIndex        =   77
            Top             =   270
            Width           =   240
         End
         Begin VB.Label BS000_00 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "mmHG"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   35
            Left            =   4110
            TabIndex        =   76
            Top             =   930
            Width           =   450
         End
         Begin VB.Label BS000_00 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "mt."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   33
            Left            =   4110
            TabIndex        =   75
            Top             =   600
            Width           =   240
         End
         Begin VB.Label BS000_00 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "pul/min"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   37
            Left            =   4110
            TabIndex        =   74
            Top             =   1260
            Width           =   510
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Estado de accesos venosos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   38
            Left            =   90
            TabIndex        =   73
            Top             =   1590
            Width           =   2040
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Peso"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   30
            Left            =   1770
            TabIndex        =   72
            Top             =   270
            Width           =   345
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Presión Arterial"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   34
            Left            =   1020
            TabIndex        =   71
            Top             =   930
            Width           =   1095
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Talla"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   1785
            TabIndex        =   70
            Top             =   600
            Width           =   330
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   39
            Left            =   1035
            TabIndex        =   69
            Top             =   1910
            Width           =   1065
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Pulso"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   36
            Left            =   1740
            TabIndex        =   68
            Top             =   1260
            Width           =   375
         End
      End
      Begin VB.CheckBox BS000_07 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Es para Donación por reposición"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   45
         Top             =   3660
         Width           =   2655
      End
      Begin VB.Frame BS000_05 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1. Datos del Donador"
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
         Height          =   2925
         Index           =   0
         Left            =   810
         TabIndex        =   29
         Top             =   540
         Width           =   11355
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   705
            Index           =   9
            Left            =   7350
            MaxLength       =   150
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   1865
            Width           =   3885
         End
         Begin VB.ComboBox BS000_03 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   5
            Left            =   7350
            TabIndex        =   19
            Top             =   1205
            Width           =   3870
         End
         Begin VB.ComboBox BS000_03 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   3
            Left            =   7350
            TabIndex        =   17
            Top             =   555
            Width           =   3915
         End
         Begin VB.ComboBox BS000_03 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   8
            Left            =   1410
            TabIndex        =   14
            Top             =   2195
            Width           =   4140
         End
         Begin VB.ComboBox BS000_03 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   7
            Left            =   1410
            TabIndex        =   12
            Top             =   1865
            Width           =   2265
         End
         Begin VB.ComboBox BS000_03 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   6
            Left            =   1410
            TabIndex        =   11
            Top             =   1535
            Width           =   1830
         End
         Begin VB.ComboBox BS000_03 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   1410
            TabIndex        =   10
            Top             =   1205
            Width           =   1875
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   7350
            MaxLength       =   10
            TabIndex        =   20
            Top             =   1535
            Width           =   1515
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   3
            Left            =   1410
            MaxLength       =   35
            TabIndex        =   8
            Top             =   555
            Width           =   4155
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   2
            Left            =   1410
            MaxLength       =   35
            TabIndex        =   7
            Top             =   225
            Width           =   4155
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   4
            Left            =   1410
            MaxLength       =   35
            TabIndex        =   9
            Top             =   885
            Width           =   4155
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   6
            Left            =   4080
            MaxLength       =   9
            TabIndex        =   13
            Top             =   1865
            Width           =   1485
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   5
            Left            =   7350
            MaxLength       =   35
            TabIndex        =   18
            Top             =   885
            Width           =   3915
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   8
            Left            =   1410
            MaxLength       =   35
            TabIndex        =   15
            Top             =   2525
            Width           =   4155
         End
         Begin MSMask.MaskEdBox BS000_02 
            Height          =   315
            Index           =   1
            Left            =   7350
            TabIndex        =   16
            Top             =   225
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   0
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
            PromptChar      =   " "
         End
         Begin VB.Label BS000_00 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nº"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   74
            Left            =   3855
            TabIndex        =   41
            Top             =   1920
            Width           =   195
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   64
            Left            =   6000
            TabIndex        =   44
            Top             =   1895
            Width           =   1320
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Ocupación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   56
            Left            =   6000
            TabIndex        =   43
            Top             =   585
            Width           =   1320
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar Procedencia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   60
            Left            =   6000
            TabIndex        =   42
            Top             =   1235
            Width           =   1320
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   63
            Left            =   555
            TabIndex        =   40
            Top             =   1895
            Width           =   810
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   59
            Left            =   1005
            TabIndex        =   39
            Top             =   1235
            Width           =   360
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   62
            Left            =   6000
            TabIndex        =   38
            Top             =   1565
            Width           =   1320
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar Nacimiento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   65
            Left            =   135
            TabIndex        =   37
            Top             =   2225
            Width           =   1230
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Estado Civil"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   61
            Left            =   540
            TabIndex        =   36
            Top             =   1565
            Width           =   825
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nacimiento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   6000
            TabIndex        =   35
            Top             =   255
            Width           =   1320
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido Materno"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   55
            Left            =   165
            TabIndex        =   34
            Top             =   585
            Width           =   1200
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombres"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   57
            Left            =   735
            TabIndex        =   33
            Top             =   915
            Width           =   630
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido Paterno"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   195
            TabIndex        =   32
            Top             =   255
            Width           =   1170
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Centro de Trabajo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   58
            Left            =   6000
            TabIndex        =   31
            Top             =   915
            Width           =   1320
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   66
            Left            =   720
            TabIndex        =   30
            Top             =   2555
            Width           =   645
         End
      End
      Begin VB.Frame BS000_05 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos Personales del Postulante"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1600
         Index           =   1
         Left            =   840
         TabIndex        =   46
         Top             =   3660
         Visible         =   0   'False
         Width           =   11295
         Begin VB.ComboBox BS000_03 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   9
            Left            =   8700
            TabIndex        =   60
            Top             =   915
            Width           =   2475
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   15
            Left            =   1680
            MaxLength       =   35
            TabIndex        =   58
            Top             =   1235
            Width           =   4155
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   13
            Left            =   8700
            MaxLength       =   35
            TabIndex        =   56
            Top             =   585
            Width           =   1515
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   12
            Left            =   1680
            MaxLength       =   35
            TabIndex        =   50
            Top             =   585
            Width           =   4155
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   10
            Left            =   1680
            MaxLength       =   35
            TabIndex        =   49
            Top             =   255
            Width           =   4155
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   14
            Left            =   1680
            MaxLength       =   35
            TabIndex        =   48
            Top             =   915
            Width           =   4155
         End
         Begin VB.TextBox BS000_01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   11
            Left            =   8700
            MaxLength       =   35
            TabIndex        =   47
            Top             =   255
            Width           =   1515
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Grado de Parentesco"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   73
            Left            =   120
            TabIndex        =   59
            Top             =   1265
            Width           =   1515
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Cama"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   70
            Left            =   8235
            TabIndex        =   57
            Top             =   615
            Width           =   405
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Atención"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   72
            Left            =   7350
            TabIndex        =   55
            Top             =   945
            Width           =   1320
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Sala Hospitalización"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   69
            Left            =   240
            TabIndex        =   54
            Top             =   615
            Width           =   1395
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Diagnóstico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   71
            Left            =   810
            TabIndex        =   53
            Top             =   945
            Width           =   825
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre de Receptor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   67
            Left            =   150
            TabIndex        =   52
            Top             =   285
            Width           =   1485
         End
         Begin VB.Label BS000_00 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Historia Clínica"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   68
            Left            =   7350
            TabIndex        =   51
            Top             =   285
            Width           =   1320
         End
      End
      Begin VB.Shape BS000_06 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   6300
         Index           =   1
         Left            =   -74925
         Top             =   360
         Width           =   12855
      End
      Begin VB.Shape BS000_06 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   6300
         Index           =   0
         Left            =   75
         Top             =   360
         Width           =   12855
      End
      Begin VB.Shape BS000_06 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   6300
         Index           =   2
         Left            =   -74925
         Top             =   360
         Width           =   12855
      End
      Begin VB.Shape BS000_06 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   6300
         Index           =   3
         Left            =   -74925
         Top             =   360
         Width           =   12855
      End
   End
   Begin VB.TextBox BS000_01 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   5580
      TabIndex        =   1
      Top             =   540
      Width           =   1575
   End
   Begin VB.TextBox BS000_01 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   540
      Width           =   1575
   End
   Begin VB.ComboBox BS000_03 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      ItemData        =   "ucBS_Seleccion.ctx":212F
      Left            =   9360
      List            =   "ucBS_Seleccion.ctx":2131
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   870
      Width           =   2055
   End
   Begin VB.ComboBox BS000_03 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      ItemData        =   "ucBS_Seleccion.ctx":2133
      Left            =   5580
      List            =   "ucBS_Seleccion.ctx":213D
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   870
      Width           =   1575
   End
   Begin VB.ComboBox BS000_03 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      ItemData        =   "ucBS_Seleccion.ctx":2155
      Left            =   1560
      List            =   "ucBS_Seleccion.ctx":2165
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   870
      Width           =   1575
   End
   Begin MSMask.MaskEdBox BS000_02 
      Height          =   315
      Index           =   0
      Left            =   9360
      TabIndex        =   2
      Top             =   540
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      HideSelection   =   0   'False
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
      PromptChar      =   " "
   End
   Begin VB.Label BS000_00 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Postulante"
      Height          =   195
      Index           =   1
      Left            =   3960
      TabIndex        =   28
      Top             =   570
      Width           =   1515
   End
   Begin VB.Label BS000_00 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Registro"
      Height          =   195
      Index           =   2
      Left            =   7920
      TabIndex        =   27
      Top             =   570
      Width           =   1305
   End
   Begin VB.Label BS000_00 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Donante"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   26
      Top             =   570
      Width           =   1380
   End
   Begin VB.Label BS000_00 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo postulante"
      Height          =   195
      Index           =   5
      Left            =   7920
      TabIndex        =   25
      Top             =   900
      Width           =   1305
   End
   Begin VB.Label BS000_00 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Factor RH"
      Height          =   195
      Index           =   4
      Left            =   3960
      TabIndex        =   24
      Top             =   900
      Width           =   1515
   End
   Begin VB.Label BS000_00 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo Sanguíneo"
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   23
      Top             =   900
      Width           =   1380
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "BANCO DE SANGRE - SELECCIÓN DEL POSTULANTES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   13365
   End
End
Attribute VB_Name = "ucBS_Seleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio

Dim ml_idUsuario As Long
Dim ml_idPrueba As Long
Dim ml_idOrden As Long
Dim ml_nombrePrueba As String
Dim ml_idAnalisis As Long
Dim ml_idPaciente As Long
Dim ml_resultado As String
Dim ml_observacion As String
Dim ml_IdMovimiento As Long
Dim ms_MensajeError As String

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property

Property Let idOrden(lValue As Long)
   ml_idOrden = lValue
End Property

Property Get idOrden() As Long
   idOrden = ml_idOrden
End Property

Property Let idPrueba(lValue As Long)
   ml_idPrueba = lValue
End Property

Property Get idPrueba() As Long
   idPrueba = ml_idPrueba
End Property

Property Let nombrePrueba(lValue As String)
   ml_nombrePrueba = lValue
End Property

Property Get nombrePrueba() As String
   nombrePrueba = ml_nombrePrueba
End Property

Property Let idAnalisis(lValue As Long)
   ml_idAnalisis = lValue
End Property

Property Get idAnalisis() As Long
   idAnalisis = ml_idAnalisis
End Property

Property Let idPaciente(lValue As Long)
   ml_idPaciente = lValue
End Property

Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property

Private Sub chkReposicion_Click()
  If chkReposicion.Value = True Then
    fraDatosPersonales.Visible = True
  Else
    fraDatosPersonales.Visible = False
  End If
End Sub

Private Sub cmdGrabar_Click()
  ml_resultado = BS000_01(2) & "/" & BS000_01(3) & "/" & BS000_01(4) & "/" & BS000_03(4) & "/" & BS000_03(6) & "/" & BS000_03(7) & "/" & BS000_01(6) & "/" & BS000_03(8) & "/" & BS000_01(8) & "/" & BS000_02(1) & "/" & BS000_03(3) & "/" & BS000_01(5) & "/" & BS000_03(5) & "/" & BS000_01(7) & "/" & BS000_01(9) & "/" & _
                 BS000_07 & "/" & BS000_01(10) & "/" & BS000_01(12) & "/" & BS000_01(14) & "/" & BS000_01(15) & "/" & BS000_01(11) & "/" & BS000_01(13) & "/" & BS000_03(9) & "/" & _
                 optSi(0) & "/" & optNo(0) & "/" & optSi(1) & "/" & optNo(1) & "/" & optSi(2) & "/" & optNo(2) & "/" & optSi(3) & "/" & optNo(3) & "/" & optSi(4) & "/" & optNo(4) & "/" & optSi(5) & "/" & optNo(5) & "/" & optSi(6) & "/" & optNo(6) & "/" & optSi(7) & "/" & optNo(7) & "/" & optSi(8) & "/" & optNo(8) & "/" & optSi(9) & "/" & optNo(9) & "/" & optSi(10) & "/" & optNo(10) & "/" & optSi(12) & "/" & optNo(12) & "/" & optSi(13) & "/" & optNo(13) & "/" & optSi(16) & "/" & optNo(16) & "/" & optSi(17) & "/" & optNo(17) & "/" & optSi(18) & "/" & optNo(18) & "/" & optSi(19) & "/" & optNo(19) & "/" & _
                 chkEnf(0) & "/" & chkEnf(1) & "/" & chkEnf(2) & "/" & chkEnf(3) & "/" & chkEnf(4) & "/" & chkEnf(5) & "/" & chkEnf(6) & "/" & chkEnf(7) & "/" & chkEnf(8) & "/" & chkEnf(9) & "/" & chkEnf(10) & "/" & chkEnf(11) & "/" & chkEnf(12) & "/" & chkEnf(13) & "/" & chkEnf(14) & "/" & chkEnf(15) & "/" & chkEnf(16) & "/" & chkEnf(17) & "/" & chkEnf(18) & "/" & chkEnf(19) & "/" & optSi(20) & "/" & optNo(21) & "/" & chkEnf(22) & "/" & chkEnf(23) & "/" & optSi(11) & "/" & optNo(11) & "/" & _
                 chkPGR(0) & "/" & chkPGR(1) & "/" & chkPGR(2) & "/" & chkPGR(3) & "/" & optSi(14) & "/" & optNo(14) & "/" & _
                 chkCGR(0) & "/" & chkCGR(1) & "/" & chkCGR(2) & "/" & chkCGR(3) & "/" & optSi(15) & "/" & optNo(15) & "/" & _
                 BS000_01(16) & "/" & BS000_01(17) & "/" & optMens(0) & "/" & optMens(1) & "/" & optMens(2) & "/" & optSi(19) & "/" & optNo(19) & "/" & BS000_01(18) & "/" & optSi(20) & "/" & optNo(20) & "/" & _
                 BS000_01(19) & "/" & BS000_01(20) & "/" & BS000_01(21) & "/" & BS000_01(22) & "/" & BS000_01(23) & "/" & BS000_01(24) & "/" & _
                 BS000_01(25) & "/" & BS000_01(27) & "/" & BS000_01(29) & "/" & BS000_01(31) & "/" & BS000_01(33) & "/" & BS000_01(35) & "/" & BS000_01(37) & "/" & BS000_01(39) & "/" & BS000_01(26) & "/" & BS000_01(28) & "/" & BS000_01(30) & "/" & BS000_01(32) & "/" & BS000_01(34) & "/" & BS000_01(36) & "/" & BS000_01(38) & "/" & _
                 optApto(0) & "/" & optApto(1) & "/" & optApto(2)
End Sub

Private Sub UserControl_Resize()
  On Error Resume Next
  
End Sub
