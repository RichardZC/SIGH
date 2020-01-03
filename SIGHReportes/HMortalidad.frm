VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form HMortalidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mortalidad Hospitalaria por Departamentos y/o Servicios"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   60
      TabIndex        =   2
      Top             =   6240
      Width           =   13305
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HMortalidad.frx":0000
         DownPicture     =   "HMortalidad.frx":04C4
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
         Left            =   6780
         Picture         =   "HMortalidad.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HMortalidad.frx":0E9C
         DownPicture     =   "HMortalidad.frx":12FC
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
         Left            =   5250
         Picture         =   "HMortalidad.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6180
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   13290
      Begin VB.CheckBox chkSoloUnDxPorPaciente 
         Caption         =   "Considerar solo un Diagnóstico por Paciente"
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
         TabIndex        =   91
         Top             =   5775
         Value           =   1  'Checked
         Width           =   6210
      End
      Begin VB.Frame Frame2 
         Height          =   525
         Left            =   6510
         TabIndex        =   87
         Top             =   5610
         Width           =   6585
         Begin Threed.SSOption optTodos 
            Height          =   255
            Left            =   90
            TabIndex        =   88
            Top             =   180
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Todos"
            Value           =   -1
         End
         Begin Threed.SSOption optSoloSIS 
            Height          =   255
            Left            =   3960
            TabIndex        =   89
            Top             =   180
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Sólo SIS"
         End
         Begin Threed.SSOption optSoloNOSIS 
            Height          =   255
            Left            =   8700
            TabIndex        =   90
            Top             =   165
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Sólo NO SIS"
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Ciclos de Vida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3405
         Left            =   120
         TabIndex        =   46
         Top             =   1845
         Width           =   4995
         Begin VB.TextBox txtCol31 
            Alignment       =   1  'Right Justify
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
            Left            =   1725
            TabIndex        =   70
            Text            =   "1"
            Top             =   1020
            Width           =   540
         End
         Begin VB.TextBox txtCol22 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   69
            Text            =   "11"
            Top             =   630
            Width           =   540
         End
         Begin VB.TextBox txtCol21 
            Alignment       =   1  'Right Justify
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
            Left            =   1725
            TabIndex        =   68
            Text            =   "1"
            Top             =   630
            Width           =   540
         End
         Begin VB.ComboBox cmbCol2 
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
            ItemData        =   "HMortalidad.frx":1BE6
            Left            =   3345
            List            =   "HMortalidad.frx":1BF3
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   630
            Width           =   1530
         End
         Begin VB.TextBox txtCol12 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   66
            Text            =   "29"
            Top             =   240
            Width           =   540
         End
         Begin VB.TextBox txtCol11 
            Alignment       =   1  'Right Justify
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
            Left            =   1725
            TabIndex        =   65
            Text            =   "0"
            Top             =   240
            Width           =   540
         End
         Begin VB.ComboBox cmbCol1 
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
            ItemData        =   "HMortalidad.frx":1C0A
            Left            =   3345
            List            =   "HMortalidad.frx":1C17
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   240
            Width           =   1530
         End
         Begin VB.TextBox txtCol42 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   63
            Text            =   "9"
            Top             =   1410
            Width           =   540
         End
         Begin VB.TextBox txtCol41 
            Alignment       =   1  'Right Justify
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
            Left            =   1725
            TabIndex        =   62
            Text            =   "5"
            Top             =   1410
            Width           =   540
         End
         Begin VB.ComboBox cmbCol4 
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
            ItemData        =   "HMortalidad.frx":1C2E
            Left            =   3345
            List            =   "HMortalidad.frx":1C3B
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   1410
            Width           =   1530
         End
         Begin VB.TextBox txtCol32 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   60
            Text            =   "4"
            Top             =   1020
            Width           =   540
         End
         Begin VB.ComboBox cmbCol3 
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
            ItemData        =   "HMortalidad.frx":1C52
            Left            =   3345
            List            =   "HMortalidad.frx":1C5F
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   1020
            Width           =   1530
         End
         Begin VB.TextBox txtCol62 
            Alignment       =   1  'Right Justify
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
            Left            =   2625
            TabIndex        =   58
            Text            =   "19"
            Top             =   2175
            Width           =   540
         End
         Begin VB.TextBox txtCol61 
            Alignment       =   1  'Right Justify
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
            Left            =   1710
            TabIndex        =   57
            Text            =   "15"
            Top             =   2175
            Width           =   540
         End
         Begin VB.ComboBox cmbCol6 
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
            ItemData        =   "HMortalidad.frx":1C76
            Left            =   3330
            List            =   "HMortalidad.frx":1C83
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   2175
            Width           =   1530
         End
         Begin VB.TextBox txtCol52 
            Alignment       =   1  'Right Justify
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
            Left            =   2625
            TabIndex        =   55
            Text            =   "14"
            Top             =   1785
            Width           =   540
         End
         Begin VB.TextBox txtCol51 
            Alignment       =   1  'Right Justify
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
            Left            =   1710
            TabIndex        =   54
            Text            =   "10"
            Top             =   1785
            Width           =   540
         End
         Begin VB.ComboBox cmbCol5 
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
            ItemData        =   "HMortalidad.frx":1C9A
            Left            =   3330
            List            =   "HMortalidad.frx":1CA7
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   1785
            Width           =   1530
         End
         Begin VB.TextBox txtCol82 
            Alignment       =   1  'Right Justify
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
            Left            =   2610
            TabIndex        =   52
            Text            =   "150"
            Top             =   2970
            Width           =   540
         End
         Begin VB.TextBox txtCol81 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   51
            Text            =   "65"
            Top             =   2955
            Width           =   540
         End
         Begin VB.ComboBox cmbCol8 
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
            ItemData        =   "HMortalidad.frx":1CBE
            Left            =   3315
            List            =   "HMortalidad.frx":1CCB
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   2955
            Width           =   1530
         End
         Begin VB.TextBox txtCol72 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   49
            Text            =   "64"
            Top             =   2565
            Width           =   540
         End
         Begin VB.TextBox txtCol71 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   48
            Text            =   "20"
            Top             =   2565
            Width           =   540
         End
         Begin VB.ComboBox cmbCol7 
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
            ItemData        =   "HMortalidad.frx":1CE2
            Left            =   3315
            List            =   "HMortalidad.frx":1CEF
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   2565
            Width           =   1530
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "a"
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
            Left            =   2400
            TabIndex        =   86
            Top             =   690
            Width           =   90
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Columna 2,  Desde:"
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
            TabIndex        =   85
            Top             =   690
            Width           =   1605
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "a"
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
            Left            =   2400
            TabIndex        =   84
            Top             =   300
            Width           =   90
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Columna 1,  Desde:"
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
            TabIndex        =   83
            Top             =   300
            Width           =   1605
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "a"
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
            Left            =   2400
            TabIndex        =   82
            Top             =   1470
            Width           =   90
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Columna 4,  Desde:"
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
            TabIndex        =   81
            Top             =   1470
            Width           =   1605
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "a"
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
            Left            =   2400
            TabIndex        =   80
            Top             =   1080
            Width           =   90
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Columna 3,  Desde:"
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
            TabIndex        =   79
            Top             =   1080
            Width           =   1605
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "a"
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
            Left            =   2385
            TabIndex        =   78
            Top             =   2235
            Width           =   90
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Columna 6,  Desde:"
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
            Left            =   75
            TabIndex        =   77
            Top             =   2235
            Width           =   1605
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "a"
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
            Left            =   2385
            TabIndex        =   76
            Top             =   1845
            Width           =   90
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Columna 5,  Desde:"
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
            Left            =   75
            TabIndex        =   75
            Top             =   1845
            Width           =   1605
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "a"
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
            Left            =   2370
            TabIndex        =   74
            Top             =   3015
            Width           =   90
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Columna 8,  Desde:"
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
            Left            =   60
            TabIndex        =   73
            Top             =   3015
            Width           =   1605
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "a"
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
            Left            =   2370
            TabIndex        =   72
            Top             =   2625
            Width           =   90
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Columna 7,  Desde:"
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
            Left            =   60
            TabIndex        =   71
            Top             =   2625
            Width           =   1605
         End
      End
      Begin VB.Frame frmDiagnosticos 
         Height          =   615
         Left            =   3390
         TabIndex        =   41
         Top             =   540
         Visible         =   0   'False
         Width           =   9780
         Begin VB.TextBox lblDescripcionDx 
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
            Left            =   2490
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   195
            Width           =   7155
         End
         Begin VB.TextBox txtIdDiagnostico 
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
            Left            =   1050
            TabIndex        =   43
            Top             =   195
            Width           =   1005
         End
         Begin VB.CommandButton btnBusquedaDiagnostico 
            Caption         =   "..."
            Height          =   315
            Left            =   2160
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   210
            Width           =   315
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Diagnóstico"
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
            Left            =   75
            TabIndex        =   45
            Top             =   240
            Width           =   930
         End
      End
      Begin VB.ComboBox cmbDiagnosticos 
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
         ItemData        =   "HMortalidad.frx":1D06
         Left            =   1425
         List            =   "HMortalidad.frx":1D10
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   720
         Width           =   1875
      End
      Begin VB.CheckBox chkDetallaHC 
         Caption         =   "Relación detalladas de Nº Historias Clínicas, debajo de cada Diagnóstico"
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
         Left            =   135
         TabIndex        =   38
         Top             =   5385
         Width           =   6210
      End
      Begin VB.Frame frmDistrito 
         Height          =   630
         Left            =   3375
         TabIndex        =   31
         Top             =   1185
         Visible         =   0   'False
         Width           =   9795
         Begin VB.ComboBox cmbIdDpto 
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
            Left            =   570
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   180
            Width           =   1605
         End
         Begin VB.ComboBox cmbIdProv 
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
            Left            =   3390
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   180
            Width           =   1950
         End
         Begin VB.ComboBox cmbIdDist 
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
            Left            =   6465
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   180
            Width           =   3240
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Dpto"
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
            Left            =   105
            TabIndex        =   37
            Top             =   225
            Width           =   405
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
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
            Height          =   210
            Left            =   2640
            TabIndex        =   36
            Top             =   225
            Width           =   705
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
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
            Height          =   210
            Left            =   5865
            TabIndex        =   35
            Top             =   225
            Width           =   570
         End
      End
      Begin VB.ComboBox cmbDistrito 
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
         ItemData        =   "HMortalidad.frx":1D25
         Left            =   1440
         List            =   "HMortalidad.frx":1D2F
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1365
         Width           =   1875
      End
      Begin VB.Frame frmTipoRep 
         Height          =   3270
         Left            =   6540
         TabIndex        =   8
         Top             =   2340
         Visible         =   0   'False
         Width           =   6555
         Begin VB.ComboBox cmbIdServicio2 
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
            Left            =   2205
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   2790
            Width           =   4200
         End
         Begin VB.ComboBox cmbIdEspecialidad2 
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
            Left            =   2205
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2400
            Width           =   4200
         End
         Begin VB.ComboBox cmbIdDepartamento2 
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
            Left            =   2205
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1995
            Width           =   4200
         End
         Begin VB.ComboBox cmbIdServicio1 
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
            Left            =   2205
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1305
            Width           =   4200
         End
         Begin VB.ComboBox cmbIdEspecialidad1 
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
            Left            =   2205
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   900
            Width           =   4200
         End
         Begin VB.ComboBox cmbIdDepartamento1 
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
            Left            =   2190
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   495
            Width           =   4200
         End
         Begin VB.Label lblTitulo2 
            Caption         =   "Servicio2"
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
            Left            =   105
            TabIndex        =   22
            Top             =   1695
            Width           =   5010
         End
         Begin VB.Label lblServicio2 
            Caption         =   "Servicio"
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
            Left            =   735
            TabIndex        =   21
            Top             =   2865
            Width           =   1275
         End
         Begin VB.Label lblEspecialidad2 
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
            Height          =   255
            Left            =   735
            TabIndex        =   20
            Top             =   2445
            Width           =   1395
         End
         Begin VB.Label lblDpto2 
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
            Height          =   285
            Left            =   735
            TabIndex        =   19
            Top             =   2055
            Width           =   1260
         End
         Begin VB.Label lblTitulo1 
            Caption         =   "Servicio1"
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
            Left            =   105
            TabIndex        =   15
            Top             =   195
            Width           =   5010
         End
         Begin VB.Label lblServicio1 
            Caption         =   "Servicio"
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
            Left            =   735
            TabIndex        =   14
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label lblEspecialidad1 
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
            Height          =   255
            Left            =   735
            TabIndex        =   13
            Top             =   945
            Width           =   1395
         End
         Begin VB.Label lblDpto1 
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
            Height          =   285
            Left            =   735
            TabIndex        =   12
            Top             =   555
            Width           =   1260
         End
      End
      Begin VB.ComboBox cmbTipoRep 
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
         ItemData        =   "HMortalidad.frx":1D44
         Left            =   6540
         List            =   "HMortalidad.frx":1D54
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1980
         Width           =   6525
      End
      Begin VB.ComboBox cmbTipoDx 
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
         ItemData        =   "HMortalidad.frx":1DBE
         Left            =   11160
         List            =   "HMortalidad.frx":1DCB
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   1995
      End
      Begin VB.ComboBox cmbSexo 
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
         ItemData        =   "HMortalidad.frx":1DFC
         Left            =   6495
         List            =   "HMortalidad.frx":1E09
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   2010
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1440
         TabIndex        =   25
         Top             =   180
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
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   3375
         TabIndex        =   26
         Top             =   180
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Diagnósticos"
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
         TabIndex        =   40
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Distrito Proced"
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
         Left            =   165
         TabIndex        =   30
         Top             =   1410
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
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
         Left            =   2880
         TabIndex        =   28
         Top             =   225
         Width           =   435
      End
      Begin VB.Label Label4 
         Caption         =   "F.Alta Médica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   165
         TabIndex        =   27
         Top             =   225
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Rep"
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
         TabIndex        =   7
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Diagnóstico"
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
         Left            =   9765
         TabIndex        =   5
         Top             =   225
         Width           =   1350
      End
      Begin VB.Label Departamento 
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
         Left            =   6015
         TabIndex        =   1
         Top             =   240
         Width           =   405
      End
   End
End
Attribute VB_Name = "HMortalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mortalidad
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbIdDepartamento1 As New sighentidades.ListaDespleglable
Dim mo_cmbIdServicio1 As New sighentidades.ListaDespleglable
Dim mo_cmbIdEspecialidad1 As New sighentidades.ListaDespleglable
Dim mo_cmbIdDepartamento2 As New sighentidades.ListaDespleglable
Dim mo_cmbIdServicio2 As New sighentidades.ListaDespleglable
Dim mo_cmbIdEspecialidad2 As New sighentidades.ListaDespleglable
Dim mo_cmbIdDpto As New sighentidades.ListaDespleglable
Dim mo_cmbIdProv As New sighentidades.ListaDespleglable
Dim mo_cmbIdDist As New sighentidades.ListaDespleglable
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_cmbTipoDx As New sighentidades.ListaDespleglable
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim idDiagnostico As Long
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_Titulo As String
Dim ml_TextoDelFiltro As String

Private Sub btnBusquedaDiagnostico_Click()
Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
Dim oDODiagnostico As DODiagnostico
    'mgaray20141023
    oBusqueda.MostrarSoloActivos = False
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            txtIdDiagnostico.Text = oDODiagnostico.CodigoCIE2004
            lblDescripcionDx.Text = oDODiagnostico.descripcion
            idDiagnostico = oDODiagnostico.idDiagnostico
        Else
            txtIdDiagnostico.Text = ""
            lblDescripcionDx.Text = ""
            idDiagnostico = 0
        End If
    End If
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub


Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRptMortalidad As New RptHMortalidad
        oRptMortalidad.FechaInicio = txtFechaInicio.Text
        oRptMortalidad.FechaFin = txtFechaFin.Text
        oRptMortalidad.idTipoSexo = cmbSexo.ListIndex
        oRptMortalidad.idTipoDiagnostico = Val(mo_cmbTipoDx.BoundText)
        oRptMortalidad.idDiagnostico = IIf(cmbDiagnosticos.ListIndex = 0, 0, idDiagnostico)
        oRptMortalidad.IdDistrito = IIf(cmbDistrito.ListIndex = 0, 0, mo_cmbIdDist.BoundText)
        oRptMortalidad.AnioCol11 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol11.Text), cmbCol1.ListIndex)
        oRptMortalidad.AnioCol12 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol12.Text), cmbCol1.ListIndex)
        oRptMortalidad.AnioCol21 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol21.Text), cmbCol2.ListIndex)
        oRptMortalidad.AnioCol22 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol22.Text), cmbCol2.ListIndex)
        oRptMortalidad.AnioCol31 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol31.Text), cmbCol3.ListIndex)
        oRptMortalidad.AnioCol32 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol32.Text), cmbCol3.ListIndex)
        oRptMortalidad.AnioCol41 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol41.Text), cmbCol4.ListIndex)
        oRptMortalidad.AnioCol42 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol42.Text), cmbCol4.ListIndex)
        oRptMortalidad.AnioCol51 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol51.Text), cmbCol5.ListIndex)
        oRptMortalidad.AnioCol52 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol52.Text), cmbCol5.ListIndex)
        oRptMortalidad.AnioCol61 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol61.Text), cmbCol6.ListIndex)
        oRptMortalidad.AnioCol62 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol62.Text), cmbCol6.ListIndex)
        oRptMortalidad.AnioCol71 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol71.Text), cmbCol7.ListIndex)
        oRptMortalidad.AnioCol72 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol72.Text), cmbCol7.ListIndex)
        oRptMortalidad.AnioCol81 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol81.Text), cmbCol8.ListIndex)
        oRptMortalidad.AnioCol82 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol82.Text), cmbCol8.ListIndex)
        oRptMortalidad.TipoReporte = cmbTipoRep.ListIndex
        oRptMortalidad.idDepartamento1 = IIf(mo_cmbIdDepartamento1.BoundText = "", 0, mo_cmbIdDepartamento1.BoundText)
        oRptMortalidad.idEspecialidad1 = IIf(mo_cmbIdEspecialidad1.BoundText = "", 0, mo_cmbIdEspecialidad1.BoundText)
        oRptMortalidad.idServicio1 = IIf(mo_cmbIdServicio1.BoundText = "", 0, mo_cmbIdServicio1.BoundText)
        oRptMortalidad.idDepartamento2 = IIf(mo_cmbIdDepartamento2.BoundText = "", 0, mo_cmbIdDepartamento2.BoundText)
        oRptMortalidad.idEspecialidad2 = IIf(mo_cmbIdEspecialidad2.BoundText = "", 0, mo_cmbIdEspecialidad2.BoundText)
        oRptMortalidad.idServicio2 = IIf(mo_cmbIdServicio2.BoundText = "", 0, mo_cmbIdServicio2.BoundText)
        oRptMortalidad.DetallaHC = IIf(chkDetallaHC.Value = 1, True, False)
        oRptMortalidad.Titulo = ml_Titulo
        oRptMortalidad.TextoDelFiltro = ml_TextoDelFiltro
        oRptMortalidad.TituloCol1 = txtCol11.Text & " - " & txtCol12.Text & " " & cmbCol1.Text
        oRptMortalidad.TituloCol2 = txtCol21.Text & " - " & txtCol22.Text & " " & cmbCol2.Text
        oRptMortalidad.TituloCol3 = txtCol31.Text & " - " & txtCol32.Text & " " & cmbCol3.Text
        oRptMortalidad.TituloCol4 = txtCol41.Text & " - " & txtCol42.Text & " " & cmbCol4.Text
        oRptMortalidad.TituloCol5 = txtCol51.Text & " - " & txtCol52.Text & " " & cmbCol5.Text
        oRptMortalidad.TituloCol6 = txtCol61.Text & " - " & txtCol62.Text & " " & cmbCol6.Text
        oRptMortalidad.TituloCol7 = txtCol71.Text & " - " & txtCol72.Text & " " & cmbCol7.Text
        oRptMortalidad.TituloCol8 = txtCol81.Text & " - " & txtCol82.Text & " " & cmbCol8.Text
        oRptMortalidad.lnSis = IIf(Me.optTodos.Value = True, 0, IIf(Me.optSoloSIS.Value = True, 1, 2))
        oRptMortalidad.CrearReporte Me.hwnd, IIf(chkSoloUnDxPorPaciente.Value = 1, True, False)
        Me.MousePointer = 1
    End If
End Sub


Function ValidaDatosObligatorios() As Boolean
    Dim sMensaje As String
    sMensaje = ""
    If Me.txtFechaInicio = sighentidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de alta médica inicial"
    Else
        If Not sighentidades.EsFecha(Me.txtFechaInicio, "DD/MM/AAAA") Then
            sMensaje = "La fecha de alta médica inicial, no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFechaFin = sighentidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de alta médica final"
    Else
        If Not sighentidades.EsFecha(Me.txtFechaFin, "DD/MM/AAAA") Then
            sMensaje = "La fecha de alta médica final, no tiene el formato correcto"
        End If
    End If
    If CDate(Me.txtFechaInicio.Text) > CDate(Me.txtFechaFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
       Exit Function
    End If
    ml_TextoDelFiltro = "FILTROS:   F.Alta Médica: (" & txtFechaInicio.Text & " hasta " & txtFechaFin.Text & "),   se consideró: " & cmbTipoDx.Text
    ml_TextoDelFiltro = ml_TextoDelFiltro & IIf(cmbSexo.ListIndex = 0, "", IIf(cmbSexo.ListIndex = 1, ",     (Sólo Masculinos)", ",     (Sólo Femeninos)"))
    ml_TextoDelFiltro = ml_TextoDelFiltro & IIf(Me.optTodos.Value = True, "", IIf(Me.optSoloSIS.Value = True, " (Solo SIS)", " (Sólo NO SIS)"))
    ml_TextoDelFiltro = ml_TextoDelFiltro & IIf(chkSoloUnDxPorPaciente.Value = 1, "  (Un Dx por Paciente)", "  (Varios Dx por Paciente)")
    If cmbTipoDx.Text = "" Then
       sMensaje = sMensaje + "Por favor elija el Tipo Diagnóstico" + Chr(13)
    End If
    
    Select Case cmbDiagnosticos.ListIndex
    Case 1
        If idDiagnostico = 0 Then
           sMensaje = sMensaje + "Por favor elija el Diagnóstico" + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Diagnóstico: " & txtIdDiagnostico.Text & " - " & lblDescripcionDx
    End Select
    Select Case cmbDistrito.ListIndex
    Case 1
        If cmbIdDist.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Distrito de Procedencia" + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Distrito de Procedencia: (" & cmbIdDpto.Text & ")/(" & cmbIdProv.Text & ")/(" & cmbIdDist.Text
    End Select
    Select Case cmbTipoRep.ListIndex
    Case 0
        ml_Titulo = "MORTALIDAD HOSPITALARIA"
    Case 1
        ml_Titulo = "MORTALIDAD HOSPITALARIA POR DEPARTAMENTO"
        If cmbIdDepartamento1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento" + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Departamento: " & cmbIdDepartamento1.Text
    Case 2    'por ESPECIALIDAD
        ml_Titulo = "MORTALIDAD HOSPITALARIA POR ESPECIALIDAD"
        If cmbIdDepartamento1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento" + Chr(13)
        End If
        If cmbIdEspecialidad1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija la Especialidad" + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Departamento: " & cmbIdDepartamento1.Text & "     Especialidad: " & cmbIdEspecialidad1.Text
    Case 3    'por 2 SERVICIOS
        ml_Titulo = "MORTALIDAD HOSPITALARIA CONSOLIDANDO DOS SERVICIOS"
        If cmbIdDepartamento1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento (para el primer Servicio)" + Chr(13)
        End If
        If cmbIdEspecialidad1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija la Especialidad (para  el primer Servicio)" + Chr(13)
        End If
        If cmbIdServicio1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Servicio (para  el primer Servicio)" + Chr(13)
        End If
        If cmbIdDepartamento2.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento (para el segundo Servicio)" + Chr(13)
        End If
        If cmbIdEspecialidad2.Text = "" Then
           sMensaje = sMensaje + "Por favor elija la Especialidad (para  el segundo Servicio)" + Chr(13)
        End If
        If cmbIdServicio2.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Servicio (para  el segundo Servicio)" + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Servicio1: (" & cmbIdDepartamento1.Text & ")/(" & cmbIdEspecialidad1.Text & ")/(" & cmbIdServicio1.Text & "),     Servicio2: (" & cmbIdDepartamento2.Text & ")/(" & cmbIdEspecialidad2.Text & ")/(" & cmbIdServicio2.Text & ")"
    End Select
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function



Private Sub cmbDiagnosticos_Change()
    Select Case cmbDiagnosticos.ListIndex
    Case 0
        frmDiagnosticos.Visible = False
    Case 1
        frmDiagnosticos.Visible = True
    End Select
End Sub

Private Sub cmbDiagnosticos_Click()
    cmbDiagnosticos_Change
End Sub

Private Sub cmbDistrito_Change()
    Select Case cmbDistrito.ListIndex
    Case 0
        frmDistrito.Visible = False
    Case 1
        frmDistrito.Visible = True
    End Select
 
End Sub

Private Sub cmbDistrito_Click()
   cmbDistrito_Change
End Sub

Private Sub cmbIdDepartamento1_Click()
       Dim sMensaje As String
       mo_cmbIdEspecialidad1.BoundColumn = "IdEspecialidad"
       mo_cmbIdEspecialidad1.ListField = "DescripcionLarga"
       Set mo_cmbIdEspecialidad1.RowSource = mo_AdminServiciosHosp.EspecialidadesSeleccionarporDepartamento(Val(mo_cmbIdDepartamento1.BoundText))
       mo_cmbIdEspecialidad1.BoundText = ""
       If mo_AdminServiciosHosp.MensajeError <> "" Then
          MsgBox mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
       End If
End Sub


Private Sub cmbIdDepartamento1_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento1
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdDepartamento2_Click()
       Dim sMensaje As String
       mo_cmbIdEspecialidad2.BoundColumn = "IdEspecialidad"
       mo_cmbIdEspecialidad2.ListField = "DescripcionLarga"
       Set mo_cmbIdEspecialidad2.RowSource = mo_AdminServiciosHosp.EspecialidadesSeleccionarporDepartamento(Val(mo_cmbIdDepartamento2.BoundText))
       mo_cmbIdEspecialidad2.BoundText = ""
       If mo_AdminServiciosHosp.MensajeError <> "" Then
          MsgBox mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
       End If
End Sub

Private Sub cmbIdDepartamento2_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento2
    AdministrarKeyPreview KeyCode

End Sub





Private Sub cmbIdDpto_Click()
       If cmbIdDpto.ListIndex = -1 Then Exit Sub
       mo_cmbIdProv.BoundColumn = "IdProvincia"
       mo_cmbIdProv.ListField = "Nombre"
       Set mo_cmbIdProv.RowSource = mo_AdminServiciosGeograficos.ProvinciasSeleccionarPorDepartamento(Val(cmbIdDpto.ItemData(cmbIdDpto.ListIndex)))
       mo_cmbIdProv.BoundText = ""
       mo_cmbIdDist.BoundText = ""
End Sub

Private Sub cmbIdEspecialidad1_Click()
    mo_cmbIdServicio1.BoundColumn = "IdServicio"
    mo_cmbIdServicio1.ListField = "DescripcionLarga"
    Set mo_cmbIdServicio1.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarPorTipoServicioDptoEspecialidad(3, Val(mo_cmbIdDepartamento1.BoundText), Val(mo_cmbIdEspecialidad1.BoundText))
End Sub



Private Sub cmbIdEspecialidad2_Click()
    mo_cmbIdServicio2.BoundColumn = "IdServicio"
    mo_cmbIdServicio2.ListField = "DescripcionLarga"
    Set mo_cmbIdServicio2.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarPorTipoServicioDptoEspecialidad(3, Val(mo_cmbIdDepartamento2.BoundText), Val(mo_cmbIdEspecialidad2.BoundText))
End Sub


Private Sub cmbIdProv_Click()
       If cmbIdProv.ListIndex = -1 Then Exit Sub
       mo_cmbIdDist.BoundColumn = "IdDistrito"
       mo_cmbIdDist.ListField = "Nombre"
       Set mo_cmbIdDist.RowSource = mo_AdminServiciosGeograficos.DistritoSeleccionarPorProvincia(Val(cmbIdProv.ItemData(cmbIdProv.ListIndex)))
       If mo_AdminServiciosGeograficos.MensajeError <> "" Then
            MsgBox mo_AdminServiciosGeograficos.MensajeError, vbInformation, "Datos de paciente"
       End If
       mo_cmbIdDist.BoundText = ""
End Sub

Private Sub cmbTipoRep_Change()
    Select Case cmbTipoRep.ListIndex
    Case 0
        frmTipoRep.Visible = False
        lblTitulo1.Visible = False
        lblDpto1.Visible = False
        lblServicio1.Visible = False
        lblEspecialidad1.Visible = False
        cmbIdDepartamento1.Visible = False
        cmbIdServicio1.Visible = False
        cmbIdEspecialidad1.Visible = False
        lblTitulo2.Visible = False
        lblDpto2.Visible = False
        lblServicio2.Visible = False
        lblEspecialidad2.Visible = False
        cmbIdDepartamento2.Visible = False
        cmbIdServicio2.Visible = False
        cmbIdEspecialidad2.Visible = False
    Case 1
        frmTipoRep.Visible = True
        lblTitulo1.Visible = True
        lblDpto1.Visible = True
        lblServicio1.Visible = False
        lblEspecialidad1.Visible = False
        cmbIdDepartamento1.Visible = True
        cmbIdServicio1.Visible = False
        cmbIdEspecialidad1.Visible = False
        lblTitulo2.Visible = False
        lblDpto2.Visible = False
        lblServicio2.Visible = False
        lblEspecialidad2.Visible = False
        cmbIdDepartamento2.Visible = False
        cmbIdServicio2.Visible = False
        cmbIdEspecialidad2.Visible = False
        lblTitulo1.Caption = "Elegir el Departamento:"
    Case 2
        frmTipoRep.Visible = True
        lblTitulo1.Visible = True
        lblDpto1.Visible = True
        lblServicio1.Visible = False
        lblEspecialidad1.Visible = True
        cmbIdDepartamento1.Visible = True
        cmbIdServicio1.Visible = False
        cmbIdEspecialidad1.Visible = True
        lblTitulo2.Visible = False
        lblDpto2.Visible = False
        lblServicio2.Visible = False
        lblEspecialidad2.Visible = False
        cmbIdDepartamento2.Visible = False
        cmbIdServicio2.Visible = False
        cmbIdEspecialidad2.Visible = False
        lblTitulo1.Caption = "Elegir el Servicio:"
    Case 3
        frmTipoRep.Visible = True
        lblTitulo1.Visible = True
        lblDpto1.Visible = True
        lblServicio1.Visible = True
        lblEspecialidad1.Visible = True
        cmbIdDepartamento1.Visible = True
        cmbIdServicio1.Visible = True
        cmbIdEspecialidad1.Visible = True
        lblTitulo2.Visible = True
        lblDpto2.Visible = True
        lblServicio2.Visible = True
        lblEspecialidad2.Visible = True
        cmbIdDepartamento2.Visible = True
        cmbIdServicio2.Visible = True
        cmbIdEspecialidad2.Visible = True
        lblTitulo1.Caption = "Elegir la primera Especialidad:"
        lblTitulo2.Caption = "Elegir la segunda Especialidad:"
    End Select
End Sub

Private Sub cmbTipoRep_Click()
    cmbTipoRep_Change
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdDepartamento1.MiComboBox = cmbIdDepartamento1
    Set mo_cmbIdEspecialidad1.MiComboBox = cmbIdEspecialidad1
    Set mo_cmbIdServicio1.MiComboBox = cmbIdServicio1
    Set mo_cmbIdDepartamento2.MiComboBox = cmbIdDepartamento2
    Set mo_cmbIdEspecialidad2.MiComboBox = cmbIdEspecialidad2
    Set mo_cmbIdServicio2.MiComboBox = cmbIdServicio2
    Set mo_cmbIdDpto.MiComboBox = cmbIdDpto
    Set mo_cmbIdProv.MiComboBox = cmbIdProv
    Set mo_cmbIdDist.MiComboBox = cmbIdDist
    Set mo_cmbTipoDx.MiComboBox = cmbTipoDx
End Sub

Private Sub Form_Load()
    Me.txtFechaInicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
    Me.txtFechaFin.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
    cmbSexo.ListIndex = 0
    cmbDiagnosticos.ListIndex = 0
    cmbDistrito.ListIndex = 0
    cmbTipoRep.ListIndex = 0
    cmbCol1.ListIndex = 0
    cmbCol2.ListIndex = 1
    cmbCol3.ListIndex = 2
    cmbCol4.ListIndex = 2
    cmbCol5.ListIndex = 2
    cmbCol6.ListIndex = 2
    cmbCol7.ListIndex = 2
    cmbCol8.ListIndex = 2
    mo_cmbTipoDx.BoundText = "305"
    CargaCombos
End Sub

Sub CargaCombos()
       mo_cmbIdDepartamento1.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento1.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento1.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()
       
       mo_cmbIdDepartamento2.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento2.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento2.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()
       
       mo_cmbIdDpto.BoundColumn = "IdDepartamento"
       mo_cmbIdDpto.ListField = "DescripcionLarga"
       Set mo_cmbIdDpto.RowSource = mo_AdminServiciosGeograficos.DepartamentosSeleccionarTodos()
       
       mo_cmbTipoDx.BoundColumn = "IdSubclasificacionDx"
       mo_cmbTipoDx.ListField = "DescripcionLarga"
       Set mo_cmbTipoDx.RowSource = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxHospMortalidad
       
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub



Private Sub txtFechaFin_LostFocus()
    If txtFechaFin <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFechaFin, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaFin = sighentidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFechaInicio_LostFocus()
    If txtFechaInicio <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFechaInicio, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaInicio = sighentidades.FECHA_VACIA_DMY
        End If
    End If
End Sub
