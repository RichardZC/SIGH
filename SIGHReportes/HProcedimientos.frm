VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form HProcedimientos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procedimientos Hospitalarios por Departamentos y/o Servicios"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   45
      TabIndex        =   2
      Top             =   4755
      Width           =   12510
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HProcedimientos.frx":0000
         DownPicture     =   "HProcedimientos.frx":04C4
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
         Left            =   6368
         Picture         =   "HProcedimientos.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HProcedimientos.frx":0E9C
         DownPicture     =   "HProcedimientos.frx":12FC
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
         Left            =   4838
         Picture         =   "HProcedimientos.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4710
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   12510
      Begin VB.ComboBox cmbTipoServicio 
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
         ItemData        =   "HProcedimientos.frx":1BE6
         Left            =   6480
         List            =   "HProcedimientos.frx":1BF3
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   4305
         Width           =   5880
      End
      Begin VB.CheckBox chkEnCantidadesCPT 
         Caption         =   "En cantidades de CPT"
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
         TabIndex        =   71
         Top             =   4350
         Width           =   2850
      End
      Begin VB.ComboBox cmbIdPtoCarga 
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
         ItemData        =   "HProcedimientos.frx":1C26
         Left            =   10260
         List            =   "HProcedimientos.frx":1C28
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   180
         Width           =   2085
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
         TabIndex        =   68
         Top             =   4035
         Width           =   6210
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
         Height          =   3420
         Left            =   135
         TabIndex        =   27
         Top             =   510
         Width           =   4995
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
            ItemData        =   "HProcedimientos.frx":1C2A
            Left            =   3315
            List            =   "HProcedimientos.frx":1C37
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   2565
            Width           =   1530
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
            TabIndex        =   62
            Text            =   "20"
            Top             =   2565
            Width           =   540
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
            TabIndex        =   61
            Text            =   "64"
            Top             =   2565
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
            ItemData        =   "HProcedimientos.frx":1C4E
            Left            =   3315
            List            =   "HProcedimientos.frx":1C5B
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   2955
            Width           =   1530
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
            TabIndex        =   59
            Text            =   "65"
            Top             =   2955
            Width           =   540
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
            TabIndex        =   58
            Text            =   "150"
            Top             =   2970
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
            ItemData        =   "HProcedimientos.frx":1C72
            Left            =   3330
            List            =   "HProcedimientos.frx":1C7F
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   1785
            Width           =   1530
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
            TabIndex        =   52
            Text            =   "10"
            Top             =   1785
            Width           =   540
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
            TabIndex        =   51
            Text            =   "14"
            Top             =   1785
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
            ItemData        =   "HProcedimientos.frx":1C96
            Left            =   3330
            List            =   "HProcedimientos.frx":1CA3
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   2175
            Width           =   1530
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
            TabIndex        =   49
            Text            =   "15"
            Top             =   2175
            Width           =   540
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
            TabIndex        =   48
            Text            =   "19"
            Top             =   2175
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
            ItemData        =   "HProcedimientos.frx":1CBA
            Left            =   3345
            List            =   "HProcedimientos.frx":1CC7
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1020
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
            TabIndex        =   42
            Text            =   "4"
            Top             =   1020
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
            ItemData        =   "HProcedimientos.frx":1CDE
            Left            =   3345
            List            =   "HProcedimientos.frx":1CEB
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1410
            Width           =   1530
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
            TabIndex        =   40
            Text            =   "5"
            Top             =   1410
            Width           =   540
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
            TabIndex        =   39
            Text            =   "9"
            Top             =   1410
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
            ItemData        =   "HProcedimientos.frx":1D02
            Left            =   3345
            List            =   "HProcedimientos.frx":1D0F
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   240
            Width           =   1530
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
            TabIndex        =   33
            Text            =   "0"
            Top             =   240
            Width           =   540
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
            TabIndex        =   32
            Text            =   "29"
            Top             =   240
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
            ItemData        =   "HProcedimientos.frx":1D26
            Left            =   3345
            List            =   "HProcedimientos.frx":1D33
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   630
            Width           =   1530
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
            TabIndex        =   30
            Text            =   "1"
            Top             =   630
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
            TabIndex        =   29
            Text            =   "11"
            Top             =   630
            Width           =   540
         End
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
            TabIndex        =   28
            Text            =   "1"
            Top             =   1020
            Width           =   540
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
            TabIndex        =   67
            Top             =   2625
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
            TabIndex        =   66
            Top             =   2625
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
            TabIndex        =   65
            Top             =   3015
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
            TabIndex        =   64
            Top             =   3015
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
            TabIndex        =   57
            Top             =   1845
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
            TabIndex        =   56
            Top             =   1845
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
            TabIndex        =   55
            Top             =   2235
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
            TabIndex        =   54
            Top             =   2235
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
            TabIndex        =   47
            Top             =   1080
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
            TabIndex        =   46
            Top             =   1080
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
            TabIndex        =   45
            Top             =   1470
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
            TabIndex        =   44
            Top             =   1470
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
            TabIndex        =   38
            Top             =   300
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
            TabIndex        =   37
            Top             =   300
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
            TabIndex        =   36
            Top             =   690
            Width           =   1605
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
            TabIndex        =   35
            Top             =   690
            Width           =   90
         End
      End
      Begin VB.Frame frmTipoRep 
         Height          =   3270
         Left            =   6495
         TabIndex        =   6
         Top             =   975
         Visible         =   0   'False
         Width           =   5865
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
            TabIndex        =   16
            Top             =   2790
            Width           =   3555
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
            TabIndex        =   15
            Top             =   2400
            Width           =   3555
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
            TabIndex        =   14
            Top             =   1995
            Width           =   3555
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
            TabIndex        =   9
            Top             =   1305
            Width           =   3555
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
            TabIndex        =   8
            Top             =   900
            Width           =   3555
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
            Left            =   2205
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   495
            Width           =   3555
         End
         Begin VB.Label lblTitulo2 
            Caption         =   "Especialidad2"
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
            Top             =   2055
            Width           =   1260
         End
         Begin VB.Label lblTitulo1 
            Caption         =   "Especialidad1"
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            TabIndex        =   10
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
         ItemData        =   "HProcedimientos.frx":1D4A
         Left            =   6495
         List            =   "HProcedimientos.frx":1D5D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   615
         Width           =   5880
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
         ItemData        =   "HProcedimientos.frx":1DE3
         Left            =   6495
         List            =   "HProcedimientos.frx":1DF0
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   2010
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1800
         TabIndex        =   23
         Top             =   210
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
         Left            =   3735
         TabIndex        =   24
         Top             =   210
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Servicio"
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
         Left            =   5430
         TabIndex        =   73
         Top             =   4365
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pto.Carga"
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
         Left            =   9390
         TabIndex        =   70
         Top             =   240
         Width           =   795
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
         Left            =   3240
         TabIndex        =   26
         Top             =   255
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "F.Procedimiento"
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
         TabIndex        =   25
         Top             =   225
         Width           =   1320
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
         Left            =   5445
         TabIndex        =   5
         Top             =   675
         Width           =   1005
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
Attribute VB_Name = "HProcedimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Procedimientos en Hospitalización
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
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_cmbIdPuntoCarga As New sighentidades.ListaDespleglable
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_Titulo As String
Dim ml_TextoDelFiltro As String

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub


Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRptProcedimientos As New RptHProcedimientos
        
        oRptProcedimientos.FechaInicio = Format(txtFechaInicio.Text & " 00:01", sighentidades.DevuelveFechaSoloFormato_DMY_HM)
        oRptProcedimientos.FechaFin = Format(txtFechaFin.Text & " 23:59", sighentidades.DevuelveFechaSoloFormato_DMY_HM)
        oRptProcedimientos.idTipoSexo = cmbSexo.ListIndex
        oRptProcedimientos.AnioCol11 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol11.Text), cmbCol1.ListIndex)
        oRptProcedimientos.AnioCol12 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol12.Text), cmbCol1.ListIndex)
        oRptProcedimientos.AnioCol21 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol21.Text), cmbCol2.ListIndex)
        oRptProcedimientos.AnioCol22 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol22.Text), cmbCol2.ListIndex)
        oRptProcedimientos.AnioCol31 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol31.Text), cmbCol3.ListIndex)
        oRptProcedimientos.AnioCol32 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol32.Text), cmbCol3.ListIndex)
        oRptProcedimientos.AnioCol41 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol41.Text), cmbCol4.ListIndex)
        oRptProcedimientos.AnioCol42 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol42.Text), cmbCol4.ListIndex)
        oRptProcedimientos.AnioCol51 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol51.Text), cmbCol5.ListIndex)
        oRptProcedimientos.AnioCol52 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol52.Text), cmbCol5.ListIndex)
        oRptProcedimientos.AnioCol61 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol61.Text), cmbCol6.ListIndex)
        oRptProcedimientos.AnioCol62 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol62.Text), cmbCol6.ListIndex)
        oRptProcedimientos.AnioCol71 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol71.Text), cmbCol7.ListIndex)
        oRptProcedimientos.AnioCol72 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol72.Text), cmbCol7.ListIndex)
        oRptProcedimientos.AnioCol81 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol81.Text), cmbCol8.ListIndex)
        oRptProcedimientos.AnioCol82 = sighentidades.ConvierteEnAnioUnMesOdia(Val(txtCol82.Text), cmbCol8.ListIndex)
        oRptProcedimientos.TipoReporte = cmbTipoRep.ListIndex
        oRptProcedimientos.idDepartamento1 = IIf(mo_cmbIdDepartamento1.BoundText = "", 0, mo_cmbIdDepartamento1.BoundText)
        oRptProcedimientos.idEspecialidad1 = IIf(mo_cmbIdEspecialidad1.BoundText = "", 0, mo_cmbIdEspecialidad1.BoundText)
        oRptProcedimientos.idServicio1 = IIf(mo_cmbIdServicio1.BoundText = "", 0, mo_cmbIdServicio1.BoundText)
        oRptProcedimientos.idDepartamento2 = IIf(mo_cmbIdDepartamento2.BoundText = "", 0, mo_cmbIdDepartamento2.BoundText)
        oRptProcedimientos.idEspecialidad2 = IIf(mo_cmbIdEspecialidad2.BoundText = "", 0, mo_cmbIdEspecialidad2.BoundText)
        oRptProcedimientos.idServicio2 = IIf(mo_cmbIdServicio2.BoundText = "", 0, mo_cmbIdServicio2.BoundText)
        oRptProcedimientos.DetallaHC = IIf(chkDetallaHC.Value = 1, True, False)
        oRptProcedimientos.Titulo = ml_Titulo
        oRptProcedimientos.TextoDelFiltro = ml_TextoDelFiltro
        oRptProcedimientos.TituloCol1 = txtCol11.Text & " - " & txtCol12.Text & " " & cmbCol1.Text
        oRptProcedimientos.TituloCol2 = txtCol21.Text & " - " & txtCol22.Text & " " & cmbCol2.Text
        oRptProcedimientos.TituloCol3 = txtCol31.Text & " - " & txtCol32.Text & " " & cmbCol3.Text
        oRptProcedimientos.TituloCol4 = txtCol41.Text & " - " & txtCol42.Text & " " & cmbCol4.Text
        oRptProcedimientos.TituloCol5 = txtCol51.Text & " - " & txtCol52.Text & " " & cmbCol5.Text
        oRptProcedimientos.TituloCol6 = txtCol61.Text & " - " & txtCol62.Text & " " & cmbCol6.Text
        oRptProcedimientos.TituloCol7 = txtCol71.Text & " - " & txtCol72.Text & " " & cmbCol7.Text
        oRptProcedimientos.TituloCol8 = txtCol81.Text & " - " & txtCol82.Text & " " & cmbCol8.Text
        oRptProcedimientos.PuntoCarga = mo_cmbIdPuntoCarga.BoundText
        oRptProcedimientos.idTipoServicio = cmbTipoServicio.ListIndex + 1
        oRptProcedimientos.CrearReporte Me.hwnd, IIf(chkEnCantidadesCPT.Value = 1, True, False)
        Me.MousePointer = 1
    End If
End Sub


Function ValidaDatosObligatorios() As Boolean
    Dim sMensaje As String
    sMensaje = ""
    If Me.txtFechaInicio = sighentidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de procedimiento inicial"
    Else
        If Not sighentidades.EsFecha(Me.txtFechaInicio, "DD/MM/AAAA") Then
            sMensaje = "La fecha de procedimiento inicial, no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFechaFin = sighentidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de procedimiento final"
    Else
        If Not sighentidades.EsFecha(Me.txtFechaFin, "DD/MM/AAAA") Then
            sMensaje = "La fecha de procedimiento final, no tiene el formato correcto"
        End If
    End If
    If CDate(Me.txtFechaInicio.Text) > CDate(Me.txtFechaFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
       Exit Function
    End If
    ml_TextoDelFiltro = "FILTROS:   F.Procedimiento: (" & txtFechaInicio.Text & " hasta " & txtFechaFin.Text & ")"
    ml_TextoDelFiltro = ml_TextoDelFiltro & IIf(cmbSexo.ListIndex = 0, "", IIf(cmbSexo.ListIndex = 1, ",     (Sólo Masculinos)", ",     (Sólo Femeninos)"))
    ml_TextoDelFiltro = ml_TextoDelFiltro & IIf(Me.chkEnCantidadesCPT.Value = 1, " (en CANTIDADES de CPT)", " (en CANTIDADES de Pacientes)")
    If cmbIdPtoCarga.Text = "" Then
        sMensaje = sMensaje + "Por favor elija el Punto de Carga" + Chr(13)
    Else
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Punto de Carga: " & Trim(cmbIdPtoCarga.Text) & " (solo " & IIf(Me.cmbTipoServicio.ListIndex = 0, "CE", IIf(Me.cmbTipoServicio.ListIndex = 2, "HOSP", "EMER")) & ")"
    End If
    Select Case cmbTipoRep.ListIndex
    Case 0
        ml_Titulo = "PROCEDIMIENTOS"
    Case 1
        ml_Titulo = "PROCEDIMIENTOS POR DEPARTAMENTO"
        If cmbIdDepartamento1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento" + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Departamento: " & cmbIdDepartamento1.Text
    Case 2    'por un Servicio
        ml_Titulo = "PROCEDIMIENTOS POR ESPECIALIDAD"
        If cmbIdDepartamento1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento" + Chr(13)
        End If
        If cmbIdEspecialidad1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Especialidad" + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Departamento: " & cmbIdDepartamento1.Text & "     Especialidad: " & cmbIdEspecialidad1.Text
    Case 3    'por 2 Especialidades
        ml_Titulo = "PROCEDIMIENTOS CONSOLIDANDO DOS SERVICIOS"
        If cmbIdDepartamento1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento (para el primer Servicio)" + Chr(13)
        End If
        If cmbIdEspecialidad1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija la Especialidad (para el primer Servicio)" + Chr(13)
        End If
        If cmbIdServicio1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Servicio (para el primer Servicio)" + Chr(13)
        End If
        If cmbIdDepartamento2.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento (para el segundo Servicio)" + Chr(13)
        End If
        If cmbIdEspecialidad2.Text = "" Then
           sMensaje = sMensaje + "Por favor elija la Especialidad (para el segundo Servicio)" + Chr(13)
        End If
        If cmbIdServicio2.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Servicio (para el segundo Servicio)" + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Servicio1: (" & cmbIdDepartamento1.Text & ")/(" & cmbIdEspecialidad1.Text & ")/(" & cmbIdServicio1.Text & "),     Servicio2: (" & cmbIdDepartamento2.Text & ")/(" & cmbIdEspecialidad2.Text & ")/(" & cmbIdServicio2.Text & ")"
    Case 4
        ml_Titulo = "PROCEDIMIENTOS CONSOLIDANDO UN SERVICIO"
        If cmbIdServicio1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Servicio " + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Servicio: (" & cmbIdDepartamento1.Text & ")/(" & cmbIdEspecialidad1.Text & ")/(" & cmbIdServicio1.Text & ")"
    End Select
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function




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






Private Sub cmbIdEspecialidad1_Click()
    mo_cmbIdServicio1.BoundColumn = "IdServicio"
    mo_cmbIdServicio1.ListField = "DescripcionLarga"
    Set mo_cmbIdServicio1.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarPorTipoServicioDptoEspecialidad(cmbTipoServicio.ListIndex + 1, Val(mo_cmbIdDepartamento1.BoundText), Val(mo_cmbIdEspecialidad1.BoundText))
End Sub



Private Sub cmbIdEspecialidad2_Click()
    mo_cmbIdServicio2.BoundColumn = "IdServicio"
    mo_cmbIdServicio2.ListField = "DescripcionLarga"
    Set mo_cmbIdServicio2.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarPorTipoServicioDptoEspecialidad(3, Val(mo_cmbIdDepartamento2.BoundText), Val(mo_cmbIdEspecialidad2.BoundText))
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
        lblTitulo1.Caption = "Elegir el primer Servicio:"
        lblTitulo2.Caption = "Elegir el segundo Servicio:"
    Case 4
        frmTipoRep.Visible = True
        lblTitulo1.Visible = True
        lblDpto1.Visible = True
        lblServicio1.Visible = True
        lblEspecialidad1.Visible = True
        cmbIdDepartamento1.Visible = True
        cmbIdServicio1.Visible = True
        cmbIdEspecialidad1.Visible = True
        lblTitulo1.Caption = "Elegir el Servicio:"
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
End Sub

Private Sub Form_Load()
    Me.txtFechaInicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
    Me.txtFechaFin.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
    cmbSexo.ListIndex = 0
    cmbTipoRep.ListIndex = 0
    cmbCol1.ListIndex = 0
    cmbCol2.ListIndex = 1
    cmbCol3.ListIndex = 2
    cmbCol4.ListIndex = 2
    cmbCol5.ListIndex = 2
    cmbCol6.ListIndex = 2
    cmbCol7.ListIndex = 2
    cmbCol8.ListIndex = 2
    CargaCombos
    cmbTipoServicio.ListIndex = 2
End Sub

Sub CargaCombos()
       mo_cmbIdDepartamento1.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento1.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento1.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()
       
       mo_cmbIdDepartamento2.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento2.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento2.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()
       
        Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPtoCarga
        mo_cmbIdPuntoCarga.ListField = "Descripcion"
        mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
        Set mo_cmbIdPuntoCarga.RowSource = mo_reglasComunes.SeleccionarPuntosDeCarga()
      
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
