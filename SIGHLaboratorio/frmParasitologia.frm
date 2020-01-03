VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmParasitologia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PARASITOLOGÍA"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   ForeColor       =   &H00000000&
   Icon            =   "frmParasitologia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   74
      Top             =   1710
      Width           =   7155
      Begin VB.ComboBox cmbResponsable 
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
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   180
         Width           =   3090
      End
      Begin MSMask.MaskEdBox txtFresultado 
         Height          =   315
         Left            =   5610
         TabIndex        =   76
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Realiza Prueba"
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
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "F.Resultado"
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
         Left            =   4590
         TabIndex        =   77
         Top             =   240
         Width           =   945
      End
   End
   Begin SIGHLaboratorio.UcPacienteDatos1 UcPacienteDatos1 
      Height          =   1815
      Left            =   60
      TabIndex        =   53
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3201
   End
   Begin VB.Frame fraBoton 
      ForeColor       =   &H00000000&
      Height          =   870
      Left            =   60
      TabIndex        =   68
      Top             =   6990
      Width           =   7200
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmParasitologia.frx":0CCA
         DownPicture     =   "frmParasitologia.frx":118E
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
         Left            =   3713
         Picture         =   "frmParasitologia.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprime  (F3)"
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
         Left            =   90
         Picture         =   "frmParasitologia.frx":1B66
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmParasitologia.frx":203F
         DownPicture     =   "frmParasitologia.frx":249F
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
         Left            =   2243
         Picture         =   "frmParasitologia.frx":2914
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   180
         Width           =   1365
      End
   End
   Begin VB.Frame COA005 
      Caption         =   "Examen Coprofuncional de heces (incluye reaccion inflamatoria)"
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
      Height          =   4635
      Left            =   60
      TabIndex        =   32
      Top             =   2340
      Visible         =   0   'False
      Width           =   7200
      Begin VB.ComboBox COA005_26 
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
         ItemData        =   "frmParasitologia.frx":2D89
         Left            =   2080
         List            =   "frmParasitologia.frx":2D93
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox COA005_28 
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
         Left            =   3465
         TabIndex        =   19
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox COA005_30 
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
         Left            =   6000
         TabIndex        =   20
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox COA005_27 
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
         Left            =   1305
         TabIndex        =   18
         Top             =   3720
         Width           =   855
      End
      Begin VB.ComboBox COA005_23 
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
         ItemData        =   "frmParasitologia.frx":2DAB
         Left            =   5640
         List            =   "frmParasitologia.frx":2DB5
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2760
         Width           =   1335
      End
      Begin VB.ComboBox COA005_16 
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
         ItemData        =   "frmParasitologia.frx":2DCD
         Left            =   2805
         List            =   "frmParasitologia.frx":2DDA
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2145
         Width           =   1575
      End
      Begin VB.ComboBox COA005_14 
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
         ItemData        =   "frmParasitologia.frx":2DF7
         Left            =   2805
         List            =   "frmParasitologia.frx":2E01
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1815
         Width           =   1575
      End
      Begin VB.ComboBox COA005_11 
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
         ItemData        =   "frmParasitologia.frx":2E19
         Left            =   2805
         List            =   "frmParasitologia.frx":2E23
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1485
         Width           =   1575
      End
      Begin VB.ComboBox COA005_09 
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
         ItemData        =   "frmParasitologia.frx":2E3B
         Left            =   2805
         List            =   "frmParasitologia.frx":2E45
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1155
         Width           =   1575
      End
      Begin VB.TextBox COA005_29 
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
         Left            =   1420
         TabIndex        =   21
         Top             =   4215
         Width           =   5670
      End
      Begin VB.TextBox COA005_02 
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
         Left            =   890
         TabIndex        =   6
         Text            =   "Diarreico"
         Top             =   460
         Width           =   1455
      End
      Begin VB.TextBox COA005_06 
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
         Left            =   5745
         TabIndex        =   8
         Text            =   "Fecal"
         Top             =   460
         Width           =   1335
      End
      Begin VB.TextBox COA005_04 
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
         Left            =   3360
         TabIndex        =   7
         Text            =   "Marrón"
         Top             =   460
         Width           =   1335
      End
      Begin VB.TextBox COA005_12 
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
         Left            =   4440
         TabIndex        =   11
         Top             =   1485
         Width           =   615
      End
      Begin VB.TextBox COA005_25 
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
         Left            =   5640
         TabIndex        =   16
         Top             =   3105
         Width           =   1335
      End
      Begin VB.TextBox COA005_19 
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
         Left            =   1320
         TabIndex        =   14
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label COA005_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "x c"
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
         Index           =   21
         Left            =   2280
         TabIndex        =   67
         Top             =   2790
         Width           =   240
      End
      Begin VB.Label COA005_00 
         AutoSize        =   -1  'True
         Caption         =   "Reacción Inflamatoria"
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
         Height          =   195
         Index           =   20
         Left            =   60
         TabIndex        =   66
         Top             =   3405
         Width           =   1875
      End
      Begin VB.Label COA005_00 
         AutoSize        =   -1  'True
         Caption         =   "MN"
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
         Index           =   17
         Left            =   5685
         TabIndex        =   65
         Top             =   3750
         Width           =   255
      End
      Begin VB.Label COA005_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Leucocitos"
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
         Index           =   13
         Left            =   360
         TabIndex        =   64
         Top             =   3750
         Width           =   870
      End
      Begin VB.Label COA005_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PMN"
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
         Index           =   15
         Left            =   3030
         TabIndex        =   63
         Top             =   3750
         Width           =   360
      End
      Begin VB.Label COA005_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "x c"
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
         Index           =   14
         Left            =   2310
         TabIndex        =   62
         Top             =   3750
         Width           =   240
      End
      Begin VB.Label COA005_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Index           =   16
         Left            =   4440
         TabIndex        =   61
         Top             =   3750
         Width           =   180
      End
      Begin VB.Label COA005_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Index           =   18
         Left            =   6900
         TabIndex        =   60
         Top             =   3750
         Width           =   180
      End
      Begin VB.Label COA005_00 
         AutoSize        =   -1  'True
         Caption         =   "Otros"
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
         Index           =   12
         Left            =   4860
         TabIndex        =   46
         Top             =   3135
         Width           =   705
      End
      Begin VB.Label COA005_00 
         AutoSize        =   -1  'True
         Caption         =   "Examen Microscópico"
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
         Height          =   195
         Index           =   9
         Left            =   60
         TabIndex        =   45
         Top             =   2520
         Width           =   1860
      End
      Begin VB.Label COA005_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Leucocitos"
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
         Index           =   10
         Left            =   375
         TabIndex        =   44
         Top             =   2790
         Width           =   870
      End
      Begin VB.Label COA005_00 
         AutoSize        =   -1  'True
         Caption         =   "Parásitos"
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
         Index           =   11
         Left            =   4860
         TabIndex        =   43
         Top             =   2790
         Width           =   705
      End
      Begin VB.Label COA005_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Reacción (pH)"
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
         Index           =   8
         Left            =   1500
         TabIndex        =   42
         Top             =   2175
         Width           =   1155
      End
      Begin VB.Label COA005_00 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
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
         Height          =   195
         Index           =   19
         Left            =   60
         TabIndex        =   41
         Top             =   4245
         Width           =   1275
      End
      Begin VB.Label COA005_00 
         AutoSize        =   -1  'True
         Caption         =   "Examen Químico"
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
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   40
         Top             =   960
         Width           =   1440
      End
      Begin VB.Label COA005_00 
         AutoSize        =   -1  'True
         Caption         =   "Examen Físico"
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
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   39
         Top             =   225
         Width           =   1260
      End
      Begin VB.Label COA005_00 
         AutoSize        =   -1  'True
         Caption         =   "Aspecto"
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
         Index           =   1
         Left            =   135
         TabIndex        =   38
         Top             =   495
         Width           =   675
      End
      Begin VB.Label COA005_00 
         AutoSize        =   -1  'True
         Caption         =   "Olor"
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
         Index           =   3
         Left            =   5360
         TabIndex        =   37
         Top             =   495
         Width           =   330
      End
      Begin VB.Label COA005_00 
         AutoSize        =   -1  'True
         Caption         =   "Color"
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
         Index           =   2
         Left            =   2895
         TabIndex        =   36
         Top             =   495
         Width           =   405
      End
      Begin VB.Label COA005_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Azúcares reductores (Benedict)"
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
         Index           =   5
         Left            =   105
         TabIndex        =   35
         Top             =   1200
         Width           =   2595
      End
      Begin VB.Label COA005_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sangre oculta (Thevenon)"
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
         Index           =   6
         Left            =   525
         TabIndex        =   34
         Top             =   1515
         Width           =   2175
      End
      Begin VB.Label COA005_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Grasas (Sudan)"
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
         Index           =   7
         Left            =   1425
         TabIndex        =   33
         Top             =   1845
         Width           =   1230
      End
   End
   Begin VB.Frame COA004 
      Caption         =   "Investigación de Criptosporidium y coccidias"
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
      Height          =   1280
      Left            =   60
      TabIndex        =   69
      Top             =   2340
      Visible         =   0   'False
      Width           =   7220
      Begin VB.TextBox COA004_01 
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
         Left            =   60
         TabIndex        =   4
         Top             =   590
         Width           =   7050
      End
      Begin VB.TextBox COA004_03 
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
         Left            =   1150
         TabIndex        =   5
         Top             =   910
         Width           =   5970
      End
      Begin VB.ComboBox COA004_00 
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
         ItemData        =   "frmParasitologia.frx":2E5D
         Left            =   60
         List            =   "frmParasitologia.frx":2E67
         TabIndex        =   3
         Text            =   "COA004_00"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label COA004_02 
         Caption         =   "Observación"
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
         Left            =   60
         TabIndex        =   70
         Top             =   940
         Width           =   1335
      End
   End
   Begin VB.Frame COA030 
      Caption         =   "Sangre Oculta (Thevenon)"
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
      Height          =   1035
      Left            =   60
      TabIndex        =   50
      Top             =   2340
      Visible         =   0   'False
      Width           =   7200
      Begin VB.ComboBox COA030_01 
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
         ItemData        =   "frmParasitologia.frx":2E7F
         Left            =   1950
         List            =   "frmParasitologia.frx":2E89
         TabIndex        =   22
         Text            =   "COA030_01"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox COA030_02 
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
         Left            =   3225
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox COA030_03 
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
         Left            =   1420
         TabIndex        =   24
         Top             =   615
         Width           =   5670
      End
      Begin VB.Label COA030_00 
         AutoSize        =   -1  'True
         Caption         =   "Sangre oculta"
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   270
         Width           =   1200
      End
      Begin VB.Label COA030_00 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
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
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   51
         Top             =   645
         Width           =   1275
      End
   End
   Begin VB.Frame COA031 
      Caption         =   "Reacción Inflamatoria"
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
      Height          =   1515
      Left            =   60
      TabIndex        =   47
      Top             =   2340
      Visible         =   0   'False
      Width           =   7200
      Begin VB.TextBox COA031_02 
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
         Left            =   1065
         TabIndex        =   26
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox COA031_04 
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
         Left            =   5760
         TabIndex        =   28
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox COA031_03 
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
         Left            =   3225
         TabIndex        =   27
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox COA031_05 
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
         Left            =   1420
         TabIndex        =   29
         Top             =   1095
         Width           =   5670
      End
      Begin VB.ComboBox COA031_01 
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
         ItemData        =   "frmParasitologia.frx":2EA1
         Left            =   2085
         List            =   "frmParasitologia.frx":2EAB
         TabIndex        =   25
         Text            =   "COA031_01"
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label COA031_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Index           =   7
         Left            =   6660
         TabIndex        =   59
         Top             =   750
         Width           =   180
      End
      Begin VB.Label COA031_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Index           =   5
         Left            =   4200
         TabIndex        =   58
         Top             =   750
         Width           =   180
      End
      Begin VB.Label COA031_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "x c"
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
         Index           =   3
         Left            =   2070
         TabIndex        =   57
         Top             =   750
         Width           =   240
      End
      Begin VB.Label COA031_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PMN"
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
         Index           =   4
         Left            =   2790
         TabIndex        =   56
         Top             =   750
         Width           =   360
      End
      Begin VB.Label COA031_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Leucocitos"
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
         Index           =   2
         Left            =   120
         TabIndex        =   55
         Top             =   750
         Width           =   870
      End
      Begin VB.Label COA031_00 
         AutoSize        =   -1  'True
         Caption         =   "MN"
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
         Index           =   6
         Left            =   5445
         TabIndex        =   54
         Top             =   750
         Width           =   255
      End
      Begin VB.Label COA031_00 
         AutoSize        =   -1  'True
         Caption         =   "Reacción Inflamatoria"
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
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   49
         Top             =   360
         Width           =   1875
      End
      Begin VB.Label COA031_00 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
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
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   48
         Top             =   1125
         Width           =   1275
      End
   End
   Begin VB.Frame COA003 
      Caption         =   "Test de Graham (Oxiuros)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1280
      Left            =   60
      TabIndex        =   30
      Top             =   2340
      Visible         =   0   'False
      Width           =   7220
      Begin VB.TextBox COA003_01 
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
         Left            =   60
         TabIndex        =   1
         Top             =   590
         Width           =   7080
      End
      Begin VB.TextBox COA003_03 
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
         Left            =   1150
         TabIndex        =   2
         Top             =   910
         Width           =   5970
      End
      Begin VB.ComboBox COA003_00 
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
         ItemData        =   "frmParasitologia.frx":2EC3
         Left            =   60
         List            =   "frmParasitologia.frx":2ECD
         TabIndex        =   0
         Text            =   "COA003_00"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label COA003_02 
         Caption         =   "Observación"
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
         Left            =   60
         TabIndex        =   31
         Top             =   940
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmParasitologia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Resultado para Parasitología
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_cmbResponsable As New sighentidades.ListaDespleglable
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idUsuario As Long
Dim ml_idOrden As Long
Dim ml_nombrePrueba As String
Dim ml_idAnalisis As Long
Dim ml_idPaciente As Long
Dim ml_resultado As String
Dim ml_observacion As String
Dim ml_IdMovimiento As Long
Dim ms_MensajeError As String
Dim ml_nombreMedico As String
Dim ml_nombrePaciente As String
Dim ml_nombreRealiza As Long
Dim ml_areaTrabajo As Long
Dim ml_CodigoPruebaSeleccionada As String
Dim ml_idPrueba As String
Dim ml_DetalleOrden As New ADODB.Recordset
Dim ml_idOrdenLab As Long
Dim ml_FechaNacimiento As Date
Dim ml_idTipoSexo As Long
Dim ml_NoMuestraBotonGrabar As Boolean
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS  As Long
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let NoMuestraBotonGrabar(lValue As Boolean)
   ml_NoMuestraBotonGrabar = lValue
   If ml_NoMuestraBotonGrabar = True Then
      cmdGrabar.Visible = False
   End If
End Property




Property Let idTipoSexo(lValue As Long)
    ml_idTipoSexo = lValue
End Property
Property Let FechaNacimiento(lValue As Date)
    ml_FechaNacimiento = lValue
End Property
Property Let idOrdenLab(lValue As Long)
   ml_idOrdenLab = lValue
End Property

Property Let CodigoPruebaSeleccionada(lValue As String)
   ml_CodigoPruebaSeleccionada = lValue
End Property

Property Let DetalleOrden(lValue As ADODB.Recordset)
  Set ml_DetalleOrden = lValue
End Property

Sub CargaDataCombos()
  mo_cmbResponsable.BoundColumn = "idEmpleado"
  mo_cmbResponsable.ListField = "ApNom"
  'Set mo_cmbResponsable.RowSource = mo_ReglasLaboratorio.EmpleadosDeLab(ml_areaTrabajo)
  Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =20")

  Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
  If mo_CabeceraReportes.NOpuedeModificarResponsable(sghAgregar, sighentidades.Usuario, mo_cmbResponsable.RowSource) Then
     mo_cmbResponsable.BoundText = Trim(Str(sighentidades.Usuario))
     Me.cmbResponsable.Enabled = False
  End If
  Set mo_CabeceraReportes = Nothing
End Sub

Property Let AreaTrabajo(lValue As Long)
    ml_areaTrabajo = lValue
End Property

Property Get AreaTrabajo() As Long
  AreaTrabajo = ml_areaTrabajo
End Property

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

Property Let idPrueba(lValue As String)
   ml_idPrueba = lValue
End Property

Property Get idPrueba() As String
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

Property Let nombreMedico(lValue As String)
   ml_nombreMedico = lValue
End Property

Property Get nombreMedico() As String
   nombreMedico = ml_nombreMedico
End Property

Property Let nombrePaciente(lValue As String)
   ml_nombrePaciente = lValue
End Property

Property Get nombrePaciente() As String
   nombrePaciente = ml_nombrePaciente
End Property

Sub AdministrarKeyPreview(KeyCode As Integer)
  Select Case KeyCode
    Case vbKeyReturn
      SendKeys "{TAB}"
    Case vbKeyF3
      cmdImprimir_Click
    Case vbKeyEscape
      cmdCancelar_Click
    Case vbKeyF2
      cmdGrabar_Click
  End Select
End Sub

Private Sub TopBoton(Fra As Frame)
  'If EmpleadoTrabajaEnLaboratorio(sighEntidades.Usuario) = True Then
    Fra.Enabled = True
  'Else
  '  Fra.Enabled = False
  'End If
  Fra.Visible = True
  Fra.Caption = ml_nombrePrueba
  Me.Caption = Fra.Caption
  Fra.Top = UcPacienteDatos1.Top + UcPacienteDatos1.Height + 500   '350
  fraBoton.Top = Fra.Top + Fra.Height
  Me.Height = fraBoton.Top + fraBoton.Height + 500 '455
End Sub

Private Sub cmbResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdGrabar_Click()
  If cmbResponsable.Text = "" Then
    MsgBox "Debe Seleccionar el personal que realizó la prueba", vbInformation, "SIGH "
    cmbResponsable.SetFocus
    Exit Sub
  End If
  If Me.txtFresultado.Text = sighentidades.FECHA_VACIA_DMY Then
    MsgBox "Por favor ingresar la Fecha del Resultado", vbInformation, "SIGH "
    Exit Sub
  End If
  ml_nombreRealiza = mo_cmbResponsable.BoundText
  If ml_CodigoPruebaSeleccionada = "COA003" Then  'Test de Graham
    'COA003
    ml_resultado = COA003_00.Text & "\" & COA003_01.Text
    ml_observacion = COA003_03.Text
  ElseIf ml_CodigoPruebaSeleccionada = "COA004" Then 'Investigación de Criptosporidium y Coccidias
    'COA004
    ml_resultado = COA004_00.Text & "\" & COA004_01.Text
    ml_observacion = COA004_03.Text
  ElseIf ml_CodigoPruebaSeleccionada = "COA031" Then 'Reacción Inflamatoria
    'COA031
    ml_resultado = COA031_01.Text & "\" & COA031_02.Text & "\" & COA031_03.Text & "\" & COA031_04.Text
    ml_observacion = COA031_05.Text
  ElseIf ml_CodigoPruebaSeleccionada = "COA005" Then 'Coprofuncional de Heces
    'COA005
    ml_resultado = COA005_02.Text & "\" & COA005_04.Text & "\" & COA005_06.Text & "\" & COA005_09.Text & "\" & COA005_11.Text & "\" & COA005_12.Text & "\" & COA005_14.Text & "\" & COA005_16.Text & "\" & COA005_19.Text & "\" & COA005_23.Text & "\" & COA005_25.Text & "\" & COA005_26.Text & "\" & COA005_27.Text & "\" & COA005_28.Text & "\" & COA005_30.Text
    ml_observacion = COA005_29.Text
  ElseIf ml_CodigoPruebaSeleccionada = "COA030" Then 'Sangre oculta
    'COA030
    ml_resultado = COA030_01.Text & "\" & COA030_02.Text
    ml_observacion = COA030_03.Text
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
  mo_ReglasLaboratorio.LabIngresaResultados idPrueba, idOrden, ml_resultado, ml_observacion, idUsuario, _
                       ml_nombreRealiza, ml_DetalleOrden, ml_idOrdenLab, "", "", 0, _
                       CDate(Me.txtFresultado.Text), mo_lcNombrePc, mo_lnIdTablaLISTBARITEMS, _
                       Me.UcPacienteDatos1.DevuelveHistoriaApellidosYnombre, Me.Caption, , _
                       IIf(Len(Trim(COA005_23.Text)) = 0, 0, IIf(COA005_23.ListIndex = 0, 1, 2)) '1-parasitosis positiva, 2-parasitosis negativa
End Sub

Private Sub cmdImprimir_Click()
  ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(idPrueba, idOrden)
  ml_observacion = mo_ReglasLaboratorio.LabRecuperaResultados_Obs(idPrueba, idOrden)
  Dim ldFechaResultado As Date
  ml_nombreRealiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(idPrueba, idOrden, ldFechaResultado)
  If ml_CodigoPruebaSeleccionada <> "" And Trim(ml_resultado) <> "" Then
    mo_ReglasLaboratorio.LabImprimeCabeceraResultados UcPacienteDatos1.idPaciente, nombrePaciente, UcPacienteDatos1.NroHistoriaClinica, ldFechaResultado, _
                         nombreMedico + mo_ReglasLaboratorio.DevuelveDatosParaImpresionResultadoLaboratorio(ml_idOrden)
    mo_ReglasLaboratorio.LabImprimeResultadosCOA ml_resultado, CStr(ml_CodigoPruebaSeleccionada), Me.Caption, ml_observacion, ml_nombreRealiza
    mo_ReglasLaboratorio.LabImprimePieResultados
  Else
    MsgBox "Debe grabar los resultados antes de poder imprimirlos", vbInformation, ""
  End If
End Sub

Private Sub COA003_00_Click()
  If COA003_00.ListIndex = 0 Then
    COA003_01.Text = "Se observan huevos de Enterobius vermicularis: 1+"
  Else
    COA003_01.Text = ""
  End If
End Sub

Private Sub COA003_00_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA003_01_GotFocus()
  SeleccionaTexto COA003_01
End Sub

Private Sub COA003_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA003_03_GotFocus()
  SeleccionaTexto COA003_03
End Sub

Private Sub COA003_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA004_00_Click()
  If COA004_00.ListIndex = 0 Then
    COA004_01.Text = "Se observan "
  Else
    COA004_01.Text = ""
  End If
End Sub

Private Sub COA004_00_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA004_01_GotFocus()
  SeleccionaTexto COA004_01
End Sub

Private Sub COA004_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA004_03_GotFocus()
  SeleccionaTexto COA004_03
End Sub

Private Sub COA004_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_02_GotFocus()
  SeleccionaTexto COA005_02
End Sub

Private Sub COA005_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_04_GotFocus()
  SeleccionaTexto COA005_04
End Sub

Private Sub COA005_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_06_GotFocus()
  SeleccionaTexto COA005_06
End Sub

Private Sub COA005_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_11_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_12_GotFocus()
  SeleccionaTexto COA005_12
End Sub

Private Sub COA005_12_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_14_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_16_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_19_GotFocus()
  SeleccionaTexto COA005_19
End Sub

Private Sub COA005_19_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_23_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_25_GotFocus()
  SeleccionaTexto COA005_25
End Sub

Private Sub COA005_25_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_26_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_27_GotFocus()
  SeleccionaTexto COA005_27
End Sub

Private Sub COA005_27_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_28_GotFocus()
  SeleccionaTexto COA005_28
End Sub

Private Sub COA005_28_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_29_GotFocus()
  SeleccionaTexto COA005_29
End Sub

Private Sub COA005_29_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA005_30_GotFocus()
  SeleccionaTexto COA005_30
End Sub

Private Sub COA005_30_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA030_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA030_02_GotFocus()
  SeleccionaTexto COA030_02
End Sub

Private Sub COA030_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA030_03_GotFocus()
  SeleccionaTexto COA030_03
End Sub

Private Sub COA030_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA031_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA031_02_GotFocus()
  SeleccionaTexto COA031_02
End Sub

Private Sub COA031_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA031_03_GotFocus()
  SeleccionaTexto COA031_03
End Sub

Private Sub COA031_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA031_04_GotFocus()
  SeleccionaTexto COA031_04
End Sub

Private Sub COA031_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub COA031_05_GotFocus()
  SeleccionaTexto COA031_05
End Sub

Private Sub COA031_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Initialize()
  Set mo_cmbResponsable.MiComboBox = cmbResponsable
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
  Me.txtFresultado.Text = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
  LimpiaVAloresDefault
  Me.UcPacienteDatos1.idPaciente = ml_idPaciente
  Me.UcPacienteDatos1.FechaRegistro = Now
  If ml_idPaciente = 0 Then
     Me.UcPacienteDatos1.idTipoSexo = ml_idTipoSexo
     Me.UcPacienteDatos1.FechaNacimiento = ml_FechaNacimiento
     Me.UcPacienteDatos1.CargaAlgunosDatosDesdeBoleta ml_nombrePaciente
  Else
     Me.UcPacienteDatos1.CargarDatosDePacienteALosControles
  End If
  Me.UcPacienteDatos1.DeshabilitarFrames True
  CargaDataCombos
  cmbResponsable.ListIndex = Ubica_En_Combo(cmbResponsable, sighentidades.NombreUsuario)
  'If EmpleadoTrabajaEnLaboratorio(sighEntidades.Usuario) = True Then
    cmdGrabar.Enabled = True
  'Else
  '  cmdGrabar.Enabled = False
  'End If
  ml_resultado = ""
  ml_observacion = ""
  
  If ml_CodigoPruebaSeleccionada = "COA003" Then  'Test de Graham
    TopBoton COA003
  ElseIf ml_CodigoPruebaSeleccionada = "COA004" Then 'Investigación de Criptosporidium y Coccidias
    TopBoton COA004
  ElseIf ml_CodigoPruebaSeleccionada = "COA031" Then 'Reacción Inflamatoria
    TopBoton COA031
  ElseIf ml_CodigoPruebaSeleccionada = "COA005" Then 'Coprofuncional de Heces
    TopBoton COA005
  ElseIf ml_CodigoPruebaSeleccionada = "COA030" Then 'Sangre oculta
    TopBoton COA030
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
  End If
  'Recupera información si es que ya esta grabado
  ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(idPrueba, idOrden)
  Dim ldFechaResultado As Date
  ml_nombreRealiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(idPrueba, idOrden, ldFechaResultado)
  If ml_resultado = "" Or Val(ml_nombreRealiza) = 0 Then Exit Sub
  Me.txtFresultado.Text = Format(IIf(ldFechaResultado = 0, Now, ldFechaResultado), sighentidades.DevuelveFechaSoloFormato_DMY_HM)
  cmbResponsable.ListIndex = Ubica_En_Combo(cmbResponsable, mo_ReglasLaboratorio.LabEmpleado(ml_nombreRealiza))
  'If cmbResponsable.Text <> "" Then cmbResponsable.Enabled = False
  ml_observacion = mo_ReglasLaboratorio.LabRecuperaResultados_Obs(idPrueba, idOrden)
  Dim Temp As String
  'Asigna la información recuperada en el formulario
  If ml_CodigoPruebaSeleccionada = "COA003" Then  'Test de Graham
    'COA003
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA003_00.ListIndex = Ubica_En_Combo(COA003_00, Temp)
    COA003_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA003_03.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "COA004" Then 'Investigación de Criptosporidium y Coccidias
    'COA004
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA004_00.ListIndex = Ubica_En_Combo(COA004_00, Temp)
    COA004_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA004_03.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "COA031" Then 'Reacción Inflamatoria
    'COA031
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA031_01.ListIndex = Ubica_En_Combo(COA031_01, Temp)
    COA031_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA031_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA031_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA031_05.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "COA005" Then 'Coprofuncional de Heces
    'COA005
    COA005_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA005_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA005_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA005_09.ListIndex = Ubica_En_Combo(COA005_09, Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA005_11.ListIndex = Ubica_En_Combo(COA005_11, Temp)
    COA005_12.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA005_14.ListIndex = Ubica_En_Combo(COA005_14, Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA005_16.ListIndex = Ubica_En_Combo(COA005_16, Temp)
    COA005_19.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA005_23.ListIndex = Ubica_En_Combo(COA005_23, Temp)
    COA005_25.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA005_26.ListIndex = Ubica_En_Combo(COA005_26, Temp)
    COA005_27.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA005_28.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA005_30.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA005_29.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "COA030" Then 'Sangre oculta
    'COA030
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA030_01.ListIndex = Ubica_En_Combo(COA030_01, Temp)
    COA030_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    COA030_03.Text = ml_observacion
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
  End If
End Sub

Sub LimpiaVAloresDefault()
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       Exit Sub
    End If
    COA005_02.Text = ""
    COA005_04.Text = ""
    COA005_06.Text = ""
    COA003_00.Text = ""
    COA004_00.Text = ""
    COA030_01.Text = ""
    COA031_01.Text = ""
End Sub

