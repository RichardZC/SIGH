VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmHematologia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HEMATOLOGÍA"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmHematologia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   212
      Top             =   1740
      Width           =   7185
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
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   213
         Top             =   180
         Width           =   3090
      End
      Begin MSMask.MaskEdBox txtFresultado 
         Height          =   315
         Left            =   5580
         TabIndex        =   214
         Top             =   180
         Width           =   1470
         _ExtentX        =   2593
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
         Left            =   150
         TabIndex        =   216
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
         Left            =   4605
         TabIndex        =   215
         Top             =   225
         Width           =   945
      End
   End
   Begin VB.Frame fraBoton 
      ForeColor       =   &H00000000&
      Height          =   870
      Left            =   60
      TabIndex        =   207
      Top             =   5790
      Width           =   7215
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmHematologia.frx":0CCA
         DownPicture     =   "frmHematologia.frx":118E
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
         Left            =   3705
         Picture         =   "frmHematologia.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   210
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprime (F3)"
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
         Picture         =   "frmHematologia.frx":1B66
         Style           =   1  'Graphical
         TabIndex        =   209
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmHematologia.frx":203F
         DownPicture     =   "frmHematologia.frx":249F
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
         Left            =   2265
         Picture         =   "frmHematologia.frx":2914
         Style           =   1  'Graphical
         TabIndex        =   208
         Top             =   180
         Width           =   1365
      End
   End
   Begin SIGHLaboratorio.UcPacienteDatos1 UcPacienteDatos1 
      Height          =   1695
      Left            =   60
      TabIndex        =   206
      Top             =   0
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   2990
   End
   Begin VB.Frame HEM002 
      Caption         =   "Hemoglobina"
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
      Height          =   1120
      Left            =   60
      TabIndex        =   110
      Top             =   2400
      Visible         =   0   'False
      Width           =   7185
      Begin VB.ComboBox HEM002_03 
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
         ItemData        =   "frmHematologia.frx":2D89
         Left            =   3000
         List            =   "frmHematologia.frx":2D93
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   390
         Width           =   1215
      End
      Begin VB.TextBox HEM002_05 
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
         Left            =   5820
         TabIndex        =   19
         Text            =   "Enzimático"
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox HEM002_04 
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
         Left            =   4320
         TabIndex        =   18
         Text            =   "13 - 15 gr %"
         Top             =   420
         Width           =   1215
      End
      Begin VB.TextBox HEM002_06 
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
         Left            =   1365
         TabIndex        =   20
         Top             =   750
         Width           =   5670
      End
      Begin VB.TextBox HEM002_02 
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
         Left            =   1380
         MaxLength       =   5
         TabIndex        =   16
         Top             =   420
         Width           =   675
      End
      Begin VB.TextBox HEM002_01 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Hemoglobina"
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label HEM002_00 
         AutoSize        =   -1  'True
         Caption         =   "gr %"
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
         Left            =   2100
         TabIndex        =   116
         Top             =   450
         Width           =   405
      End
      Begin VB.Label HEM002_00 
         Alignment       =   2  'Center
         Caption         =   "Método"
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
         Index           =   3
         Left            =   6000
         TabIndex        =   115
         Top             =   200
         Width           =   975
      End
      Begin VB.Label HEM002_00 
         Alignment       =   2  'Center
         Caption         =   "Valor Referencial"
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
         Index           =   2
         Left            =   3000
         TabIndex        =   114
         Top             =   195
         Width           =   2535
      End
      Begin VB.Label HEM002_00 
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
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   113
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label HEM002_00 
         Alignment       =   2  'Center
         Caption         =   "Resultado"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   112
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label HEM002_00 
         Alignment       =   2  'Center
         Caption         =   "Dosaje de "
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
         Index           =   0
         Left            =   180
         TabIndex        =   111
         Top             =   195
         Width           =   975
      End
   End
   Begin VB.Frame HEM003 
      Caption         =   "VSG (Velocidad de Sedimentación Globular)"
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
      Height          =   1100
      Left            =   60
      TabIndex        =   168
      Top             =   2370
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox HEM003_01 
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
         Left            =   120
         MaxLength       =   5
         TabIndex        =   21
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox HEM003_04 
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
         Left            =   1365
         TabIndex        =   24
         Top             =   740
         Width           =   5790
      End
      Begin VB.TextBox HEM003_02 
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
         Left            =   3120
         TabIndex        =   22
         Text            =   "0 - 20 mm / hora"
         Top             =   420
         Width           =   1575
      End
      Begin VB.TextBox HEM003_03 
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
         Left            =   5820
         TabIndex        =   23
         Text            =   "Wintrobe"
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label HEM003_00 
         Alignment       =   2  'Center
         Caption         =   "Resultado"
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
         Index           =   0
         Left            =   120
         TabIndex        =   173
         Top             =   195
         Width           =   1575
      End
      Begin VB.Label HEM003_00 
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
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   172
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label HEM003_00 
         Alignment       =   2  'Center
         Caption         =   "Valor Referencial"
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
         Index           =   1
         Left            =   2520
         TabIndex        =   171
         Top             =   195
         Width           =   2775
      End
      Begin VB.Label HEM003_00 
         Alignment       =   2  'Center
         Caption         =   "Método"
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
         Index           =   2
         Left            =   5880
         TabIndex        =   170
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label HEM003_00 
         AutoSize        =   -1  'True
         Caption         =   "mm / hora"
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
         Left            =   780
         TabIndex        =   169
         Top             =   450
         Width           =   855
      End
   End
   Begin VB.Frame HEM004 
      Caption         =   "Recuento de Reticulocitos"
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
      Height          =   1100
      Left            =   60
      TabIndex        =   150
      Top             =   2370
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox HEM004_03 
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
         Left            =   4740
         TabIndex        =   27
         Text            =   "Coloración Azul Cresil Brillante"
         Top             =   420
         Width           =   2295
      End
      Begin VB.TextBox HEM004_02 
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
         Left            =   2880
         TabIndex        =   26
         Text            =   "0.5 - 2.5 %"
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox HEM004_04 
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
         Left            =   1365
         TabIndex        =   28
         Top             =   740
         Width           =   5670
      End
      Begin VB.TextBox HEM004_01 
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
         Left            =   480
         MaxLength       =   5
         TabIndex        =   25
         Top             =   420
         Width           =   615
      End
      Begin VB.Label HEM004_00 
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
         Index           =   3
         Left            =   1140
         TabIndex        =   155
         Top             =   450
         Width           =   180
      End
      Begin VB.Label HEM004_00 
         Alignment       =   2  'Center
         Caption         =   "Método"
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
         Index           =   2
         Left            =   4800
         TabIndex        =   154
         Top             =   195
         Width           =   2175
      End
      Begin VB.Label HEM004_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor Referencial"
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
         Index           =   1
         Left            =   2520
         TabIndex        =   153
         Top             =   195
         Width           =   1575
      End
      Begin VB.Label HEM004_00 
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
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   152
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label HEM004_00 
         Alignment       =   2  'Center
         Caption         =   "Resultado"
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
         Index           =   0
         Left            =   120
         TabIndex        =   151
         Top             =   195
         Width           =   1575
      End
   End
   Begin VB.Frame HEM005 
      Caption         =   "Recuento de Plaquetas"
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
      Height          =   1100
      Left            =   60
      TabIndex        =   145
      Top             =   2370
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox HEM005_01 
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
         Left            =   240
         MaxLength       =   7
         TabIndex        =   29
         Top             =   420
         Width           =   1575
      End
      Begin VB.TextBox HEM005_04 
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
         Left            =   1365
         TabIndex        =   32
         Top             =   740
         Width           =   5670
      End
      Begin VB.TextBox HEM005_02 
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
         Left            =   3120
         TabIndex        =   30
         Text            =   "150000 - 300000 mm3"
         Top             =   420
         Width           =   1815
      End
      Begin VB.TextBox HEM005_03 
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
         Left            =   5220
         TabIndex        =   31
         Text            =   "Coloración Rees Ecker"
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label HEM005_00 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "mm3"
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
         Left            =   1920
         TabIndex        =   211
         Top             =   480
         Width           =   465
      End
      Begin VB.Label HEM005_00 
         Alignment       =   2  'Center
         Caption         =   "Resultado"
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
         Index           =   0
         Left            =   240
         TabIndex        =   149
         Top             =   195
         Width           =   1575
      End
      Begin VB.Label HEM005_00 
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
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   148
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label HEM005_00 
         Alignment       =   2  'Center
         Caption         =   "Valor Referencial"
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
         Index           =   1
         Left            =   3120
         TabIndex        =   147
         Top             =   195
         Width           =   1815
      End
      Begin VB.Label HEM005_00 
         Alignment       =   2  'Center
         Caption         =   "Método"
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
         Index           =   2
         Left            =   5280
         TabIndex        =   146
         Top             =   195
         Width           =   1695
      End
   End
   Begin VB.Frame HEM006 
      Caption         =   "Tiempo de Coagulación y Sangría"
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
      Height          =   1420
      Left            =   60
      TabIndex        =   156
      Top             =   2370
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox HEM006_09 
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
         Left            =   6300
         TabIndex        =   40
         Text            =   "Duke"
         Top             =   735
         Width           =   855
      End
      Begin VB.TextBox HEM006_01 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   158
         TabStop         =   0   'False
         Text            =   "Tiempo de Coagulación"
         Top             =   420
         Width           =   1935
      End
      Begin VB.TextBox HEM006_08 
         Alignment       =   2  'Center
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
         Left            =   4800
         TabIndex        =   39
         Text            =   "1 - 4 min"
         Top             =   735
         Width           =   975
      End
      Begin VB.TextBox HEM006_06 
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
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   37
         Top             =   735
         Width           =   495
      End
      Begin VB.TextBox HEM006_01 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   157
         TabStop         =   0   'False
         Text            =   "Tiempo de Sangría"
         Top             =   735
         Width           =   1935
      End
      Begin VB.TextBox HEM006_02 
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
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   33
         Top             =   420
         Width           =   495
      End
      Begin VB.TextBox HEM006_10 
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
         Left            =   1365
         TabIndex        =   41
         Top             =   1060
         Width           =   5670
      End
      Begin VB.TextBox HEM006_04 
         Alignment       =   2  'Center
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
         Left            =   4800
         TabIndex        =   35
         Text            =   "4 - 8 min"
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox HEM006_05 
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
         Left            =   6300
         TabIndex        =   36
         Text            =   "Burker"
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox HEM006_03 
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
         MaxLength       =   5
         TabIndex        =   34
         Top             =   420
         Width           =   495
      End
      Begin VB.TextBox HEM006_07 
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
         MaxLength       =   5
         TabIndex        =   38
         Top             =   735
         Width           =   495
      End
      Begin VB.Label HEM006_00 
         AutoSize        =   -1  'True
         Caption         =   "min"
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
         Left            =   2940
         TabIndex        =   166
         Top             =   765
         Width           =   285
      End
      Begin VB.Label HEM006_00 
         Alignment       =   2  'Center
         Caption         =   "Resultado"
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
         Index           =   1
         Left            =   2400
         TabIndex        =   165
         Top             =   195
         Width           =   1815
      End
      Begin VB.Label HEM006_00 
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
         Height          =   255
         Index           =   8
         Left            =   60
         TabIndex        =   164
         Top             =   1090
         Width           =   1335
      End
      Begin VB.Label HEM006_00 
         Alignment       =   2  'Center
         Caption         =   "Valor Referencial"
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
         Index           =   2
         Left            =   4440
         TabIndex        =   163
         Top             =   195
         Width           =   1695
      End
      Begin VB.Label HEM006_00 
         Alignment       =   2  'Center
         Caption         =   "Método"
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
         Index           =   3
         Left            =   6360
         TabIndex        =   162
         Top             =   195
         Width           =   735
      End
      Begin VB.Label HEM006_00 
         AutoSize        =   -1  'True
         Caption         =   "min"
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
         Left            =   2940
         TabIndex        =   161
         Top             =   450
         Width           =   285
      End
      Begin VB.Label HEM006_00 
         AutoSize        =   -1  'True
         Caption         =   "seg"
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
         Left            =   3900
         TabIndex        =   160
         Top             =   450
         Width           =   285
      End
      Begin VB.Label HEM006_00 
         AutoSize        =   -1  'True
         Caption         =   "seg"
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
         Left            =   3900
         TabIndex        =   159
         Top             =   765
         Width           =   285
      End
      Begin VB.Label HEM006_00 
         Alignment       =   2  'Center
         Caption         =   "Dosaje de "
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
         Index           =   0
         Left            =   180
         TabIndex        =   167
         Top             =   195
         Width           =   1815
      End
   End
   Begin VB.Frame HEM008 
      Caption         =   "Leishmania"
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
      Height          =   1100
      Left            =   60
      TabIndex        =   174
      Top             =   2400
      Visible         =   0   'False
      Width           =   7185
      Begin VB.ComboBox HEM008_01 
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
         ItemData        =   "frmHematologia.frx":2DA5
         Left            =   120
         List            =   "frmHematologia.frx":2DAF
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   380
         Width           =   1215
      End
      Begin VB.TextBox HEM008_02 
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
         Left            =   1440
         TabIndex        =   44
         Top             =   400
         Width           =   3495
      End
      Begin VB.TextBox HEM008_04 
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
         Left            =   1455
         TabIndex        =   46
         Top             =   720
         Width           =   5700
      End
      Begin VB.TextBox HEM008_03 
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
         Left            =   5820
         TabIndex        =   45
         Text            =   "Coloración"
         Top             =   400
         Width           =   1335
      End
      Begin VB.Label HEM008_00 
         Alignment       =   2  'Center
         Caption         =   "Resultado"
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
         Index           =   0
         Left            =   120
         TabIndex        =   177
         Top             =   195
         Width           =   4815
      End
      Begin VB.Label HEM008_00 
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
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   176
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label HEM008_00 
         Alignment       =   2  'Center
         Caption         =   "Método"
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
         Index           =   1
         Left            =   6000
         TabIndex        =   175
         Top             =   200
         Width           =   975
      End
   End
   Begin VB.Frame HEM009 
      Caption         =   "Células L.E."
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
      Height          =   1100
      Left            =   60
      TabIndex        =   124
      Top             =   2370
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox HEM009_03 
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
         Left            =   5820
         TabIndex        =   128
         Text            =   "Coloración"
         Top             =   400
         Width           =   1335
      End
      Begin VB.TextBox HEM009_04 
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
         Left            =   1425
         TabIndex        =   127
         Top             =   720
         Width           =   5730
      End
      Begin VB.TextBox HEM009_02 
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
         Left            =   1440
         TabIndex        =   126
         Top             =   400
         Width           =   3495
      End
      Begin VB.ComboBox HEM009_01 
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
         ItemData        =   "frmHematologia.frx":2DC7
         Left            =   120
         List            =   "frmHematologia.frx":2DD1
         Style           =   2  'Dropdown List
         TabIndex        =   125
         Top             =   380
         Width           =   1215
      End
      Begin VB.Label HEM009_00 
         Alignment       =   2  'Center
         Caption         =   "Método"
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
         Index           =   1
         Left            =   6000
         TabIndex        =   131
         Top             =   200
         Width           =   975
      End
      Begin VB.Label HEM009_00 
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
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   130
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label HEM009_00 
         Alignment       =   2  'Center
         Caption         =   "Resultado"
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
         Index           =   0
         Left            =   120
         TabIndex        =   129
         Top             =   195
         Width           =   4815
      End
   End
   Begin VB.Frame HEM010 
      Caption         =   "Lámina Periférica"
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
      Height          =   2055
      Left            =   60
      TabIndex        =   132
      Top             =   2370
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox HEM010_07 
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
         Left            =   1365
         TabIndex        =   55
         Top             =   1680
         Width           =   5670
      End
      Begin VB.TextBox HEM010_01 
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
         Left            =   1980
         MaxLength       =   5
         TabIndex        =   49
         Top             =   400
         Width           =   1095
      End
      Begin VB.TextBox HEM010_02 
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
         Left            =   5700
         MaxLength       =   5
         TabIndex        =   50
         Top             =   400
         Width           =   1095
      End
      Begin VB.TextBox HEM010_03 
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
         Left            =   1980
         MaxLength       =   5
         TabIndex        =   51
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox HEM010_04 
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
         Left            =   900
         TabIndex        =   52
         Top             =   1290
         Width           =   1215
      End
      Begin VB.TextBox HEM010_05 
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
         Left            =   3540
         TabIndex        =   53
         Top             =   1290
         Width           =   1095
      End
      Begin VB.TextBox HEM010_06 
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
         Left            =   5940
         TabIndex        =   54
         Top             =   1290
         Width           =   1095
      End
      Begin VB.Label HEM010_00 
         Alignment       =   2  'Center
         Caption         =   "Serie Blanca"
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
         Index           =   0
         Left            =   30
         TabIndex        =   144
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label HEM010_00 
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
         Height          =   255
         Index           =   11
         Left            =   60
         TabIndex        =   143
         Top             =   1710
         Width           =   1335
      End
      Begin VB.Label HEM010_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Leucocitos típicos"
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
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   142
         Top             =   435
         Width           =   1695
      End
      Begin VB.Label HEM010_00 
         Alignment       =   2  'Center
         Caption         =   "Leucocitos atípicos"
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
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   141
         Top             =   435
         Width           =   1695
      End
      Begin VB.Label HEM010_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Blastos"
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
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   140
         Top             =   750
         Width           =   1695
      End
      Begin VB.Label HEM010_00 
         Alignment       =   2  'Center
         Caption         =   "Serie Roja"
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
         Index           =   7
         Left            =   30
         TabIndex        =   139
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label HEM010_00 
         Alignment       =   2  'Center
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
         Index           =   8
         Left            =   420
         TabIndex        =   138
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label HEM010_00 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Tamaño"
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
         Index           =   9
         Left            =   2820
         TabIndex        =   137
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label HEM010_00 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Forma"
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
         Left            =   5385
         TabIndex        =   136
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label HEM010_00 
         Alignment       =   2  'Center
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
         Index           =   2
         Left            =   3135
         TabIndex        =   135
         Top             =   435
         Width           =   195
      End
      Begin VB.Label HEM010_00 
         Alignment       =   2  'Center
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
         Index           =   6
         Left            =   3135
         TabIndex        =   134
         Top             =   750
         Width           =   195
      End
      Begin VB.Label HEM010_00 
         Alignment       =   2  'Center
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
         Index           =   4
         Left            =   6825
         TabIndex        =   133
         Top             =   435
         Width           =   195
      End
   End
   Begin VB.Frame HEM011 
      Caption         =   "Hematocrito"
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
      Height          =   1150
      Left            =   60
      TabIndex        =   117
      Top             =   2370
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox HEM011_01 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   "Hematocrito"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox HEM011_02 
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
         Left            =   1380
         MaxLength       =   5
         TabIndex        =   57
         Top             =   420
         Width           =   675
      End
      Begin VB.TextBox HEM011_06 
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
         Left            =   1365
         TabIndex        =   82
         Top             =   750
         Width           =   5790
      End
      Begin VB.TextBox HEM011_04 
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
         Left            =   4320
         TabIndex        =   59
         Text            =   "43 - 49 %"
         Top             =   420
         Width           =   1215
      End
      Begin VB.TextBox HEM011_05 
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
         Left            =   5820
         TabIndex        =   81
         Text            =   "Microhematocrito"
         Top             =   420
         Width           =   1335
      End
      Begin VB.ComboBox HEM011_03 
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
         ItemData        =   "frmHematologia.frx":2DE9
         Left            =   3000
         List            =   "frmHematologia.frx":2DF3
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label HEM011_00 
         Alignment       =   2  'Center
         Caption         =   "Dosaje de "
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
         Index           =   0
         Left            =   180
         TabIndex        =   123
         Top             =   180
         Width           =   975
      End
      Begin VB.Label HEM011_00 
         Alignment       =   2  'Center
         Caption         =   "Resultado"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   122
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label HEM011_00 
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
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   121
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label HEM011_00 
         Alignment       =   2  'Center
         Caption         =   "Valor Referencial"
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
         Index           =   2
         Left            =   3000
         TabIndex        =   120
         Top             =   195
         Width           =   2535
      End
      Begin VB.Label HEM011_00 
         Alignment       =   2  'Center
         Caption         =   "Método"
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
         Index           =   3
         Left            =   6000
         TabIndex        =   119
         Top             =   200
         Width           =   975
      End
      Begin VB.Label HEM011_00 
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
         Index           =   4
         Left            =   2100
         TabIndex        =   118
         Top             =   450
         Width           =   180
      End
   End
   Begin VB.Frame HEM012 
      Caption         =   "Tiempo de Protombina"
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
      Height          =   1450
      Left            =   60
      TabIndex        =   182
      Top             =   2370
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox HEM012_05 
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
         Left            =   3720
         MaxLength       =   5
         TabIndex        =   64
         Top             =   750
         Width           =   615
      End
      Begin VB.TextBox HEM012_04 
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
         Left            =   1350
         MaxLength       =   5
         TabIndex        =   63
         Top             =   750
         Width           =   615
      End
      Begin VB.TextBox HEM012_03 
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
         Left            =   5790
         TabIndex        =   62
         Text            =   "Microhematocrito"
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox HEM012_02 
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
         Left            =   3720
         TabIndex        =   61
         Text            =   "10 - 15 seg"
         Top             =   420
         Width           =   1215
      End
      Begin VB.TextBox HEM012_06 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1365
         TabIndex        =   65
         Top             =   1080
         Width           =   5790
      End
      Begin VB.TextBox HEM012_01 
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
         Left            =   1350
         MaxLength       =   5
         TabIndex        =   60
         Top             =   420
         Width           =   615
      End
      Begin VB.Label HEM012_00 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "INR"
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
         Left            =   3330
         TabIndex        =   191
         Top             =   780
         Width           =   315
      End
      Begin VB.Label HEM012_00 
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
         Index           =   6
         Left            =   1980
         TabIndex        =   190
         Top             =   780
         Width           =   180
      End
      Begin VB.Label HEM012_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Protombina"
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
         Left            =   390
         TabIndex        =   189
         Top             =   780
         Width           =   945
      End
      Begin VB.Label HEM012_00 
         AutoSize        =   -1  'True
         Caption         =   "segundos"
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
         Left            =   1980
         TabIndex        =   188
         Top             =   450
         Width           =   780
      End
      Begin VB.Label HEM012_00 
         Alignment       =   2  'Center
         Caption         =   "Método"
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
         Index           =   2
         Left            =   6000
         TabIndex        =   187
         Top             =   200
         Width           =   975
      End
      Begin VB.Label HEM012_00 
         Alignment       =   2  'Center
         Caption         =   "Valor Referencial"
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
         Index           =   1
         Left            =   3000
         TabIndex        =   186
         Top             =   195
         Width           =   2535
      End
      Begin VB.Label HEM012_00 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   8
         Left            =   30
         TabIndex        =   185
         Top             =   1110
         Width           =   1335
      End
      Begin VB.Label HEM012_00 
         Alignment       =   2  'Center
         Caption         =   "Resultado"
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
         Index           =   0
         Left            =   1170
         TabIndex        =   184
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label HEM012_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Tiempo"
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
         Height          =   255
         Index           =   3
         Left            =   390
         TabIndex        =   183
         Top             =   450
         Width           =   975
      End
   End
   Begin VB.Frame HEM013 
      Caption         =   "Grupo y Factor Sanguíneo"
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
      Height          =   975
      Left            =   60
      TabIndex        =   178
      Top             =   2370
      Visible         =   0   'False
      Width           =   7185
      Begin VB.ComboBox HEM013_01 
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
         ItemData        =   "frmHematologia.frx":2E05
         Left            =   1650
         List            =   "frmHematologia.frx":2E15
         TabIndex        =   66
         Text            =   "HEM013_01"
         Top             =   220
         Width           =   1215
      End
      Begin VB.ComboBox HEM013_02 
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
         ItemData        =   "frmHematologia.frx":2E26
         Left            =   5550
         List            =   "frmHematologia.frx":2E30
         TabIndex        =   67
         Text            =   "HEM013_02"
         Top             =   220
         Width           =   1215
      End
      Begin VB.TextBox HEM013_03 
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
         Left            =   1365
         TabIndex        =   68
         Top             =   600
         Width           =   5670
      End
      Begin VB.Label HEM013_00 
         Caption         =   "Grupo Sanguíneo"
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
         Index           =   0
         Left            =   60
         TabIndex        =   181
         Top             =   250
         Width           =   1695
      End
      Begin VB.Label HEM013_00 
         Caption         =   "Factor RH"
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
         Index           =   1
         Left            =   4620
         TabIndex        =   180
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label HEM013_00 
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
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   179
         Top             =   630
         Width           =   1335
      End
   End
   Begin VB.Frame HEM030 
      Caption         =   "Tiempo de Coagulación"
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
      Height          =   1185
      Left            =   60
      TabIndex        =   192
      Top             =   2370
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox HEM030_03 
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
         MaxLength       =   5
         TabIndex        =   77
         Top             =   420
         Width           =   495
      End
      Begin VB.TextBox HEM030_05 
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
         Left            =   6300
         TabIndex        =   79
         Text            =   "Burker"
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox HEM030_04 
         Alignment       =   2  'Center
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
         Left            =   4800
         TabIndex        =   78
         Text            =   "4 - 8 min"
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox HEM030_06 
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
         Left            =   1485
         TabIndex        =   80
         Top             =   750
         Width           =   5670
      End
      Begin VB.TextBox HEM030_02 
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
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   70
         Top             =   420
         Width           =   495
      End
      Begin VB.TextBox HEM030_01 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   69
         TabStop         =   0   'False
         Text            =   "Tiempo de Coagulación"
         Top             =   420
         Width           =   1935
      End
      Begin VB.Label HEM030_00 
         AutoSize        =   -1  'True
         Caption         =   "seg"
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
         Left            =   3900
         TabIndex        =   198
         Top             =   450
         Width           =   285
      End
      Begin VB.Label HEM030_00 
         AutoSize        =   -1  'True
         Caption         =   "min"
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
         Left            =   2940
         TabIndex        =   197
         Top             =   450
         Width           =   285
      End
      Begin VB.Label HEM030_00 
         Alignment       =   2  'Center
         Caption         =   "Método"
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
         Index           =   2
         Left            =   6360
         TabIndex        =   196
         Top             =   195
         Width           =   735
      End
      Begin VB.Label HEM030_00 
         Alignment       =   2  'Center
         Caption         =   "Valor Referencial"
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
         Index           =   1
         Left            =   4440
         TabIndex        =   195
         Top             =   195
         Width           =   1695
      End
      Begin VB.Label HEM030_00 
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
         Height          =   255
         Index           =   5
         Left            =   150
         TabIndex        =   194
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label HEM030_00 
         Alignment       =   2  'Center
         Caption         =   "Resultado"
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
         Index           =   0
         Left            =   2400
         TabIndex        =   193
         Top             =   195
         Width           =   1815
      End
   End
   Begin VB.Frame HEM031 
      Caption         =   "Tiempo de Sangría"
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
      Height          =   1080
      Left            =   60
      TabIndex        =   199
      Top             =   2400
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox HEM031_03 
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
         MaxLength       =   5
         TabIndex        =   73
         Top             =   375
         Width           =   495
      End
      Begin VB.TextBox HEM031_06 
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
         Left            =   1365
         TabIndex        =   76
         Top             =   705
         Width           =   5670
      End
      Begin VB.TextBox HEM031_01 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Text            =   "Tiempo de Sangría"
         Top             =   375
         Width           =   1935
      End
      Begin VB.TextBox HEM031_02 
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
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   72
         Top             =   375
         Width           =   495
      End
      Begin VB.TextBox HEM031_04 
         Alignment       =   2  'Center
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
         Left            =   4800
         TabIndex        =   74
         Text            =   "1 - 4 min"
         Top             =   375
         Width           =   975
      End
      Begin VB.TextBox HEM031_05 
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
         Left            =   6300
         TabIndex        =   75
         Text            =   "Duke"
         Top             =   375
         Width           =   855
      End
      Begin VB.Label HEM031_00 
         AutoSize        =   -1  'True
         Caption         =   "seg"
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
         Left            =   3900
         TabIndex        =   205
         Top             =   405
         Width           =   285
      End
      Begin VB.Label HEM031_00 
         Alignment       =   2  'Center
         Caption         =   "Método"
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
         Index           =   2
         Left            =   6360
         TabIndex        =   204
         Top             =   180
         Width           =   735
      End
      Begin VB.Label HEM031_00 
         Alignment       =   2  'Center
         Caption         =   "Valor Referencial"
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
         Index           =   1
         Left            =   4440
         TabIndex        =   203
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label HEM031_00 
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
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   202
         Top             =   735
         Width           =   1335
      End
      Begin VB.Label HEM031_00 
         Alignment       =   2  'Center
         Caption         =   "Resultado"
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
         Index           =   0
         Left            =   2400
         TabIndex        =   201
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label HEM031_00 
         AutoSize        =   -1  'True
         Caption         =   "min"
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
         Left            =   2940
         TabIndex        =   200
         Top             =   405
         Width           =   285
      End
   End
   Begin VB.Frame HEM001 
      Caption         =   "Hemograma Completo"
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
      Height          =   3405
      Left            =   60
      TabIndex        =   42
      Top             =   2370
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox HEM001_06 
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
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "6700"
         Top             =   1260
         Width           =   630
      End
      Begin VB.TextBox HEM001_03 
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
         Left            =   4710
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "47"
         Top             =   270
         Width           =   1110
      End
      Begin VB.TextBox HEM001_05 
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
         Left            =   4710
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "6700"
         Top             =   600
         Width           =   1110
      End
      Begin VB.TextBox HEM001_16 
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
         Left            =   1365
         TabIndex        =   14
         Top             =   3015
         Width           =   5790
      End
      Begin VB.TextBox HEM001_04 
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
         Left            =   1290
         MaxLength       =   9
         TabIndex        =   2
         Text            =   "4 800000"
         Top             =   570
         Width           =   1110
      End
      Begin VB.TextBox HEM001_01 
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
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   0
         Text            =   "13.5"
         Top             =   240
         Width           =   1110
      End
      Begin VB.TextBox HEM001_10 
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
         Left            =   6480
         MaxLength       =   5
         TabIndex        =   8
         Text            =   "0"
         Top             =   1560
         Width           =   390
      End
      Begin VB.TextBox HEM001_07 
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
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "2"
         Top             =   1590
         Width           =   390
      End
      Begin VB.TextBox HEM001_11 
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
         MaxLength       =   5
         TabIndex        =   9
         Text            =   "64"
         Top             =   1920
         Width           =   390
      End
      Begin VB.TextBox HEM001_08 
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
         Left            =   2970
         MaxLength       =   5
         TabIndex        =   6
         Text            =   "2"
         Top             =   1560
         Width           =   390
      End
      Begin VB.TextBox HEM001_09 
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
         Left            =   4710
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "28"
         Top             =   1560
         Width           =   390
      End
      Begin VB.Frame HEM001_02 
         Caption         =   "Neutrófilos           %"
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
         Height          =   975
         Left            =   90
         TabIndex        =   101
         Top             =   1920
         Width           =   7035
         Begin VB.TextBox HEM001_15 
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
            Left            =   4620
            MaxLength       =   5
            TabIndex        =   13
            Text            =   "0"
            Top             =   630
            Width           =   390
         End
         Begin VB.TextBox HEM001_12 
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
            Left            =   1230
            MaxLength       =   5
            TabIndex        =   10
            Text            =   "2"
            Top             =   330
            Width           =   390
         End
         Begin VB.TextBox HEM001_13 
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
            Left            =   4620
            MaxLength       =   5
            TabIndex        =   11
            Text            =   "2"
            Top             =   300
            Width           =   390
         End
         Begin VB.TextBox HEM001_14 
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
            Left            =   1230
            MaxLength       =   5
            TabIndex        =   12
            Text            =   "28"
            Top             =   660
            Width           =   390
         End
         Begin VB.Label HEM001_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Segmentados"
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
            Index           =   25
            Left            =   3450
            TabIndex        =   109
            Top             =   660
            Width           =   1125
         End
         Begin VB.Label HEM001_00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Mielocitos"
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
            Index           =   19
            Left            =   180
            TabIndex        =   108
            Top             =   360
            Width           =   780
         End
         Begin VB.Label HEM001_00 
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
            Index           =   26
            Left            =   5055
            TabIndex        =   107
            Top             =   660
            Width           =   180
         End
         Begin VB.Label HEM001_00 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Left            =   1665
            TabIndex        =   106
            Top             =   360
            Width           =   165
         End
         Begin VB.Label HEM001_00 
            AutoSize        =   -1  'True
            Caption         =   "Metamielocitos"
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
            Left            =   3360
            TabIndex        =   105
            Top             =   330
            Width           =   1200
         End
         Begin VB.Label HEM001_00 
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
            Index           =   22
            Left            =   5055
            TabIndex        =   104
            Top             =   330
            Width           =   180
         End
         Begin VB.Label HEM001_00 
            AutoSize        =   -1  'True
            Caption         =   "Abastonados"
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
            Index           =   23
            Left            =   180
            TabIndex        =   103
            Top             =   690
            Width           =   1050
         End
         Begin VB.Label HEM001_00 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Left            =   1665
            TabIndex        =   102
            Top             =   690
            Width           =   165
         End
      End
      Begin VB.Label HEM001_00 
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
         Index           =   9
         Left            =   210
         TabIndex        =   100
         Top             =   1260
         Width           =   870
      End
      Begin VB.Label HEM001_00 
         AutoSize        =   -1  'True
         Caption         =   "x mm3"
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
         Left            =   1980
         TabIndex        =   99
         Top             =   1290
         Width           =   555
      End
      Begin VB.Label HEM001_00 
         AutoSize        =   -1  'True
         Caption         =   "Fórmula Leucocitaria"
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
         Index           =   8
         Left            =   120
         TabIndex        =   98
         Top             =   990
         Width           =   1785
      End
      Begin VB.Label HEM001_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hemoglobina"
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
         Left            =   210
         TabIndex        =   97
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label HEM001_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hematocrito"
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
         Left            =   3675
         TabIndex        =   96
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label HEM001_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Plaquetas"
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
         Left            =   3900
         TabIndex        =   95
         Top             =   630
         Width           =   780
      End
      Begin VB.Label HEM001_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hematíes"
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
         Left            =   210
         TabIndex        =   94
         Top             =   600
         Width           =   750
      End
      Begin VB.Label HEM001_00 
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
         Index           =   27
         Left            =   60
         TabIndex        =   93
         Top             =   3045
         Width           =   1275
      End
      Begin VB.Label HEM001_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Monocitos"
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
         Left            =   5625
         TabIndex        =   92
         Top             =   1590
         Width           =   825
      End
      Begin VB.Label HEM001_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Eosinofilos"
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
         Left            =   210
         TabIndex        =   91
         Top             =   1590
         Width           =   825
      End
      Begin VB.Label HEM001_00 
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
         Left            =   6915
         TabIndex        =   90
         Top             =   1590
         Width           =   180
      End
      Begin VB.Label HEM001_00 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   1785
         TabIndex        =   89
         Top             =   1620
         Width           =   165
      End
      Begin VB.Label HEM001_00 
         AutoSize        =   -1  'True
         Caption         =   "Basófilos"
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
         Left            =   2310
         TabIndex        =   88
         Top             =   1590
         Width           =   675
      End
      Begin VB.Label HEM001_00 
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
         Index           =   14
         Left            =   3405
         TabIndex        =   87
         Top             =   1590
         Width           =   180
      End
      Begin VB.Label HEM001_00 
         AutoSize        =   -1  'True
         Caption         =   "Linfocitos"
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
         Left            =   3840
         TabIndex        =   86
         Top             =   1590
         Width           =   765
      End
      Begin VB.Label HEM001_00 
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
         Left            =   5145
         TabIndex        =   85
         Top             =   1590
         Width           =   180
      End
      Begin VB.Label HEM001_00 
         AutoSize        =   -1  'True
         Caption         =   "mm3"
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
         Left            =   5850
         TabIndex        =   84
         Top             =   630
         Width           =   405
      End
      Begin VB.Label HEM001_00 
         AutoSize        =   -1  'True
         Caption         =   "/ mm3"
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
         Left            =   2460
         TabIndex        =   83
         Top             =   630
         Width           =   540
      End
      Begin VB.Label HEM001_00 
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
         Index           =   3
         Left            =   5850
         TabIndex        =   48
         Top             =   300
         Width           =   180
      End
      Begin VB.Label HEM001_00 
         AutoSize        =   -1  'True
         Caption         =   "g %"
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
         Left            =   2460
         TabIndex        =   47
         Top             =   270
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmHematologia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Resultados de Hematología
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
Dim ml_idPaciente As Long
Dim ml_idAnalisis As Long
Dim ml_resultado As String
Dim ml_observacion As String
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
  Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =13")
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
  Fra.Top = UcPacienteDatos1.Top + UcPacienteDatos1.Height + 700   '350
  fraBoton.Top = Fra.Top + Fra.Height
  Me.Height = fraBoton.Top + fraBoton.Height + 500
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
  If ml_CodigoPruebaSeleccionada = "HEM001" Then 'Hemograma completo
    ml_resultado = HEM001_01.Text & "\" & HEM001_03.Text & "\" & HEM001_04.Text & "\" & HEM001_05.Text & "\" & HEM001_06.Text & "\" & HEM001_07.Text & "\" & HEM001_08.Text & "\" & HEM001_09.Text & "\" & HEM001_10.Text & "\" & HEM001_11.Text & "\" & HEM001_12.Text & "\" & HEM001_13.Text & "\" & HEM001_14.Text & "\" & HEM001_15.Text
    ml_observacion = HEM001_16.Text
  ElseIf ml_CodigoPruebaSeleccionada = "HEM002" Then 'Hemoglobina
    ml_resultado = HEM002_02.Text & "\" & HEM002_03.Text & "\" & HEM002_04.Text & "\" & HEM002_05.Text
    ml_observacion = HEM002_06.Text
  ElseIf ml_CodigoPruebaSeleccionada = "HEM003" Then 'VSG
    ml_resultado = HEM003_01.Text & "\" & HEM003_02.Text & "\" & HEM003_03.Text
    ml_observacion = HEM003_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "2013" Or ml_CodigoPruebaSeleccionada = "2056" Then 'Recuento de reticulocitos
    ml_resultado = HEM004_01.Text & "\" & HEM004_02.Text & "\" & HEM004_03.Text
    ml_observacion = HEM004_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "HEM005" Then 'Recuento de plaquetas
    ml_resultado = HEM005_01.Text & "\" & HEM005_02.Text & "\" & HEM005_03.Text
    ml_observacion = HEM005_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "HEM006" Then 'Tiempo de Coagulación y sangría
    ml_resultado = HEM006_02.Text & "\" & HEM006_03.Text & "\" & HEM006_04.Text & "\" & HEM006_05.Text & "\" & HEM006_06.Text & "\" & HEM006_07.Text & "\" & HEM006_08.Text & "\" & HEM006_09.Text
    ml_observacion = HEM006_10.Text
  ElseIf ml_CodigoPruebaSeleccionada = "HEM008" Then 'Leishmania
    ml_resultado = HEM008_01.Text & "\" & HEM008_02.Text & "\" & HEM008_03.Text
    ml_observacion = HEM008_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "HEM009" Then 'Células L.E.
    ml_resultado = HEM009_01.Text & "\" & HEM009_02.Text & "\" & HEM009_03.Text
    ml_observacion = HEM009_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "HEM010" Then 'Lámina periférica
    ml_resultado = HEM010_01.Text & "\" & HEM010_02.Text & "\" & HEM010_03.Text & "\" & HEM010_04.Text & "\" & HEM010_05.Text & "\" & HEM010_06.Text
    ml_observacion = HEM010_07.Text
  ElseIf ml_CodigoPruebaSeleccionada = "HEM011" Then 'Hematocrito
    ml_resultado = HEM011_02.Text & "\" & HEM011_03.Text & "\" & HEM011_04.Text & "\" & HEM011_05.Text
    ml_observacion = HEM011_06.Text
  ElseIf ml_CodigoPruebaSeleccionada = "HEM012" Then  'Tiempo de protombina
    ml_resultado = HEM012_01.Text & "\" & HEM012_02.Text & "\" & HEM012_03.Text & "\" & HEM012_04.Text & "\" & HEM012_05.Text
    ml_observacion = HEM012_06.Text
  ElseIf ml_CodigoPruebaSeleccionada = "HEM013" Then 'Grupo y factor sanguíneo
    ml_resultado = HEM013_01.Text & "\" & HEM013_02.Text
    ml_observacion = HEM013_03.Text
  ElseIf ml_CodigoPruebaSeleccionada = "HEM030" Then 'Tiempo de coagulación
    ml_resultado = HEM030_02.Text & "\" & HEM030_03.Text & "\" & HEM030_04.Text & "\" & HEM030_05.Text
    ml_observacion = HEM030_06.Text
  ElseIf ml_CodigoPruebaSeleccionada = "HEM031" Then 'Tiempo de sangria
    ml_resultado = HEM031_02.Text & "\" & HEM031_03.Text & "\" & HEM031_04.Text & "\" & HEM031_05.Text
    ml_observacion = HEM031_06.Text
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
  
  'debb-2/3/2015
  mo_ReglasLaboratorio.LabIngresaResultados idPrueba, idOrden, ml_resultado, ml_observacion, idUsuario, _
                                            ml_nombreRealiza, ml_DetalleOrden, ml_idOrdenLab, Me.HEM013_01.Text, _
                                            Me.HEM013_02.Text, ml_idPaciente, CDate(Me.txtFresultado.Text), _
                                            mo_lcNombrePc, mo_lnIdTablaLISTBARITEMS, _
                                            Me.UcPacienteDatos1.DevuelveHistoriaApellidosYnombre, Me.Caption, _
                                            Val(HEM002_02.Text)
End Sub

Private Sub cmdImprimir_Click()
  ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(idPrueba, idOrden)
  ml_observacion = mo_ReglasLaboratorio.LabRecuperaResultados_Obs(idPrueba, idOrden)
  Dim ldFechaResultado As Date
  ml_nombreRealiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(idPrueba, idOrden, ldFechaResultado)
  If ml_CodigoPruebaSeleccionada <> "" And Trim(ml_resultado) <> "" Then
    mo_ReglasLaboratorio.LabImprimeCabeceraResultados UcPacienteDatos1.idPaciente, nombrePaciente, UcPacienteDatos1.NroHistoriaClinica, ldFechaResultado, _
                         nombreMedico + mo_ReglasLaboratorio.DevuelveDatosParaImpresionResultadoLaboratorio(ml_idOrden)
    mo_ReglasLaboratorio.LabImprimeResultadoshem ml_resultado, CStr(ml_CodigoPruebaSeleccionada), Me.Caption, ml_observacion, ml_nombreRealiza
    mo_ReglasLaboratorio.LabImprimePieResultados
  Else
    MsgBox "Debe grabar los resultados antes de poder imprimirlos", vbInformation, ""
  End If
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
  
  If ml_CodigoPruebaSeleccionada = "HEM001" Then 'Hemograma completo
    TopBoton HEM001
  ElseIf ml_CodigoPruebaSeleccionada = "HEM002" Then 'Hemoglobina
    TopBoton HEM002
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       HEM002_03.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "HEM003" Then 'VSG
    TopBoton HEM003
  ElseIf ml_CodigoPruebaSeleccionada = "HEM004" Then  'Recuento de reticulocitos
    TopBoton HEM004
  ElseIf ml_CodigoPruebaSeleccionada = "HEM005" Then 'Recuento de plaquetas
    TopBoton HEM005
  ElseIf ml_CodigoPruebaSeleccionada = "HEM006" Then 'Tiempo de Coagulación y sangría
    TopBoton HEM006
  ElseIf ml_CodigoPruebaSeleccionada = "HEM008" Then 'Leishmania
    TopBoton HEM008
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       HEM008_01.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "HEM009" Then 'Células L.E.
    TopBoton HEM009
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       HEM009_01.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "HEM010" Then 'lámina periférica
    TopBoton HEM010
  ElseIf ml_CodigoPruebaSeleccionada = "HEM011" Then 'Hematocrito
    TopBoton HEM011
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       HEM011_03.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "HEM012" Then 'tiempo de protombina
    TopBoton HEM012
  ElseIf ml_CodigoPruebaSeleccionada = "HEM013" Then 'grupo y factor sanguíneo
    TopBoton HEM013
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
        HEM013_01.ListIndex = 0
        HEM013_02.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "HEM030" Then 'tiempo de coagulación
    TopBoton HEM030
  ElseIf ml_CodigoPruebaSeleccionada = "HEM031" Then 'tiempo de sangria
    TopBoton HEM031
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
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
  If ml_CodigoPruebaSeleccionada = "HEM001" Then 'Hemograma completo
    HEM001_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM001_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM001_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM001_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM001_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM001_07.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM001_08.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM001_09.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM001_10.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM001_11.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM001_12.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM001_13.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM001_14.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM001_15.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM001_16.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "HEM002" Then 'Hemoglobina
    HEM002_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM002_03.ListIndex = Ubica_En_Combo(HEM002_03, Temp)
    HEM002_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM002_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM002_06.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "HEM003" Then 'VSG
    HEM003_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM003_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM003_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM003_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "2013" Or ml_CodigoPruebaSeleccionada = "2056" Then  'Recuento de reticulocitos
    HEM004_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM004_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM004_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM004_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "HEM005" Then 'Recuento de plaquetas
    HEM005_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM005_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM005_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM005_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "HEM006" Then 'Tiempo de Coagulación y sangría
    HEM006_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM006_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM006_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM006_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM006_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM006_07.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM006_08.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM006_09.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM006_10.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "HEM008" Then 'Leishmania
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM008_01.ListIndex = Ubica_En_Combo(HEM008_01, Temp)
    HEM008_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM008_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM008_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "HEM009" Then 'Células L.E.
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM009_01.ListIndex = Ubica_En_Combo(HEM009_01, Temp)
    HEM009_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM009_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM009_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "HEM010" Then 'lámina periférica
    HEM010_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM010_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM010_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM010_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM010_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM010_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM010_07.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "HEM011" Then 'Hematocrito
    HEM011_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM011_03.ListIndex = Ubica_En_Combo(HEM011_03, Temp)
    HEM011_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM011_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM011_06.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "HEM012" Then  'tiempo de protombina
    HEM012_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM012_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM012_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM012_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM012_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM012_06.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "HEM013" Then 'grupo y factor sanguíneo
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM013_01.ListIndex = Ubica_En_Combo(HEM013_01, Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM013_02.ListIndex = Ubica_En_Combo(HEM013_02, Temp)
    HEM013_03.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "HEM030" Then 'tiempo de coagulación
    HEM030_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM030_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM030_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM030_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM030_06.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "HEM031" Then 'tiempo de sangria
    HEM031_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM031_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM031_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM031_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    HEM031_06.Text = ml_observacion
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
End Sub

Private Sub HEM001_01_GotFocus()
  SeleccionaTexto HEM001_01
End Sub

Private Sub HEM001_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM001_03_GotFocus()
  SeleccionaTexto HEM001_03
End Sub

Private Sub HEM001_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM001_04_GotFocus()
  SeleccionaTexto HEM001_04
End Sub

Private Sub HEM001_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM001_05_GotFocus()
  SeleccionaTexto HEM001_05
End Sub

Private Sub HEM001_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM001_06_GotFocus()
  SeleccionaTexto HEM001_06
End Sub

Private Sub HEM001_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM001_07_GotFocus()
  SeleccionaTexto HEM001_07
End Sub

Private Sub HEM001_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM001_08_GotFocus()
  SeleccionaTexto HEM001_08
End Sub

Private Sub HEM001_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM001_09_GotFocus()
  SeleccionaTexto HEM001_09
End Sub

Private Sub HEM001_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM001_10_GotFocus()
  SeleccionaTexto HEM001_10
End Sub

Private Sub HEM001_10_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM001_11_GotFocus()
  SeleccionaTexto HEM001_11
End Sub

Private Sub HEM001_11_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM001_12_GotFocus()
  SeleccionaTexto HEM001_12
End Sub

Private Sub HEM001_12_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM001_13_GotFocus()
  SeleccionaTexto HEM001_13
End Sub

Private Sub HEM001_13_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM001_14_GotFocus()
  SeleccionaTexto HEM001_14
End Sub

Private Sub HEM001_14_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM001_15_GotFocus()
  SeleccionaTexto HEM001_15
End Sub

Private Sub HEM001_15_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM001_16_GotFocus()
  SeleccionaTexto HEM001_16
End Sub

Private Sub HEM001_16_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM002_02_GotFocus()
  SeleccionaTexto HEM002_02
End Sub

Private Sub HEM002_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM002_03_Click()
  If HEM002_03.ListIndex = 0 Then
    HEM002_04.Text = "13 - 15 gr %"
  Else
    HEM002_04.Text = "12 - 14 gr %"
  End If
End Sub

Private Sub HEM002_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM002_04_GotFocus()
  SeleccionaTexto HEM002_04
End Sub

Private Sub HEM002_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM002_05_GotFocus()
  SeleccionaTexto HEM002_05
End Sub

Private Sub HEM002_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM002_06_GotFocus()
  SeleccionaTexto HEM002_06
End Sub

Private Sub HEM002_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM003_01_GotFocus()
  SeleccionaTexto HEM003_01
End Sub

Private Sub HEM003_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM003_02_GotFocus()
  SeleccionaTexto HEM003_02
End Sub

Private Sub HEM003_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM003_03_GotFocus()
  SeleccionaTexto HEM003_03
End Sub

Private Sub HEM003_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM003_04_GotFocus()
  SeleccionaTexto HEM003_04
End Sub

Private Sub HEM003_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM004_01_GotFocus()
  SeleccionaTexto HEM004_01
End Sub

Private Sub HEM004_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM004_02_GotFocus()
  SeleccionaTexto HEM004_02
End Sub

Private Sub HEM004_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM004_03_GotFocus()
  SeleccionaTexto HEM004_03
End Sub

Private Sub HEM004_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM004_04_GotFocus()
  SeleccionaTexto HEM004_04
End Sub

Private Sub HEM004_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM005_01_GotFocus()
  SeleccionaTexto HEM005_01
End Sub

Private Sub HEM005_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM005_02_GotFocus()
  SeleccionaTexto HEM005_02
End Sub

Private Sub HEM005_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM005_03_GotFocus()
  SeleccionaTexto HEM005_03
End Sub

Private Sub HEM005_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM005_04_GotFocus()
  SeleccionaTexto HEM005_04
End Sub

Private Sub HEM005_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM006_02_GotFocus()
  SeleccionaTexto HEM006_02
End Sub

Private Sub HEM006_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM006_03_GotFocus()
  SeleccionaTexto HEM006_03
End Sub

Private Sub HEM006_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM006_04_GotFocus()
  SeleccionaTexto HEM006_04
End Sub

Private Sub HEM006_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM006_05_GotFocus()
  SeleccionaTexto HEM006_05
End Sub

Private Sub HEM006_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM006_06_GotFocus()
  SeleccionaTexto HEM006_06
End Sub

Private Sub HEM006_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM006_07_GotFocus()
  SeleccionaTexto HEM006_07
End Sub

Private Sub HEM006_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM006_08_GotFocus()
  SeleccionaTexto HEM006_08
End Sub

Private Sub HEM006_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM006_09_GotFocus()
  SeleccionaTexto HEM006_09
End Sub

Private Sub HEM006_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM006_10_GotFocus()
  SeleccionaTexto HEM006_10
End Sub

Private Sub HEM006_10_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM008_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM008_02_GotFocus()
  SeleccionaTexto HEM008_02
End Sub

Private Sub HEM008_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM008_03_GotFocus()
  SeleccionaTexto HEM008_03
End Sub

Private Sub HEM008_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM008_04_GotFocus()
  SeleccionaTexto HEM008_04
End Sub

Private Sub HEM008_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM009_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM009_02_GotFocus()
  SeleccionaTexto HEM009_02
End Sub

Private Sub HEM009_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM009_03_GotFocus()
  SeleccionaTexto HEM009_03
End Sub

Private Sub HEM009_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM009_04_GotFocus()
  SeleccionaTexto HEM009_04
End Sub

Private Sub HEM009_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM010_01_GotFocus()
  SeleccionaTexto HEM010_01
End Sub

Private Sub HEM010_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM010_02_GotFocus()
  SeleccionaTexto HEM010_02
End Sub

Private Sub HEM010_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM010_03_GotFocus()
  SeleccionaTexto HEM010_03
End Sub

Private Sub HEM010_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM010_04_GotFocus()
  SeleccionaTexto HEM010_04
End Sub

Private Sub HEM010_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM010_05_GotFocus()
  SeleccionaTexto HEM010_05
End Sub

Private Sub HEM010_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM010_06_GotFocus()
  SeleccionaTexto HEM010_06
End Sub

Private Sub HEM010_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM010_07_GotFocus()
  SeleccionaTexto HEM010_07
End Sub

Private Sub HEM010_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM011_02_GotFocus()
  SeleccionaTexto HEM011_02
End Sub

Private Sub HEM011_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM011_03_Click()
  If HEM011_03.ListIndex = 0 Then
    HEM011_04.Text = "43 - 49 %"
  Else
    HEM011_04.Text = "40 - 47 %"
  End If
End Sub

Private Sub HEM011_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM011_04_GotFocus()
  SeleccionaTexto HEM011_04
End Sub

Private Sub HEM011_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM011_05_GotFocus()
  SeleccionaTexto HEM011_05
End Sub

Private Sub HEM011_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM011_06_GotFocus()
  SeleccionaTexto HEM011_06
End Sub

Private Sub HEM011_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM012_01_GotFocus()
  SeleccionaTexto HEM012_01
End Sub

Private Sub HEM012_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM012_02_GotFocus()
  SeleccionaTexto HEM012_02
End Sub

Private Sub HEM012_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM012_03_GotFocus()
  SeleccionaTexto HEM012_03
End Sub

Private Sub HEM012_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM012_04_GotFocus()
  SeleccionaTexto HEM012_04
End Sub

Private Sub HEM012_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM012_05_GotFocus()
  SeleccionaTexto HEM012_05
End Sub

Private Sub HEM012_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM012_06_GotFocus()
  SeleccionaTexto HEM012_06
End Sub

Private Sub HEM012_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM013_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM013_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM013_03_GotFocus()
  SeleccionaTexto HEM013_03
End Sub

Private Sub HEM013_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM030_02_GotFocus()
  SeleccionaTexto HEM030_02
End Sub

Private Sub HEM030_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM030_03_GotFocus()
  SeleccionaTexto HEM030_03
End Sub

Private Sub HEM030_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM030_04_GotFocus()
  SeleccionaTexto HEM030_04
End Sub

Private Sub HEM030_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM030_05_GotFocus()
  SeleccionaTexto HEM030_05
End Sub

Private Sub HEM030_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM030_06_GotFocus()
  SeleccionaTexto HEM030_06
End Sub

Private Sub HEM030_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM031_02_GotFocus()
  SeleccionaTexto HEM031_02
End Sub

Private Sub HEM031_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM031_03_GotFocus()
  SeleccionaTexto HEM031_03
End Sub

Private Sub HEM031_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM031_04_GotFocus()
  SeleccionaTexto HEM031_04
End Sub

Private Sub HEM031_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM031_05_GotFocus()
  SeleccionaTexto HEM031_05
End Sub

Private Sub HEM031_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub HEM031_06_GotFocus()
  SeleccionaTexto HEM031_06
End Sub

Private Sub HEM031_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub


Sub LimpiaVAloresDefault()
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       Exit Sub
    End If
    HEM001_01.Text = ""
    HEM001_04.Text = ""
    HEM001_03.Text = ""
    HEM001_05.Text = ""
    HEM001_06.Text = ""
    HEM001_07.Text = ""
    HEM001_11.Text = ""
    HEM001_08.Text = ""
    HEM001_09.Text = ""
    HEM001_10.Text = ""
    HEM001_12.Text = ""
    HEM001_14.Text = ""
    HEM001_13.Text = ""
    HEM001_15.Text = ""
    HEM002_04.Text = ""
    HEM002_05.Text = ""
    HEM003_02.Text = ""
    HEM003_03.Text = ""
    HEM004_02.Text = ""
    HEM004_03.Text = ""
    HEM005_02.Text = ""
    HEM005_03.Text = ""
    HEM006_04.Text = ""
    HEM006_08.Text = ""
    HEM006_05.Text = ""
    HEM006_09.Text = ""
    HEM008_03.Text = ""
    HEM009_03.Text = ""
    HEM011_04.Text = ""
    HEM011_05.Text = ""
    HEM012_02.Text = ""
    HEM012_03.Text = ""
    HEM030_04.Text = ""
    HEM030_05.Text = ""
    HEM031_04.Text = ""
    HEM031_05.Text = ""
    HEM013_01.Text = ""
    HEM013_02.Text = ""
End Sub

