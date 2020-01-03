VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmUrianalisis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "URIANÁLISIS"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   ForeColor       =   &H00000000&
   Icon            =   "frmUrianalisis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   60
      TabIndex        =   79
      Top             =   1740
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
         TabIndex        =   80
         Top             =   180
         Width           =   3090
      End
      Begin MSMask.MaskEdBox txtFresultado 
         Height          =   315
         Left            =   5580
         TabIndex        =   81
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
         TabIndex        =   83
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
         Left            =   4620
         TabIndex        =   82
         Top             =   240
         Width           =   945
      End
   End
   Begin SIGHLaboratorio.UcPacienteDatos1 UcPacienteDatos1 
      Height          =   1815
      Left            =   60
      TabIndex        =   73
      Top             =   15
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3201
   End
   Begin VB.Frame fraBoton 
      ForeColor       =   &H00000000&
      Height          =   870
      Left            =   60
      TabIndex        =   74
      Top             =   6600
      Width           =   7200
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmUrianalisis.frx":0CCA
         DownPicture     =   "frmUrianalisis.frx":118E
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
         Picture         =   "frmUrianalisis.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprime (F3)"
         Height          =   615
         Left            =   135
         Picture         =   "frmUrianalisis.frx":1B66
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmUrianalisis.frx":203F
         DownPicture     =   "frmUrianalisis.frx":249F
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
         Left            =   2273
         Picture         =   "frmUrianalisis.frx":2914
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   180
         Width           =   1365
      End
   End
   Begin VB.Frame ANA003 
      Caption         =   "Proteinas en orina de 24 horas"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   58
      Top             =   2340
      Visible         =   0   'False
      Width           =   7155
      Begin VB.TextBox ANA003_01 
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
         Left            =   165
         MaxLength       =   5
         TabIndex        =   20
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox ANA003_04 
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
         TabIndex        =   23
         Top             =   740
         Width           =   5670
      End
      Begin VB.TextBox ANA003_02 
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
         Left            =   2640
         TabIndex        =   21
         Text            =   "Positivo > 0.3 gr / l / 24 horas"
         Top             =   420
         Width           =   2295
      End
      Begin VB.TextBox ANA003_03 
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
         TabIndex        =   22
         Text            =   "Ferrocianuro de Potasio"
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label ANA003_00 
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
         Left            =   300
         TabIndex        =   63
         Top             =   225
         Width           =   1935
      End
      Begin VB.Label ANA003_00 
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
         TabIndex        =   62
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label ANA003_00 
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
         Left            =   2700
         TabIndex        =   61
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label ANA003_00 
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
         TabIndex        =   60
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label ANA003_00 
         AutoSize        =   -1  'True
         Caption         =   "gr / l / 24 horas"
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
         Left            =   1020
         TabIndex        =   59
         Top             =   450
         Width           =   1290
      End
   End
   Begin VB.Frame ANA010 
      Caption         =   "Prueba de Addis"
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
      Height          =   1660
      Left            =   60
      TabIndex        =   64
      Top             =   2400
      Visible         =   0   'False
      Width           =   7200
      Begin VB.TextBox ANA010_05 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1730
         TabIndex        =   28
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox ANA010_07 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1420
         TabIndex        =   30
         Top             =   1290
         Width           =   5670
      End
      Begin VB.TextBox ANA010_03 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1730
         TabIndex        =   26
         Top             =   630
         Width           =   855
      End
      Begin VB.TextBox ANA010_02 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5625
         TabIndex        =   25
         Top             =   300
         Width           =   1485
      End
      Begin VB.TextBox ANA010_04 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5625
         TabIndex        =   27
         Top             =   630
         Width           =   1485
      End
      Begin VB.TextBox ANA010_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1730
         TabIndex        =   24
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox ANA010_06 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5625
         TabIndex        =   29
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label ANA010_00 
         AutoSize        =   -1  'True
         Caption         =   "mg / l"
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
         Left            =   2640
         TabIndex        =   72
         Top             =   330
         Width           =   480
      End
      Begin VB.Label ANA010_00 
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
         Index           =   7
         Left            =   60
         TabIndex        =   71
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label ANA010_00 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   70
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label ANA010_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Cilindros Hialinos"
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
         Left            =   120
         TabIndex        =   69
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label ANA010_00 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   2
         Left            =   3735
         TabIndex        =   68
         Top             =   330
         Width           =   1815
      End
      Begin VB.Label ANA010_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Cel. Epiteliales"
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
         Index           =   4
         Left            =   3735
         TabIndex        =   67
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label ANA010_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Volumen"
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
         Index           =   0
         Left            =   150
         TabIndex        =   66
         Top             =   330
         Width           =   1575
      End
      Begin VB.Label ANA010_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Cilindros Granulosos"
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
         Index           =   6
         Left            =   3735
         TabIndex        =   65
         Top             =   990
         Width           =   1815
      End
   End
   Begin VB.Frame ANA001 
      Caption         =   "Examen Completo de Orina"
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
      Height          =   4260
      Left            =   60
      TabIndex        =   31
      Top             =   2340
      Visible         =   0   'False
      Width           =   7200
      Begin VB.ComboBox ANA001_20 
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
         ItemData        =   "frmUrianalisis.frx":2D89
         Left            =   3960
         List            =   "frmUrianalisis.frx":2D93
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox ANA001_07 
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
         Left            =   3705
         TabIndex        =   13
         Text            =   "0 - 0"
         Top             =   2730
         Width           =   615
      End
      Begin VB.TextBox ANA001_11 
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
         TabIndex        =   17
         Text            =   "0 - 0"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox ANA001_09 
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
         TabIndex        =   15
         Text            =   "0 - 1"
         Top             =   3045
         Width           =   615
      End
      Begin VB.TextBox ANA001_06 
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
         TabIndex        =   12
         Text            =   "1 - 3"
         Top             =   2730
         Width           =   615
      End
      Begin VB.TextBox ANA001_01 
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
         Left            =   720
         TabIndex        =   0
         Text            =   "Amarillo Pajizo"
         Top             =   460
         Width           =   1335
      End
      Begin VB.TextBox ANA001_04 
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
         Left            =   705
         TabIndex        =   3
         Text            =   "Suigéneris"
         Top             =   780
         Width           =   1335
      End
      Begin VB.TextBox ANA001_02 
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
         TabIndex        =   1
         Text            =   "Transparente"
         Top             =   460
         Width           =   1455
      End
      Begin VB.TextBox ANA001_03 
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
         Left            =   6225
         TabIndex        =   2
         Text            =   "1,010"
         Top             =   460
         Width           =   855
      End
      Begin VB.TextBox ANA001_19 
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
         TabIndex        =   19
         Top             =   3795
         Width           =   5670
      End
      Begin VB.ComboBox ANA001_05 
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
         ItemData        =   "frmUrianalisis.frx":2DAB
         Left            =   3465
         List            =   "frmUrianalisis.frx":2DB8
         Style           =   1  'Simple Combo
         TabIndex        =   4
         Text            =   "ANA001_05"
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox ANA001_08 
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
         Left            =   5490
         TabIndex        =   14
         Top             =   2730
         Width           =   1575
      End
      Begin VB.TextBox ANA001_12 
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
         Left            =   3690
         TabIndex        =   18
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox ANA001_10 
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
         Left            =   3690
         TabIndex        =   16
         Top             =   3045
         Width           =   3375
      End
      Begin VB.ComboBox ANA001_13 
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
         ItemData        =   "frmUrianalisis.frx":2DD5
         Left            =   1065
         List            =   "frmUrianalisis.frx":2DDF
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1395
         Width           =   975
      End
      Begin VB.ComboBox ANA001_15 
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
         ItemData        =   "frmUrianalisis.frx":2DF7
         Left            =   6180
         List            =   "frmUrianalisis.frx":2E01
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1395
         Width           =   975
      End
      Begin VB.ComboBox ANA001_17 
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
         ItemData        =   "frmUrianalisis.frx":2E19
         Left            =   3960
         List            =   "frmUrianalisis.frx":2E23
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1725
         Width           =   975
      End
      Begin VB.ComboBox ANA001_14 
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
         ItemData        =   "frmUrianalisis.frx":2E3B
         Left            =   3960
         List            =   "frmUrianalisis.frx":2E45
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1395
         Width           =   975
      End
      Begin VB.ComboBox ANA001_16 
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
         ItemData        =   "frmUrianalisis.frx":2E5D
         Left            =   1080
         List            =   "frmUrianalisis.frx":2E67
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1725
         Width           =   975
      End
      Begin VB.ComboBox ANA001_18 
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
         ItemData        =   "frmUrianalisis.frx":2E7F
         Left            =   1080
         List            =   "frmUrianalisis.frx":2E89
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label ANA001_00 
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
         Index           =   14
         Left            =   5100
         TabIndex        =   78
         Top             =   1485
         Width           =   1050
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   8
         Left            =   2760
         TabIndex        =   57
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pus"
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
         Left            =   1110
         TabIndex        =   56
         Top             =   3390
         Width           =   285
      End
      Begin VB.Label ANA001_00 
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
         Index           =   11
         Left            =   525
         TabIndex        =   55
         Top             =   3075
         Width           =   870
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cél. Epiteliales"
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
         Left            =   255
         TabIndex        =   54
         Top             =   2760
         Width           =   1140
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   53
         Top             =   495
         Width           =   495
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   52
         Top             =   810
         Width           =   495
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   2
         Left            =   2535
         TabIndex        =   51
         Top             =   495
         Width           =   855
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Reacción"
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
         Left            =   2535
         TabIndex        =   50
         Top             =   810
         Width           =   855
      End
      Begin VB.Label ANA001_00 
         Caption         =   "Densidad"
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
         Left            =   5295
         TabIndex        =   49
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label ANA001_00 
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
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   48
         Top             =   220
         Width           =   1575
      End
      Begin VB.Label ANA001_00 
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
         Height          =   255
         Index           =   27
         Left            =   60
         TabIndex        =   47
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label ANA001_00 
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
         Index           =   26
         Left            =   60
         TabIndex        =   46
         Top             =   3825
         Width           =   1335
      End
      Begin VB.Label ANA001_00 
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
         Height          =   255
         Index           =   10
         Left            =   4920
         TabIndex        =   45
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Cristales"
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
         Index           =   17
         Left            =   2760
         TabIndex        =   44
         Top             =   3390
         Width           =   855
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Cilindros"
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
         Index           =   13
         Left            =   2760
         TabIndex        =   43
         Top             =   3075
         Width           =   855
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Urobilinogeno"
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
         Left            =   2820
         TabIndex        =   42
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bilirrubina"
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
         Index           =   24
         Left            =   255
         TabIndex        =   41
         Top             =   2070
         Width           =   750
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Proteinas"
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
         Left            =   255
         TabIndex        =   40
         Top             =   1755
         Width           =   750
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Glucosa"
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
         Index           =   20
         Left            =   390
         TabIndex        =   39
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label ANA001_00 
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
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cuerpos Cetónicos"
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
         Left            =   2415
         TabIndex        =   37
         Top             =   1755
         Width           =   1515
      End
      Begin VB.Label ANA001_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nitritos"
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
         Left            =   3360
         TabIndex        =   36
         Top             =   2070
         Width           =   570
      End
      Begin VB.Label ANA001_00 
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
         Index           =   9
         Left            =   4365
         TabIndex        =   35
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label ANA001_00 
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
         Index           =   16
         Left            =   2115
         TabIndex        =   34
         Top             =   3390
         Width           =   240
      End
      Begin VB.Label ANA001_00 
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
         Index           =   12
         Left            =   2115
         TabIndex        =   33
         Top             =   3075
         Width           =   240
      End
      Begin VB.Label ANA001_00 
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
         Index           =   7
         Left            =   2115
         TabIndex        =   32
         Top             =   2760
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmUrianalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Resultados para Urianálisis
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
  Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =15")
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
'  If EmpleadoTrabajaEnLaboratorio(sighEntidades.Usuario) = True Then
    Fra.Enabled = True
'  Else
'    Fra.Enabled = False
'  End If
  Fra.Visible = True
  Fra.Caption = ml_nombrePrueba
  Me.Caption = Fra.Caption
  Fra.Top = UcPacienteDatos1.Top + UcPacienteDatos1.Height + 500  '350
  fraBoton.Top = Fra.Top + Fra.Height
  Me.Height = fraBoton.Top + fraBoton.Height + 500
End Sub

Private Sub ANA001_01_GotFocus()
  SeleccionaTexto ANA001_01
End Sub

Private Sub ANA001_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_02_GotFocus()
  SeleccionaTexto ANA001_02
End Sub

Private Sub ANA001_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_03_GotFocus()
  SeleccionaTexto ANA001_03
End Sub

Private Sub ANA001_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_04_GotFocus()
  SeleccionaTexto ANA001_04
End Sub

Private Sub ANA001_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_06_GotFocus()
  SeleccionaTexto ANA001_06
End Sub

Private Sub ANA001_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_07_GotFocus()
  SeleccionaTexto ANA001_07
End Sub

Private Sub ANA001_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_08_GotFocus()
  SeleccionaTexto ANA001_08
End Sub

Private Sub ANA001_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_09_GotFocus()
  SeleccionaTexto ANA001_09
End Sub

Private Sub ANA001_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_10_GotFocus()
  SeleccionaTexto ANA001_10
End Sub

Private Sub ANA001_10_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_11_GotFocus()
  SeleccionaTexto ANA001_11
End Sub

Private Sub ANA001_11_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_12_GotFocus()
  SeleccionaTexto ANA001_12
End Sub

Private Sub ANA001_12_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_13_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_14_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_15_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_16_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_17_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_18_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_19_GotFocus()
  SeleccionaTexto ANA001_19
End Sub

Private Sub ANA001_19_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA001_20_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA003_01_GotFocus()
  SeleccionaTexto ANA003_01
End Sub

Private Sub ANA003_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA003_02_GotFocus()
  SeleccionaTexto ANA003_02
End Sub

Private Sub ANA003_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA003_03_GotFocus()
  SeleccionaTexto ANA003_03
End Sub

Private Sub ANA003_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA003_04_GotFocus()
  SeleccionaTexto ANA003_04
End Sub

Private Sub ANA003_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA010_01_GotFocus()
  SeleccionaTexto ANA010_01
End Sub

Private Sub ANA010_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA010_02_GotFocus()
  SeleccionaTexto ANA010_02
End Sub

Private Sub ANA010_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA010_03_GotFocus()
  SeleccionaTexto ANA010_03
End Sub

Private Sub ANA010_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA010_04_GotFocus()
  SeleccionaTexto ANA010_04
End Sub

Private Sub ANA010_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA010_05_GotFocus()
  SeleccionaTexto ANA010_05
End Sub

Private Sub ANA010_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA010_06_GotFocus()
  SeleccionaTexto ANA010_06
End Sub

Private Sub ANA010_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub ANA010_07_GotFocus()
  SeleccionaTexto ANA010_07
End Sub

Private Sub ANA010_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
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
  If ml_CodigoPruebaSeleccionada = "ANA001" Then 'Examen completo de orina
    ml_resultado = ANA001_01.Text & "\" & ANA001_02.Text & "\" & ANA001_03.Text & "\" & ANA001_04.Text & "\" & ANA001_05.Text & "\" & ANA001_13.Text & "\" & ANA001_14.Text & "\" & ANA001_15.Text & "\" & ANA001_16.Text & "\" & ANA001_17.Text & "\" & ANA001_18.Text & "\" & ANA001_20.Text & "\" & ANA001_06.Text & "\" & ANA001_07.Text & "\" & ANA001_08.Text & "\" & ANA001_09.Text & "\" & ANA001_10.Text & "\" & ANA001_11.Text & "\" & ANA001_12.Text
    ml_observacion = ANA001_19.Text
  ElseIf ml_CodigoPruebaSeleccionada = "82042" Then 'proteina en orina de 24 horas
    ml_resultado = ANA003_01.Text & "\" & ANA003_02.Text & "\" & ANA003_03.Text
    ml_observacion = ANA003_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "ANA010" Then 'prueba de Addis
    ml_resultado = ANA010_01.Text & "\" & ANA010_02.Text & "\" & ANA010_03.Text & "\" & ANA010_04.Text & "\" & ANA010_05.Text & "\" & ANA010_06.Text
    ml_observacion = ANA010_07.Text
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
  mo_ReglasLaboratorio.LabIngresaResultados idPrueba, idOrden, ml_resultado, ml_observacion, idUsuario, ml_nombreRealiza, ml_DetalleOrden, ml_idOrdenLab, "", "", 0, CDate(Me.txtFresultado.Text), mo_lcNombrePc, mo_lnIdTablaLISTBARITEMS, Me.UcPacienteDatos1.DevuelveHistoriaApellidosYnombre, Me.Caption
End Sub

Private Sub cmdImprimir_Click()
  ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(idPrueba, idOrden)
  ml_observacion = mo_ReglasLaboratorio.LabRecuperaResultados_Obs(idPrueba, idOrden)
  Dim ldFechaResultado As Date
  ml_nombreRealiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(idPrueba, idOrden, ldFechaResultado)
  If ml_CodigoPruebaSeleccionada <> "" And Trim(ml_resultado) <> "" Then
    mo_ReglasLaboratorio.LabImprimeCabeceraResultados UcPacienteDatos1.idPaciente, nombrePaciente, UcPacienteDatos1.NroHistoriaClinica, ldFechaResultado, _
                         nombreMedico + mo_ReglasLaboratorio.DevuelveDatosParaImpresionResultadoLaboratorio(ml_idOrden)
    mo_ReglasLaboratorio.LabImprimeResultadosANA ml_resultado, CStr(ml_CodigoPruebaSeleccionada), Me.Caption, ml_observacion, ml_nombreRealiza
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
     If ml_FechaNacimiento <> 0 Then
        Me.UcPacienteDatos1.FechaNacimiento = ml_FechaNacimiento
     End If
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
  
  If ml_CodigoPruebaSeleccionada = "ANA001" Then 'Examen completo de orina
    TopBoton ANA001
  ElseIf ml_CodigoPruebaSeleccionada = "ANA003" Then 'proteina en orina de 24 horas
    TopBoton ANA003
  ElseIf ml_CodigoPruebaSeleccionada = "ANA010" Then 'prueba de Addis
    TopBoton ANA010
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
  'Recupera información si es que ya esta grabado
  ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(idPrueba, idOrden)
  Dim ldFehaResultado As Date
  ml_nombreRealiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(idPrueba, idOrden, ldFehaResultado)
  If ml_resultado = "" Or Val(ml_nombreRealiza) = 0 Then Exit Sub
  Me.txtFresultado.Text = Format(IIf(ldFehaResultado = 0, Now, ldFehaResultado), sighentidades.DevuelveFechaSoloFormato_DMY_HM)
  cmbResponsable.ListIndex = Ubica_En_Combo(cmbResponsable, mo_ReglasLaboratorio.LabEmpleado(ml_nombreRealiza))
  'If cmbResponsable.Text <> "" Then cmbResponsable.Enabled = False
  ml_observacion = mo_ReglasLaboratorio.LabRecuperaResultados_Obs(idPrueba, idOrden)
  Dim Temp As String
  'Asigna la información recuperada en el formulario
  If ml_CodigoPruebaSeleccionada = "ANA001" Then 'Examen completo de orina
    ANA001_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_05.ListIndex = Ubica_En_Combo(ANA001_05, Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_13.ListIndex = Ubica_En_Combo(ANA001_13, Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_14.ListIndex = Ubica_En_Combo(ANA001_14, Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_15.ListIndex = Ubica_En_Combo(ANA001_15, Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_16.ListIndex = Ubica_En_Combo(ANA001_16, Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_17.ListIndex = Ubica_En_Combo(ANA001_17, Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_18.ListIndex = Ubica_En_Combo(ANA001_18, Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_20.ListIndex = Ubica_En_Combo(ANA001_20, Temp)
    ANA001_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_07.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_08.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_09.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_10.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_11.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_12.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA001_19.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "ANA003" Then 'proteina en orina de 24 horas
    ANA003_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA003_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA003_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA003_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "ANA010" Then 'prueba de Addis
    ANA010_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA010_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA010_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA010_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA010_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA010_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    ANA010_07.Text = ml_observacion
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
 
End Sub

Sub LimpiaVAloresDefault()
    
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       Exit Sub
    End If
    ANA003_02.Text = ""
    ANA003_03.Text = ""
    ANA001_06.Text = ""
    ANA001_09.Text = ""
    ANA001_11.Text = ""
    ANA001_07.Text = ""
    ANA001_01.Text = ""
    ANA001_04.Text = ""
    ANA001_02.Text = ""
    ANA001_03.Text = ""
    ANA001_05.Text = ""
End Sub
