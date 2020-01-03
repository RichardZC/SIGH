VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmBioquimica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BIOQUÍMICA"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   ForeColor       =   &H00000000&
   Icon            =   "frmBioquimica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBoton 
      ForeColor       =   &H00000000&
      Height          =   870
      Left            =   60
      TabIndex        =   308
      Top             =   4440
      Width           =   7185
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmBioquimica.frx":0CCA
         DownPicture     =   "frmBioquimica.frx":118E
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
         Left            =   3765
         Picture         =   "frmBioquimica.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   311
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
         Left            =   120
         Picture         =   "frmBioquimica.frx":1B66
         Style           =   1  'Graphical
         TabIndex        =   310
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmBioquimica.frx":203F
         DownPicture     =   "frmBioquimica.frx":249F
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
         Left            =   2325
         Picture         =   "frmBioquimica.frx":2914
         Style           =   1  'Graphical
         TabIndex        =   309
         Top             =   180
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   60
      TabIndex        =   314
      Top             =   1710
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
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   315
         Top             =   180
         Width           =   3120
      End
      Begin MSMask.MaskEdBox txtFresultado 
         Height          =   315
         Left            =   5580
         TabIndex        =   316
         Top             =   180
         Width           =   1515
         _ExtentX        =   2672
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
         TabIndex        =   318
         Top             =   240
         Width           =   1215
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
         TabIndex        =   317
         Top             =   210
         Width           =   945
      End
   End
   Begin SIGHLaboratorio.UcPacienteDatos1 UcPacienteDatos1 
      Height          =   1695
      Left            =   60
      TabIndex        =   302
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   2990
   End
   Begin VB.Frame BQM031 
      Caption         =   "Proteina Total y Fraccionada"
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
      Height          =   2130
      Left            =   60
      TabIndex        =   288
      Top             =   5370
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox BQM031_18 
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
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   296
         TabStop         =   0   'False
         Text            =   "Proteína Total"
         Top             =   750
         Width           =   1815
      End
      Begin VB.TextBox BQM031_05 
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
         Left            =   2220
         MaxLength       =   5
         TabIndex        =   140
         Top             =   750
         Width           =   615
      End
      Begin VB.TextBox BQM031_07 
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
         Left            =   4020
         TabIndex        =   142
         Text            =   "6.1 - 7.9"
         Top             =   750
         Width           =   855
      End
      Begin VB.TextBox BQM031_08 
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
         TabIndex        =   143
         Text            =   "Colorimétrico"
         Top             =   750
         Width           =   1095
      End
      Begin VB.TextBox BQM031_16 
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
         TabIndex        =   151
         Text            =   "Colorimétrico"
         Top             =   1410
         Width           =   1095
      End
      Begin VB.TextBox BQM031_12 
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
         TabIndex        =   147
         Text            =   "Colorimétrico"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox BQM031_11 
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
         Left            =   4020
         TabIndex        =   146
         Text            =   "3.5 - 4.8"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox BQM031_09 
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
         Left            =   2220
         MaxLength       =   5
         TabIndex        =   144
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox BQM031_18 
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
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   295
         TabStop         =   0   'False
         Text            =   "Albúmina"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox BQM031_15 
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
         Left            =   4020
         TabIndex        =   150
         Text            =   "1.6 - 3.0"
         Top             =   1410
         Width           =   855
      End
      Begin VB.TextBox BQM031_13 
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
         Left            =   2220
         MaxLength       =   5
         TabIndex        =   148
         Top             =   1410
         Width           =   615
      End
      Begin VB.TextBox BQM031_18 
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
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   294
         TabStop         =   0   'False
         Text            =   "Globulinas"
         Top             =   1410
         Width           =   1815
      End
      Begin VB.TextBox BQM031_06 
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
         Left            =   2940
         TabIndex        =   141
         Text            =   "g / dl"
         Top             =   750
         Width           =   495
      End
      Begin VB.TextBox BQM031_10 
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
         Left            =   2940
         TabIndex        =   145
         Text            =   "g / dl"
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox BQM031_14 
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
         Left            =   2940
         TabIndex        =   149
         Text            =   "g / dl"
         Top             =   1410
         Width           =   495
      End
      Begin VB.TextBox BQM031_04 
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
         TabIndex        =   139
         Text            =   "Colorimétrico"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox BQM031_03 
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
         Left            =   4020
         TabIndex        =   138
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox BQM031_17 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1470
         TabIndex        =   152
         Top             =   1740
         Width           =   5580
      End
      Begin VB.TextBox BQM031_01 
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
         Left            =   2220
         MaxLength       =   5
         TabIndex        =   136
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox BQM031_02 
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
         Left            =   2940
         TabIndex        =   137
         Text            =   "g / dl"
         Top             =   420
         Width           =   495
      End
      Begin VB.TextBox BQM031_18 
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
         TabIndex        =   301
         TabStop         =   0   'False
         Text            =   "Proteína Total"
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label BQM031_00 
         AutoSize        =   -1  'True
         Caption         =   "g / dl"
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
         Left            =   4920
         TabIndex        =   300
         Top             =   780
         Width           =   435
      End
      Begin VB.Label BQM031_00 
         AutoSize        =   -1  'True
         Caption         =   "g / dl"
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
         Left            =   4920
         TabIndex        =   299
         Top             =   1110
         Width           =   435
      End
      Begin VB.Label BQM031_00 
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
         Left            =   780
         TabIndex        =   298
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label BQM031_00 
         AutoSize        =   -1  'True
         Caption         =   "g / dl"
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
         Left            =   4920
         TabIndex        =   297
         Top             =   1425
         Width           =   435
      End
      Begin VB.Label BQM031_00 
         AutoSize        =   -1  'True
         Caption         =   "g / dl"
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
         Left            =   4920
         TabIndex        =   293
         Top             =   450
         Width           =   435
      End
      Begin VB.Label BQM031_00 
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
         TabIndex        =   292
         Top             =   210
         Width           =   975
      End
      Begin VB.Label BQM031_00 
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
         Left            =   3840
         TabIndex        =   291
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label BQM031_00 
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
         TabIndex        =   290
         Top             =   1770
         Width           =   1335
      End
      Begin VB.Label BQM031_00 
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
         Left            =   1800
         TabIndex        =   289
         Top             =   210
         Width           =   1935
      End
   End
   Begin VB.Frame BQM001 
      Caption         =   "Glucosa"
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
      Height          =   1455
      Left            =   60
      TabIndex        =   260
      Top             =   2340
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox BQM001_03 
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
         Left            =   1720
         TabIndex        =   2
         Text            =   "mg / dl"
         Top             =   780
         Width           =   735
      End
      Begin VB.TextBox BQM001_02 
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
         Left            =   360
         MaxLength       =   5
         TabIndex        =   1
         Top             =   780
         Width           =   1335
      End
      Begin VB.TextBox BQM001_06 
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
         TabIndex        =   5
         Top             =   1080
         Width           =   5790
      End
      Begin VB.TextBox BQM001_04 
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
         Left            =   3240
         TabIndex        =   3
         Text            =   "60 - 110"
         Top             =   780
         Width           =   975
      End
      Begin VB.TextBox BQM001_05 
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
         Left            =   6060
         TabIndex        =   4
         Text            =   "Enzimático"
         Top             =   780
         Width           =   1095
      End
      Begin VB.ComboBox BQM001_01 
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
         ItemData        =   "frmBioquimica.frx":2D89
         Left            =   840
         List            =   "frmBioquimica.frx":2D93
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   200
         Width           =   2055
      End
      Begin VB.Label BQM001_00 
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
         Height          =   285
         Index           =   0
         Left            =   420
         TabIndex        =   266
         Top             =   550
         Width           =   1935
      End
      Begin VB.Label BQM001_00 
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
         Height          =   285
         Index           =   5
         Left            =   60
         TabIndex        =   265
         Top             =   1125
         Width           =   1335
      End
      Begin VB.Label BQM001_00 
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
         Height          =   285
         Index           =   2
         Left            =   3300
         TabIndex        =   264
         Top             =   550
         Width           =   1575
      End
      Begin VB.Label BQM001_00 
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
         Height          =   285
         Index           =   3
         Left            =   6120
         TabIndex        =   263
         Top             =   550
         Width           =   975
      End
      Begin VB.Label BQM001_00 
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
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
         Index           =   4
         Left            =   4260
         TabIndex        =   262
         Top             =   810
         Width           =   585
      End
      Begin VB.Label BQM001_00 
         Alignment       =   2  'Center
         Caption         =   "Tipo: "
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
         Index           =   1
         Left            =   240
         TabIndex        =   261
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame BQM012 
      Caption         =   "Proteina Fraccionada (Albúmina y Globulinas)"
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
      Height          =   1815
      Left            =   60
      TabIndex        =   224
      Top             =   2340
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox BQM012_11 
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
         Left            =   3000
         TabIndex        =   72
         Text            =   "g / dl"
         Top             =   1095
         Width           =   495
      End
      Begin VB.TextBox BQM012_07 
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
         Left            =   3000
         TabIndex        =   68
         Text            =   "g / dl"
         Top             =   780
         Width           =   495
      End
      Begin VB.TextBox BQM012_03 
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
         Left            =   3000
         TabIndex        =   64
         Text            =   "g / dl"
         Top             =   460
         Width           =   495
      End
      Begin VB.TextBox BQM012_01 
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
         Index           =   10
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   227
         TabStop         =   0   'False
         Text            =   "Globulinas"
         Top             =   1095
         Width           =   1815
      End
      Begin VB.TextBox BQM012_10 
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
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   71
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox BQM012_12 
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
         Left            =   4080
         TabIndex        =   73
         Text            =   "1.6 - 3.0"
         Top             =   1095
         Width           =   855
      End
      Begin VB.TextBox BQM012_01 
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   226
         TabStop         =   0   'False
         Text            =   "Albúmina"
         Top             =   780
         Width           =   1815
      End
      Begin VB.TextBox BQM012_06 
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
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   67
         Top             =   780
         Width           =   615
      End
      Begin VB.TextBox BQM012_14 
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
         TabIndex        =   75
         Top             =   1425
         Width           =   5670
      End
      Begin VB.TextBox BQM012_08 
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
         Left            =   4080
         TabIndex        =   69
         Text            =   "3.5 - 4.8"
         Top             =   780
         Width           =   855
      End
      Begin VB.TextBox BQM012_09 
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
         TabIndex        =   70
         Text            =   "Colorimétrico"
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox BQM012_13 
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
         TabIndex        =   74
         Text            =   "Colorimétrico"
         Top             =   1095
         Width           =   1095
      End
      Begin VB.TextBox BQM012_05 
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
         TabIndex        =   66
         Text            =   "Colorimétrico"
         Top             =   460
         Width           =   1095
      End
      Begin VB.TextBox BQM012_04 
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
         Left            =   4080
         TabIndex        =   65
         Text            =   "6.1 - 7.9"
         Top             =   460
         Width           =   855
      End
      Begin VB.TextBox BQM012_02 
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
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   63
         Top             =   460
         Width           =   615
      End
      Begin VB.TextBox BQM012_01 
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   225
         TabStop         =   0   'False
         Text            =   "Proteína Total"
         Top             =   460
         Width           =   1815
      End
      Begin VB.Label BQM012_00 
         AutoSize        =   -1  'True
         Caption         =   "g / dl"
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
         Left            =   4980
         TabIndex        =   235
         Top             =   1125
         Width           =   435
      End
      Begin VB.Label BQM012_00 
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
         Left            =   840
         TabIndex        =   234
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label BQM012_00 
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
         Left            =   2280
         TabIndex        =   233
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label BQM012_00 
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
         Left            =   120
         TabIndex        =   232
         Top             =   1455
         Width           =   1335
      End
      Begin VB.Label BQM012_00 
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
         Left            =   4080
         TabIndex        =   231
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label BQM012_00 
         AutoSize        =   -1  'True
         Caption         =   "g / dl"
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
         Left            =   4980
         TabIndex        =   230
         Top             =   810
         Width           =   435
      End
      Begin VB.Label BQM012_00 
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
         Left            =   6120
         TabIndex        =   229
         Top             =   240
         Width           =   975
      End
      Begin VB.Label BQM012_00 
         AutoSize        =   -1  'True
         Caption         =   "g / dl"
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
         Left            =   4980
         TabIndex        =   228
         Top             =   510
         Width           =   435
      End
   End
   Begin VB.Frame BQM021 
      Caption         =   "Depuración de Creatinina"
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
      Height          =   2055
      Left            =   60
      TabIndex        =   273
      Top             =   2340
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox BQM021_14 
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
         Left            =   3210
         TabIndex        =   123
         Text            =   "lt"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox BQM021_16 
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
         Left            =   6120
         TabIndex        =   125
         Top             =   1320
         Width           =   1065
      End
      Begin VB.TextBox BQM021_15 
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
         TabIndex        =   124
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox BQM021_13 
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
         Left            =   2445
         MaxLength       =   5
         TabIndex        =   122
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox BQM021_09 
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
         Left            =   2445
         MaxLength       =   5
         TabIndex        =   118
         Top             =   1020
         Width           =   735
      End
      Begin VB.TextBox BQM021_11 
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
         TabIndex        =   120
         Text            =   "80 - 140"
         Top             =   1020
         Width           =   735
      End
      Begin VB.TextBox BQM021_12 
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
         Left            =   6120
         TabIndex        =   121
         Text            =   "Colorimétrico"
         Top             =   1020
         Width           =   1065
      End
      Begin VB.TextBox BQM021_10 
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
         Left            =   3210
         TabIndex        =   119
         Text            =   "ml / min"
         Top             =   1020
         Width           =   735
      End
      Begin VB.TextBox BQM021_05 
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
         Left            =   2445
         MaxLength       =   5
         TabIndex        =   114
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox BQM021_07 
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
         TabIndex        =   116
         Text            =   "0.9 - 1.5"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox BQM021_08 
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
         Left            =   6120
         TabIndex        =   117
         Text            =   "Colorimétrico"
         Top             =   720
         Width           =   1065
      End
      Begin VB.TextBox BQM021_06 
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
         Left            =   3210
         TabIndex        =   115
         Text            =   "g / 24 h"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox BQM021_02 
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
         Left            =   3210
         TabIndex        =   111
         Text            =   "mg / l"
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox BQM021_04 
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
         Left            =   6120
         TabIndex        =   113
         Text            =   "Colorimétrico"
         Top             =   420
         Width           =   1065
      End
      Begin VB.TextBox BQM021_03 
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
         TabIndex        =   112
         Text            =   "8 - 14"
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox BQM021_17 
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
         TabIndex        =   126
         Top             =   1660
         Width           =   5670
      End
      Begin VB.TextBox BQM021_01 
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
         Left            =   2445
         MaxLength       =   5
         TabIndex        =   110
         Top             =   420
         Width           =   735
      End
      Begin VB.Label BQM021_00 
         Caption         =   "Volumen de Orina"
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
         Left            =   120
         TabIndex        =   313
         Top             =   1350
         Width           =   2175
      End
      Begin VB.Label BQM021_00 
         AutoSize        =   -1  'True
         Caption         =   "lt"
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
         Left            =   5100
         TabIndex        =   312
         Top             =   1350
         Width           =   105
      End
      Begin VB.Label BQM021_00 
         AutoSize        =   -1  'True
         Caption         =   "ml / min"
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
         Left            =   5100
         TabIndex        =   307
         Top             =   1050
         Width           =   660
      End
      Begin VB.Label BQM021_00 
         AutoSize        =   -1  'True
         Caption         =   "g / 24 h"
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
         Left            =   5100
         TabIndex        =   306
         Top             =   750
         Width           =   675
      End
      Begin VB.Label BQM021_00 
         Caption         =   "Dep. Creatinina Endógena"
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
         Left            =   120
         TabIndex        =   305
         Top             =   1050
         Width           =   2175
      End
      Begin VB.Label BQM021_00 
         Caption         =   "Creatinina en Orina"
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
         Index           =   6
         Left            =   120
         TabIndex        =   304
         Top             =   750
         Width           =   1695
      End
      Begin VB.Label BQM021_00 
         Caption         =   "Creatinina en Suero"
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
         Left            =   120
         TabIndex        =   303
         Top             =   450
         Width           =   1695
      End
      Begin VB.Label BQM021_00 
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
         Index           =   3
         Left            =   5100
         TabIndex        =   278
         Top             =   450
         Width           =   480
      End
      Begin VB.Label BQM021_00 
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
         Left            =   6120
         TabIndex        =   277
         Top             =   210
         Width           =   975
      End
      Begin VB.Label BQM021_00 
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
         Left            =   4260
         TabIndex        =   276
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label BQM021_00 
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
         Left            =   120
         TabIndex        =   275
         Top             =   1690
         Width           =   1335
      End
      Begin VB.Label BQM021_00 
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
         Left            =   2460
         TabIndex        =   274
         Top             =   210
         Width           =   1455
      End
   End
   Begin VB.Frame BQM009 
      Caption         =   "Bilirrubina Total y  Fraccionada"
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
      Height          =   1815
      Left            =   60
      TabIndex        =   242
      Top             =   2370
      Visible         =   0   'False
      Width           =   7185
      Begin VB.ComboBox BQM009_15 
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
         ItemData        =   "frmBioquimica.frx":2DAB
         Left            =   3240
         List            =   "frmBioquimica.frx":2DB8
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   460
         Width           =   975
      End
      Begin VB.TextBox BQM009_11 
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
         Left            =   2520
         TabIndex        =   49
         Text            =   "mg / dl"
         Top             =   1095
         Width           =   615
      End
      Begin VB.TextBox BQM009_07 
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
         Left            =   2520
         TabIndex        =   45
         Text            =   "mg / dl"
         Top             =   780
         Width           =   615
      End
      Begin VB.TextBox BQM009_03 
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
         Left            =   2520
         TabIndex        =   40
         Text            =   "mg / dl"
         Top             =   460
         Width           =   615
      End
      Begin VB.TextBox BQM009_13 
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
         TabIndex        =   51
         Text            =   "Colorimétrico"
         Top             =   1095
         Width           =   1095
      End
      Begin VB.TextBox BQM009_09 
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
         TabIndex        =   47
         Text            =   "Colorimétrico"
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox BQM009_08 
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
         TabIndex        =   46
         Text            =   "0.1 - 0.4"
         Top             =   780
         Width           =   855
      End
      Begin VB.TextBox BQM009_14 
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
         TabIndex        =   52
         Top             =   1425
         Width           =   5790
      End
      Begin VB.TextBox BQM009_06 
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
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   44
         Top             =   780
         Width           =   615
      End
      Begin VB.TextBox BQM009_01 
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
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   245
         TabStop         =   0   'False
         Text            =   "Bilirrubina Directa"
         Top             =   780
         Width           =   1575
      End
      Begin VB.TextBox BQM009_12 
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
         TabIndex        =   50
         Text            =   "0.2 - 0.8"
         Top             =   1095
         Width           =   855
      End
      Begin VB.TextBox BQM009_10 
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
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   48
         Top             =   1095
         Width           =   615
      End
      Begin VB.TextBox BQM009_01 
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
         Index           =   10
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   244
         TabStop         =   0   'False
         Text            =   "Bilirrubina Indirecta"
         Top             =   1095
         Width           =   1575
      End
      Begin VB.TextBox BQM009_01 
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
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   243
         TabStop         =   0   'False
         Text            =   "Bilirrubina Total"
         Top             =   460
         Width           =   1575
      End
      Begin VB.TextBox BQM009_02 
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
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   39
         Top             =   460
         Width           =   615
      End
      Begin VB.TextBox BQM009_04 
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
         TabIndex        =   42
         Text            =   "Hasta 1.0"
         Top             =   460
         Width           =   855
      End
      Begin VB.TextBox BQM009_05 
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
         TabIndex        =   43
         Text            =   "Colorimétrico"
         Top             =   460
         Width           =   1095
      End
      Begin VB.Label BQM009_00 
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
         TabIndex        =   253
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label BQM009_00 
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
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
         Left            =   5220
         TabIndex        =   252
         Top             =   810
         Width           =   585
      End
      Begin VB.Label BQM009_00 
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
         Left            =   3600
         TabIndex        =   251
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label BQM009_00 
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
         TabIndex        =   250
         Top             =   1455
         Width           =   1335
      End
      Begin VB.Label BQM009_00 
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
         Left            =   1920
         TabIndex        =   249
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label BQM009_00 
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
         Left            =   120
         TabIndex        =   248
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label BQM009_00 
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
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
         Left            =   5220
         TabIndex        =   247
         Top             =   1125
         Width           =   585
      End
      Begin VB.Label BQM009_00 
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
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
         Left            =   5220
         TabIndex        =   246
         Top             =   495
         Width           =   585
      End
   End
   Begin VB.Frame BQM004 
      Caption         =   "Colesterol Fraccionado (HDL, LDL)"
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
      Height          =   1455
      Left            =   60
      TabIndex        =   154
      Top             =   2340
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox BQM004_03 
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
         Left            =   2700
         TabIndex        =   12
         Text            =   "mg / dl"
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox BQM004_07 
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
         Left            =   2700
         TabIndex        =   16
         Text            =   "mg / dl"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox BQM004_01 
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   156
         TabStop         =   0   'False
         Text            =   "Colesterol LDL"
         Top             =   770
         Width           =   1695
      End
      Begin VB.TextBox BQM004_06 
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
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   15
         Top             =   740
         Width           =   615
      End
      Begin VB.TextBox BQM004_08 
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
         Left            =   4080
         TabIndex        =   17
         Text            =   "< 129"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox BQM004_01 
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   155
         TabStop         =   0   'False
         Text            =   "Colesterol HDL"
         Top             =   450
         Width           =   1695
      End
      Begin VB.TextBox BQM004_02 
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
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   11
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox BQM004_10 
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
         TabIndex        =   153
         Top             =   1065
         Width           =   5730
      End
      Begin VB.TextBox BQM004_04 
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
         Left            =   4080
         TabIndex        =   13
         Text            =   "40 - 60"
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox BQM004_05 
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
         TabIndex        =   14
         Text            =   "Enzimático"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox BQM004_09 
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
         TabIndex        =   18
         Text            =   "Enzimático"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label BQM004_00 
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
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
         Index           =   5
         Left            =   4860
         TabIndex        =   163
         Top             =   750
         Width           =   480
      End
      Begin VB.Label BQM004_00 
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
         Left            =   420
         TabIndex        =   162
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label BQM004_00 
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
         Left            =   2100
         TabIndex        =   161
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label BQM004_00 
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
         Index           =   6
         Left            =   60
         TabIndex        =   160
         Top             =   1095
         Width           =   1335
      End
      Begin VB.Label BQM004_00 
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
         Left            =   4020
         TabIndex        =   159
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label BQM004_00 
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
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
         Index           =   4
         Left            =   4860
         TabIndex        =   158
         Top             =   450
         Width           =   480
      End
      Begin VB.Label BQM004_00 
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
         Left            =   6060
         TabIndex        =   157
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame BQM008 
      Caption         =   "Billirubina Total"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   254
      Top             =   2310
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox BQM008_02 
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
         Left            =   1730
         TabIndex        =   35
         Text            =   "mg / dl"
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox BQM008_04 
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
         TabIndex        =   37
         Text            =   "Colorimétrico"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox BQM008_03 
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
         Left            =   3480
         TabIndex        =   36
         Text            =   "0.3 - 1.3"
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox BQM008_05 
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
         TabIndex        =   38
         Top             =   740
         Width           =   5670
      End
      Begin VB.TextBox BQM008_01 
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
         Left            =   360
         MaxLength       =   5
         TabIndex        =   34
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label BQM008_00 
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
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
         Left            =   4260
         TabIndex        =   259
         Top             =   450
         Width           =   585
      End
      Begin VB.Label BQM008_00 
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
         TabIndex        =   258
         Top             =   210
         Width           =   975
      End
      Begin VB.Label BQM008_00 
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
         Left            =   3360
         TabIndex        =   257
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label BQM008_00 
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
         TabIndex        =   256
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label BQM008_00 
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
         Left            =   480
         TabIndex        =   255
         Top             =   210
         Width           =   1935
      End
   End
   Begin VB.Frame BQM003 
      Caption         =   "Colesterol Total"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   218
      Top             =   2340
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox BQM003_02 
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
         Left            =   1730
         TabIndex        =   7
         Text            =   "mg / dl"
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox BQM003_04 
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
         Left            =   6060
         TabIndex        =   9
         Text            =   "Enzimático"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox BQM003_03 
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
         Left            =   3240
         TabIndex        =   8
         Text            =   "140 - 200"
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox BQM003_05 
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
         TabIndex        =   10
         Top             =   740
         Width           =   5790
      End
      Begin VB.TextBox BQM003_01 
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
         Left            =   360
         MaxLength       =   5
         TabIndex        =   6
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label BQM003_00 
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
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
         Left            =   4260
         TabIndex        =   223
         Top             =   450
         Width           =   585
      End
      Begin VB.Label BQM003_00 
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
         Left            =   6120
         TabIndex        =   222
         Top             =   220
         Width           =   975
      End
      Begin VB.Label BQM003_00 
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
         Left            =   3300
         TabIndex        =   221
         Top             =   220
         Width           =   1575
      End
      Begin VB.Label BQM003_00 
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
         TabIndex        =   220
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label BQM003_00 
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
         Left            =   420
         TabIndex        =   219
         Top             =   220
         Width           =   1935
      End
   End
   Begin VB.Frame BQM005 
      Caption         =   "Triglicéridos"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   164
      Top             =   2340
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox BQM005_02 
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
         Left            =   1740
         TabIndex        =   20
         Text            =   "mg / dl"
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox BQM005_01 
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
         Left            =   360
         MaxLength       =   5
         TabIndex        =   19
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox BQM005_05 
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
      Begin VB.TextBox BQM005_03 
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
         TabIndex        =   21
         Text            =   "25 - 160"
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox BQM005_04 
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
         TabIndex        =   22
         Text            =   "Enzimático"
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label BQM005_00 
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
         Left            =   420
         TabIndex        =   169
         Top             =   220
         Width           =   1935
      End
      Begin VB.Label BQM005_00 
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
         TabIndex        =   168
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label BQM005_00 
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
         Left            =   3300
         TabIndex        =   167
         Top             =   220
         Width           =   1575
      End
      Begin VB.Label BQM005_00 
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
         TabIndex        =   166
         Top             =   220
         Width           =   975
      End
      Begin VB.Label BQM005_00 
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
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
         Left            =   4260
         TabIndex        =   165
         Top             =   450
         Width           =   585
      End
   End
   Begin VB.Frame BQM006 
      Caption         =   "Transaminasa GPT"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   170
      Top             =   2340
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox BQM006_02 
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
         Left            =   1740
         TabIndex        =   25
         Text            =   "U / L"
         Top             =   420
         Width           =   495
      End
      Begin VB.TextBox BQM006_04 
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
         TabIndex        =   27
         Text            =   "Colorimétrico"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox BQM006_03 
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
         TabIndex        =   26
         Text            =   "Hasta 12"
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox BQM006_05 
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
      Begin VB.TextBox BQM006_01 
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
         Left            =   360
         MaxLength       =   5
         TabIndex        =   24
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label BQM006_00 
         AutoSize        =   -1  'True
         Caption         =   "U / L"
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
         Index           =   3
         Left            =   4260
         TabIndex        =   175
         Top             =   450
         Width           =   330
      End
      Begin VB.Label BQM006_00 
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
         TabIndex        =   174
         Top             =   220
         Width           =   975
      End
      Begin VB.Label BQM006_00 
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
         Left            =   3300
         TabIndex        =   173
         Top             =   220
         Width           =   1575
      End
      Begin VB.Label BQM006_00 
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
      Begin VB.Label BQM006_00 
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
         Left            =   420
         TabIndex        =   171
         Top             =   220
         Width           =   1935
      End
   End
   Begin VB.Frame BQM007 
      Caption         =   "Transaminasa GOT"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   176
      Top             =   2340
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox BQM007_02 
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
         Left            =   1730
         TabIndex        =   30
         Text            =   "U / L"
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox BQM007_01 
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
         Left            =   360
         MaxLength       =   5
         TabIndex        =   29
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox BQM007_05 
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
         TabIndex        =   33
         Top             =   740
         Width           =   5670
      End
      Begin VB.TextBox BQM007_03 
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
         Left            =   3480
         TabIndex        =   31
         Text            =   "Hasta 12"
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox BQM007_04 
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
         TabIndex        =   32
         Text            =   "Colorimétrico"
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label BQM007_00 
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
         Left            =   450
         TabIndex        =   181
         Top             =   210
         Width           =   1935
      End
      Begin VB.Label BQM007_00 
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
         TabIndex        =   180
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label BQM007_00 
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
         Left            =   3360
         TabIndex        =   179
         Top             =   220
         Width           =   1575
      End
      Begin VB.Label BQM007_00 
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
         TabIndex        =   178
         Top             =   220
         Width           =   975
      End
      Begin VB.Label BQM007_00 
         AutoSize        =   -1  'True
         Caption         =   "U / L"
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
         Index           =   3
         Left            =   4380
         TabIndex        =   177
         Top             =   450
         Width           =   330
      End
   End
   Begin VB.Frame BQM010 
      Caption         =   "Fosfatasa Alcalina"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   212
      Top             =   2340
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox BQM010_02 
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
         Left            =   1710
         TabIndex        =   54
         Text            =   "UI /  L"
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox BQM010_04 
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
         Left            =   6060
         TabIndex        =   56
         Text            =   "Colorimétrico"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox BQM010_03 
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
         TabIndex        =   55
         Text            =   "68 - 240"
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox BQM010_05 
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
         TabIndex        =   57
         Top             =   740
         Width           =   5790
      End
      Begin VB.TextBox BQM010_01 
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
         Left            =   360
         MaxLength       =   5
         TabIndex        =   53
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label BQM010_00 
         AutoSize        =   -1  'True
         Caption         =   "UI /  L"
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
         Left            =   4260
         TabIndex        =   217
         Top             =   450
         Width           =   525
      End
      Begin VB.Label BQM010_00 
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
         Left            =   6120
         TabIndex        =   216
         Top             =   220
         Width           =   975
      End
      Begin VB.Label BQM010_00 
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
         Left            =   3300
         TabIndex        =   215
         Top             =   220
         Width           =   1575
      End
      Begin VB.Label BQM010_00 
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
         TabIndex        =   214
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label BQM010_00 
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
         Left            =   420
         TabIndex        =   213
         Top             =   220
         Width           =   1935
      End
   End
   Begin VB.Frame BQM011 
      Caption         =   "Proteina Total"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   236
      Top             =   2340
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox BQM011_02 
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
         Left            =   1740
         TabIndex        =   59
         Text            =   "g / dl"
         Top             =   420
         Width           =   495
      End
      Begin VB.TextBox BQM011_01 
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
         Left            =   360
         MaxLength       =   5
         TabIndex        =   58
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox BQM011_05 
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
         TabIndex        =   62
         Top             =   740
         Width           =   5670
      End
      Begin VB.TextBox BQM011_03 
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
         Left            =   3480
         TabIndex        =   60
         Text            =   "6.1 - 7.9"
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox BQM011_04 
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
         TabIndex        =   61
         Text            =   "Colorimétrico"
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label BQM011_00 
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
         Left            =   480
         TabIndex        =   241
         Top             =   210
         Width           =   1935
      End
      Begin VB.Label BQM011_00 
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
         TabIndex        =   240
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label BQM011_00 
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
         Left            =   3360
         TabIndex        =   239
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label BQM011_00 
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
         TabIndex        =   238
         Top             =   210
         Width           =   975
      End
      Begin VB.Label BQM011_00 
         AutoSize        =   -1  'True
         Caption         =   "g / dl"
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
         Left            =   4380
         TabIndex        =   237
         Top             =   450
         Width           =   435
      End
   End
   Begin VB.Frame BQM014 
      Caption         =   "Amilasa"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   182
      Top             =   2370
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox BQM014_02 
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
         Left            =   1500
         TabIndex        =   77
         Text            =   "U / dl"
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox BQM014_01 
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
         Left            =   360
         MaxLength       =   5
         TabIndex        =   76
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox BQM014_05 
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
         TabIndex        =   80
         Top             =   740
         Width           =   5670
      End
      Begin VB.TextBox BQM014_03 
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
         Left            =   3480
         TabIndex        =   78
         Text            =   "< 120"
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox BQM014_04 
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
         TabIndex        =   79
         Text            =   "Colorimétrico"
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label BQM014_00 
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
         Left            =   360
         TabIndex        =   187
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label BQM014_00 
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
         TabIndex        =   186
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label BQM014_00 
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
         Left            =   3360
         TabIndex        =   185
         Top             =   220
         Width           =   1575
      End
      Begin VB.Label BQM014_00 
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
         TabIndex        =   184
         Top             =   220
         Width           =   975
      End
      Begin VB.Label BQM014_00 
         AutoSize        =   -1  'True
         Caption         =   "U / dl"
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
         Left            =   4260
         TabIndex        =   183
         Top             =   450
         Width           =   450
      End
   End
   Begin VB.Frame BQM015 
      Caption         =   "Úrea"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   188
      Top             =   2370
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox BQM015_02 
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
         Left            =   1720
         TabIndex        =   82
         Text            =   "mg / dl"
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox BQM015_04 
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
         TabIndex        =   84
         Text            =   "Enzimático"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox BQM015_03 
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
         TabIndex        =   83
         Text            =   "20 - 45"
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox BQM015_05 
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
         TabIndex        =   85
         Top             =   740
         Width           =   5670
      End
      Begin VB.TextBox BQM015_01 
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
         Left            =   360
         MaxLength       =   5
         TabIndex        =   81
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label BQM015_00 
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
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
         Left            =   4260
         TabIndex        =   193
         Top             =   450
         Width           =   585
      End
      Begin VB.Label BQM015_00 
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
         TabIndex        =   192
         Top             =   210
         Width           =   975
      End
      Begin VB.Label BQM015_00 
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
         Left            =   3360
         TabIndex        =   191
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label BQM015_00 
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
         TabIndex        =   190
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label BQM015_00 
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
         Left            =   480
         TabIndex        =   189
         Top             =   210
         Width           =   1935
      End
   End
   Begin VB.Frame BQM016 
      Caption         =   "Creatinina"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   194
      Top             =   2340
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox BQM016_02 
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
         Left            =   1120
         TabIndex        =   87
         Text            =   "mg / dl"
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox BQM016_01 
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
         TabIndex        =   86
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox BQM016_06 
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
         TabIndex        =   91
         Top             =   740
         Width           =   5670
      End
      Begin VB.TextBox BQM016_04 
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
         Left            =   3600
         TabIndex        =   89
         Text            =   "0.7 - 1.4"
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox BQM016_05 
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
         TabIndex        =   90
         Text            =   "Colorimétrico"
         Top             =   420
         Width           =   1095
      End
      Begin VB.ComboBox BQM016_03 
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
         ItemData        =   "frmBioquimica.frx":2DD0
         Left            =   2280
         List            =   "frmBioquimica.frx":2DDA
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label BQM016_00 
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
         TabIndex        =   199
         Top             =   210
         Width           =   1455
      End
      Begin VB.Label BQM016_00 
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
         TabIndex        =   198
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label BQM016_00 
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
         TabIndex        =   196
         Top             =   210
         Width           =   975
      End
      Begin VB.Label BQM016_00 
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
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
         Left            =   4500
         TabIndex        =   195
         Top             =   450
         Width           =   585
      End
      Begin VB.Label BQM016_00 
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
         Left            =   2280
         TabIndex        =   197
         Top             =   210
         Width           =   2775
      End
   End
   Begin VB.Frame BQM017 
      Caption         =   "Ácido Úrico"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   200
      Top             =   2340
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox BQM017_02 
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
         Left            =   1140
         TabIndex        =   93
         Text            =   "mg / dl"
         Top             =   420
         Width           =   735
      End
      Begin VB.ComboBox BQM017_03 
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
         ItemData        =   "frmBioquimica.frx":2DEC
         Left            =   2280
         List            =   "frmBioquimica.frx":2DF6
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   390
         Width           =   1215
      End
      Begin VB.TextBox BQM017_05 
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
         TabIndex        =   96
         Text            =   "Enzimático"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox BQM017_04 
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
         Left            =   3600
         TabIndex        =   95
         Text            =   "25 - 60"
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox BQM017_06 
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
         TabIndex        =   97
         Top             =   735
         Width           =   5670
      End
      Begin VB.TextBox BQM017_01 
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
         TabIndex        =   92
         Top             =   420
         Width           =   975
      End
      Begin VB.Label BQM017_00 
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
         TabIndex        =   201
         Top             =   200
         Width           =   1455
      End
      Begin VB.Label BQM017_00 
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
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
         Left            =   4380
         TabIndex        =   205
         Top             =   450
         Width           =   585
      End
      Begin VB.Label BQM017_00 
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
         TabIndex        =   204
         Top             =   200
         Width           =   975
      End
      Begin VB.Label BQM017_00 
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
         Left            =   2280
         TabIndex        =   203
         Top             =   200
         Width           =   2775
      End
      Begin VB.Label BQM017_00 
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
         TabIndex        =   202
         Top             =   765
         Width           =   1335
      End
   End
   Begin VB.Frame BQM018 
      Caption         =   "ADA (Adenosina Deaminasa)"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   267
      Top             =   2310
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox BQM018_02 
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
         Left            =   1140
         TabIndex        =   99
         Text            =   "U / L"
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox BQM018_01 
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
         TabIndex        =   98
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox BQM018_06 
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
         TabIndex        =   103
         Top             =   740
         Width           =   5790
      End
      Begin VB.TextBox BQM018_04 
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
         Left            =   4080
         TabIndex        =   101
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox BQM018_05 
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
         Left            =   6060
         TabIndex        =   102
         Text            =   "Colorimétrico"
         Top             =   420
         Width           =   1095
      End
      Begin VB.ComboBox BQM018_03 
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
         ItemData        =   "frmBioquimica.frx":2E08
         Left            =   2040
         List            =   "frmBioquimica.frx":2E1B
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   400
         Width           =   1935
      End
      Begin VB.Label BQM018_00 
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
         TabIndex        =   272
         Top             =   195
         Width           =   1575
      End
      Begin VB.Label BQM018_00 
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
         TabIndex        =   271
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label BQM018_00 
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
         Left            =   2160
         TabIndex        =   270
         Top             =   195
         Width           =   3495
      End
      Begin VB.Label BQM018_00 
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
         Left            =   6120
         TabIndex        =   269
         Top             =   195
         Width           =   975
      End
      Begin VB.Label BQM018_00 
         AutoSize        =   -1  'True
         Caption         =   "U / L"
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
         Left            =   5220
         TabIndex        =   268
         Top             =   450
         Width           =   405
      End
   End
   Begin VB.Frame BQM019 
      Caption         =   "Calcio"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   206
      Top             =   2370
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox BQM019_02 
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
         Left            =   1140
         TabIndex        =   105
         Text            =   "mg / dl"
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox BQM019_01 
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
         TabIndex        =   104
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox BQM019_06 
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
         TabIndex        =   109
         Top             =   740
         Width           =   5670
      End
      Begin VB.TextBox BQM019_04 
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
         Left            =   3840
         TabIndex        =   107
         Text            =   "8.8 - 10.7"
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox BQM019_05 
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
         TabIndex        =   108
         Text            =   "Colorimétrico"
         Top             =   420
         Width           =   1095
      End
      Begin VB.ComboBox BQM019_03 
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
         ItemData        =   "frmBioquimica.frx":2E6E
         Left            =   2400
         List            =   "frmBioquimica.frx":2E78
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label BQM019_00 
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
         TabIndex        =   211
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label BQM019_00 
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
         TabIndex        =   210
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label BQM019_00 
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
         Left            =   2400
         TabIndex        =   209
         Top             =   210
         Width           =   2895
      End
      Begin VB.Label BQM019_00 
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
         TabIndex        =   208
         Top             =   210
         Width           =   975
      End
      Begin VB.Label BQM019_00 
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
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
         Left            =   4740
         TabIndex        =   207
         Top             =   450
         Width           =   585
      End
   End
   Begin VB.Frame BQM030 
      Caption         =   "Transaminasa GPT y GOT"
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
      Height          =   1470
      Left            =   60
      TabIndex        =   279
      Top             =   2370
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox BQM030_05 
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
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   131
         Top             =   770
         Width           =   1335
      End
      Begin VB.TextBox BQM030_07 
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
         Left            =   4080
         TabIndex        =   133
         Text            =   "Hasta 12"
         Top             =   770
         Width           =   855
      End
      Begin VB.TextBox BQM030_08 
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
         TabIndex        =   134
         Text            =   "Colorimétrico"
         Top             =   770
         Width           =   1095
      End
      Begin VB.TextBox BQM030_06 
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
         Left            =   3180
         TabIndex        =   132
         Text            =   "U / L"
         Top             =   770
         Width           =   495
      End
      Begin VB.TextBox BQM030_01 
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
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   127
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox BQM030_09 
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
         TabIndex        =   135
         Top             =   1100
         Width           =   5670
      End
      Begin VB.TextBox BQM030_03 
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
         Left            =   4080
         TabIndex        =   129
         Text            =   "Hasta 12"
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox BQM030_04 
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
         TabIndex        =   130
         Text            =   "Colorimétrico"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox BQM030_02 
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
         Left            =   3180
         TabIndex        =   128
         Text            =   "U / L"
         Top             =   420
         Width           =   495
      End
      Begin VB.Label BQM030_00 
         AutoSize        =   -1  'True
         Caption         =   "U / L"
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
         Left            =   4980
         TabIndex        =   287
         Top             =   795
         Width           =   405
      End
      Begin VB.Label BQM030_00 
         AutoSize        =   -1  'True
         Caption         =   "Transaminasa GOT"
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
         Left            =   60
         TabIndex        =   286
         Top             =   795
         Width           =   1515
      End
      Begin VB.Label BQM030_00 
         AutoSize        =   -1  'True
         Caption         =   "Transaminasa GPT"
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
         Left            =   60
         TabIndex        =   285
         Top             =   450
         Width           =   1485
      End
      Begin VB.Label BQM030_00 
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
         Index           =   9
         Left            =   1860
         TabIndex        =   284
         Top             =   225
         Width           =   1815
      End
      Begin VB.Label BQM030_00 
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
         TabIndex        =   283
         Top             =   1130
         Width           =   1335
      End
      Begin VB.Label BQM030_00 
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
         Index           =   7
         Left            =   4020
         TabIndex        =   282
         Top             =   225
         Width           =   1575
      End
      Begin VB.Label BQM030_00 
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
         Index           =   6
         Left            =   6000
         TabIndex        =   281
         Top             =   220
         Width           =   975
      End
      Begin VB.Label BQM030_00 
         AutoSize        =   -1  'True
         Caption         =   "U / L"
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
         Left            =   4980
         TabIndex        =   280
         Top             =   450
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmBioquimica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Resultados de Bioquímica
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
  Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =16")
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
  Fra.Top = UcPacienteDatos1.Top + UcPacienteDatos1.Height + 700   '350
  fraBoton.Top = Fra.Top + Fra.Height
  Me.Height = fraBoton.Top + fraBoton.Height + 500
End Sub

Private Sub BQM001_01_Click()
  If BQM001_01.ListIndex = 0 Then
    BQM001_04.Text = "60 - 110"
  Else
    BQM001_04.Text = "70 - 140"
  End If
End Sub

Private Sub BQM001_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM001_02_GotFocus()
  SeleccionaTexto BQM001_02
End Sub

Private Sub BQM001_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM001_03_Change()
  BQM001_00(3).Caption = BQM003_03.Text
End Sub

Private Sub BQM001_03_GotFocus()
  SeleccionaTexto BQM001_03
End Sub

Private Sub BQM001_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM001_04_GotFocus()
  SeleccionaTexto BQM001_04
End Sub

Private Sub BQM001_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM001_05_GotFocus()
  SeleccionaTexto BQM001_05
End Sub

Private Sub BQM001_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM001_06_GotFocus()
  SeleccionaTexto BQM001_06
End Sub

Private Sub BQM001_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM003_01_GotFocus()
  SeleccionaTexto BQM003_01
End Sub

Private Sub BQM003_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM003_02_Change()
  BQM003_00(3).Caption = BQM003_02.Text
End Sub

Private Sub BQM003_02_GotFocus()
  SeleccionaTexto BQM003_02
End Sub

Private Sub BQM003_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM003_03_GotFocus()
  SeleccionaTexto BQM003_03
End Sub

Private Sub BQM003_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM003_04_GotFocus()
  SeleccionaTexto BQM003_04
End Sub

Private Sub BQM003_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM003_05_GotFocus()
  SeleccionaTexto BQM003_05
End Sub

Private Sub BQM003_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM004_02_GotFocus()
  SeleccionaTexto BQM004_02
End Sub

Private Sub BQM004_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM004_03_Change()
  BQM004_00(4).Caption = BQM004_03.Text
End Sub

Private Sub BQM004_03_GotFocus()
  SeleccionaTexto BQM004_03
End Sub

Private Sub BQM004_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM004_04_GotFocus()
  SeleccionaTexto BQM004_04
End Sub

Private Sub BQM004_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM004_05_GotFocus()
  SeleccionaTexto BQM004_05
End Sub

Private Sub BQM004_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM004_06_GotFocus()
  SeleccionaTexto BQM004_06
End Sub

Private Sub BQM004_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM004_07_Change()
  BQM004_00(5).Caption = BQM004_07.Text
End Sub

Private Sub BQM004_07_GotFocus()
  SeleccionaTexto BQM004_07
End Sub

Private Sub BQM004_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM004_08_GotFocus()
  SeleccionaTexto BQM004_08
End Sub

Private Sub BQM004_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM004_09_GotFocus()
  SeleccionaTexto BQM004_09
End Sub

Private Sub BQM004_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM004_10_GotFocus()
  SeleccionaTexto BQM004_10
End Sub

Private Sub BQM004_10_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM005_01_GotFocus()
  SeleccionaTexto BQM005_01
End Sub

Private Sub BQM005_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM005_02_Change()
  BQM005_00(3).Caption = BQM005_02.Text
End Sub

Private Sub BQM005_02_GotFocus()
  SeleccionaTexto BQM005_02
End Sub

Private Sub BQM005_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM005_03_GotFocus()
  SeleccionaTexto BQM005_03
End Sub

Private Sub BQM005_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM005_04_GotFocus()
  SeleccionaTexto BQM005_04
End Sub

Private Sub BQM005_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM005_05_GotFocus()
  SeleccionaTexto BQM005_05
End Sub

Private Sub BQM005_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM006_01_GotFocus()
  SeleccionaTexto BQM006_01
End Sub

Private Sub BQM006_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM006_02_Change()
  BQM006_00(3).Caption = BQM006_02.Text
End Sub

Private Sub BQM006_02_GotFocus()
  SeleccionaTexto BQM006_02
End Sub

Private Sub BQM006_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM006_03_GotFocus()
  SeleccionaTexto BQM006_03
End Sub

Private Sub BQM006_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM006_04_GotFocus()
  SeleccionaTexto BQM006_04
End Sub

Private Sub BQM006_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM006_05_GotFocus()
  SeleccionaTexto BQM006_05
End Sub

Private Sub BQM006_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM007_01_GotFocus()
  SeleccionaTexto BQM007_01
End Sub

Private Sub BQM007_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM007_02_Change()
  BQM007_00(3).Caption = BQM007_02.Text
End Sub

Private Sub BQM007_02_GotFocus()
  SeleccionaTexto BQM007_02
End Sub

Private Sub BQM007_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM007_03_GotFocus()
  SeleccionaTexto BQM007_03
End Sub

Private Sub BQM007_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM007_04_GotFocus()
  SeleccionaTexto BQM007_04
End Sub

Private Sub BQM007_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM007_05_GotFocus()
  SeleccionaTexto BQM007_05
End Sub

Private Sub BQM007_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM008_01_GotFocus()
  SeleccionaTexto BQM008_01
End Sub

Private Sub BQM008_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM008_02_Change()
  BQM008_00(3).Caption = BQM008_02.Text
End Sub

Private Sub BQM008_02_GotFocus()
  SeleccionaTexto BQM008_02
End Sub

Private Sub BQM008_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM008_03_GotFocus()
  SeleccionaTexto BQM008_03
End Sub

Private Sub BQM008_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM008_04_GotFocus()
  SeleccionaTexto BQM008_04
End Sub

Private Sub BQM008_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM008_05_GotFocus()
  SeleccionaTexto BQM008_05
End Sub

Private Sub BQM008_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM009_02_GotFocus()
  SeleccionaTexto BQM009_02
End Sub

Private Sub BQM009_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM009_03_Change()
  BQM009_00(4).Caption = BQM009_03.Text
End Sub

Private Sub BQM009_03_GotFocus()
  SeleccionaTexto BQM009_03
End Sub

Private Sub BQM009_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM009_04_GotFocus()
  SeleccionaTexto BQM009_04
End Sub

Private Sub BQM009_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM009_05_GotFocus()
  SeleccionaTexto BQM009_05
End Sub

Private Sub BQM009_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM009_06_GotFocus()
  SeleccionaTexto BQM009_06
End Sub

Private Sub BQM009_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM009_07_Change()
  BQM009_00(5).Caption = BQM009_07.Text
End Sub

Private Sub BQM009_07_GotFocus()
  SeleccionaTexto BQM009_07
End Sub

Private Sub BQM009_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM009_08_GotFocus()
  SeleccionaTexto BQM009_08
End Sub

Private Sub BQM009_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM009_09_GotFocus()
  SeleccionaTexto BQM009_09
End Sub

Private Sub BQM009_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM009_10_GotFocus()
  SeleccionaTexto BQM009_10
End Sub

Private Sub BQM009_10_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM009_11_Change()
  BQM009_00(6).Caption = BQM009_11.Text
End Sub

Private Sub BQM009_11_GotFocus()
  SeleccionaTexto BQM009_11
End Sub

Private Sub BQM009_11_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM009_12_GotFocus()
  SeleccionaTexto BQM009_12
End Sub

Private Sub BQM009_12_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM009_13_GotFocus()
  SeleccionaTexto BQM009_13
End Sub

Private Sub BQM009_13_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM009_14_GotFocus()
  SeleccionaTexto BQM009_14
End Sub

Private Sub BQM009_14_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM009_15_Click()
  If BQM009_15.ListIndex = 0 Then
    BQM009_04.Text = "Hasta 1.0"
    BQM009_08.Text = "0.1 - 0.4"
    BQM009_12.Text = "0.2 - 0.8"
  ElseIf BQM009_15.ListIndex = 1 Then
    BQM009_04.Text = ""
    BQM009_08.Text = ""
    BQM009_12.Text = ""
  Else
    BQM009_04.Text = ""
    BQM009_08.Text = ""
    BQM009_12.Text = ""
  End If
End Sub

Private Sub BQM009_15_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM010_01_GotFocus()
  SeleccionaTexto BQM010_01
End Sub

Private Sub BQM010_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM010_02_Change()
  BQM010_00(3).Caption = BQM010_02.Text
End Sub

Private Sub BQM010_02_GotFocus()
  SeleccionaTexto BQM010_02
End Sub

Private Sub BQM010_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM010_03_GotFocus()
  SeleccionaTexto BQM010_03
End Sub

Private Sub BQM010_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM010_04_GotFocus()
  SeleccionaTexto BQM010_04
End Sub

Private Sub BQM010_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM010_05_GotFocus()
  SeleccionaTexto BQM010_05
End Sub

Private Sub BQM010_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM011_01_GotFocus()
  SeleccionaTexto BQM011_01
End Sub

Private Sub BQM011_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM011_02_Change()
  BQM011_00(3).Caption = BQM011_02.Text
End Sub

Private Sub BQM011_02_GotFocus()
  SeleccionaTexto BQM011_02
End Sub

Private Sub BQM011_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM011_03_GotFocus()
  SeleccionaTexto BQM011_03
End Sub

Private Sub BQM011_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM011_04_GotFocus()
  SeleccionaTexto BQM011_04
End Sub

Private Sub BQM011_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM011_05_GotFocus()
  SeleccionaTexto BQM011_05
End Sub

Private Sub BQM011_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM012_02_GotFocus()
  SeleccionaTexto BQM012_02
End Sub

Private Sub BQM012_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM012_03_Change()
  BQM012_00(4).Caption = BQM012_03.Text
End Sub

Private Sub BQM012_03_GotFocus()
  SeleccionaTexto BQM012_03
End Sub

Private Sub BQM012_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM012_04_GotFocus()
  SeleccionaTexto BQM012_04
End Sub

Private Sub BQM012_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM012_05_GotFocus()
  SeleccionaTexto BQM012_05
End Sub

Private Sub BQM012_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM012_06_GotFocus()
  SeleccionaTexto BQM012_06
End Sub

Private Sub BQM012_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM012_07_Change()
  BQM012_00(5).Caption = BQM012_07.Text
End Sub

Private Sub BQM012_07_GotFocus()
  SeleccionaTexto BQM012_07
End Sub

Private Sub BQM012_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM012_08_GotFocus()
  SeleccionaTexto BQM012_08
End Sub

Private Sub BQM012_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM012_09_GotFocus()
  SeleccionaTexto BQM012_09
End Sub

Private Sub BQM012_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM012_10_GotFocus()
  SeleccionaTexto BQM012_10
End Sub

Private Sub BQM012_10_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM012_11_Change()
  BQM012_00(6).Caption = BQM012_11.Text
End Sub

Private Sub BQM012_11_GotFocus()
  SeleccionaTexto BQM012_11
End Sub

Private Sub BQM012_11_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM012_12_GotFocus()
  SeleccionaTexto BQM012_12
End Sub

Private Sub BQM012_12_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM012_13_GotFocus()
  SeleccionaTexto BQM012_13
End Sub

Private Sub BQM012_13_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM012_14_GotFocus()
  SeleccionaTexto BQM012_14
End Sub

Private Sub BQM012_14_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM014_01_GotFocus()
  SeleccionaTexto BQM014_01
End Sub

Private Sub BQM014_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM014_02_Change()
  BQM014_00(3).Caption = BQM014_02.Text
End Sub

Private Sub BQM014_02_GotFocus()
  SeleccionaTexto BQM014_02
End Sub

Private Sub BQM014_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM014_03_GotFocus()
  SeleccionaTexto BQM014_03
End Sub

Private Sub BQM014_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM014_04_GotFocus()
  SeleccionaTexto BQM014_04
End Sub

Private Sub BQM014_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM014_05_GotFocus()
  SeleccionaTexto BQM014_05
End Sub

Private Sub BQM014_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM015_01_GotFocus()
  SeleccionaTexto BQM015_01
End Sub

Private Sub BQM015_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM015_02_Change()
  BQM015_00(3).Caption = BQM015_02.Text
End Sub

Private Sub BQM015_02_GotFocus()
  SeleccionaTexto BQM015_02
End Sub

Private Sub BQM015_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM015_03_GotFocus()
  SeleccionaTexto BQM015_03
End Sub

Private Sub BQM015_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM015_04_GotFocus()
  SeleccionaTexto BQM015_04
End Sub

Private Sub BQM015_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM015_05_GotFocus()
  SeleccionaTexto BQM015_05
End Sub

Private Sub BQM015_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM016_01_GotFocus()
  SeleccionaTexto BQM016_01
End Sub

Private Sub BQM016_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM016_02_Change()
  BQM016_00(3).Caption = BQM016_02.Text
End Sub

Private Sub BQM016_02_GotFocus()
  SeleccionaTexto BQM016_02
End Sub

Private Sub BQM016_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM016_03_Click()
  If BQM016_03.ListIndex = 0 Then
    BQM016_04.Text = "0.7 - 1.4"
  Else
    BQM016_04.Text = "0.6 - 1.2"
  End If
End Sub

Private Sub BQM016_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM016_04_GotFocus()
  SeleccionaTexto BQM016_04
End Sub

Private Sub BQM016_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM016_05_GotFocus()
  SeleccionaTexto BQM016_05
End Sub

Private Sub BQM016_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM016_06_GotFocus()
  SeleccionaTexto BQM016_06
End Sub

Private Sub BQM016_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM017_01_GotFocus()
  SeleccionaTexto BQM017_01
End Sub

Private Sub BQM017_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM017_02_Change()
  BQM017_00(3).Caption = BQM017_02.Text
End Sub

Private Sub BQM017_02_GotFocus()
  SeleccionaTexto BQM017_02
End Sub

Private Sub BQM017_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM017_03_Click()
  If BQM017_03.ListIndex = 0 Then
    BQM017_04.Text = "25 - 60"
  Else
    BQM017_04.Text = "20 - 50"
  End If
End Sub

Private Sub BQM017_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM017_04_GotFocus()
  SeleccionaTexto BQM017_04
End Sub

Private Sub BQM017_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM017_05_GotFocus()
  SeleccionaTexto BQM017_05
End Sub

Private Sub BQM017_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM017_06_GotFocus()
  SeleccionaTexto BQM017_06
End Sub

Private Sub BQM017_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM018_01_GotFocus()
  SeleccionaTexto BQM018_01
End Sub

Private Sub BQM018_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM018_02_Change()
  BQM018_00(3).Caption = BQM018_02.Text
End Sub

Private Sub BQM018_02_GotFocus()
  SeleccionaTexto BQM018_02
End Sub

Private Sub BQM018_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM018_03_Click()
  If BQM018_03.ListIndex = 0 Then
    BQM018_04.Text = "13 - 21"
  ElseIf BQM018_03.ListIndex = 1 Then
    BQM018_04.Text = "0 - 45"
  ElseIf BQM018_03.ListIndex = 2 Then
    BQM018_04.Text = "0 - 6"
  ElseIf BQM018_03.ListIndex = 3 Then
    BQM018_04.Text = "0 - 40"
  Else
    BQM018_04.Text = "0 - 20"
  End If
End Sub

Private Sub BQM018_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM018_04_GotFocus()
  SeleccionaTexto BQM018_04
End Sub

Private Sub BQM018_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM018_05_GotFocus()
  SeleccionaTexto BQM018_05
End Sub

Private Sub BQM018_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM018_06_GotFocus()
  SeleccionaTexto BQM018_06
End Sub

Private Sub BQM018_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM019_01_GotFocus()
  SeleccionaTexto BQM019_01
End Sub

Private Sub BQM019_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM019_02_Change()
  BQM019_00(3).Caption = BQM019_02.Text
End Sub

Private Sub BQM019_02_GotFocus()
  SeleccionaTexto BQM019_02
End Sub

Private Sub BQM019_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM019_03_Click()
  If BQM019_03.ListIndex = 0 Then
    BQM019_04.Text = "8.8 - 10.7"
  Else
    BQM019_04.Text = "10.0 - 12.7"
  End If
End Sub

Private Sub BQM019_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM019_04_GotFocus()
  SeleccionaTexto BQM019_04
End Sub

Private Sub BQM019_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM019_05_GotFocus()
  SeleccionaTexto BQM019_05
End Sub

Private Sub BQM019_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM019_06_GotFocus()
  SeleccionaTexto BQM019_06
End Sub

Private Sub BQM019_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_01_GotFocus()
  SeleccionaTexto BQM021_01
End Sub

Private Sub BQM021_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_02_Change()
  BQM021_00(3).Caption = BQM021_02.Text
End Sub

Private Sub BQM021_02_GotFocus()
  SeleccionaTexto BQM021_02
End Sub

Private Sub BQM021_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_03_GotFocus()
  SeleccionaTexto BQM021_03
End Sub

Private Sub BQM021_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_04_GotFocus()
  SeleccionaTexto BQM021_04
End Sub

Private Sub BQM021_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_05_GotFocus()
  SeleccionaTexto BQM021_05
End Sub

Private Sub BQM021_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_06_Change()
  BQM021_00(8).Caption = BQM021_06.Text
End Sub

Private Sub BQM021_06_GotFocus()
  SeleccionaTexto BQM021_06
End Sub

Private Sub BQM021_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_07_GotFocus()
  SeleccionaTexto BQM021_07
End Sub

Private Sub BQM021_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_08_GotFocus()
  SeleccionaTexto BQM021_08
End Sub

Private Sub BQM021_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_09_GotFocus()
  SeleccionaTexto BQM021_09
End Sub

Private Sub BQM021_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_10_Change()
  BQM021_00(9).Caption = BQM021_10.Text
End Sub

Private Sub BQM021_10_GotFocus()
  SeleccionaTexto BQM021_10
End Sub

Private Sub BQM021_10_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_11_GotFocus()
  SeleccionaTexto BQM021_11
End Sub

Private Sub BQM021_11_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_12_GotFocus()
  SeleccionaTexto BQM021_12
End Sub

Private Sub BQM021_12_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_13_GotFocus()
  SeleccionaTexto BQM021_13
End Sub

Private Sub BQM021_13_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_14_Change()
  BQM021_00(10).Caption = BQM021_14.Text
End Sub

Private Sub BQM021_14_GotFocus()
  SeleccionaTexto BQM021_14
End Sub

Private Sub BQM021_14_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_15_GotFocus()
  SeleccionaTexto BQM021_15
End Sub

Private Sub BQM021_15_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_16_GotFocus()
  SeleccionaTexto BQM021_16
End Sub

Private Sub BQM021_16_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM021_17_GotFocus()
  SeleccionaTexto BQM021_17
End Sub

Private Sub BQM021_17_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM030_01_GotFocus()
  SeleccionaTexto BQM030_01
End Sub

Private Sub BQM030_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM030_02_Change()
  BQM030_00(5).Caption = BQM030_02.Text
End Sub

Private Sub BQM030_02_GotFocus()
  SeleccionaTexto BQM030_02
End Sub

Private Sub BQM030_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM030_03_GotFocus()
  SeleccionaTexto BQM030_03
End Sub

Private Sub BQM030_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM030_04_GotFocus()
  SeleccionaTexto BQM030_04
End Sub

Private Sub BQM030_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM030_05_GotFocus()
  SeleccionaTexto BQM030_05
End Sub

Private Sub BQM030_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM030_06_Change()
  BQM030_00(11).Caption = BQM030_06.Text
End Sub

Private Sub BQM030_06_GotFocus()
  SeleccionaTexto BQM030_06
End Sub

Private Sub BQM030_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM030_07_GotFocus()
  SeleccionaTexto BQM030_07
End Sub

Private Sub BQM030_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM030_08_GotFocus()
  SeleccionaTexto BQM030_08
End Sub

Private Sub BQM030_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM030_09_GotFocus()
  SeleccionaTexto BQM030_09
End Sub

Private Sub BQM030_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_02_Change()
  BQM031_00(4).Caption = BQM031_02.Text
End Sub

Private Sub BQM031_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_06_Change()
  BQM031_00(5).Caption = BQM031_06.Text
End Sub

Private Sub BQM031_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_10_Change()
  BQM031_00(6).Caption = BQM031_10.Text
End Sub

Private Sub BQM031_10_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_11_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_12_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_13_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_14_Change()
  BQM031_00(7).Caption = BQM031_14.Text
End Sub

Private Sub BQM031_14_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_15_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_16_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub BQM031_17_KeyDown(KeyCode As Integer, Shift As Integer)
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
  If txtFresultado.Text = sighentidades.FECHA_VACIA_DMY Then
     MsgBox "Por favor ingrese la Fecha de Resultado", vbInformation, "SIGH "
     Exit Sub
  End If
  ml_nombreRealiza = mo_cmbResponsable.BoundText
  If ml_CodigoPruebaSeleccionada = "BQM001" Then 'Glucosa
    'BQM001
    ml_resultado = BQM001_01.Text & "\" & BQM001_02.Text & "\" & BQM001_03.Text & "\" & BQM001_04.Text & "\" & BQM001_05.Text
    ml_observacion = BQM001_06.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM003" Then  'Colesterol total
    'BQM003
    ml_resultado = BQM003_01.Text & "\" & BQM003_02.Text & "\" & BQM003_03.Text & "\" & BQM003_04.Text
    ml_observacion = BQM003_05.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM004" Then 'Colesterol fraccionado
    'BQM004
    ml_resultado = BQM004_02.Text & "\" & BQM004_03.Text & "\" & BQM004_04.Text & "\" & BQM004_05.Text & "\" & BQM004_06.Text & "\" & BQM004_07.Text & "\" & BQM004_08.Text & "\" & BQM004_09.Text
    ml_observacion = BQM004_10.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM005" Then  'Triglicéridos
    'BQM005
    ml_resultado = BQM005_01.Text & "\" & BQM005_02.Text & "\" & BQM005_03.Text & "\" & BQM005_04.Text
    ml_observacion = BQM005_05.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM006" Then 'transaminasa GOT/GPT
    'BQM006
    ml_resultado = BQM006_01.Text & "\" & BQM006_02.Text & "\" & BQM006_03.Text & "\" & BQM006_04.Text
    ml_observacion = BQM006_05.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM007" Then 'transaminasa GOT/GPT
    'BQM007
    ml_resultado = BQM007_01.Text & "\" & BQM007_02.Text & "\" & BQM007_03.Text & "\" & BQM007_04.Text
    ml_observacion = BQM007_05.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM008" Then 'Bilirrubina Total
    'BQM008
    ml_resultado = BQM008_01.Text & "\" & BQM008_02.Text & "\" & BQM008_03.Text & "\" & BQM008_04.Text
    ml_observacion = BQM008_05.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM009" Then 'Billirubina total y fraccionada
    'BQM009
    ml_resultado = BQM009_02.Text & "\" & BQM009_03.Text & "\" & BQM009_04.Text & "\" & BQM009_05.Text & "\" & BQM009_06.Text & "\" & BQM009_07.Text & "\" & BQM009_08.Text & "\" & BQM009_09.Text & "\" & BQM009_10.Text & "\" & BQM009_11.Text & "\" & BQM009_12.Text & "\" & BQM009_13.Text
    ml_observacion = BQM009_14.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM010" Then 'Fosfatasa alcalina
    'BQM010
    ml_resultado = BQM010_01.Text & "\" & BQM010_02.Text & "\" & BQM010_03.Text & "\" & BQM010_04.Text
    ml_observacion = BQM010_05.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM011" Then 'Proteina total
    'BQM011
    ml_resultado = BQM011_01.Text & "\" & BQM011_02.Text & "\" & BQM011_03.Text & "\" & BQM011_04.Text
    ml_observacion = BQM011_05.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM012" Then 'Proteina total y fraccionada
    'BQM012
    ml_resultado = BQM012_02.Text & "\" & BQM012_03.Text & "\" & BQM012_04.Text & "\" & BQM012_05.Text & "\" & BQM012_06.Text & "\" & BQM012_07.Text & "\" & BQM012_08.Text & "\" & BQM012_09.Text & "\" & BQM012_10.Text & "\" & BQM012_11.Text & "\" & BQM012_12.Text & "\" & BQM012_13.Text
    ml_observacion = BQM012_14.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM014" Then 'Amilasa en orina
    'BQM014
    ml_resultado = BQM014_01.Text & "\" & BQM014_02.Text & "\" & BQM014_03.Text & "\" & BQM014_04.Text
    ml_observacion = BQM014_05.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM015" Then 'Úrea
    'BQM015
    ml_resultado = BQM015_01.Text & "\" & BQM015_02.Text & "\" & BQM015_03.Text & "\" & BQM015_04.Text
    ml_observacion = BQM015_05.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM016" Then  'Creatinina
    'BQM016
    ml_resultado = BQM016_01.Text & "\" & BQM016_02.Text & "\" & BQM016_03.Text & "\" & BQM016_04.Text & "\" & BQM016_05.Text
    ml_observacion = BQM016_06.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM017" Then 'Ácido Úrico
    'BQM017
    ml_resultado = BQM017_01.Text & "\" & BQM017_02.Text & "\" & BQM017_03.Text & "\" & BQM017_04.Text & "\" & BQM017_05.Text
    ml_observacion = BQM017_06.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM018" Then 'ADA
    'BQM018
    ml_resultado = BQM018_01.Text & "\" & BQM018_02.Text & "\" & BQM018_03.Text & "\" & BQM018_04.Text & "\" & BQM018_05.Text
    ml_observacion = BQM018_06.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM019" Then  'Calcio
    'BQM019
    ml_resultado = BQM019_01.Text & "\" & BQM019_02.Text & "\" & BQM019_03.Text & "\" & BQM019_04.Text & "\" & BQM019_05.Text
    ml_observacion = BQM019_06.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM021" Then  'Depuración de Creatinina
    'BQM021
    ml_resultado = BQM021_01.Text & "\" & BQM021_02.Text & "\" & BQM021_03.Text & "\" & BQM021_04.Text & "\" & BQM021_05.Text & "\" & BQM021_06.Text & "\" & BQM021_07.Text & "\" & BQM021_08.Text & "\" & BQM021_09.Text & "\" & BQM021_10.Text & "\" & BQM021_11.Text & "\" & BQM021_12.Text & "\" & BQM021_13.Text & "\" & BQM021_14.Text & "\" & BQM021_15.Text & "\" & BQM021_16.Text
    ml_observacion = BQM021_17.Text
  ElseIf ml_CodigoPruebaSeleccionada = "BQM030" Then 'transaminasa GOT/GPT
    'BQM030
    ml_resultado = BQM030_01.Text & "\" & BQM030_02.Text & "\" & BQM030_03.Text & "\" & BQM030_04.Text & "\" & BQM030_05.Text & "\" & BQM030_06.Text & "\" & BQM030_07.Text & "\" & BQM030_08.Text
    ml_observacion = BQM030_09.Text
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
  mo_ReglasLaboratorio.LabIngresaResultados idPrueba, idOrden, ml_resultado, ml_observacion, idUsuario, ml_nombreRealiza, ml_DetalleOrden, ml_idOrdenLab, "", "", 0, CDate(txtFresultado.Text), mo_lcNombrePc, mo_lnIdTablaLISTBARITEMS, Me.UcPacienteDatos1.DevuelveHistoriaApellidosYnombre, Me.Caption
End Sub

Private Sub cmdImprimir_Click()
  ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(idPrueba, idOrden)
  ml_observacion = mo_ReglasLaboratorio.LabRecuperaResultados_Obs(idPrueba, idOrden)
  Dim ldFechaResultado As Date
  ml_nombreRealiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(idPrueba, idOrden, ldFechaResultado)
  If ml_CodigoPruebaSeleccionada <> "" And Trim(ml_resultado) <> "" Then
    mo_ReglasLaboratorio.LabImprimeCabeceraResultados UcPacienteDatos1.idPaciente, nombrePaciente, _
                         UcPacienteDatos1.NroHistoriaClinica, ldFechaResultado, _
                         nombreMedico + mo_ReglasLaboratorio.DevuelveDatosParaImpresionResultadoLaboratorio(ml_idOrden)
    mo_ReglasLaboratorio.LabImprimeResultadosBQM ml_resultado, CStr(ml_CodigoPruebaSeleccionada), Me.Caption, ml_observacion, ml_nombreRealiza
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
 ' Else
 '   cmdGrabar.Enabled = False
 ' End If
 
 
  
  ml_resultado = ""
  ml_observacion = ""
  
  If ml_CodigoPruebaSeleccionada = "BQM001" Then  'Glucosa
    TopBoton BQM001
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       BQM001_01.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "BQM003" Then  'Colesterol total
    TopBoton BQM003
  ElseIf ml_CodigoPruebaSeleccionada = "BQM004" Then  'Colesterol fraccionado
    TopBoton BQM004
  ElseIf ml_CodigoPruebaSeleccionada = "BQM005" Then  'Triglicéridos
    TopBoton BQM005
  ElseIf ml_CodigoPruebaSeleccionada = "BQM006" Then 'transaminasa GOT/GPT
    TopBoton BQM006
  ElseIf ml_CodigoPruebaSeleccionada = "BQM007" Then 'transaminasa GOT/GPT
    TopBoton BQM007
  ElseIf ml_CodigoPruebaSeleccionada = "BQM008" Then  'Bilirrubina Total
    TopBoton BQM008
  ElseIf ml_CodigoPruebaSeleccionada = "BQM009" Then 'Billirubina total y fraccionada
    TopBoton BQM009
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       BQM009_15.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "BQM010" Then 'Fosfatasa alcalina
    TopBoton BQM010
  ElseIf ml_CodigoPruebaSeleccionada = "BQM011" Then 'Proteina total
    TopBoton BQM011
  ElseIf ml_CodigoPruebaSeleccionada = "BQM012" Then 'Proteina total y fraccionada
    TopBoton BQM012
  ElseIf ml_CodigoPruebaSeleccionada = "BQM014" Then  'Amilasa en orina
    TopBoton BQM014
  ElseIf ml_CodigoPruebaSeleccionada = "BQM015" Then 'Úrea
    TopBoton BQM015
  ElseIf ml_CodigoPruebaSeleccionada = "BQM016" Then  'Creatinina
    TopBoton BQM016
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       BQM016_03.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "BQM017" Then 'Ácido Úrico
    TopBoton BQM017
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       BQM017_03.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "BQM018" Then 'ADA
    TopBoton BQM018
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       BQM018_03.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "BQM019" Then 'Calcio
    TopBoton BQM019
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       BQM019_03.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "BQM021" Then 'Depuración de Creatinina
    TopBoton BQM021
  ElseIf ml_CodigoPruebaSeleccionada = "BQM030" Then 'transaminasa GOT/GPT
    TopBoton BQM030
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
  'Recupera información si es que ya esta grabado
  Dim ldFechaResultado As Date
  ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(idPrueba, idOrden)
  ml_nombreRealiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(idPrueba, idOrden, ldFechaResultado)
  If ml_resultado = "" Or Val(ml_nombreRealiza) = 0 Then Exit Sub
  Me.txtFresultado.Text = Format(IIf(ldFechaResultado = 0, Now, ldFechaResultado), sighentidades.DevuelveFechaSoloFormato_DMY_HM)
  cmbResponsable.ListIndex = Ubica_En_Combo(cmbResponsable, mo_ReglasLaboratorio.LabEmpleado(ml_nombreRealiza))
  'If cmbResponsable.Text <> "" Then cmbResponsable.Enabled = False
  ml_observacion = mo_ReglasLaboratorio.LabRecuperaResultados_Obs(idPrueba, idOrden)
  Dim Temp As String
  'Asigna la información recuperada en el formulario
  If ml_CodigoPruebaSeleccionada = "BQM001" Then 'Glucosa
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM001_01.ListIndex = Ubica_En_Combo(BQM001_01, Temp)
    BQM001_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM001_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM001_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM001_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM001_06.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM003" Then  'Colesterol total
    BQM003_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM003_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM003_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM003_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM003_05.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM004" Then 'Colesterol fraccionado
    BQM004_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM004_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM004_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM004_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM004_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM004_07.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM004_08.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM004_09.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM004_10.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM005" Then 'Triglicéridos
    BQM005_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM005_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM005_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM005_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM005_05.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM006" Then 'transaminasa GOT/GPT
    BQM006_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM006_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM006_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM006_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM006_05.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM007" Then 'transaminasa GOT/GPT
    BQM007_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM007_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM007_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM007_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM007_05.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM008" Then  'Bilirrubina Total
    BQM008_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM008_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM008_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM008_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM008_05.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM009" Then 'Billirubina total y fraccionada
    BQM009_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM009_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM009_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM009_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM009_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM009_07.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM009_08.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM009_09.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM009_10.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM009_11.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM009_12.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM009_13.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM009_14.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM010" Then 'Fosfatasa alcalina
    BQM010_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM010_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM010_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM010_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM010_05.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM011" Then 'Proteina total
    BQM011_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM011_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM011_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM011_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM011_05.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM012" Then 'proteina total y fraccionada
    BQM012_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM012_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM012_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM012_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM012_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM012_07.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM012_08.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM012_09.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM012_10.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM012_11.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM012_12.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM012_13.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM012_14.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM014" Then 'Amilasa en orina
    BQM014_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM014_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM014_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM014_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM014_05.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM015" Then 'Úrea
    BQM015_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM015_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM015_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM015_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM015_05.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM016" Then  'Creatinina
    BQM016_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM016_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM016_03.ListIndex = Ubica_En_Combo(BQM016_03, Temp)
    BQM016_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM016_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM016_06.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM017" Then 'Ácido Úrico
    BQM017_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM017_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM017_03.ListIndex = Ubica_En_Combo(BQM017_03, Temp)
    BQM017_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM017_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM017_06.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM018" Then 'ADA
    BQM018_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM018_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM018_03.ListIndex = Ubica_En_Combo(BQM018_03, Temp)
    BQM018_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM018_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM018_06.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM019" Then  'Calcio
    BQM019_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM019_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM019_03.ListIndex = Ubica_En_Combo(BQM019_03, Temp)
    BQM019_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM019_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM019_06.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM021" Then 'Depuración de Creatinina
    BQM021_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_07.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_08.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_09.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_10.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_11.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_12.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_13.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_14.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_15.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_16.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM021_17.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "BQM030" Then 'transaminasa GOT/GPT
    BQM030_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM030_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM030_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM030_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM030_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM030_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM030_07.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM030_08.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    BQM030_09.Text = ml_observacion
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
  End If
End Sub

Sub LimpiaVAloresDefault()
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       Exit Sub
    End If
    BQM008_03.Text = ""
    BQM008_04.Text = ""
    BQM001_04.Text = ""
    BQM001_05.Text = ""
    BQM003_03.Text = ""
    BQM003_04.Text = ""
    BQM004_04.Text = ""
    BQM004_08.Text = ""
    BQM004_05.Text = ""
    BQM004_09.Text = ""
    BQM005_03.Text = ""
    BQM005_04.Text = ""
    BQM006_03.Text = ""
    BQM006_04.Text = ""
    BQM007_03.Text = ""
    BQM007_04.Text = ""
    BQM009_04.Text = ""
    BQM009_08.Text = ""
    BQM009_12.Text = ""
    BQM009_05.Text = ""
    BQM009_09.Text = ""
    BQM009_13.Text = ""
    BQM010_03.Text = ""
    BQM010_04.Text = ""
    BQM011_03.Text = ""
    BQM011_04.Text = ""
    BQM012_04.Text = ""
    BQM012_08.Text = ""
    BQM012_12.Text = ""
    BQM012_05.Text = ""
    BQM012_09.Text = ""
    BQM012_13.Text = ""
    BQM014_03.Text = ""
    BQM014_04.Text = ""
    BQM015_03.Text = ""
    BQM015_04.Text = ""
    BQM016_04.Text = ""
    BQM016_05.Text = ""
    BQM017_04.Text = ""
    BQM017_05.Text = ""
    BQM018_05.Text = ""
    BQM019_04.Text = ""
    BQM019_05.Text = ""
    BQM021_03.Text = ""
    BQM021_07.Text = ""
    BQM021_11.Text = ""
    BQM021_04.Text = ""
    BQM021_08.Text = ""
    BQM021_12.Text = ""
    BQM030_03.Text = ""
    BQM030_07.Text = ""
    BQM030_04.Text = ""
    BQM030_08.Text = ""
    BQM031_07.Text = ""
    BQM031_11.Text = ""
    BQM031_15.Text = ""
    BQM031_04.Text = ""
    BQM031_08.Text = ""
    BQM031_12.Text = ""
    BQM031_16.Text = ""
End Sub


