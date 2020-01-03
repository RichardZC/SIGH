VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmMicrobiologia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MICROBIOLOGÍA"
   ClientHeight    =   10890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   ForeColor       =   &H00000000&
   Icon            =   "frmMicrobiologia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10890
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   387
      Top             =   1710
      Width           =   9255
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
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   388
         Top             =   210
         Width           =   3120
      End
      Begin MSMask.MaskEdBox txtFresultado 
         Height          =   315
         Left            =   6630
         TabIndex        =   389
         Top             =   210
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
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
         Left            =   930
         TabIndex        =   391
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label4 
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
         Left            =   5640
         TabIndex        =   390
         Top             =   255
         Width           =   945
      End
   End
   Begin SIGHLaboratorio.UcPacienteDatos1 UcPacienteDatos1 
      Height          =   1665
      Left            =   1080
      TabIndex        =   317
      Top             =   15
      Width           =   7185
      _ExtentX        =   12726
      _ExtentY        =   3201
   End
   Begin VB.Frame fraATB 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Fármacos"
      ForeColor       =   &H00000000&
      Height          =   3585
      Left            =   6840
      TabIndex        =   318
      Top             =   -1920
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CommandButton cmdCerrar1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ocultar"
         DisabledPicture =   "frmMicrobiologia.frx":0CCA
         DownPicture     =   "frmMicrobiologia.frx":118E
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
         Left            =   540
         Picture         =   "frmMicrobiologia.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   323
         Top             =   2930
         Width           =   1365
      End
      Begin VB.ComboBox cboGF 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   320
         Top             =   240
         Width           =   2415
      End
      Begin VB.ListBox lstF 
         Appearance      =   0  'Flat
         Height          =   1980
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   319
         Top             =   900
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         Height          =   3585
         Left            =   0
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Grupo Farmacológico:"
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
         Height          =   255
         Left            =   60
         TabIndex        =   322
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre Genérico:"
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
         Height          =   255
         Left            =   60
         TabIndex        =   321
         Top             =   660
         Width           =   2055
      End
   End
   Begin VB.Frame fraBoton 
      ForeColor       =   &H00000000&
      Height          =   870
      Left            =   60
      TabIndex        =   324
      Top             =   9990
      Width           =   9255
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmMicrobiologia.frx":1B66
         DownPicture     =   "frmMicrobiologia.frx":202A
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
         Left            =   4733
         Picture         =   "frmMicrobiologia.frx":2516
         Style           =   1  'Graphical
         TabIndex        =   327
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprime  (F3)"
         Height          =   615
         Left            =   90
         Picture         =   "frmMicrobiologia.frx":2A02
         Style           =   1  'Graphical
         TabIndex        =   326
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmMicrobiologia.frx":2EDB
         DownPicture     =   "frmMicrobiologia.frx":333B
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
         Left            =   3293
         Picture         =   "frmMicrobiologia.frx":37B0
         Style           =   1  'Graphical
         TabIndex        =   325
         Top             =   180
         Width           =   1365
      End
   End
   Begin VB.Frame MIC031 
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
      Height          =   7620
      Left            =   60
      TabIndex        =   305
      Top             =   2340
      Visible         =   0   'False
      Width           =   9255
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Vancomicina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   28
         Left            =   6480
         TabIndex        =   351
         Top             =   6840
         Width           =   1260
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Tetraciclina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   27
         Left            =   4560
         TabIndex        =   350
         Top             =   6840
         Width           =   1455
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Sulfametosaxol"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   26
         Left            =   3000
         TabIndex        =   349
         Top             =   6840
         Width           =   1455
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Norfloxacina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   25
         Left            =   1560
         TabIndex        =   348
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Cirugía contaminada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   4200
         TabIndex        =   136
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Traqueostomía"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   4200
         TabIndex        =   144
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Diabetes Mellitus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   3600
         TabIndex        =   139
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Neoplasia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   134
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Cirugía previa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   2160
         TabIndex        =   141
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Cat. venoso central"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   2160
         TabIndex        =   143
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Desnutrición"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   138
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Insuficiencia renal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   133
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Coma"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   132
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Cirrosis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   137
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Ventilación mecánica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   146
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Enf. Pulmonar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   5520
         TabIndex        =   135
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Infección por VIH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   5520
         TabIndex        =   140
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Sonda nasogástrica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   6360
         TabIndex        =   145
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox MIC031_09 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5880
         TabIndex        =   131
         Top             =   1440
         Width           =   3210
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "S. Vaginal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   22
         Left            =   3600
         TabIndex        =   121
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "L. Pericárdico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   7320
         TabIndex        =   118
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "S. Ótica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   111
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "S. Uretral"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   21
         Left            =   1680
         TabIndex        =   116
         Top             =   1440
         Width           =   1380
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "Otro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   23
         Left            =   5160
         TabIndex        =   130
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "L. Sinovial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   1680
         TabIndex        =   120
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "L. Pleural"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   119
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "Esputo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   5160
         TabIndex        =   110
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "L.Biliar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   115
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "S. Cerviz"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   125
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "S. Nasal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   19
         Left            =   7320
         TabIndex        =   129
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "L. Cefalorraquídeo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   5160
         TabIndex        =   117
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "S. Faríngea"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   18
         Left            =   5160
         TabIndex        =   128
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "Sangre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   7320
         TabIndex        =   124
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "L. Ascítico-Peritoneal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   114
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "Esperma"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   109
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "S. de Herida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   17
         Left            =   3600
         TabIndex        =   127
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "Raspado de piel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   5160
         TabIndex        =   123
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "Hisopado rectal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         TabIndex        =   113
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "Cat. venoso central"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   108
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "Aspirado gástrico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   107
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "Heces"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   112
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "Orina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   3600
         TabIndex        =   122
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox MIC031_10 
         Caption         =   "S. Conjuntiva"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   1680
         TabIndex        =   126
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Nitropurantoina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   177
         Top             =   6840
         Width           =   1455
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Meropenem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   23
         Left            =   7920
         TabIndex        =   175
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Lincomicina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   22
         Left            =   6480
         TabIndex        =   169
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Impenem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   21
         Left            =   4560
         TabIndex        =   163
         Top             =   6600
         Width           =   1260
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Gentamicina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   20
         Left            =   3000
         TabIndex        =   157
         Top             =   6600
         Width           =   1335
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Estreptomicina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   19
         Left            =   1560
         TabIndex        =   176
         Top             =   6600
         Width           =   1575
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Eritromicina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   174
         Top             =   6600
         Width           =   1575
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Dicloxacilina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   17
         Left            =   7920
         TabIndex        =   168
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Cotrimoxazol"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   6480
         TabIndex        =   162
         Top             =   6360
         Width           =   1335
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Cloranfenicol"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   4560
         TabIndex        =   156
         Top             =   6360
         Width           =   1575
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Ampicilina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   6480
         TabIndex        =   158
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Amox/Ac.Clavulanico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   171
         Top             =   5880
         Width           =   1815
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Amikacina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   159
         Top             =   5880
         Width           =   1455
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Ac. Nalidixico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   153
         Top             =   5880
         Width           =   1575
      End
      Begin VB.TextBox MIC031_06 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2955
         TabIndex        =   151
         Top             =   4980
         Width           =   6210
      End
      Begin VB.TextBox MIC031_07 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2955
         TabIndex        =   152
         Top             =   5310
         Width           =   6210
      End
      Begin VB.TextBox MIC031_05 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2955
         MultiLine       =   -1  'True
         TabIndex        =   150
         Top             =   4570
         Width           =   6210
      End
      Begin VB.TextBox MIC031_08 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1035
         TabIndex        =   178
         Top             =   7200
         Width           =   8130
      End
      Begin VB.TextBox MIC031_03 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2955
         MultiLine       =   -1  'True
         TabIndex        =   148
         Top             =   3910
         Width           =   6210
      End
      Begin VB.TextBox MIC031_01 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2955
         MultiLine       =   -1  'True
         TabIndex        =   147
         Top             =   3580
         Width           =   6210
      End
      Begin VB.TextBox MIC031_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2955
         MultiLine       =   -1  'True
         TabIndex        =   149
         Top             =   4240
         Width           =   6210
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Aztreonam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   7920
         TabIndex        =   154
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Cefadroxil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   160
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Cefalotina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   166
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Cefaclor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   172
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Cefepime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   4560
         TabIndex        =   164
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Cefotaxima"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   6480
         TabIndex        =   155
         Top             =   6120
         Width           =   1335
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Ceftazidima"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   7920
         TabIndex        =   161
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Ceftriaxona"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   167
         Top             =   6360
         Width           =   1335
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Ciprofloxacina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   1560
         TabIndex        =   173
         Top             =   6360
         Width           =   1335
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Clindamicina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   3000
         TabIndex        =   170
         Top             =   6360
         Width           =   1575
      End
      Begin VB.CheckBox MIC031_02 
         Caption         =   "Amoxicilina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   165
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CheckBox MIC031_11 
         Caption         =   "Sonda Vesical"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   142
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Line Line2 
         X1              =   1920
         X2              =   2160
         Y1              =   5280
         Y2              =   5400
      End
      Begin VB.Line Line1 
         X1              =   1920
         X2              =   2160
         Y1              =   5280
         Y2              =   5160
      End
      Begin VB.Label MIC031_00 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Informe de Resultado de Examen Microbiológico"
         BeginProperty Font 
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
         Left            =   2895
         TabIndex        =   347
         Top             =   3360
         Width           =   3450
      End
      Begin VB.Label MIC031_00 
         AutoSize        =   -1  'True
         Caption         =   "Factores en infecciones Intrahospitalarias Extrínsecos"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   346
         Top             =   2520
         Width           =   3870
      End
      Begin VB.Label MIC031_00 
         AutoSize        =   -1  'True
         Caption         =   "Factores en infecciones Intrahospitalarias Intrínsecos"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   316
         Top             =   1800
         Width           =   3840
      End
      Begin VB.Label MIC031_00 
         Caption         =   "Muestra"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   315
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label MIC031_00 
         Caption         =   "Sensibilidad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   314
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label MIC031_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Germen 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   313
         Top             =   5010
         Width           =   735
      End
      Begin VB.Label MIC031_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Germen 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   312
         Top             =   5340
         Width           =   855
      End
      Begin VB.Label MIC031_00 
         Caption         =   "Aislamiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   1020
         TabIndex        =   311
         Top             =   5160
         Width           =   855
      End
      Begin VB.Label MIC031_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Prueba de Antibióticos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   310
         Top             =   4600
         Width           =   2775
      End
      Begin VB.Label MIC031_00 
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   60
         TabIndex        =   309
         Top             =   7230
         Width           =   975
      End
      Begin VB.Label MIC031_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Gram / K OH / Otros"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   308
         Top             =   3940
         Width           =   2775
      End
      Begin VB.Label MIC031_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Sedimento Urinario / Examen Directo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   307
         Top             =   3610
         Width           =   2775
      End
      Begin VB.Label MIC031_00 
         Alignment       =   1  'Right Justify
         Caption         =   "Recuento de Colonias (UFC x ml)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   306
         Top             =   4270
         Width           =   2775
      End
   End
   Begin VB.Frame MIC030 
      Caption         =   "Parasitológico Seriado"
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
      Height          =   2295
      Left            =   1110
      TabIndex        =   300
      Top             =   2370
      Visible         =   0   'False
      Width           =   7220
      Begin VB.ComboBox MIC030_03 
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3C25
         Left            =   1150
         List            =   "frmMicrobiologia.frx":3C2F
         TabIndex        =   103
         Text            =   "MIC030_03"
         Top             =   900
         Width           =   1215
      End
      Begin VB.ComboBox MIC030_01 
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3C47
         Left            =   1150
         List            =   "frmMicrobiologia.frx":3C51
         TabIndex        =   101
         Text            =   "MIC030_01"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox MIC030_05 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1150
         TabIndex        =   105
         Text            =   "Las muestras se procesaron por observación directa y sedimentación espontánea."
         Top             =   1575
         Width           =   5970
      End
      Begin VB.TextBox MIC030_02 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1150
         TabIndex        =   102
         Text            =   "No se observan quistes ni trofozoitos"
         Top             =   585
         Width           =   5970
      End
      Begin VB.TextBox MIC030_04 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1150
         TabIndex        =   104
         Text            =   "No se observan huevos ni larvas."
         Top             =   1260
         Width           =   5970
      End
      Begin VB.TextBox MIC030_06 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1150
         TabIndex        =   106
         Top             =   1920
         Width           =   5970
      End
      Begin VB.Label MIC030_00 
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
         Left            =   60
         TabIndex        =   304
         Top             =   1605
         Width           =   1335
      End
      Begin VB.Label MIC030_00 
         Caption         =   "Protozoarios"
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
         TabIndex        =   303
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label MIC030_00 
         Caption         =   "Helmintos"
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
         Left            =   60
         TabIndex        =   302
         Top             =   930
         Width           =   1095
      End
      Begin VB.Label MIC030_00 
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
         Index           =   3
         Left            =   60
         TabIndex        =   301
         Top             =   1950
         Width           =   1335
      End
   End
   Begin VB.Frame MIC003 
      Caption         =   "Examen Completo Secreción Vaginal"
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
      Height          =   4395
      Left            =   1110
      TabIndex        =   276
      Top             =   2340
      Visible         =   0   'False
      Width           =   6360
      Begin VB.TextBox MIC003_16 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         Top             =   3615
         Width           =   4815
      End
      Begin VB.TextBox MIC003_15 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5775
         TabIndex        =   26
         Text            =   "1+"
         Top             =   3285
         Width           =   495
      End
      Begin VB.TextBox MIC003_04 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4710
         TabIndex        =   15
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox MIC003_03 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1350
         TabIndex        =   14
         Text            =   "Blanquesino"
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox MIC003_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1350
         TabIndex        =   12
         Text            =   "Grumoso"
         Top             =   465
         Width           =   1455
      End
      Begin VB.ComboBox MIC003_02 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3C69
         Left            =   4710
         List            =   "frmMicrobiologia.frx":3C76
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   440
         Width           =   1455
      End
      Begin VB.ComboBox MIC003_06 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3C93
         Left            =   5130
         List            =   "frmMicrobiologia.frx":3C9D
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1410
         Width           =   1095
      End
      Begin VB.ComboBox MIC003_07 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3CB5
         Left            =   2160
         List            =   "frmMicrobiologia.frx":3CBF
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1725
         Width           =   1095
      End
      Begin VB.ComboBox MIC003_05 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3CD7
         Left            =   2160
         List            =   "frmMicrobiologia.frx":3CE1
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1410
         Width           =   1095
      End
      Begin VB.ComboBox MIC003_08 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3CF9
         Left            =   5130
         List            =   "frmMicrobiologia.frx":3D03
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1725
         Width           =   1095
      End
      Begin VB.ComboBox MIC003_09 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3D1B
         Left            =   2160
         List            =   "frmMicrobiologia.frx":3D25
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2055
         Width           =   1095
      End
      Begin VB.TextBox MIC003_17 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1395
         TabIndex        =   28
         Top             =   4050
         Width           =   4830
      End
      Begin VB.ComboBox MIC003_10 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3D3D
         Left            =   4320
         List            =   "frmMicrobiologia.frx":3D4D
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2470
         Width           =   1455
      End
      Begin VB.ComboBox MIC003_11 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3D75
         Left            =   1995
         List            =   "frmMicrobiologia.frx":3D85
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2940
         Width           =   1215
      End
      Begin VB.ComboBox MIC003_12 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3DAD
         Left            =   4275
         List            =   "frmMicrobiologia.frx":3DBD
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2940
         Width           =   1215
      End
      Begin VB.ComboBox MIC003_13 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3DE5
         Left            =   1995
         List            =   "frmMicrobiologia.frx":3DEF
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3270
         Width           =   1215
      End
      Begin VB.ComboBox MIC003_14 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3E08
         Left            =   4275
         List            =   "frmMicrobiologia.frx":3E12
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3270
         Width           =   1455
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Otros"
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
         Index           =   13
         Left            =   60
         TabIndex        =   296
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Flora Bacteriana"
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
         Index           =   11
         Left            =   60
         TabIndex        =   295
         Top             =   2535
         Width           =   1410
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Test de K (OH)"
         BeginProperty Font 
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
         Left            =   3300
         TabIndex        =   294
         Top             =   810
         Width           =   1065
      End
      Begin VB.Label MIC003_00 
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
         TabIndex        =   293
         Top             =   225
         Width           =   1260
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Otros"
         BeginProperty Font 
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
         Left            =   735
         TabIndex        =   292
         Top             =   3645
         Width           =   405
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Bacterias"
         BeginProperty Font 
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
         Left            =   3435
         TabIndex        =   291
         Top             =   3315
         Width           =   660
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Hematíes"
         BeginProperty Font 
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
         Left            =   3435
         TabIndex        =   290
         Top             =   3000
         Width           =   660
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Leucocitos"
         BeginProperty Font 
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
         Left            =   975
         TabIndex        =   289
         Top             =   3000
         Width           =   750
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "pH"
         BeginProperty Font 
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
         Index           =   2
         Left            =   3300
         TabIndex        =   288
         Top             =   495
         Width           =   195
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Color"
         BeginProperty Font 
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
         Left            =   135
         TabIndex        =   287
         Top             =   810
         Width           =   375
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Aspecto"
         BeginProperty Font 
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
         Index           =   1
         Left            =   135
         TabIndex        =   286
         Top             =   495
         Width           =   585
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Mobiluncus sp"
         BeginProperty Font 
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
         Left            =   3780
         TabIndex        =   285
         Top             =   1785
         Width           =   990
      End
      Begin VB.Label MIC003_00 
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
         Index           =   5
         Left            =   60
         TabIndex        =   284
         Top             =   1200
         Width           =   1860
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Candida sp"
         BeginProperty Font 
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
         Left            =   3780
         TabIndex        =   283
         Top             =   1470
         Width           =   795
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Neisseria gonorhoeae"
         BeginProperty Font 
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
         Left            =   135
         TabIndex        =   282
         Top             =   1785
         Width           =   1560
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Trichomonas vaginalis"
         BeginProperty Font 
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
         Left            =   135
         TabIndex        =   281
         Top             =   1470
         Width           =   1560
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Gardnerella vaginalis"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   280
         Top             =   2085
         Width           =   1485
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Cél. Guía (Clue Cells)"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   279
         Top             =   3315
         Width           =   1500
      End
      Begin VB.Label MIC003_00 
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
         TabIndex        =   278
         Top             =   4080
         Width           =   1275
      End
      Begin VB.Label MIC003_00 
         AutoSize        =   -1  'True
         Caption         =   "Lactobacillus acidophilus"
         BeginProperty Font 
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
         Left            =   2040
         TabIndex        =   277
         Top             =   2520
         Width           =   1725
      End
   End
   Begin VB.Frame MIC008 
      Caption         =   "Líquido Céfalo-Raquídeo"
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
      Height          =   3045
      Left            =   1080
      TabIndex        =   202
      Top             =   2370
      Visible         =   0   'False
      Width           =   6360
      Begin VB.TextBox MIC008_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   5640
         TabIndex        =   98
         Top             =   1605
         Width           =   495
      End
      Begin VB.TextBox MIC008_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   2550
         TabIndex        =   94
         Top             =   930
         Width           =   615
      End
      Begin VB.TextBox MIC008_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   240
         TabIndex        =   99
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox MIC008_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   870
         TabIndex        =   91
         Text            =   "Claro"
         Top             =   470
         Width           =   1095
      End
      Begin VB.TextBox MIC008_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   5160
         TabIndex        =   93
         Text            =   "Ausente"
         Top             =   470
         Width           =   1095
      End
      Begin VB.TextBox MIC008_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   1230
         TabIndex        =   96
         Top             =   1605
         Width           =   615
      End
      Begin VB.TextBox MIC008_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   3765
         TabIndex        =   97
         Top             =   1605
         Width           =   615
      End
      Begin VB.TextBox MIC008_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   92
         Text            =   "Amarillento"
         Top             =   470
         Width           =   1095
      End
      Begin VB.TextBox MIC008_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   5150
         TabIndex        =   95
         Top             =   930
         Width           =   615
      End
      Begin VB.TextBox MIC008_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   1485
         TabIndex        =   100
         Top             =   2640
         Width           =   4830
      End
      Begin VB.Label MIC008_00 
         AutoSize        =   -1  'True
         Caption         =   "PMN"
         BeginProperty Font 
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
         Left            =   5205
         TabIndex        =   222
         Top             =   1635
         Width           =   315
      End
      Begin VB.Label MIC008_00 
         AutoSize        =   -1  'True
         Caption         =   "Color"
         BeginProperty Font 
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
         Index           =   2
         Left            =   2400
         TabIndex        =   221
         Top             =   495
         Width           =   375
      End
      Begin VB.Label MIC008_00 
         AutoSize        =   -1  'True
         Caption         =   "U / l"
         BeginProperty Font 
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
         Left            =   960
         TabIndex        =   220
         Top             =   2310
         Width           =   285
      End
      Begin VB.Label MIC008_00 
         AutoSize        =   -1  'True
         Caption         =   "Adenosin Deaminasa (ADA)"
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
         Index           =   16
         Left            =   60
         TabIndex        =   219
         Top             =   2040
         Width           =   2340
      End
      Begin VB.Label MIC008_00 
         AutoSize        =   -1  'True
         Caption         =   "Glucosa"
         BeginProperty Font 
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
         Left            =   1815
         TabIndex        =   218
         Top             =   960
         Width           =   555
      End
      Begin VB.Label MIC008_00 
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
         TabIndex        =   217
         Top             =   960
         Width           =   1440
      End
      Begin VB.Label MIC008_00 
         AutoSize        =   -1  'True
         Caption         =   "Aspecto"
         BeginProperty Font 
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
         Index           =   1
         Left            =   135
         TabIndex        =   216
         Top             =   495
         Width           =   585
      End
      Begin VB.Label MIC008_00 
         AutoSize        =   -1  'True
         Caption         =   "Coagulo"
         BeginProperty Font 
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
         Left            =   4335
         TabIndex        =   215
         Top             =   495
         Width           =   585
      End
      Begin VB.Label MIC008_00 
         AutoSize        =   -1  'True
         Caption         =   "Linfocitos"
         BeginProperty Font 
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
         Left            =   2880
         TabIndex        =   214
         Top             =   1635
         Width           =   675
      End
      Begin VB.Label MIC008_00 
         AutoSize        =   -1  'True
         Caption         =   "Leucocitos"
         BeginProperty Font 
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
         Left            =   225
         TabIndex        =   213
         Top             =   1635
         Width           =   750
      End
      Begin VB.Label MIC008_00 
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
         TabIndex        =   212
         Top             =   225
         Width           =   1260
      End
      Begin VB.Label MIC008_00 
         AutoSize        =   -1  'True
         Caption         =   "Examen Citológico"
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
         TabIndex        =   211
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label MIC008_00 
         AutoSize        =   -1  'True
         Caption         =   "Proteinas"
         BeginProperty Font 
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
         Left            =   4200
         TabIndex        =   210
         Top             =   960
         Width           =   675
      End
      Begin VB.Label MIC008_00 
         AutoSize        =   -1  'True
         Caption         =   "g / dl"
         BeginProperty Font 
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
         Left            =   5775
         TabIndex        =   209
         Top             =   960
         Width           =   360
      End
      Begin VB.Label MIC008_00 
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
         Index           =   6
         Left            =   3240
         TabIndex        =   208
         Top             =   960
         Width           =   480
      End
      Begin VB.Label MIC008_00 
         AutoSize        =   -1  'True
         Caption         =   "(Valor de Referencia: 0 - 6 U / l)"
         BeginProperty Font 
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
         Left            =   1800
         TabIndex        =   207
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label MIC008_00 
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
         Index           =   13
         Left            =   4440
         TabIndex        =   206
         Top             =   1635
         Width           =   165
      End
      Begin VB.Label MIC008_00 
         AutoSize        =   -1  'True
         Caption         =   "/ mm3"
         BeginProperty Font 
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
         Left            =   1905
         TabIndex        =   205
         Top             =   1635
         Width           =   435
      End
      Begin VB.Label MIC008_00 
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
         Index           =   15
         Left            =   6150
         TabIndex        =   204
         Top             =   1635
         Width           =   165
      End
      Begin VB.Label MIC008_00 
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
         Index           =   19
         Left            =   120
         TabIndex        =   203
         Top             =   2670
         Width           =   1335
      End
   End
   Begin VB.Frame MIC007 
      Caption         =   "Líquido Sinovial"
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
      Height          =   3765
      Left            =   1110
      TabIndex        =   223
      Top             =   2340
      Visible         =   0   'False
      Width           =   6360
      Begin VB.TextBox MIC007_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   150
         TabIndex        =   85
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox MIC007_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   4250
         TabIndex        =   84
         Top             =   1605
         Width           =   615
      End
      Begin VB.TextBox MIC007_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   1830
         TabIndex        =   83
         Top             =   1605
         Width           =   615
      End
      Begin VB.TextBox MIC007_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   4440
         TabIndex        =   80
         Top             =   470
         Width           =   1095
      End
      Begin VB.TextBox MIC007_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   870
         TabIndex        =   79
         Top             =   470
         Width           =   1095
      End
      Begin VB.ComboBox MIC007_02 
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         ItemData        =   "frmMicrobiologia.frx":3E34
         Left            =   1920
         List            =   "frmMicrobiologia.frx":3E3E
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   2850
         Width           =   1215
      End
      Begin VB.TextBox MIC007_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   4530
         TabIndex        =   89
         Text            =   "En proceso"
         Top             =   2850
         Width           =   1455
      End
      Begin VB.TextBox MIC007_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   4200
         TabIndex        =   82
         Top             =   930
         Width           =   1095
      End
      Begin VB.TextBox MIC007_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   4710
         TabIndex        =   87
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox MIC007_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   2775
         TabIndex        =   86
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox MIC007_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   1485
         TabIndex        =   90
         Top             =   3360
         Width           =   4830
      End
      Begin VB.ComboBox MIC007_02 
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         ItemData        =   "frmMicrobiologia.frx":3E56
         Left            =   2880
         List            =   "frmMicrobiologia.frx":3E60
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "Análisis Citológico"
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
         Index           =   7
         Left            =   60
         TabIndex        =   242
         Top             =   1350
         Width           =   1560
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "Análisis Bioquímico"
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
         TabIndex        =   241
         Top             =   225
         Width           =   1665
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "Recuento Diferencial:"
         BeginProperty Font 
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
         Left            =   135
         TabIndex        =   240
         Top             =   1950
         Width           =   1545
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "leucocitos / mm3"
         BeginProperty Font 
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
         Left            =   2505
         TabIndex        =   239
         Top             =   1635
         Width           =   1185
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "hematíes / mm3"
         BeginProperty Font 
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
         Left            =   4920
         TabIndex        =   238
         Top             =   1635
         Width           =   1125
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "Recuento Celular:"
         BeginProperty Font 
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
         Left            =   135
         TabIndex        =   237
         Top             =   1635
         Width           =   1290
      End
      Begin VB.Label MIC007_00 
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
         Left            =   5640
         TabIndex        =   236
         Top             =   495
         Width           =   480
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "Proteinas"
         BeginProperty Font 
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
         Left            =   3615
         TabIndex        =   235
         Top             =   495
         Width           =   675
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "Glucosa"
         BeginProperty Font 
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
         Index           =   1
         Left            =   135
         TabIndex        =   234
         Top             =   495
         Width           =   555
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "Análisi Serológico"
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
         Index           =   5
         Left            =   60
         TabIndex        =   233
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "Látex"
         BeginProperty Font 
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
         Left            =   2295
         TabIndex        =   232
         Top             =   960
         Width           =   405
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "Análisis Bacteriológico"
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
         Index           =   15
         Left            =   120
         TabIndex        =   231
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "Cultivo"
         BeginProperty Font 
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
         Left            =   3795
         TabIndex        =   230
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "Coloración de Gram"
         BeginProperty Font 
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
         Left            =   195
         TabIndex        =   229
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Label MIC007_00 
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
         Index           =   2
         Left            =   2040
         TabIndex        =   228
         Top             =   495
         Width           =   480
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "% Linfocitos"
         BeginProperty Font 
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
         Left            =   3285
         TabIndex        =   227
         Top             =   2190
         Width           =   885
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "% Monocitos"
         BeginProperty Font 
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
         Left            =   5220
         TabIndex        =   226
         Top             =   2190
         Width           =   930
      End
      Begin VB.Label MIC007_00 
         AutoSize        =   -1  'True
         Caption         =   "% Polimorfonucleares"
         BeginProperty Font 
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
         Left            =   660
         TabIndex        =   225
         Top             =   2190
         Width           =   1560
      End
      Begin VB.Label MIC007_00 
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
         Index           =   18
         Left            =   120
         TabIndex        =   224
         Top             =   3390
         Width           =   1335
      End
   End
   Begin VB.Frame MIC005 
      Caption         =   "Raspado de Piel: Examen completo de Lesión Dérmica"
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
      Height          =   3675
      Left            =   1080
      TabIndex        =   187
      Top             =   2370
      Visible         =   0   'False
      Width           =   6360
      Begin VB.TextBox MIC005_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2910
         TabIndex        =   39
         Text            =   "Manchas"
         Top             =   220
         Width           =   3375
      End
      Begin VB.TextBox MIC005_02 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2910
         TabIndex        =   40
         Text            =   "A nivel de "
         Top             =   540
         Width           =   3375
      End
      Begin VB.TextBox MIC005_03 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2910
         TabIndex        =   41
         Top             =   860
         Width           =   3375
      End
      Begin VB.TextBox MIC005_04 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1470
         TabIndex        =   42
         Text            =   "Ausencia de levaduras e hifas de hongos"
         Top             =   1320
         Width           =   4815
      End
      Begin VB.TextBox MIC005_07 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1950
         TabIndex        =   45
         Top             =   2250
         Width           =   735
      End
      Begin VB.TextBox MIC005_08 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4950
         TabIndex        =   46
         Top             =   2250
         Width           =   735
      End
      Begin VB.TextBox MIC005_09 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1950
         TabIndex        =   47
         Top             =   2570
         Width           =   4335
      End
      Begin VB.TextBox MIC005_10 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1950
         TabIndex        =   48
         Top             =   2890
         Width           =   4335
      End
      Begin VB.TextBox MIC005_05 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1470
         TabIndex        =   43
         Text            =   "Ausencia de "
         Top             =   1780
         Width           =   1455
      End
      Begin VB.TextBox MIC005_11 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1455
         TabIndex        =   49
         Top             =   3315
         Width           =   4830
      End
      Begin VB.TextBox MIC005_06 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3000
         TabIndex        =   44
         Text            =   "Demodex follicullorum y Sarcoptes scabei"
         Top             =   1780
         Width           =   3255
      End
      Begin VB.Label MIC005_00 
         AutoSize        =   -1  'True
         Caption         =   "Lesión"
         BeginProperty Font 
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
         Index           =   1
         Left            =   1695
         TabIndex        =   201
         Top             =   255
         Width           =   450
      End
      Begin VB.Label MIC005_00 
         AutoSize        =   -1  'True
         Caption         =   "Localización"
         BeginProperty Font 
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
         Index           =   2
         Left            =   1695
         TabIndex        =   200
         Top             =   570
         Width           =   840
      End
      Begin VB.Label MIC005_00 
         AutoSize        =   -1  'True
         Caption         =   "Color"
         BeginProperty Font 
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
         Left            =   1695
         TabIndex        =   199
         Top             =   885
         Width           =   375
      End
      Begin VB.Label MIC005_00 
         AutoSize        =   -1  'True
         Caption         =   "Leucocitos"
         BeginProperty Font 
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
         Left            =   855
         TabIndex        =   198
         Top             =   2280
         Width           =   750
      End
      Begin VB.Label MIC005_00 
         AutoSize        =   -1  'True
         Caption         =   "Hematíes"
         BeginProperty Font 
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
         Left            =   3855
         TabIndex        =   197
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label MIC005_00 
         AutoSize        =   -1  'True
         Caption         =   "Bacterias"
         BeginProperty Font 
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
         Left            =   855
         TabIndex        =   196
         Top             =   2595
         Width           =   660
      End
      Begin VB.Label MIC005_00 
         AutoSize        =   -1  'True
         Caption         =   "Otros"
         BeginProperty Font 
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
         Left            =   855
         TabIndex        =   195
         Top             =   2925
         Width           =   405
      End
      Begin VB.Label MIC005_00 
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
         TabIndex        =   194
         Top             =   225
         Width           =   1260
      End
      Begin VB.Label MIC005_00 
         AutoSize        =   -1  'True
         Caption         =   "Test de K (OH)"
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
         TabIndex        =   193
         Top             =   1320
         Width           =   1290
      End
      Begin VB.Label MIC005_00 
         AutoSize        =   -1  'True
         Caption         =   "Ectoparásitos"
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
         Index           =   5
         Left            =   60
         TabIndex        =   192
         Top             =   1815
         Width           =   1170
      End
      Begin VB.Label MIC005_00 
         AutoSize        =   -1  'True
         Caption         =   "Otros"
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
         Index           =   6
         Left            =   60
         TabIndex        =   191
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label MIC005_00 
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
         Index           =   13
         Left            =   60
         TabIndex        =   190
         Top             =   3345
         Width           =   1275
      End
      Begin VB.Label MIC005_00 
         AutoSize        =   -1  'True
         Caption         =   "x c"
         BeginProperty Font 
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
         Left            =   2715
         TabIndex        =   189
         Top             =   2280
         Width           =   210
      End
      Begin VB.Label MIC005_00 
         AutoSize        =   -1  'True
         Caption         =   "x c"
         BeginProperty Font 
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
         Left            =   5715
         TabIndex        =   188
         Top             =   2280
         Width           =   210
      End
   End
   Begin VB.Frame MIC004 
      Caption         =   "Examen Completo Secreción Uretral"
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
      Height          =   2685
      Left            =   1080
      TabIndex        =   243
      Top             =   2340
      Visible         =   0   'False
      Width           =   6360
      Begin VB.ComboBox MIC004_04 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3E78
         Left            =   4320
         List            =   "frmMicrobiologia.frx":3E82
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   900
         Width           =   1095
      End
      Begin VB.ComboBox MIC004_03 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3E9A
         Left            =   5070
         List            =   "frmMicrobiologia.frx":3EA7
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   450
         Width           =   1215
      End
      Begin VB.TextBox MIC004_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   870
         TabIndex        =   29
         Text            =   "Mucoide"
         Top             =   465
         Width           =   1455
      End
      Begin VB.TextBox MIC004_02 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3030
         TabIndex        =   30
         Text            =   "Cremoso"
         Top             =   465
         Width           =   1455
      End
      Begin VB.TextBox MIC004_08 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1110
         TabIndex        =   36
         Text            =   "0 - 1"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox MIC004_07 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5000
         TabIndex        =   35
         Text            =   "1+"
         Top             =   1605
         Width           =   495
      End
      Begin VB.TextBox MIC004_09 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3510
         TabIndex        =   37
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox MIC004_10 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1395
         TabIndex        =   38
         Top             =   2320
         Width           =   4830
      End
      Begin VB.ComboBox MIC004_06 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3EC4
         Left            =   3510
         List            =   "frmMicrobiologia.frx":3ECE
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1580
         Width           =   1455
      End
      Begin VB.ComboBox MIC004_05 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3EF0
         Left            =   1110
         List            =   "frmMicrobiologia.frx":3F00
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1580
         Width           =   1215
      End
      Begin VB.Label MIC004_00 
         AutoSize        =   -1  'True
         Caption         =   "Neisseria Gonorhoeae"
         BeginProperty Font 
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
         Left            =   2295
         TabIndex        =   256
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label MIC004_00 
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
         Index           =   4
         Left            =   60
         TabIndex        =   255
         Top             =   960
         Width           =   1860
      End
      Begin VB.Label MIC004_00 
         AutoSize        =   -1  'True
         Caption         =   "Aspecto"
         BeginProperty Font 
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
         Index           =   1
         Left            =   135
         TabIndex        =   254
         Top             =   495
         Width           =   585
      End
      Begin VB.Label MIC004_00 
         AutoSize        =   -1  'True
         Caption         =   "Color"
         BeginProperty Font 
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
         Index           =   2
         Left            =   2535
         TabIndex        =   253
         Top             =   495
         Width           =   375
      End
      Begin VB.Label MIC004_00 
         AutoSize        =   -1  'True
         Caption         =   "pH"
         BeginProperty Font 
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
         Left            =   4740
         TabIndex        =   252
         Top             =   495
         Width           =   195
      End
      Begin VB.Label MIC004_00 
         AutoSize        =   -1  'True
         Caption         =   "Leucocitos"
         BeginProperty Font 
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
         Left            =   135
         TabIndex        =   251
         Top             =   1635
         Width           =   750
      End
      Begin VB.Label MIC004_00 
         AutoSize        =   -1  'True
         Caption         =   "Hematíes"
         BeginProperty Font 
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
         Left            =   135
         TabIndex        =   250
         Top             =   1950
         Width           =   660
      End
      Begin VB.Label MIC004_00 
         AutoSize        =   -1  'True
         Caption         =   "Bacterias"
         BeginProperty Font 
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
         Left            =   2655
         TabIndex        =   249
         Top             =   1635
         Width           =   660
      End
      Begin VB.Label MIC004_00 
         AutoSize        =   -1  'True
         Caption         =   "Otros"
         BeginProperty Font 
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
         Left            =   2655
         TabIndex        =   248
         Top             =   1950
         Width           =   405
      End
      Begin VB.Label MIC004_00 
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
         TabIndex        =   247
         Top             =   225
         Width           =   1260
      End
      Begin VB.Label MIC004_00 
         AutoSize        =   -1  'True
         Caption         =   "Otros"
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
         Index           =   6
         Left            =   60
         TabIndex        =   246
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label MIC004_00 
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
         Index           =   12
         Left            =   60
         TabIndex        =   245
         Top             =   2350
         Width           =   1275
      End
      Begin VB.Label MIC004_00 
         AutoSize        =   -1  'True
         Caption         =   "x c"
         BeginProperty Font 
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
         Left            =   2040
         TabIndex        =   244
         Top             =   1950
         Width           =   210
      End
   End
   Begin VB.Frame MIC001 
      Caption         =   "Cultivo + Antibiograma"
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
      Height          =   2595
      Left            =   1080
      TabIndex        =   179
      Top             =   2340
      Visible         =   0   'False
      Width           =   6360
      Begin VB.TextBox MIC001_06 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1425
         TabIndex        =   5
         Top             =   1845
         Width           =   4480
      End
      Begin VB.TextBox MIC001_05 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1425
         TabIndex        =   3
         Top             =   1530
         Width           =   4480
      End
      Begin VB.TextBox MIC001_02 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1665
         TabIndex        =   1
         Text            =   "Escherichia coli"
         Top             =   945
         Width           =   2175
      End
      Begin VB.TextBox MIC001_07 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1420
         TabIndex        =   7
         Top             =   2235
         Width           =   4830
      End
      Begin VB.CommandButton MIC001_03 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   5910
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1545
         Width           =   315
      End
      Begin VB.CommandButton MIC001_03 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5910
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1875
         Width           =   315
      End
      Begin VB.ComboBox MIC001_01 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":3F28
         Left            =   960
         List            =   "frmMicrobiologia.frx":3F83
         TabIndex        =   0
         Text            =   "MIC001_01"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox MIC001_04 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4080
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label MIC001_00 
         Caption         =   "Resistente a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         TabIndex        =   186
         Top             =   1875
         Width           =   1215
      End
      Begin VB.Label MIC001_00 
         Caption         =   "Sensible a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   185
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label MIC001_00 
         Caption         =   "Análisis Microbiológico"
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
         Left            =   60
         TabIndex        =   184
         Top             =   705
         Width           =   2175
      End
      Begin VB.Label MIC001_00 
         Caption         =   "Antibiograma"
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
         TabIndex        =   183
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label MIC001_00 
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
         TabIndex        =   182
         Top             =   2265
         Width           =   1335
      End
      Begin VB.Label MIC001_00 
         Caption         =   "Bacteria aislada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   181
         Top             =   975
         Width           =   1815
      End
      Begin VB.Label MIC001_00 
         Caption         =   "Muestra:"
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
         TabIndex        =   180
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame MIC002 
      Caption         =   "BK Seriado (Esputo)"
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
      Height          =   1335
      Left            =   1110
      TabIndex        =   297
      Top             =   2340
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox MIC002_04 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1370
         TabIndex        =   11
         Top             =   960
         Width           =   4950
      End
      Begin VB.TextBox MIC002_03 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1370
         TabIndex        =   10
         Top             =   600
         Width           =   4950
      End
      Begin VB.ComboBox MIC002_02 
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":4153
         Left            =   60
         List            =   "frmMicrobiologia.frx":415D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox MIC002_01 
         Height          =   315
         ItemData        =   "frmMicrobiologia.frx":4175
         Left            =   1370
         List            =   "frmMicrobiologia.frx":41A0
         TabIndex        =   8
         Text            =   "MIC002_01"
         Top             =   220
         Width           =   2895
      End
      Begin VB.Label MIC002_00 
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
         Index           =   1
         Left            =   60
         TabIndex        =   299
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label MIC002_00 
         Caption         =   "Muestra:"
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
         TabIndex        =   298
         Top             =   290
         Width           =   855
      End
   End
   Begin VB.Frame MIC006 
      Caption         =   "Semen: Espermatograma"
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
      Height          =   6600
      Left            =   1080
      TabIndex        =   257
      Top             =   2340
      Visible         =   0   'False
      Width           =   7005
      Begin VB.TextBox MIC006_29 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5370
         TabIndex        =   55
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox MIC006_26 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1530
         TabIndex        =   52
         Top             =   570
         Width           =   1455
      End
      Begin VB.TextBox MIC006_28 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1530
         TabIndex        =   54
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox MIC006_25 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5370
         TabIndex        =   51
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox MIC006_27 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5370
         TabIndex        =   53
         Top             =   570
         Width           =   1455
      End
      Begin VB.TextBox MIC006_24 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1530
         TabIndex        =   50
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox MIC006_22 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1725
         TabIndex        =   77
         Top             =   5820
         Width           =   615
      End
      Begin VB.TextBox MIC006_21 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5925
         TabIndex        =   76
         Top             =   5490
         Width           =   615
      End
      Begin VB.TextBox MIC006_19 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5925
         TabIndex        =   74
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox MIC006_18 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1725
         TabIndex        =   73
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox MIC006_20 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1725
         TabIndex        =   75
         Top             =   5490
         Width           =   615
      End
      Begin VB.TextBox MIC006_17 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5925
         TabIndex        =   72
         Top             =   4530
         Width           =   615
      End
      Begin VB.TextBox MIC006_15 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5925
         TabIndex        =   70
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox MIC006_14 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2325
         TabIndex        =   69
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox MIC006_16 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2325
         TabIndex        =   71
         Top             =   4530
         Width           =   615
      End
      Begin VB.TextBox MIC006_13 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5925
         TabIndex        =   68
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox MIC006_11 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2205
         TabIndex        =   66
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox MIC006_10 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   525
         TabIndex        =   65
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox MIC006_12 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4125
         TabIndex        =   67
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox MIC006_23 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1420
         TabIndex        =   78
         Top             =   6195
         Width           =   4830
      End
      Begin VB.TextBox MIC006_05 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1550
         TabIndex        =   60
         Top             =   2820
         Width           =   1215
      End
      Begin VB.TextBox MIC006_04 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1550
         TabIndex        =   59
         Top             =   2505
         Width           =   1215
      End
      Begin VB.TextBox MIC006_03 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1550
         TabIndex        =   58
         Top             =   2175
         Width           =   1215
      End
      Begin VB.TextBox MIC006_02 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1550
         TabIndex        =   57
         Top             =   1860
         Width           =   1215
      End
      Begin VB.TextBox MIC006_01 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1550
         TabIndex        =   56
         Top             =   1545
         Width           =   1215
      End
      Begin VB.TextBox MIC006_06 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5505
         TabIndex        =   61
         Top             =   1530
         Width           =   855
      End
      Begin VB.TextBox MIC006_07 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5505
         TabIndex        =   62
         Top             =   1845
         Width           =   855
      End
      Begin VB.TextBox MIC006_08 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5505
         TabIndex        =   63
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox MIC006_09 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5505
         TabIndex        =   64
         Top             =   2490
         Width           =   855
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "Hora procesamiento"
         BeginProperty Font 
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
         Left            =   3840
         TabIndex        =   386
         Top             =   930
         Width           =   1440
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   51
         Left            =   120
         TabIndex        =   385
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "Hora toma muestra"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   384
         Top             =   930
         Width           =   1380
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   49
         Left            =   3840
         TabIndex        =   383
         Top             =   270
         Width           =   1440
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Abstinencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   48
         Left            =   3840
         TabIndex        =   382
         Top             =   600
         Width           =   1440
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Muestra Nº"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   47
         Left            =   120
         TabIndex        =   381
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label MIC006_00 
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
         Index           =   46
         Left            =   2400
         TabIndex        =   380
         Top             =   5850
         Width           =   165
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "Anomalías en cola"
         BeginProperty Font 
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
         Left            =   195
         TabIndex        =   379
         Top             =   5850
         Width           =   1275
      End
      Begin VB.Label MIC006_00 
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
         Index           =   44
         Left            =   2400
         TabIndex        =   378
         Top             =   5520
         Width           =   165
      End
      Begin VB.Label MIC006_00 
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
         Index           =   43
         Left            =   6585
         TabIndex        =   377
         Top             =   5190
         Width           =   165
      End
      Begin VB.Label MIC006_00 
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
         Index           =   42
         Left            =   6585
         TabIndex        =   376
         Top             =   5520
         Width           =   165
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "Anomalías en segmento intermedio"
         BeginProperty Font 
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
         Left            =   3420
         TabIndex        =   375
         Top             =   5520
         Width           =   2490
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "Inmaduros"
         BeginProperty Font 
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
         Left            =   3420
         TabIndex        =   374
         Top             =   5190
         Width           =   765
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "Normales"
         BeginProperty Font 
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
         Left            =   195
         TabIndex        =   373
         Top             =   5190
         Width           =   660
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "Anomalías en cabeza"
         BeginProperty Font 
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
         Left            =   195
         TabIndex        =   372
         Top             =   5520
         Width           =   1500
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Morfología"
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
         Index           =   37
         Left            =   120
         TabIndex        =   371
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label MIC006_00 
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
         Height          =   255
         Index           =   36
         Left            =   2400
         TabIndex        =   370
         Top             =   5190
         Width           =   375
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "x 10"
         BeginProperty Font 
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
         Left            =   3000
         TabIndex        =   369
         Top             =   4560
         Width           =   315
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "x 10"
         BeginProperty Font 
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
         Left            =   6585
         TabIndex        =   368
         Top             =   4230
         Width           =   315
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "x 10"
         BeginProperty Font 
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
         Left            =   6585
         TabIndex        =   367
         Top             =   4560
         Width           =   315
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "Espermatozoides eyectados / ml"
         BeginProperty Font 
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
         Left            =   3540
         TabIndex        =   366
         Top             =   4560
         Width           =   2310
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "Espermatozoides motiles / ml"
         BeginProperty Font 
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
         Left            =   3540
         TabIndex        =   365
         Top             =   4230
         Width           =   2295
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "Espermatozoides / ml"
         BeginProperty Font 
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
         Left            =   195
         TabIndex        =   364
         Top             =   4230
         Width           =   1515
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "Espermatozoides / eyaculado"
         BeginProperty Font 
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
         Left            =   195
         TabIndex        =   363
         Top             =   4560
         Width           =   2100
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Concentración"
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
         Index           =   28
         Left            =   120
         TabIndex        =   362
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label MIC006_00 
         Caption         =   "x 10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   27
         Left            =   3000
         TabIndex        =   361
         Top             =   4230
         Width           =   375
      End
      Begin VB.Label MIC006_00 
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
         Index           =   26
         Left            =   4800
         TabIndex        =   360
         Top             =   3510
         Width           =   165
      End
      Begin VB.Label MIC006_00 
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
         Index           =   25
         Left            =   2865
         TabIndex        =   359
         Top             =   3510
         Width           =   165
      End
      Begin VB.Label MIC006_00 
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
         Left            =   6585
         TabIndex        =   358
         Top             =   3510
         Width           =   165
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "G0"
         BeginProperty Font 
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
         Left            =   5700
         TabIndex        =   357
         Top             =   3510
         Width           =   195
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "G2"
         BeginProperty Font 
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
         Left            =   1980
         TabIndex        =   356
         Top             =   3510
         Width           =   195
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "G3"
         BeginProperty Font 
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
         Left            =   315
         TabIndex        =   355
         Top             =   3510
         Width           =   195
      End
      Begin VB.Label MIC006_00 
         AutoSize        =   -1  'True
         Caption         =   "G1"
         BeginProperty Font 
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
         Left            =   3915
         TabIndex        =   354
         Top             =   3510
         Width           =   195
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Motilidad"
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
         Index           =   10
         Left            =   120
         TabIndex        =   353
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label MIC006_00 
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
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   352
         Top             =   3510
         Width           =   375
      End
      Begin VB.Label MIC006_00 
         Caption         =   "min"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   2820
         TabIndex        =   275
         Top             =   2850
         Width           =   375
      End
      Begin VB.Label MIC006_00 
         Caption         =   "ml"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   2820
         TabIndex        =   274
         Top             =   1575
         Width           =   375
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Otros datos"
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
         Left            =   60
         TabIndex        =   273
         Top             =   6225
         Width           =   1335
      End
      Begin VB.Label MIC006_00 
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
         Index           =   9
         Left            =   4020
         TabIndex        =   272
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Examen Macroscópico"
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
         TabIndex        =   271
         Top             =   1305
         Width           =   1935
      End
      Begin VB.Label MIC006_00 
         Caption         =   "T. Licuefacción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   255
         TabIndex        =   270
         Top             =   2850
         Width           =   1335
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Aspecto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   255
         TabIndex        =   269
         Top             =   2535
         Width           =   1335
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   255
         TabIndex        =   268
         Top             =   2205
         Width           =   1095
      End
      Begin VB.Label MIC006_00 
         Caption         =   "pH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   255
         TabIndex        =   267
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Volumen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   255
         TabIndex        =   266
         Top             =   1575
         Width           =   1095
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Leucocitos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   4200
         TabIndex        =   265
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Glóbulos Rojos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   4200
         TabIndex        =   264
         Top             =   1875
         Width           =   1095
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Gérmenes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   4200
         TabIndex        =   263
         Top             =   2190
         Width           =   1095
      End
      Begin VB.Label MIC006_00 
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   17
         Left            =   4200
         TabIndex        =   262
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label MIC006_00 
         Caption         =   "x c"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   6405
         TabIndex        =   261
         Top             =   1875
         Width           =   375
      End
      Begin VB.Label MIC006_00 
         Caption         =   " x c"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   6405
         TabIndex        =   260
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label MIC006_00 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   18
         Left            =   6405
         TabIndex        =   259
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label MIC006_00 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   6405
         TabIndex        =   258
         Top             =   2190
         Width           =   375
      End
   End
   Begin VB.Frame MIC032 
      Caption         =   "Parasitológico Directo"
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
      Left            =   1110
      TabIndex        =   328
      Top             =   2370
      Visible         =   0   'False
      Width           =   7220
      Begin VB.TextBox MIC032_05 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   1150
         MultiLine       =   -1  'True
         TabIndex        =   332
         Text            =   "frmMicrobiologia.frx":4240
         Top             =   855
         Width           =   5970
      End
      Begin VB.ComboBox MIC032_01 
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
         ItemData        =   "frmMicrobiologia.frx":428F
         Left            =   1150
         List            =   "frmMicrobiologia.frx":4299
         TabIndex        =   331
         Text            =   "MIC032_01"
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox MIC032_03 
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
         ItemData        =   "frmMicrobiologia.frx":42B1
         Left            =   1150
         List            =   "frmMicrobiologia.frx":42BB
         TabIndex        =   330
         Text            =   "MIC032_03"
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox MIC032_06 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   329
         Top             =   1440
         Width           =   5970
      End
      Begin VB.Label MIC032_00 
         Caption         =   "Helmintos"
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
         Left            =   60
         TabIndex        =   336
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label MIC032_00 
         Caption         =   "Protozoarios"
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
         TabIndex        =   335
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label MIC032_00 
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
         Left            =   60
         TabIndex        =   334
         Top             =   885
         Width           =   1335
      End
      Begin VB.Label MIC032_00 
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
         Index           =   3
         Left            =   60
         TabIndex        =   333
         Top             =   1470
         Width           =   1335
      End
   End
   Begin VB.Frame MIC033 
      Caption         =   "Parasitológico Seriado"
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
      Left            =   1050
      TabIndex        =   337
      Top             =   2370
      Visible         =   0   'False
      Width           =   7220
      Begin VB.ComboBox MIC033_03 
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
         ItemData        =   "frmMicrobiologia.frx":42D3
         Left            =   1150
         List            =   "frmMicrobiologia.frx":42DD
         TabIndex        =   341
         Text            =   "MIC033_03"
         Top             =   540
         Width           =   1215
      End
      Begin VB.ComboBox MIC033_01 
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
         ItemData        =   "frmMicrobiologia.frx":42F5
         Left            =   1150
         List            =   "frmMicrobiologia.frx":42FF
         TabIndex        =   340
         Text            =   "MIC033_01"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox MIC033_05 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   1140
         MultiLine       =   -1  'True
         TabIndex        =   339
         Text            =   "frmMicrobiologia.frx":4317
         Top             =   870
         Width           =   5970
      End
      Begin VB.TextBox MIC033_06 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   338
         Top             =   1440
         Width           =   5970
      End
      Begin VB.Label MIC033_00 
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
         Left            =   60
         TabIndex        =   345
         Top             =   885
         Width           =   1335
      End
      Begin VB.Label MIC033_00 
         Caption         =   "Protozoarios"
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
         TabIndex        =   344
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label MIC033_00 
         Caption         =   "Helmintos"
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
         Left            =   60
         TabIndex        =   343
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label MIC033_00 
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
         Index           =   3
         Left            =   60
         TabIndex        =   342
         Top             =   1470
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMicrobiologia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Resultado de Microbiología
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
Dim ml_boton As String
Dim I As Integer
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
  Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =12")
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
      cmdCerrar1_Click
    Case vbKeyF2
      cmdGrabar_Click
  End Select
End Sub

Public Sub LlenaComboFarmacos(GF As ComboBox)
  GF.AddItem "Todos los grupos"
  GF.AddItem "PENICILINAS"
  GF.AddItem "CEFALOSPIRINAS"
  GF.AddItem "MACROLIDOS"
  GF.AddItem "AMINOGLUCOSIDOS"
  GF.AddItem "TETRACICLINAS"
  GF.AddItem "QUINOLONAS"
  GF.AddItem "SULFONAMIDAS"
  GF.AddItem "NITROFURANOS"
  GF.AddItem "DER. ÁCIDO TRICLORACET"
  GF.AddItem "LINCOSAMIDA"
  GF.AddItem "GLICOPEPTIDO"
  GF.AddItem "MONOBACTANS"
  GF.AddItem "DERIVADOS ÁCIDO FOSFÓRICO"
  GF.AddItem "NITROIMIDAZOL"
End Sub

Public Sub LlenaListFarmacos(Indice As Integer, f As ListBox)
  If Indice = 1 Then
    f.AddItem "Penicilina"
    f.AddItem "Ampicilina"
    f.AddItem "Amox. + Ac. Clavulánico"
    f.AddItem "Oxacilina"
    f.AddItem "Dicloxacilina"
  ElseIf Indice = 2 Then
    f.AddItem "Cefazolina"
    f.AddItem "Ceftazidina"
    f.AddItem "Ceftriaxona"
    f.AddItem "Cefradina"
    f.AddItem "Cefpirome"
  ElseIf Indice = 3 Then
    f.AddItem "Eritromicina"
    f.AddItem "Claritromicina"
    f.AddItem "Azitromicina"
  ElseIf Indice = 4 Then
    f.AddItem "Amikacina"
    f.AddItem "Gentamicina"
  ElseIf Indice = 5 Then
    f.AddItem "Tetraciclina"
    f.AddItem "Doxiciclina"
  ElseIf Indice = 6 Then
    f.AddItem "Ciprofloxacina"
    f.AddItem "Norfloxacina"
    f.AddItem "Ofloxacina"
    f.AddItem "Ácido Pipemídico"
    f.AddItem "Ácido Nalidíxico"
  ElseIf Indice = 7 Then
    f.AddItem "Cotrimoxazol"
  ElseIf Indice = 8 Then
    f.AddItem "Nitrofurantoína"
    f.AddItem "Furazolidona"
  ElseIf Indice = 9 Then
    f.AddItem "Cloranfenicol"
  ElseIf Indice = 10 Then
    f.AddItem "Clindamicina"
  ElseIf Indice = 11 Then
    f.AddItem "Vancomicina"
  ElseIf Indice = 12 Then
    f.AddItem "Aztreonam"
    f.AddItem "Inipenem"
  ElseIf Indice = 13 Then
    f.AddItem "Fosfomicina Trometanol"
  Else
    f.AddItem "Metronidazol"
  End If
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
  Fra.Top = UcPacienteDatos1.Top + UcPacienteDatos1.Height + 700  '350
  fraBoton.Top = Fra.Top + Fra.Height
  Me.Height = fraBoton.Top + fraBoton.Height + 500
End Sub

Private Sub cboGF_Click()
  lstF.Clear
  If cboGF.ListIndex = 0 Then
    For I = 1 To cboGF.ListCount - 1
      Call LlenaListFarmacos(I, lstF)
    Next I
  Else
    Call LlenaListFarmacos(cboGF.ListIndex, lstF)
  End If
End Sub

Private Sub cmbResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdCerrar1_Click()
  fraATB.Visible = False
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
  If ml_CodigoPruebaSeleccionada = "MIC032" Then 'Parasitologico Directo
    'MIC032
    ml_resultado = MIC032_01.Text & "\" & MIC032_03.Text & "\" & MIC032_05.Text
    ml_observacion = MIC032_06.Text
  ElseIf ml_CodigoPruebaSeleccionada = "MIC033" Then 'Parasitológico seriado
    'MIC033
    ml_resultado = MIC033_01.Text & "\" & MIC033_03.Text & "\" & MIC033_05.Text
    ml_observacion = MIC033_06.Text
  ElseIf ml_CodigoPruebaSeleccionada = "MIC006" Then 'Espermatograma
    'MIC006
    ml_resultado = MIC006_24.Text & "\" & MIC006_25.Text & "\" & MIC006_26.Text & "\" & MIC006_27.Text & "\" & MIC006_28.Text & "\" & MIC006_29.Text & "\" & MIC006_01.Text & "\" & MIC006_02.Text & "\" & MIC006_03.Text & "\" & MIC006_04.Text & "\" & MIC006_05.Text & "\" & MIC006_06.Text & "\" & MIC006_07.Text & "\" & MIC006_08.Text & "\" & MIC006_09.Text & "\" & MIC006_10.Text & "\" & MIC006_11.Text & "\" & MIC006_12.Text & "\" & MIC006_13.Text & "\" & MIC006_14.Text & "\" & MIC006_15.Text & "\" & MIC006_16.Text & "\" & MIC006_17.Text & "\" & MIC006_18.Text & "\" & MIC006_19.Text & "\" & MIC006_20.Text & "\" & MIC006_21.Text & "\" & MIC006_22.Text
    ml_observacion = MIC006_23.Text
  ElseIf ml_CodigoPruebaSeleccionada = "MIC031" Then
    ml_resultado = MIC031_10(0).Value & "\" & MIC031_10(1).Value & "\" & MIC031_10(2).Value & "\" & MIC031_10(3).Value & "\" & MIC031_10(4).Value & "\" & MIC031_10(5).Value & "\" & MIC031_10(6).Value & "\" & MIC031_10(7).Value & "\" & MIC031_10(8).Value & "\" & MIC031_10(9).Value & "\" & MIC031_10(10).Value & "\" & MIC031_10(11).Value & "\" & MIC031_10(12).Value & "\" & MIC031_10(13).Value & "\" & MIC031_10(14).Value & "\" & MIC031_10(15).Value & "\" & MIC031_10(16).Value & "\" & MIC031_10(17).Value & "\" & MIC031_10(18).Value & "\" & MIC031_10(19).Value & "\" & MIC031_10(20).Value & "\" & MIC031_10(21).Value & "\" & MIC031_10(22).Value & "\" & MIC031_10(23).Value & "\" & _
                   MIC031_09.Text & "\" & _
                   MIC031_11(0).Value & "\" & MIC031_11(1).Value & "\" & MIC031_11(2).Value & "\" & MIC031_11(3).Value & "\" & MIC031_11(4).Value & "\" & MIC031_11(5).Value & "\" & MIC031_11(6).Value & "\" & MIC031_11(7).Value & "\" & MIC031_11(8).Value & "\" & MIC031_11(9).Value & "\" & MIC031_11(10).Value & "\" & MIC031_11(11).Value & "\" & MIC031_11(12).Value & "\" & MIC031_11(13).Value & "\" & MIC031_11(14).Value & "\" & _
                   MIC031_01.Text & "\" & MIC031_03.Text & "\" & MIC031_04.Text & "\" & MIC031_05.Text & "\" & MIC031_06.Text & "\" & MIC031_07.Text & "\" & _
                   MIC031_02(0).Value & "\" & MIC031_02(1).Value & "\" & MIC031_02(2).Value & "\" & MIC031_02(3).Value & "\" & MIC031_02(4).Value & "\" & MIC031_02(5).Value & "\" & MIC031_02(6).Value & "\" & MIC031_02(7).Value & "\" & MIC031_02(8).Value & "\" & MIC031_02(9).Value & "\" & MIC031_02(10).Value & "\" & MIC031_02(11).Value & "\" & MIC031_02(12).Value & "\" & MIC031_02(13).Value & "\" & MIC031_02(14).Value & "\" & MIC031_02(15).Value & "\" & MIC031_02(16).Value & "\" & MIC031_02(17).Value & "\" & MIC031_02(18).Value & "\" & MIC031_02(19).Value & "\" & MIC031_02(20).Value & "\" & MIC031_02(21).Value & "\" & MIC031_02(22).Value & "\" & MIC031_02(23).Value & "\" & MIC031_02(24).Value & "\" & MIC031_02(25).Value & "\" & MIC031_02(26).Value & "\" & MIC031_02(27).Value & "\" & MIC031_02(28).Value
    ml_observacion = MIC031_08.Text
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
  'debb-2/3/2015
  mo_ReglasLaboratorio.LabIngresaResultados idPrueba, idOrden, ml_resultado, ml_observacion, idUsuario, ml_nombreRealiza, _
                        ml_DetalleOrden, ml_idOrdenLab, "", "", ml_idPaciente, CDate(Me.txtFresultado.Text), mo_lcNombrePc, _
                        mo_lnIdTablaLISTBARITEMS, Me.UcPacienteDatos1.DevuelveHistoriaApellidosYnombre, Me.Caption, _
                        , IIf(Len(Trim(MIC033_01.Text)) = 0, 3, IIf(MIC033_01.ListIndex = 0, 1, 2))   '1-parasitosis positiva, 2-parasitosis negativa
End Sub
Private Sub cmdImprimir_Click()
  ml_resultado = mo_ReglasLaboratorio.LabRecuperaResultados_Res(idPrueba, idOrden)
  ml_observacion = mo_ReglasLaboratorio.LabRecuperaResultados_Obs(idPrueba, idOrden)
  Dim ldFechaResultado As Date
  ml_nombreRealiza = mo_ReglasLaboratorio.LabRecuperaResultados_ReaP(idPrueba, idOrden, ldFechaResultado)
  If ml_CodigoPruebaSeleccionada <> "" And Trim(ml_resultado) <> "" Then
    mo_ReglasLaboratorio.LabImprimeCabeceraResultados UcPacienteDatos1.idPaciente, nombrePaciente, UcPacienteDatos1.NroHistoriaClinica, ldFechaResultado, _
                         nombreMedico + mo_ReglasLaboratorio.DevuelveDatosParaImpresionResultadoLaboratorio(ml_idOrden)
    mo_ReglasLaboratorio.LabImprimeResultadosMIC ml_resultado, CStr(ml_CodigoPruebaSeleccionada), Me.Caption, ml_observacion, ml_nombreRealiza
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
  
  If ml_CodigoPruebaSeleccionada = "MIC032" Then 'Parasitologico Directo
    TopBoton MIC032
  ElseIf ml_CodigoPruebaSeleccionada = "MIC033" Then 'Parasitológico seriado
    TopBoton MIC033
  ElseIf ml_CodigoPruebaSeleccionada = "MIC006" Then 'Espermatograma
    TopBoton MIC006
  ElseIf ml_CodigoPruebaSeleccionada = "MIC031" Then
    TopBoton MIC031
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
  If ml_CodigoPruebaSeleccionada = "MIC032" Then 'Parasitologico Directo
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC032_01.ListIndex = Ubica_En_Combo(MIC032_01, Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC032_03.ListIndex = Ubica_En_Combo(MIC032_03, Temp)
    MIC032_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC032_06.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "MIC033" Then 'Parasitológico seriado
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC033_01.ListIndex = Ubica_En_Combo(MIC033_01, Temp)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC033_03.ListIndex = Ubica_En_Combo(MIC033_03, Temp)
    MIC033_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC033_06.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "MIC006" Then 'Espermatograma
    MIC006_24.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_25.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_26.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_27.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_28.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_29.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_07.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_08.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_09.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_10.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_11.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_12.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_13.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_14.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_15.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_16.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_17.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_18.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_19.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_20.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_21.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_22.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC006_23.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "MIC031" Then
    MIC031_10(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(3).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(4).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(5).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(6).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(7).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(8).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(9).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(10).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(11).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(12).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(13).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(14).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(15).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(16).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(17).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(18).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(19).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(20).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(21).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(22).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_10(23).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_09.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(3).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(4).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(5).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(6).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(7).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(8).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(9).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(10).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(11).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(12).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(13).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_11(14).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    'MIC031_02(Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_07.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(0).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(1).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(2).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(3).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(4).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(5).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(6).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(7).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(8).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(9).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(10).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(11).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(12).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(13).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(14).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(15).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(16).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(17).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(18).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(19).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(20).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(21).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(22).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(23).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(24).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(25).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(26).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(27).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_02(28).Value = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    MIC031_08.Text = ml_observacion
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
End Sub

Private Sub lstF_Click()
  Dim Temp1 As String, Temp2 As String
  If ml_boton = 0 Then
    Temp1 = Trim(MIC001_05.Text)
  Else
    Temp1 = Trim(MIC001_06.Text)
  End If
  
  Temp2 = lstF.List(lstF.ListIndex)
  If InStr(1, Temp1, Temp2, vbTextCompare) <> 0 Then
    MsgBox "El fármaco " & Chr(34) & UCase(Temp2) & Chr(34) & " ya ha sido agregado a la lista."
    Exit Sub
  End If
  If Len(Temp1) = 0 Then
    If ml_boton = 0 Then
      MIC001_05.Text = Temp2
    Else
      MIC001_06.Text = Temp2
    End If
  Else
    If ml_boton = 0 Then
      MIC001_05.Text = Temp1 & ", " & Temp2
    Else
      MIC001_06.Text = Temp1 & ", " & Temp2
    End If
  End If
End Sub

Private Sub MIC001_03_Click(Index As Integer)
  'frmFarmacos.idFormulario = Index
  'frmFarmacos.Show vbModal
  ml_boton = Index
  fraATB.Visible = True
  'fraATB.Left = MIC001.Left + MIC001_03(Index).Left
  'fraATB.Left = MIC001.Top + MIC001_03(Index).Top
End Sub



Private Sub MIC006_01_GotFocus()
  SeleccionaTexto MIC006_01
End Sub

Private Sub MIC006_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_02_GotFocus()
  SeleccionaTexto MIC006_02
End Sub

Private Sub MIC006_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_03_GotFocus()
  SeleccionaTexto MIC006_03
End Sub

Private Sub MIC006_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_04_GotFocus()
  SeleccionaTexto MIC006_04
End Sub

Private Sub MIC006_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_05_GotFocus()
  SeleccionaTexto MIC006_05
End Sub

Private Sub MIC006_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_06_GotFocus()
  SeleccionaTexto MIC006_06
End Sub

Private Sub MIC006_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_07_GotFocus()
  SeleccionaTexto MIC006_07
End Sub

Private Sub MIC006_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_08_GotFocus()
  SeleccionaTexto MIC006_08
End Sub

Private Sub MIC006_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_09_GotFocus()
  SeleccionaTexto MIC006_09
End Sub

Private Sub MIC006_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_10_GotFocus()
  SeleccionaTexto MIC006_10
End Sub

Private Sub MIC006_10_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_11_GotFocus()
  SeleccionaTexto MIC006_11
End Sub

Private Sub MIC006_11_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_12_GotFocus()
  SeleccionaTexto MIC006_12
End Sub

Private Sub MIC006_12_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_13_GotFocus()
  SeleccionaTexto MIC006_13
End Sub

Private Sub MIC006_13_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_14_GotFocus()
  SeleccionaTexto MIC006_14
End Sub

Private Sub MIC006_14_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_15_GotFocus()
  SeleccionaTexto MIC006_15
End Sub

Private Sub MIC006_15_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_16_GotFocus()
  SeleccionaTexto MIC006_16
End Sub

Private Sub MIC006_16_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_17_GotFocus()
  SeleccionaTexto MIC006_17
End Sub

Private Sub MIC006_17_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_18_GotFocus()
  SeleccionaTexto MIC006_18
End Sub

Private Sub MIC006_18_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_19_GotFocus()
  SeleccionaTexto MIC006_19
End Sub

Private Sub MIC006_19_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_20_GotFocus()
  SeleccionaTexto MIC006_20
End Sub

Private Sub MIC006_20_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_21_GotFocus()
  SeleccionaTexto MIC006_21
End Sub

Private Sub MIC006_21_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_22_GotFocus()
  SeleccionaTexto MIC006_22
End Sub

Private Sub MIC006_22_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_23_GotFocus()
  SeleccionaTexto MIC006_23
End Sub

Private Sub MIC006_23_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_24_GotFocus()
  SeleccionaTexto MIC006_24
End Sub

Private Sub MIC006_24_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_25_GotFocus()
  SeleccionaTexto MIC006_25
End Sub


Private Sub MIC006_25_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_26_GotFocus()
  SeleccionaTexto MIC006_26
End Sub

Private Sub MIC006_26_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_27_GotFocus()
  SeleccionaTexto MIC006_27
End Sub

Private Sub MIC006_27_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_28_GotFocus()
  SeleccionaTexto MIC006_28
End Sub

Private Sub MIC006_28_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC006_29_GotFocus()
  SeleccionaTexto MIC006_29
End Sub

Private Sub MIC006_29_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC030_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC030_02_GotFocus()
  SeleccionaTexto MIC030_02
End Sub

Private Sub MIC030_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC030_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC030_04_GotFocus()
  SeleccionaTexto MIC030_04
End Sub

Private Sub MIC030_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC030_05_GotFocus()
  SeleccionaTexto MIC030_05
End Sub

Private Sub MIC030_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC030_06_GotFocus()
  SeleccionaTexto MIC030_06
End Sub

Private Sub MIC030_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC031_01_GotFocus()
  SeleccionaTexto MIC031_01
End Sub

Private Sub MIC031_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC031_02_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC031_03_GotFocus()
  SeleccionaTexto MIC031_03
End Sub

Private Sub MIC031_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC031_04_GotFocus()
  SeleccionaTexto MIC031_04
End Sub

Private Sub MIC031_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC031_05_GotFocus()
  SeleccionaTexto MIC031_05
End Sub

Private Sub MIC031_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC031_06_GotFocus()
  SeleccionaTexto MIC031_06
End Sub

Private Sub MIC031_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC031_07_GotFocus()
  SeleccionaTexto MIC031_07
End Sub

Private Sub MIC031_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC031_08_GotFocus()
  SeleccionaTexto MIC031_08
End Sub

Private Sub MIC031_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC031_09_GotFocus()
  SeleccionaTexto MIC031_09
End Sub

Private Sub MIC031_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC031_10_Click(Index As Integer)
  If Index = 23 Then
    If MIC031_10(23).Value = vbChecked Then
      MIC031_09.Visible = True
    Else
      MIC031_09.Visible = False
    End If
  End If
End Sub

Private Sub MIC031_10_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC031_11_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC032_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC032_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC032_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC032_06_GotFocus()
  SeleccionaTexto MIC032_06
End Sub

Private Sub MIC032_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC033_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC033_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC033_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub MIC033_06_GotFocus()
  SeleccionaTexto MIC033_06
End Sub

Private Sub MIC033_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Sub LimpiaVAloresDefault()
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       Exit Sub
    End If
    MIC030_02.Text = ""
    MIC030_04.Text = ""
    MIC030_05.Text = ""
    MIC003_01.Text = ""
    MIC003_03.Text = ""
    MIC003_15.Text = ""
    MIC008_01(0).Text = ""
    MIC008_01(1).Text = ""
    MIC008_01(2).Text = ""
    MIC007_01(8).Text = ""
    MIC005_01.Text = ""
    MIC005_02.Text = ""
    MIC005_04.Text = ""
    MIC005_05.Text = ""
    MIC005_06.Text = ""
    MIC004_01.Text = ""
    MIC004_02.Text = ""
    MIC004_08.Text = ""
    MIC004_07.Text = ""
    MIC001_02.Text = ""
    MIC032_05.Text = ""
    MIC033_05.Text = ""
    MIC030_01.Text = ""
    MIC030_03.Text = ""
    MIC001_01.Text = ""
    MIC002_01.Text = ""
    MIC032_01.Text = ""
    MIC032_03.Text = ""
    MIC033_01.Text = ""
    MIC033_03.Text = ""
End Sub
