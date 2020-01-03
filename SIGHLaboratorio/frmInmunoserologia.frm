VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmInmunoserologia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INMUNOSEROLOGÍA"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   Icon            =   "frmInmunoserologia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   90
      TabIndex        =   163
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
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   164
         Top             =   210
         Width           =   3090
      End
      Begin MSMask.MaskEdBox txtFresultado 
         Height          =   315
         Left            =   5580
         TabIndex        =   165
         Top             =   210
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         TabIndex        =   167
         Top             =   270
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
         TabIndex        =   166
         Top             =   270
         Width           =   945
      End
   End
   Begin VB.Frame fraBoton 
      ForeColor       =   &H00000000&
      Height          =   870
      Left            =   90
      TabIndex        =   158
      Top             =   4020
      Width           =   7155
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmInmunoserologia.frx":0CCA
         DownPicture     =   "frmInmunoserologia.frx":118E
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
         Left            =   3728
         Picture         =   "frmInmunoserologia.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   161
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
         Picture         =   "frmInmunoserologia.frx":1B66
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmInmunoserologia.frx":203F
         DownPicture     =   "frmInmunoserologia.frx":249F
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
         Left            =   2288
         Picture         =   "frmInmunoserologia.frx":2914
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   180
         Width           =   1365
      End
   End
   Begin SIGHLaboratorio.UcPacienteDatos1 UcPacienteDatos1 
      Height          =   1665
      Left            =   60
      TabIndex        =   145
      Top             =   30
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   2937
   End
   Begin VB.Frame INM003 
      Caption         =   "VIH - SIDA (Elisa)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   885
      Left            =   90
      TabIndex        =   77
      Top             =   2370
      Visible         =   0   'False
      Width           =   7155
      Begin VB.ComboBox INM003_01 
         BeginProperty Font 
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
         ItemData        =   "frmInmunoserologia.frx":2D89
         Left            =   120
         List            =   "frmInmunoserologia.frx":2D93
         TabIndex        =   8
         Text            =   "INM003_01"
         Top             =   420
         Width           =   1455
      End
      Begin VB.TextBox INM003_03 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2925
         TabIndex        =   10
         Top             =   435
         Width           =   4170
      End
      Begin VB.TextBox INM003_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   9
         Text            =   "E.I.A."
         Top             =   435
         Width           =   975
      End
      Begin VB.Label INM003_00 
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
         Left            =   180
         TabIndex        =   80
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label INM003_00 
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
         Left            =   5100
         TabIndex        =   79
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label INM003_00 
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
         Left            =   1800
         TabIndex        =   78
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame INM006 
      Caption         =   "ELISA HBsAg (Antígeno Australiano)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   90
      TabIndex        =   81
      Top             =   2370
      Visible         =   0   'False
      Width           =   7155
      Begin VB.TextBox INM006_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Top             =   750
         Width           =   5670
      End
      Begin VB.TextBox INM006_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Text            =   "Positivo si es > de 0.254"
         Top             =   435
         Width           =   2535
      End
      Begin VB.TextBox INM006_03 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   13
         Text            =   "E.I.A."
         Top             =   435
         Width           =   1095
      End
      Begin VB.TextBox INM006_01 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   11
         Top             =   435
         Width           =   1215
      End
      Begin VB.Label INM006_00 
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
         Left            =   60
         TabIndex        =   85
         Top             =   225
         Width           =   2055
      End
      Begin VB.Label INM006_00 
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
         TabIndex        =   84
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label INM006_00 
         Alignment       =   2  'Center
         Caption         =   "Valor de Referencia"
         BeginProperty Font 
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
         Left            =   3060
         TabIndex        =   83
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label INM006_00 
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
         Left            =   5760
         TabIndex        =   82
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame INM009 
      Caption         =   "RPR ó VDRL  (Sïfilis)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   90
      TabIndex        =   86
      Top             =   2370
      Visible         =   0   'False
      Width           =   7155
      Begin VB.ComboBox INM009_01 
         BeginProperty Font 
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
         ItemData        =   "frmInmunoserologia.frx":2DAC
         Left            =   600
         List            =   "frmInmunoserologia.frx":2DB6
         TabIndex        =   26
         Text            =   "INM009_01"
         Top             =   420
         Width           =   1575
      End
      Begin VB.TextBox INM009_03 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1365
         TabIndex        =   28
         Top             =   750
         Width           =   5670
      End
      Begin VB.TextBox INM009_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2580
         TabIndex        =   27
         Text            =   "Aglutinación"
         Top             =   435
         Width           =   1815
      End
      Begin VB.Label INM009_00 
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
         Left            =   660
         TabIndex        =   89
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label INM009_00 
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
         TabIndex        =   88
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label INM009_00 
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
         Left            =   2640
         TabIndex        =   87
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame INM010 
      Caption         =   "PCR (Proteína C Reactiva)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   90
      TabIndex        =   90
      Top             =   2370
      Visible         =   0   'False
      Width           =   7155
      Begin VB.TextBox INM010_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   30
         Top             =   435
         Width           =   975
      End
      Begin VB.TextBox INM010_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5100
         TabIndex        =   31
         Text            =   "Aglutinación"
         Top             =   435
         Width           =   1815
      End
      Begin VB.TextBox INM010_03 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1365
         TabIndex        =   32
         Top             =   750
         Width           =   5670
      End
      Begin VB.ComboBox INM010_01 
         BeginProperty Font 
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
         ItemData        =   "frmInmunoserologia.frx":2DCF
         Left            =   600
         List            =   "frmInmunoserologia.frx":2DD9
         TabIndex        =   29
         Text            =   "INM010_01"
         Top             =   420
         Width           =   1575
      End
      Begin VB.Label INM010_00 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
         BeginProperty Font 
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
         Left            =   3885
         TabIndex        =   147
         Top             =   480
         Width           =   615
      End
      Begin VB.Label INM010_00 
         Alignment       =   2  'Center
         Caption         =   "Valor"
         BeginProperty Font 
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
         Left            =   2820
         TabIndex        =   146
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label INM010_00 
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
         Left            =   5160
         TabIndex        =   93
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label INM010_00 
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
         TabIndex        =   92
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label INM010_00 
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
         Left            =   660
         TabIndex        =   91
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.Frame INM011 
      Caption         =   "Prueba de Látex (Artritest)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   90
      TabIndex        =   94
      Top             =   2400
      Visible         =   0   'False
      Width           =   7155
      Begin VB.TextBox INM011_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   148
         Top             =   435
         Width           =   975
      End
      Begin VB.ComboBox INM011_01 
         BeginProperty Font 
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
         ItemData        =   "frmInmunoserologia.frx":2DF1
         Left            =   600
         List            =   "frmInmunoserologia.frx":2DFB
         TabIndex        =   33
         Text            =   "INM011_01"
         Top             =   420
         Width           =   1575
      End
      Begin VB.TextBox INM011_03 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   35
         Top             =   750
         Width           =   5670
      End
      Begin VB.TextBox INM011_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   34
         Text            =   "Aglutinación"
         Top             =   435
         Width           =   1815
      End
      Begin VB.Label INM011_00 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "mg / dl"
         BeginProperty Font 
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
         Left            =   3885
         TabIndex        =   150
         Top             =   480
         Width           =   615
      End
      Begin VB.Label INM011_00 
         Alignment       =   2  'Center
         Caption         =   "Valor"
         BeginProperty Font 
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
         Left            =   2820
         TabIndex        =   149
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label INM011_00 
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
         Left            =   660
         TabIndex        =   97
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label INM011_00 
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
         TabIndex        =   96
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label INM011_00 
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
         Left            =   5280
         TabIndex        =   95
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame INM030 
      Caption         =   "VIH Prueba Rápida"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   90
      TabIndex        =   141
      Top             =   2400
      Visible         =   0   'False
      Width           =   7155
      Begin VB.ComboBox INM030_01 
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
         ItemData        =   "frmInmunoserologia.frx":2E13
         Left            =   240
         List            =   "frmInmunoserologia.frx":2E1D
         TabIndex        =   162
         Text            =   "INM030_01"
         Top             =   420
         Width           =   1575
      End
      Begin VB.TextBox INM030_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   46
         Top             =   750
         Width           =   5670
      End
      Begin VB.TextBox INM030_03 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   45
         Text            =   "Prueba Rápida"
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label INM030_00 
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
         TabIndex        =   144
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label INM030_00 
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
         TabIndex        =   143
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label INM030_00 
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
         Left            =   5760
         TabIndex        =   142
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Frame INM031 
      Caption         =   "ELISA HTLV 1-2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   90
      TabIndex        =   136
      Top             =   2370
      Visible         =   0   'False
      Width           =   7155
      Begin VB.TextBox INM031_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   50
         Top             =   750
         Width           =   5670
      End
      Begin VB.TextBox INM031_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   48
         Top             =   435
         Width           =   2535
      End
      Begin VB.TextBox INM031_03 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   49
         Text            =   "E.I.A."
         Top             =   435
         Width           =   1095
      End
      Begin VB.TextBox INM031_01 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   47
         Top             =   435
         Width           =   1215
      End
      Begin VB.Label INM031_00 
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
         Left            =   60
         TabIndex        =   140
         Top             =   225
         Width           =   2055
      End
      Begin VB.Label INM031_00 
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
         TabIndex        =   139
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label INM031_00 
         Alignment       =   2  'Center
         Caption         =   "Valor de Referencia"
         BeginProperty Font 
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
         Left            =   3060
         TabIndex        =   138
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label INM031_00 
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
         Left            =   5760
         TabIndex        =   137
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame INM032 
      Caption         =   "ELISA HEPATITIS C - HVC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   90
      TabIndex        =   131
      Top             =   2370
      Visible         =   0   'False
      Width           =   7155
      Begin VB.TextBox INM032_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   54
         Top             =   750
         Width           =   5670
      End
      Begin VB.TextBox INM032_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   52
         Top             =   435
         Width           =   2535
      End
      Begin VB.TextBox INM032_03 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   53
         Text            =   "E.I.A."
         Top             =   435
         Width           =   1095
      End
      Begin VB.TextBox INM032_01 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   51
         Top             =   435
         Width           =   1215
      End
      Begin VB.Label INM032_00 
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
         Left            =   60
         TabIndex        =   135
         Top             =   225
         Width           =   2055
      End
      Begin VB.Label INM032_00 
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
         TabIndex        =   134
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label INM032_00 
         Alignment       =   2  'Center
         Caption         =   "Valor de Referencia"
         BeginProperty Font 
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
         Left            =   3060
         TabIndex        =   133
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label INM032_00 
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
         Left            =   5760
         TabIndex        =   132
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame INM033 
      Caption         =   "ELISA ANTICORE - HBC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   90
      TabIndex        =   126
      Top             =   2400
      Visible         =   0   'False
      Width           =   7155
      Begin VB.TextBox INM033_01 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   55
         Top             =   435
         Width           =   1215
      End
      Begin VB.TextBox INM033_03 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   57
         Text            =   "E.I.A."
         Top             =   435
         Width           =   1095
      End
      Begin VB.TextBox INM033_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   56
         Top             =   435
         Width           =   2535
      End
      Begin VB.TextBox INM033_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   58
         Top             =   750
         Width           =   5670
      End
      Begin VB.Label INM033_00 
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
         Left            =   5760
         TabIndex        =   130
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label INM033_00 
         Alignment       =   2  'Center
         Caption         =   "Valor de Referencia"
         BeginProperty Font 
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
         Left            =   3060
         TabIndex        =   129
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label INM033_00 
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
         TabIndex        =   128
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label INM033_00 
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
         Left            =   60
         TabIndex        =   127
         Top             =   225
         Width           =   2055
      End
   End
   Begin VB.Frame INM034 
      Caption         =   "Coombs Directo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   90
      TabIndex        =   116
      Top             =   2370
      Visible         =   0   'False
      Width           =   7155
      Begin VB.ComboBox INM034_01 
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
         ItemData        =   "frmInmunoserologia.frx":2E35
         Left            =   600
         List            =   "frmInmunoserologia.frx":2E3F
         TabIndex        =   59
         Text            =   "INM034_01"
         Top             =   420
         Width           =   1575
      End
      Begin VB.TextBox INM034_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Top             =   750
         Width           =   5670
      End
      Begin VB.TextBox INM034_03 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5100
         TabIndex        =   61
         Text            =   "Aglutinación"
         Top             =   435
         Width           =   1815
      End
      Begin VB.TextBox INM034_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   60
         Text            =   "Negativo"
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label INM034_00 
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
         Left            =   660
         TabIndex        =   120
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label INM034_00 
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
         TabIndex        =   119
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label INM034_00 
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
         Left            =   5160
         TabIndex        =   118
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label INM034_00 
         Alignment       =   2  'Center
         Caption         =   "Valor de Referencia"
         BeginProperty Font 
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
         Left            =   2770
         TabIndex        =   117
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame INM035 
      Caption         =   "Coombs Indirecto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   90
      TabIndex        =   121
      Top             =   2430
      Visible         =   0   'False
      Width           =   7155
      Begin VB.TextBox INM035_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   64
         Text            =   "Negativo"
         Top             =   435
         Width           =   1815
      End
      Begin VB.TextBox INM035_03 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5100
         TabIndex        =   65
         Text            =   "Aglutinación"
         Top             =   435
         Width           =   1935
      End
      Begin VB.TextBox INM035_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   66
         Top             =   750
         Width           =   5670
      End
      Begin VB.ComboBox INM035_01 
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
         ItemData        =   "frmInmunoserologia.frx":2E57
         Left            =   600
         List            =   "frmInmunoserologia.frx":2E61
         TabIndex        =   63
         Text            =   "INM035_01"
         Top             =   420
         Width           =   1575
      End
      Begin VB.Label INM035_00 
         Alignment       =   2  'Center
         Caption         =   "Valor de Referencia"
         BeginProperty Font 
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
         Left            =   2770
         TabIndex        =   125
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label INM035_00 
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
         Left            =   5160
         TabIndex        =   124
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label INM035_00 
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
         TabIndex        =   123
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label INM035_00 
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
         Left            =   660
         TabIndex        =   122
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.Frame INM036 
      Caption         =   "Antígeno Australiano (HBsAg)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   90
      TabIndex        =   151
      Top             =   2400
      Visible         =   0   'False
      Width           =   7155
      Begin VB.TextBox INM036_01 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   157
         Top             =   435
         Width           =   1335
      End
      Begin VB.TextBox INM036_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2580
         TabIndex        =   153
         Text            =   "Prueba Rápida"
         Top             =   435
         Width           =   1815
      End
      Begin VB.TextBox INM036_03 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1365
         TabIndex        =   152
         Top             =   750
         Width           =   5670
      End
      Begin VB.Label INM036_00 
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
         Index           =   5
         Left            =   2640
         TabIndex        =   156
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label INM036_00 
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
         TabIndex        =   155
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label INM036_00 
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
         Index           =   3
         Left            =   420
         TabIndex        =   154
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.Frame INM012 
      Caption         =   "ASO (Antiestreptolisina O)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   90
      TabIndex        =   98
      Top             =   2430
      Visible         =   0   'False
      Width           =   7155
      Begin VB.ComboBox INM012_01 
         BeginProperty Font 
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
         ItemData        =   "frmInmunoserologia.frx":2E79
         Left            =   600
         List            =   "frmInmunoserologia.frx":2E83
         TabIndex        =   36
         Text            =   "INM012_01"
         Top             =   420
         Width           =   1575
      End
      Begin VB.TextBox INM012_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   39
         Top             =   750
         Width           =   5670
      End
      Begin VB.TextBox INM012_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   37
         Text            =   "Negativo"
         Top             =   435
         Width           =   1215
      End
      Begin VB.TextBox INM012_03 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   38
         Text            =   "Aglutinación"
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label INM012_00 
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
         TabIndex        =   102
         Top             =   225
         Width           =   2175
      End
      Begin VB.Label INM012_00 
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
         TabIndex        =   101
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label INM012_00 
         Alignment       =   2  'Center
         Caption         =   "Valor de Referencia"
         BeginProperty Font 
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
         Left            =   2940
         TabIndex        =   100
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label INM012_00 
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
         Left            =   5760
         TabIndex        =   99
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame INM021 
      Caption         =   "Examen Toxicológico: Marihuana y Cocaina"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1500
      Left            =   90
      TabIndex        =   103
      Top             =   2430
      Visible         =   0   'False
      Width           =   7155
      Begin VB.TextBox INM021_01 
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
         TabIndex        =   115
         Text            =   "Cocaína"
         Top             =   760
         Width           =   2175
      End
      Begin VB.TextBox INM021_01 
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
         TabIndex        =   114
         Text            =   "Marihuana"
         Top             =   430
         Width           =   2175
      End
      Begin VB.TextBox INM021_05 
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
         Left            =   4620
         TabIndex        =   43
         Text            =   "Inmunocromatografia"
         Top             =   760
         Width           =   1815
      End
      Begin VB.ComboBox INM021_04 
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
         ItemData        =   "frmInmunoserologia.frx":2E9B
         Left            =   2640
         List            =   "frmInmunoserologia.frx":2EA5
         TabIndex        =   42
         Text            =   "INM021_04"
         Top             =   760
         Width           =   1575
      End
      Begin VB.TextBox INM021_03 
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
         Left            =   4620
         TabIndex        =   41
         Text            =   "Inmunocromatografia"
         Top             =   430
         Width           =   1815
      End
      Begin VB.TextBox INM021_06 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   44
         Top             =   1110
         Width           =   5670
      End
      Begin VB.ComboBox INM021_02 
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
         ItemData        =   "frmInmunoserologia.frx":2EBD
         Left            =   2640
         List            =   "frmInmunoserologia.frx":2EC7
         TabIndex        =   40
         Text            =   "INM021_02"
         Top             =   430
         Width           =   1575
      End
      Begin VB.Label INM021_00 
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
         Left            =   4680
         TabIndex        =   106
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label INM021_00 
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
         TabIndex        =   105
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label INM021_00 
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
         Left            =   2700
         TabIndex        =   104
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.Frame INM002 
      Caption         =   "Sub Unidad Beta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1230
      Left            =   75
      TabIndex        =   72
      Top             =   2400
      Visible         =   0   'False
      Width           =   7185
      Begin VB.ComboBox INM002_01 
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
         ItemData        =   "frmInmunoserologia.frx":2EDF
         Left            =   240
         List            =   "frmInmunoserologia.frx":2EE9
         TabIndex        =   4
         Text            =   "INM002_01"
         Top             =   435
         Width           =   1575
      End
      Begin VB.TextBox INM002_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   7
         Top             =   840
         Width           =   5670
      End
      Begin VB.TextBox INM002_03 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   6
         Text            =   "E.I.A."
         Top             =   435
         Width           =   1095
      End
      Begin VB.TextBox INM002_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   5
         Top             =   435
         Width           =   855
      End
      Begin VB.Label INM002_00 
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
         TabIndex        =   76
         Top             =   225
         Width           =   3255
      End
      Begin VB.Label INM002_00 
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
         TabIndex        =   75
         Top             =   870
         Width           =   1335
      End
      Begin VB.Label INM002_00 
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
         Left            =   5760
         TabIndex        =   74
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label INM002_00 
         Alignment       =   2  'Center
         Caption         =   "mUI / ml"
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
         Left            =   2760
         TabIndex        =   73
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame INM008 
      Caption         =   "Aglutinaciones Febriles"
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
      Height          =   1605
      Left            =   90
      TabIndex        =   107
      Top             =   2340
      Visible         =   0   'False
      Width           =   7155
      Begin VB.ComboBox INM008_05 
         BeginProperty Font 
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
         ItemData        =   "frmInmunoserologia.frx":2F01
         Left            =   1360
         List            =   "frmInmunoserologia.frx":2F0B
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   525
         Width           =   1215
      End
      Begin VB.TextBox INM008_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   16
         Text            =   "1 / "
         Top             =   240
         Width           =   750
      End
      Begin VB.TextBox INM008_11 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1360
         TabIndex        =   25
         Top             =   1200
         Width           =   5790
      End
      Begin VB.TextBox INM008_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6315
         TabIndex        =   18
         Text            =   "1 / "
         Top             =   240
         Width           =   750
      End
      Begin VB.TextBox INM008_08 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6315
         TabIndex        =   22
         Text            =   "1 / "
         Top             =   555
         Width           =   750
      End
      Begin VB.ComboBox INM008_09 
         BeginProperty Font 
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
         ItemData        =   "frmInmunoserologia.frx":2F23
         Left            =   1360
         List            =   "frmInmunoserologia.frx":2F2D
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   855
         Width           =   1215
      End
      Begin VB.ComboBox INM008_01 
         BeginProperty Font 
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
         ItemData        =   "frmInmunoserologia.frx":2F45
         Left            =   1360
         List            =   "frmInmunoserologia.frx":2F4F
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox INM008_06 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   20
         Text            =   "1 / "
         Top             =   555
         Width           =   750
      End
      Begin VB.ComboBox INM008_07 
         BeginProperty Font 
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
         ItemData        =   "frmInmunoserologia.frx":2F67
         Left            =   5070
         List            =   "frmInmunoserologia.frx":2F71
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   525
         Width           =   1215
      End
      Begin VB.ComboBox INM008_03 
         BeginProperty Font 
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
         ItemData        =   "frmInmunoserologia.frx":2F89
         Left            =   5070
         List            =   "frmInmunoserologia.frx":2F93
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox INM008_10 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   24
         Text            =   "1 / "
         Top             =   885
         Width           =   750
      End
      Begin VB.Label INM008_00 
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
         Index           =   5
         Left            =   60
         TabIndex        =   113
         Top             =   1245
         Width           =   1275
      End
      Begin VB.Label INM008_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tífico ""O"""
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
         Left            =   555
         TabIndex        =   112
         Top             =   300
         Width           =   780
      End
      Begin VB.Label INM008_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tífico ""H"""
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
         Left            =   4260
         TabIndex        =   111
         Top             =   300
         Width           =   765
      End
      Begin VB.Label INM008_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tífico ""A"""
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
         Left            =   570
         TabIndex        =   110
         Top             =   585
         Width           =   765
      End
      Begin VB.Label INM008_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Paratífico ""B"""
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
         Left            =   3975
         TabIndex        =   109
         Top             =   585
         Width           =   1050
      End
      Begin VB.Label INM008_00 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Brucella"
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
         Left            =   720
         TabIndex        =   108
         Top             =   885
         Width           =   615
      End
   End
   Begin VB.Frame INM001 
      Caption         =   "Pregnosticon (Diagnóstico de Embarazo)"
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
      Height          =   1110
      Left            =   90
      TabIndex        =   67
      Top             =   2370
      Visible         =   0   'False
      Width           =   7155
      Begin VB.TextBox INM001_03 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   2
         Text            =   "E.I.A."
         Top             =   435
         Width           =   1095
      End
      Begin VB.TextBox INM001_02 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   1
         Text            =   "25 mUI / ml"
         Top             =   435
         Width           =   1215
      End
      Begin VB.TextBox INM001_04 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   3
         Top             =   780
         Width           =   5670
      End
      Begin VB.ComboBox INM001_01 
         BeginProperty Font 
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
         ItemData        =   "frmInmunoserologia.frx":2FAB
         Left            =   600
         List            =   "frmInmunoserologia.frx":2FB5
         TabIndex        =   0
         Text            =   "INM001_01"
         Top             =   420
         Width           =   1575
      End
      Begin VB.Label INM001_00 
         Alignment       =   2  'Center
         Caption         =   "Tipo de Prueba"
         BeginProperty Font 
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
         Left            =   5760
         TabIndex        =   71
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label INM001_00 
         Alignment       =   2  'Center
         Caption         =   "Sensibilidad"
         BeginProperty Font 
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
         TabIndex        =   70
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label INM001_00 
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
         TabIndex        =   69
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label INM001_00 
         Alignment       =   2  'Center
         Caption         =   "Examen Cualitativo"
         BeginProperty Font 
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
         TabIndex        =   68
         Top             =   225
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmInmunoserologia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Resultados para Inmunoserología
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
  Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =14")
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
  Fra.Top = UcPacienteDatos1.Top + UcPacienteDatos1.Height + 700 '350
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
    MsgBox "Por favor ingrese la Fecha del Resultado", vbInformation, "SIGH "
    Exit Sub
  End If
  ml_nombreRealiza = mo_cmbResponsable.BoundText
  If ml_CodigoPruebaSeleccionada = "INM001" Then 'Diagnóstico de embarazo
    ml_resultado = INM001_01.Text & "\ " & INM001_02.Text & "\ " & INM001_03.Text
    ml_observacion = INM001_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM002" Then 'Sub unidad beta cuantitativo
    ml_resultado = INM002_01.Text & "\ " & INM002_02.Text & "\ " & INM002_03.Text
    ml_observacion = INM002_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM003" Then 'VIH-SIDA
    ml_resultado = INM003_01.Text & "\ " & INM003_02.Text
    ml_observacion = INM003_03.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM006" Then 'Elisa HBSAg (Antígeno australiano)
    ml_resultado = INM006_01.Text & "\ " & INM006_02.Text & "\ " & INM006_03.Text
    ml_observacion = INM006_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM008" Then 'Aglutinaciones febriles
    ml_resultado = INM008_01.Text & "\ " & INM008_02.Text & "\ " & INM008_03.Text & "\ " & INM008_04.Text & "\ " & INM008_05.Text & "\ " & INM008_06.Text & "\ " & INM008_07.Text & "\ " & INM008_08.Text & "\ " & INM008_09.Text & "\ " & INM008_10.Text
    ml_observacion = INM008_11.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM009" Then  'RPR o VDRL (Sífilis)
    ml_resultado = INM009_01.Text & "\ " & INM009_02.Text
    ml_observacion = INM009_03.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM010" Then 'PCR
    ml_resultado = INM010_01.Text & "\ " & INM010_04.Text & "\ " & INM010_02.Text
    ml_observacion = INM010_03.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM011" Then 'Prueba de Latex
    ml_resultado = INM011_01.Text & "\ " & INM011_04.Text & "\ " & INM011_02.Text
    ml_observacion = INM011_03.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM012" Then 'ASO
    ml_resultado = INM012_01.Text & "\ " & INM012_02.Text & "\ " & INM012_03.Text
    ml_observacion = INM012_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM021" Then 'Toxicológico
    ml_resultado = INM021_02.Text & "\ " & INM021_03.Text & "\ " & INM021_04.Text & "\ " & INM021_05.Text
    ml_observacion = INM021_06.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM030" Then 'ELISA Prueba Rápida
    ml_resultado = INM030_01.Text & "\ " & INM030_03.Text
    ml_observacion = INM030_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM031" Then  'ELISA HTLV 1-2
    ml_resultado = INM031_01.Text & "\ " & INM031_02.Text & "\ " & INM031_03.Text
    ml_observacion = INM031_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM032" Then 'Elisa Hepatitis C
    ml_resultado = INM032_01.Text & "\ " & INM032_02.Text & "\ " & INM032_03.Text
    ml_observacion = INM032_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM033" Then 'Elisa Anticore HBC
    ml_resultado = INM033_01.Text & "\ " & INM033_02.Text & "\ " & INM033_03.Text
    ml_observacion = INM033_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM034" Then 'Coombs directo
    ml_resultado = INM034_01.Text & "\ " & INM034_02.Text & "\ " & INM034_03.Text
    ml_observacion = INM034_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM035" Then   'Coombs indirecto
    ml_resultado = INM035_01.Text & "\ " & INM035_02.Text & "\ " & INM035_03.Text
    ml_observacion = INM035_04.Text
  ElseIf ml_CodigoPruebaSeleccionada = "INM036" Then  'Antígeno Australiano (HBsAg)
    ml_resultado = INM036_01.Text & "\ " & INM036_02.Text
    ml_observacion = INM036_03.Text
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
    mo_ReglasLaboratorio.LabImprimeResultadosinm ml_resultado, CStr(ml_CodigoPruebaSeleccionada), Me.Caption, ml_observacion, ml_nombreRealiza
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
  
  If ml_CodigoPruebaSeleccionada = "INM001" Then 'Diagnóstico de embarazo
    TopBoton INM001
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       INM001_01.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "INM002" Then 'Sub unidad beta cuantitativo
    TopBoton INM002
  ElseIf ml_CodigoPruebaSeleccionada = "INM003" Then 'VIH-SIDA
    TopBoton INM003
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       INM003_01.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "INM006" Then 'Elisa HBSAg (Antígeno australiano)
    TopBoton INM006
  ElseIf ml_CodigoPruebaSeleccionada = "INM008" Then 'Aglutinaciones febriles
    TopBoton INM008
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
        INM008_01.ListIndex = 1
        INM008_03.ListIndex = 1
        INM008_05.ListIndex = 1
        INM008_07.ListIndex = 1
        INM008_09.ListIndex = 1
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "INM009" Then  'RPR o VDRL (Sífilis)
    TopBoton INM009
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       INM009_01.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "INM010" Then 'PCR
    TopBoton INM010
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       INM010_01.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "INM011" Then 'Prueba de Latex
    TopBoton INM011
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       INM011_01.ListIndex = 0
    End If
  ElseIf ml_CodigoPruebaSeleccionada = "INM012" Then 'ASO
    TopBoton INM012
  ElseIf ml_CodigoPruebaSeleccionada = "INM021" Then 'Toxicológico
    TopBoton INM021
  ElseIf ml_CodigoPruebaSeleccionada = "INM030" Then 'ELISA Prueba Rápida
    TopBoton INM030
  ElseIf ml_CodigoPruebaSeleccionada = "INM031" Then  'ELISA HTLV 1-2
    TopBoton INM031
  ElseIf ml_CodigoPruebaSeleccionada = "INM032" Then 'Elisa Hepatitis C
    TopBoton INM032
  ElseIf ml_CodigoPruebaSeleccionada = "INM033" Then 'Elisa AntiCore HBC
    TopBoton INM033
  ElseIf ml_CodigoPruebaSeleccionada = "INM034" Then  'Coombs directo
    TopBoton INM034
  ElseIf ml_CodigoPruebaSeleccionada = "INM035" Then   'Coombs indirecto
    TopBoton INM035
  ElseIf ml_CodigoPruebaSeleccionada = "INM036" Then  'Antígeno Australiano (HBsAg)
    TopBoton INM036
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
  If ml_CodigoPruebaSeleccionada = "INM001" Then  'Diagnóstico de embarazo
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM001_01.ListIndex = Ubica_En_Combo(INM001_01, Temp)
    INM001_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM001_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM001_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM002" Then 'Sub unidad beta cuantitativo
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM002_01.ListIndex = Ubica_En_Combo(INM002_01, Temp)
    INM002_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM002_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM002_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM003" Then 'VIH-SIDA
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM003_01.ListIndex = Ubica_En_Combo(INM003_01, Temp)
    INM003_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM003_03.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM006" Then 'Elisa HBSAg (Antígeno australiano)
    INM006_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM006_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM006_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM006_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM008" Then 'Aglutinaciones febriles
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM008_01.ListIndex = Ubica_En_Combo(INM008_01, Temp)
    INM008_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM008_03.ListIndex = Ubica_En_Combo(INM008_03, Temp)
    INM008_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM008_05.ListIndex = Ubica_En_Combo(INM008_05, Temp)
    INM008_06.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM008_07.ListIndex = Ubica_En_Combo(INM008_07, Temp)
    INM008_08.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM008_09.ListIndex = Ubica_En_Combo(INM008_09, Temp)
    INM008_10.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM008_11.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM009" Then  'RPR o VDRL (Sífilis)
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM009_01.ListIndex = Ubica_En_Combo(INM009_01, Temp)
    INM009_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM009_03.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM010" Then 'PCR
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM010_01.ListIndex = Ubica_En_Combo(INM010_01, Temp)
    INM010_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM010_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM010_03.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM011" Then 'Prueba de Latex
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM011_01.ListIndex = Ubica_En_Combo(INM011_01, Temp)
    INM011_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM011_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM011_03.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM012" Then 'ASO
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM012_01.ListIndex = Ubica_En_Combo(INM012_01, Temp)
    INM012_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM012_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM012_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM021" Then 'Toxicológico
    INM021_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM021_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM021_04.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM021_05.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM021_06.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM030" Then 'ELISA Prueba Rápida
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM030_01.ListIndex = Ubica_En_Combo(INM030_01, Temp)
    INM030_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM030_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM031" Then  'ELISA HTLV 1-2
    INM031_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM031_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM031_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM031_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM032" Then 'Elisa Hepatitis C
    INM032_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM032_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM032_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM032_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM033" Then 'Elisa AntiCore HBC
    INM033_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM033_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM033_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM033_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM034" Then  'Coombs directo
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM034_01.ListIndex = Ubica_En_Combo(INM034_01, Temp)
    INM034_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM034_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM034_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM035" Then 'Coombs indirecto
    Temp = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM035_01.ListIndex = Ubica_En_Combo(INM035_01, Temp)
    INM035_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM035_03.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM035_04.Text = ml_observacion
  ElseIf ml_CodigoPruebaSeleccionada = "INM036" Then  'Antígeno Australiano (HBsAg)
    INM036_01.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM036_02.Text = mo_ReglasLaboratorio.SeparaCadena(ml_resultado)
    INM036_03.Text = ml_observacion
  Else
    MsgBox "El formato para el ingreso de resultados de la prueba no esta implementado", vbCritical
    Exit Sub
  End If
End Sub

Private Sub INM001_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM001_02_GotFocus()
  SeleccionaTexto INM001_02
End Sub

Private Sub INM001_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM001_03_GotFocus()
  SeleccionaTexto INM001_03
End Sub

Private Sub INM001_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM001_04_GotFocus()
  SeleccionaTexto INM001_04
End Sub

Private Sub INM001_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM002_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM002_02_GotFocus()
  SeleccionaTexto INM002_02
End Sub

Private Sub INM002_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM002_03_GotFocus()
  SeleccionaTexto INM002_03
End Sub

Private Sub INM002_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM002_04_GotFocus()
  SeleccionaTexto INM002_04
End Sub

Private Sub INM002_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM003_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM003_02_GotFocus()
  SeleccionaTexto INM003_02
End Sub

Private Sub INM003_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM003_03_GotFocus()
  SeleccionaTexto INM003_03
End Sub

Private Sub INM003_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM006_01_GotFocus()
  SeleccionaTexto INM006_01
End Sub

Private Sub INM006_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM006_02_GotFocus()
  SeleccionaTexto INM006_02
End Sub

Private Sub INM006_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM006_03_GotFocus()
  SeleccionaTexto INM006_03
End Sub

Private Sub INM006_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM006_04_GotFocus()
  SeleccionaTexto INM006_04
End Sub

Private Sub INM006_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM008_01_Click()
  If INM008_01.ListIndex = 0 Then
    INM008_02.Text = "1 / 80"
  Else
    INM008_02.Text = ""
  End If
End Sub

Private Sub INM008_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM008_02_GotFocus()
  SeleccionaTexto INM008_02
End Sub

Private Sub INM008_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM008_03_Click()
  If INM008_03.ListIndex = 0 Then
    INM008_04.Text = "1 / 80"
  Else
    INM008_04.Text = ""
  End If
End Sub

Private Sub INM008_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM008_04_GotFocus()
  SeleccionaTexto INM008_04
End Sub

Private Sub INM008_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM008_05_Click()
  If INM008_05.ListIndex = 0 Then
    INM008_06.Text = "1 / 160"
  Else
    INM008_06.Text = ""
  End If
End Sub

Private Sub INM008_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM008_06_GotFocus()
  SeleccionaTexto INM008_06
End Sub

Private Sub INM008_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM008_07_Click()
  If INM008_07.ListIndex = 0 Then
    INM008_08.Text = "1 / 160"
  Else
    INM008_08.Text = ""
  End If
End Sub

Private Sub INM008_07_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM008_08_GotFocus()
  SeleccionaTexto INM008_08
End Sub

Private Sub INM008_08_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM008_09_Click()
  If INM008_09.ListIndex = 0 Then
    INM008_10.Text = "1 / 100"
  Else
    INM008_10.Text = ""
  End If
End Sub

Private Sub INM008_09_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM008_10_GotFocus()
  SeleccionaTexto INM008_10
End Sub

Private Sub INM008_10_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM008_11_GotFocus()
  SeleccionaTexto INM008_11
End Sub

Private Sub INM008_11_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM009_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM009_02_GotFocus()
  SeleccionaTexto INM009_02
End Sub

Private Sub INM009_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM009_03_GotFocus()
  SeleccionaTexto INM009_03
End Sub

Private Sub INM009_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM010_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM010_02_GotFocus()
  SeleccionaTexto INM010_02
End Sub

Private Sub INM010_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM010_03_GotFocus()
  SeleccionaTexto INM010_03
End Sub

Private Sub INM010_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM010_04_GotFocus()
  SeleccionaTexto INM010_04
End Sub

Private Sub INM010_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM011_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM011_02_GotFocus()
  SeleccionaTexto INM011_02
End Sub

Private Sub INM011_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM011_03_GotFocus()
  SeleccionaTexto INM011_03
End Sub

Private Sub INM011_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM011_04_GotFocus()
  SeleccionaTexto INM011_04
End Sub

Private Sub INM011_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM012_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM012_02_GotFocus()
  SeleccionaTexto INM012_02
End Sub

Private Sub INM012_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM012_03_GotFocus()
  SeleccionaTexto INM012_03
End Sub

Private Sub INM012_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM012_04_GotFocus()
  SeleccionaTexto INM012_04
End Sub

Private Sub INM012_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM021_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM021_03_GotFocus()
  SeleccionaTexto INM021_03
End Sub

Private Sub INM021_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM021_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM021_05_GotFocus()
  SeleccionaTexto INM021_05
End Sub

Private Sub INM021_05_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM021_06_GotFocus()
  SeleccionaTexto INM021_06
End Sub

Private Sub INM021_06_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM030_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM030_03_GotFocus()
  SeleccionaTexto INM030_03
End Sub

Private Sub INM030_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM030_04_GotFocus()
  SeleccionaTexto INM030_04
End Sub

Private Sub INM030_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM031_01_GotFocus()
  SeleccionaTexto INM031_01
End Sub

Private Sub INM031_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM031_02_GotFocus()
  SeleccionaTexto INM031_02
End Sub

Private Sub INM031_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM031_03_GotFocus()
  SeleccionaTexto INM031_03
End Sub

Private Sub INM031_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM031_04_GotFocus()
  SeleccionaTexto INM031_04
End Sub

Private Sub INM031_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM032_01_GotFocus()
  SeleccionaTexto INM032_01
End Sub

Private Sub INM032_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM032_02_GotFocus()
  SeleccionaTexto INM032_02
End Sub

Private Sub INM032_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM032_03_GotFocus()
  SeleccionaTexto INM032_03
End Sub

Private Sub INM032_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM032_04_GotFocus()
  SeleccionaTexto INM032_04
End Sub

Private Sub INM032_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM033_01_GotFocus()
  SeleccionaTexto INM033_01
End Sub

Private Sub INM033_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM033_02_GotFocus()
  SeleccionaTexto INM033_02
End Sub

Private Sub INM033_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM033_03_GotFocus()
  SeleccionaTexto INM033_03
End Sub

Private Sub INM033_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM033_04_GotFocus()
  SeleccionaTexto INM033_04
End Sub

Private Sub INM033_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM034_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM034_02_GotFocus()
  SeleccionaTexto INM034_02
End Sub

Private Sub INM034_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM034_03_GotFocus()
  SeleccionaTexto INM034_03
End Sub

Private Sub INM034_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM034_04_GotFocus()
  SeleccionaTexto INM034_04
End Sub

Private Sub INM034_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM035_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM035_02_GotFocus()
  SeleccionaTexto INM035_02
End Sub

Private Sub INM035_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM035_03_GotFocus()
  SeleccionaTexto INM035_03
End Sub

Private Sub INM035_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM035_04_GotFocus()
  SeleccionaTexto INM035_04
End Sub

Private Sub INM035_04_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM036_01_GotFocus()
  SeleccionaTexto INM036_01
End Sub

Private Sub INM036_01_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM036_02_GotFocus()
  SeleccionaTexto INM036_02
End Sub

Private Sub INM036_02_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub INM036_03_GotFocus()
  SeleccionaTexto INM036_03
End Sub

Private Sub INM036_03_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub txtRealiza_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Sub LimpiaVAloresDefault()
    If lcBuscaParametro.SeleccionaFilaParametro(299) = "DEBB" Then
       Exit Sub
    End If
    INM001_02.Text = ""
    INM001_03.Text = ""
    INM002_03.Text = ""
    INM003_02.Text = ""
    INM006_02.Text = ""
    INM006_03.Text = ""
    INM008_02.Text = ""
    INM008_06.Text = ""
    INM008_10.Text = ""
    INM008_04.Text = ""
    INM008_08.Text = ""
    INM009_02.Text = ""
    INM010_02.Text = ""
    INM011_02.Text = ""
    INM021_03.Text = ""
    INM021_05.Text = ""
    INM030_03.Text = ""
    INM031_03.Text = ""
    INM032_03.Text = ""
    INM033_03.Text = ""
    INM034_02.Text = ""
    INM034_03.Text = ""
    INM035_02.Text = ""
    INM035_03.Text = ""
    INM036_02.Text = ""
    INM012_02.Text = ""
    INM012_03.Text = ""
    
    INM001_01.Text = ""
    INM002_01.Text = ""
    INM003_01.Text = ""
    INM009_01.Text = ""
    INM010_01.Text = ""
    INM011_01.Text = ""
    INM021_02.Text = ""
    INM021_04.Text = ""
    INM030_01.Text = ""
    INM034_01.Text = ""
    INM035_01.Text = ""
    INM012_01.Text = ""
End Sub

