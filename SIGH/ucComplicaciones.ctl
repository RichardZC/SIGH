VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucComplicaciones 
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10950
   ScaleHeight     =   4005
   ScaleWidth      =   10950
   Begin VB.Frame Frame8 
      Caption         =   "Diagnósticos de ingreso"
      Height          =   675
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   10785
      Begin VB.CommandButton btnQuitarDx 
         DisabledPicture =   "ucComplicaciones.ctx":0000
         DownPicture     =   "ucComplicaciones.ctx":038B
         Height          =   315
         Left            =   9600
         Picture         =   "ucComplicaciones.ctx":071E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   1005
      End
      Begin VB.CommandButton btnAgregarDx 
         DisabledPicture =   "ucComplicaciones.ctx":0AAF
         DownPicture     =   "ucComplicaciones.ctx":0E98
         Height          =   315
         Left            =   8520
         Picture         =   "ucComplicaciones.ctx":12A4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1005
      End
      Begin VB.CommandButton Command1 
         Caption         =   ".."
         Height          =   315
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   345
      End
      Begin VB.TextBox txtIdDiagnostico 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label56 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3150
         TabIndex        =   5
         Top             =   240
         Width           =   4515
      End
      Begin VB.Label Label60 
         Caption         =   "Diagnostico"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   1065
      End
   End
   Begin UltraGrid.SSUltraGrid SSUltraGrid4 
      Height          =   3105
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   5477
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108864
      Caption         =   "Lista de interconsultas"
   End
End
Attribute VB_Name = "ucComplicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
