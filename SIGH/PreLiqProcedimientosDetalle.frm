VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PreLiquidacionDetalle 
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   60
      TabIndex        =   9
      Top             =   1470
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del paciente"
      Height          =   1425
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.CommandButton btnBuscarHistoriaClinica 
         Caption         =   "..."
         Height          =   315
         Left            =   3540
         TabIndex        =   2
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox txtNroHistoriaClinica 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2250
         TabIndex        =   1
         Top             =   240
         Width           =   1250
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   3540
         TabIndex        =   3
         Top             =   600
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2250
         TabIndex        =   8
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label lblNombres 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2250
         TabIndex        =   7
         Top             =   960
         Width           =   6525
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Historia:"
         Height          =   225
         Left            =   330
         TabIndex        =   6
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Nombres"
         Height          =   285
         Left            =   330
         TabIndex        =   5
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Nº de cuenta"
         Height          =   255
         Left            =   330
         TabIndex        =   4
         Top             =   300
         Width           =   1365
      End
   End
End
Attribute VB_Name = "PreLiquidacionDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
