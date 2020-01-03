VERSION 5.00
Begin VB.Form Complicaciones 
   Caption         =   "Form1"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Complicaciones"
      Height          =   1485
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      Begin VB.TextBox txtIdComplicacion 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1470
         TabIndex        =   6
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton btnBusquedaComplicacion 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   2550
         TabIndex        =   5
         Top             =   240
         Width           =   345
      End
      Begin VB.TextBox txtIdComplicacion 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1470
         TabIndex        =   4
         Top             =   600
         Width           =   1005
      End
      Begin VB.CommandButton btnBusquedaComplicacion 
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   2550
         TabIndex        =   3
         Top             =   600
         Width           =   345
      End
      Begin VB.TextBox txtIdComplicacion 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1470
         TabIndex        =   2
         Top             =   960
         Width           =   1005
      End
      Begin VB.CommandButton btnBusquedaComplicacion 
         Caption         =   "..."
         Height          =   315
         Index           =   2
         Left            =   2550
         TabIndex        =   1
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label66 
         Caption         =   "Diagnostico"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label lblDescripcionComplicacion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   8235
      End
      Begin VB.Label Label68 
         Caption         =   "Diagnostico"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label lblDescripcionComplicacion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   3000
         TabIndex        =   9
         Top             =   600
         Width           =   8235
      End
      Begin VB.Label Label70 
         Caption         =   "Diagnostico"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblDescripcionComplicacion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   3000
         TabIndex        =   7
         Top             =   960
         Width           =   8235
      End
   End
End
Attribute VB_Name = "Complicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
