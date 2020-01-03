VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ExoneracionDetalle 
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1875
      Left            =   60
      TabIndex        =   12
      Top             =   1470
      Width           =   9015
      Begin VB.TextBox txtIdResponsableExoneracion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1905
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   210
         Width           =   1185
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   315
         Left            =   3180
         TabIndex        =   14
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtPrecioUnitario 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1905
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   585
         Width           =   1185
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   1890
         TabIndex        =   16
         Top             =   960
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   11
         Mask            =   "##/ ##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFechaRealizacion 
         Caption         =   "Fecha resultado"
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   1020
         Width           =   1485
      End
      Begin VB.Label lblIdResponsableExoneracion 
         Caption         =   "IdResponsableExoneracion"
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   270
         Width           =   1905
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3570
         TabIndex        =   18
         Top             =   210
         Width           =   5265
      End
      Begin VB.Label lblPrecioUnitario 
         Caption         =   "Valor exonerado"
         Height          =   315
         Left            =   180
         TabIndex        =   17
         Top             =   645
         Width           =   1425
      End
   End
   Begin VB.Frame frame 
      Height          =   1065
      Left            =   60
      TabIndex        =   9
      Top             =   3360
      Width           =   9015
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         Height          =   700
         Left            =   3000
         Picture         =   "ExoneracionDetalle.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         Height          =   700
         Left            =   4560
         Picture         =   "ExoneracionDetalle.frx":0475
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del paciente"
      Height          =   1425
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.TextBox txtIdCuentaAtencion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton btnBuscarHistoriaClinica 
         Caption         =   "..."
         Height          =   315
         Left            =   3240
         TabIndex        =   2
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox txtNroHistoriaClinica 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   1250
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   3210
         TabIndex        =   4
         Top             =   600
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin VB.Label lblNombres 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   960
         Width           =   6915
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Historia:"
         Height          =   225
         Left            =   330
         TabIndex        =   7
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Nombres"
         Height          =   285
         Left            =   330
         TabIndex        =   6
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Nº de cuenta"
         Height          =   255
         Left            =   330
         TabIndex        =   5
         Top             =   300
         Width           =   1365
      End
   End
End
Attribute VB_Name = "ExoneracionDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
