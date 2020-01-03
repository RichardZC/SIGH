VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form FactRecetaDetalle 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Datos del paciente"
      Height          =   1425
      Left            =   60
      TabIndex        =   14
      Top             =   30
      Width           =   9015
      Begin VB.TextBox txtNroHistoriaClinica 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   17
         Top             =   600
         Width           =   1250
      End
      Begin VB.CommandButton btnBuscarHistoriaClinica 
         Caption         =   "..."
         Height          =   315
         Left            =   3240
         TabIndex        =   16
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox txtIdCuentaAtencion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   240
         Width           =   1245
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   3210
         TabIndex        =   18
         Top             =   600
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label8 
         Caption         =   "Nº de cuenta"
         Height          =   255
         Left            =   330
         TabIndex        =   22
         Top             =   300
         Width           =   1365
      End
      Begin VB.Label Label15 
         Caption         =   "Nombres"
         Height          =   285
         Left            =   330
         TabIndex        =   21
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Historia:"
         Height          =   225
         Left            =   330
         TabIndex        =   20
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblNombres 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1920
         TabIndex        =   19
         Top             =   960
         Width           =   6915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   60
      TabIndex        =   4
      Top             =   1470
      Width           =   9015
      Begin VB.TextBox txtNroReceta 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4575
         MaxLength       =   10
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1080
         Width           =   1000
      End
      Begin VB.TextBox txtIdMedico 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1935
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox txtIdServicio 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1935
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   720
         Width           =   1185
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Left            =   3180
         TabIndex        =   7
         Top             =   360
         Width           =   315
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   315
         Left            =   3180
         TabIndex        =   6
         Top             =   720
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtFechaReceta 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   1080
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   11
         Mask            =   "##/ ##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblNroReceta 
         Caption         =   "NroReceta"
         Height          =   315
         Left            =   3570
         TabIndex        =   24
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label lblFechaReceta 
         Caption         =   "FechaReceta"
         Height          =   315
         Left            =   330
         TabIndex        =   23
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lblIdMedico 
         Caption         =   "IdMedico"
         Height          =   315
         Left            =   330
         TabIndex        =   13
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label lblIdServicio 
         Caption         =   "IdServicio"
         Height          =   315
         Left            =   330
         TabIndex        =   12
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3570
         TabIndex        =   11
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3570
         TabIndex        =   10
         Top             =   720
         Width           =   5295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   2715
      Left            =   60
      TabIndex        =   3
      Top             =   3060
      Width           =   9015
      Begin VB.CommandButton btnQuitar 
         Caption         =   "Quitar"
         Height          =   345
         Left            =   1710
         TabIndex        =   27
         Top             =   210
         Width           =   1425
      End
      Begin VB.CommandButton btnAgregar 
         Caption         =   "Agregar"
         Height          =   345
         Left            =   210
         TabIndex        =   26
         Top             =   210
         Width           =   1425
      End
      Begin UltraGrid.SSUltraGrid grdPlanProducto 
         Height          =   1935
         Left            =   210
         TabIndex        =   28
         Top             =   630
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   3413
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108864
         Caption         =   "Medicamentos"
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1065
      Left            =   60
      TabIndex        =   0
      Top             =   5790
      Width           =   9015
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         Height          =   700
         Left            =   4560
         Picture         =   "FactRecetaDetalle.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         Height          =   700
         Left            =   3000
         Picture         =   "FactRecetaDetalle.frx":04EC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FactRecetaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
