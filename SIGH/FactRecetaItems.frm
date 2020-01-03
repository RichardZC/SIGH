VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FactRecetaItems 
   Caption         =   "Form1"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1515
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   9015
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   315
         Left            =   3180
         TabIndex        =   9
         Top             =   210
         Width           =   315
      End
      Begin VB.CheckBox chkCubiertoPorSeguro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Cubierto por el seguro"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6990
         TabIndex        =   8
         Top             =   990
         Width           =   1875
      End
      Begin VB.TextBox txtIdProducto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   225
         Width           =   1185
      End
      Begin VB.TextBox txtPrecioTotal 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7650
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   570
         Width           =   1215
      End
      Begin VB.TextBox txtCantidad 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4530
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   570
         Width           =   1215
      End
      Begin VB.TextBox txtPrecioUnitario 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1905
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   585
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo cmbIdMotivoNoAtencion 
         Height          =   315
         Left            =   5370
         TabIndex        =   10
         Top             =   930
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo cmbIdEstadoProducto 
         Height          =   315
         Left            =   1890
         TabIndex        =   11
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
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
         Width           =   5295
      End
      Begin VB.Label lblIdProducto 
         Caption         =   "IdProducto"
         Height          =   315
         Left            =   315
         TabIndex        =   17
         Top             =   315
         Width           =   1005
      End
      Begin VB.Label lblPrecioTotal 
         Caption         =   "PrecioTotal"
         Height          =   315
         Left            =   6630
         TabIndex        =   16
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label lblCantidad 
         Caption         =   "Cantidad"
         Height          =   315
         Left            =   3600
         TabIndex        =   15
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label lblPrecioUnitario 
         Caption         =   "PrecioUnitario"
         Height          =   315
         Left            =   300
         TabIndex        =   14
         Top             =   615
         Width           =   1005
      End
      Begin VB.Label lblIdEstadoProducto 
         Caption         =   "IdEstadoProducto"
         Height          =   315
         Left            =   300
         TabIndex        =   13
         Top             =   990
         Width           =   1395
      End
      Begin VB.Label lblIdMotivoNoAtencion 
         Caption         =   "IdMotivoNoAtencion"
         Height          =   315
         Left            =   3600
         TabIndex        =   12
         Top             =   990
         Width           =   1425
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1065
      Left            =   60
      TabIndex        =   0
      Top             =   1590
      Width           =   9015
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         Height          =   700
         Left            =   3000
         Picture         =   "FactRecetaItems.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         Height          =   700
         Left            =   4560
         Picture         =   "FactRecetaItems.frx":0475
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FactRecetaItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
