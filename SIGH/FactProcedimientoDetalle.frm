VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FactProcedimientoDetalle 
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   1065
      Left            =   60
      TabIndex        =   38
      Top             =   4620
      Width           =   9015
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         Height          =   700
         Left            =   3000
         Picture         =   "FactProcedimientoDetalle.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         Height          =   700
         Left            =   4560
         Picture         =   "FactProcedimientoDetalle.frx":0475
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1515
      Left            =   60
      TabIndex        =   22
      Top             =   3060
      Width           =   9015
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   315
         Left            =   3180
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   990
         Width           =   1875
      End
      Begin VB.TextBox txtIdProducto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   225
         Width           =   1185
      End
      Begin VB.TextBox txtPrecioTotal 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7650
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   570
         Width           =   1215
      End
      Begin VB.TextBox txtCantidad 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4530
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   570
         Width           =   1215
      End
      Begin VB.TextBox txtPrecioUnitario 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1905
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   585
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo cmbIdMotivoNoAtencion 
         Height          =   315
         Left            =   5370
         TabIndex        =   24
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
         TabIndex        =   26
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
         TabIndex        =   37
         Top             =   210
         Width           =   5295
      End
      Begin VB.Label lblIdProducto 
         Caption         =   "IdProducto"
         Height          =   315
         Left            =   315
         TabIndex        =   33
         Top             =   315
         Width           =   1005
      End
      Begin VB.Label lblPrecioTotal 
         Caption         =   "PrecioTotal"
         Height          =   315
         Left            =   6630
         TabIndex        =   31
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label lblCantidad 
         Caption         =   "Cantidad"
         Height          =   315
         Left            =   3600
         TabIndex        =   29
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label lblPrecioUnitario 
         Caption         =   "PrecioUnitario"
         Height          =   315
         Left            =   300
         TabIndex        =   27
         Top             =   615
         Width           =   1005
      End
      Begin VB.Label lblIdEstadoProducto 
         Caption         =   "IdEstadoProducto"
         Height          =   315
         Left            =   300
         TabIndex        =   25
         Top             =   990
         Width           =   1395
      End
      Begin VB.Label lblIdMotivoNoAtencion 
         Caption         =   "IdMotivoNoAtencion"
         Height          =   315
         Left            =   3600
         TabIndex        =   23
         Top             =   990
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   60
      TabIndex        =   8
      Top             =   1470
      Width           =   9015
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   315
         Left            =   4920
         TabIndex        =   21
         Top             =   1080
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   1920
         TabIndex        =   20
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
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   315
         Left            =   3180
         TabIndex        =   18
         Top             =   720
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Left            =   3180
         TabIndex        =   16
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox txtIdServicio 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1935
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   720
         Width           =   1185
      End
      Begin VB.TextBox txtIdMedico 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1935
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3570
         TabIndex        =   19
         Top             =   720
         Width           =   5295
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3570
         TabIndex        =   17
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label lblFechaRealizacion 
         Caption         =   "FechaRealizacion"
         Height          =   315
         Left            =   330
         TabIndex        =   15
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Label lblHoraRealizacion 
         Caption         =   "HoraRealizacion"
         Height          =   315
         Left            =   3390
         TabIndex        =   14
         Top             =   1140
         Width           =   1275
      End
      Begin VB.Label lblIdServicio 
         Caption         =   "IdServicio"
         Height          =   315
         Left            =   330
         TabIndex        =   11
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label lblIdMedico 
         Caption         =   "IdMedico"
         Height          =   315
         Left            =   330
         TabIndex        =   9
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del paciente"
      Height          =   1425
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   9015
      Begin VB.TextBox txtIdCuentaAtencion 
         Height          =   315
         Left            =   1920
         TabIndex        =   13
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
         TabIndex        =   3
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
         TabIndex        =   7
         Top             =   960
         Width           =   6915
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
Attribute VB_Name = "FactProcedimientoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
