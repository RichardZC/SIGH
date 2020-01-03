VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form ImagEquipos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "ImagEquipos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraNuevo 
      Caption         =   "Nuevo Equipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4725
      Left            =   5865
      TabIndex        =   4
      Top             =   495
      Width           =   3690
      Begin VB.TextBox txtRuta 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   840
         TabIndex        =   14
         Top             =   3120
         Width           =   2805
      End
      Begin VB.TextBox txtTipo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   825
         TabIndex        =   13
         Top             =   2175
         Width           =   2805
      End
      Begin VB.TextBox txtModelo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   825
         TabIndex        =   12
         Top             =   1290
         Width           =   2805
      End
      Begin VB.TextBox txtMarca 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   810
         TabIndex        =   11
         Top             =   360
         Width           =   2805
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ImagEquipos.frx":0CCA
         DownPicture     =   "ImagEquipos.frx":118E
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
         Left            =   2235
         Picture         =   "ImagEquipos.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3990
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Grabar"
         DisabledPicture =   "ImagEquipos.frx":1B66
         DownPicture     =   "ImagEquipos.frx":1FC6
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
         Left            =   825
         Picture         =   "ImagEquipos.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4005
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ruta"
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
         Left            =   120
         TabIndex        =   10
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         Left            =   135
         TabIndex        =   9
         Top             =   2220
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Modelo"
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
         Left            =   135
         TabIndex        =   8
         Top             =   1275
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
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
         Left            =   135
         TabIndex        =   7
         Top             =   375
         Width           =   465
      End
   End
   Begin VB.CommandButton btnQuitar 
      DisabledPicture =   "ImagEquipos.frx":28B0
      DownPicture     =   "ImagEquipos.frx":2C3B
      Height          =   480
      Left            =   5355
      Picture         =   "ImagEquipos.frx":2FCE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar un Item"
      Top             =   1020
      Width           =   450
   End
   Begin VB.CommandButton btnAgregar 
      DisabledPicture =   "ImagEquipos.frx":335F
      DownPicture     =   "ImagEquipos.frx":3748
      Height          =   480
      Left            =   5355
      Picture         =   "ImagEquipos.frx":3B54
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Agregar un Item"
      Top             =   540
      Width           =   450
   End
   Begin UltraGrid.SSUltraGrid grdEquipos 
      Height          =   4650
      Left            =   30
      TabIndex        =   0
      Top             =   540
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   8202
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de Equipos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Equipos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "ImagEquipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca procedimientos
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit


Private Sub Form_Load()

End Sub
