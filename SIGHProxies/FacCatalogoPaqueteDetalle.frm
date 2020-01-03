VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form FacCatalogoPaqueteDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   Icon            =   "FacCatalogoPaqueteDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Farme 
      Height          =   6105
      Left            =   0
      TabIndex        =   8
      Top             =   1530
      Width           =   10785
      Begin VB.CommandButton cmdFarmacia 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         Picture         =   "FacCatalogoPaqueteDetalle.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3000
         Width           =   465
      End
      Begin VB.CommandButton cmdAdministrativo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         Picture         =   "FacCatalogoPaqueteDetalle.frx":1254
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2625
         Width           =   465
      End
      Begin VB.CommandButton cmdCita 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         Picture         =   "FacCatalogoPaqueteDetalle.frx":17DE
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2250
         Width           =   465
      End
      Begin VB.CommandButton cmdPatolClinica 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         Picture         =   "FacCatalogoPaqueteDetalle.frx":1D68
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1860
         Width           =   465
      End
      Begin VB.CommandButton cmdEcogGeneral 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         Picture         =   "FacCatalogoPaqueteDetalle.frx":22F2
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1485
         Width           =   465
      End
      Begin VB.CommandButton cmdEcogObs 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         Picture         =   "FacCatalogoPaqueteDetalle.frx":287C
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1080
         Width           =   465
      End
      Begin VB.CommandButton cmdTomografia 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         Picture         =   "FacCatalogoPaqueteDetalle.frx":2E06
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   720
         Width           =   465
      End
      Begin VB.CommandButton cmdRx 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         Picture         =   "FacCatalogoPaqueteDetalle.frx":3390
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   330
         Width           =   465
      End
      Begin VB.CommandButton btnFarmacia 
         DisabledPicture =   "FacCatalogoPaqueteDetalle.frx":391A
         DownPicture     =   "FacCatalogoPaqueteDetalle.frx":3D03
         Height          =   315
         Left            =   10005
         Picture         =   "FacCatalogoPaqueteDetalle.frx":410F
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2970
         Width           =   645
      End
      Begin VB.CommandButton cmdRefrescaTotal 
         Caption         =   "Refresca Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5715
         TabIndex        =   34
         Top             =   5745
         Width           =   1845
      End
      Begin VB.CommandButton btnAdministrativo 
         DisabledPicture =   "FacCatalogoPaqueteDetalle.frx":451B
         DownPicture     =   "FacCatalogoPaqueteDetalle.frx":4904
         Height          =   315
         Left            =   10005
         Picture         =   "FacCatalogoPaqueteDetalle.frx":4D10
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2610
         Width           =   645
      End
      Begin VB.CommandButton btnCita 
         DisabledPicture =   "FacCatalogoPaqueteDetalle.frx":511C
         DownPicture     =   "FacCatalogoPaqueteDetalle.frx":5505
         Height          =   315
         Left            =   10005
         Picture         =   "FacCatalogoPaqueteDetalle.frx":5911
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2250
         Width           =   645
      End
      Begin VB.CommandButton btnAddPatolClin 
         DisabledPicture =   "FacCatalogoPaqueteDetalle.frx":5D1D
         DownPicture     =   "FacCatalogoPaqueteDetalle.frx":6106
         Height          =   315
         Left            =   10005
         Picture         =   "FacCatalogoPaqueteDetalle.frx":6512
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1860
         Width           =   645
      End
      Begin VB.CommandButton btnAddEcogGeneral 
         DisabledPicture =   "FacCatalogoPaqueteDetalle.frx":691E
         DownPicture     =   "FacCatalogoPaqueteDetalle.frx":6D07
         Height          =   315
         Left            =   10005
         Picture         =   "FacCatalogoPaqueteDetalle.frx":7113
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1470
         Width           =   645
      End
      Begin VB.CommandButton btnAddEcogObst 
         DisabledPicture =   "FacCatalogoPaqueteDetalle.frx":751F
         DownPicture     =   "FacCatalogoPaqueteDetalle.frx":7908
         Height          =   315
         Left            =   10005
         Picture         =   "FacCatalogoPaqueteDetalle.frx":7D14
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1080
         Width           =   645
      End
      Begin VB.CommandButton btnAddTomografia 
         DisabledPicture =   "FacCatalogoPaqueteDetalle.frx":8120
         DownPicture     =   "FacCatalogoPaqueteDetalle.frx":8509
         Height          =   315
         Left            =   10005
         Picture         =   "FacCatalogoPaqueteDetalle.frx":8915
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   690
         Width           =   645
      End
      Begin VB.CommandButton btnQuitar 
         DisabledPicture =   "FacCatalogoPaqueteDetalle.frx":8D21
         DownPicture     =   "FacCatalogoPaqueteDetalle.frx":90AC
         Height          =   315
         Left            =   10035
         Picture         =   "FacCatalogoPaqueteDetalle.frx":943F
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3600
         Width           =   645
      End
      Begin VB.CommandButton btnAddRayos 
         DisabledPicture =   "FacCatalogoPaqueteDetalle.frx":97D0
         DownPicture     =   "FacCatalogoPaqueteDetalle.frx":9BB9
         Height          =   315
         Left            =   10005
         Picture         =   "FacCatalogoPaqueteDetalle.frx":9FC5
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   300
         Width           =   645
      End
      Begin UltraGrid.SSUltraGrid grdCpt 
         Height          =   2295
         Left            =   150
         TabIndex        =   9
         Top             =   3390
         Width           =   9780
         _ExtentX        =   17251
         _ExtentY        =   4048
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
         Caption         =   "Detalle del Paquete"
      End
      Begin VB.Label lblFarmacia 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Left            =   1950
         TabIndex        =   37
         Top             =   3030
         Width           =   7965
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Farmacia"
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
         Left            =   210
         TabIndex        =   36
         Top             =   3030
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Administrativos"
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
         Left            =   210
         TabIndex        =   33
         Top             =   2670
         Width           =   1215
      End
      Begin VB.Label lblAdministrativo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Left            =   1950
         TabIndex        =   32
         Top             =   2670
         Width           =   7965
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "..."
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
         Height          =   210
         Left            =   7710
         TabIndex        =   30
         Top             =   5760
         Width           =   1470
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cita en CE"
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
         Left            =   210
         TabIndex        =   29
         Top             =   2310
         Width           =   840
      End
      Begin VB.Label lblEspecialidad 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Left            =   1950
         TabIndex        =   28
         Top             =   2310
         Width           =   7965
      End
      Begin VB.Label lblPatolClinica 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Left            =   1950
         TabIndex        =   24
         Top             =   1920
         Width           =   7965
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Patol.Clínica"
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
         Left            =   210
         TabIndex        =   23
         Top             =   1920
         Width           =   945
      End
      Begin VB.Label lblEcogGeneral 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Left            =   1950
         TabIndex        =   21
         Top             =   1530
         Width           =   7965
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ecog.General"
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
         Left            =   210
         TabIndex        =   20
         Top             =   1530
         Width           =   1080
      End
      Begin VB.Label lblEcogObst 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Left            =   1950
         TabIndex        =   18
         Top             =   1140
         Width           =   7965
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ecog.Obstét"
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
         Left            =   210
         TabIndex        =   17
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Label lblTomografia 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Left            =   1950
         TabIndex        =   15
         Top             =   750
         Width           =   7965
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tomografía"
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
         Left            =   210
         TabIndex        =   14
         Top             =   750
         Width           =   915
      End
      Begin VB.Label lblRayosX 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Left            =   1950
         TabIndex        =   11
         Top             =   360
         Width           =   7965
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Rayos X"
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
         Left            =   210
         TabIndex        =   10
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame fraDatosGenerales 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   -30
      TabIndex        =   5
      Top             =   0
      Width           =   10800
      Begin VB.CommandButton cmdCopia 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8340
         Picture         =   "FacCatalogoPaqueteDetalle.frx":A3D1
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Copia de un Paquete"
         Top             =   600
         Width           =   345
      End
      Begin VB.TextBox txtCPT 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         MaxLength       =   20
         TabIndex        =   40
         Top             =   960
         Width           =   1275
      End
      Begin VB.CheckBox chkItemFarmacia 
         Alignment       =   1  'Right Justify
         Caption         =   "Es un ITEM más en FARMACIA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7890
         TabIndex        =   38
         Top             =   975
         Width           =   2790
      End
      Begin VB.CheckBox chkEstado 
         Alignment       =   1  'Right Justify
         Caption         =   "Habilitado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9585
         TabIndex        =   26
         Top             =   615
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         MaxLength       =   250
         TabIndex        =   1
         Top             =   600
         Width           =   7425
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblCpt 
         AutoSize        =   -1  'True
         Caption         =   "(en el FUA no se considera los ITEMS consumidos en FARMACIA)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   2175
         TabIndex        =   42
         Top             =   1035
         Width           =   5325
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "CPT (sis)"
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
         Left            =   180
         TabIndex        =   41
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lblTipo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "....."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   10380
         TabIndex        =   39
         Top             =   330
         Width           =   300
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
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
         Left            =   180
         TabIndex        =   7
         Top             =   660
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   0
      TabIndex        =   4
      Top             =   7680
      Width           =   10770
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FacCatalogoPaqueteDetalle.frx":A543
         DownPicture     =   "FacCatalogoPaqueteDetalle.frx":A9A3
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   3990
         Picture         =   "FacCatalogoPaqueteDetalle.frx":AE18
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FacCatalogoPaqueteDetalle.frx":B28D
         DownPicture     =   "FacCatalogoPaqueteDetalle.frx":B751
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   5520
         Picture         =   "FacCatalogoPaqueteDetalle.frx":BC3D
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FacCatalogoPaqueteDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Paquetes para Farmacia y CAJA
'        Programado por: Barrantes D
'        Fecha: Agosto 2010
'
'------------------------------------------------------------------------------------

Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdFactPaquete As Long
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim mrs_Cpt As New ADODB.Recordset
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim lnIdRayosX As Long, lnIdEspecialidadCE As Long, lnIdTomografia As Long
Dim lnIdEcogObst As Long, lnIdEcogGen As Long, lnIdPatolClinica As Long
Dim lnIdFarmacia As Long
Dim lnIdPagoCE As Long, lcPagoCE As String, lnIdAdministrativo As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lcSql As String
Dim lnIdTipoPaquete As Long               ', lnIdTipoSalidaBienInsumo As Long

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property


Property Let ExistenDatos(bValue As Boolean)
    
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let idFactPaquete(lValue As Long)
   ml_IdFactPaquete = lValue
End Property
Property Get idFactPaquete() As Long
   idFactPaquete = ml_IdFactPaquete
End Property

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()
 'debb-08/11/2016
 If mi_Opcion <> sghAgregar Then
    mo_Formulario.HabilitarDeshabilitar txtCodigo, False
    chkItemFarmacia.Enabled = False
 Else
    Dim lbTieneDerecho As Boolean, lcMensajeLicencia As String
    lbTieneDerecho = True  'licencia
    If lbTieneDerecho = False Then
       chkItemFarmacia.Enabled = False
       MsgBox lcMensajeLicencia, vbInformation, Me.Caption
    End If
 End If
 
 '
 Select Case mi_Opcion
     Case sghAgregar
         
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         fraDatosGenerales.Enabled = False
         CargarDatosALosControles
     Case sghEliminar
         fraDatosGenerales.Enabled = False
         CargarDatosALosControles
 End Select
End Sub

Private Sub btnAddEcogGeneral_Click()
    If Me.lblEcogGeneral.Caption = "" Then
       MsgBox "Tiene que elegir un Procedimiento de Ecografía General", vbInformation, Me.Caption
       Exit Sub
    End If
    Dim lbAgregar As Boolean
    lbAgregar = True
    If mrs_Cpt.RecordCount > 0 Then
       mrs_Cpt.MoveFirst
       mrs_Cpt.Find "idProducto=" & lnIdEcogGen
       If Not mrs_Cpt.EOF Then
          MsgBox "Ya está registrado", vbInformation, Me.Caption
          lbAgregar = False
       End If
    End If
    If lbAgregar = True Then
        mrs_Cpt.AddNew
        mrs_Cpt.Fields!idProducto = lnIdEcogGen
        mrs_Cpt.Fields!Producto = lblEcogGeneral.Caption
        mrs_Cpt.Fields!Cantidad = 1
        mrs_Cpt.Fields!precio = 1
        mrs_Cpt.Fields!Importe = 1
        mrs_Cpt.Fields!IdPuntoCarga = sghPtoCargaEcogGeneral
        mrs_Cpt.Fields!idEspecialidadServicio = DevuelveIdEspecialidadServicioSegunPuntoCarga(sghPtoCargaEcogGeneral)
        mrs_Cpt.Update
        SumaTotales
    End If
End Sub

Private Sub btnAddEcogObst_Click()
    If Me.lblEcogObst.Caption = "" Then
       MsgBox "Tiene que elegir un Procedimiento de Ecografía Obstétrica", vbInformation, Me.Caption
       Exit Sub
    End If
    Dim lbAgregar As Boolean
    lbAgregar = True
    If mrs_Cpt.RecordCount > 0 Then
       mrs_Cpt.MoveFirst
       mrs_Cpt.Find "idProducto=" & lnIdEcogObst
       If Not mrs_Cpt.EOF Then
          MsgBox "Ya está registrado", vbInformation, Me.Caption
          lbAgregar = False
       End If
    End If
    If lbAgregar = True Then
        mrs_Cpt.AddNew
        mrs_Cpt.Fields!idProducto = lnIdEcogObst
        mrs_Cpt.Fields!Producto = lblEcogObst.Caption
        mrs_Cpt.Fields!Cantidad = 1
        mrs_Cpt.Fields!precio = 1
        mrs_Cpt.Fields!Importe = 1
        mrs_Cpt.Fields!IdPuntoCarga = sghPtoCargaEcogObstetrica
        mrs_Cpt.Fields!idEspecialidadServicio = DevuelveIdEspecialidadServicioSegunPuntoCarga(sghPtoCargaEcogObstetrica)
        mrs_Cpt.Update
        SumaTotales
    End If
End Sub

Private Sub btnAddPatolClin_Click()
    If Me.lblPatolClinica.Caption = "" Then
       MsgBox "Tiene que elegir un Procedimiento de Patología Clínica", vbInformation, Me.Caption
       Exit Sub
    End If
    Dim lbAgregar As Boolean
    lbAgregar = True
    If mrs_Cpt.RecordCount > 0 Then
       mrs_Cpt.MoveFirst
       mrs_Cpt.Find "idProducto=" & lnIdPatolClinica
       If Not mrs_Cpt.EOF Then
          MsgBox "Ya está registrado", vbInformation, Me.Caption
          lbAgregar = False
       End If
    End If
    If lbAgregar = True Then
        mrs_Cpt.AddNew
        mrs_Cpt.Fields!idProducto = lnIdPatolClinica
        mrs_Cpt.Fields!Producto = lblPatolClinica.Caption
        mrs_Cpt.Fields!Cantidad = 1
        mrs_Cpt.Fields!precio = 1
        mrs_Cpt.Fields!Importe = 1
        mrs_Cpt.Fields!IdPuntoCarga = sghPtoCargaPatologiaClinica
        mrs_Cpt.Fields!idEspecialidadServicio = DevuelveIdEspecialidadServicioSegunPuntoCarga(sghPtoCargaPatologiaClinica)
        mrs_Cpt.Update
        SumaTotales
    End If
End Sub

Private Sub btnAddRayos_Click()
    If Me.lblRayosX.Caption = "" Then
       MsgBox "Tiene que elegir un Procedimiento de Rayos X", vbInformation, Me.Caption
       Exit Sub
    End If
    Dim lbAgregar As Boolean
    lbAgregar = True
    If mrs_Cpt.RecordCount > 0 Then
       mrs_Cpt.MoveFirst
       mrs_Cpt.Find "idProducto=" & lnIdRayosX
       If Not mrs_Cpt.EOF Then
          MsgBox "Ya está registrado", vbInformation, Me.Caption
          lbAgregar = False
       End If
    End If
    If lbAgregar = True Then
        mrs_Cpt.AddNew
        mrs_Cpt.Fields!idProducto = lnIdRayosX
        mrs_Cpt.Fields!Producto = lblRayosX.Caption
        mrs_Cpt.Fields!Cantidad = 1
        mrs_Cpt.Fields!precio = 1
        mrs_Cpt.Fields!Importe = 1
        mrs_Cpt.Fields!IdPuntoCarga = sghPtoCargaRayosX
        mrs_Cpt.Fields!idEspecialidadServicio = DevuelveIdEspecialidadServicioSegunPuntoCarga(sghPtoCargaRayosX)
        mrs_Cpt.Update
        SumaTotales
    End If
End Sub

Private Sub btnAddTomografia_Click()
    If Me.lblTomografia.Caption = "" Then
       MsgBox "Tiene que elegir un Procedimiento de Tomografía", vbInformation, Me.Caption
       Exit Sub
    End If
    Dim lbAgregar As Boolean
    lbAgregar = True
    If mrs_Cpt.RecordCount > 0 Then
       mrs_Cpt.MoveFirst
       mrs_Cpt.Find "idProducto=" & lnIdTomografia
       If Not mrs_Cpt.EOF Then
          MsgBox "Ya está registrado", vbInformation, Me.Caption
          lbAgregar = False
       End If
    End If
    If lbAgregar = True Then
        mrs_Cpt.AddNew
        mrs_Cpt.Fields!idProducto = lnIdTomografia
        mrs_Cpt.Fields!Producto = lblTomografia.Caption
        mrs_Cpt.Fields!Cantidad = 1
        mrs_Cpt.Fields!precio = 1
        mrs_Cpt.Fields!Importe = 1
        mrs_Cpt.Fields!IdPuntoCarga = sghPtoCargaTomografia
        mrs_Cpt.Fields!idEspecialidadServicio = DevuelveIdEspecialidadServicioSegunPuntoCarga(sghPtoCargaTomografia)
        mrs_Cpt.Update
        SumaTotales
    End If
End Sub

Private Sub btnAdministrativo_Click()
    If Me.lblAdministrativo.Caption = "" Then
       MsgBox "Tiene que elegir un Procedimiento Administrativo", vbInformation, Me.Caption
       Exit Sub
    End If
    Dim lbAgregar As Boolean
    lbAgregar = True
    If mrs_Cpt.RecordCount > 0 Then
       mrs_Cpt.MoveFirst
       mrs_Cpt.Find "idProducto=" & lnIdAdministrativo
       If Not mrs_Cpt.EOF Then
          MsgBox "Ya está registrado", vbInformation, Me.Caption
          lbAgregar = False
       End If
    End If
    If lbAgregar = True Then
        mrs_Cpt.AddNew
        mrs_Cpt.Fields!idProducto = lnIdAdministrativo
        mrs_Cpt.Fields!Producto = Me.lblAdministrativo.Caption
        mrs_Cpt.Fields!Cantidad = 1
        mrs_Cpt.Fields!precio = 1
        mrs_Cpt.Fields!Importe = 1
        mrs_Cpt.Fields!IdPuntoCarga = sghPtoCargaCaja
        mrs_Cpt.Fields!idEspecialidadServicio = Val(lcBuscaParametro.SeleccionaFilaParametro(262))      'secretaria, por ahora
        mrs_Cpt.Update
        SumaTotales
    End If

End Sub

Private Sub btnCita_Click()
    If Me.lblEspecialidad.Caption = "" Then
       MsgBox "Tiene que elegir una Especialidad para la CITA", vbInformation, Me.Caption
       Exit Sub
    End If
    Dim lbAgregar As Boolean
    Dim lcProducto As String
    lcProducto = lcPagoCE & " (" & lblEspecialidad.Caption & ")"
    lbAgregar = True
    If mrs_Cpt.RecordCount > 0 Then
       mrs_Cpt.MoveFirst
       mrs_Cpt.Find "Producto='" & lcProducto & "'"
       If Not mrs_Cpt.EOF Then
          Do While Not mrs_Cpt.EOF
             If mrs_Cpt.Fields!idProducto = lnIdPagoCE And mrs_Cpt.Fields!idEspecialidadServicio = lnIdEspecialidadCE And mrs_Cpt.Fields!Producto = lcProducto Then
                MsgBox "Ya está registrado", vbInformation, Me.Caption
                lbAgregar = False
                Exit Do
             End If
             mrs_Cpt.MoveNext
          Loop
       End If
    End If
    If lbAgregar = True Then
        mrs_Cpt.AddNew
        mrs_Cpt.Fields!idProducto = lnIdPagoCE
        mrs_Cpt.Fields!Producto = lcProducto     'lcPagoCE & "  (" & lblEspecialidad.Caption & ")"
        mrs_Cpt.Fields!Cantidad = 1
        mrs_Cpt.Fields!precio = 1
        mrs_Cpt.Fields!Importe = 1
        mrs_Cpt.Fields!IdPuntoCarga = sghPtoCargaAdmisionCE
        mrs_Cpt.Fields!idEspecialidadServicio = lnIdEspecialidadCE
        mrs_Cpt.Update
        SumaTotales
    End If

End Sub

Private Sub btnFarmacia_Click()
    If Me.lblFarmacia.Caption = "" Then
       MsgBox "Tiene que elegir un Medicamento o Insumo de Farmacia", vbInformation, Me.Caption
       Exit Sub
    End If
    Dim lbAgregar As Boolean
    lbAgregar = True
    If mrs_Cpt.RecordCount > 0 Then
       mrs_Cpt.MoveFirst
       mrs_Cpt.Find "idProducto=" & lnIdFarmacia
       If Not mrs_Cpt.EOF Then
          Do While Not mrs_Cpt.EOF
            If mrs_Cpt.Fields!IdPuntoCarga = sghPtoCargaFarmacia And mrs_Cpt.Fields!idProducto = lnIdFarmacia Then
                MsgBox "Ya está registrado", vbInformation, Me.Caption
                lbAgregar = False
                Exit Do
            End If
          Loop
       End If
    End If
    If lbAgregar = True Then
        mrs_Cpt.AddNew
        mrs_Cpt.Fields!idProducto = lnIdFarmacia
        mrs_Cpt.Fields!Producto = Me.lblFarmacia.Caption
        mrs_Cpt.Fields!Cantidad = 1
        mrs_Cpt.Fields!precio = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(lnIdFarmacia, sghPrecioVentaContado)
        mrs_Cpt.Fields!Importe = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(lnIdFarmacia, sghPrecioVentaContado)
        mrs_Cpt.Fields!IdPuntoCarga = sghPtoCargaFarmacia
        mrs_Cpt.Fields!idEspecialidadServicio = Val(lcBuscaParametro.SeleccionaFilaParametro(261))
        mrs_Cpt.Update
        SumaTotales
    End If
   
End Sub

Private Sub btnQuitar_Click()
    On Error Resume Next
    mrs_Cpt.Delete
    mrs_Cpt.Update
    SumaTotales
End Sub
'debb-08/11/2016
Private Sub chkItemFarmacia_Click()
    cmdCopia.Visible = False
    If chkItemFarmacia.Value = 1 Then
       mo_Formulario.HabilitarDeshabilitar txtCodigo, False
       If mi_Opcion = sghAgregar Then
          Me.txtCodigo.Text = mo_ReglasFarmacia.CatalogoDIGEMIDAsignaCodigoPaquete
          cmdCopia.Visible = True
       End If
    Else
       mo_Formulario.HabilitarDeshabilitar txtCodigo, True
    End If
End Sub

Private Sub cmdAdministrativo_Click()
    'Dim oPaquetesBuscar As New FacCatalogoPqteBuscar
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    oPaquetesBuscar.IdPuntoCarga = 1500
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
       lnIdAdministrativo = oPaquetesBuscar.idProducto
       Me.lblAdministrativo.Caption = oPaquetesBuscar.Producto
       btnAdministrativo.SetFocus
    End If
    Set oPaquetesBuscar = Nothing

End Sub

Private Sub cmdCita_Click()
    'Dim oPaquetesBuscar As New FacCatalogoPqteBuscar
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    oPaquetesBuscar.IdPuntoCarga = 0
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
       lnIdEspecialidadCE = oPaquetesBuscar.idProducto
       lblEspecialidad.Caption = oPaquetesBuscar.Producto
       btnCita.SetFocus
    End If
    Set oPaquetesBuscar = Nothing

End Sub



Private Sub cmdCopia_Click()
    If mrs_Cpt.RecordCount > 0 Then
       MsgBox "Solo es usado cuando no se tienen ITEMS registrados", vbInformation, Me.Caption
       Exit Sub
    End If
    Dim oPaquetesBuscar As New SIGHNegocios.BuscaPaquetes
    Dim lnIdFactPaquete1 As Long
    Dim oRsTmp As New Recordset
    oPaquetesBuscar.DebeConsiderarPaquete = sghTipoPaqueteSolofarmacia
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
       lnIdFactPaquete1 = oPaquetesBuscar.idFactPaquete
       Set oRsTmp = mo_ReglasFacturacion.FactCatalogoPaqueteSeleccionarPorFiltro(" idFactPaquete=" & lnIdFactPaquete1)
       If oRsTmp.RecordCount > 0 Then
          txtDescripcion.Text = Trim(oRsTmp.Fields!Descripcion) & "-" & Right(txtCodigo.Text, 3)
       End If
       oRsTmp.Close
       Set oRsTmp = mo_ReglasFacturacion.FacturacionCatalogoPaqueteFarmSeleccionarXid(lnIdFactPaquete1)
       If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
              mrs_Cpt.AddNew
              mrs_Cpt!idProducto = oRsTmp!idProducto
              mrs_Cpt!Producto = oRsTmp.Fields!codigo & "//" & oRsTmp!nombre
              mrs_Cpt!Cantidad = oRsTmp!Cantidad
              mrs_Cpt!precio = oRsTmp!precio
              mrs_Cpt!Importe = oRsTmp!Importe
              mrs_Cpt!IdPuntoCarga = sghPtoCargaFarmacia
              mrs_Cpt!idEspecialidadServicio = Val(lcBuscaParametro.SeleccionaFilaParametro(261))
              mrs_Cpt.Update
              oRsTmp.MoveNext
           Loop
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
       If mrs_Cpt.RecordCount > 0 Then
           SumaTotales
           mrs_Cpt.MoveFirst
       End If
    End If
    Set oPaquetesBuscar = Nothing
    Set oRsTmp = Nothing
End Sub

Private Sub cmdEcogGeneral_Click()
    'Dim oPaquetesBuscar As New FacCatalogoPqteBuscar
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    oPaquetesBuscar.IdPuntoCarga = sghPtoCargaEcogGeneral
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
       lnIdEcogGen = oPaquetesBuscar.idProducto
       lblEcogGeneral.Caption = oPaquetesBuscar.Producto
       btnAddEcogGeneral.SetFocus
    End If
    Set oPaquetesBuscar = Nothing
End Sub

Private Sub cmdEcogObs_Click()
    'Dim oPaquetesBuscar As New FacCatalogoPqteBuscar
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    oPaquetesBuscar.IdPuntoCarga = sghPtoCargaEcogObstetrica
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
       lnIdEcogObst = oPaquetesBuscar.idProducto
       lblEcogObst.Caption = oPaquetesBuscar.Producto
       btnAddEcogObst.SetFocus
    End If
    Set oPaquetesBuscar = Nothing

End Sub

Private Sub cmdFarmacia_Click()
    'Dim oPaquetesBuscar As New FacCatalogoPqteBuscar
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    oPaquetesBuscar.IdPuntoCarga = sghPtoCargaFarmacia
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
       lnIdFarmacia = oPaquetesBuscar.idProducto
       Me.lblFarmacia.Caption = oPaquetesBuscar.Producto
       btnFarmacia.SetFocus
    End If
    Set oPaquetesBuscar = Nothing
End Sub

Private Sub cmdPatolClinica_Click()
Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    oPaquetesBuscar.IdPuntoCarga = sghPtoCargaPatologiaClinica
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
       lnIdPatolClinica = oPaquetesBuscar.idProducto
       lblPatolClinica.Caption = oPaquetesBuscar.Producto
       btnAddPatolClin.SetFocus
    End If
    Set oPaquetesBuscar = Nothing
End Sub

Private Sub cmdRefrescaTotal_Click()
       SumaTotales
End Sub

Private Sub cmdRx_Click()
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    oPaquetesBuscar.IdPuntoCarga = sghPtoCargaRayosX
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
       lnIdRayosX = oPaquetesBuscar.idProducto
       lblRayosX.Caption = oPaquetesBuscar.Producto
       btnAddRayos.SetFocus
    End If
    Set oPaquetesBuscar = Nothing

End Sub

Private Sub cmdTomografia_Click()
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    oPaquetesBuscar.IdPuntoCarga = sghPtoCargaTomografia
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
       lnIdTomografia = oPaquetesBuscar.idProducto
       lblTomografia.Caption = oPaquetesBuscar.Producto
       btnAddTomografia.SetFocus
    End If
    Set oPaquetesBuscar = Nothing

End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       CreaTemporal
       DevuelveIdPagoCE
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Paquete"
       Case sghModificar
           Me.Caption = "Modificar Paquete"
       Case sghConsultar
           Me.Caption = "Consultar Paquete"
       Case sghEliminar
           Me.Caption = "Eliminar Paquete"
       End Select
       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()

If wxFranklin = "*" Then Exit Sub

   
   
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                   If ActualizaItemFarmacia = False Then
                      MsgBox "Problemas para AGREGAR/MODIFICAR en CATALOGO DE FARMACIA", vbInformation, Me.Caption
                   End If
                   Me.Visible = False
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                   If ActualizaItemFarmacia = False Then
                      MsgBox "Problemas para AGREGAR/MODIFICAR en CATALOGO DE FARMACIA", vbInformation, Me.Caption
                   End If
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If ActualizaItemFarmacia = True Then
                    If EliminarDatos() Then
                        
                        MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                        Me.Visible = False
                    Else
                        MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                    End If
               Else
                    MsgBox "Problemas para ElIMINAR en CATALOGO DE FARMACIA", vbInformation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   SumaTotales
   If Trim(Me.txtCodigo) = "" Then
       sMensaje = sMensaje + "Ingrese el código" + Chr(13)
   End If
   If Trim(Me.txtDescripcion) = "" Then
       sMensaje = sMensaje + "Ingrese el nombre" + Chr(13)
   End If
   If mrs_Cpt.RecordCount <= 0 Then
       sMensaje = sMensaje + "Al menos debe tener 1 detalle" + Chr(13)
   End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   Dim sMensaje As String, lcTipo As String
   Dim oRsTmp As New Recordset
   If chkItemFarmacia.Value = 1 And lnIdTipoPaquete <> 2 Then
      MsgBox "Al Paquete de FARMACIA no se puede añadir items de CPT", vbInformation, Me.Caption
      Exit Function
   End If
   '
   sMensaje = ""
'   lnIdTipoSalidaBienInsumo = 0
'   If chkItemFarmacia.Value = 1 Then
'      If mrs_Cpt.RecordCount > 0 Then
'         mrs_Cpt.MoveFirst
'         Set oRsTmp = mo_ReglasFacturacion.CatalogoBienesSeleccionarPorIdProducto(mrs_Cpt!IdProducto)
'         lnIdTipoSalidaBienInsumo = oRsTmp!idTipoSalidaBienInsumo
'         lcTipo = oRsTmp!tipo
'         If lnIdTipoSalidaBienInsumo <> 3 Then
'            Do While Not mrs_Cpt.EOF
'               Set oRsTmp = mo_ReglasFacturacion.CatalogoBienesSeleccionarPorIdProducto(mrs_Cpt!IdProducto)
'               If Not (oRsTmp!idTipoSalidaBienInsumo = 3 Or oRsTmp!idTipoSalidaBienInsumo = lnIdTipoSalidaBienInsumo) Then
'                  sMensaje = sMensaje & "El item " & Trim(mrs_Cpt!Producto) & " tiene TIPO_SALIDA diferente" & Chr(13)
'               End If
'               mrs_Cpt.MoveNext
'            Loop
'         End If
'      End If
'   End If
'   If sMensaje <> "" Then
'      MsgBox "Problemas con TIPO SALIDA " & lcTipo & " en el PAQUETE DE FARMACIA para: " & Chr(13) & sMensaje, vbInformation, Me.Caption
'      Exit Function
'   End If
   '
   ValidarReglas = True
   Set oRsTmp = Nothing
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oRsTmp As New ADODB.Recordset
    Dim oFactCatalogoPaquete As New FactCatalogoPaquete, oDOFactCatalogoPaquete As New DOFactCatalogoPaquete
    Dim oFacturacionCatalogoPqte As New FacturacionCatalogoPqte, oDOFacturacionCatalogoPqtes As New DOFacturacionCatalogoPqtes
    On Error GoTo ErrAgregar
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oFactCatalogoPaquete.Conexion = oConexion
    Set oFacturacionCatalogoPqte.Conexion = oConexion
    ms_MensajeError = ""
    '
    Set oRsTmp = mo_ReglasFacturacion.FactCatalogoPaqueteSeleccionarPorFiltro("")
    
    oDOFactCatalogoPaquete.codigo = Me.txtCodigo.Text
    oDOFactCatalogoPaquete.Descripcion = Me.txtDescripcion.Text
    oDOFactCatalogoPaquete.IdTipoFinanciamiento = 1
    oDOFactCatalogoPaquete.fechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
    oDOFactCatalogoPaquete.idUsuario = ml_idUsuario
    oDOFactCatalogoPaquete.IdEstado = IIf(Me.chkEstado.Value = 1, True, False)
    oDOFactCatalogoPaquete.TipoPaquete = lnIdTipoPaquete
    oDOFactCatalogoPaquete.IdUsuarioAuditoria = ml_idUsuario
    oDOFactCatalogoPaquete.esItemFarmacia = Me.chkItemFarmacia.Value
    oDOFactCatalogoPaquete.cpt = Me.txtCPT.Text
    If Not oFactCatalogoPaquete.Insertar(oDOFactCatalogoPaquete) Then
       MsgBox oFactCatalogoPaquete.MensajeError: GoTo ErrAgregar
    End If
    ml_IdFactPaquete = oDOFactCatalogoPaquete.idFactPaquete
    '
    Set oRsTmp = mo_ReglasFacturacion.FacturacionCatalogoPaquetesSeleccionarTodos
    mrs_Cpt.MoveFirst
    Do While Not mrs_Cpt.EOF
        oDOFacturacionCatalogoPqtes.idFactPaquete = ml_IdFactPaquete
        oDOFacturacionCatalogoPqtes.IdPuntoCarga = mrs_Cpt.Fields!IdPuntoCarga
        oDOFacturacionCatalogoPqtes.idEspecialidadServicio = mrs_Cpt.Fields!idEspecialidadServicio
        oDOFacturacionCatalogoPqtes.idProducto = mrs_Cpt.Fields!idProducto
        oDOFacturacionCatalogoPqtes.Cantidad = mrs_Cpt.Fields!Cantidad
        oDOFacturacionCatalogoPqtes.precio = mrs_Cpt.Fields!precio
        oDOFacturacionCatalogoPqtes.Importe = mrs_Cpt.Fields!Importe
        oDOFacturacionCatalogoPqtes.IdUsuarioAuditoria = ml_idUsuario
        If Not oFacturacionCatalogoPqte.Insertar(oDOFacturacionCatalogoPqtes) Then
           MsgBox oFacturacionCatalogoPqte.MensajeError: GoTo ErrAgregar
        End If
        mrs_Cpt.MoveNext
    Loop
    oRsTmp.Close
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", ml_IdFactPaquete, "FactCatalogoPaquete", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Paquete N° " & Me.txtCodigo.Text)
    '
    oConexion.CommitTrans
    AgregarDatos = True
ErrAgregar:
    If AgregarDatos = False Then
    'Resume
       oConexion.RollbackTrans
       ms_MensajeError = Err.Description
    End If
    oConexion.Close
    Set oRsTmp = Nothing
    Set oConexion = Nothing
    Set oFactCatalogoPaquete = Nothing
    Set oDOFactCatalogoPaquete = Nothing
    Set oFacturacionCatalogoPqte = Nothing
    Set oDOFacturacionCatalogoPqtes = Nothing
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oRsTmp As New ADODB.Recordset
    Dim oFactCatalogoPaquete As New FactCatalogoPaquete, oDOFactCatalogoPaquete As New DOFactCatalogoPaquete
    Dim oFacturacionCatalogoPqte As New FacturacionCatalogoPqte, oDOFacturacionCatalogoPqtes As New DOFacturacionCatalogoPqtes
    On Error GoTo ErrModificar
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    Set oRsTmp = mo_ReglasFacturacion.FactCatalogoPaqueteSeleccionarPorFiltro("idFactPaquete=" & Trim(str(ml_IdFactPaquete)))
    oConexion.BeginTrans
    Set oFactCatalogoPaquete.Conexion = oConexion
    Set oFacturacionCatalogoPqte.Conexion = oConexion
    ms_MensajeError = ""
    '
    
    If oRsTmp.RecordCount > 0 Then
        oDOFactCatalogoPaquete.idFactPaquete = ml_IdFactPaquete
        oDOFactCatalogoPaquete.codigo = Me.txtCodigo.Text
        oDOFactCatalogoPaquete.Descripcion = Me.txtDescripcion.Text
        oDOFactCatalogoPaquete.IdTipoFinanciamiento = 1
        oDOFactCatalogoPaquete.fechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
        oDOFactCatalogoPaquete.idUsuario = ml_idUsuario
        oDOFactCatalogoPaquete.IdEstado = IIf(Me.chkEstado.Value = 1, True, False)
        oDOFactCatalogoPaquete.TipoPaquete = lnIdTipoPaquete
        oDOFactCatalogoPaquete.IdUsuarioAuditoria = ml_idUsuario
        oDOFactCatalogoPaquete.esItemFarmacia = Me.chkItemFarmacia.Value
        oDOFactCatalogoPaquete.cpt = Me.txtCPT.Text
        If Not oFactCatalogoPaquete.Modificar(oDOFactCatalogoPaquete) Then
           MsgBox oFactCatalogoPaquete.MensajeError: GoTo ErrModificar
        End If
    End If
    oRsTmp.Close
    '
    mo_ReglasFacturacion.FacturacionCatalogoPaquetesEliminarXidFactPaquete ml_IdFactPaquete
    Set oRsTmp = mo_ReglasFacturacion.FacturacionCatalogoPaquetesSeleccionarTodos
    mrs_Cpt.MoveFirst
    Do While Not mrs_Cpt.EOF

        oDOFacturacionCatalogoPqtes.idFactPaquete = ml_IdFactPaquete
        oDOFacturacionCatalogoPqtes.IdPuntoCarga = mrs_Cpt.Fields!IdPuntoCarga
        oDOFacturacionCatalogoPqtes.idEspecialidadServicio = mrs_Cpt.Fields!idEspecialidadServicio
        oDOFacturacionCatalogoPqtes.idProducto = mrs_Cpt.Fields!idProducto
        oDOFacturacionCatalogoPqtes.Cantidad = mrs_Cpt.Fields!Cantidad
        oDOFacturacionCatalogoPqtes.precio = mrs_Cpt.Fields!precio
        oDOFacturacionCatalogoPqtes.Importe = mrs_Cpt.Fields!Importe
        oDOFacturacionCatalogoPqtes.IdUsuarioAuditoria = ml_idUsuario
        If Not oFacturacionCatalogoPqte.Insertar(oDOFacturacionCatalogoPqtes) Then
           MsgBox oFacturacionCatalogoPqte.MensajeError: GoTo ErrModificar
        End If
        mrs_Cpt.MoveNext
    Loop
    oRsTmp.Close
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "M", ml_IdFactPaquete, "FactCatalogoPaquete", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Paquete N° " & Me.txtCodigo.Text)
    '
    oConexion.CommitTrans
    ModificarDatos = True
ErrModificar:
    If ModificarDatos = False Then
       
       oConexion.RollbackTrans
       ms_MensajeError = Err.Description
    End If
    oConexion.Close
    Set oRsTmp = Nothing
    Set oConexion = Nothing
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oRsTmp As New ADODB.Recordset
    On Error GoTo ErrAnular
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    ms_MensajeError = ""
    '
    Set oRsTmp = mo_ReglasFacturacion.FacturacionPaquetesSeleccionarPorFiltro("idFactPaquete=" & Trim(str(ml_IdFactPaquete)))
    If oRsTmp.RecordCount > 0 Then
       MsgBox "Ya existe MOVIMIENTOS PARA ESE PAQUETE", vbInformation, Me.Caption
    Else
        oRsTmp.Close
        mo_ReglasFacturacion.FacturacionCatalogoPaquetesEliminarXidFactPaquete ml_IdFactPaquete
        '
         mo_ReglasFacturacion.FactCatalogoPaqueteEliminarXidFactPaquete ml_IdFactPaquete
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "E", ml_IdFactPaquete, "FactCatalogoPaquete", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Paquete N° " & Me.txtCodigo.Text)
    '
    oConexion.CommitTrans
    EliminarDatos = True
ErrAnular:
    If EliminarDatos = False Then
       oConexion.RollbackTrans
       ms_MensajeError = Err.Description
    End If
    oConexion.Close
    Set oRsTmp = Nothing
    Set oConexion = Nothing
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
        
    
    Dim oRsTmp As New ADODB.Recordset
    Dim oRsTmp1 As New ADODB.Recordset
    Dim lcEspecialidad As String, lcDproducto As String
    '
    Set oRsTmp = mo_ReglasFacturacion.FactCatalogoPaqueteSeleccionarPorFiltro("idFactPaquete=" & Trim(str(ml_IdFactPaquete)))
    If oRsTmp.RecordCount > 0 Then
        txtCodigo.Text = oRsTmp.Fields!codigo
        txtDescripcion.Text = oRsTmp.Fields!Descripcion
        chkEstado.Value = IIf(oRsTmp.Fields!IdEstado = 0, 0, 1)
        Me.chkItemFarmacia.Value = IIf(IsNull(oRsTmp!esItemFarmacia), 0, oRsTmp!esItemFarmacia)
        Me.txtCPT.Text = IIf(IsNull(oRsTmp!cpt), "", oRsTmp!cpt)
        If Me.chkItemFarmacia.Value = 1 And (mi_Opcion = sghModificar Or mi_Opcion = sghEliminar) Then
            oRsTmp.Close
            Set oRsTmp = mo_ReglasFarmacia.FarmMovimientoDetalleSeleccionarXcodigo(txtCodigo.Text)
            If oRsTmp.RecordCount > 0 Then
               MsgBox "NO PODRA MODIFICAR/ELIMINAR PORQUE YA EXISTE UN MOVIMIENTO: " & oRsTmp!movNumero & " EN FARMACIA", vbInformation, Me.Caption
               Me.btnAceptar.Enabled = False
            End If
        End If
    End If
    oRsTmp.Close
    '
    Set oRsTmp = mo_ReglasFacturacion.FacturacionCatalogoPaquetesSeleccionarPorIdFactPaquete(ml_IdFactPaquete)
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
            lcEspecialidad = ""
            If lnIdPagoCE = oRsTmp.Fields!idProducto Then
               Set oRsTmp1 = mo_AdminComun.EspecialidadesSeleccionarPorFiltro("idEspecialidad=" & Trim(str(oRsTmp.Fields!idEspecialidadServicio)))
               If oRsTmp1.RecordCount > 0 Then
                  lcEspecialidad = "   (" & Trim(oRsTmp1.Fields!nombre) & ")"
               End If
               oRsTmp1.Close
            End If
            If oRsTmp.Fields!IdPuntoCarga = sghPtoCargaFarmacia Then
               Set oRsTmp1 = mo_ReglasFacturacion.CatalogoBienesSeleccionarPorIdProducto(oRsTmp.Fields!idProducto)
               If oRsTmp1.RecordCount > 0 Then
                  lcDproducto = Trim(oRsTmp1.Fields!codigo) & " " & oRsTmp1.Fields!nombre
               End If
               oRsTmp1.Close
            Else
'               lcDproducto = lcPagoCE & lcEspecialidad  'Trim(oRsTmp.Fields!Codigo) & " " & oRsTmp.Fields!Nombre & lcEspecialidad
                lcDproducto = Trim(oRsTmp.Fields!codigo) & " " & oRsTmp.Fields!nombre & lcEspecialidad 'Actualizado 07102014
            End If
            mrs_Cpt.AddNew
            mrs_Cpt.Fields!idProducto = oRsTmp.Fields!idProducto
            mrs_Cpt.Fields!Producto = lcDproducto
            mrs_Cpt.Fields!Cantidad = oRsTmp.Fields!Cantidad
            mrs_Cpt.Fields!precio = oRsTmp.Fields!precio
            mrs_Cpt.Fields!Importe = oRsTmp.Fields!Importe
            mrs_Cpt.Fields!IdPuntoCarga = oRsTmp.Fields!IdPuntoCarga
            mrs_Cpt.Fields!idEspecialidadServicio = oRsTmp.Fields!idEspecialidadServicio
            mrs_Cpt.Update
            oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    SumaTotales
    mb_ExistenDatos = True
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

    
    Me.txtDescripcion = ""
    Me.txtCodigo = ""
    
End Sub

Sub CargarComboBoxes()
    
End Sub




Private Sub grdCpt_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
'   SumaTotales
End Sub

Sub SumaTotales()
    Dim lnTotal As Double
    Dim lbHayServicios As Boolean, lbHayFarmacia As Boolean
    lbHayFarmacia = False: lbHayServicios = False
    lnTotal = 0
    
    If mrs_Cpt.RecordCount > 0 Then
       mrs_Cpt.MoveFirst
       Do While Not mrs_Cpt.EOF
          If mrs_Cpt.Fields!IdPuntoCarga = sghPtoCargaFarmacia Then
             lbHayFarmacia = True
          Else
             lbHayServicios = True
          End If
          mrs_Cpt.Fields!Importe = Round(mrs_Cpt.Fields!Cantidad * mrs_Cpt.Fields!precio, 2)
          mrs_Cpt.Update
          lnTotal = lnTotal + mrs_Cpt.Fields!Importe
          mrs_Cpt.MoveNext
       Loop
    End If
    lblTotal.Caption = Format(lnTotal, "####,###,##0.00")
    Set Me.grdCpt.DataSource = mrs_Cpt
    '
    If lbHayServicios = True And lbHayFarmacia = True Then
       lnIdTipoPaquete = 3
       lblTipo.Caption = "Es un Paquete con items de FARMACIA y CPT"
    ElseIf lbHayServicios = True Then
       lnIdTipoPaquete = 1
       lblTipo.Caption = "Es un Paquete solo con items de  CPT"
    Else
       lnIdTipoPaquete = 2
       lblTipo.Caption = "Es un Paquete solo con items de FARMACIA"
    End If
    '
End Sub

Private Sub grdCpt_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdCpt.Bands(0).Columns("IdProducto").Hidden = True
    grdCpt.Bands(0).Columns("idPuntoCarga").Hidden = True
    grdCpt.Bands(0).Columns("idTipoSalidaBienInsumo").Hidden = True
    grdCpt.Bands(0).Columns("idEspecialidadServicio").Hidden = True
    grdCpt.Bands(0).Columns("Producto").Header.Caption = "Procedimiento"
    grdCpt.Bands(0).Columns("Producto").Width = 6600
    grdCpt.Bands(0).Columns("Producto").Activation = ssActivationActivateNoEdit
    grdCpt.Bands(0).Columns("cantidad").Width = 500
    grdCpt.Bands(0).Columns("cantidad").Format = "#0"
    grdCpt.Bands(0).Columns("cantidad").Header.Appearance.ForeColor = vbWhite
    grdCpt.Bands(0).Columns("cantidad").Header.Appearance.BackColor = vbRed
    grdCpt.Bands(0).Columns("Precio").Width = 900
    'grdCpt.Bands(0).Columns("Precio").Activation = ssActivationActivateNoEdit
    grdCpt.Bands(0).Columns("Precio").Format = "#0.000"
    grdCpt.Bands(0).Columns("Precio").Header.Appearance.ForeColor = vbWhite
    grdCpt.Bands(0).Columns("Precio").Header.Appearance.BackColor = vbRed
    grdCpt.Bands(0).Columns("Importe").Width = 900
    grdCpt.Bands(0).Columns("Importe").Format = "#0.000"
    grdCpt.Bands(0).Columns("Importe").Activation = ssActivationActivateNoEdit
End Sub

Sub CreaTemporal()
    With mrs_Cpt
          .Fields.Append "IdProducto", adInteger, 4, adFldIsNullable
          .Fields.Append "Producto", adVarChar, 249, adFldIsNullable
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Precio", adDouble
          .Fields.Append "Importe", adDouble
          .Fields.Append "idPuntoCarga", adInteger
          .Fields.Append "idEspecialidadServicio", adInteger
          .Fields.Append "idTipoSalidaBienInsumo", adInteger
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    mo_Apariencia.ConfigurarFilasBiColores Me.grdCpt, sighentidades.GrillaConFilasBicolor
    Set Me.grdCpt.DataSource = mrs_Cpt
    
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If

End Sub


Function DevuelveIdEspecialidadServicioSegunPuntoCarga(lnIdPuntoCarga As Long) As Long
    Dim oRsTmp As New ADODB.Recordset
    Dim lcSql As String
    Set oRsTmp = mo_AdminComun.FactPuntosCargaSeleccionarPorIdPuntoCarga(lnIdPuntoCarga)
    DevuelveIdEspecialidadServicioSegunPuntoCarga = 0
    If oRsTmp.RecordCount > 0 Then
       DevuelveIdEspecialidadServicioSegunPuntoCarga = oRsTmp.Fields!IdEspecialidad   ' oRsTmp.Fields!IdPuntoCarga
    End If
    Set oRsTmp = Nothing
End Function

Sub DevuelveIdPagoCE()
    Dim oRsTmp As New ADODB.Recordset
    Dim lcSql As String
    lnIdPagoCE = Val(lcBuscaParametro.SeleccionaFilaParametro(257))
    Set oRsTmp = mo_ReglasFacturacion.CatalogoServiciosSeleccionarPorIdProductodebb(lnIdPagoCE)
    lcPagoCE = ""
    If oRsTmp.RecordCount > 0 Then
       lcPagoCE = oRsTmp.Fields!codigo & " " & oRsTmp.Fields!nombre
    End If
    Set oRsTmp = Nothing
End Sub

'debb-08/11/2016
Function ActualizaItemFarmacia() As Boolean
    ActualizaItemFarmacia = False
    If Me.chkItemFarmacia.Value = 1 Then
       Dim oDOCatalogoBienesInsumos As New DOCatalogoBienesInsumos
       Dim oCatalogoBienesInsumos As New CatalogoBienesInsumos
       Dim oRsTmp1 As New Recordset
       Dim oConexion As New Connection
       Dim lnPrecioDistribucion As Double, lnPrecioCompra As Double, lnPrecioVenta As Double, lnIdProducto As Long
       
       lnPrecioVenta = CDbl(lblTotal.Caption)
       lnPrecioDistribucion = lnPrecioVenta  '  Round((100 * lnPrecioVenta) / (100 + CDbl(lcBuscaParametro.SeleccionaFilaParametro(308))), 2)
       lnPrecioCompra = lnPrecioVenta ' Round((100 * lnPrecioDistribucion) / (100 + CDbl(lcBuscaParametro.SeleccionaFilaParametro(307))), 2)

       oConexion.CommandTimeout = 900
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighentidades.CadenaConexion
       Set oRsTmp1 = oCatalogoBienesInsumos.SeleccionarPorCodigo(Me.txtCodigo.Text, oConexion)
       Set oCatalogoBienesInsumos.Conexion = oConexion
       If oRsTmp1.RecordCount = 0 Then
            If mi_Opcion = sghAgregar Or mi_Opcion = sghModificar Then
               oDOCatalogoBienesInsumos.codigo = Me.txtCodigo.Text
               oDOCatalogoBienesInsumos.IdCentroCosto = 999
               oDOCatalogoBienesInsumos.IdGrupoFarmacologico = 999
               oDOCatalogoBienesInsumos.IdPartida = 1
               oDOCatalogoBienesInsumos.idTipoSalidaBienInsumo = sghTipoSalidaItemFarmacia.sghSoloVenta     'lnIdTipoSalidaBienInsumo
               oDOCatalogoBienesInsumos.IdUsuarioAuditoria = ml_idUsuario
               oDOCatalogoBienesInsumos.nombre = Me.txtDescripcion.Text
               oDOCatalogoBienesInsumos.TipoProducto = 0
               oDOCatalogoBienesInsumos.TipoProductoSismed = "_"
               oDOCatalogoBienesInsumos.IdSubGrupoFarmacologico = 999
               oDOCatalogoBienesInsumos.FormaFarmaceutica = "UNI"
               oDOCatalogoBienesInsumos.PrecioCompra = lnPrecioCompra
               oDOCatalogoBienesInsumos.PrecioDistribucion = lnPrecioDistribucion
               oDOCatalogoBienesInsumos.PrecioDonacion = lnPrecioCompra
               If oCatalogoBienesInsumos.Insertar(oDOCatalogoBienesInsumos) = True Then
                  ActualizaItemFarmacia = True
                  lnIdProducto = oDOCatalogoBienesInsumos.idProducto
                  mo_ReglasFacturacion.FactCatalogoBienesInsumosActualizarEsPaquete lnIdProducto, True, oConexion
               End If
            Else
               ActualizaItemFarmacia = True
            End If
       Else
            lnIdProducto = oRsTmp1!idProducto
            oDOCatalogoBienesInsumos.idProducto = oRsTmp1!idProducto
            If oCatalogoBienesInsumos.SeleccionarPorId(oDOCatalogoBienesInsumos) = True Then
                oDOCatalogoBienesInsumos.IdUsuarioAuditoria = ml_idUsuario
                If mi_Opcion = sghEliminar Then
                   If oCatalogoBienesInsumos.Eliminar(oDOCatalogoBienesInsumos) = True Then
                      ActualizaItemFarmacia = True
                   End If
                Else
                   oDOCatalogoBienesInsumos.PrecioCompra = lnPrecioCompra
                   oDOCatalogoBienesInsumos.PrecioDistribucion = lnPrecioDistribucion
                   oDOCatalogoBienesInsumos.PrecioDonacion = lnPrecioCompra
                   oDOCatalogoBienesInsumos.nombre = Me.txtDescripcion.Text
                   oDOCatalogoBienesInsumos.FormaFarmaceutica = "UNI"
                   If oCatalogoBienesInsumos.Modificar(oDOCatalogoBienesInsumos) = True Then
                      mo_ReglasFacturacion.FactCatalogoBienesInsumosActualizarEsPaquete lnIdProducto, True, oConexion
                      ActualizaItemFarmacia = True
                   End If
                End If
            End If
       End If
       Set oDOCatalogoBienesInsumos = Nothing
       Set oCatalogoBienesInsumos = Nothing
        
        Dim oRsTmp As New Recordset
        Dim lcSql As String
        Dim lbNuevo As Boolean
        Dim oDoFinanciamientoCatalogoBien As New DoFinanciamientoCatalogoBien, oFinanciamientoCatalogoBien As New FinanciamientoCatalogoBien
        Set oFinanciamientoCatalogoBien.Conexion = oConexion
        'Agrega Insumo en otros Tipo de Financiamiento y actualiza Precio
        Set oRsTmp = mo_ReglasFacturacion.CatalogoBienesInsumosHospSeleccionarXIdProducto(lnIdProducto)
        Set oRsTmp1 = mo_AdminComun.TiposFinanciamientoSegunFiltro("SeIngresPrecios=1")
        If oRsTmp1.RecordCount > 0 Then
           Set oFinanciamientoCatalogoBien.Conexion = oConexion
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              lbNuevo = True
              If oRsTmp.RecordCount > 0 Then
                 oRsTmp.MoveFirst
                 oRsTmp.Find "idTipoFinanciamiento=" & oRsTmp1.Fields!IdTipoFinanciamiento
                 If Not oRsTmp.EOF Then
                    lbNuevo = False
                 End If
              End If

              oDoFinanciamientoCatalogoBien.idProducto = lnIdProducto
              oDoFinanciamientoCatalogoBien.IdTipoFinanciamiento = oRsTmp1.Fields!IdTipoFinanciamiento
              oDoFinanciamientoCatalogoBien.Activo = 1
              oDoFinanciamientoCatalogoBien.PrecioUnitario = lnPrecioVenta
              oDoFinanciamientoCatalogoBien.IdUsuarioAuditoria = ml_idUsuario
              If lbNuevo = True Then
                 If oFinanciamientoCatalogoBien.Insertar(oDoFinanciamientoCatalogoBien) = False Then
                 End If
              Else
                  oDoFinanciamientoCatalogoBien.IdPlanCatalogo = oRsTmp.Fields!IdPlanCatalogo
                  If oFinanciamientoCatalogoBien.Modificar(oDoFinanciamientoCatalogoBien) = False Then
                  End If
              End If
              oRsTmp1.MoveNext
           Loop
        End If
        Set oRsTmp = Nothing
        Set oRsTmp1 = Nothing
        Set oDoFinanciamientoCatalogoBien = Nothing
        Set oFinanciamientoCatalogoBien = Nothing
        oConexion.Close
        Set oConexion = Nothing
    
    
    
    Else
       ActualizaItemFarmacia = True
    End If
End Function

Private Sub txtDescripcion_LostFocus()
   txtDescripcion.Text = UCase(txtDescripcion.Text)
End Sub
