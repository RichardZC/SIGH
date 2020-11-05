VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl UcRecetaCE 
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11685
   ScaleHeight     =   9225
   ScaleWidth      =   11685
   Begin SISGalenPlus.ucRecetaCpt ucRecetaCpt1 
      Height          =   1695
      Left            =   0
      TabIndex        =   43
      Top             =   7530
      Width           =   10900
      _ExtentX        =   19235
      _ExtentY        =   2990
   End
   Begin UltraGrid.SSUltraGrid grdDiag 
      Height          =   3600
      Left            =   3600
      TabIndex        =   41
      Top             =   1905
      Visible         =   0   'False
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   6350
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   71303188
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   "UcRecetaCE.ctx":0000
      Caption         =   "grdDiag"
   End
   Begin VB.CheckBox ChkRegistraTodosItems 
      Height          =   315
      Left            =   11310
      TabIndex        =   40
      ToolTipText     =   "Registra todos los ITEMS en una sola VENTANA"
      Top             =   780
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton cmdPaquetes 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11280
      Picture         =   "UcRecetaCE.ctx":003C
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Lista de PAQUETES"
      Top             =   390
      Width           =   405
   End
   Begin VB.Frame FraRayos 
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
      Height          =   1515
      Left            =   0
      TabIndex        =   31
      Top             =   1920
      Width           =   5805
      Begin VB.CommandButton btnAddRayosX 
         DisabledPicture =   "UcRecetaCE.ctx":05C6
         DownPicture     =   "UcRecetaCE.ctx":09AF
         Height          =   315
         Left            =   5490
         Picture         =   "UcRecetaCE.ctx":0DBB
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   270
      End
      Begin UltraGrid.SSUltraGrid grdRayos 
         Height          =   1215
         Left            =   75
         TabIndex        =   33
         Top             =   270
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   2143
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SSUltraGrid1"
      End
      Begin VB.Label lblRayosX 
         AutoSize        =   -1  'True
         Caption         =   "()"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   990
         TabIndex        =   34
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.Frame FraEcografiaO 
      Caption         =   "Ecografía Obstétrica"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   0
      TabIndex        =   27
      Top             =   3480
      Width           =   5805
      Begin VB.CommandButton btnAddEcoO 
         DisabledPicture =   "UcRecetaCE.ctx":11C7
         DownPicture     =   "UcRecetaCE.ctx":15B0
         Height          =   315
         Left            =   5475
         Picture         =   "UcRecetaCE.ctx":19BC
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   255
         Width           =   270
      End
      Begin UltraGrid.SSUltraGrid grdEcografiaO 
         Height          =   1245
         Left            =   90
         TabIndex        =   29
         Top             =   255
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   2196
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SSUltraGrid1"
      End
      Begin VB.Label lblEcografiaO 
         AutoSize        =   -1  'True
         Caption         =   "()"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1950
         TabIndex        =   30
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.Frame FraEcografiaG 
      Caption         =   "Ecografía General"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   0
      TabIndex        =   23
      Top             =   5070
      Width           =   5805
      Begin VB.CommandButton btnAddEcoG 
         DisabledPicture =   "UcRecetaCE.ctx":1DC8
         DownPicture     =   "UcRecetaCE.ctx":21B1
         Height          =   345
         Left            =   5490
         Picture         =   "UcRecetaCE.ctx":25BD
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   255
         Width           =   270
      End
      Begin UltraGrid.SSUltraGrid grdEcografiaG 
         Height          =   915
         Left            =   60
         TabIndex        =   25
         Top             =   240
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   1614
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SSUltraGrid1"
      End
      Begin VB.Label lblEcografiaG 
         AutoSize        =   -1  'True
         Caption         =   "()"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1770
         TabIndex        =   26
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.Frame FraTomografia 
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
      Height          =   1245
      Left            =   0
      TabIndex        =   19
      Top             =   6300
      Width           =   5775
      Begin VB.CommandButton btnAddTomografia 
         DisabledPicture =   "UcRecetaCE.ctx":29C9
         DownPicture     =   "UcRecetaCE.ctx":2DB2
         Height          =   345
         Left            =   5460
         Picture         =   "UcRecetaCE.ctx":31BE
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   270
      End
      Begin UltraGrid.SSUltraGrid grdTomografia 
         Height          =   975
         Left            =   60
         TabIndex        =   21
         Top             =   240
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   1720
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SSUltraGrid1"
      End
      Begin VB.Label lblTomografia 
         AutoSize        =   -1  'True
         Caption         =   "()"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1320
         TabIndex        =   22
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.Frame FraFarmacia 
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
      Height          =   1875
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11225
      Begin VB.CheckBox chkVerDx 
         Caption         =   "Ver Dx"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   42
         Top             =   255
         Width           =   1200
      End
      Begin VB.ComboBox cmdFarmacias 
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
         Left            =   3795
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   135
         Width           =   4680
      End
      Begin VB.CommandButton btnAddFarmacia 
         DisabledPicture =   "UcRecetaCE.ctx":35CA
         DownPicture     =   "UcRecetaCE.ctx":39B3
         Height          =   345
         Left            =   10905
         Picture         =   "UcRecetaCE.ctx":3DBF
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   495
         Width           =   270
      End
      Begin UltraGrid.SSUltraGrid grdFarmacia 
         Height          =   1365
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   2408
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SSUltraGrid1"
      End
      Begin MSMask.MaskEdBox txtFechaVigencia 
         Height          =   330
         Left            =   9570
         TabIndex        =   35
         Top             =   165
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Elegir Farmacia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   2430
         TabIndex        =   39
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "F.Vigencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   8670
         TabIndex        =   36
         Top             =   210
         Width           =   900
      End
      Begin VB.Label lblFarmacia 
         AutoSize        =   -1  'True
         Caption         =   "()"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   960
         TabIndex        =   18
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.Frame FraAnatomia 
      Caption         =   "Anatomía Patológica"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   5850
      TabIndex        =   11
      Top             =   4230
      Width           =   5805
      Begin VB.CommandButton btnAddAnatomia 
         DisabledPicture =   "UcRecetaCE.ctx":41CB
         DownPicture     =   "UcRecetaCE.ctx":45B4
         Height          =   345
         Left            =   5475
         Picture         =   "UcRecetaCE.ctx":49C0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   270
      End
      Begin UltraGrid.SSUltraGrid grdAnatomia 
         Height          =   1455
         Left            =   60
         TabIndex        =   13
         Top             =   255
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   2566
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SSUltraGrid1"
      End
      Begin VB.Label lblAnatomia 
         AutoSize        =   -1  'True
         Caption         =   "()"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1980
         TabIndex        =   14
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.Frame FraPatologia 
      Caption         =   "Patología Clínica"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   5865
      TabIndex        =   7
      Top             =   1935
      Width           =   5805
      Begin VB.CommandButton btnAddPatologia 
         DisabledPicture =   "UcRecetaCE.ctx":4DCC
         DownPicture     =   "UcRecetaCE.ctx":51B5
         Height          =   345
         Left            =   5475
         Picture         =   "UcRecetaCE.ctx":55C1
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   255
         Width           =   270
      End
      Begin UltraGrid.SSUltraGrid grdPatologia 
         Height          =   1935
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   3413
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SSUltraGrid1"
      End
      Begin VB.Label lblPatologia 
         AutoSize        =   -1  'True
         Caption         =   "()"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1590
         TabIndex        =   10
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.Frame FraBancoS 
      Caption         =   "Banco de Sangre"
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
      Left            =   5850
      TabIndex        =   3
      Top             =   6030
      Width           =   5805
      Begin VB.CommandButton btnAddBanco 
         DisabledPicture =   "UcRecetaCE.ctx":59CD
         DownPicture     =   "UcRecetaCE.ctx":5DB6
         Height          =   345
         Left            =   5475
         Picture         =   "UcRecetaCE.ctx":61C2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   300
         Width           =   270
      End
      Begin UltraGrid.SSUltraGrid grdBanco 
         Height          =   1155
         Left            =   60
         TabIndex        =   5
         Top             =   300
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   2037
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SSUltraGrid1"
      End
      Begin VB.Label lblBancoS 
         AutoSize        =   -1  'True
         Caption         =   "()"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1740
         TabIndex        =   6
         Top             =   30
         Width           =   90
      End
   End
   Begin VB.CommandButton btnImprimir 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11280
      Picture         =   "UcRecetaCE.ctx":65CE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir recetas de farmacia"
      Top             =   0
      Width           =   405
   End
   Begin VB.TextBox txtCitaExClinicos 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   11250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1095
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton btnImprimirOrden 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11280
      Picture         =   "UcRecetaCE.ctx":6AA7
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir ordenes médicas"
      Top             =   1575
      Width           =   405
   End
End
Attribute VB_Name = "UcRecetaCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para registro de Receta0
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Const lcLinea As String = "----------------------------------------------------------------------------------------"
Const lcLineaChar As String = "¨"
Dim oRsRayosX As New Recordset, lnIdRecetaRayosX As Long
Dim oRsEcografiaO As New Recordset, lnIdRecetaEcografiaO As Long
Dim oRsEcografiaG As New Recordset, lnIdRecetaEcografiaG As Long
Dim oRsTomografia As New Recordset, lnIdRecetaTomografia As Long
Dim oRsAnatomia As New Recordset, lnIdRecetaAnatomia As Long
Dim oRsPatologia As New Recordset, lnIdRecetaPatología As Long
Dim oRsBanco As New Recordset, lnIdRecetaBanco As Long
Dim oRsFarmacia As New Recordset, lnIdRecetaFarmacia As Long
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_Formulario As New sighEntidades.Formulario
Dim lcSql As String
Dim ml_IdTipoFinanciamiento As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_sighProxies As New SIGHProxies.Procesos
Dim ml_DatoCabeceraReceta As String
Dim lc_Tratamiento As String
Dim lnIdDosisDefault As Long
Dim ml_idCuentaAtencion As Long
Dim ml_AScorrelativo As Long
Dim mRs_Fua As New Recordset
Dim ml_IdFuenteFinanciamiento As Long
'debb-24/06/2015
Dim lnDiasMaximoVigencia As Long
Dim lnMaximoItems As Long
Dim lnTotalReg As Long
Dim mi_Opcion As sghOpciones
Dim lnIdFarmaciaElegida As Long                       'debb-14/07/2015
Dim mi_lnWnd As Long
Dim ml_idTipoSexo As Long
Dim lcDxUnico As String

Property Let idTipoSexo(lValue As Long)
   ml_idTipoSexo = lValue
End Property

Property Let lnWnd(lValue As Long)
    mi_lnWnd = lValue
End Property

'debb-21/07/2015
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
   If mi_Opcion <> sghAgregar Then
      cmdPaquetes.Visible = False
      ChkRegistraTodosItems.Visible = False
   Else
      cmdPaquetes.Visible = True
      'If mo_reglasComunes.EESSconDerechosAmejoras(2, "61008", lcMensajeLicencia) = True Then
         ChkRegistraTodosItems.Visible = True
      'End If
   End If
End Property

Property Let IdFuenteFinanciamiento(lValue As Long)
   ml_IdFuenteFinanciamiento = lValue
End Property

Property Set RsServiciosAtenSimultaneaFuaXcorrelativo(oValue As Recordset)
    Set mRs_Fua = oValue
    If oRsRayosX.RecordCount > 0 Then
        mRs_Fua.Filter = "idtipo=1"
        If mRs_Fua.RecordCount > 0 Then
           mRs_Fua.MoveFirst
           Do While Not mRs_Fua.EOF
              oRsRayosX.MoveFirst
              Do While Not oRsRayosX.EOF
                 If oRsRayosX.Fields!ID = mRs_Fua!Item Then
                     oRsRayosX.Fields!FUA = mRs_Fua!idFuaCorrelativo
                 End If
                 oRsRayosX.MoveNext
              Loop
              mRs_Fua.MoveNext
           Loop
        End If
    End If
    If oRsEcografiaO.RecordCount > 0 Then
        mRs_Fua.Filter = "idtipo=1"
        If mRs_Fua.RecordCount > 0 Then
           mRs_Fua.MoveFirst
           Do While Not mRs_Fua.EOF
              oRsEcografiaO.MoveFirst
              Do While Not oRsEcografiaO.EOF
                 If oRsEcografiaO.Fields!ID = mRs_Fua!Item Then
                     oRsEcografiaO.Fields!FUA = mRs_Fua!idFuaCorrelativo
                 End If
                 oRsEcografiaO.MoveNext
              Loop
              mRs_Fua.MoveNext
           Loop
        End If
    End If
    If oRsEcografiaG.RecordCount > 0 Then
        mRs_Fua.Filter = "idtipo=1"
        If mRs_Fua.RecordCount > 0 Then
           mRs_Fua.MoveFirst
           Do While Not mRs_Fua.EOF
              oRsEcografiaG.MoveFirst
              Do While Not oRsEcografiaG.EOF
                 If oRsEcografiaG.Fields!ID = mRs_Fua!Item Then
                     oRsEcografiaG.Fields!FUA = mRs_Fua!idFuaCorrelativo
                 End If
                 oRsEcografiaG.MoveNext
              Loop
              mRs_Fua.MoveNext
           Loop
        End If
    End If
    If oRsTomografia.RecordCount > 0 Then
        mRs_Fua.Filter = "idtipo=1"
        If mRs_Fua.RecordCount > 0 Then
           mRs_Fua.MoveFirst
           Do While Not mRs_Fua.EOF
              oRsTomografia.MoveFirst
              Do While Not oRsTomografia.EOF
                 If oRsTomografia.Fields!ID = mRs_Fua!Item Then
                     oRsTomografia.Fields!FUA = mRs_Fua!idFuaCorrelativo
                 End If
                 oRsTomografia.MoveNext
              Loop
              mRs_Fua.MoveNext
           Loop
        End If
    End If
    If oRsAnatomia.RecordCount > 0 Then
        mRs_Fua.Filter = "idtipo=1"
        If mRs_Fua.RecordCount > 0 Then
           mRs_Fua.MoveFirst
           Do While Not mRs_Fua.EOF
              oRsAnatomia.MoveFirst
              Do While Not oRsAnatomia.EOF
                 If oRsAnatomia.Fields!ID = mRs_Fua!Item Then
                     oRsAnatomia.Fields!FUA = mRs_Fua!idFuaCorrelativo
                 End If
                 oRsAnatomia.MoveNext
              Loop
              mRs_Fua.MoveNext
           Loop
        End If
    End If
    If oRsPatologia.RecordCount > 0 Then
        mRs_Fua.Filter = "idtipo=1"
        If mRs_Fua.RecordCount > 0 Then
           mRs_Fua.MoveFirst
           Do While Not mRs_Fua.EOF
              oRsPatologia.MoveFirst
              Do While Not oRsPatologia.EOF
                 If oRsPatologia.Fields!ID = mRs_Fua!Item Then
                     oRsPatologia.Fields!FUA = mRs_Fua!idFuaCorrelativo
                 End If
                 oRsPatologia.MoveNext
              Loop
              mRs_Fua.MoveNext
           Loop
        End If
    End If
    If oRsBanco.RecordCount > 0 Then
        mRs_Fua.Filter = "idtipo=1"
        If mRs_Fua.RecordCount > 0 Then
           mRs_Fua.MoveFirst
           Do While Not mRs_Fua.EOF
              oRsBanco.MoveFirst
              Do While Not oRsBanco.EOF
                 If oRsBanco.Fields!ID = mRs_Fua!Item Then
                     oRsBanco.Fields!FUA = mRs_Fua!idFuaCorrelativo
                 End If
                 oRsBanco.MoveNext
              Loop
              mRs_Fua.MoveNext
           Loop
        End If
    End If
    If oRsFarmacia.RecordCount > 0 Then
        mRs_Fua.Filter = "idtipo=2"
        If mRs_Fua.RecordCount > 0 Then
           mRs_Fua.MoveFirst
           Do While Not mRs_Fua.EOF
              oRsFarmacia.MoveFirst
              Do While Not oRsFarmacia.EOF
                 If oRsFarmacia.Fields!ID = mRs_Fua!Item Then
                     oRsFarmacia.Fields!FUA = mRs_Fua!idFuaCorrelativo
                 End If
                 oRsFarmacia.MoveNext
              Loop
              mRs_Fua.MoveNext
           Loop
        End If
    End If
End Property

Property Let AScorrelativo(lValue As Long)
    ml_AScorrelativo = lValue
    '
    'If ml_AScorrelativo > 0 Then
       Dim lnFor As Integer
       If wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
            grdFarmacia.Bands(0).Columns("fua").Header.Appearance.ForeColor = vbWhite
            grdFarmacia.Bands(0).Columns("fua").Header.Appearance.BackColor = vbRed
            grdFarmacia.Bands(0).Columns("Fua").Width = 700
            On Error Resume Next
            With grdFarmacia.ValueLists.Add("FuaList").ValueListItems
                For lnFor = 1 To 20
                    .Add lnFor, "N° " & Trim(Str(lnFor))
                Next
            End With
            grdFarmacia.Bands(0).Columns("Fua").ValueList = "FuaList"
            grdFarmacia.Bands(0).Columns("fua").Hidden = False
       Else
            grdFarmacia.Bands(0).Columns("fua").Hidden = True
       End If
       '
       If wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
            grdRayos.Bands(0).Columns("fua").Header.Appearance.ForeColor = vbWhite
            grdRayos.Bands(0).Columns("fua").Header.Appearance.BackColor = vbRed
            grdRayos.Bands(0).Columns("Fua").Width = 700
            With grdRayos.ValueLists.Add("FuaList").ValueListItems
                 For lnFor = 1 To 20
                     .Add lnFor, "N° " & Trim(Str(lnFor))
                 Next
            End With
            grdRayos.Bands(0).Columns("Fua").ValueList = "FuaList"
            grdRayos.Bands(0).Columns("fua").Hidden = False
       Else
            grdRayos.Bands(0).Columns("fua").Hidden = True
       End If
       '
       If wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
            grdEcografiaO.Bands(0).Columns("fua").Header.Appearance.ForeColor = vbWhite
            grdEcografiaO.Bands(0).Columns("fua").Header.Appearance.BackColor = vbRed
            grdEcografiaO.Bands(0).Columns("Fua").Width = 700
            With grdEcografiaO.ValueLists.Add("FuaList").ValueListItems
                 For lnFor = 1 To 20
                     .Add lnFor, "N° " & Trim(Str(lnFor))
                 Next
            End With
            grdEcografiaO.Bands(0).Columns("Fua").ValueList = "FuaList"
            grdEcografiaO.Bands(0).Columns("fua").Hidden = False
       Else
            grdEcografiaO.Bands(0).Columns("fua").Hidden = True
       End If
       '
       If wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
            grdEcografiaG.Bands(0).Columns("fua").Header.Appearance.ForeColor = vbWhite
            grdEcografiaG.Bands(0).Columns("fua").Header.Appearance.BackColor = vbRed
            grdEcografiaG.Bands(0).Columns("Fua").Width = 700
            With grdEcografiaG.ValueLists.Add("FuaList").ValueListItems
                 For lnFor = 1 To 20
                     .Add lnFor, "N° " & Trim(Str(lnFor))
                 Next
            End With
            grdEcografiaG.Bands(0).Columns("Fua").ValueList = "FuaList"
            grdEcografiaG.Bands(0).Columns("fua").Hidden = False
       Else
            grdEcografiaG.Bands(0).Columns("fua").Hidden = True
       End If
       '
       If wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
            grdTomografia.Bands(0).Columns("fua").Header.Appearance.ForeColor = vbWhite
            grdTomografia.Bands(0).Columns("fua").Header.Appearance.BackColor = vbRed
            grdTomografia.Bands(0).Columns("Fua").Width = 700
            With grdTomografia.ValueLists.Add("FuaList").ValueListItems
                 For lnFor = 1 To 20
                     .Add lnFor, "N° " & Trim(Str(lnFor))
                 Next
            End With
            grdTomografia.Bands(0).Columns("Fua").ValueList = "FuaList"
            grdTomografia.Bands(0).Columns("fua").Hidden = False
       Else
            grdTomografia.Bands(0).Columns("fua").Hidden = True
       End If
       '
       If wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
            grdPatologia.Bands(0).Columns("fua").Header.Appearance.ForeColor = vbWhite
            grdPatologia.Bands(0).Columns("fua").Header.Appearance.BackColor = vbRed
            grdPatologia.Bands(0).Columns("Fua").Width = 700
            With grdPatologia.ValueLists.Add("FuaList").ValueListItems
                 For lnFor = 1 To 20
                     .Add lnFor, "N° " & Trim(Str(lnFor))
                 Next
            End With
            grdPatologia.Bands(0).Columns("Fua").ValueList = "FuaList"
            grdPatologia.Bands(0).Columns("fua").Hidden = False
       Else
            grdPatologia.Bands(0).Columns("fua").Hidden = True
       End If
       '
       If wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
            grdAnatomia.Bands(0).Columns("fua").Header.Appearance.ForeColor = vbWhite
            grdAnatomia.Bands(0).Columns("fua").Header.Appearance.BackColor = vbRed
            grdAnatomia.Bands(0).Columns("Fua").Width = 700
            With grdAnatomia.ValueLists.Add("FuaList").ValueListItems
                 For lnFor = 1 To 20
                     .Add lnFor, "N° " & Trim(Str(lnFor))
                 Next
            End With
            grdAnatomia.Bands(0).Columns("Fua").ValueList = "FuaList"
            grdAnatomia.Bands(0).Columns("fua").Hidden = False
       Else
            grdAnatomia.Bands(0).Columns("fua").Hidden = True
       End If
       '
       If wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
            grdBanco.Bands(0).Columns("fua").Header.Appearance.ForeColor = vbWhite
            grdBanco.Bands(0).Columns("fua").Header.Appearance.BackColor = vbRed
            grdBanco.Bands(0).Columns("Fua").Width = 700
            With grdBanco.ValueLists.Add("FuaList").ValueListItems
                 For lnFor = 1 To 20
                     .Add lnFor, "N° " & Trim(Str(lnFor))
                 Next
            End With
            grdBanco.Bands(0).Columns("Fua").ValueList = "FuaList"
            grdBanco.Bands(0).Columns("fua").Hidden = False
       Else
            grdBanco.Bands(0).Columns("fua").Hidden = True
       End If
'    Else
'       grdFarmacia.Bands(0).Columns("fua").Hidden = True
'       grdRayos.Bands(0).Columns("fua").Hidden = True
'       grdEcografiaO.Bands(0).Columns("fua").Hidden = True
'       grdEcografiaG.Bands(0).Columns("fua").Hidden = True
'       grdTomografia.Bands(0).Columns("fua").Hidden = True
'       grdPatologia.Bands(0).Columns("fua").Hidden = True
'       grdAnatomia.Bands(0).Columns("fua").Hidden = True
'       grdBanco.Bands(0).Columns("fua").Hidden = True
'    End If
End Property
Property Let Tratamiento(lValue As String)
   lc_Tratamiento = lValue
End Property

Property Let DatoCabeceraReceta(lValue As String)
   ml_DatoCabeceraReceta = lValue
End Property
Property Let idTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
   UserControl.ucRecetaCpt1.idTipoFinanciamiento = lValue
End Property
'Actualizado 2209
Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property

Public Sub Inicializar()
    CreaTemporales True, True, True, True, True, True, True, True, True
    '
    InicializarLaGrilla grdFarmacia
    InicializarLaGrilla grdRayos
    InicializarLaGrilla grdEcografiaO
    InicializarLaGrilla grdEcografiaG
    InicializarLaGrilla grdTomografia
    InicializarLaGrilla grdPatologia
    InicializarLaGrilla grdAnatomia
    InicializarLaGrilla grdBanco
    '
    'debb-24/06/2015
    lnDiasMaximoVigencia = Val(lcBuscaParametro.SeleccionaFilaParametro(356))
    '
    lnMaximoItems = Val(lcBuscaParametro.SeleccionaFilaParametro(355))
    If UCase(lcBuscaParametro.SeleccionaFilaParametro(500)) = "S" Then  'debb-18/05/2016
       lnMaximoItems = 500                                              'debb-18/05/2016
    End If                                                              'debb-18/05/2016
    '
    txtFechaVigencia.Text = CDate(lcBuscaParametro.RetornaFechaServidorSQL) + lnDiasMaximoVigencia
    '
    '
    CargaFarmaciasAelegir            'debb-14/07/2015
    '
    wxParametro513 = lcBuscaParametro.SeleccionaFilaParametro(513)
    '
    ucRecetaCpt1.Inicializar
    ucRecetaCpt1.MaximoItems = lnMaximoItems
    ucRecetaCpt1.idTipoFinanciamiento = ml_IdTipoFinanciamiento
    ucRecetaCpt1.Dx = lcDxUnico
    '
End Sub
'debb-11/04/2016
Private Sub btnAddAnatomia_Click()
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    Dim oRsItemsElegidos As New Recordset
    Dim oRsDevuelveTodosLosItemsServ As New Recordset
    oPaquetesBuscar.idPuntoCarga = sghPtoCargaAnatomiaPatologica1
    oPaquetesBuscar.RegistraTodosLosItems = IIf(ChkRegistraTodosItems.Value = 1, True, False)
    oPaquetesBuscar.idTipoFinanciamiento = ml_IdTipoFinanciamiento
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
        Set oRsItemsElegidos = oPaquetesBuscar.ItemsMasivosElegidos
        Set oRsDevuelveTodosLosItemsServ = oPaquetesBuscar.DevuelveTodosLosItemsServ
        mo_sighProxies.AgregaItemsDeReceta sghPtoCargaAnatomiaPatologica1, oRsItemsElegidos, oRsPatologia, _
                         oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                         lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, 0, False, _
                         oRsDevuelveTodosLosItemsServ, Nothing, lcDxUnico
    End If
    Set oPaquetesBuscar = Nothing
    Set oRsItemsElegidos = Nothing
    Set oRsDevuelveTodosLosItemsServ = Nothing

End Sub
'debb-11/04/2016
Private Sub btnAddBanco_Click()
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    Dim oRsItemsElegidos As New Recordset
    Dim oRsDevuelveTodosLosItemsServ As New Recordset
    oPaquetesBuscar.idPuntoCarga = sghPtoCargaBancoSangre1
    oPaquetesBuscar.RegistraTodosLosItems = IIf(ChkRegistraTodosItems.Value = 1, True, False)
    oPaquetesBuscar.idTipoFinanciamiento = ml_IdTipoFinanciamiento
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
        Set oRsItemsElegidos = oPaquetesBuscar.ItemsMasivosElegidos
        Set oRsDevuelveTodosLosItemsServ = oPaquetesBuscar.DevuelveTodosLosItemsServ
        mo_sighProxies.AgregaItemsDeReceta sghPtoCargaBancoSangre1, oRsItemsElegidos, oRsPatologia, _
                         oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                         lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, 0, False, _
                         oRsDevuelveTodosLosItemsServ, Nothing, lcDxUnico
    End If
    Set oPaquetesBuscar = Nothing
    Set oRsItemsElegidos = Nothing
    Set oRsDevuelveTodosLosItemsServ = Nothing

End Sub
'debb-11/04/2016
Private Sub btnAddEcoG_Click()
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    Dim oRsItemsElegidos As New Recordset
    Dim oRsDevuelveTodosLosItemsServ As New Recordset
    oPaquetesBuscar.idPuntoCarga = sghPtoCargaEcogGeneral
    oPaquetesBuscar.RegistraTodosLosItems = IIf(ChkRegistraTodosItems.Value = 1, True, False)
    oPaquetesBuscar.idTipoFinanciamiento = ml_IdTipoFinanciamiento
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
        Set oRsItemsElegidos = oPaquetesBuscar.ItemsMasivosElegidos
        Set oRsDevuelveTodosLosItemsServ = oPaquetesBuscar.DevuelveTodosLosItemsServ
        mo_sighProxies.AgregaItemsDeReceta sghPtoCargaEcogGeneral, oRsItemsElegidos, oRsPatologia, _
                         oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                         lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, 0, False, _
                         oRsDevuelveTodosLosItemsServ, Nothing, lcDxUnico
    End If
    Set oPaquetesBuscar = Nothing
    Set oRsItemsElegidos = Nothing
    Set oRsDevuelveTodosLosItemsServ = Nothing
End Sub
'debb-11/04/2016
Private Sub btnAddEcoO_Click()
    If ml_idTipoSexo = 1 Then
       MsgBox "Solo se agrega CPT a Pacientes de sexo FEMENINO", vbInformation, ""
       Exit Sub
    End If
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    Dim oRsItemsElegidos As New Recordset
    Dim oRsDevuelveTodosLosItemsServ As New Recordset
    oPaquetesBuscar.idPuntoCarga = sghPtoCargaEcogObstetrica
    oPaquetesBuscar.RegistraTodosLosItems = IIf(ChkRegistraTodosItems.Value = 1, True, False)
    oPaquetesBuscar.idTipoFinanciamiento = ml_IdTipoFinanciamiento
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
        Set oRsItemsElegidos = oPaquetesBuscar.ItemsMasivosElegidos
        Set oRsDevuelveTodosLosItemsServ = oPaquetesBuscar.DevuelveTodosLosItemsServ
        mo_sighProxies.AgregaItemsDeReceta sghPtoCargaEcogObstetrica, oRsItemsElegidos, oRsPatologia, _
                         oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                         lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, 0, False, _
                         oRsDevuelveTodosLosItemsServ, Nothing, lcDxUnico
    End If
    Set oPaquetesBuscar = Nothing
    Set oRsItemsElegidos = Nothing
    Set oRsDevuelveTodosLosItemsServ = Nothing
End Sub
'debb-11/04/2016
Private Sub btnAddFarmacia_Click()
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    Dim oRsItemsElegidos As New Recordset
    Dim oRsDevuelveTodosLosItemsFarm As New Recordset
    oPaquetesBuscar.idPuntoCarga = sghPtoCargaFarmacia
    'debb-14/07/2015 (INICIO)
    oPaquetesBuscar.FarmaciaElegida = IIf(lnIdFarmaciaElegida > 1, _
                                      "Farmacia: " & UCase(Mid(cmdFarmacias.Text, InStr(cmdFarmacias.Text, "-") + 1)), _
                                      cmdFarmacias.Text)
    oPaquetesBuscar.IdFarmaciaElegida = lnIdFarmaciaElegida
    oPaquetesBuscar.RegistraTodosLosItems = IIf(ChkRegistraTodosItems.Value = 1, True, False)
    'debb-14/07/2015 (FIN
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
    
        Set oRsItemsElegidos = oPaquetesBuscar.ItemsMasivosElegidos
        Set oRsDevuelveTodosLosItemsFarm = oPaquetesBuscar.DevuelveTodosLosItemsFarm
        mo_sighProxies.AgregaItemsDeReceta sghPtoCargaFarmacia, oRsItemsElegidos, oRsPatologia, _
                         oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                         lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, lnIdFarmaciaElegida, False, _
                         Nothing, oRsDevuelveTodosLosItemsFarm, lcDxUnico
        
    End If
    Set oPaquetesBuscar = Nothing
    Set oRsItemsElegidos = Nothing
    Set oRsDevuelveTodosLosItemsFarm = Nothing
End Sub
'debb-11/04/2016
Private Sub btnAddPatologia_Click()
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    Dim oRsItemsElegidos As New Recordset
    Dim oRsDevuelveTodosLosItemsServ As New Recordset
    oPaquetesBuscar.idPuntoCarga = sghPtoCargaPatologiaClinica
    oPaquetesBuscar.RegistraTodosLosItems = IIf(ChkRegistraTodosItems.Value = 1, True, False)
    oPaquetesBuscar.idTipoFinanciamiento = ml_IdTipoFinanciamiento
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
        Set oRsItemsElegidos = oPaquetesBuscar.ItemsMasivosElegidos
        Set oRsDevuelveTodosLosItemsServ = oPaquetesBuscar.DevuelveTodosLosItemsServ
        mo_sighProxies.AgregaItemsDeReceta sghPtoCargaPatologiaClinica, oRsItemsElegidos, oRsPatologia, _
                         oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                         lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, 0, False, _
                         oRsDevuelveTodosLosItemsServ, Nothing, lcDxUnico
    End If
    Set oPaquetesBuscar = Nothing
    Set oRsItemsElegidos = Nothing
    Set oRsDevuelveTodosLosItemsServ = Nothing
End Sub
'debb-11/04/2016
Private Sub btnAddRayosX_Click()
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    Dim oRsItemsElegidos As New Recordset
    Dim oRsDevuelveTodosLosItemsServ As New Recordset
    oPaquetesBuscar.idPuntoCarga = sghPtoCargaRayosX
    oPaquetesBuscar.RegistraTodosLosItems = IIf(ChkRegistraTodosItems.Value = 1, True, False)
    oPaquetesBuscar.idTipoFinanciamiento = ml_IdTipoFinanciamiento
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
        Set oRsItemsElegidos = oPaquetesBuscar.ItemsMasivosElegidos
        Set oRsDevuelveTodosLosItemsServ = oPaquetesBuscar.DevuelveTodosLosItemsServ
        mo_sighProxies.AgregaItemsDeReceta sghPtoCargaRayosX, oRsItemsElegidos, oRsPatologia, _
                         oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                         lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, 0, False, _
                         oRsDevuelveTodosLosItemsServ, Nothing, lcDxUnico
    End If
    Set oPaquetesBuscar = Nothing
    Set oRsItemsElegidos = Nothing
    Set oRsDevuelveTodosLosItemsServ = Nothing
End Sub
'debb-11/04/2016
Private Sub btnAddTomografia_Click()
    Dim oPaquetesBuscar As New SIGHNegocios.BuscarFactCatalogoPqte
    Dim oRsItemsElegidos As New Recordset
    Dim oRsDevuelveTodosLosItemsServ As New Recordset
    oPaquetesBuscar.idPuntoCarga = sghPtoCargaTomografia
    oPaquetesBuscar.RegistraTodosLosItems = IIf(ChkRegistraTodosItems.Value = 1, True, False)
    oPaquetesBuscar.idTipoFinanciamiento = ml_IdTipoFinanciamiento
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
        Set oRsItemsElegidos = oPaquetesBuscar.ItemsMasivosElegidos
        Set oRsDevuelveTodosLosItemsServ = oPaquetesBuscar.DevuelveTodosLosItemsServ
        mo_sighProxies.AgregaItemsDeReceta sghPtoCargaTomografia, oRsItemsElegidos, oRsPatologia, _
                         oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                         lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, 0, False, _
                         oRsDevuelveTodosLosItemsServ, Nothing, lcDxUnico
    End If
    Set oPaquetesBuscar = Nothing
    Set oRsItemsElegidos = Nothing
    Set oRsDevuelveTodosLosItemsServ = Nothing
End Sub

Sub CreaTemporales(lbHabilitaFrame As Boolean, _
                   lbSoloLimpiaEcografiaO As Boolean, lbSoloLimpiaEcografiaG As Boolean, _
                   lbSoloLimpiaRayosX As Boolean, lbSoloLimpiaTomografia As Boolean, _
                   lbSoloLimpiaAnatomiaP As Boolean, lbSoloLimpiaBancoS As Boolean, _
                   lbSoloLimpiaPatologiaC As Boolean, lbSoloLimpiaFarmacia As Boolean)
    On Error Resume Next
    If lbSoloLimpiaRayosX = True Then
        With oRsRayosX
              .Fields.Append "Fua", adInteger
              .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
              .Fields.Append "Id", adInteger
              .Fields.Append "Procedimiento", adVarChar, 255, adFldIsNullable
              .Fields.Append "Cantidad", adInteger
              .Fields.Append "idDosisRecetada", adInteger, , adFldIsNullable
              .Fields.Append "HayCpt", adBoolean
              .Fields.Append "Precio", adDouble
              .Fields.Append "SaldoActual", adInteger
              .Fields.Append "Receta", adInteger
              .Fields.Append "idEstadoDetalle", adInteger
              .Fields.Append "MotivoAnulacionMedico", adVarChar, 300, adFldIsNullable
              .Fields.Append "Observaciones", adVarChar, 300, adFldIsNullable
              .CursorType = adOpenDynamic
              .LockType = adLockOptimistic
              .Open
        End With
        Set grdRayos.DataSource = oRsRayosX
        mo_Apariencia.ConfigurarFilasBiColores grdRayos, sighEntidades.GrillaConFilasBicolor
        grdRayos.Caption = ""
        If lbHabilitaFrame = True Then
           FraRayos.Enabled = True
        End If
    End If
    '
    If lbSoloLimpiaEcografiaO = True Then
        With oRsEcografiaO
              .Fields.Append "Fua", adInteger
              .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
              .Fields.Append "Id", adInteger
              .Fields.Append "Procedimiento", adVarChar, 255, adFldIsNullable
              .Fields.Append "Cantidad", adInteger
              .Fields.Append "idDosisRecetada", adInteger
              .Fields.Append "HayCpt", adBoolean
              .Fields.Append "Precio", adDouble
              .Fields.Append "SaldoActual", adInteger
              .Fields.Append "Receta", adInteger
              .Fields.Append "idEstadoDetalle", adInteger
              .Fields.Append "MotivoAnulacionMedico", adVarChar, 300, adFldIsNullable
              .Fields.Append "Observaciones", adVarChar, 300, adFldIsNullable
              .CursorType = adOpenDynamic
              .LockType = adLockOptimistic
              .Open
        End With
        Set grdEcografiaO.DataSource = oRsEcografiaO
        mo_Apariencia.ConfigurarFilasBiColores grdEcografiaO, sighEntidades.GrillaConFilasBicolor
        grdEcografiaO.Caption = ""
        If lbHabilitaFrame = True Then
           FraEcografiaO.Enabled = True
        End If
    End If
    '
    If lbSoloLimpiaEcografiaG = True Then
        With oRsEcografiaG
              .Fields.Append "Fua", adInteger
              .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
              .Fields.Append "Id", adInteger
              .Fields.Append "Procedimiento", adVarChar, 255, adFldIsNullable
              .Fields.Append "Cantidad", adInteger
              .Fields.Append "idDosisRecetada", adInteger
              .Fields.Append "HayCpt", adBoolean
              .Fields.Append "Precio", adDouble
              .Fields.Append "SaldoActual", adInteger
              .Fields.Append "Receta", adInteger
              .Fields.Append "idEstadoDetalle", adInteger
              .Fields.Append "MotivoAnulacionMedico", adVarChar, 300, adFldIsNullable
              .Fields.Append "Observaciones", adVarChar, 300, adFldIsNullable
              .CursorType = adOpenDynamic
              .LockType = adLockOptimistic
              .Open
        End With
        Set grdEcografiaG.DataSource = oRsEcografiaG
        mo_Apariencia.ConfigurarFilasBiColores grdEcografiaG, sighEntidades.GrillaConFilasBicolor
        grdEcografiaG.Caption = ""
        If lbHabilitaFrame = True Then
           FraEcografiaG.Enabled = True
        End If
    End If
    '
    If lbSoloLimpiaTomografia = True Then
        With oRsTomografia
              .Fields.Append "Fua", adInteger
              .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
              .Fields.Append "Id", adInteger
              .Fields.Append "Procedimiento", adVarChar, 255, adFldIsNullable
              .Fields.Append "idDosisRecetada", adInteger
              .Fields.Append "Cantidad", adInteger
              .Fields.Append "HayCpt", adBoolean
              .Fields.Append "Precio", adDouble
              .Fields.Append "SaldoActual", adInteger
              .Fields.Append "Receta", adInteger
              .Fields.Append "idEstadoDetalle", adInteger
              .Fields.Append "MotivoAnulacionMedico", adVarChar, 300, adFldIsNullable
              .Fields.Append "Observaciones", adVarChar, 300, adFldIsNullable
              .CursorType = adOpenDynamic
              .LockType = adLockOptimistic
              .Open
        End With
        Set grdTomografia.DataSource = oRsTomografia
        mo_Apariencia.ConfigurarFilasBiColores grdTomografia, sighEntidades.GrillaConFilasBicolor
        grdTomografia.Caption = ""
        If lbHabilitaFrame = True Then
            FraTomografia.Enabled = True
        End If
    End If
    '
    If lbSoloLimpiaAnatomiaP = True Then
        With oRsAnatomia
              .Fields.Append "Fua", adInteger
              .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
              .Fields.Append "Id", adInteger
              .Fields.Append "Procedimiento", adVarChar, 255, adFldIsNullable
              .Fields.Append "Cantidad", adInteger
              .Fields.Append "idDosisRecetada", adInteger
              .Fields.Append "HayCpt", adBoolean
              .Fields.Append "Precio", adDouble
              .Fields.Append "SaldoActual", adInteger
              .Fields.Append "Receta", adInteger
              .Fields.Append "idEstadoDetalle", adInteger
              .Fields.Append "MotivoAnulacionMedico", adVarChar, 300, adFldIsNullable
              .Fields.Append "Observaciones", adVarChar, 300, adFldIsNullable
              .CursorType = adOpenDynamic
              .LockType = adLockOptimistic
              .Open
        End With
        Set grdAnatomia.DataSource = oRsAnatomia
        mo_Apariencia.ConfigurarFilasBiColores grdAnatomia, sighEntidades.GrillaConFilasBicolor
        grdAnatomia.Caption = ""
        If lbHabilitaFrame = True Then
           FraAnatomia.Enabled = True
        End If
    End If
    '
    If lbSoloLimpiaPatologiaC = True Then
        With oRsPatologia
              .Fields.Append "Fua", adInteger
              .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
              .Fields.Append "Id", adInteger
              .Fields.Append "Procedimiento", adVarChar, 255, adFldIsNullable
              .Fields.Append "Cantidad", adInteger
              .Fields.Append "idDosisRecetada", adInteger
              .Fields.Append "HayCpt", adBoolean
              .Fields.Append "Precio", adDouble
              .Fields.Append "SaldoActual", adInteger
              .Fields.Append "Receta", adInteger
              .Fields.Append "idEstadoDetalle", adInteger
              .Fields.Append "MotivoAnulacionMedico", adVarChar, 300, adFldIsNullable
              .Fields.Append "Observaciones", adVarChar, 300, adFldIsNullable
              .CursorType = adOpenDynamic
              .LockType = adLockOptimistic
              .Open
        End With
        Set grdPatologia.DataSource = oRsPatologia
        mo_Apariencia.ConfigurarFilasBiColores grdPatologia, sighEntidades.GrillaConFilasBicolor
        grdPatologia.Caption = ""
        If lbHabilitaFrame = True Then
           FraPatologia.Enabled = True
        End If
    End If
    '
    If lbSoloLimpiaBancoS = True Then
        With oRsBanco
              .Fields.Append "Fua", adInteger
              .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
              .Fields.Append "Id", adInteger
              .Fields.Append "Procedimiento", adVarChar, 255, adFldIsNullable
              .Fields.Append "Cantidad", adInteger
              .Fields.Append "idDosisRecetada", adInteger
              .Fields.Append "HayCpt", adBoolean
              .Fields.Append "Precio", adDouble
              .Fields.Append "SaldoActual", adInteger
              .Fields.Append "Receta", adInteger
              .Fields.Append "idEstadoDetalle", adInteger
              .Fields.Append "MotivoAnulacionMedico", adVarChar, 300, adFldIsNullable
              .Fields.Append "Observaciones", adVarChar, 300, adFldIsNullable
              .CursorType = adOpenDynamic
              .LockType = adLockOptimistic
              .Open
        End With
        Set grdBanco.DataSource = oRsBanco
        mo_Apariencia.ConfigurarFilasBiColores grdBanco, sighEntidades.GrillaConFilasBicolor
        grdBanco.Caption = ""
        If lbHabilitaFrame = True Then
           FraBancoS.Enabled = True
        End If
    End If
    '
    If lbSoloLimpiaFarmacia = True Then
        With oRsFarmacia
              .Fields.Append "Fua", adInteger
              .Fields.Append "Id", adInteger
              .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
              .Fields.Append "Procedimiento", adVarChar, 300, adFldIsNullable
              .Fields.Append "Cantidad", adInteger
              .Fields.Append "idDosisRecetada", adInteger
              .Fields.Append "IdViaAdministracion", adInteger 'Actualizado 26092014
              .Fields.Append "HaySaldo", adBoolean
              .Fields.Append "SaldoActual", adInteger
              .Fields.Append "Almacen", adVarChar, 255, adFldIsNullable
              .Fields.Append "IdAlmacen", adInteger
              .Fields.Append "Precio", adDouble
              .Fields.Append "Receta", adInteger
              .Fields.Append "idEstadoDetalle", adInteger
              .Fields.Append "MotivoAnulacionMedico", adVarChar, 300, adFldIsNullable
              .Fields.Append "Observaciones", adVarChar, 300, adFldIsNullable
              .Fields.Append "fechaVigencia", adDate, , adFldIsNullable                   'debb-24/06/2015
              .CursorType = adOpenDynamic
              .LockType = adLockOptimistic
              .Open
        End With
        Set grdFarmacia.DataSource = oRsFarmacia
        mo_Apariencia.ConfigurarFilasBiColores grdFarmacia, sighEntidades.GrillaConFilasBicolor
        grdFarmacia.Caption = ""
        If lbHabilitaFrame = True Then
           FraFarmacia.Enabled = True
        End If
    End If
End Sub

Private Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
    Dim oRsTmp1 As New Recordset
    Dim lnItem As Integer
    Dim lnFor As Integer
    oGrilla.Bands(0).Columns("idEstadoDetalle").Hidden = True
    oGrilla.Bands(0).Columns("MotivoAnulacionMedico").Hidden = True
    Select Case oGrilla.Name
    Case "grdFarmacia"
         'oGrilla.Override.RowSizing = ssRowSizingFixed
         oGrilla.Bands(0).Columns("Dx").Width = 700
         oGrilla.Bands(0).Columns("fechaVigencia").Hidden = True     'debb-24/06/2015
         oGrilla.Bands(0).Columns("Id").Hidden = True
         oGrilla.Bands(0).Columns("SaldoActual").Hidden = True
         oGrilla.Bands(0).Columns("Precio").Hidden = True
         oGrilla.Bands(0).Columns("Almacen").Hidden = True
         oGrilla.Bands(0).Columns("idalmacen").Hidden = True
         oGrilla.Bands(0).Columns("Receta").Hidden = True
         oGrilla.Bands(0).Columns("Procedimiento").Header.Caption = "Medicamento/Insumo"
         oGrilla.Bands(0).Columns("Procedimiento").Width = 4350 '5350 'Actualizado 26092014
         oGrilla.Bands(0).Columns("Procedimiento").Activation = ssActivationActivateNoEdit
         oGrilla.Bands(0).Columns("Cantidad").Width = 650 'Actualizado 26092014
         oGrilla.Bands(0).Columns("Cantidad").Activation = ssActivationAllowEdit
         '
         oGrilla.Bands(0).Columns("idDosisRecetada").Activation = ssActivationActivateNoEdit
         oGrilla.Bands(0).Columns("idDosisRecetada").Width = 700
         oGrilla.Bands(0).Columns("idDosisRecetada").Header.Caption = "N°Dosis"
         Set oRsTmp1 = mo_ReglasComunes.RecetaDosisSelecionarTodos
         lnItem = 0
         With oGrilla.ValueLists.Add("Dosis").ValueListItems
                If oRsTmp1.RecordCount > 0 Then
                   oRsTmp1.MoveFirst
                   Do While Not oRsTmp1.EOF
                      If lnItem = 1 Then
                         lnIdDosisDefault = oRsTmp1.Fields!idDosis
                      End If
                      lnItem = lnItem + 1
                      .Add Trim(Str(oRsTmp1.Fields!idDosis)), oRsTmp1.Fields!numeroDosis
                      oRsTmp1.MoveNext
                   Loop
                Else
                   MsgBox "Falta ingresar DATOS en tabla RecetaDosis"
                End If
         End With
         oRsTmp1.Close
         oGrilla.Bands(0).Columns("idDosisRecetada").ValueList = "Dosis"
         
         oGrilla.Bands(0).Columns("Observaciones").Header.Caption = "Frecuencia" 'Actualizado 01102014
         
        'Actualizado 26092014''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set oRsTmp1 = mo_ReglasComunes.RecetasListadoViasAdministracion
        With oGrilla.ValueLists.Add("ViaAdministracion").ValueListItems
'            grdFarmacia.ValueLists.Add ("ViaAdministracion")
            If oRsTmp1.RecordCount > 0 Then
                oRsTmp1.MoveFirst
                While Not oRsTmp1.EOF
                    .Add CInt(oRsTmp1!IdViaAdministracion), CStr(oRsTmp1!descripcion)
                    oRsTmp1.MoveNext
                Wend
            End If
        End With
        oRsTmp1.Close
        oGrilla.Bands(0).Columns("IdViaAdministracion").ValueList = "ViaAdministracion"
        oGrilla.Bands(0).Columns("IdViaAdministracion").Header.Caption = "Via"
        oGrilla.Bands(0).Columns("IdViaAdministracion").Width = 1400
        '
        '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         '
         '
    Case "grdRayos", "grdEcografiaO", "grdEcografiaG", "grdTomografia"
         oGrilla.Bands(0).Columns("Id").Hidden = True
         oGrilla.Bands(0).Columns("SaldoActual").Hidden = True
         oGrilla.Bands(0).Columns("Precio").Hidden = True
         oGrilla.Bands(0).Columns("Procedimiento").Header.Caption = "Procedimiento"
         oGrilla.Bands(0).Columns("Procedimiento").Width = 2650 '3650
         oGrilla.Bands(0).Columns("Procedimiento").Activation = ssActivationActivateNoEdit
         oGrilla.Bands(0).Columns("Cantidad").Width = 400
         oGrilla.Bands(0).Columns("Cantidad").Activation = ssActivationAllowEdit
         oGrilla.Bands(0).Columns("hayCpt").Activation = ssActivationActivateNoEdit
         oGrilla.Bands(0).Columns("hayCpt").Width = 400
         oGrilla.Bands(0).Columns("Receta").Hidden = True
         oGrilla.Bands(0).Columns("idDosisRecetada").Hidden = True
         '
         'If ml_AScorrelativo > 0 Then
            oGrilla.Bands(0).Columns("fua").Header.Appearance.ForeColor = vbWhite
            oGrilla.Bands(0).Columns("fua").Header.Appearance.BackColor = vbRed
            oGrilla.Bands(0).Columns("Fua").Width = 700
            With oGrilla.ValueLists.Add("FuaList").ValueListItems
                For lnFor = 1 To 20
                    .Add lnFor, "N° " & Trim(Str(lnFor))
                Next
            End With
            oGrilla.Bands(0).Columns("Fua").ValueList = "FuaList"
         'Else
         '   oGrilla.Bands(0).Columns("fua").Hidden = True
         'End If
         '
    Case "grdPatologia", "grdAnatomia", "grdBanco"
         oGrilla.Bands(0).Columns("Id").Hidden = True
         oGrilla.Bands(0).Columns("SaldoActual").Hidden = True
         oGrilla.Bands(0).Columns("Precio").Hidden = True
         oGrilla.Bands(0).Columns("Procedimiento").Header.Caption = "Procedimiento"
         oGrilla.Bands(0).Columns("Procedimiento").Width = 2650 ' 3650
         oGrilla.Bands(0).Columns("Procedimiento").Activation = ssActivationActivateNoEdit
         oGrilla.Bands(0).Columns("Cantidad").Width = 400
         oGrilla.Bands(0).Columns("Cantidad").Activation = ssActivationAllowEdit
         oGrilla.Bands(0).Columns("hayCpt").Activation = ssActivationActivateNoEdit
         oGrilla.Bands(0).Columns("hayCpt").Width = 400
         oGrilla.Bands(0).Columns("Receta").Hidden = True
         oGrilla.Bands(0).Columns("idDosisRecetada").Hidden = True
         
         oGrilla.Bands(0).Columns("fua").Header.Appearance.ForeColor = vbWhite
         oGrilla.Bands(0).Columns("fua").Header.Appearance.BackColor = vbRed
         oGrilla.Bands(0).Columns("Fua").Width = 700
         With oGrilla.ValueLists.Add("FuaList").ValueListItems
            For lnFor = 1 To 20
                .Add lnFor, "N° " & Trim(Str(lnFor))
            Next
         End With
         oGrilla.Bands(0).Columns("Fua").ValueList = "FuaList"
         '
    End Select
    Set oRsTmp1 = Nothing
End Sub

Public Function AlMenosHayUnaReceta() As Boolean
    AlMenosHayUnaReceta = True
    If oRsRayosX.RecordCount = 0 And oRsEcografiaO.RecordCount = 0 And oRsEcografiaG.RecordCount = 0 And _
       oRsTomografia.RecordCount = 0 And oRsAnatomia.RecordCount = 0 And oRsPatologia.RecordCount = 0 And _
       oRsBanco.RecordCount = 0 And oRsFarmacia.RecordCount = 0 Then
       AlMenosHayUnaReceta = False
    End If
End Function


Public Function DevuelveRayosX() As Recordset
    Set DevuelveRayosX = oRsRayosX
End Function
Public Function DevuelveEcografiaO() As Recordset
    Set DevuelveEcografiaO = oRsEcografiaO
End Function
Public Function DevuelveEcografiaG() As Recordset
    Set DevuelveEcografiaG = oRsEcografiaG
End Function
Public Function DevuelveTomografia() As Recordset
    Set DevuelveTomografia = oRsTomografia
End Function
Public Function DevuelveAnatomia() As Recordset
    Set DevuelveAnatomia = oRsAnatomia
End Function
Public Function DevuelvePatologia() As Recordset
    Set DevuelvePatologia = oRsPatologia
End Function
Public Function DevuelveBancoSangre() As Recordset
    Set DevuelveBancoSangre = oRsBanco
End Function
Public Function DevuelveFarmacia() As Recordset
    'debb-24/06/2015
    If oRsFarmacia.RecordCount > 0 Then
       oRsFarmacia.MoveFirst
       Do While Not oRsFarmacia.EOF
          oRsFarmacia!FechaVigencia = CDate(txtFechaVigencia.Text)
          oRsFarmacia.Update
          oRsFarmacia.MoveNext
       Loop
    End If
    '
    Set DevuelveFarmacia = oRsFarmacia
End Function


Public Sub CargaDatosAcontroles(oRsCabeceraRecetas As Recordset, _
                       ByRef lnRecetaRayosX As Long, ByRef lnRecetaEcografiaO As Long, ByRef lnRecetaEcografiaG As Long, _
                       ByRef lnRecetaTomografia As Long, ByRef lnRecetaAnatomiaP As Long, ByRef lnRecetaPatologiaC As Long, _
                       ByRef lnRecetaBancoS As Long, ByRef lnRecetaFarmacia As Long, ByRef lnRecetaOtrosCpt As Long)
       Dim oRsDetalleReceta As New Recordset, oRsTmp1 As New Recordset
       lblRayosX.Caption = ""
       lblEcografiaO.Caption = ""
       lblEcografiaG.Caption = ""
       lblTomografia.Caption = ""
       lblPatologia.Caption = ""
       lblAnatomia.Caption = ""
       lblBancoS.Caption = ""
       lblFarmacia.Caption = ""
       oRsCabeceraRecetas.MoveFirst
       Do While Not oRsCabeceraRecetas.EOF
          Select Case oRsCabeceraRecetas.Fields!idPuntoCarga
          Case sghPtoCargaServicioHospitalizacion
               lnRecetaOtrosCpt = oRsCabeceraRecetas.Fields!idReceta
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  lnRecetaOtrosCpt = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                     If IsNull(oRsCabeceraRecetas.Fields!IdComprobantePago) Then
                        UserControl.ucRecetaCpt1.InhabilitaControles " (Movim: " & oRsCabeceraRecetas.Fields!DocumentoDespacho & ")"
                     Else
                        Set oRsTmp1 = mo_ReglasCaja.CajaComprobantesSeleccionarPorId(oRsCabeceraRecetas.Fields!IdComprobantePago)
                        If oRsTmp1.RecordCount > 0 Then
                           UserControl.ucRecetaCpt1.InhabilitaControles lblPatologia.Caption & " (Boleta: " & Trim(oRsTmp1.Fields!nroSerie) & "-" & Trim(oRsTmp1.Fields!nrodocumento) & ")"
                        End If
                        oRsTmp1.Close
                     End If
                  End Select
               End If
               Set oRsDetalleReceta = mo_ReglasComunes.RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!idReceta, oRsCabeceraRecetas.Fields!idPuntoCarga)
               UserControl.ucRecetaCpt1.CargaDatosAcontroles oRsDetalleReceta
               oRsDetalleReceta.Close
          Case sghPtoCargaRayosX
               lnRecetaRayosX = oRsCabeceraRecetas.Fields!idReceta
               lblRayosX.Caption = "(Receta N° " & Trim(Str(oRsCabeceraRecetas.Fields!idReceta)) & ") (" & mo_ReglasComunes.DevuelveEstadoReceta(oRsCabeceraRecetas.Fields!idEstado) & ")"
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  FraRayos.Enabled = False
                  lnRecetaRayosX = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                     If IsNull(oRsCabeceraRecetas.Fields!IdComprobantePago) Then
                        lblRayosX.Caption = lblRayosX.Caption & " (Movim: " & oRsCabeceraRecetas.Fields!DocumentoDespacho & ")"
                     Else
                        Set oRsTmp1 = mo_ReglasCaja.CajaComprobantesSeleccionarPorId(oRsCabeceraRecetas.Fields!IdComprobantePago)
                        If oRsTmp1.RecordCount > 0 Then
                           lblRayosX.Caption = lblRayosX.Caption & " (Boleta: " & Trim(oRsTmp1.Fields!nroSerie) & "-" & Trim(oRsTmp1.Fields!nrodocumento) & ")"
                        End If
                        oRsTmp1.Close
                     End If
                  End Select
               End If
               Set oRsDetalleReceta = mo_ReglasComunes.RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!idReceta, oRsCabeceraRecetas.Fields!idPuntoCarga)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsRayosX.AddNew
                     oRsRayosX.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsRayosX.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsRayosX.Fields!Cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsRayosX.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsRayosX.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     oRsRayosX.Fields!Receta = oRsCabeceraRecetas.Fields!idReceta
                     oRsRayosX.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsRayosX.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsRayosX.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     oRsRayosX.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     If oRsDetalleReceta.Fields!precio > 0 Then
                        oRsRayosX.Fields!hayCpt = True
                     End If
                     oRsRayosX.Fields!Dx = IIf(IsNull(oRsDetalleReceta!Dx), "", oRsDetalleReceta!Dx)
                     oRsRayosX.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          Case sghPtoCargaEcogObstetrica
               lnRecetaEcografiaO = oRsCabeceraRecetas.Fields!idReceta
               lblEcografiaO.Caption = "(Receta N° " & Trim(Str(oRsCabeceraRecetas.Fields!idReceta)) & ") (Estado: " & mo_ReglasComunes.DevuelveEstadoReceta(oRsCabeceraRecetas.Fields!idEstado) & ")"
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  FraEcografiaO.Enabled = False
                  lnRecetaEcografiaO = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                     If IsNull(oRsCabeceraRecetas.Fields!IdComprobantePago) Then
                        lblEcografiaO.Caption = lblEcografiaO.Caption & " (Movim: " & oRsCabeceraRecetas.Fields!DocumentoDespacho & ")"
                     Else
                        Set oRsTmp1 = mo_ReglasCaja.CajaComprobantesSeleccionarPorId(oRsCabeceraRecetas.Fields!IdComprobantePago)
                        If oRsTmp1.RecordCount > 0 Then
                           lblEcografiaO.Caption = lblEcografiaO.Caption & " (Boleta: " & Trim(oRsTmp1.Fields!nroSerie) & "-" & Trim(oRsTmp1.Fields!nrodocumento) & ")"
                        End If
                        oRsTmp1.Close
                     End If
                  End Select
               End If
               Set oRsDetalleReceta = mo_ReglasComunes.RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!idReceta, oRsCabeceraRecetas.Fields!idPuntoCarga)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsEcografiaO.AddNew
                     oRsEcografiaO.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsEcografiaO.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsEcografiaO.Fields!Cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsEcografiaO.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsEcografiaO.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     If oRsDetalleReceta.Fields!precio > 0 Then
                        oRsEcografiaO.Fields!hayCpt = True
                     End If
                     oRsEcografiaO.Fields!Receta = oRsCabeceraRecetas.Fields!idReceta
                     oRsEcografiaO.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsEcografiaO.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsEcografiaO.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     oRsEcografiaO.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsEcografiaO.Fields!Dx = IIf(IsNull(oRsDetalleReceta!Dx), "", oRsDetalleReceta!Dx)
                     oRsEcografiaO.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          Case sghPtoCargaEcogGeneral
               lnRecetaEcografiaG = oRsCabeceraRecetas.Fields!idReceta
               lblEcografiaG.Caption = "(Receta N° " & Trim(Str(oRsCabeceraRecetas.Fields!idReceta)) & ") (Estado: " & mo_ReglasComunes.DevuelveEstadoReceta(oRsCabeceraRecetas.Fields!idEstado) & ")"
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  FraEcografiaG.Enabled = False
                  lnRecetaEcografiaG = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                     If IsNull(oRsCabeceraRecetas.Fields!IdComprobantePago) Then
                        lblEcografiaG.Caption = lblEcografiaG.Caption & " (Movim: " & oRsCabeceraRecetas.Fields!DocumentoDespacho & ")"
                     Else
                        Set oRsTmp1 = mo_ReglasCaja.CajaComprobantesSeleccionarPorId(oRsCabeceraRecetas.Fields!IdComprobantePago)
                        If oRsTmp1.RecordCount > 0 Then
                           lblEcografiaG.Caption = lblEcografiaG.Caption & " (Boleta: " & Trim(oRsTmp1.Fields!nroSerie) & "-" & Trim(oRsTmp1.Fields!nrodocumento) & ")"
                        End If
                        oRsTmp1.Close
                     End If
                  End Select
               End If
               Set oRsDetalleReceta = mo_ReglasComunes.RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!idReceta, oRsCabeceraRecetas.Fields!idPuntoCarga)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsEcografiaG.AddNew
                     oRsEcografiaG.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsEcografiaG.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsEcografiaG.Fields!Cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsEcografiaG.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsEcografiaG.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     If oRsDetalleReceta.Fields!precio > 0 Then
                        oRsEcografiaG.Fields!hayCpt = True
                     End If
                     oRsEcografiaG.Fields!Receta = oRsCabeceraRecetas.Fields!idReceta
                     oRsEcografiaG.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsEcografiaG.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsEcografiaG.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     oRsEcografiaG.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsEcografiaG.Fields!Dx = IIf(IsNull(oRsDetalleReceta!Dx), "", oRsDetalleReceta!Dx)
                     oRsEcografiaG.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          Case sghPtoCargaTomografia
               lnRecetaTomografia = oRsCabeceraRecetas.Fields!idReceta
               lblTomografia.Caption = "(Receta N° " & Trim(Str(oRsCabeceraRecetas.Fields!idReceta)) & ") (Estado: " & mo_ReglasComunes.DevuelveEstadoReceta(oRsCabeceraRecetas.Fields!idEstado) & ")"
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  FraTomografia.Enabled = False
                  lnRecetaTomografia = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                     If IsNull(oRsCabeceraRecetas.Fields!IdComprobantePago) Then
                        lblTomografia.Caption = lblTomografia.Caption & " (Movim: " & oRsCabeceraRecetas.Fields!DocumentoDespacho & ")"
                     Else
                        Set oRsTmp1 = mo_ReglasCaja.CajaComprobantesSeleccionarPorId(oRsCabeceraRecetas.Fields!IdComprobantePago)
                        If oRsTmp1.RecordCount > 0 Then
                           lblTomografia.Caption = lblTomografia.Caption & " (Boleta: " & Trim(oRsTmp1.Fields!nroSerie) & "-" & Trim(oRsTmp1.Fields!nrodocumento) & ")"
                        End If
                        oRsTmp1.Close
                     End If
                  End Select
               End If
               Set oRsDetalleReceta = mo_ReglasComunes.RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!idReceta, oRsCabeceraRecetas.Fields!idPuntoCarga)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsTomografia.AddNew
                     oRsTomografia.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsTomografia.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsTomografia.Fields!Cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsTomografia.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsTomografia.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     If oRsDetalleReceta.Fields!precio > 0 Then
                        oRsTomografia.Fields!hayCpt = True
                     End If
                     oRsTomografia.Fields!Receta = oRsCabeceraRecetas.Fields!idReceta
                     oRsTomografia.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsTomografia.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsTomografia.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     oRsTomografia.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsTomografia.Fields!Dx = IIf(IsNull(oRsDetalleReceta!Dx), "", oRsDetalleReceta!Dx)
                     oRsTomografia.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          Case sghPtoCargaPatologiaClinica
               lnRecetaPatologiaC = oRsCabeceraRecetas.Fields!idReceta
               lblPatologia.Caption = "(Receta N° " & Trim(Str(oRsCabeceraRecetas.Fields!idReceta)) & ") (Estado: " & mo_ReglasComunes.DevuelveEstadoReceta(oRsCabeceraRecetas.Fields!idEstado) & ")"
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  FraPatologia.Enabled = False
                  lnRecetaPatologiaC = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                     If IsNull(oRsCabeceraRecetas.Fields!IdComprobantePago) Then
                        lblPatologia.Caption = lblPatologia.Caption & " (Movim: " & oRsCabeceraRecetas.Fields!DocumentoDespacho & ")"
                     Else
                        Set oRsTmp1 = mo_ReglasCaja.CajaComprobantesSeleccionarPorId(oRsCabeceraRecetas.Fields!IdComprobantePago)
                        If oRsTmp1.RecordCount > 0 Then
                           lblPatologia.Caption = lblPatologia.Caption & " (Boleta: " & Trim(oRsTmp1.Fields!nroSerie) & "-" & Trim(oRsTmp1.Fields!nrodocumento) & ")"
                        End If
                        oRsTmp1.Close
                     End If
                  End Select
               End If
               Set oRsDetalleReceta = mo_ReglasComunes.RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!idReceta, oRsCabeceraRecetas.Fields!idPuntoCarga)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsPatologia.AddNew
                     oRsPatologia.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsPatologia.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsPatologia.Fields!Cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsPatologia.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsPatologia.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     If oRsDetalleReceta.Fields!precio > 0 Then
                        oRsPatologia.Fields!hayCpt = True
                     End If
                     oRsPatologia.Fields!Receta = oRsCabeceraRecetas.Fields!idReceta
                     oRsPatologia.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsPatologia.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsPatologia.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     oRsPatologia.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsPatologia.Fields!Dx = IIf(IsNull(oRsDetalleReceta!Dx), "", oRsDetalleReceta!Dx)
                     oRsPatologia.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          Case sghPtoCargaAnatomiaPatologica1
               lnRecetaAnatomiaP = oRsCabeceraRecetas.Fields!idReceta
               lblAnatomia.Caption = "(Receta N° " & Trim(Str(oRsCabeceraRecetas.Fields!idReceta)) & ") (Estado: " & mo_ReglasComunes.DevuelveEstadoReceta(oRsCabeceraRecetas.Fields!idEstado) & ")"
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  FraAnatomia.Enabled = False
                  lnRecetaAnatomiaP = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                     If IsNull(oRsCabeceraRecetas.Fields!IdComprobantePago) Then
                        lblAnatomia.Caption = lblAnatomia.Caption & " (Movim: " & oRsCabeceraRecetas.Fields!DocumentoDespacho & ")"
                     Else
                        Set oRsTmp1 = mo_ReglasCaja.CajaComprobantesSeleccionarPorId(oRsCabeceraRecetas.Fields!IdComprobantePago)
                        If oRsTmp1.RecordCount > 0 Then
                           lblAnatomia.Caption = lblAnatomia.Caption & " (Boleta: " & Trim(oRsTmp1.Fields!nroSerie) & "-" & Trim(oRsTmp1.Fields!nrodocumento) & ")"
                        End If
                        oRsTmp1.Close
                     End If
                  End Select
               End If
               Set oRsDetalleReceta = mo_ReglasComunes.RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!idReceta, oRsCabeceraRecetas.Fields!idPuntoCarga)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsAnatomia.AddNew
                     oRsAnatomia.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsAnatomia.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsAnatomia.Fields!Cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsAnatomia.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsAnatomia.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     If oRsDetalleReceta.Fields!precio > 0 Then
                        oRsAnatomia.Fields!hayCpt = True
                     End If
                     oRsAnatomia.Fields!Receta = oRsCabeceraRecetas.Fields!idReceta
                     oRsAnatomia.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsAnatomia.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsAnatomia.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     oRsAnatomia.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsAnatomia.Fields!Dx = IIf(IsNull(oRsDetalleReceta!Dx), "", oRsDetalleReceta!Dx)
                     oRsAnatomia.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          Case sghPtoCargaBancoSangre1
               lnRecetaBancoS = oRsCabeceraRecetas.Fields!idReceta
               lblBancoS.Caption = "(Receta N° " & Trim(Str(oRsCabeceraRecetas.Fields!idReceta)) & ") (Estado: " & mo_ReglasComunes.DevuelveEstadoReceta(oRsCabeceraRecetas.Fields!idEstado) & ")"
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  FraBancoS.Enabled = False
                  lnRecetaBancoS = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                     If IsNull(oRsCabeceraRecetas.Fields!IdComprobantePago) Then
                        lblBancoS.Caption = lblBancoS.Caption & " (Movim: " & oRsCabeceraRecetas.Fields!DocumentoDespacho & ")"
                     Else
                        Set oRsTmp1 = mo_ReglasCaja.CajaComprobantesSeleccionarPorId(oRsCabeceraRecetas.Fields!IdComprobantePago)
                        If oRsTmp1.RecordCount > 0 Then
                           lblBancoS.Caption = lblBancoS.Caption & " (Boleta: " & Trim(oRsTmp1.Fields!nroSerie) & "-" & Trim(oRsTmp1.Fields!nrodocumento) & ")"
                        End If
                        oRsTmp1.Close
                     End If
                  End Select
               End If
               Set oRsDetalleReceta = mo_ReglasComunes.RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!idReceta, oRsCabeceraRecetas.Fields!idPuntoCarga)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsBanco.AddNew
                     oRsBanco.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsBanco.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsBanco.Fields!Cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsBanco.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsBanco.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     If oRsDetalleReceta.Fields!precio > 0 Then
                        oRsBanco.Fields!hayCpt = True
                     End If
                     oRsBanco.Fields!Receta = oRsCabeceraRecetas.Fields!idReceta
                     oRsBanco.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsBanco.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsBanco.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     oRsBanco.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsBanco.Fields!Dx = IIf(IsNull(oRsDetalleReceta!Dx), "", oRsDetalleReceta!Dx)
                     oRsBanco.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          Case sghPtoCargaFarmacia
               'debb-24/06/2015
               If IsNull(oRsCabeceraRecetas.Fields!FechaVigencia) Then
                  txtFechaVigencia.Text = CDate(lcBuscaParametro.RetornaFechaServidorSQL) + lnDiasMaximoVigencia
               Else
                  txtFechaVigencia.Text = oRsCabeceraRecetas.Fields!FechaVigencia
               End If
               '
               lnRecetaFarmacia = oRsCabeceraRecetas.Fields!idReceta
               lblFarmacia.Caption = "(Receta N° " & Trim(Str(oRsCabeceraRecetas.Fields!idReceta)) & ") (Estado: " & mo_ReglasComunes.DevuelveEstadoReceta(oRsCabeceraRecetas.Fields!idEstado) & ")"
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  lnRecetaFarmacia = -100
                  FraFarmacia.Enabled = False
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                     If IsNull(oRsCabeceraRecetas.Fields!IdComprobantePago) Then
                        lblFarmacia.Caption = lblFarmacia.Caption & " (Movim: " & oRsCabeceraRecetas.Fields!DocumentoDespacho & ")"
                     Else
                        Set oRsTmp1 = mo_ReglasCaja.CajaComprobantesSeleccionarPorId(oRsCabeceraRecetas.Fields!IdComprobantePago)
                        If oRsTmp1.RecordCount > 0 Then
                           lblFarmacia.Caption = lblFarmacia.Caption & " (Boleta: " & Trim(oRsTmp1.Fields!nroSerie) & "-" & Trim(oRsTmp1.Fields!nrodocumento) & ")"
                        End If
                        oRsTmp1.Close
                     End If
                  End Select
               End If
               Set oRsDetalleReceta = mo_ReglasComunes.RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!idReceta, oRsCabeceraRecetas.Fields!idPuntoCarga)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsFarmacia.AddNew
                     oRsFarmacia.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsFarmacia.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsFarmacia.Fields!Cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsFarmacia.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsFarmacia.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     If oRsDetalleReceta.Fields!SaldoEnRegistroReceta > 0 Then
                        oRsFarmacia.Fields!haySaldo = True
                     End If
                     oRsFarmacia.Fields!Receta = oRsCabeceraRecetas.Fields!idReceta
                     oRsFarmacia.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsFarmacia.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsFarmacia.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     
                     oRsFarmacia.Fields!IdViaAdministracion = IIf(IsNull(oRsDetalleReceta.Fields!IdViaAdministracion), 0, oRsDetalleReceta.Fields!IdViaAdministracion) 'Actualizado 26092014
                     
                     oRsFarmacia.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsFarmacia.Fields!Dx = IIf(IsNull(oRsDetalleReceta!Dx), "", oRsDetalleReceta!Dx)
                     oRsFarmacia.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          End Select
          oRsCabeceraRecetas.MoveNext
       Loop
       Set oRsTmp1 = Nothing
       Set oRsDetalleReceta = Nothing
End Sub




'Function DevuelvePrecioItem(lnIdProducto As Long, lnIdPuntoCarga As Long, Optional oConexion1 As Connection) As Double
'      Dim oRsTmp As New Recordset
'      If lnIdPuntoCarga = sghPtoCargaFarmacia Then
'         Set oRsTmp = mo_ReglasComunes.FactCatalogoBienesInsumosHospXfiltro("idProducto=" & lnIdProducto & " and idTipoFinanciamiento=" & ml_IdTipoFinanciamiento)
'      Else
'         Set oRsTmp = mo_ReglasComunes.FactCatalogoServiciosHospXfiltro("idProducto=" & lnIdProducto & " and idTipoFinanciamiento=" & ml_IdTipoFinanciamiento)
'      End If
'      DevuelvePrecioItem = 0
'      If oRsTmp.RecordCount > 0 Then
'         DevuelvePrecioItem = oRsTmp.Fields!PrecioUnitario
'      End If
'End Function

Function DevuelveRecetaAntesDeImprimir() As String
    LlenaRecetaParaImpresion
    DevuelveRecetaAntesDeImprimir = txtCitaExClinicos.Text
End Function

Sub LlenaRecetaParaImpresion()
    Dim lcCabecera As String
    txtCitaExClinicos.Text = ""
    'Llenado de datos
    If oRsRayosX.RecordCount > 0 Then
        oRsRayosX.MoveFirst
        lcCabecera = "(Rayos X) (N° Receta: " & Trim(Str(oRsRayosX.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLineaChar & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcCabecera & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsRayosX.EOF
           txtCitaExClinicos.Text = txtCitaExClinicos.Text & Right("000" & Trim(Str(oRsRayosX.Fields!Cantidad)), 4) & " " & oRsRayosX.Fields!procedimiento & " " & Trim(oRsRayosX.Fields!Observaciones) & Chr(13) & Chr(10)
           oRsRayosX.MoveNext
        Loop
    End If
    If oRsEcografiaO.RecordCount > 0 Then
        oRsEcografiaO.MoveFirst
        lcCabecera = "(Ecografía Obstétrica) (N° Receta: " & Trim(Str(oRsEcografiaO.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLineaChar & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcCabecera & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsEcografiaO.EOF
           txtCitaExClinicos.Text = txtCitaExClinicos.Text & Right("000" & Trim(Str(oRsEcografiaO.Fields!Cantidad)), 4) & " " & oRsEcografiaO.Fields!procedimiento & " " & Trim(oRsEcografiaO.Fields!Observaciones) & Chr(13) & Chr(10)
           oRsEcografiaO.MoveNext
        Loop
     End If
    If oRsEcografiaG.RecordCount > 0 Then
        oRsEcografiaG.MoveFirst
        lcCabecera = "(Ecografía General) (N° Receta: " & Trim(Str(oRsEcografiaG.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLineaChar & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcCabecera & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsEcografiaG.EOF
           txtCitaExClinicos.Text = txtCitaExClinicos.Text & Right("000" & Trim(Str(oRsEcografiaG.Fields!Cantidad)), 4) & " " & oRsEcografiaG.Fields!procedimiento & " " & Trim(oRsEcografiaG.Fields!Observaciones) & Chr(13) & Chr(10)
           oRsEcografiaG.MoveNext
        Loop
     End If
    If oRsTomografia.RecordCount > 0 Then
        oRsTomografia.MoveFirst
        lcCabecera = "(Tomografía) (N° Receta: " & Trim(Str(oRsTomografia.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLineaChar & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcCabecera & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsTomografia.EOF
           txtCitaExClinicos.Text = txtCitaExClinicos.Text & Right("000" & Trim(Str(oRsTomografia.Fields!Cantidad)), 4) & " " & oRsTomografia.Fields!procedimiento & " " & Trim(oRsTomografia.Fields!Observaciones) & Chr(13) & Chr(10)
           oRsTomografia.MoveNext
        Loop
     End If
     If oRsAnatomia.RecordCount > 0 Then
        oRsAnatomia.MoveFirst
        lcCabecera = "(Anatomía Patológica) (N° Receta: " & Trim(Str(oRsAnatomia.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLineaChar & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcCabecera & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsAnatomia.EOF
           txtCitaExClinicos.Text = txtCitaExClinicos.Text & Right("000" & Trim(Str(oRsAnatomia.Fields!Cantidad)), 4) & " " & oRsAnatomia.Fields!procedimiento & Chr(13) & Chr(10)
           oRsAnatomia.MoveNext
        Loop
     End If
     If oRsPatologia.RecordCount > 0 Then
        oRsPatologia.MoveFirst
        lcCabecera = "(Patológica Clínica) (N° Receta: " & Trim(Str(oRsPatologia.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLineaChar & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcCabecera & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsPatologia.EOF
           txtCitaExClinicos.Text = txtCitaExClinicos.Text & Right("000" & Trim(Str(oRsPatologia.Fields!Cantidad)), 4) & " " & oRsPatologia.Fields!procedimiento & " " & Trim(oRsPatologia.Fields!Observaciones) & Chr(13) & Chr(10)
           oRsPatologia.MoveNext
        Loop
     End If
     If oRsBanco.RecordCount > 0 Then
        oRsBanco.MoveFirst
        lcCabecera = "(Banco Sangre) (N° Receta: " & Trim(Str(oRsBanco.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLineaChar & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcCabecera & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsBanco.EOF
           txtCitaExClinicos.Text = txtCitaExClinicos.Text & Right("000" & Trim(Str(oRsBanco.Fields!Cantidad)), 4) & " " & oRsBanco.Fields!procedimiento & " " & Trim(oRsBanco.Fields!Observaciones) & Chr(13) & Chr(10)
           oRsBanco.MoveNext
        Loop
     End If
     If oRsFarmacia.RecordCount > 0 Then
        oRsFarmacia.MoveFirst
        lcCabecera = "(Farmacia:" & Trim(oRsFarmacia.Fields!Almacen) & ") (N° Receta: " & Trim(Str(oRsFarmacia.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLineaChar & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcCabecera & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsFarmacia.EOF
           txtCitaExClinicos.Text = txtCitaExClinicos.Text & Right("000" & Trim(Str(oRsFarmacia.Fields!Cantidad)), 4) & " " & oRsFarmacia.Fields!procedimiento & " " & Trim(oRsFarmacia.Fields!Observaciones) & Chr(13) & Chr(10)
           oRsFarmacia.MoveNext
        Loop
     End If
     If lc_Tratamiento <> "" Then
        lcCabecera = "(TRATAMIENTO)" & ml_DatoCabeceraReceta
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLineaChar & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcCabecera & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lcLinea & Chr(13) & Chr(10)
        txtCitaExClinicos.Text = txtCitaExClinicos.Text & lc_Tratamiento
     End If
End Sub

Function LimpiarDatos(lbDesdeModuloMaternoOperinatal As Boolean)
    On Error Resume Next
    If lbDesdeModuloMaternoOperinatal = True Then
        If UserControl.FraRayos.Enabled = True Then
           Set oRsRayosX = Nothing
        End If
        If UserControl.FraEcografiaO.Enabled = True Then
           Set oRsEcografiaO = Nothing
        End If
        If UserControl.FraEcografiaG.Enabled = True Then
           Set oRsEcografiaG = Nothing
        End If
        If UserControl.FraTomografia.Enabled = True Then
           Set oRsTomografia = Nothing
        End If
        If UserControl.FraAnatomia.Enabled = True Then
           Set oRsAnatomia = Nothing
        End If
        If UserControl.FraPatologia.Enabled = True Then
           Set oRsPatologia = Nothing
        End If
        If UserControl.FraBancoS.Enabled = True Then
           Set oRsBanco = Nothing
        End If
        If UserControl.FraFarmacia.Enabled = True Then
           Set oRsFarmacia = Nothing
        End If
        CreaTemporales False, True, True, True, True, True, True, True, True
    Else
        Set oRsRayosX = Nothing
        Set oRsEcografiaO = Nothing
        Set oRsEcografiaG = Nothing
        Set oRsTomografia = Nothing
        Set oRsAnatomia = Nothing
        Set oRsPatologia = Nothing
        Set oRsBanco = Nothing
        Set oRsFarmacia = Nothing
        CreaTemporales True, True, True, True, True, True, True, True, True
        UserControl.ucRecetaCpt1.CreaTemporales True, True
    End If
    '
       lblRayosX.Caption = ""
       lblEcografiaO.Caption = ""
       lblEcografiaG.Caption = ""
       lblTomografia.Caption = ""
       lblPatologia.Caption = ""
       lblAnatomia.Caption = ""
       lblBancoS.Caption = ""
       lblFarmacia.Caption = ""
       '
       txtFechaVigencia.Text = CDate(lcBuscaParametro.RetornaFechaServidorSQL) + lnDiasMaximoVigencia    'debb-03/09/2015
End Function

'Yamill Palomino  2409
Private Sub btnImprimir_Click()
'      UserControl.FraFarmacia.Tag
'      ImprimeOrdenMedica True
      If oRsFarmacia.RecordCount > 0 Then
        oRsFarmacia.MoveFirst
        If Not IsNull(oRsFarmacia.Fields!Receta) Then
            ImprimeOrdenMedicaPorIdReceta oRsFarmacia.Fields!Receta, True
        End If
      End If
'      ImprimeOrdenMedicaPorIdReceta
End Sub

'Actualizado 2609
Private Sub btnImprimirOrden_Click()
'    ImprimeOrdenMedica False
    
    If oRsRayosX.RecordCount > 0 Then
      oRsRayosX.MoveFirst
      If Not IsNull(oRsRayosX.Fields!Receta) Then
          If Not oRsRayosX.Fields!Receta = 0 Then ImprimeOrdenMedicaPorIdReceta oRsRayosX.Fields!Receta, False
      End If
    End If
      
    If oRsEcografiaO.RecordCount > 0 Then
      oRsEcografiaO.MoveFirst
      If Not IsNull(oRsEcografiaO.Fields!Receta) Then
          If Not oRsEcografiaO.Fields!Receta = 0 Then ImprimeOrdenMedicaPorIdReceta oRsEcografiaO.Fields!Receta, False
      End If
    End If
    
    If oRsEcografiaG.RecordCount > 0 Then
      oRsEcografiaG.MoveFirst
      If Not IsNull(oRsEcografiaG.Fields!Receta) Then
          If Not oRsEcografiaG.Fields!Receta = 0 Then ImprimeOrdenMedicaPorIdReceta oRsEcografiaG.Fields!Receta, False
      End If
    End If
    
    If oRsTomografia.RecordCount > 0 Then
      oRsTomografia.MoveFirst
      If Not IsNull(oRsTomografia.Fields!Receta) Then
          If Not oRsTomografia.Fields!Receta = 0 Then ImprimeOrdenMedicaPorIdReceta oRsTomografia.Fields!Receta, False
      End If
    End If
    
    If oRsAnatomia.RecordCount > 0 Then
      oRsAnatomia.MoveFirst
      If Not IsNull(oRsAnatomia.Fields!Receta) Then
          If Not oRsAnatomia.Fields!Receta = 0 Then ImprimeOrdenMedicaPorIdReceta oRsAnatomia.Fields!Receta, False
      End If
    End If
    
    If oRsPatologia.RecordCount > 0 Then
      oRsPatologia.MoveFirst
      If Not IsNull(oRsPatologia.Fields!Receta) Then
          If Not oRsPatologia.Fields!Receta = 0 Then ImprimeOrdenMedicaPorIdReceta oRsPatologia.Fields!Receta, False
      End If
    End If
      
    If oRsBanco.RecordCount > 0 Then
      oRsBanco.MoveFirst
      If Not IsNull(oRsBanco.Fields!Receta) Then
          If Not oRsBanco.Fields!Receta = 0 Then ImprimeOrdenMedicaPorIdReceta oRsBanco.Fields!Receta, False
      End If
    End If
    
End Sub

Sub ImprimeOrdenMedicaPorIdReceta(lnIdReceta As Long, lbEsFarmacia As Boolean)
    Dim oReporte As New RptCaja
    Dim lbImpresionIndirecta As Boolean
    lbImpresionIndirecta = False
    If lcBuscaParametro.SeleccionaFilaParametro(348) = "S" Then lbImpresionIndirecta = True
    If lbEsFarmacia Then
        'Farmacia
        oReporte.imprimirReceta_Version2 lnIdReceta, lbImpresionIndirecta, mi_lnWnd
    Else
        'Examenes Auxiliareas
        oReporte.ImpresionOrdenMedica lnIdReceta, lbImpresionIndirecta
    End If
    Set oReporte = Nothing
 End Sub


'Actualizado 30092014
Sub ImprimeOrdenMedica(lbEsFarmacia As Boolean)
    Dim oReporte As New RptCaja
    Dim oRsCabeceraReceta As Recordset
    Dim oRecetaCabecera As New RecetaCabecera
    Dim oConexion As New ADODB.Connection
    Dim lbImpresionIndirecta As Boolean
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient

    Set oRecetaCabecera.Conexion = oConexion
    Set oRsCabeceraReceta = oRecetaCabecera.SeleccionarPorIdCuentaAtencion(ml_idCuentaAtencion)
        
    lbImpresionIndirecta = False
    If lcBuscaParametro.SeleccionaFilaParametro(348) = "S" Then lbImpresionIndirecta = True
    
    If oRsCabeceraReceta.RecordCount > 0 Then
        If lbEsFarmacia Then
            'Farmacia
            oRsCabeceraReceta.MoveFirst
            oRsCabeceraReceta.Find "IdPuntoCarga=" & 5
            If Not oRsCabeceraReceta.EOF Then
    '            oReporte.ImpresionOrdenMedica oRsCabeceraReceta.Fields!idReceta, True
                oReporte.imprimirReceta_Version2 oRsCabeceraReceta.Fields!idReceta, lbImpresionIndirecta, mi_lnWnd       'Actualizado 26092014
            End If
        Else
            'Rayos X
            oRsCabeceraReceta.MoveFirst
            oRsCabeceraReceta.Find "IdPuntoCarga=" & 21
            If Not oRsCabeceraReceta.EOF Then
                oReporte.ImpresionOrdenMedica oRsCabeceraReceta.Fields!idReceta, lbImpresionIndirecta
            End If
            'Ecografia Obstetrica
            oRsCabeceraReceta.MoveFirst
            oRsCabeceraReceta.Find "IdPuntoCarga=" & 23
            If Not oRsCabeceraReceta.EOF Then
                oReporte.ImpresionOrdenMedica oRsCabeceraReceta.Fields!idReceta, lbImpresionIndirecta
            End If
            'Ecografia General
            oRsCabeceraReceta.MoveFirst
            oRsCabeceraReceta.Find "IdPuntoCarga=" & 20
            If Not oRsCabeceraReceta.EOF Then
                oReporte.ImpresionOrdenMedica oRsCabeceraReceta.Fields!idReceta, lbImpresionIndirecta
            End If
            'Tomografia
            oRsCabeceraReceta.MoveFirst
            oRsCabeceraReceta.Find "IdPuntoCarga=" & 22
            If Not oRsCabeceraReceta.EOF Then
                oReporte.ImpresionOrdenMedica oRsCabeceraReceta.Fields!idReceta, lbImpresionIndirecta
            End If
            'Patologia Clinica
            oRsCabeceraReceta.MoveFirst
            oRsCabeceraReceta.Find "IdPuntoCarga=" & 2
            If Not oRsCabeceraReceta.EOF Then
                oReporte.ImpresionOrdenMedica oRsCabeceraReceta.Fields!idReceta, lbImpresionIndirecta
            End If
            'Anatomia Patologica
            oRsCabeceraReceta.MoveFirst
            oRsCabeceraReceta.Find "IdPuntoCarga=" & 32
            If Not oRsCabeceraReceta.EOF Then
                oReporte.ImpresionOrdenMedica oRsCabeceraReceta.Fields!idReceta, lbImpresionIndirecta
            End If
            'Banco de Sangre
            oRsCabeceraReceta.MoveFirst
            oRsCabeceraReceta.Find "IdPuntoCarga=" & 11
            If Not oRsCabeceraReceta.EOF Then
                oReporte.ImpresionOrdenMedica oRsCabeceraReceta.Fields!idReceta, lbImpresionIndirecta
            End If
        End If
    Else
        MsgBox "No existe ordenes para imprimir", vbInformation, "Ordenes Médicas"
    End If
    

    Set oRsCabeceraReceta = Nothing
    Set oReporte = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub

Public Sub CargaNumeroDeRecetaEimprime(lnRecetaRayosX As Long, lnRecetaEcografiaO As Long, lnRecetaEcografiaG As Long, _
                       lnRecetaTomografia As Long, lnRecetaAnatomiaP As Long, lnRecetaPatologiaC As Long, _
                       lnRecetaBancoS As Long, lnRecetaFarmacia As Long, lbImprimeReceta As Boolean, _
                       lnRecetaOtrosCpt As Long)
       On Error Resume Next
       If lnRecetaOtrosCpt > 0 Then
          Dim oRecetasOtrosCpt As New Recordset
          Dim oRecetaCabecera1 As New RecetaCabecera
          Dim oConexion As New ADODB.Connection
          sighEntidades.AbreConexionSIGH oConexion
          Set oRecetaCabecera1.Conexion = oConexion
          Set oRecetasOtrosCpt = oRecetaCabecera1.SeleccionarPorIdCuentaAtencion(ml_idCuentaAtencion)
          oRecetasOtrosCpt.Filter = "idPuntoCarga = " & sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion
          If oRecetasOtrosCpt.RecordCount > 0 Then
             oRecetasOtrosCpt.MoveFirst
             Do While Not oRecetasOtrosCpt.EOF
               ImprimeOrdenMedicaPorIdReceta oRecetasOtrosCpt!idReceta, False
               oRecetasOtrosCpt.MoveNext
             Loop
          End If
          oRecetasOtrosCpt.Close
          oConexion.Close
          Set oRecetasOtrosCpt = Nothing
          Set oRecetaCabecera1 = Nothing
          Set oConexion = Nothing
       End If
       If lnRecetaRayosX > 0 Then
          oRsRayosX.MoveFirst
          oRsRayosX.Fields!Receta = lnRecetaRayosX
          oRsRayosX.Update
          If lbImprimeReceta = True Then ImprimeOrdenMedicaPorIdReceta lnRecetaRayosX, False
       End If
       If lnRecetaEcografiaO > 0 Then
          oRsEcografiaO.MoveFirst
          oRsEcografiaO.Fields!Receta = lnRecetaEcografiaO
          oRsEcografiaO.Update
          If lbImprimeReceta = True Then ImprimeOrdenMedicaPorIdReceta lnRecetaEcografiaO, False
       End If
       If lnRecetaEcografiaG > 0 Then
          oRsEcografiaG.MoveFirst
          oRsEcografiaG.Fields!Receta = lnRecetaEcografiaG
          oRsEcografiaG.Update
          If lbImprimeReceta = True Then ImprimeOrdenMedicaPorIdReceta lnRecetaEcografiaG, False
       End If
       If lnRecetaTomografia > 0 Then
          oRsTomografia.MoveFirst
          oRsTomografia.Fields!Receta = lnRecetaTomografia
          oRsTomografia.Update
          If lbImprimeReceta = True Then ImprimeOrdenMedicaPorIdReceta lnRecetaTomografia, False
       End If
       If lnRecetaAnatomiaP > 0 Then
          oRsAnatomia.MoveFirst
          oRsAnatomia.Fields!Receta = lnRecetaAnatomiaP
          oRsAnatomia.Update
          If lbImprimeReceta = True Then ImprimeOrdenMedicaPorIdReceta lnRecetaAnatomiaP, False
       End If
       If lnRecetaPatologiaC > 0 Then
          oRsPatologia.MoveFirst
          oRsPatologia.Fields!Receta = lnRecetaPatologiaC
          oRsPatologia.Update
          If lbImprimeReceta = True Then ImprimeOrdenMedicaPorIdReceta lnRecetaPatologiaC, False
       End If
       If lnRecetaBancoS > 0 Then
          oRsBanco.MoveFirst
          oRsBanco.Fields!Receta = lnRecetaBancoS
          oRsBanco.Update
          If lbImprimeReceta = True Then ImprimeOrdenMedicaPorIdReceta lnRecetaBancoS, False
       End If
       If lnRecetaFarmacia > 0 Then
          oRsFarmacia.MoveFirst
          oRsFarmacia.Fields!Receta = lnRecetaFarmacia
          oRsFarmacia.Update
          If lbImprimeReceta = True Then ImprimeOrdenMedicaPorIdReceta lnRecetaFarmacia, True
       End If
       '
'       If lbImprimeReceta = True Then 'Actualizado 30092014
''          btnImprimir_Click
'            ImprimeOrdenMedica True 'Farmacia
'            ImprimeOrdenMedica False 'Ordenes Medicas
'       End If
End Sub


Private Sub chkVerDx_Click()
    If chkVerDx.Value = 1 Then
       grdDiag.Visible = True
    Else
       grdDiag.Visible = False
    End If
End Sub

Private Sub grdAnatomia_AfterRowsDeleted()
   Set grdAnatomia.DataSource = oRsAnatomia
End Sub
Private Sub grdBanco_AfterRowsDeleted()
    Set grdBanco.DataSource = oRsBanco
End Sub

Private Sub grdEcografiaG_AfterRowsDeleted()
   Set grdEcografiaG.DataSource = oRsEcografiaG
End Sub
Private Sub grdEcografiaO_AfterRowsDeleted()
    Set grdEcografiaO.DataSource = oRsEcografiaO
End Sub



Private Sub grdFarmacia_AfterRowsDeleted()
    Set grdFarmacia.DataSource = oRsFarmacia
End Sub






Private Sub grdPatologia_AfterRowsDeleted()
    Set grdPatologia.DataSource = oRsPatologia
End Sub


Private Sub grdRayos_AfterRowsDeleted()
   Set grdRayos.DataSource = oRsRayosX
End Sub
Private Sub grdTomografia_AfterRowsDeleted()
    Set grdTomografia.DataSource = oRsTomografia
End Sub


Public Sub CargaRecetaDesdePerinatal(oRsTmpFarmacia As Recordset, oRsTmpCpt As Recordset)
    Dim lbNuevoRegistro As Boolean, lnNroItems As Long
    Dim oRsTmp1 As New Recordset, lnIdPuntoCarga As Long
    Dim oRsTmp As New Recordset
    Dim lnSaldoActual As Long, lnPrecio As Double, lcAlmacen As String, lnIdAlmacen As Long
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    LimpiarDatos True
    
'    If oRsTmpFarmacia.RecordCount > 0 And UserControl.FraFarmacia.Enabled = True Then
'
'       lnNroItems = oRsFarmacia.RecordCount
'       oRsTmpFarmacia.MoveFirst
'       Do While Not oRsTmpFarmacia.EOF
'          If oRsTmpFarmacia!seleccionar = True Then
'                lbNuevoRegistro = False
'                If lnNroItems > 0 Then
'                   oRsFarmacia.MoveFirst
'                   oRsFarmacia.Find "id=" & oRsTmpFarmacia.Fields!Id
'                   If oRsFarmacia.EOF Then
'                      lbNuevoRegistro = True
'                   End If
'                Else
'                   lbNuevoRegistro = True
'                   lnNroItems = 1
'                End If
'                If lbNuevoRegistro = True Then
'                    Set oRsTmp = mo_ReglasFarmacia.farmSaldoSoloFarmaciasSismed(oRsTmpFarmacia.Fields!Id, oConexion)
'                    lnSaldoActual = 0
'                    lnPrecio = 0
'                    lcAlmacen = ""
'                    If oRsTmp.RecordCount > 0 Then
'                       oRsTmp.MoveFirst
'                       Do While Not oRsTmp.EOF
'                          lnIdAlmacen = oRsTmp.Fields!IdAlmacen
'                          lnSaldoActual = 0
'                          lnPrecio = oRsTmp.Fields!Precio
'                          lcAlmacen = oRsTmp.Fields!Descripcion
'                          Do While Not oRsTmp.EOF And lnIdAlmacen = oRsTmp.Fields!IdAlmacen
'                             lnSaldoActual = lnSaldoActual + oRsTmp.Fields!cantidad
'                             oRsTmp.MoveNext
'                             If oRsTmp.EOF Then
'                                Exit Do
'                             End If
'                          Loop
'                          If lnSaldoActual > 1 Then
'                             Exit Do
'                          End If
'                       Loop
'                   End If
'                   oRsTmp.Close
'                   '
'                   oRsFarmacia.AddNew
'                   oRsFarmacia.Fields!Id = oRsTmpFarmacia.Fields!Id
'                   oRsFarmacia.Fields!procedimiento = oRsTmpFarmacia.Fields!Medicamento
'                   oRsFarmacia.Fields!cantidad = 1
'                   oRsFarmacia.Fields!haySaldo = IIf(lnSaldoActual > 0, True, False)
'                   oRsFarmacia.Fields!saldoActual = lnSaldoActual
'                   oRsFarmacia.Fields!IdAlmacen = lnIdAlmacen
'                   oRsFarmacia.Fields!Almacen = lcAlmacen
'                   oRsFarmacia.Fields!Precio = mo_reglasComunes.DevuelvePrecioItem(oRsTmpFarmacia.Fields!Id, _
'                                                                 sghPtoCargaFarmacia, ml_IdTipoFinanciamiento, oConexion)
'                   oRsFarmacia.Update
'                End If
'          End If
'          oRsTmpFarmacia.MoveNext
'       Loop
'    End If
    If oRsTmpCpt.RecordCount > 0 Then
       oRsTmpCpt.MoveFirst
       Do While Not oRsTmpCpt.EOF
          lbNuevoRegistro = False
          'mgaray201410e
          If mo_ReglasComunes.ProcedimientoEsParaReceta(oRsTmpCpt!ID) = True Then
            Set oRsTmp1 = mo_ReglasComunes.FactCatalogoServiciosPtosSeleccionarXidProducto(oRsTmpCpt!ID, oConexion)
            lnIdPuntoCarga = 0
            If oRsTmp1.RecordCount > 0 Then
               oRsTmp1.MoveFirst
               Do While Not oRsTmp1.EOF
                    lnIdPuntoCarga = oRsTmp1!idPuntoCarga
                    If lnIdPuntoCarga = sghPtoCargaPatologiaClinica Or lnIdPuntoCarga = sghPtoCargaAnatomiaPatologica1 Or _
                       lnIdPuntoCarga = sghPtoCargaBancoSangre1 Or lnIdPuntoCarga = sghPtoCargaRayosX Or _
                       lnIdPuntoCarga = sghPtoCargaEcogObstetrica Or lnIdPuntoCarga = sghPtoCargaEcogGeneral Or _
                       lnIdPuntoCarga = sghPtoCargaTomografia Then
                       Exit Do
                    End If
                    oRsTmp1.MoveNext
               Loop
            End If
            oRsTmp1.Close
            Select Case lnIdPuntoCarga
            Case sghPtoCargaPatologiaClinica
                  If UserControl.FraPatologia.Enabled = True Then
                      lnNroItems = oRsPatologia.RecordCount
                      If lnNroItems > 0 Then
                         oRsPatologia.MoveFirst
                         oRsPatologia.Find "id=" & oRsTmpCpt.Fields!ID
                         If oRsPatologia.EOF Then
                            lbNuevoRegistro = True
                         End If
                      Else
                         lbNuevoRegistro = True
                         lnNroItems = 1
                      End If
                      If lbNuevoRegistro = True Then
                         lnPrecio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpCpt!ID, sghPtoCargaPatologiaClinica, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                         oRsPatologia.AddNew
                         oRsPatologia.Fields!ID = oRsTmpCpt!ID
                         oRsPatologia.Fields!procedimiento = oRsTmpCpt!procedimiento
                         oRsPatologia.Fields!Cantidad = 1
                         oRsPatologia.Fields!precio = lnPrecio
                         If lnPrecio > 0 Then
                             oRsPatologia.Fields!hayCpt = True
                         End If
                         oRsPatologia.Fields!saldoActual = 0
                         oRsPatologia.Update
                      End If
                  End If
            Case sghPtoCargaAnatomiaPatologica1
                  If UserControl.FraAnatomia.Enabled = True Then
                      lnNroItems = oRsAnatomia.RecordCount
                      If lnNroItems > 0 Then
                         oRsAnatomia.MoveFirst
                         oRsAnatomia.Find "id=" & oRsTmpCpt.Fields!ID
                         If oRsAnatomia.EOF Then
                            lbNuevoRegistro = True
                         End If
                      Else
                         lbNuevoRegistro = True
                         lnNroItems = 1
                      End If
                      If lbNuevoRegistro = True Then
                         lnPrecio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpCpt!ID, sghPtoCargaAnatomiaPatologica1, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                         oRsAnatomia.AddNew
                         oRsAnatomia.Fields!ID = oRsTmpCpt!ID
                         oRsAnatomia.Fields!procedimiento = oRsTmpCpt!procedimiento
                         oRsAnatomia.Fields!Cantidad = 1
                         oRsAnatomia.Fields!precio = lnPrecio
                         If lnPrecio > 0 Then
                             oRsAnatomia.Fields!hayCpt = True
                         End If
                         oRsAnatomia.Fields!saldoActual = 0
                         oRsAnatomia.Update
                      End If
                  End If
            Case sghPtoCargaBancoSangre1
                  If UserControl.FraBancoS.Enabled = True Then
                      lnNroItems = oRsBanco.RecordCount
                      If lnNroItems > 0 Then
                         oRsBanco.MoveFirst
                         oRsBanco.Find "id=" & oRsTmpCpt.Fields!ID
                         If oRsBanco.EOF Then
                            lbNuevoRegistro = True
                         End If
                      Else
                         lbNuevoRegistro = True
                         lnNroItems = 1
                      End If
                      If lbNuevoRegistro = True Then
                         lnPrecio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpCpt!ID, sghPtoCargaBancoSangre1, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                         oRsBanco.AddNew
                         oRsBanco.Fields!ID = oRsTmpCpt!ID
                         oRsBanco.Fields!procedimiento = oRsTmpCpt!procedimiento
                         oRsBanco.Fields!Cantidad = 1
                         oRsBanco.Fields!precio = lnPrecio
                         If lnPrecio > 0 Then
                             oRsBanco.Fields!hayCpt = True
                         End If
                         oRsBanco.Fields!saldoActual = 0
                         oRsBanco.Update
                      End If
                  End If
            Case sghPtoCargaRayosX
                  If UserControl.FraRayos.Enabled = True Then
                      lnNroItems = oRsRayosX.RecordCount
                      If lnNroItems > 0 Then
                         oRsRayosX.MoveFirst
                         oRsRayosX.Find "id=" & oRsTmpCpt.Fields!ID
                         If oRsRayosX.EOF Then
                            lbNuevoRegistro = True
                         End If
                      Else
                         lbNuevoRegistro = True
                         lnNroItems = 1
                      End If
                      If lbNuevoRegistro = True Then
                         lnPrecio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpCpt!ID, sghPtoCargaRayosX, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                         oRsRayosX.AddNew
                         oRsRayosX.Fields!ID = oRsTmpCpt!ID
                         oRsRayosX.Fields!procedimiento = oRsTmpCpt!procedimiento
                         oRsRayosX.Fields!Cantidad = 1
                         oRsRayosX.Fields!precio = lnPrecio
                         If lnPrecio > 0 Then
                             oRsRayosX.Fields!hayCpt = True
                         End If
                         oRsRayosX.Fields!saldoActual = 0
                         oRsRayosX.Update
                      End If
                  End If
            Case sghPtoCargaEcogObstetrica
                  If UserControl.FraEcografiaO.Enabled = True Then
                      lnNroItems = oRsEcografiaO.RecordCount
                      If lnNroItems > 0 Then
                         oRsEcografiaO.MoveFirst
                         oRsEcografiaO.Find "id=" & oRsTmpCpt.Fields!ID
                         If oRsEcografiaO.EOF Then
                            lbNuevoRegistro = True
                         End If
                      Else
                         lbNuevoRegistro = True
                         lnNroItems = 1
                      End If
                      If lbNuevoRegistro = True Then
                         lnPrecio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpCpt!ID, sghPtoCargaEcogObstetrica, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                         oRsEcografiaO.AddNew
                         oRsEcografiaO.Fields!ID = oRsTmpCpt!ID
                         oRsEcografiaO.Fields!procedimiento = oRsTmpCpt!procedimiento
                         oRsEcografiaO.Fields!Cantidad = 1
                         oRsEcografiaO.Fields!precio = lnPrecio
                         If lnPrecio > 0 Then
                             oRsEcografiaO.Fields!hayCpt = True
                         End If
                         oRsEcografiaO.Fields!saldoActual = 0
                         oRsEcografiaO.Update
                      End If
                  End If
            Case sghPtoCargaEcogGeneral
                  If UserControl.FraEcografiaG.Enabled = True Then
                      lnNroItems = oRsEcografiaG.RecordCount
                      If lnNroItems > 0 Then
                         oRsEcografiaG.MoveFirst
                         oRsEcografiaG.Find "id=" & oRsTmpCpt.Fields!ID
                         If oRsEcografiaG.EOF Then
                            lbNuevoRegistro = True
                         End If
                      Else
                         lbNuevoRegistro = True
                         lnNroItems = 1
                      End If
                      If lbNuevoRegistro = True Then
                         lnPrecio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpCpt!ID, sghPtoCargaEcogGeneral, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                         oRsEcografiaG.AddNew
                         oRsEcografiaG.Fields!ID = oRsTmpCpt!ID
                         oRsEcografiaG.Fields!procedimiento = oRsTmpCpt!procedimiento
                         oRsEcografiaG.Fields!Cantidad = 1
                         oRsEcografiaG.Fields!precio = lnPrecio
                         If lnPrecio > 0 Then
                             oRsEcografiaG.Fields!hayCpt = True
                         End If
                         oRsEcografiaG.Fields!saldoActual = 0
                         oRsEcografiaG.Update
                      End If
                  End If
            Case sghPtoCargaTomografia
                  If UserControl.FraTomografia.Enabled = True Then
                      lnNroItems = oRsTomografia.RecordCount
                      If lnNroItems > 0 Then
                         oRsTomografia.MoveFirst
                         oRsTomografia.Find "id=" & oRsTmpCpt.Fields!ID
                         If oRsTomografia.EOF Then
                            lbNuevoRegistro = True
                         End If
                      Else
                         lbNuevoRegistro = True
                         lnNroItems = 1
                      End If
                      If lbNuevoRegistro = True Then
                         lnPrecio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpCpt!ID, sghPtoCargaTomografia, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                         oRsTomografia.AddNew
                         oRsTomografia.Fields!ID = oRsTmpCpt!ID
                         oRsTomografia.Fields!procedimiento = oRsTmpCpt!procedimiento
                         oRsTomografia.Fields!Cantidad = 1
                         oRsTomografia.Fields!precio = lnPrecio
                         If lnPrecio > 0 Then
                             oRsTomografia.Fields!hayCpt = True
                         End If
                         oRsTomografia.Fields!saldoActual = 0
                         oRsTomografia.Update
                      End If
                  End If
            End Select
        End If
          oRsTmpCpt.MoveNext
       Loop
    End If
    oConexion.Close
    Set oRsTmp1 = Nothing
    Set oRsTmp = Nothing
    Set oConexion = Nothing
End Sub


Public Sub CargaRecetaDesdeMaterno(oRsTmpFarmacia As Recordset, oRsTmpCpt As Recordset)
    Dim lbNuevoRegistro As Boolean, lnNroItems As Long
    Dim oRsTmp1 As New Recordset, lnIdPuntoCarga As Long
    Dim oRsTmp As New Recordset
    Dim lnSaldoActual As Long, lnPrecio As Double, lcAlmacen As String, lnIdAlmacen As Long
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    LimpiarDatos True
    
    If oRsTmpFarmacia.RecordCount > 0 And UserControl.FraFarmacia.Enabled = True Then
       
       lnNroItems = oRsFarmacia.RecordCount
       oRsTmpFarmacia.MoveFirst
       Do While Not oRsTmpFarmacia.EOF
          'If oRsTmpFarmacia!seleccionar = True Then
                lbNuevoRegistro = False
                If lnNroItems > 0 Then
                   oRsFarmacia.MoveFirst
                   oRsFarmacia.Find "id=" & oRsTmpFarmacia.Fields!idProducto
                   If oRsFarmacia.EOF Then
                      lbNuevoRegistro = True
                   End If
                Else
                   lbNuevoRegistro = True
                   lnNroItems = 1
                End If
                If lbNuevoRegistro = True Then
                    Set oRsTmp = mo_ReglasFarmacia.farmSaldoSoloFarmaciasSismed(oRsTmpFarmacia.Fields!idProducto, oConexion)
                    lnSaldoActual = 0
                    lnPrecio = 0
                    lcAlmacen = ""
                    If oRsTmp.RecordCount > 0 Then
                       oRsTmp.MoveFirst
                       Do While Not oRsTmp.EOF
                          lnIdAlmacen = oRsTmp.Fields!IdAlmacen
                          lnSaldoActual = 0
                          lnPrecio = oRsTmp.Fields!precio
                          lcAlmacen = oRsTmp.Fields!descripcion
                          Do While Not oRsTmp.EOF And lnIdAlmacen = oRsTmp.Fields!IdAlmacen
                             lnSaldoActual = lnSaldoActual + oRsTmp.Fields!Cantidad
                             oRsTmp.MoveNext
                             If oRsTmp.EOF Then
                                Exit Do
                             End If
                          Loop
                          If lnSaldoActual > 1 Then
                             Exit Do
                          End If
                       Loop
                   End If
                   oRsTmp.Close
                   '
                   oRsFarmacia.AddNew
                   oRsFarmacia.Fields!ID = oRsTmpFarmacia.Fields!idProducto
                   oRsFarmacia.Fields!procedimiento = oRsTmpFarmacia.Fields!procedimiento
                   oRsFarmacia.Fields!Cantidad = 1
                   oRsFarmacia.Fields!haySaldo = IIf(lnSaldoActual > 0, True, False)
                   oRsFarmacia.Fields!saldoActual = lnSaldoActual
                   oRsFarmacia.Fields!IdAlmacen = lnIdAlmacen
                   oRsFarmacia.Fields!Almacen = lcAlmacen
                   oRsFarmacia.Fields!precio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpFarmacia.Fields!idProducto, sghPtoCargaFarmacia, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                   oRsFarmacia.Update
                End If
         ' End If
          oRsTmpFarmacia.MoveNext
       Loop
    End If
    If oRsTmpCpt.RecordCount > 0 Then
       oRsTmpCpt.MoveFirst
       Do While Not oRsTmpCpt.EOF
          lbNuevoRegistro = False
          'mgaray201410e
          If mo_ReglasComunes.ProcedimientoEsParaReceta(oRsTmpCpt!idProducto) = True Then
            Set oRsTmp1 = mo_ReglasComunes.FactCatalogoServiciosPtosSeleccionarXidProducto(oRsTmpCpt!idProducto, oConexion)
            lnIdPuntoCarga = 0
            If oRsTmp1.RecordCount > 0 Then
               lnIdPuntoCarga = oRsTmp1!idPuntoCarga
            End If
            oRsTmp1.Close
            Select Case lnIdPuntoCarga
            Case sghPtoCargaPatologiaClinica
                 If UserControl.FraPatologia.Enabled = True Then
                      lnNroItems = oRsPatologia.RecordCount
                      If lnNroItems > 0 Then
                         oRsPatologia.MoveFirst
                         oRsPatologia.Find "id=" & oRsTmpCpt.Fields!idProducto
                         If oRsPatologia.EOF Then
                            lbNuevoRegistro = True
                         End If
                      Else
                         lbNuevoRegistro = True
                         lnNroItems = 1
                      End If
                      If lbNuevoRegistro = True Then
                         lnPrecio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpCpt!idProducto, sghPtoCargaPatologiaClinica, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                         oRsPatologia.AddNew
                         oRsPatologia.Fields!ID = oRsTmpCpt!idProducto
                         oRsPatologia.Fields!procedimiento = oRsTmpCpt!procedimiento
                         oRsPatologia.Fields!Cantidad = 1
                         oRsPatologia.Fields!precio = lnPrecio
                         If lnPrecio > 0 Then
                             oRsPatologia.Fields!hayCpt = True
                         End If
                         oRsPatologia.Fields!saldoActual = 0
                         oRsPatologia.Update
                      End If
                  End If
            Case sghPtoCargaAnatomiaPatologica1
                  If UserControl.FraAnatomia.Enabled = True Then
                      lnNroItems = oRsAnatomia.RecordCount
                      If lnNroItems > 0 Then
                         oRsAnatomia.MoveFirst
                         oRsAnatomia.Find "id=" & oRsTmpCpt.Fields!idProducto
                         If oRsAnatomia.EOF Then
                            lbNuevoRegistro = True
                         End If
                      Else
                         lbNuevoRegistro = True
                         lnNroItems = 1
                      End If
                      If lbNuevoRegistro = True Then
                         lnPrecio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpCpt!idProducto, sghPtoCargaAnatomiaPatologica1, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                         oRsAnatomia.AddNew
                         oRsAnatomia.Fields!ID = oRsTmpCpt!idProducto
                         oRsAnatomia.Fields!procedimiento = oRsTmpCpt!procedimiento
                         oRsAnatomia.Fields!Cantidad = 1
                         oRsAnatomia.Fields!precio = lnPrecio
                         If lnPrecio > 0 Then
                             oRsAnatomia.Fields!hayCpt = True
                         End If
                         oRsAnatomia.Fields!saldoActual = 0
                         oRsAnatomia.Update
                      End If
                  End If
            Case sghPtoCargaBancoSangre1
                  If UserControl.FraBancoS.Enabled = True Then
                      lnNroItems = oRsBanco.RecordCount
                      If lnNroItems > 0 Then
                         oRsBanco.MoveFirst
                         oRsBanco.Find "id=" & oRsTmpCpt.Fields!idProducto
                         If oRsBanco.EOF Then
                            lbNuevoRegistro = True
                         End If
                      Else
                         lbNuevoRegistro = True
                         lnNroItems = 1
                      End If
                      If lbNuevoRegistro = True Then
                         lnPrecio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpCpt!idProducto, sghPtoCargaBancoSangre1, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                         oRsBanco.AddNew
                         oRsBanco.Fields!ID = oRsTmpCpt!idProducto
                         oRsBanco.Fields!procedimiento = oRsTmpCpt!procedimiento
                         oRsBanco.Fields!Cantidad = 1
                         oRsBanco.Fields!precio = lnPrecio
                         If lnPrecio > 0 Then
                             oRsBanco.Fields!hayCpt = True
                         End If
                         oRsBanco.Fields!saldoActual = 0
                         oRsBanco.Update
                      End If
                  End If
            Case sghPtoCargaRayosX
                  If UserControl.FraRayos.Enabled = True Then
                      lnNroItems = oRsRayosX.RecordCount
                      If lnNroItems > 0 Then
                         oRsRayosX.MoveFirst
                         oRsRayosX.Find "id=" & oRsTmpCpt.Fields!idProducto
                         If oRsRayosX.EOF Then
                            lbNuevoRegistro = True
                         End If
                      Else
                         lbNuevoRegistro = True
                         lnNroItems = 1
                      End If
                      If lbNuevoRegistro = True Then
                         lnPrecio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpCpt!idProducto, sghPtoCargaRayosX, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                         oRsRayosX.AddNew
                         oRsRayosX.Fields!ID = oRsTmpCpt!idProducto
                         oRsRayosX.Fields!procedimiento = oRsTmpCpt!procedimiento
                         oRsRayosX.Fields!Cantidad = 1
                         oRsRayosX.Fields!precio = lnPrecio
                         If lnPrecio > 0 Then
                             oRsRayosX.Fields!hayCpt = True
                         End If
                         oRsRayosX.Fields!saldoActual = 0
                         oRsRayosX.Update
                      End If
                  End If
            Case sghPtoCargaEcogObstetrica
                  If UserControl.FraEcografiaO.Enabled = True Then
                      lnNroItems = oRsEcografiaO.RecordCount
                      If lnNroItems > 0 Then
                         oRsEcografiaO.MoveFirst
                         oRsEcografiaO.Find "id=" & oRsTmpCpt.Fields!idProducto
                         If oRsEcografiaO.EOF Then
                            lbNuevoRegistro = True
                         End If
                      Else
                         lbNuevoRegistro = True
                         lnNroItems = 1
                      End If
                      If lbNuevoRegistro = True Then
                         lnPrecio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpCpt!idProducto, sghPtoCargaEcogObstetrica, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                         oRsEcografiaO.AddNew
                         oRsEcografiaO.Fields!ID = oRsTmpCpt!idProducto
                         oRsEcografiaO.Fields!procedimiento = oRsTmpCpt!procedimiento
                         oRsEcografiaO.Fields!Cantidad = 1
                         oRsEcografiaO.Fields!precio = lnPrecio
                         If lnPrecio > 0 Then
                             oRsEcografiaO.Fields!hayCpt = True
                         End If
                         oRsEcografiaO.Fields!saldoActual = 0
                         oRsEcografiaO.Update
                      End If
                  End If
            Case sghPtoCargaEcogGeneral
                  If UserControl.FraEcografiaG.Enabled = True Then
                      lnNroItems = oRsEcografiaG.RecordCount
                      If lnNroItems > 0 Then
                         oRsEcografiaG.MoveFirst
                         oRsEcografiaG.Find "id=" & oRsTmpCpt.Fields!idProducto
                         If oRsEcografiaG.EOF Then
                            lbNuevoRegistro = True
                         End If
                      Else
                         lbNuevoRegistro = True
                         lnNroItems = 1
                      End If
                      If lbNuevoRegistro = True Then
                         lnPrecio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpCpt!idProducto, sghPtoCargaEcogGeneral, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                         oRsEcografiaG.AddNew
                         oRsEcografiaG.Fields!ID = oRsTmpCpt!idProducto
                         oRsEcografiaG.Fields!procedimiento = oRsTmpCpt!procedimiento
                         oRsEcografiaG.Fields!Cantidad = 1
                         oRsEcografiaG.Fields!precio = lnPrecio
                         If lnPrecio > 0 Then
                             oRsEcografiaG.Fields!hayCpt = True
                         End If
                         oRsEcografiaG.Fields!saldoActual = 0
                         oRsEcografiaG.Update
                      End If
                  End If
            Case sghPtoCargaTomografia
                  If UserControl.FraTomografia.Enabled = True Then
                      lnNroItems = oRsTomografia.RecordCount
                      If lnNroItems > 0 Then
                         oRsTomografia.MoveFirst
                         oRsTomografia.Find "id=" & oRsTmpCpt.Fields!idProducto
                         If oRsTomografia.EOF Then
                            lbNuevoRegistro = True
                         End If
                      Else
                         lbNuevoRegistro = True
                         lnNroItems = 1
                      End If
                      If lbNuevoRegistro = True Then
                         lnPrecio = mo_ReglasComunes.DevuelvePrecioItem(oRsTmpCpt!idProducto, sghPtoCargaTomografia, _
                                                                        ml_IdTipoFinanciamiento, oConexion)
                         oRsTomografia.AddNew
                         oRsTomografia.Fields!ID = oRsTmpCpt!idProducto
                         oRsTomografia.Fields!procedimiento = oRsTmpCpt!procedimiento
                         oRsTomografia.Fields!Cantidad = 1
                         oRsTomografia.Fields!precio = lnPrecio
                         If lnPrecio > 0 Then
                             oRsTomografia.Fields!hayCpt = True
                         End If
                         oRsTomografia.Fields!saldoActual = 0
                         oRsTomografia.Update
                      End If
                  End If
            End Select
        End If
          oRsTmpCpt.MoveNext
       Loop
    End If
    oConexion.Close
    Set oRsTmp1 = Nothing
    Set oRsTmp = Nothing
    Set oConexion = Nothing
End Sub

Public Sub OcultarBotonesImpresionReceta(lbOcultar As Boolean)  'Actualizado 21102014
    btnImprimir.Visible = lbOcultar
    btnImprimirOrden.Visible = lbOcultar
End Sub

'mgaray201410d
Public Function DevuelveSoloExamenesParaImpresion()
    Dim sCadenaImprimir As String
    
    Dim lcCabecera As String
    sCadenaImprimir = ""
    'Llenado de datos
    If oRsRayosX.RecordCount > 0 Then
        oRsRayosX.MoveFirst
        lcCabecera = "(Rayos X) (N° Receta: " & Trim(Str(oRsRayosX.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        lcCabecera = "(Rayos X) (N° Receta: " & Trim(Str(oRsRayosX.Fields!Receta)) & ")"
        sCadenaImprimir = sCadenaImprimir & lcLineaChar & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcCabecera & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsRayosX.EOF
           sCadenaImprimir = sCadenaImprimir & Right("000" & Trim(Str(oRsRayosX.Fields!Cantidad)), 4) & " " & oRsRayosX.Fields!procedimiento & " " & Trim(oRsRayosX.Fields!Observaciones) & Chr(13) & Chr(10)
           oRsRayosX.MoveNext
        Loop
    End If
    If oRsEcografiaO.RecordCount > 0 Then
        oRsEcografiaO.MoveFirst
        lcCabecera = "(Ecografía Obstétrica) (N° Receta: " & Trim(Str(oRsEcografiaO.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        lcCabecera = "(Ecografía Obstétrica) (N° Receta: " & Trim(Str(oRsEcografiaO.Fields!Receta)) & ")"
        sCadenaImprimir = sCadenaImprimir & lcLineaChar & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcCabecera & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsEcografiaO.EOF
           sCadenaImprimir = sCadenaImprimir & Right("000" & Trim(Str(oRsEcografiaO.Fields!Cantidad)), 4) & " " & oRsEcografiaO.Fields!procedimiento & " " & Trim(oRsEcografiaO.Fields!Observaciones) & Chr(13) & Chr(10)
           oRsEcografiaO.MoveNext
        Loop
     End If
    If oRsEcografiaG.RecordCount > 0 Then
        oRsEcografiaG.MoveFirst
        lcCabecera = "(Ecografía General) (N° Receta: " & Trim(Str(oRsEcografiaG.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        lcCabecera = "(Ecografía General) (N° Receta: " & Trim(Str(oRsEcografiaG.Fields!Receta)) & ")"
        sCadenaImprimir = sCadenaImprimir & lcLineaChar & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcCabecera & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsEcografiaG.EOF
           sCadenaImprimir = sCadenaImprimir & Right("000" & Trim(Str(oRsEcografiaG.Fields!Cantidad)), 4) & " " & oRsEcografiaG.Fields!procedimiento & " " & Trim(oRsEcografiaG.Fields!Observaciones) & Chr(13) & Chr(10)
           oRsEcografiaG.MoveNext
        Loop
     End If
    If oRsTomografia.RecordCount > 0 Then
        oRsTomografia.MoveFirst
        lcCabecera = "(Tomografía) (N° Receta: " & Trim(Str(oRsTomografia.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        lcCabecera = "(Tomografía) (N° Receta: " & Trim(Str(oRsTomografia.Fields!Receta)) & ")"
        sCadenaImprimir = sCadenaImprimir & lcLineaChar & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcCabecera & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsTomografia.EOF
           sCadenaImprimir = sCadenaImprimir & Right("000" & Trim(Str(oRsTomografia.Fields!Cantidad)), 4) & " " & oRsTomografia.Fields!procedimiento & " " & Trim(oRsTomografia.Fields!Observaciones) & Chr(13) & Chr(10)
           oRsTomografia.MoveNext
        Loop
     End If
     If oRsAnatomia.RecordCount > 0 Then
        oRsAnatomia.MoveFirst
        lcCabecera = "(Anatomía Patológica) (N° Receta: " & Trim(Str(oRsAnatomia.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        lcCabecera = "(Anatomía Patológica) (N° Receta: " & Trim(Str(oRsAnatomia.Fields!Receta)) & ")"
        sCadenaImprimir = sCadenaImprimir & lcLineaChar & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcCabecera & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsAnatomia.EOF
           sCadenaImprimir = sCadenaImprimir & Right("000" & Trim(Str(oRsAnatomia.Fields!Cantidad)), 4) & " " & oRsAnatomia.Fields!procedimiento & Chr(13) & Chr(10)
           oRsAnatomia.MoveNext
        Loop
     End If
     If oRsPatologia.RecordCount > 0 Then
        oRsPatologia.MoveFirst
        lcCabecera = "(Patológica Clínica) (N° Receta: " & Trim(Str(oRsPatologia.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        lcCabecera = "(Patológica Clínica) (N° Receta: " & Trim(Str(oRsPatologia.Fields!Receta)) & ")"
        sCadenaImprimir = sCadenaImprimir & lcLineaChar & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcCabecera & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsPatologia.EOF
           sCadenaImprimir = sCadenaImprimir & Right("000" & Trim(Str(oRsPatologia.Fields!Cantidad)), 4) & " " & oRsPatologia.Fields!procedimiento & " " & Trim(oRsPatologia.Fields!Observaciones) & Chr(13) & Chr(10)
           oRsPatologia.MoveNext
        Loop
     End If
     If oRsBanco.RecordCount > 0 Then
        oRsBanco.MoveFirst
        lcCabecera = "(Banco Sangre) (N° Receta: " & Trim(Str(oRsBanco.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        lcCabecera = "(Banco Sangre) (N° Receta: " & Trim(Str(oRsBanco.Fields!Receta)) & ")"
        sCadenaImprimir = sCadenaImprimir & lcLineaChar & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcCabecera & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsBanco.EOF
           sCadenaImprimir = sCadenaImprimir & Right("000" & Trim(Str(oRsBanco.Fields!Cantidad)), 4) & " " & oRsBanco.Fields!procedimiento & " " & Trim(oRsBanco.Fields!Observaciones) & Chr(13) & Chr(10)
           oRsBanco.MoveNext
        Loop
     End If
     DevuelveSoloExamenesParaImpresion = sCadenaImprimir
End Function

Public Function DevuelveSoloRecetaParaImpresion()
    Dim sCadenaImprimir As String
    
    Dim lcCabecera As String
    sCadenaImprimir = ""
    'Llenado de datos
     If oRsFarmacia.RecordCount > 0 Then
        oRsFarmacia.MoveFirst
        lcCabecera = "(Farmacia:" & Trim(oRsFarmacia.Fields!Almacen) & _
                     ") (N° Receta: " & Trim(Str(oRsFarmacia.Fields!Receta)) & ")" & ml_DatoCabeceraReceta
        lcCabecera = "(Farmacia:" & Trim(oRsFarmacia.Fields!Almacen) & _
                     ") (N° Receta: " & Trim(Str(oRsFarmacia.Fields!Receta)) & ")"
        sCadenaImprimir = sCadenaImprimir & lcLineaChar & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcCabecera & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        Do While Not oRsFarmacia.EOF
           sCadenaImprimir = sCadenaImprimir & Right("000" & Trim(Str(oRsFarmacia.Fields!Cantidad)), 4) & " " & oRsFarmacia.Fields!procedimiento & " " & Trim(oRsFarmacia.Fields!Observaciones) & Chr(13) & Chr(10)
           oRsFarmacia.MoveNext
        Loop
     End If
     If lc_Tratamiento <> "" Then
        lcCabecera = "(TRATAMIENTO)" & ml_DatoCabeceraReceta
        sCadenaImprimir = sCadenaImprimir & lcLineaChar & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcCabecera & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lcLinea & Chr(13) & Chr(10)
        sCadenaImprimir = sCadenaImprimir & lc_Tratamiento
     End If
     
     DevuelveSoloRecetaParaImpresion = sCadenaImprimir
End Function

'debb-11/04/2016
Private Sub cmdPaquetes_Click()
    Dim oPaquetesBuscar As New SIGHNegocios.BuscaPaquetes
    Dim oRsTmp1 As New Recordset
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim oRsItems As New Recordset
    Dim lnIdFactPaquete As Long
    oPaquetesBuscar.MostrarFormulario
    If oPaquetesBuscar.BotonPresionado = sghAceptar Then
        

        'LimpiarDatos False
        lnIdFactPaquete = oPaquetesBuscar.idFactPaquete
        Set oRsTmp1 = mo_ReglasFacturacion.FacturacionCatalogoPaquetesXpaquete(lnIdFactPaquete)
        Set oRsItems = mo_sighProxies.CargaDatosEnTemporalesRecetas(1, oRsTmp1)
        

        If FraPatologia.Enabled = True Then
            oRsItems.Filter = "idPuntoCarga=" & sghPtoCargaPatologiaClinica
            If oRsItems.RecordCount > 0 Then
               Set oRsPatologia = Nothing
               CreaTemporales True, False, False, False, False, False, False, True, False
               mo_sighProxies.AgregaItemsDeReceta sghPtoCargaPatologiaClinica, oRsItems, oRsPatologia, _
                                oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                                lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, 0, True, oRsPatologia, Nothing, lcDxUnico
            End If
        End If
        If FraAnatomia.Enabled = True Then
            oRsItems.Filter = "idPuntoCarga=" & sghPtoCargaAnatomiaPatologica1
            If oRsItems.RecordCount > 0 Then
               Set oRsAnatomia = Nothing
               CreaTemporales True, False, False, False, False, True, False, False, False
               mo_sighProxies.AgregaItemsDeReceta sghPtoCargaAnatomiaPatologica1, oRsItems, oRsPatologia, _
                                oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                                lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, 0, True, oRsAnatomia, Nothing, lcDxUnico
            End If
        End If
        oRsItems.Filter = "idPuntoCarga=" & sghPtoCargaBancoSangre1
        If oRsItems.RecordCount > 0 Then
           Set oRsBanco = Nothing
           CreaTemporales True, False, False, False, False, False, True, False, False
           mo_sighProxies.AgregaItemsDeReceta sghPtoCargaBancoSangre1, oRsItems, oRsPatologia, _
                                oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                                lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, 0, True, oRsBanco, Nothing, lcDxUnico
        End If
        If FraRayos.Enabled = True Then
            oRsItems.Filter = "idPuntoCarga=" & sghPtoCargaRayosX
            If oRsItems.RecordCount > 0 Then
               Set oRsRayosX = Nothing
               CreaTemporales True, False, False, True, False, False, False, False, False
               mo_sighProxies.AgregaItemsDeReceta sghPtoCargaRayosX, oRsItems, oRsPatologia, _
                                oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                                lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, 0, True, oRsRayosX, Nothing, lcDxUnico
            End If
        End If
        If FraEcografiaO.Enabled = True Then
            oRsItems.Filter = "idPuntoCarga=" & sghPtoCargaEcogObstetrica
            If oRsItems.RecordCount > 0 Then
               Set oRsEcografiaO = Nothing
               CreaTemporales True, True, False, False, False, False, False, False, False
               mo_sighProxies.AgregaItemsDeReceta sghPtoCargaEcogObstetrica, oRsItems, oRsPatologia, _
                                oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                                lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, 0, True, oRsEcografiaO, Nothing, lcDxUnico
            End If
        End If
        If FraEcografiaG.Enabled = True Then
            oRsItems.Filter = "idPuntoCarga=" & sghPtoCargaEcogGeneral
            If oRsItems.RecordCount > 0 Then
               Set oRsEcografiaG = Nothing
               CreaTemporales True, False, True, False, False, False, False, False, False
               mo_sighProxies.AgregaItemsDeReceta sghPtoCargaEcogGeneral, oRsItems, oRsPatologia, _
                                oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                                 lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, 0, True, oRsEcografiaG, Nothing, lcDxUnico
            End If
        End If
        If FraTomografia.Enabled = True Then
            oRsItems.Filter = "idPuntoCarga=" & sghPtoCargaTomografia
            If oRsItems.RecordCount > 0 Then
               Set oRsTomografia = Nothing
               CreaTemporales True, False, False, False, True, False, False, False, False
               mo_sighProxies.AgregaItemsDeReceta sghPtoCargaTomografia, oRsItems, oRsPatologia, _
                                oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                                lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, 0, True, oRsTomografia, Nothing, lcDxUnico
            End If
        End If
        If FraFarmacia.Enabled = True Then
           oRsItems.Filter = "idPuntoCarga=" & sghPtoCargaFarmacia
           If oRsItems.RecordCount > 0 Then
              Set oRsFarmacia = Nothing
              CreaTemporales True, False, False, False, False, False, False, False, True
              mo_sighProxies.AgregaItemsDeReceta sghPtoCargaFarmacia, oRsItems, oRsPatologia, _
                                oRsAnatomia, oRsBanco, oRsRayosX, oRsEcografiaO, oRsEcografiaG, oRsTomografia, oRsFarmacia, _
                                lnMaximoItems, lnIdDosisDefault, ml_IdTipoFinanciamiento, True, lnIdFarmaciaElegida, True, _
                                Nothing, oRsFarmacia, lcDxUnico
           End If
        End If
    End If
    Set oPaquetesBuscar = Nothing
    Set oRsTmp1 = Nothing
    Set oRsItems = Nothing
    Set mo_ReglasFacturacion = Nothing
End Sub
'debb-14/07/2015
Sub CargaFarmaciasAelegir()
    Dim oFarmAlmacen As New SIGHDatos.FarmAlmacen
    Dim oRsTmp1 As New Recordset
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    cmdFarmacias.AddItem "Muestra todos los ITEMS"
    cmdFarmacias.AddItem "Muestra sólo los q tienen SALDOS mayores a CERO"
    Set oFarmAlmacen.Conexion = oConexion
    'SCCQ 05/11/2020 Cambio40 Inicio
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim lcParametro601 As String
    lcParametro601 = Trim(lcBuscaParametro.SeleccionaFilaParametro(601))
    If lcParametro601 > 0 Then
        Set oRsTmp1 = oFarmAlmacen.SeleccionarSegunFiltro("idTipoLocales='F' and idTipoSuministro='01' and idEstado=1 and idAlmacen=" + lcParametro601)
        cmdFarmacias.Enabled = False
    Else
        cmdFarmacias.Enabled = True
   'SCCQ 05/11/2020 Cambio40 Fin
        Set oRsTmp1 = oFarmAlmacen.SeleccionarSegunFiltro("idTipoLocales='F' and idTipoSuministro='01' and idEstado=1")
   'SCCQ 05/11/2020 Cambio40 Inicio
    End If
   'SCCQ 05/11/2020 Cambio40 Fin
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          cmdFarmacias.AddItem Trim(Str(oRsTmp1.Fields!IdAlmacen)) & " - " & oRsTmp1.Fields!descripcion
          oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    '
    
    lnIdFarmaciaElegida = Val(sighEntidades.IdFarmaciaParaReceta)
    If lnIdFarmaciaElegida > 1 Then
       Dim lnFor As Integer, lcFarmacia As String
       For lnFor = 2 To cmdFarmacias.ListCount - 1
           lcFarmacia = cmdFarmacias.List(lnFor)
           If Val(Left(lcFarmacia, InStr(lcFarmacia, "-") - 1)) = lnIdFarmaciaElegida Then
              cmdFarmacias.ListIndex = lnFor
              Exit For
           End If
       Next
    Else
       cmdFarmacias.ListIndex = lnIdFarmaciaElegida
    End If
    'SCCQ 05/11/2020 Cambio40 Inicio
    If lcParametro601 > 0 Then
        cmdFarmacias.ListIndex = 2
    End If
    Set lcBuscaParametro = Nothing
   'SCCQ 05/11/2020 Cambio40 Fin
    Set oFarmAlmacen = Nothing
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    
End Sub
'debb-14/07/2015
Private Sub cmdFarmacias_Click()
    If cmdFarmacias.ListIndex > 1 Then
       lnIdFarmaciaElegida = Val(Left(cmdFarmacias.Text, InStr(cmdFarmacias.Text, "-") - 2))
    Else
       lnIdFarmaciaElegida = cmdFarmacias.ListIndex
    End If
    sighEntidades.IdFarmaciaParaReceta = Trim(Str(lnIdFarmaciaElegida))
    On Error Resume Next
    If oRsFarmacia.RecordCount > 0 Then
       oRsFarmacia.MoveFirst
       Do While Not oRsFarmacia.EOF
          oRsFarmacia.Delete
          oRsFarmacia.Update
          oRsFarmacia.MoveNext
       Loop
    End If
    Set grdFarmacia.DataSource = oRsFarmacia
End Sub

Sub ActualizaDxEnGrilla(oRsDx As Recordset)
    On Error Resume Next
    lcDxUnico = ""
    btnAddFarmacia.Enabled = True
    grdFarmacia.ValueLists.Remove ("DxPrincipal1")
    grdPatologia.ValueLists.Remove ("DxPrincipal2")
    grdRayos.ValueLists.Remove ("DxPrincipal3")
    grdEcografiaO.ValueLists.Remove ("DxPrincipal4")
    grdEcografiaG.ValueLists.Remove ("DxPrincipal5")
    grdTomografia.ValueLists.Remove ("DxPrincipal6")
    grdAnatomia.ValueLists.Remove ("DxPrincipal7")
    grdBanco.ValueLists.Remove ("DxPrincipal8")
    If oRsDx.RecordCount > 0 Then
        With grdFarmacia.ValueLists.Add("DxPrincipal1").ValueListItems
           oRsDx.MoveFirst
           Do While Not oRsDx.EOF
              .Add Trim(oRsDx!CodigoCIE2004), oRsDx!CodigoCIE2004
              oRsDx.MoveNext
           Loop
        End With
        grdFarmacia.Bands(0).Columns("dx").ValueList = "DxPrincipal1"
        grdFarmacia.Bands(0).Columns("dx").Width = 1000
        grdFarmacia.Bands(0).Columns("fua").Hidden = True
        grdFarmacia.Bands(0).Columns("dx").ButtonDisplayStyle = ssButtonDisplayStyleAlways
        grdFarmacia.Bands(0).Columns("idDosisRecetada").ButtonDisplayStyle = ssButtonDisplayStyleAlways
        grdFarmacia.Bands(0).Columns("IdViaAdministracion").ButtonDisplayStyle = ssButtonDisplayStyleAlways
        
        '
        With grdPatologia.ValueLists.Add("DxPrincipal2").ValueListItems
           oRsDx.MoveFirst
           Do While Not oRsDx.EOF
              .Add Trim(oRsDx!CodigoCIE2004), oRsDx!CodigoCIE2004
              oRsDx.MoveNext
           Loop
        End With
        grdPatologia.Bands(0).Columns("dx").ValueList = "DxPrincipal2"
        grdPatologia.Bands(0).Columns("dx").Width = 1000
        grdPatologia.Bands(0).Columns("fua").Hidden = True
        grdPatologia.Bands(0).Columns("dx").ButtonDisplayStyle = ssButtonDisplayStyleAlways
        '
        With grdRayos.ValueLists.Add("DxPrincipal3").ValueListItems
           oRsDx.MoveFirst
           Do While Not oRsDx.EOF
              .Add Trim(oRsDx!CodigoCIE2004), oRsDx!CodigoCIE2004
              oRsDx.MoveNext
           Loop
        End With
        grdRayos.Bands(0).Columns("dx").ValueList = "DxPrincipal3"
        grdRayos.Bands(0).Columns("dx").Width = 1000
        grdRayos.Bands(0).Columns("fua").Hidden = True
        grdRayos.Bands(0).Columns("dx").ButtonDisplayStyle = ssButtonDisplayStyleAlways
        '
        With grdEcografiaO.ValueLists.Add("DxPrincipal4").ValueListItems
           oRsDx.MoveFirst
           Do While Not oRsDx.EOF
              .Add Trim(oRsDx!CodigoCIE2004), oRsDx!CodigoCIE2004
              oRsDx.MoveNext
           Loop
        End With
        grdEcografiaO.Bands(0).Columns("dx").ValueList = "DxPrincipal4"
        grdEcografiaO.Bands(0).Columns("dx").Width = 1000
        grdEcografiaO.Bands(0).Columns("fua").Hidden = True
        grdEcografiaO.Bands(0).Columns("dx").ButtonDisplayStyle = ssButtonDisplayStyleAlways
        
        '
        With grdEcografiaG.ValueLists.Add("DxPrincipal5").ValueListItems
           oRsDx.MoveFirst
           Do While Not oRsDx.EOF
              .Add Trim(oRsDx!CodigoCIE2004), oRsDx!CodigoCIE2004
              oRsDx.MoveNext
           Loop
        End With
        grdEcografiaG.Bands(0).Columns("dx").ValueList = "DxPrincipal5"
        grdEcografiaG.Bands(0).Columns("dx").Width = 1000
        grdEcografiaG.Bands(0).Columns("fua").Hidden = True
        grdEcografiaG.Bands(0).Columns("dx").ButtonDisplayStyle = ssButtonDisplayStyleAlways
        '
        With grdTomografia.ValueLists.Add("DxPrincipal6").ValueListItems
           oRsDx.MoveFirst
           Do While Not oRsDx.EOF
              .Add Trim(oRsDx!CodigoCIE2004), oRsDx!CodigoCIE2004
              oRsDx.MoveNext
           Loop
        End With
        grdTomografia.Bands(0).Columns("dx").ValueList = "DxPrincipal6"
        grdTomografia.Bands(0).Columns("dx").Width = 1000
        grdTomografia.Bands(0).Columns("fua").Hidden = True
        grdTomografia.Bands(0).Columns("dx").ButtonDisplayStyle = ssButtonDisplayStyleAlways
        '
        With grdAnatomia.ValueLists.Add("DxPrincipal7").ValueListItems
           oRsDx.MoveFirst
           Do While Not oRsDx.EOF
              .Add Trim(oRsDx!CodigoCIE2004), oRsDx!CodigoCIE2004
              oRsDx.MoveNext
           Loop
        End With
        grdAnatomia.Bands(0).Columns("dx").ValueList = "DxPrincipal7"
        grdAnatomia.Bands(0).Columns("dx").Width = 1000
        grdAnatomia.Bands(0).Columns("fua").Hidden = True
        grdAnatomia.Bands(0).Columns("dx").ButtonDisplayStyle = ssButtonDisplayStyleAlways
        '
        With grdBanco.ValueLists.Add("DxPrincipal8").ValueListItems
           oRsDx.MoveFirst
           Do While Not oRsDx.EOF
              .Add Trim(oRsDx!CodigoCIE2004), oRsDx!CodigoCIE2004
              oRsDx.MoveNext
           Loop
        End With
        grdBanco.Bands(0).Columns("dx").ValueList = "DxPrincipal8"
        grdBanco.Bands(0).Columns("dx").Width = 1000
        grdBanco.Bands(0).Columns("fua").Hidden = True
        grdBanco.Bands(0).Columns("dx").ButtonDisplayStyle = ssButtonDisplayStyleAlways
        '
        If wxParametro545 <> "S" Then
            oRsDx.MoveFirst
            lcDxUnico = oRsDx!CodigoCIE2004
        Else
            If oRsDx.RecordCount = 1 Then
               oRsDx.MoveFirst
               lcDxUnico = oRsDx!CodigoCIE2004
            End If
        End If
    Else
        btnAddFarmacia.Enabled = False
        MsgBox "Debe registrar DIAGNOSTICOS antes de AGREGAR MEDICAMENTOS/INSUMOS", vbInformation, ""
    End If
    Set grdDiag.DataSource = oRsDx
    mo_Apariencia.ConfigurarFilasBiColores grdDiag, sighEntidades.GrillaConFilasBicolor
         grdDiag.Bands(0).Columns("idCuentaAtencion").Hidden = True
         grdDiag.Bands(0).Columns("idTipoDiagnostico").Hidden = True
         grdDiag.Bands(0).Columns("DescripcionTipoDx").Hidden = True
         grdDiag.Bands(0).Columns("idDiagnostico").Hidden = True
         grdDiag.Bands(0).Columns("labConfHIS").Hidden = True
         grdDiag.Bands(0).Columns("CodigoCie2004").Header.Caption = "Dx"
         grdDiag.Bands(0).Columns("CodigoCie2004").Width = 800
         grdDiag.Bands(0).Columns("Descripcion").Width = 3000
    UserControl.ucRecetaCpt1.Dx = lcDxUnico
End Sub


Function ValidaReglas() As Boolean
  ValidaReglas = False
  If wxParametro545 = "S" Then
     If oRsRayosX.RecordCount > 0 Then
        oRsRayosX.MoveFirst
        Do While Not oRsRayosX.EOF
           If Len(oRsRayosX!Dx) = 0 Or IsNull(oRsRayosX!Dx) Then
              MsgBox "Tiene que ingresar DX para cada item de RAYOS  (Ficha 3.3)", vbInformation, ""
              Exit Function
           End If
           oRsRayosX.MoveNext
        Loop
     End If
     If oRsEcografiaO.RecordCount > 0 Then
        oRsEcografiaO.MoveFirst
        Do While Not oRsEcografiaO.EOF
           If Len(oRsEcografiaO!Dx) = 0 Or IsNull(oRsEcografiaO!Dx) Then
              MsgBox "Tiene que ingresar DX para cada  item de ECOGRAFIA OBSTETRICA  (Ficha 3.3)", vbInformation, ""
              Exit Function
           End If
           oRsEcografiaO.MoveNext
        Loop
     End If
     If oRsEcografiaG.RecordCount > 0 Then
        oRsEcografiaG.MoveFirst
        Do While Not oRsEcografiaG.EOF
           If Len(oRsEcografiaG!Dx) = 0 Or IsNull(oRsEcografiaG!Dx) Then
              MsgBox "Tiene que ingresar DX para cada  item de ECOGRAFIA GENERAL  (Ficha 3.3)", vbInformation, ""
              Exit Function
           End If
           oRsEcografiaG.MoveNext
        Loop
     End If
     If oRsTomografia.RecordCount > 0 Then
        oRsTomografia.MoveFirst
        Do While Not oRsTomografia.EOF
           If Len(oRsTomografia!Dx) = 0 Or IsNull(oRsTomografia!Dx) Then
              MsgBox "Tiene que ingresar DX para cada  item de TOMOGRAFIA  (Ficha 3.3)", vbInformation, ""
              Exit Function
           End If
           oRsTomografia.MoveNext
        Loop
     End If
     If oRsAnatomia.RecordCount > 0 Then
        oRsAnatomia.MoveFirst
        Do While Not oRsAnatomia.EOF
           If Len(oRsAnatomia!Dx) = 0 Or IsNull(oRsAnatomia!Dx) Then
              MsgBox "Tiene que ingresar DX para cada  item de ANATOMIA PATOLOGICA  (Ficha 3.3)", vbInformation, ""
              Exit Function
           End If
           oRsAnatomia.MoveNext
        Loop
     End If
     If oRsPatologia.RecordCount > 0 Then
        oRsPatologia.MoveFirst
        Do While Not oRsPatologia.EOF
           If Len(oRsPatologia!Dx) = 0 Or IsNull(oRsPatologia!Dx) Then
              MsgBox "Tiene que ingresar DX para cada  item de PATOLOGIA CLINICA (Ficha 3.3)", vbInformation, ""
              Exit Function
           End If
           oRsPatologia.MoveNext
        Loop
     End If
     If oRsBanco.RecordCount > 0 Then
        oRsBanco.MoveFirst
        Do While Not oRsBanco.EOF
           If Len(oRsBanco!Dx) = 0 Or IsNull(oRsBanco!Dx) Then
              MsgBox "Tiene que ingresar DX para cada  item de BANCO DE SANGRE  (Ficha 3.3)", vbInformation, ""
              Exit Function
           End If
           oRsBanco.MoveNext
        Loop
     End If
     If oRsFarmacia.RecordCount > 0 Then
        oRsFarmacia.MoveFirst
        Do While Not oRsFarmacia.EOF
           If Len(oRsFarmacia!Dx) = 0 Or IsNull(oRsFarmacia!Dx) Then
              MsgBox "Tiene que ingresar DX para cada Medicamento/Insumo  (Ficha 3.3)", vbInformation, ""
              Exit Function
           End If
           oRsFarmacia.MoveNext
        Loop
     End If
  End If
  '
  If sighEntidades.Parametro551 = "S" Then
    If oRsTomografia.RecordCount > 0 Then
         oRsTomografia.MoveFirst
         Do While Not oRsTomografia.EOF
            If IsNull(oRsTomografia!Observaciones) Then
               MsgBox "Tiene que ingresar OBSERVACIONES para cada  item de TOMOGRAFIA (mínimo 50 caracteres)", vbInformation, ""
               Exit Function
            ElseIf Len(oRsTomografia!Observaciones) < 50 Then
               MsgBox "Tiene que ingresar OBSERVACIONES para cada  item de TOMOGRAFIA (mínimo 50 caracteres)", vbInformation, ""
               Exit Function
            End If
            oRsTomografia.MoveNext
         Loop
    End If
  End If
  '
  ValidaReglas = True
End Function


Public Function DevuelveOtrosCpt() As Recordset
    Set DevuelveOtrosCpt = UserControl.ucRecetaCpt1.DevuelveOtrosCpt
End Function
'SCCQ 05/11/2020 Cambio40 Inicio
Private Sub UserControl_Initialize()
btnAddRayosX.Enabled = False
btnAddEcoO.Enabled = False
btnAddEcoG.Enabled = False
btnAddTomografia.Enabled = False
btnAddPatologia.Enabled = False
btnAddAnatomia.Enabled = False
btnAddBanco.Enabled = False
End Sub
'SCCQ 05/11/2020 Cambio40 Fin
