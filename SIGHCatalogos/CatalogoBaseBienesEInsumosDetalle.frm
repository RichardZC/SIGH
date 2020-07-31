VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form CatalogoBaseBienesEInsumosDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13755
   Icon            =   "CatalogoBaseBienesEInsumosDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   13755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   600
      Left            =   60
      TabIndex        =   45
      Top             =   6015
      Width           =   13650
      Begin VB.TextBox txtCodigoSunat 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1785
         MaxLength       =   20
         TabIndex        =   47
         Top             =   150
         Width           =   1065
      End
      Begin VB.CommandButton cmdBuscaCodigoSunat 
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
         Left            =   2880
         Picture         =   "CatalogoBaseBienesEInsumosDetalle.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Buscar CODIGO SUNAT"
         Top             =   135
         Width           =   300
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Código SUNAT"
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
         Height          =   255
         Left            =   150
         TabIndex        =   48
         Top             =   150
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4785
      Left            =   8220
      TabIndex        =   29
      Top             =   60
      Width           =   5475
      Begin VB.CommandButton cmdActualizaPreciosXcolumna 
         Caption         =   "Iguala Precios en los demás Producto/Plan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   210
         TabIndex        =   31
         Top             =   4170
         Width           =   5115
      End
      Begin UltraGrid.SSUltraGrid grdPrecios 
         Height          =   3780
         Left            =   90
         TabIndex        =   30
         Top             =   210
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   6668
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Precios"
      End
   End
   Begin VB.Frame fraPresupuesto 
      Caption         =   "Presupuesto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   8220
      TabIndex        =   21
      Top             =   4890
      Width           =   5490
      Begin VB.ComboBox cmbIdCentroCosto 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   270
         Width           =   3765
      End
      Begin VB.ComboBox cmbIdPartida 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   630
         Width           =   3765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
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
         Left            =   150
         TabIndex        =   25
         Top             =   300
         Width           =   1350
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Partida"
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
         Left            =   150
         TabIndex        =   24
         Top             =   660
         Width           =   555
      End
   End
   Begin VB.Frame fraGrupoFarmacologico 
      Caption         =   "Grupo Farmacológico"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   60
      TabIndex        =   20
      Top             =   4890
      Width           =   8100
      Begin VB.ComboBox cmbIdGrupoFarmacologico 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   270
         Width           =   6135
      End
      Begin VB.ComboBox cmbIdSubGrupoFarmacologico 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   630
         Width           =   6135
      End
      Begin VB.Label Label2 
         Caption         =   "Grupo"
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
         Left            =   150
         TabIndex        =   23
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Sub Grupo"
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
         Left            =   150
         TabIndex        =   22
         Top             =   660
         Width           =   975
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
      Height          =   4830
      Left            =   60
      TabIndex        =   19
      Top             =   30
      Width           =   8100
      Begin VB.ComboBox cmbTpSISMED 
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
         ItemData        =   "CatalogoBaseBienesEInsumosDetalle.frx":1254
         Left            =   6300
         List            =   "CatalogoBaseBienesEInsumosDetalle.frx":125E
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   3840
         Width           =   1635
      End
      Begin VB.CommandButton cmdBuscaEnTablasSIS 
         Caption         =   "..."
         Height          =   315
         Left            =   3300
         TabIndex        =   41
         ToolTipText     =   "Busca en Tablas del SIS: Medicamentos e Insumos"
         Top             =   300
         Width           =   315
      End
      Begin VB.CheckBox chkPetitorio 
         Alignment       =   1  'Right Justify
         Caption         =   "Petitorio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   165
         TabIndex        =   40
         Top             =   3840
         Width           =   1890
      End
      Begin VB.ComboBox cmbPaisOrigen 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5850
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2400
         Width           =   2115
      End
      Begin VB.TextBox txtFabricante 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         MaxLength       =   100
         TabIndex        =   9
         Top             =   3480
         Width           =   6075
      End
      Begin VB.TextBox txtPresentE 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         MaxLength       =   100
         TabIndex        =   8
         Top             =   3120
         Width           =   6075
      End
      Begin VB.TextBox txtMaterialE 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         MaxLength       =   100
         TabIndex        =   7
         Top             =   2760
         Width           =   6075
      End
      Begin VB.TextBox txtFF 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtPresentacion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1680
         Width           =   6075
      End
      Begin VB.TextBox txtConcentracion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         MaxLength       =   100
         TabIndex        =   4
         Top             =   2040
         Width           =   6075
      End
      Begin VB.TextBox txtDenominacion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1350
         Width           =   6075
      End
      Begin VB.TextBox txtNombreComercial 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         MaxLength       =   100
         TabIndex        =   1
         Top             =   1020
         Width           =   6075
      End
      Begin VB.TextBox txtNombre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         MaxLength       =   250
         TabIndex        =   17
         Top             =   660
         Width           =   6075
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1860
         MaxLength       =   7
         TabIndex        =   0
         Top             =   300
         Width           =   1395
      End
      Begin VB.ComboBox cmbIdClasificacionBienInsumo___ 
         Height          =   315
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   330
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Producto SISMED"
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
         Left            =   4395
         TabIndex        =   44
         Top             =   3885
         Width           =   1860
      End
      Begin VB.Label lblHalladosEnSis 
         AutoSize        =   -1  'True
         Caption         =   "Código hallado en tablas SIS"
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
         Left            =   180
         TabIndex        =   42
         Top             =   4515
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "País Origen"
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
         Left            =   4890
         TabIndex        =   39
         Top             =   2460
         Width           =   900
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante"
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
         TabIndex        =   38
         Top             =   3510
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Present. Envase"
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
         TabIndex        =   37
         Top             =   3150
         Width           =   1320
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Material de Envase"
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
         TabIndex        =   36
         Top             =   2790
         Width           =   1515
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Forma Farmacéut (f)"
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
         TabIndex        =   35
         Top             =   2430
         Width           =   1680
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Presentación (p)"
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
         TabIndex        =   34
         Top             =   1710
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Concentración (c)"
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
         TabIndex        =   33
         Top             =   2070
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación (d)"
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
         TabIndex        =   32
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Comercial"
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
         TabIndex        =   28
         Top             =   1050
         Width           =   1470
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre (d+p+c+f)"
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
         TabIndex        =   27
         Top             =   720
         Width           =   1575
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
         TabIndex        =   26
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   30
      TabIndex        =   18
      Top             =   6570
      Width           =   13650
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CatalogoBaseBienesEInsumosDetalle.frx":1277
         DownPicture     =   "CatalogoBaseBienesEInsumosDetalle.frx":16D7
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
         Left            =   5422
         Picture         =   "CatalogoBaseBienesEInsumosDetalle.frx":1B4C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CatalogoBaseBienesEInsumosDetalle.frx":1FC1
         DownPicture     =   "CatalogoBaseBienesEInsumosDetalle.frx":2485
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
         Left            =   6967
         Picture         =   "CatalogoBaseBienesEInsumosDetalle.frx":2971
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CatalogoBaseBienesEInsumosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Bienes e Insumos de Farmacia
'        Programado por: Castro W
'        Fecha: Agosto 2005
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_CatalogoBienesInsumos As New DOCatalogoBienesInsumos
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdProducto As Long
Dim mo_AdminComun As New ReglasComunes
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_cmbIdClasificacionBienInsumo As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdCentroCosto As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdGrupoFarmacologico As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdPartida As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdSubGrupoFarmacologico As New SIGHEntidades.ListaDespleglable
Dim mo_cmbPaisOrigen As New SIGHEntidades.ListaDespleglable
Dim mo_cmbNacionalidad As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTpSISMED As New SIGHEntidades.ListaDespleglable
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim mrs_Precios As New Recordset
Dim lnPrecioNew As Double, lcColumnaEditada As String

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
Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

' Actualizado yamill palomino 10102014
Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
         
     Case sghModificar
         HabilitarDeshabilitarControles
         CargarDatosALosControles
     Case sghConsultar
         HabilitarDeshabilitarControles
         'fraDatosGenerales.Enabled = False
         'fraGrupoFarmacologico.Enabled = False
         'fraPresupuesto.Enabled = False
         CargarDatosALosControles
     Case sghEliminar
         HabilitarDeshabilitarControles
         'fraDatosGenerales.Enabled = False
         'fraGrupoFarmacologico.Enabled = False
         'fraPresupuesto.Enabled = False
         CargarDatosALosControles
 End Select
End Sub
Private Sub cmbIdCentroCosto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdCentroCosto
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdGrupoFarmacologico_Click()
    'Recuperamos los  SubGrupos
    mo_cmbIdSubGrupoFarmacologico.BoundColumn = "IdSubGrupoFarmacologico"
    mo_cmbIdSubGrupoFarmacologico.ListField = "Descripcion"
    Set mo_cmbIdSubGrupoFarmacologico.RowSource = mo_AdminComun.InsumosSubGrupoFarmacologicoSeleccionarPorGrupo(Val(mo_cmbIdGrupoFarmacologico.BoundText))
End Sub

Private Sub cmbIdGrupoFarmacologico_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdGrupoFarmacologico
    AdministrarKeyPreview KeyCode
End Sub
Private Sub cmbIdPartida_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdPartida
    AdministrarKeyPreview KeyCode
End Sub
Private Sub cmbIdSubGrupoFarmacologico_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdSubGrupoFarmacologico
    AdministrarKeyPreview KeyCode
End Sub







Private Sub cmbPaisOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbPaisOrigen
    AdministrarKeyPreview KeyCode

End Sub

Private Sub cmdActualizaPreciosXcolumna_Click()
            mrs_Precios.MoveFirst
            Do While Not mrs_Precios.EOF
               Select Case lcColumnaEditada
                Case "PrecioVenta"
                     mrs_Precios.Fields!PrecioVenta = lnPrecioNew
                Case "PrecioDistribucion"
                     mrs_Precios.Fields!PrecioDistribucion = lnPrecioNew
                Case "PrecioCompra"
                     mrs_Precios.Fields!PrecioCompra = lnPrecioNew
                Case "PrecioDonacion"
                     mrs_Precios.Fields!PrecioDonacion = lnPrecioNew
               End Select
               mrs_Precios.Update
               mrs_Precios.MoveNext
            Loop

End Sub

Private Sub cmdBuscaCodigoSunat_Click()
    Dim oBuscaCodigoSunat As New SIGHNegocios.BuscaCodigoSunat
    oBuscaCodigoSunat.MostrarFormulario
    If oBuscaCodigoSunat.BotonPresionado = sghAceptar Then
       txtCodigoSunat.Text = oBuscaCodigoSunat.codigoSUNAT
    End If
    Set oBuscaCodigoSunat = Nothing
End Sub

Private Sub cmdBuscaEnTablasSIS_Click()
    ActualizaDatosDesdeSIScodigo
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdCentroCosto.MiComboBox = cmbIdCentroCosto
    Set mo_cmbIdGrupoFarmacologico.MiComboBox = cmbIdGrupoFarmacologico
    Set mo_cmbIdPartida.MiComboBox = cmbIdPartida
    Set mo_cmbIdSubGrupoFarmacologico.MiComboBox = cmbIdSubGrupoFarmacologico
    Set mo_cmbPaisOrigen.MiComboBox = cmbPaisOrigen
    Set mo_cmbTpSISMED.MiComboBox = cmbTpSISMED
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       mo_Formulario.HabilitarDeshabilitar cmbTpSISMED, False
       CreaTemporal
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Bien e Insumo"
           mo_Formulario.HabilitarDeshabilitar cmbTpSISMED, True
       Case sghModificar
           Me.Caption = "Modificar Bien e Insumo"
       Case sghConsultar
           Me.Caption = "Consultar Bien e Insumo"
       Case sghEliminar
           Me.Caption = "Eliminar Bien e Insumo"
       End Select
       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
      ' mo_Formulario.HabilitarDeshabilitar txtNombre, False
       mo_Apariencia.ConfigurarFilasBiColores Me.grdPrecios, SIGHEntidades.GrillaConFilasBicolor
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
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                   LimpiarFormulario
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminComun.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminComun.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminComun.MensajeError, vbExclamation, Me.Caption
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
   
   If mo_cmbIdSubGrupoFarmacologico.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese sub grupo farmacologico" + Chr(13)
   End If
   If Trim(Me.txtCodigo) = "" Then
       sMensaje = sMensaje + "Ingrese el código" + Chr(13)
   End If
   If Trim(Me.txtNombre) = "" Then
       sMensaje = sMensaje + "Ingrese el nombre" + Chr(13)
   End If
   If Me.cmbIdPartida.Text = "" Then
       sMensaje = sMensaje + "Elija la PARTIDA" + Chr(13)
   End If
   If cmbTpSISMED.Text = "" And mi_Opcion = sghAgregar Then
      sMensaje = sMensaje + "Elija TIPO PRODUCTO SISMED" + Chr(13)
   End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   Dim sMensaje As String
   sMensaje = ""
   'Valida codigos Repetidos
   Dim oRsBuscaCodigo As New Recordset
   Set oRsBuscaCodigo = mo_AdminComun.CatalogoBienesInsumosSeleccionarPorCodigo(txtCodigo.Text)
   Select Case mi_Opcion
   Case sghAgregar
        If oRsBuscaCodigo.RecordCount > 0 Then
            sMensaje = sMensaje + "Ese código ya esta Registrado para: " + oRsBuscaCodigo.Fields!Nombre + Chr(13)
        End If
        If mo_AdminComun.EsUnaClinicaNOminsa = False Then
            If mi_Opcion = sghAgregar And lblHalladosEnSis.Visible = False Then
               sMensaje = sMensaje + "No se encontró en las tablas SIS, tendrá problemas con el formato FUA" + Chr(13)
            End If
        End If
   Case sghModificar
        If oRsBuscaCodigo.RecordCount > 0 Then
           oRsBuscaCodigo.MoveFirst
           Do While Not oRsBuscaCodigo.EOF
              If oRsBuscaCodigo.Fields!codigo = Me.txtCodigo.Text And oRsBuscaCodigo.Fields!IdProducto <> ml_IdProducto Then
                 sMensaje = sMensaje + "Ese código ya esta Registrado para: " + oRsBuscaCodigo.Fields!Nombre + Chr(13)
                 Exit Do
              End If
              oRsBuscaCodigo.MoveNext
           Loop
        End If
   Case sghEliminar
        Set oRsBuscaCodigo = mo_ReglasFarmacia.FarmMovimientoDetalleSeleccionarXcodigo(txtCodigo.Text)
        If oRsBuscaCodigo.RecordCount > 0 Then
           sMensaje = sMensaje = "Ya existe un MOVIMIENTO: " & oRsBuscaCodigo!movNumero
        End If
        oRsBuscaCodigo.Close
   End Select
   Set oRsBuscaCodigo = Nothing
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
   'Me.txtPrecioUnitario = Replace(Me.txtPrecioUnitario, ".", ",")
   With mo_CatalogoBienesInsumos
        .codigoSUNAT = txtCodigoSunat.Text
        .codigo = Me.txtCodigo.Text
        .Nombre = Me.txtNombre.Text
        '.PrecioUnitario = CCur(Me.txtPrecioUnitario.Text)
        .NombreComercial = Me.txtNombreComercial
        '.IdClasificacionBienInsumo = Val(mo_cmbIdClasificacionBienInsumo.BoundText)
        .IdGrupoFarmacologico = Val(mo_cmbIdGrupoFarmacologico.BoundText)
        .IdSubGrupoFarmacologico = Val(mo_cmbIdSubGrupoFarmacologico.BoundText)
        .IdPartida = Val(mo_cmbIdPartida.BoundText)
        .IdCentroCosto = Val(mo_cmbIdCentroCosto.BoundText)
        mrs_Precios.MoveFirst
        .PrecioCompra = mrs_Precios.Fields!PrecioCompra
        .PrecioDistribucion = mrs_Precios.Fields!PrecioDistribucion
        .PrecioDonacion = mrs_Precios.Fields!PrecioDonacion
        .IdUsuarioAuditoria = Me.idUsuario
        .Denominacion = txtDenominacion.Text
        .Concentracion = txtConcentracion.Text
        .Presentacion = txtPresentacion.Text
        .FormaFarmaceutica = txtFF.Text
        .IdPaisOrigen = Val(mo_cmbPaisOrigen.BoundText)
        .MaterialEnvase = txtMaterialE.Text
        .PresentacionEnvase = txtPresentE.Text
        .Fabricante = txtFabricante.Text
        .Petitorio = IIf(chkPetitorio.Value = 1, True, False)
        .TipoProductoSismed = Chr(mo_cmbTpSISMED.BoundText)
   End With
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   mo_CatalogoBienesInsumos.idTipoSalidaBienInsumo = 3  'ventas/programas
   AgregarDatos = mo_AdminComun.CatalogoBienesInsumosAgregar(mo_CatalogoBienesInsumos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtCodigo.Text) & " " & txtNombre.Text)
   If AgregarDatos = True Then
        mo_AdminComun.ActualizaPreciosParaFarmacia mrs_Precios, mo_CatalogoBienesInsumos.IdProducto
   End If
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   
   ModificarDatos = mo_AdminComun.CatalogoBienesInsumosModificar(mo_CatalogoBienesInsumos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtCodigo.Text) & " " & txtNombre.Text)
   If ModificarDatos = True Then
        mo_AdminComun.ActualizaPreciosParaFarmacia mrs_Precios, mo_CatalogoBienesInsumos.IdProducto
   End If

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminComun.CatalogoBienesInsumosEliminar(mo_CatalogoBienesInsumos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtCodigo.Text) & " " & txtNombre.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
    
    Set mo_CatalogoBienesInsumos = mo_AdminComun.CatalogoBienesInsumosSeleccionarPorId(Me.IdProducto)
    If mo_AdminComun.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos " + Chr(13) + mo_AdminComun.MensajeError, vbInformation, Me.Caption
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_CatalogoBienesInsumos Is Nothing Then
        With mo_CatalogoBienesInsumos
            txtCodigoSunat.Text = .codigoSUNAT
            mo_cmbIdCentroCosto.BoundText = .IdCentroCosto
            mo_cmbIdGrupoFarmacologico.BoundText = .IdGrupoFarmacologico
            mo_cmbIdPartida.BoundText = .IdPartida
            mo_cmbIdSubGrupoFarmacologico.BoundText = .IdSubGrupoFarmacologico
            
            Me.txtNombre = .Nombre
            Me.txtCodigo = .codigo
            Me.txtNombreComercial.Text = .NombreComercial
            txtDenominacion.Text = .Denominacion
            txtConcentracion.Text = .Concentracion
            txtPresentacion.Text = .Presentacion
            txtFF.Text = .FormaFarmaceutica
            mo_cmbPaisOrigen.BoundText = .IdPaisOrigen
            txtMaterialE.Text = .MaterialEnvase
            txtPresentE.Text = .PresentacionEnvase
            txtFabricante.Text = .Fabricante
            chkPetitorio.Value = IIf(.Petitorio = True, 1, 0)
            mo_cmbTpSISMED.BoundText = Asc(.TipoProductoSismed)
            mb_ExistenDatos = True
        End With
        'Carga  Precios
        Dim oFactCatalogoServiciosPtos As New Recordset
        Set oFactCatalogoServiciosPtos = mo_ReglasFacturacion.CatalogoBienesInsumosHospSeleccionarXIdProducto(Me.IdProducto)
        If oFactCatalogoServiciosPtos.RecordCount > 0 Then
           oFactCatalogoServiciosPtos.MoveFirst
           Do While Not oFactCatalogoServiciosPtos.EOF
              mrs_Precios.MoveFirst
              mrs_Precios.Find "idTipoFinanciamiento=" & oFactCatalogoServiciosPtos.Fields!idTipoFinanciamiento
              If Not mrs_Precios.EOF Then
                 mrs_Precios.Fields!PrecioVenta = oFactCatalogoServiciosPtos.Fields!PrecioUnitario
                 mrs_Precios.Fields!PrecioDistribucion = mo_CatalogoBienesInsumos.PrecioDistribucion
                 mrs_Precios.Fields!PrecioCompra = mo_CatalogoBienesInsumos.PrecioCompra
                 mrs_Precios.Fields!PrecioDonacion = mo_CatalogoBienesInsumos.PrecioDonacion
                 mrs_Precios.Update
              End If
              oFactCatalogoServiciosPtos.MoveNext
           Loop
           mrs_Precios.MoveFirst
        End If
        oFactCatalogoServiciosPtos.Close
        Set oFactCatalogoServiciosPtos = Nothing
        'Busca si tiene Historicos
        If mi_Opcion = sghEliminar Then
           Dim oRsTmp1 As New Recordset
           Set oRsTmp1 = mo_ReglasFarmacia.FarmMovimientoDetalleSeleccionarXcodigo(Me.txtCodigo.Text)
           If oRsTmp1.RecordCount > 0 Then
              MsgBox "Existen Movimientos Historicos", vbInformation, Me.Caption
              Me.btnAceptar.Enabled = False
           End If
           oRsTmp1.Close
           Set oRsTmp1 = Nothing
        End If
        'debb-08/11/2016
        If mo_ReglasFarmacia.CatalogoDIGEMIDesCodigoPaquete(txtCodigo.Text) = True And mi_Opcion = sghModificar Then
           MsgBox "No podrá MODIFICAR PRECIOS, porque es un CODIGO DE PAQUETE", vbInformation, Me.Caption
           Frame1.Enabled = False
        End If
        '
    Else
        mb_ExistenDatos = False
        Exit Sub
    End If
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

    Me.IdProducto = 0
    mo_cmbIdCentroCosto.BoundText = ""
    mo_cmbIdGrupoFarmacologico.BoundText = ""
    mo_cmbIdPartida.BoundText = ""
    mo_cmbIdSubGrupoFarmacologico.BoundText = ""
    mo_cmbTpSISMED.BoundText = ""
    
    Me.txtNombre = ""
    Me.txtCodigo = ""
    Me.txtNombreComercial = ""
    txtDenominacion.Text = ""
    txtConcentracion.Text = ""
    txtPresentacion.Text = ""
    txtFF.Text = ""
    mo_cmbPaisOrigen.BoundText = ""
    txtMaterialE.Text = ""
    txtPresentE.Text = ""
    txtFabricante.Text = ""
End Sub

Sub CargarComboBoxes()
       
    mo_cmbIdCentroCosto.BoundColumn = "IdCentroCosto"
    mo_cmbIdCentroCosto.ListField = "Descripcion"
    Set mo_cmbIdCentroCosto.RowSource = mo_AdminComun.CentrosCostoSeleccionarTodos

    mo_cmbIdPartida.BoundColumn = "IdPartidaPresupuestal"
    mo_cmbIdPartida.ListField = "Descripcion"
    Set mo_cmbIdPartida.RowSource = mo_AdminComun.PartidasPresupuestalesSeleccionarTodos

    mo_cmbIdGrupoFarmacologico.BoundColumn = "IdGrupoFarmacologico"
    mo_cmbIdGrupoFarmacologico.ListField = "Descripcion"
    Set mo_cmbIdGrupoFarmacologico.RowSource = mo_AdminComun.InsumosGrupoFarmacologicoSeleccionarTodos
    
    mo_cmbPaisOrigen.BoundColumn = "IdPais"
    mo_cmbPaisOrigen.ListField = "Nombre"
    Set mo_cmbPaisOrigen.RowSource = mo_AdminServiciosGeograficos.PaisesSeleccionarTodos()
    
    mo_cmbTpSISMED.BoundColumn = "identificador"
    mo_cmbTpSISMED.ListField = "Descripcion"
    Set mo_cmbTpSISMED.RowSource = mo_ReglasFarmacia.farmTipoProductosSismedDevuelveTodos
    
End Sub


Private Sub grdPrecios_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
        On Error Resume Next
        Dim oRow As SSRow
        Set oRow = grdPrecios.ActiveCell.Row
        lcColumnaEditada = grdPrecios.ActiveCell.Column.Key
        Select Case lcColumnaEditada
        Case "PrecioVenta"
            lnPrecioNew = mrs_Precios.Fields!PrecioVenta
        Case "PrecioDistribucion"
            lnPrecioNew = mrs_Precios.Fields!PrecioDistribucion
        Case "PrecioCompra"
            lnPrecioNew = mrs_Precios.Fields!PrecioCompra
        Case "PrecioDonacion"
            lnPrecioNew = mrs_Precios.Fields!PrecioDonacion
        End Select
End Sub

Private Sub grdPrecios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPrecios.Bands(0).Columns("IdTipoFinanciamiento").Hidden = True
    'Actualizado Yamill Palomino 16/10/2014
    If mi_Opcion = sghConsultar Or mi_Opcion = sghEliminar Then
        grdPrecios.Bands(0).Columns("TipoFinanciamiento").Activation = ssActivationActivateNoEdit
        grdPrecios.Bands(0).Columns("PrecioVenta").Activation = ssActivationActivateNoEdit
        grdPrecios.Bands(0).Columns("PrecioDistribucion").Activation = ssActivationActivateNoEdit
        grdPrecios.Bands(0).Columns("PrecioCompra").Activation = ssActivationActivateNoEdit
        grdPrecios.Bands(0).Columns("PrecioDonacion").Activation = ssActivationActivateNoEdit
    End If
    '
    grdPrecios.Bands(0).Columns("TipoFinanciamiento").Header.Caption = "Producto/Plan"
    grdPrecios.Bands(0).Columns("TipoFinanciamiento").Width = 1500
    '
    grdPrecios.Bands(0).Columns("PrecioVenta").Header.Caption = "Pr.Venta"
    grdPrecios.Bands(0).Columns("PrecioVenta").Width = 800
    '
    grdPrecios.Bands(0).Columns("PrecioDistribucion").Header.Caption = "Pr.Distribución"
    grdPrecios.Bands(0).Columns("PrecioDistribucion").Width = 800
    '
    grdPrecios.Bands(0).Columns("PrecioCompra").Header.Caption = "Pr.Compra"
    grdPrecios.Bands(0).Columns("PrecioCompra").Width = 800
    '
    grdPrecios.Bands(0).Columns("PrecioDonacion").Header.Caption = "Pr.Donación"
    grdPrecios.Bands(0).Columns("PrecioDonacion").Width = 800

End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode
End Sub

' Actualizado yamill palomino 10102014
Private Sub txtCodigo_LostFocus()
     If mi_Opcion = sghAgregar Then
        ActualizaDatosDesdeSIScodigo
    End If
End Sub
Sub ActualizaDatosDesdeSIScodigo()
    lblHalladosEnSis.Visible = False
    If Len(txtCodigo.Text) > 0 Then
        Dim oBuscaCodigoEnSIS As New SIGHNegocios.SisConsumoWeb
        Dim rsTmp As New ADODB.Recordset
        If oBuscaCodigoEnSIS.ConsultarServicioMedicamentosxCodigo(txtCodigo.Text, rsTmp) = True Then
           If rsTmp.RecordCount = 0 Then
              Set rsTmp = oBuscaCodigoEnSIS.m_medicamentosSeleccionarPorId(txtCodigo.Text)
              If rsTmp.RecordCount > 0 Then
                lblHalladosEnSis.Visible = True
                lblHalladosEnSis.Caption = "CODIGO hallado en WEB SIS"
                ActualizaDatosUbicadoEnSis rsTmp, True
              End If
           Else
             lblHalladosEnSis.Visible = True
             lblHalladosEnSis.Caption = "CODIGO hallado en TABLAS SIS LOCAL"
             ActualizaDatosUbicadoEnSis rsTmp, True
           End If
        ElseIf oBuscaCodigoEnSIS.ConsultarServicioInsumosxCodigo(txtCodigo.Text, rsTmp) = True Then
           If rsTmp.RecordCount = 0 Then
              Set rsTmp = oBuscaCodigoEnSIS.m_insumosSeleccionarPorId(txtCodigo.Text)
              If rsTmp.RecordCount > 0 Then
                lblHalladosEnSis.Visible = True
                lblHalladosEnSis.Caption = "CODIGO hallado en WEB SIS"
                ActualizaDatosUbicadoEnSis rsTmp, False
              End If
           Else
             lblHalladosEnSis.Visible = True
             lblHalladosEnSis.Caption = "CODIGO hallado en TABLAS SIS LOCAL"
             ActualizaDatosUbicadoEnSis rsTmp, False
           End If
        End If
        Set rsTmp = Nothing
        Set oBuscaCodigoEnSIS = Nothing
    End If
End Sub
Sub ActualizaDatosUbicadoEnSis(rsTmp As Recordset, lbDesdeMedicamento As Boolean)
    If lbDesdeMedicamento = True Then
       txtDenominacion.Text = IIf(IsNull(rsTmp!med_nombre), "", rsTmp!med_nombre)
       txtPresentacion.Text = IIf(IsNull(rsTmp!med_presen), "", rsTmp!med_presen)
       txtConcentracion.Text = IIf(IsNull(rsTmp!med_concen), "", rsTmp!med_concen)
       txtFF.Text = IIf(IsNull(rsTmp!med_formaFarmaceutica), "", rsTmp!med_formaFarmaceutica)
       chkPetitorio.Value = IIf(IsNull(rsTmp!med_petitorio), 0, IIf(rsTmp!med_petitorio = "S", 1, 0))
    Else
       txtDenominacion.Text = IIf(IsNull(rsTmp!ins_nombre), "", rsTmp!ins_nombre)
       txtPresentacion.Text = IIf(IsNull(rsTmp!ins_presen), "", rsTmp!ins_presen)
       txtConcentracion.Text = IIf(IsNull(rsTmp!ins_concen), "", rsTmp!ins_concen)
       txtFF.Text = IIf(IsNull(rsTmp!ins_formaFarmaceutica), "", rsTmp!ins_formaFarmaceutica)
       chkPetitorio.Value = IIf(IsNull(rsTmp!ins_petitorio), 0, IIf(rsTmp!ins_petitorio = "S", 1, 0))
    End If
    UneDenominacionConcentracionPresentacion
End Sub



Private Sub txtConcentracion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtConcentracion
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtConcentracion_LostFocus()
    UneDenominacionConcentracionPresentacion
End Sub

Private Sub txtDenominacion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDenominacion
    AdministrarKeyPreview KeyCode

End Sub





Private Sub txtDenominacion_LostFocus()
   UneDenominacionConcentracionPresentacion
End Sub

Private Sub txtFabricante_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFabricante
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFF_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFF
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtMaterialE_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtMaterialE
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNombre
    AdministrarKeyPreview KeyCode

End Sub
Private Sub txtNombreComercial_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNombreComercial
    AdministrarKeyPreview KeyCode

End Sub

Sub CreaTemporal()
    If mrs_Precios.State = adStateOpen Then mrs_Precios.Close
    With mrs_Precios
          .Fields.Append "IdTipoFinanciamiento", adInteger, 4, adFldIsNullable
          .Fields.Append "TipoFinanciamiento", adVarChar, 50, adFldIsNullable
          .Fields.Append "PrecioVenta", adDouble
          .Fields.Append "PrecioDistribucion", adDouble
          .Fields.Append "PrecioCompra", adDouble
          .Fields.Append "PrecioDonacion", adDouble
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Dim oRsTmp As New Recordset
    Set oRsTmp = mo_AdminComun.TiposFinanciamientoSegunFiltro("seIngresPrecios=1 and idTipoFinanciamiento>0")
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          mrs_Precios.AddNew
          mrs_Precios.Fields!idTipoFinanciamiento = oRsTmp.Fields!idTipoFinanciamiento
          mrs_Precios.Fields!TipoFinanciamiento = oRsTmp.Fields!Descripcion
          mrs_Precios.Fields!PrecioVenta = 0
          mrs_Precios.Fields!PrecioCompra = 0
          mrs_Precios.Fields!PrecioDistribucion = 0
          mrs_Precios.Fields!PrecioDonacion = 0
          mrs_Precios.Update
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set Me.grdPrecios.DataSource = mrs_Precios
    mrs_Precios.MoveFirst
End Sub



Private Sub txtPresentacion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPresentacion
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtPresentacion_LostFocus()
    UneDenominacionConcentracionPresentacion
End Sub

Private Sub txtPresentE_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPresentE
    AdministrarKeyPreview KeyCode

End Sub


Sub UneDenominacionConcentracionPresentacion()
    If txtDenominacion.Text <> "" Or txtPresentacion.Text <> "" Or txtConcentracion.Text <> "" Or txtFF.Text <> "" Then
       txtNombre.Text = Left(Trim(txtDenominacion.Text) & " " & Trim(txtPresentacion.Text) & " " & Trim(txtConcentracion.Text) & " " & txtFF.Text, 300)
    End If
End Sub

' Actualizado yamill palomino 10102014
Private Sub HabilitarDeshabilitarControles()

    If mi_Opcion = sghConsultar Or mi_Opcion = sghEliminar Then
        mo_Formulario.HabilitarDeshabilitar txtCodigo, False
        mo_Formulario.HabilitarDeshabilitar txtNombre, False
        mo_Formulario.HabilitarDeshabilitar txtNombreComercial, False
        mo_Formulario.HabilitarDeshabilitar txtDenominacion, False
        mo_Formulario.HabilitarDeshabilitar txtPresentacion, False
        mo_Formulario.HabilitarDeshabilitar txtConcentracion, False
        mo_Formulario.HabilitarDeshabilitar txtFF, False
        mo_Formulario.HabilitarDeshabilitar cmbPaisOrigen, False
        mo_Formulario.HabilitarDeshabilitar txtMaterialE, False
        mo_Formulario.HabilitarDeshabilitar txtPresentE, False
        mo_Formulario.HabilitarDeshabilitar txtFabricante, False
        mo_Formulario.HabilitarDeshabilitar chkPetitorio, False
        mo_Formulario.HabilitarDeshabilitar cmbTpSISMED, False
        mo_Formulario.HabilitarDeshabilitar cmbIdGrupoFarmacologico, False
        mo_Formulario.HabilitarDeshabilitar cmbIdSubGrupoFarmacologico, False
        mo_Formulario.HabilitarDeshabilitar cmbIdCentroCosto, False
        mo_Formulario.HabilitarDeshabilitar cmbIdPartida, False
    ElseIf mo_AdminComun.EsUnaClinicaNOminsa = False Then
        mo_Formulario.HabilitarDeshabilitar txtCodigo, False
        mo_Formulario.HabilitarDeshabilitar txtNombre, False
        mo_Formulario.HabilitarDeshabilitar txtNombreComercial, False
        mo_Formulario.HabilitarDeshabilitar txtDenominacion, False
        mo_Formulario.HabilitarDeshabilitar txtPresentacion, False
        mo_Formulario.HabilitarDeshabilitar txtConcentracion, False
        mo_Formulario.HabilitarDeshabilitar txtFF, False
        mo_Formulario.HabilitarDeshabilitar cmbPaisOrigen, False
        mo_Formulario.HabilitarDeshabilitar txtMaterialE, False
        mo_Formulario.HabilitarDeshabilitar txtPresentE, False
        mo_Formulario.HabilitarDeshabilitar txtFabricante, False
        mo_Formulario.HabilitarDeshabilitar chkPetitorio, False
        mo_Formulario.HabilitarDeshabilitar cmbTpSISMED, False
    End If

End Sub

