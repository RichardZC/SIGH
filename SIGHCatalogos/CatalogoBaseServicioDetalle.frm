VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form CatalogoBaseServicioDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15120
   Icon            =   "CatalogoBaseServicioDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraLabImag 
      Height          =   540
      Left            =   30
      TabIndex        =   41
      Top             =   1530
      Width           =   15060
      Begin VB.CommandButton cmdNuevoEquipo 
         DisabledPicture =   "CatalogoBaseServicioDetalle.frx":0CCA
         DownPicture     =   "CatalogoBaseServicioDetalle.frx":10B3
         Height          =   300
         Left            =   10530
         Picture         =   "CatalogoBaseServicioDetalle.frx":14BF
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Agrega EQUIPO de Imágenes"
         Top             =   165
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.ComboBox cmbEquipoImg 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "CatalogoBaseServicioDetalle.frx":18CB
         Left            =   6360
         List            =   "CatalogoBaseServicioDetalle.frx":18D5
         Style           =   2  'Dropdown List
         TabIndex        =   45
         ToolTipText     =   "Equipo de Imágenes"
         Top             =   165
         Visible         =   0   'False
         Width           =   4200
      End
      Begin VB.TextBox txtRuc 
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
         Left            =   1230
         MaxLength       =   11
         TabIndex        =   5
         ToolTipText     =   "RUC del PROVEEDOR DE IMAGENES/LABORATORIO"
         Top             =   165
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtRazonSocial 
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
         Left            =   2340
         MaxLength       =   100
         TabIndex        =   6
         ToolTipText     =   "RAZON SOCIAL del Proveedor de IMAGENES/LABORATORIO"
         Top             =   165
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.CheckBox chkResultadoAutomatico 
         Alignment       =   1  'Right Justify
         Caption         =   "Resultado automático en Laboratorio"
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
         Left            =   11625
         TabIndex        =   42
         Top             =   195
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Eq"
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
         Left            =   6120
         TabIndex        =   44
         Top             =   225
         Width           =   210
      End
      Begin VB.Label lblRuc 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ruc Proveed"
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
         Left            =   165
         TabIndex        =   43
         Top             =   210
         Visible         =   0   'False
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   6015
      TabIndex        =   32
      Top             =   2085
      Width           =   4530
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   120
         TabIndex        =   33
         Top             =   180
         Width           =   4305
         Begin VB.ComboBox cmbIdPtoCarga 
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
            ItemData        =   "CatalogoBaseServicioDetalle.frx":18F4
            Left            =   1290
            List            =   "CatalogoBaseServicioDetalle.frx":18F6
            TabIndex        =   14
            Text            =   "cmbIdPtoCarga"
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton btnAgregar 
            DisabledPicture =   "CatalogoBaseServicioDetalle.frx":18F8
            DownPicture     =   "CatalogoBaseServicioDetalle.frx":1CE1
            Height          =   315
            Left            =   1290
            Picture         =   "CatalogoBaseServicioDetalle.frx":20ED
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   645
            Width           =   795
         End
         Begin VB.CommandButton btnQuitar 
            DisabledPicture =   "CatalogoBaseServicioDetalle.frx":24F9
            DownPicture     =   "CatalogoBaseServicioDetalle.frx":2884
            Height          =   315
            Left            =   2370
            Picture         =   "CatalogoBaseServicioDetalle.frx":2C17
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   645
            Width           =   795
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Pto de Carga"
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
            TabIndex        =   34
            Top             =   300
            Width           =   1065
         End
      End
      Begin UltraGrid.SSUltraGrid grdPuntosDeCarga 
         Height          =   3360
         Left            =   120
         TabIndex        =   17
         Top             =   1350
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   5927
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Puntos de Carga"
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   30
      TabIndex        =   29
      Top             =   6960
      Width           =   15075
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CatalogoBaseServicioDetalle.frx":2FA8
         DownPicture     =   "CatalogoBaseServicioDetalle.frx":346C
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
         Left            =   7312
         Picture         =   "CatalogoBaseServicioDetalle.frx":3958
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CatalogoBaseServicioDetalle.frx":3E44
         DownPicture     =   "CatalogoBaseServicioDetalle.frx":42A4
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
         Left            =   5767
         Picture         =   "CatalogoBaseServicioDetalle.frx":4719
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   225
         Width           =   1365
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
      Left            =   30
      TabIndex        =   26
      Top             =   30
      Width           =   15060
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
         Left            =   5700
         Picture         =   "CatalogoBaseServicioDetalle.frx":4B8E
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Buscar CODIGO SUNAT"
         Top             =   1110
         Width           =   300
      End
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
         Left            =   4605
         MaxLength       =   20
         TabIndex        =   47
         Top             =   1110
         Width           =   1065
      End
      Begin VB.TextBox txtCodigoSIS 
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
         Left            =   7770
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1080
         Width           =   1035
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
         Height          =   255
         Left            =   150
         TabIndex        =   38
         Top             =   1110
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.TextBox txtNombreMINSA 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7770
         MaxLength       =   250
         TabIndex        =   3
         Top             =   690
         Width           =   7245
      End
      Begin VB.ComboBox cmbTipoServicio 
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
         ItemData        =   "CatalogoBaseServicioDetalle.frx":5118
         Left            =   1260
         List            =   "CatalogoBaseServicioDetalle.frx":5122
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   4755
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   0
         Top             =   330
         Width           =   1035
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
         Height          =   345
         Left            =   7770
         MaxLength       =   250
         TabIndex        =   2
         Top             =   300
         Width           =   7245
      End
      Begin VB.Label Label15 
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
         Left            =   3375
         TabIndex        =   48
         Top             =   1155
         Width           =   1200
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Se usa en el FUA"
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
         Left            =   8850
         TabIndex        =   40
         Top             =   1140
         Width           =   1395
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Código SIS"
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
         Left            =   6840
         TabIndex        =   39
         Top             =   1110
         Width           =   885
      End
      Begin VB.Label lblId 
         Alignment       =   1  'Right Justify
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4110
         TabIndex        =   37
         Top             =   360
         Width           =   1875
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre MINSA"
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
         Left            =   6480
         TabIndex        =   36
         Top             =   750
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   35
         Top             =   750
         Width           =   360
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
         Left            =   150
         TabIndex        =   28
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre corto"
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
         Left            =   6585
         TabIndex        =   27
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.Frame fraGrupoFarmacologico 
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
      Height          =   1905
      Left            =   30
      TabIndex        =   23
      Top             =   2085
      Width           =   5985
      Begin VB.ComboBox cmbIdServicioSubGrupo 
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
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   615
         Width           =   4680
      End
      Begin VB.ComboBox cmbIdServicioGrupo 
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
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   255
         Width           =   4695
      End
      Begin VB.ComboBox cmbIdServicioSubSeccion 
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
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1410
         Width           =   4695
      End
      Begin VB.ComboBox cmbIdServicioSeccion 
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
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1020
         Width           =   4695
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   31
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   30
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Sub Sección"
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
         Left            =   120
         TabIndex        =   25
         Top             =   1410
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Sección"
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
         Left            =   120
         TabIndex        =   24
         Top             =   1050
         Width           =   1095
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
      Height          =   2895
      Left            =   30
      TabIndex        =   20
      Top             =   4035
      Width           =   5985
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
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   4710
      End
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
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   4710
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
         Left            =   120
         TabIndex        =   22
         Top             =   660
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Centro Costo"
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
         TabIndex        =   21
         Top             =   270
         Width           =   1080
      End
   End
   Begin UltraGrid.SSUltraGrid grdPrecios 
      Height          =   4740
      Left            =   10545
      TabIndex        =   18
      Top             =   2175
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   8361
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
Attribute VB_Name = "CatalogoBaseServicioDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Procedimientos CPT MINSA
'        Programado por: Castro W
'        Fecha: Agosto 2005
'------------------------------------------------------------------------------------

Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_CatalogoServicios As New DOCatalogoServicio
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdProducto As Long
Dim mo_AdminComun As New ReglasComunes
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_cmbIdCentroCosto As New sighentidades.ListaDespleglable
Dim mo_cmbIdPartida As New sighentidades.ListaDespleglable
Dim mo_cmbIdPuntoCarga As New sighentidades.ListaDespleglable
Dim mo_cmbIdServicioGrupo As New sighentidades.ListaDespleglable
Dim mo_cmbIdServicioSubGrupo As New sighentidades.ListaDespleglable
Dim mo_cmbIdServicioSeccion As New sighentidades.ListaDespleglable
Dim mo_cmbIdServicioSubSeccion As New sighentidades.ListaDespleglable
Dim mo_cmbEquipoImg As New sighentidades.ListaDespleglable
Dim mrs_PuntosCarga As New Recordset
Dim mrs_Precios As New Recordset
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
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

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
         
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         fraDatosGenerales.Enabled = False
         fraGrupoFarmacologico.Enabled = False
         fraPresupuesto.Enabled = False
         CargarDatosALosControles
     Case sghEliminar
         fraDatosGenerales.Enabled = False
         fraGrupoFarmacologico.Enabled = False
         fraPresupuesto.Enabled = False
         CargarDatosALosControles
 End Select
End Sub

Private Sub btnAgregar_Click()
    Dim lbNuevo As Boolean
    Dim lcIdServicio As String
    If cmbIdPtoCarga.Text <> "" Then
        lbNuevo = True
        If mrs_PuntosCarga.RecordCount > 0 Then
            mrs_PuntosCarga.MoveFirst
            mrs_PuntosCarga.Find "idPuntoCarga=" & mo_cmbIdPuntoCarga.BoundText
            If Not mrs_PuntosCarga.EOF Then
               MsgBox "Ya existe en la Lista", vbInformation, Me.Caption
               lbNuevo = False
            End If
        End If
        If lbNuevo = True Then
            lcIdServicio = devuelveIdServicio(Val(mo_cmbIdPuntoCarga.BoundText))
            mrs_PuntosCarga.AddNew
            mrs_PuntosCarga.Fields!idPuntoCarga = Val(mo_cmbIdPuntoCarga.BoundText)
            mrs_PuntosCarga.Fields!Descripcion = Trim(cmbIdPtoCarga.Text) & lcIdServicio
            mrs_PuntosCarga.Fields!TieneIdServicio = IIf(lcIdServicio = "", False, True)
            mrs_PuntosCarga.Update
        End If
    End If
End Sub

Function devuelveIdServicio(lnIdPuntoCarga As Long) As String
            Dim oRsTmp As New Recordset
            Dim lcIdServicio As String
            
            Set oRsTmp = mo_AdminComun.FactPuntosCargaSeleccionarPorId(lnIdPuntoCarga)
            lcIdServicio = ""
            If oRsTmp.RecordCount > 0 Then
               oRsTmp.MoveFirst
               Do While Not oRsTmp.EOF
                  If oRsTmp.Fields!idServicio > 0 Then
                     lcIdServicio = "(Serv=" & Trim(Str(oRsTmp.Fields!idServicio)) & ")"
                  End If
                  oRsTmp.MoveNext
               Loop
            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
            devuelveIdServicio = lcIdServicio
End Function

Private Sub btnQuitar_Click()
    On Error Resume Next
    If Not mrs_PuntosCarga.EOF Then
       mrs_PuntosCarga.Delete
       mrs_PuntosCarga.Update
    End If
End Sub

Private Sub cmbIdCentroCosto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdCentroCosto
    AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdServicioGrupo_Click()
    
    mo_cmbIdServicioSubGrupo.BoundColumn = "IdServicioSubGrupo"
    mo_cmbIdServicioSubGrupo.ListField = "Descripcion"
    Set mo_cmbIdServicioSubGrupo.RowSource = mo_AdminComun.CatalogoServiciosSubGrupoSeleccionarPorGrupo(Val(mo_cmbIdServicioGrupo.BoundText))
    
End Sub

Private Sub cmbIdServicioSeccion_Click()
    'Recuperamos los  SubGrupos
    mo_cmbIdServicioSubSeccion.BoundColumn = "IdServicioSubSeccion"
    mo_cmbIdServicioSubSeccion.ListField = "Descripcion"
    Set mo_cmbIdServicioSubSeccion.RowSource = mo_AdminComun.CatalogoServiciosSubSeccionSeleccionarPorSeccion(Val(mo_cmbIdServicioSeccion.BoundText))
End Sub
Private Sub cmbIdServicioGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdServicioGrupo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdServicioSeccion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdServicioSeccion
    AdministrarKeyPreview KeyCode
End Sub
Private Sub cmbIdPartida_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdPartida
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdServicioSubGrupo_Click()
    
    mo_cmbIdServicioSeccion.BoundColumn = "IdServicioSeccion"
    mo_cmbIdServicioSeccion.ListField = "Descripcion"
    Set mo_cmbIdServicioSeccion.RowSource = mo_AdminComun.CatalogoServiciosSeccionSeleccionarPorSubGrupo(Val(mo_cmbIdServicioSubGrupo.BoundText))
    MuestraLblAutomatico
End Sub

Sub MuestraLblAutomatico()
    Label14.Visible = False
    lblRuc.Visible = False
    txtRuc.Visible = False
    txtRazonSocial.Visible = False
    chkResultadoAutomatico.Visible = False
    cmbEquipoImg.Visible = False
    cmdNuevoEquipo.Visible = False
    Select Case Val(mo_cmbIdServicioSubGrupo.BoundText)
    Case 2, 3
        chkResultadoAutomatico.Visible = True
        If Val(mo_cmbIdServicioSubGrupo.BoundText) = 2 Then
           chkResultadoAutomatico.Caption = "Resultado automático en Imágenes"
           cmbEquipoImg.Visible = True
           cmdNuevoEquipo.Visible = True
           Label14.Visible = True
        Else
           chkResultadoAutomatico.Caption = "Resultado automático en Laboratorio"
        End If
        lblRuc.Visible = True
        txtRuc.Visible = True
        txtRazonSocial.Visible = True
    Case Else
        chkResultadoAutomatico.Value = 0
    End Select
End Sub

Private Sub cmbIdServicioSubSeccion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdServicioSubSeccion
    AdministrarKeyPreview KeyCode
End Sub
Private Sub cmbIdServicioSubGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdServicioSubGrupo
    AdministrarKeyPreview KeyCode
End Sub




Private Sub cmbTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdServicioGrupo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdBuscaCodigoSunat_Click()
    Dim oBuscaCodigoSunat As New SIGHNegocios.BuscaCodigoSunat
    oBuscaCodigoSunat.MostrarFormulario
    If oBuscaCodigoSunat.BotonPresionado = sghAceptar Then
       txtCodigoSunat.Text = oBuscaCodigoSunat.codigoSUNAT
    End If
    Set oBuscaCodigoSunat = Nothing
End Sub

Private Sub cmdNuevoEquipo_Click()
    Dim oImagEquipos As New ImagEquipos
    oImagEquipos.Show 1
    Set ImagEquipos = Nothing
    CArgaEquiposIMG
    If mi_Opcion <> sghAgregar Then
       mo_cmbEquipoImg.BoundText = mo_CatalogoServicios.EquipoCodigo
    End If
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdCentroCosto.MiComboBox = cmbIdCentroCosto
    Set mo_cmbIdServicioGrupo.MiComboBox = cmbIdServicioGrupo
    Set mo_cmbIdServicioSeccion.MiComboBox = cmbIdServicioSeccion
    Set mo_cmbIdPartida.MiComboBox = cmbIdPartida
    Set mo_cmbIdServicioSubSeccion.MiComboBox = cmbIdServicioSubSeccion
    Set mo_cmbIdServicioSubGrupo.MiComboBox = cmbIdServicioSubGrupo
    Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPtoCarga
    Set mo_cmbEquipoImg.MiComboBox = cmbEquipoImg
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       CreaTemporal
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Servicio"
       Case sghModificar
           Me.Caption = "Modificar Servicio"
       Case sghConsultar
           Me.Caption = "Consultar Servicio"
       Case sghEliminar
           Me.Caption = "Eliminar Servicio"
       End Select
       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
       mo_Apariencia.ConfigurarFilasBiColores grdPrecios, sighentidades.GrillaConFilasBicolor
       mo_Apariencia.ConfigurarFilasBiColores grdPuntosDeCarga, sighentidades.GrillaConFilasBicolor
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
                   lblId.Caption = mo_CatalogoServicios.IdProducto
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
   
   If mo_cmbIdServicioSubGrupo.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese la categoria del producto" + Chr(13)
   End If
   If Trim(Me.txtCodigo) = "" Then
       sMensaje = sMensaje + "Ingrese el código" + Chr(13)
   End If
   If Trim(Me.txtNombre) = "" Then
       sMensaje = sMensaje + "Ingrese el nombre (corto)" + Chr(13)
   End If
   If Trim(Me.txtNombreMINSA.Text) = "" Then
       sMensaje = sMensaje + "Ingrese el nombre (MINSA)" + Chr(13)
   End If
   If cmbTipoServicio.Text = "" Then
       sMensaje = sMensaje + "Elija el Tipo" + Chr(13)
   End If
   If Me.cmbIdPartida.Text = "" Then
       sMensaje = sMensaje + "Elija la PARTIDA" + Chr(13)
   End If
   If txtCodigoSIS.Text = "" Then
      sMensaje = sMensaje + "Registre el CODIGO SIS" + Chr(13)
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
   Set oRsBuscaCodigo = mo_AdminComun.CatalogoServiciosSeleccionarPorCodigo(txtCodigo.Text)
   Select Case mi_Opcion
   Case sghAgregar
        'mgaray
        If validarDuplicadoServicio(oRsBuscaCodigo, 0) = False Then
            sMensaje = sMensaje + "Ese código y descripción corta ya esta Registrado para: " + oRsBuscaCodigo.Fields!Nombre + Chr(13)
        End If
'        If oRsBuscaCodigo.RecordCount > 0 Then
'            sMensaje = sMensaje + "Ese código ya esta Registrado para: " + oRsBuscaCodigo.Fields!Nombre + Chr(13)
'        End If
   Case sghModificar
        'mgaray
        If validarDuplicadoServicio(oRsBuscaCodigo, ml_IdProducto) = False Then
            sMensaje = sMensaje + "Ese código y descripción corta ya esta Registrado para: " + oRsBuscaCodigo.Fields!Nombre + Chr(13)
        End If
'        If oRsBuscaCodigo.RecordCount > 0 Then
'           oRsBuscaCodigo.MoveFirst
'           Do While Not oRsBuscaCodigo.EOF
'              If oRsBuscaCodigo.Fields!Codigo = Me.txtCodigo.Text And oRsBuscaCodigo.Fields!idProducto <> ml_IdProducto Then
'                 sMensaje = sMensaje + "Ese código ya esta Registrado para: " + oRsBuscaCodigo.Fields!Nombre + Chr(13)
'                 Exit Do
'              End If
'              oRsBuscaCodigo.MoveNext
'           Loop
'        End If
        'debb-21/09/2015
        If Len(txtCodigo.Text) > 8 Then
           sMensaje = sMensaje + "El CODIGO no debe tener longitud mayor a 8" + Chr(13)
        End If
   Case sghEliminar
       oRsBuscaCodigo.Close
       Set oRsBuscaCodigo = mo_ReglasFacturacion.FacturacionServicioPagosSeleccionarXidProducto(mo_CatalogoServicios.IdProducto)
       If oRsBuscaCodigo.RecordCount > 0 Then
          sMensaje = sMensaje + "Ese código Tiene Movimientos en tabla: FacturacionServicioPagos " + Chr(13)
       Else
          oRsBuscaCodigo.Close
          Set oRsBuscaCodigo = mo_ReglasFacturacion.FacturacionServicioDespachoSeleccionarPorIdProducto(mo_CatalogoServicios.IdProducto)
          If oRsBuscaCodigo.RecordCount > 0 Then
             sMensaje = sMensaje + "Ese código Tiene Movimientos en tabla: FacturacionServicioDespacho " + Chr(13)
          End If
       End If
       'debb-21/09/2015
        If Len(txtCodigo.Text) > 8 Then
           sMensaje = sMensaje + "El CODIGO no debe tener longitud mayor a 8" + Chr(13)
        End If
   End Select
   Set oRsBuscaCodigo = Nothing
   If mi_Opcion = sghAgregar Or mi_Opcion = sghModificar Then
      If mrs_PuntosCarga.RecordCount > 0 Then
         mrs_PuntosCarga.MoveFirst
         Do While Not mrs_PuntosCarga.EOF
            If mrs_PuntosCarga.Fields!EsPreVenta = True And mrs_PuntosCarga.Fields!TieneIdServicio = False Then
               sMensaje = sMensaje + "Para que sea 'Cabecera de Preventa', el Punto de Carga deberá tener un 'Id Servicio'" + Chr(13)
            End If
            mrs_PuntosCarga.MoveNext
         Loop
      End If
      If chkResultadoAutomatico.Visible = True And chkResultadoAutomatico.Value = 1 Then
         If Len(txtRuc.Text) <> 11 Then
            sMensaje = sMensaje + "Ingresar el RUC del Proveedor" + Chr(13)
         End If
         If txtRazonSocial.Text = "" Then
            sMensaje = sMensaje + "Ingresar la RAZON SOCIAL del Proveedor" + Chr(13)
         End If
         If cmbEquipoImg.Visible = True And cmbEquipoImg.Text = "" Then
            sMensaje = sMensaje + "Elegir el EQUIPO DE IMAGEN del Proveedor" + Chr(13)
         End If
      End If
   End If
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
   
   With mo_CatalogoServicios
        .codigoSUNAT = txtCodigoSunat.Text
        .codigo = Me.txtCodigo.Text
        .Nombre = Me.txtNombre.Text
        .IdServicioGrupo = Val(mo_cmbIdServicioGrupo.BoundText)
        .IdServicioSubGrupo = Val(mo_cmbIdServicioSubGrupo.BoundText)
        .IdServicioSeccion = Val(mo_cmbIdServicioSeccion.BoundText)
        .IdServicioSubSeccion = Val(mo_cmbIdServicioSubSeccion.BoundText)
        .IdPartida = Val(mo_cmbIdPartida.BoundText)
        .IdCentroCosto = Val(mo_cmbIdCentroCosto.BoundText)
        .IdUsuarioAuditoria = Me.idUsuario
        .EsCPT = cmbTipoServicio.ListIndex
        .NombreMINSA = Me.txtNombreMINSA.Text
        .idEstado = IIf(chkEstado.Value = 1, 1, 0)
        .codigoSIS = txtCodigoSIS.Text
        .LabResultadoAutomatico = IIf(chkResultadoAutomatico.Value = 1, 1, 0)
        .idProveedor = Val(Me.txtRuc.Tag)
        .EquipoCodigo = Right("0" & mo_cmbEquipoImg.BoundText, 2)
   End With
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   ProveedorActualizar
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminComun.CatalogoServiciosAgregar(mo_CatalogoServicios, mrs_PuntosCarga, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtCodigo.Text) & " " & txtNombre.Text)
   If mo_CatalogoServicios.IdProducto > 0 Then
        'Graba Precios
        Dim oRsTmp As New Recordset
        Dim lcSql As String
        Dim oConexion As New ADODB.Connection
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim lnPrecio As Double
        Dim oDOFactCatalogoServiciosHosp As New DOFactCatalogoServiciosHosp, oFactCatalogoServiciosHosp As New FactCatalogoServiciosHosp
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        Set oFactCatalogoServiciosHosp.Conexion = oConexion
        If mrs_Precios.RecordCount > 0 Then
           mrs_Precios.MoveFirst
           Do While Not mrs_Precios.EOF
             lnPrecio = mrs_Precios.Fields!PrecioUnitario
             Set oRsTmp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(mo_CatalogoServicios.IdProducto, mrs_Precios.Fields!idTipoFinanciamiento)
             If oRsTmp.RecordCount > 0 Then
                
                If mrs_Precios.Fields!PrecioUnitario > 0 Or mrs_Precios.Fields!SeUsaSinPrecio = True Then
                     oDOFactCatalogoServiciosHosp.IdFinanciamientoCatalogo = oRsTmp!IdFinanciamientoCatalogo
                     If oFactCatalogoServiciosHosp.SeleccionarPorId(oDOFactCatalogoServiciosHosp) = True Then
                        oDOFactCatalogoServiciosHosp.PrecioUnitario = lnPrecio
                        oDOFactCatalogoServiciosHosp.SeUsaSinPrecio = IIf(mrs_Precios.Fields!SeUsaSinPrecio = True, 1, 0)
                        If oFactCatalogoServiciosHosp.Modificar(oDOFactCatalogoServiciosHosp) = False Then
                           MsgBox oFactCatalogoServiciosHosp.MensajeError: Exit Function
                        End If
                     End If
                Else
                     oDOFactCatalogoServiciosHosp.IdFinanciamientoCatalogo = oRsTmp!IdFinanciamientoCatalogo
                     If oFactCatalogoServiciosHosp.Eliminar(oDOFactCatalogoServiciosHosp) = False Then
                        MsgBox oFactCatalogoServiciosHosp.MensajeError: Exit Function
                     End If
                End If
             Else
                If mrs_Precios.Fields!PrecioUnitario > 0 Or mrs_Precios.Fields!SeUsaSinPrecio = True Then
                    oDOFactCatalogoServiciosHosp.PrecioUnitario = lnPrecio
                    oDOFactCatalogoServiciosHosp.IdProducto = mo_CatalogoServicios.IdProducto
                    oDOFactCatalogoServiciosHosp.idTipoFinanciamiento = mrs_Precios.Fields!idTipoFinanciamiento
                    oDOFactCatalogoServiciosHosp.Activo = 1
                    oDOFactCatalogoServiciosHosp.SeUsaSinPrecio = IIf(mrs_Precios.Fields!SeUsaSinPrecio = True, 1, 0)
                    If oFactCatalogoServiciosHosp.Insertar(oDOFactCatalogoServiciosHosp) = False Then
                       MsgBox oFactCatalogoServiciosHosp.MensajeError: Exit Function
                    End If
                End If
             End If
             oRsTmp.Close
             mrs_Precios.MoveNext
           Loop
        End If
        Set oRsTmp = Nothing
        Set oConexion = Nothing
        Set oDOFactCatalogoServiciosHosp = Nothing
        Set oFactCatalogoServiciosHosp = Nothing
        
   Else
        mo_AdminComun.MensajeError = "Error, el IDproducto=0 y no se pudo registrar en la tabla."
'        MsgBox "Error, el IDproducto=0", vbInformation, Me.Caption
        AgregarDatos = False
   End If
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminComun.CatalogoServiciosModificar(mo_CatalogoServicios, mrs_PuntosCarga, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtCodigo.Text) & " " & txtNombre.Text)
   'Graba Precios
   Dim oRsTmp As New Recordset
   Dim lcSql As String
   Dim oConexion As New ADODB.Connection
   Dim oCommand As New ADODB.Command
   Dim oParameter As ADODB.Parameter
   Dim oDOFactCatalogoServiciosHosp As New DOFactCatalogoServiciosHosp, oFactCatalogoServiciosHosp As New FactCatalogoServiciosHosp
   Dim lnPrecio As Double
   oConexion.CursorLocation = adUseClient
   oConexion.CommandTimeout = 300
   oConexion.Open sighentidades.CadenaConexion
   If mrs_Precios.RecordCount > 0 Then
      Set oFactCatalogoServiciosHosp.Conexion = oConexion
      mrs_Precios.MoveFirst
      Do While Not mrs_Precios.EOF
        lnPrecio = mrs_Precios.Fields!PrecioUnitario
        Set oRsTmp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(mo_CatalogoServicios.IdProducto, mrs_Precios.Fields!idTipoFinanciamiento)
        If oRsTmp.RecordCount > 0 Then
           If mrs_Precios.Fields!PrecioUnitario > 0 Or mrs_Precios.Fields!SeUsaSinPrecio = True Then
                oDOFactCatalogoServiciosHosp.IdFinanciamientoCatalogo = oRsTmp!IdFinanciamientoCatalogo
                If oFactCatalogoServiciosHosp.SeleccionarPorId(oDOFactCatalogoServiciosHosp) = True Then
                   oDOFactCatalogoServiciosHosp.PrecioUnitario = lnPrecio
                   oDOFactCatalogoServiciosHosp.SeUsaSinPrecio = IIf(mrs_Precios.Fields!SeUsaSinPrecio = True, 1, 0)
                   If oFactCatalogoServiciosHosp.Modificar(oDOFactCatalogoServiciosHosp) = False Then
                      MsgBox oFactCatalogoServiciosHosp.MensajeError: Exit Function
                   End If
                End If
           Else
                oDOFactCatalogoServiciosHosp.IdFinanciamientoCatalogo = oRsTmp!IdFinanciamientoCatalogo
                If oFactCatalogoServiciosHosp.Eliminar(oDOFactCatalogoServiciosHosp) = False Then
                   MsgBox oFactCatalogoServiciosHosp.MensajeError: Exit Function
                End If
           End If
        Else
           If mrs_Precios.Fields!PrecioUnitario > 0 Or mrs_Precios.Fields!SeUsaSinPrecio = True Then
                oDOFactCatalogoServiciosHosp.PrecioUnitario = lnPrecio
                oDOFactCatalogoServiciosHosp.IdProducto = mo_CatalogoServicios.IdProducto
                oDOFactCatalogoServiciosHosp.idTipoFinanciamiento = mrs_Precios.Fields!idTipoFinanciamiento
                oDOFactCatalogoServiciosHosp.Activo = 1
                oDOFactCatalogoServiciosHosp.SeUsaSinPrecio = IIf(mrs_Precios.Fields!SeUsaSinPrecio = True, 1, 0)
                If oFactCatalogoServiciosHosp.Insertar(oDOFactCatalogoServiciosHosp) = False Then
                   MsgBox oFactCatalogoServiciosHosp.MensajeError: Exit Function
                End If
           End If
        End If
        oRsTmp.Close
        mrs_Precios.MoveNext
      Loop
   End If
   Set oRsTmp = Nothing
   Set oDOFactCatalogoServiciosHosp = Nothing
   Set oFactCatalogoServiciosHosp = Nothing
   ProveedorActualizar
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   'Graba Precios
   Dim oRsTmp As New Recordset
   Dim lcSql As String
   Set oRsTmp = mo_ReglasFacturacion.CatalogoServiciosHospEliminarXidProducto(mo_CatalogoServicios.IdProducto)
   '
   EliminarDatos = mo_AdminComun.CatalogoServiciosEliminar(mo_CatalogoServicios, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(txtCodigo.Text) & " " & txtNombre.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

    Set mo_CatalogoServicios = mo_AdminComun.CatalogoServiciosSeleccionarPorId(Me.IdProducto)
    If mo_AdminComun.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminComun.MensajeError, vbInformation, Me.Caption"
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_CatalogoServicios Is Nothing Then
        With mo_CatalogoServicios
            txtCodigoSunat.Text = .codigoSUNAT
            Me.txtNombre = .Nombre
            Me.txtCodigo = .codigo
            Me.txtNombreMINSA = .NombreMINSA
            
            mo_cmbIdServicioGrupo.BoundText = .IdServicioGrupo
            mo_cmbIdServicioSubGrupo.BoundText = .IdServicioSubGrupo
            mo_cmbIdServicioSeccion.BoundText = .IdServicioSeccion
            mo_cmbIdServicioSubSeccion.BoundText = .IdServicioSubSeccion
            
            mo_cmbIdPartida.BoundText = .IdPartida
            mo_cmbIdCentroCosto.BoundText = .IdCentroCosto
            cmbTipoServicio.ListIndex = .EsCPT
            lblId.Caption = Me.IdProducto
            chkEstado.Value = IIf(.idEstado = 1, 1, 0)
            txtCodigoSIS.Text = .codigoSIS
            chkResultadoAutomatico.Value = IIf(.LabResultadoAutomatico = 1, 1, 0)
            txtRuc.Tag = IIf(.idProveedor > 0, Trim(Str(.idProveedor)), "")
            mo_cmbEquipoImg.BoundText = .EquipoCodigo
            mb_ExistenDatos = True
        End With
        'carga puntos  carga
        Dim oFactCatalogoServiciosPtos As New Recordset
        Dim lcSql As String
        Dim lcIdServicio As String
        Set oFactCatalogoServiciosPtos = mo_AdminComun.FactCatalogoServiciosPtosSeleccionarXidProducto(Me.IdProducto)
        If oFactCatalogoServiciosPtos.RecordCount > 0 Then
           oFactCatalogoServiciosPtos.MoveFirst
           Do While Not oFactCatalogoServiciosPtos.EOF
              lcIdServicio = devuelveIdServicio(oFactCatalogoServiciosPtos.Fields!idPuntoCarga)
              mrs_PuntosCarga.AddNew
              mrs_PuntosCarga.Fields!idPuntoCarga = oFactCatalogoServiciosPtos.Fields!idPuntoCarga
              mrs_PuntosCarga.Fields!Descripcion = Trim(oFactCatalogoServiciosPtos.Fields!Descripcion) & lcIdServicio
              mrs_PuntosCarga.Fields!EsPreVenta = IIf(oFactCatalogoServiciosPtos.Fields!EsPreVenta = True, True, False)
              mrs_PuntosCarga.Fields!TieneIdServicio = IIf(lcIdServicio = "", False, True)
              mrs_PuntosCarga.Update
              oFactCatalogoServiciosPtos.MoveNext
           Loop
        End If
        oFactCatalogoServiciosPtos.Close
        'Carga  Precios
        Set oFactCatalogoServiciosPtos = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarXidProducto(Me.IdProducto)
        If oFactCatalogoServiciosPtos.RecordCount > 0 Then
           oFactCatalogoServiciosPtos.MoveFirst
           Do While Not oFactCatalogoServiciosPtos.EOF
              mrs_Precios.MoveFirst
              mrs_Precios.Find "idTipoFinanciamiento=" & oFactCatalogoServiciosPtos.Fields!idTipoFinanciamiento
              If Not mrs_Precios.EOF Then
                 mrs_Precios.Fields!PrecioUnitario = oFactCatalogoServiciosPtos.Fields!PrecioUnitario
                 mrs_Precios.Fields!SeUsaSinPrecio = IIf(IsNull(oFactCatalogoServiciosPtos.Fields!SeUsaSinPrecio), False, oFactCatalogoServiciosPtos.Fields!SeUsaSinPrecio)
                 mrs_Precios.Update
              End If
              oFactCatalogoServiciosPtos.MoveNext
           Loop
           mrs_Precios.MoveFirst
        End If
        'Set Me.grdPrecios.DataSource = mrs_Precios
        oFactCatalogoServiciosPtos.Close
        Set oFactCatalogoServiciosPtos = Nothing
        MuestraLblAutomatico
        ProveedorBuscaRazonSocial
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
    mo_cmbIdServicioGrupo.BoundText = ""
    mo_cmbIdServicioSeccion.BoundText = ""
    mo_cmbIdPartida.BoundText = ""
    cmbTipoServicio.ListIndex = -1
    mo_cmbIdServicioSubSeccion.BoundText = ""
    mo_cmbIdServicioSubGrupo.BoundText = ""
    mo_cmbIdPuntoCarga.BoundText = ""
    Me.txtNombre = ""
    Me.txtCodigo = ""
    Me.txtNombreMINSA.Text = ""
    lblId.Caption = ""
    chkEstado.Value = 1
    txtCodigoSIS.Text = ""
    '*******GalenHos V3 Inicio ***********
    CreaTemporal
    '*******GalenHos V3 Final ***********
    mo_Apariencia.ConfigurarFilasBiColores grdPrecios, sighentidades.GrillaConFilasBicolor
    mo_Apariencia.ConfigurarFilasBiColores grdPuntosDeCarga, sighentidades.GrillaConFilasBicolor
    'Me.txtPrecioUnitario = ""
End Sub

Sub CargarComboBoxes()

    mo_cmbIdCentroCosto.BoundColumn = "IdCentroCosto"
    mo_cmbIdCentroCosto.ListField = "Descripcion"
    Set mo_cmbIdCentroCosto.RowSource = mo_AdminComun.CentrosCostoSeleccionarTodos

    mo_cmbIdPartida.BoundColumn = "IdPartidaPresupuestal"
    mo_cmbIdPartida.ListField = "Descripcion"
    Set mo_cmbIdPartida.RowSource = mo_AdminComun.PartidasPresupuestalesSeleccionarTodos

    mo_cmbIdServicioGrupo.BoundColumn = "IdServicioGrupo"
    mo_cmbIdServicioGrupo.ListField = "Descripcion"
    Set mo_cmbIdServicioGrupo.RowSource = mo_AdminComun.CatalogoServiciosGrupoSeleccionarTodos
    
    Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPtoCarga
    mo_cmbIdPuntoCarga.ListField = "Descripcion"
    mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
    Set mo_cmbIdPuntoCarga.RowSource = mo_AdminComun.SeleccionarPuntosDeCarga()
    
    CArgaEquiposIMG

    
End Sub

Sub CArgaEquiposIMG()
    mo_cmbEquipoImg.ListField = "Equipo"
    mo_cmbEquipoImg.BoundColumn = "codigo"
    Set mo_cmbEquipoImg.RowSource = mo_AdminComun.InteroperaEquiposSeleccionarTodos()

End Sub



Private Sub grdPrecios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPrecios.Bands(0).Columns("IdTipoFinanciamiento").Hidden = True
    '
    grdPrecios.Bands(0).Columns("TipoFinanciamiento").Header.Caption = "Producto/Plan"
    grdPrecios.Bands(0).Columns("TipoFinanciamiento").Width = 2500
    '
    grdPrecios.Bands(0).Columns("PrecioUnitario").Header.Caption = "Precio Unitario"
    grdPrecios.Bands(0).Columns("PrecioUnitario").Width = 700
    '
    grdPrecios.Bands(0).Columns("SeUsaSinPrecio").Header.Caption = "Se usará con Precio=0"
    grdPrecios.Bands(0).Columns("SeUsaSinPrecio").Width = 800
    grdPrecios.Bands(0).ColHeaderLines = 4
End Sub


Private Sub grdPuntosDeCarga_Click()
            On Error Resume Next
            Select Case grdPuntosDeCarga.ActiveCell.Column.Key
            Case "EsPreVenta"
                If mrs_PuntosCarga.Fields!TieneIdServicio = False Then
                   mrs_PuntosCarga.Fields!EsPreVenta = False
                End If
            End Select

End Sub

Private Sub grdPuntosDeCarga_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPuntosDeCarga.Bands(0).Columns("idPuntoCarga").Width = 300
    grdPuntosDeCarga.Bands(0).Columns("Descripcion").Width = 2700
    grdPuntosDeCarga.Bands(0).Columns("EsPreVenta").Width = 1000
    grdPuntosDeCarga.Bands(0).Columns("TieneIdServicio").Hidden = True
End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode

End Sub



Private Sub txtCodigoSIS_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigoSIS
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNombre
    AdministrarKeyPreview KeyCode

End Sub


Sub CreaTemporal()
    '*******GalenHos V3 Inicio ***********
    If mrs_PuntosCarga.State = adStateOpen Then mrs_PuntosCarga.Close
    '*******GalenHos V3 Final ***********
    With mrs_PuntosCarga
          .Fields.Append "IdPuntoCarga", adInteger, 4, adFldIsNullable
          .Fields.Append "Descripcion", adVarChar, 200, adFldIsNullable
          .Fields.Append "EsPreVenta", adBoolean
          .Fields.Append "TieneIdServicio", adBoolean
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdPuntosDeCarga.DataSource = mrs_PuntosCarga
    '*******GalenHos V3 Inicio ***********
    If mrs_Precios.State = adStateOpen Then mrs_Precios.Close
    '*******GalenHos V3 Final ***********
    With mrs_Precios
          .Fields.Append "IdTipoFinanciamiento", adInteger, 4, adFldIsNullable
          .Fields.Append "TipoFinanciamiento", adVarChar, 200, adFldIsNullable
          .Fields.Append "PrecioUnitario", adDouble
          .Fields.Append "SeUsaSinPrecio", adBoolean
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
          mrs_Precios.Fields!PrecioUnitario = 0
          mrs_Precios.Update
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set Me.grdPrecios.DataSource = mrs_Precios
    mrs_Precios.MoveFirst
End Sub

Private Sub txtNombreMINSA_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNombreMINSA
    AdministrarKeyPreview KeyCode

End Sub
'mgaray
Private Function validarDuplicadoServicio(oRsBuscaCodigo As ADODB.Recordset, _
            lIdProducto As Long) As Boolean
    Dim bReturnValue As Boolean
    bReturnValue = True
    If oRsBuscaCodigo.RecordCount > 0 Then
       oRsBuscaCodigo.MoveFirst
       Do While Not oRsBuscaCodigo.EOF
        'mgaray20141013
        ''And UCase(Trim(oRsBuscaCodigo.Fields!Nombre)) = UCase(Trim(Me.txtNombre.Text)))
          If (UCase(Trim(oRsBuscaCodigo.Fields!codigo)) = UCase(Trim(Me.txtCodigo.Text))) _
                And oRsBuscaCodigo.Fields!IdProducto <> lIdProducto Then
             bReturnValue = False
             Exit Do
          End If
          oRsBuscaCodigo.MoveNext
       Loop
    End If
    validarDuplicadoServicio = bReturnValue
End Function



Private Sub txtRuc_LostFocus()
    ProveedorBuscaRazonSocial
End Sub

Sub ProveedorBuscaRazonSocial()
    mo_ReglasFacturacion.ProveedorBuscaDatos txtRuc, txtRazonSocial
End Sub
Private Sub txtRuc_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtRuc
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       ElseIf KeyAscii = 13 Then
           txtRuc_LostFocus
       End If
End Sub
Private Sub txtRazonSocial_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtRazonSocial
    AdministrarKeyPreview KeyCode
End Sub

Sub ProveedorActualizar()
    On Error GoTo ErrProv
    If chkResultadoAutomatico.Visible = True Then
        Dim oProveedores As New Proveedores
        Dim oDoProveedores As New DoProveedores
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        Set oProveedores.Conexion = oConexion
        oDoProveedores.IdUsuarioAuditoria = sighentidades.Usuario
        If Val(txtRuc.Tag) > 0 Then
            oDoProveedores.idProveedor = Val(txtRuc.Tag)
            If oProveedores.SeleccionarPorId(oDoProveedores) Then
                oDoProveedores.razonSocial = txtRazonSocial.Text
                If Not oProveedores.Modificar(oDoProveedores) Then
                End If
            End If
        Else
            oDoProveedores.ruc = txtRuc.Text
            oDoProveedores.razonSocial = txtRazonSocial.Text
            If Not oProveedores.Insertar(oDoProveedores) Then
               txtRuc.Tag = Trim(Str(oDoProveedores.idProveedor))
            End If
        End If
        oConexion.Close
        Set oProveedores = Nothing
        Set oDoProveedores = Nothing
    End If
    Exit Sub
ErrProv:
    MsgBox Err.Description
    Exit Sub
    Resume
End Sub
