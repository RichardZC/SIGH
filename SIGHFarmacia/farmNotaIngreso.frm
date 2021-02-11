VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form FarmNotaIngreso 
   Caption         =   "Form1"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15180
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "farmNotaIngreso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   15180
   StartUpPosition =   2  'CenterScreen
   Begin SighFarmacia.ucNotaIngreso grdProductos 
      Height          =   3915
      Left            =   90
      TabIndex        =   44
      Top             =   3885
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   6906
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   90
      TabIndex        =   39
      Top             =   7770
      Width           =   15045
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "farmNotaIngreso.frx":0CCA
         DownPicture     =   "farmNotaIngreso.frx":118E
         Height          =   700
         Left            =   7703
         Picture         =   "farmNotaIngreso.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "farmNotaIngreso.frx":1B66
         DownPicture     =   "farmNotaIngreso.frx":1FC6
         Height          =   700
         Left            =   6120
         Picture         =   "farmNotaIngreso.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprime"
         Height          =   700
         Left            =   135
         Picture         =   "farmNotaIngreso.frx":28B0
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Visible         =   0   'False
         Width           =   1365
      End
   End
   Begin VB.Frame p 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   60
      TabIndex        =   16
      Top             =   90
      Width           =   15045
      Begin VB.TextBox txtdocext 
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
         Left            =   8010
         MaxLength       =   20
         TabIndex        =   62
         Top             =   1920
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CheckBox chkPacienteForaneo 
         Caption         =   "Paciente Foráneo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6600
         TabIndex        =   59
         Top             =   2400
         Width           =   1890
      End
      Begin VB.Frame fraPacienteForaneo 
         Height          =   1305
         Left            =   6600
         TabIndex        =   54
         Top             =   960
         Visible         =   0   'False
         Width           =   8235
         Begin VB.TextBox txtSerieBoleta 
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
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   55
            Top             =   360
            Width           =   705
         End
         Begin VB.TextBox txtNumeroBoleta 
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
            Left            =   2160
            MaxLength       =   8
            TabIndex        =   57
            Top             =   360
            Width           =   1665
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "N° Boleta"
            Height          =   210
            Left            =   360
            TabIndex        =   58
            Top             =   405
            Width           =   780
         End
         Begin VB.Label lblGuion 
            Caption         =   "-"
            Height          =   375
            Index           =   1
            Left            =   2040
            TabIndex        =   56
            Top             =   360
            Width           =   135
         End
      End
      Begin UltraGrid.SSUltraGrid grdConsumoPaciente 
         Height          =   2295
         Left            =   1500
         TabIndex        =   41
         Top             =   105
         Visible         =   0   'False
         Width           =   11880
         _ExtentX        =   20955
         _ExtentY        =   4048
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   71303188
         BorderStyle     =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   "farmNotaIngreso.frx":2D89
         Caption         =   "Consumos de la CUENTA"
      End
      Begin VB.TextBox txtNotaIngreso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1545
         MaxLength       =   30
         TabIndex        =   20
         Top             =   165
         Width           =   1635
      End
      Begin VB.TextBox txtEstado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   13455
         MaxLength       =   30
         TabIndex        =   19
         Top             =   165
         Width           =   1395
      End
      Begin VB.TextBox txtNdocum 
         Enabled         =   0   'False
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
         Left            =   8010
         MaxLength       =   20
         TabIndex        =   2
         Top             =   900
         Width           =   1425
      End
      Begin VB.TextBox txtNdocO 
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
         Left            =   8010
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1230
         Width           =   1425
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
         Left            =   1545
         MaxLength       =   11
         TabIndex        =   9
         ToolTipText     =   "Ingrese el N° RUC"
         Top             =   2325
         Width           =   1215
      End
      Begin VB.TextBox txtProveedor 
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
         Left            =   2820
         MaxLength       =   100
         TabIndex        =   18
         ToolTipText     =   "Ingrese la Razón Social"
         Top             =   2325
         Width           =   3555
      End
      Begin VB.TextBox txtNproceso 
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
         Left            =   13605
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1950
         Width           =   1245
      End
      Begin VB.TextBox txtObservaciones 
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
         Left            =   1545
         MaxLength       =   100
         TabIndex        =   10
         Top             =   2700
         Width           =   4845
      End
      Begin VB.ComboBox cmbAlmDestino 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11835
         TabIndex        =   15
         Top             =   510
         Width           =   3045
      End
      Begin VB.ComboBox cmbTipoDocum 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1545
         TabIndex        =   13
         Top             =   1260
         Width           =   4860
      End
      Begin VB.ComboBox cmbAlmOrigen 
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
         Left            =   1545
         TabIndex        =   1
         Top             =   870
         Width           =   4860
      End
      Begin VB.ComboBox cmbConcepto 
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
         Left            =   1545
         TabIndex        =   0
         Top             =   510
         Width           =   4860
      End
      Begin VB.ComboBox cmbTipodocumO 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1545
         TabIndex        =   14
         Top             =   1590
         Width           =   4860
      End
      Begin VB.ComboBox cmbTipoCompra 
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
         Left            =   1545
         TabIndex        =   6
         Top             =   1950
         Width           =   4860
      End
      Begin VB.ComboBox cmbTproceso 
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
         Left            =   8010
         TabIndex        =   7
         Top             =   1590
         Width           =   1440
      End
      Begin VB.TextBox txtHoraRegistro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   9345
         MaxLength       =   30
         TabIndex        =   17
         Top             =   180
         Width           =   735
      End
      Begin MSMask.MaskEdBox txtFregistro 
         Height          =   315
         Left            =   7950
         TabIndex        =   21
         Top             =   195
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox txtFrecepcion 
         Height          =   315
         Left            =   13590
         TabIndex        =   3
         Top             =   1260
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFdocO 
         Height          =   315
         Left            =   13590
         TabIndex        =   5
         Top             =   1590
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdBuscaZIP 
         Caption         =   "Importar"
         Height          =   330
         Left            =   9465
         TabIndex        =   43
         ToolTipText     =   "Debe existir el ODBC HIS q apunte a c:\archiv...\dig...\gal...\archivos"
         Top             =   900
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Frame frmPaciente 
         Height          =   1320
         Left            =   6585
         TabIndex        =   45
         Top             =   2415
         Width           =   8235
         Begin VB.ComboBox cmbUnidosis 
            Height          =   330
            Left            =   2745
            TabIndex        =   60
            Top             =   930
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2745
            TabIndex        =   51
            ToolTipText     =   "Busca Cuenta por Apellidos y Nombres"
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtNcuenta 
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
            Left            =   1290
            MaxLength       =   30
            TabIndex        =   50
            Top             =   240
            Width           =   1425
         End
         Begin VB.TextBox txtDatosDeCuenta 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3075
            TabIndex        =   49
            Top             =   240
            Width           =   5070
         End
         Begin VB.TextBox txtNhistoria 
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
            Left            =   1290
            MaxLength       =   30
            TabIndex        =   48
            ToolTipText     =   "Ingrese el Nro de Historia Clínica"
            Top             =   630
            Width           =   1425
         End
         Begin VB.TextBox txtNombrePaciente 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3090
            MaxLength       =   30
            TabIndex        =   47
            Top             =   615
            Width           =   5055
         End
         Begin VB.CommandButton btnBuscarPaciente 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2760
            TabIndex        =   46
            ToolTipText     =   "Busca por Apellidos y Nombres"
            Top             =   630
            Width           =   315
         End
         Begin VB.Label lblUnidosis 
            AutoSize        =   -1  'True
            Caption         =   "Farmacia UNIDOSIS (origen)"
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   345
            TabIndex        =   61
            Top             =   990
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label lblNcuenta 
            AutoSize        =   -1  'True
            Caption         =   "N° Cuenta"
            Height          =   210
            Left            =   360
            TabIndex        =   53
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Paciente"
            Height          =   210
            Left            =   480
            TabIndex        =   52
            Top             =   720
            Width           =   705
         End
      End
      Begin VB.Label lblNDoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° Doc. Externo"
         Height          =   210
         Left            =   6600
         TabIndex        =   63
         Top             =   1920
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   210
         Left            =   105
         TabIndex        =   38
         Top             =   540
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F.Registro"
         Height          =   210
         Left            =   7125
         TabIndex        =   37
         Top             =   225
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "N° Nota Ingreso"
         Height          =   210
         Left            =   105
         TabIndex        =   36
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   210
         Left            =   12825
         TabIndex        =   35
         Top             =   195
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
         Height          =   210
         Left            =   105
         TabIndex        =   34
         Top             =   930
         Width           =   540
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Almacén destino"
         Height          =   210
         Left            =   10425
         TabIndex        =   33
         Top             =   585
         Width           =   1350
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Docum"
         Height          =   210
         Left            =   105
         TabIndex        =   32
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "N° Docum"
         Height          =   210
         Left            =   7110
         TabIndex        =   31
         Top             =   900
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "F.Recepción"
         Height          =   210
         Left            =   12405
         TabIndex        =   30
         Top             =   1260
         Width           =   990
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc. Origen"
         Height          =   210
         Left            =   105
         TabIndex        =   29
         Top             =   1650
         Width           =   1395
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "N° Doc. Origen"
         Height          =   210
         Left            =   6705
         TabIndex        =   28
         Top             =   1230
         Width           =   1245
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "F.Doc. Origen"
         Height          =   210
         Left            =   12270
         TabIndex        =   27
         Top             =   1590
         Width           =   1125
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "T/compra"
         Height          =   210
         Left            =   105
         TabIndex        =   26
         Top             =   2040
         Width           =   795
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "T/proceso"
         Height          =   210
         Left            =   7110
         TabIndex        =   25
         Top             =   1590
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "N° Proceso"
         Height          =   210
         Left            =   12480
         TabIndex        =   24
         Top             =   1950
         Width           =   915
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   210
         Left            =   105
         TabIndex        =   23
         Top             =   2385
         Width           =   840
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   210
         Left            =   105
         TabIndex        =   22
         Top             =   2730
         Width           =   1170
      End
   End
   Begin UltraGrid.SSUltraGrid grdProductosDevol 
      Height          =   3690
      Left            =   0
      TabIndex        =   42
      Top             =   3885
      Visible         =   0   'False
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   6509
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   71303188
      BorderStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   "farmNotaIngreso.frx":2DC5
      Caption         =   "grdProductosDevol"
   End
End
Attribute VB_Name = "FarmNotaIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Notas de Ingreso
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim gridInfra As New GridInfragistic
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_movNumero As String
Dim oRsConceptos As New ADODB.Recordset
Dim oRsAlmacenDestino As New ADODB.Recordset
Dim rsTmp As New ADODB.Recordset
Dim mRs_ProductosDevol As New Recordset
Dim mRs_ConsumoPaciente As New Recordset
Dim mo_cmbUnidosis As New SIGHEntidades.ListaDespleglable
Dim mo_cmbConceptos As New SIGHEntidades.ListaDespleglable
Dim mo_cmbAlmacenOrigen As New SIGHEntidades.ListaDespleglable
Dim mo_cmbAlmacenDestino As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTipoDocum As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTipoDocumO As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTipoCompra As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTproceso As New SIGHEntidades.ListaDespleglable
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mRs_Productos As New ADODB.Recordset
Dim oRsItemsUnidosis As New Recordset
Dim oRsItemsUNIDOSISpNS As New Recordset
Dim mo_farmMovimiento1 As New DoFarmMovimiento
Dim mo_farmMovimiento As New sighComun.DoFarmMovimiento
Dim mo_farmMovimientoNotaIngreso As New sighComun.DOfarmMovimientoNotaIngreso
Dim oDoProveedores As New DoProveedores
Const lcConstanteMovimientoEntrada As String = "E"
Const lcConstanteMovimientoSalida As String = "S"
Const lnTipoConceptoAjusteInventario As Long = 20
Dim lnTotalDocumento As Double
Dim ml_IdPaciente As Long: Dim ml_IdComprobantePago As Long
Dim ml_IdProveedor As Long
Dim ms_MensajeError As String
Dim mo_DoPaciente As New DOPaciente
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminAdmision As New ReglasAdmision
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion

Dim mo_ReglasFacturacion1 As New SIGHNegocios.ReglasFacturacion

Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim ml_idFuenteFinanciamiento As Long
Dim ml_IdTipoFinanciamiento As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim mo_lbElEstablecimentoEsCS As Boolean
Dim wxParametro302 As String
Dim lnIdTipoServicio As Long
Dim ml_idUsuarioCreo As Long
Const LcIdTipoConceptoDevolucionPaciente As String = "21"
Const LcIdTipoDocumentoNINGUNO As Long = 22
Dim lcMensajeLicencia As String
Dim lbElUsuarioTrabajaAqui As Boolean
Dim lbLaFarmaciaEsUnidosis As Boolean, lbLaFuenteFinanciamientoUsadoEnFUnidosis As Boolean, lnCuentaUnidosis As Long


Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let movNumero(lValue As String)
   ml_movNumero = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property


Sub CargaDatosProveedor()
   Dim oProveedor As New Proveedores
   Dim oConexion As New ADODB.Connection
   oConexion.CommandTimeout = 900
   oConexion.CursorLocation = adUseClient
   oConexion.Open SIGHEntidades.CadenaConexion
   Set oProveedor.Conexion = oConexion
   If oProveedor.SeleccionarPorId(oDoProveedores) Then
      txtRuc.Text = oDoProveedores.ruc
      txtProveedor.Text = oDoProveedores.razonSocial
      mo_Formulario.HabilitarDeshabilitar Me.txtProveedor, False
   End If
   Set oProveedor = Nothing
   Set oConexion = Nothing
   
End Sub


Private Sub chkPacienteForaneo_Click()
    If chkPacienteForaneo.Value = 1 Then
        Me.fraPacienteForaneo.Visible = True
        BlanquedaVariablesUnidosis
    Else
        Me.fraPacienteForaneo.Visible = False
        
    End If
End Sub

Private Sub cmbAlmDestino_Click()
    oRsAlmacenDestino.MoveFirst
    If mo_cmbAlmacenDestino.BoundText <> "" Then
        oRsAlmacenDestino.Find "idAlmacen=" & mo_cmbAlmacenDestino.BoundText
        Set oRsConceptos = mo_ReglasFarmacia.FarmTipoConceptosDevuelveParaRegistroDeNiNs(oRsAlmacenDestino.Fields!idTipoLocales, lcConstanteMovimientoEntrada, oRsAlmacenDestino.Fields!idTipoSuministro)
        mo_cmbConceptos.BoundColumn = "IdTipoConcepto"
        mo_cmbConceptos.ListField = "Concepto"
        Set mo_cmbConceptos.RowSource = mo_ReglasFarmacia.FarmTipoConceptosDevuelveParaRegistroDeNiNs(oRsAlmacenDestino.Fields!idTipoLocales, lcConstanteMovimientoEntrada, oRsAlmacenDestino.Fields!idTipoSuministro)
        grdProductos.IdAlmacen = oRsAlmacenDestino.Fields!IdAlmacen
        Me.grdConsumoPaciente.Visible = False
        ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
        lbLaFarmaciaEsUnidosis = mo_ReglasFarmacia.FarmaciaEsUnidosis(Val(mo_cmbAlmacenDestino.BoundText))
        BlanquedaVariablesUnidosis
    End If
End Sub

Private Sub btnCancelar_Click()
   If SIGHEntidades.ParaAuditoria = "" Then
      Me.Visible = False
      LimpiarVariablesDeMemoria
   ElseIf MsgBox("Hubo cambios, desea salir de todas maneras ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
      Me.Visible = False
      LimpiarVariablesDeMemoria
      SIGHEntidades.ParaAuditoriaPorCadaDato sghAudLimpiar, ""
   End If
End Sub


Private Sub cmbAlmDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmDestino

End Sub

Private Sub cmbAlmOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmOrigen

End Sub

Private Sub cmbConcepto_Click()
    Dim lcFiltro As String
    If oRsConceptos.State = 0 Then Exit Sub
    oRsConceptos.MoveFirst
    oRsConceptos.Find "idTipoConcepto=" & mo_cmbConceptos.BoundText
    mo_cmbTipoDocum.BoundText = oRsConceptos.Fields!DocumentoId
    
    'RHA 12/01/2021 CAMBIO 50 INCIO
     If Trim(Me.cmbConcepto.Text) = "COMPRA" Then
        Me.txtdocext.Visible = True
    Me.lblNDoc.Visible = True
    Else
        Me.txtdocext.Visible = False
    Me.lblNDoc.Visible = False
    End If
    'RHA 12/01/2021 CAMBIO 50 FIN
    
    If oRsConceptos.Fields!DocumentoId = LcIdTipoDocumentoNINGUNO Then     'Ninguno
       mo_Formulario.HabilitarDeshabilitar txtNdocum, False
       mo_Formulario.HabilitarDeshabilitar txtFrecepcion, False
    Else
    'SCCQ 13/10/2020 Cambio28 Inicio
       'mo_Formulario.HabilitarDeshabilitar txtNdocum, True
    'SCCQ 13/10/2020 Cambio28 Fin
       mo_Formulario.HabilitarDeshabilitar txtFrecepcion, True
    End If
    '
    mo_cmbTipoDocumO.BoundText = oRsConceptos.Fields!NiDocumentoOrigenId
    If oRsConceptos.Fields!NiDocumentoOrigenId = LcIdTipoDocumentoNINGUNO Then    'Ninguno
       mo_Formulario.HabilitarDeshabilitar txtNdocO, False
       mo_Formulario.HabilitarDeshabilitar txtFdocO, False
    Else
       mo_Formulario.HabilitarDeshabilitar txtNdocO, True
       mo_Formulario.HabilitarDeshabilitar txtFdocO, True
    End If
    '
    
    If oRsConceptos.Fields!NiEsCompra Then
       If oRsConceptos.Fields!conceptoCodigo = "01" Then
            mo_Formulario.HabilitarDeshabilitar Me.cmbTproceso, True
            mo_Formulario.HabilitarDeshabilitar Me.cmbTipoCompra, True
            mo_Formulario.HabilitarDeshabilitar Me.txtNproceso, True
       Else
            mo_Formulario.HabilitarDeshabilitar Me.cmbTproceso, False
            mo_Formulario.HabilitarDeshabilitar Me.cmbTipoCompra, False
            mo_Formulario.HabilitarDeshabilitar Me.txtNproceso, False
       End If
       mo_Formulario.HabilitarDeshabilitar Me.txtRuc, True
       mo_Formulario.HabilitarDeshabilitar Me.txtProveedor, True
       '
       lcFiltro = "idTipoCompra<>1"
       mo_cmbTipoCompra.BoundColumn = "idTipoCompra"
       mo_cmbTipoCompra.ListField = "Descripcion"
       Set mo_cmbTipoCompra.RowSource = mo_ReglasFarmacia.FarmTipoCompraDevuelveSegunFiltro(lcFiltro)
       ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
       '
       lcFiltro = "idTipoProceso<>1"
       mo_cmbTproceso.BoundColumn = "idTipoProceso"
       mo_cmbTproceso.ListField = "Descripcion"
       Set mo_cmbTproceso.RowSource = mo_ReglasFarmacia.FarmTipoProcesoDevuelveSegunFiltro(lcFiltro)
       ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    Else
       mo_Formulario.HabilitarDeshabilitar Me.cmbTproceso, False
       mo_Formulario.HabilitarDeshabilitar Me.cmbTipoCompra, False
       mo_Formulario.HabilitarDeshabilitar Me.txtNproceso, False
       mo_Formulario.HabilitarDeshabilitar Me.txtRuc, False
       mo_Formulario.HabilitarDeshabilitar Me.txtProveedor, False
       '
       lcFiltro = ""
       mo_cmbTipoCompra.BoundColumn = "idTipoCompra"
       mo_cmbTipoCompra.ListField = "Descripcion"
       Set mo_cmbTipoCompra.RowSource = mo_ReglasFarmacia.FarmTipoCompraDevuelveSegunFiltro(lcFiltro)
       ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
       mo_cmbTipoCompra.BoundText = "1"
       '
       lcFiltro = ""
       mo_cmbTproceso.BoundColumn = "idTipoProceso"
       mo_cmbTproceso.ListField = "Descripcion"
       Set mo_cmbTproceso.RowSource = mo_ReglasFarmacia.FarmTipoProcesoDevuelveSegunFiltro(lcFiltro)
       ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
       mo_cmbTproceso.BoundText = "1"
    End If
    '
    If lbElUsuarioTrabajaAqui = True Then
       mo_Formulario.HabilitarDeshabilitar cmbAlmDestino, False
    End If

    Me.chkPacienteForaneo.Value = 0
    fraPacienteForaneo.Visible = False
    grdConsumoPaciente.Visible = False
    Me.txtSerieBoleta.Text = ""
    Me.txtNumeroBoleta.Text = ""
    grdProductos.Visible = True
    Me.grdProductosDevol.Visible = False
    If oRsConceptos.Fields!NiEsDevolucionPaciente Then
       grdProductos.Visible = False
       Me.grdProductosDevol.Visible = True
       mo_Formulario.HabilitarDeshabilitar chkPacienteForaneo, True
       mo_Formulario.HabilitarDeshabilitar txtNhistoria, True
       mo_Formulario.HabilitarDeshabilitar txtNcuenta, True
       mo_Formulario.HabilitarDeshabilitar txtNdocO, False
       mo_Formulario.HabilitarDeshabilitar txtFdocO, False
       mo_Formulario.HabilitarDeshabilitar txtFrecepcion, True

       btnBuscarPaciente.Enabled = True
       txtFrecepcion.Text = Date
    Else
       mo_Formulario.HabilitarDeshabilitar chkPacienteForaneo, False
       mo_Formulario.HabilitarDeshabilitar txtNhistoria, False
       mo_Formulario.HabilitarDeshabilitar txtNcuenta, False
       btnBuscarPaciente.Enabled = False
    End If
    
    '
    mo_cmbAlmacenOrigen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenOrigen.ListField = "Descripcion"
    If mo_lbElEstablecimentoEsCS = True Then
       Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro(oRsConceptos.Fields!NiFiltroAlmacenOrigenCS & " and idEstado=1")
    Else
       Set mo_cmbAlmacenOrigen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro(oRsConceptos.Fields!NiFiltroAlmacenOrigen & " and idEstado=1")
    End If
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    If cmbAlmOrigen.ListCount = 1 Then
       cmbAlmOrigen.ListIndex = 0
    End If
    '
    BlanquedaVariablesUnidosis
    '
    grdProductos.TipoPrecioParaNiNs = oRsConceptos.Fields!TipoPrecioParaNiNs
    grdProductos.TipoConcepto = oRsConceptos.Fields!idTipoConcepto
    If Val(mo_cmbConceptos.BoundText) = 3 Then
       'Donaciones
       grdProductos.EsUnaDonacionOestrategico = sghTipoSalidaItemFarmacia.sghDonaciones
       mo_Formulario.HabilitarDeshabilitar txtRuc, True
    ElseIf Val(mo_cmbConceptos.BoundText) = 8 Then
       'Estrategicos
       grdProductos.EsUnaDonacionOestrategico = sghTipoSalidaItemFarmacia.sghSoloEstrategico
    Else
       grdProductos.EsUnaDonacionOestrategico = 0
    End If
    
    
End Sub

Private Sub cmbConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbConcepto

End Sub







Private Sub cmbTipoCompra_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoCompra

End Sub



Private Sub cmbTproceso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTproceso

End Sub



Private Sub cmbUnidosis_Click()
   FiltraItemsDeLaFarmaciaUnidosisElejida
End Sub

Private Sub cmdBuscaCuentaPorApellidos_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New DOPaciente
    Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            ml_IdPaciente = oDOPaciente.IdPaciente
            txtNhistoria.Text = oDOPaciente.NroHistoriaClinica
            txtNombrePaciente.Text = Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + oDOPaciente.PrimerNombre
            Dim oRsTmp As New Recordset
            Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(ml_IdPaciente, oConexion, True)
            If oRsTmp.RecordCount > 0 Then
               txtNcuenta.Text = oRsTmp.Fields!idCuentaAtencion
            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
            txtNcuenta_LostFocus
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub



Private Sub cmdBuscaZIP_Click()
    If cmbAlmDestino.Text = "" Then
       MsgBox "Elija el DESTINO", vbInformation, Me.Caption
       cmbAlmDestino.SetFocus
       Exit Sub
    End If
    If cmbConcepto.Text = "" Then
       MsgBox "Elija el TIPO CONCEPTO", vbInformation, Me.Caption
       cmbConcepto.SetFocus
       Exit Sub
    End If
    On Error GoTo ErrBuscarZIP
    Dim oCrypKey As New CrypKey.Util
    Dim oBuscaZIP As New SIGHNegocios.BuscaArchivo
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oConexion As New Connection
    Dim oConexionFox As New Connection
    Dim lc_ArchivoElegido As String, lcRuta As String, lcSql As String, lcTabla As String
    oBuscaZIP.MuestraImagen = False
    oBuscaZIP.TipoArchivo = "*.zip"
    oBuscaZIP.PathDefault = lcBuscaParametro.SeleccionaFilaParametro(236)
    oBuscaZIP.MostrarFormulario
    lc_ArchivoElegido = oBuscaZIP.ArchivoElegido
    If lc_ArchivoElegido <> "" Then
       Me.MousePointer = 11
       lcRuta = "C:\Archivos de programa\Digital Works Corporation\GalenHos\Archivos"    'app.Path & "\archivos"
       SIGHEntidades.DescomprimeArchivoZIP oCrypKey.DecryptString(lcBuscaParametro.SeleccionaFilaParametro(350)), _
                                           lc_ArchivoElegido, lcRuta, False
       lcTabla = Right(lc_ArchivoElegido, 10)
       oConexionFox.CommandTimeout = 300
       oConexionFox.Open "DSN=his"
       lcSql = "select * from mv" & Left(lcTabla, 6)
       oRsTmp1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
       If oRsTmp1.RecordCount = 0 Then
          MsgBox "No existe datos en el archivo: pr" & Left(lcTabla, 6)
       Else
          Set oRsTmp2 = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("")
          If oRsTmp2.RecordCount = 0 Then
                MsgBox "No existe Farmacias/Almacenes"
          Else
                oRsTmp1.MoveFirst
                oRsTmp2.MoveFirst
                oRsTmp2.Find "idAlmacen=" & mo_cmbAlmacenDestino.BoundText
                If Trim(oRsTmp2!CodigoSismed) <> Trim(oRsTmp1!almDstVir) Then
                   MsgBox "El Almacén o Farmacia (DESTINO) no coinciden: " & Trim(oRsTmp1!almDstVir) & Chr(13) & _
                          "Verifique la opción de FARMACIAS", vbInformation, Me.Caption
                Else
                   oRsTmp2.MoveFirst
                   oRsTmp2.Find "codigoSismed='" & Trim(oRsTmp1!almCodiOrg) & "'"
                   If oRsTmp2.EOF Then
                      MsgBox "No se encuentra en la tabla FarmAlmacen el ORIGEN: " & Trim(oRsTmp1!almCodiOrg) & Chr(13) & _
                             "Verifique la opción de DEPENDENCIAS EXTERNAS", vbInformation, Me.Caption
                   Else
                      If Val(mo_cmbAlmacenOrigen.BoundText) <> oRsTmp2.Fields!IdAlmacen Then
                            MsgBox "El ORIGEN elejido no coincide con " & Trim(oRsTmp1!almCodiOrg) & Chr(13) & _
                                   "Verifique la opción de DEPENDENCIAS EXTERNAS", vbInformation, Me.Caption
                      Else
                            txtNdocum.Text = IIf(IsNull(oRsTmp1!movNumeDco), "", Trim(oRsTmp1!movNumeDco))
                            txtNdocO.Text = IIf(IsNull(oRsTmp1!movNumeDci), "", Trim(oRsTmp1!movNumeDci))
                            txtRuc.Text = IIf(IsNull(oRsTmp1!prvNumeRuc), "", Trim(oRsTmp1!prvNumeRuc))
                            txtRuc_LostFocus
                            If txtProveedor.Text = "" Then
                               txtProveedor.Text = IIf(IsNull(oRsTmp1!prvDescrip), "", oRsTmp1!prvDescrip)
                            End If
                            Me.grdProductos.CargaProductosPorTemporal oRsTmp1
                      End If
                   End If
                End If
          End If
          oRsTmp2.Close
       End If
       oRsTmp1.Close
    End If
ErrBuscarZIP:
    If Err.Number <> 0 Then
       MsgBox Err.Description
    End If
    Set oBuscaZIP = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oConexionFox = Nothing
    Set oConexion = Nothing
    Set oCrypKey = Nothing
    Me.MousePointer = 1
End Sub

Private Sub Form_Activate()
   If mo_ReglasFarmacia.LaFarmaciaEstaRegenerandoSaldos(Val(mo_cmbAlmacenDestino.BoundText)) = True Then
        btnCancelar_Click
   End If
End Sub

Private Sub Form_Initialize()

    Set mo_cmbConceptos.MiComboBox = cmbConcepto
    Set mo_cmbAlmacenOrigen.MiComboBox = cmbAlmOrigen
    Set mo_cmbAlmacenDestino.MiComboBox = cmbAlmDestino
    Set mo_cmbTipoDocum.MiComboBox = cmbTipoDocum
    Set mo_cmbTipoDocumO.MiComboBox = cmbTipodocumO
    Set mo_cmbTipoCompra.MiComboBox = cmbTipoCompra
    Set mo_cmbTproceso.MiComboBox = cmbTproceso
     Set mo_cmbUnidosis.MiComboBox = cmbUnidosis
End Sub

Private Sub Form_Load()
       SIGHEntidades.ParaAuditoriaPorCadaDato sghAudLimpiar, ""
       mo_lbElEstablecimentoEsCS = IIf(lcBuscaParametro.SeleccionaFilaParametro(282) = "S", True, False)
       CargarComboBoxes
       ConfigurarGrdProductos
       GenerarRecordsetProductos
        Select Case mi_Opcion
        Case sghAgregar
            Me.Caption = "Agregar Nota Ingreso"
        Case sghModificar
            Me.Caption = "Modificar Nota Ingreso"
        Case sghConsultar
            Me.Caption = "Consultar Nota Ingreso"
            btnImprimir.Visible = True
        Case sghEliminar
            Me.Caption = "Anular Nota Ingreso"
        End Select
        CargarDatosAlFormulario
End Sub


Sub CargarComboBoxes()
    Dim rsIdAlmacen As Recordset
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    Set oBuscaDondeLabora = Nothing
    '
    Set oRsItemsUnidosis = mo_ReglasFarmacia.farmUnidosisSeleccionarTodos
    '
    If mo_lnIdTablaLISTBARITEMS <> 1304 Then
       Set oRsAlmacenDestino = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idEstado=1")
       Label5.Caption = "Farmacia destino"
       cmdBuscaZIP.Visible = IIf(mi_Opcion = sghAgregar, True, False)
    Else
       Set oRsAlmacenDestino = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='A' and idEstado=1")
       Label5.Caption = "Almacén destino"
       Me.fraPacienteForaneo.Visible = False
       Me.frmPaciente.Visible = False
       Me.chkPacienteForaneo.Visible = False
    End If
    '
    mo_cmbAlmacenDestino.BoundColumn = "IdAlmacen"
    mo_cmbAlmacenDestino.ListField = "Descripcion"
    'Set mo_cmbAlmacenDestino.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarTodosMenosExternos
    If mo_lnIdTablaLISTBARITEMS <> 1304 Then
       Set mo_cmbAlmacenDestino.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idEstado=1")
    Else
       Set mo_cmbAlmacenDestino.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='A' and idEstado=1")
    End If
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    lbElUsuarioTrabajaAqui = False
    If rsIdAlmacen.RecordCount > 0 Then
       mo_cmbAlmacenDestino.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmDestino, False
       lbLaFarmaciaEsUnidosis = mo_ReglasFarmacia.FarmaciaEsUnidosis(Val(mo_cmbAlmacenDestino.BoundText))
       lbElUsuarioTrabajaAqui = True
    End If
    'UNIDOSIS
    Dim oRsTmpA As New Recordset
    Set oRsTmpA = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idtipoLocales='F' and idTipoSuministro='01' and idEstado=1 and esUnidosis=1")
    mo_cmbUnidosis.BoundColumn = "IdAlmacen"
    mo_cmbUnidosis.ListField = "Descripcion"
    Set mo_cmbUnidosis.RowSource = oRsTmpA
    If oRsTmpA.RecordCount = 1 Then
       oRsTmpA.MoveFirst
       mo_cmbUnidosis.BoundText = Trim(Str(oRsTmpA!IdAlmacen))
    End If
    Set oRsTmpA = Nothing
    '
    
   '
    mo_cmbTipoDocum.BoundColumn = "idTipoDocumento"
    mo_cmbTipoDocum.ListField = "Nombre"
    Set mo_cmbTipoDocum.RowSource = mo_ReglasFarmacia.FarmTipoDocumentosDevuelveTodos
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
   '
    mo_cmbTipoDocumO.BoundColumn = "idTipoDocumento"
    mo_cmbTipoDocumO.ListField = "Nombre"
    Set mo_cmbTipoDocumO.RowSource = mo_ReglasFarmacia.FarmTipoDocumentosDevuelveTodos
    ms_MensajeError = ms_MensajeError + mo_ReglasFarmacia.MensajeError
    If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError
    End If
End Sub

Sub ConfigurarGrdProductos()
    grdProductos.movNumero = ml_movNumero
    grdProductos.IdAlmacen = 0
    grdProductos.inicializar
    
End Sub

Sub BlanquedaVariablesUnidosis()
    lnCuentaUnidosis = 0
    lbLaFuenteFinanciamientoUsadoEnFUnidosis = False
    cmbUnidosis.Visible = False
    lblUnidosis.Visible = False
End Sub

Sub CargarDatosAlFormulario()
'SCCQ 13/10/2020 Cambio28 Inicio
mo_Formulario.HabilitarDeshabilitar txtNdocum, False
'SCCQ 13/10/2020 Cambio28 Fin
    mo_Formulario.HabilitarDeshabilitar Me.txtNotaIngreso, False
    mo_Formulario.HabilitarDeshabilitar Me.txtFregistro, False
    mo_Formulario.HabilitarDeshabilitar Me.txtHoraRegistro, False
    mo_Formulario.HabilitarDeshabilitar Me.txtEstado, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbTipoDocum, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbTipodocumO, False
    mo_Formulario.HabilitarDeshabilitar Me.txtProveedor, False
    mo_Formulario.HabilitarDeshabilitar Me.txtDatosDeCuenta, False
    mo_Formulario.HabilitarDeshabilitar txtNombrePaciente, False
    wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
    ml_IdPaciente = 0: ml_IdComprobantePago = 0
    ml_IdProveedor = 0
     Select Case mi_Opcion
     Case sghAgregar
        txtFregistro.Text = lcBuscaParametro.RetornaFechaServidorSQL      'Format(Now, sighentidades.DevuelveHoraSoloFormato_HM)
        txtHoraRegistro.Text = lcBuscaParametro.RetornaHoraServidorSQL
        grdProductos.movNumero = ""
        grdProductos.LimpiarGrilla
        grdProductos.CargaProductosPorMovNumero
        grdProductos.AgregaRegistro
     Case sghModificar
        DeshabilitaCabecera
        CargarDatosALosControles
     Case sghConsultar
        DeshabilitaCabecera
        CargarDatosALosControles
        btnAceptar.Enabled = False
     Case sghEliminar
        DeshabilitaCabecera
        CargarDatosALosControles
 End Select

End Sub

Sub DeshabilitaCabecera()
    mo_Formulario.HabilitarDeshabilitar Me.cmbAlmOrigen, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbAlmDestino, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbConcepto, False
End Sub
Sub CargarDatosALosControles()
   Dim oConexion As New Connection
   oConexion.CommandTimeout = 300
   oConexion.CursorLocation = adUseClient
   oConexion.Open SIGHEntidades.CadenaConexion
   
   '**************Datos de la tabla FarmMovimiento *****************
   mo_farmMovimiento.movNumero = ml_movNumero
   mo_farmMovimiento.MovTipo = lcConstanteMovimientoEntrada
   If Not mo_ReglasFarmacia.FarmMovimientoSeleccionarPorId(mo_farmMovimiento) Then
      MsgBox mo_ReglasFarmacia.MensajeError
      Exit Sub
   End If
   
   mo_farmMovimiento1.movNumero = Trim(mo_farmMovimiento.Observaciones)
   lbLaFarmaciaEsUnidosis = mo_ReglasFarmacia.FarmaciaEsUnidosis(mo_farmMovimiento.IdAlmacenDestino)
   
   txtNotaIngreso.Text = ml_movNumero
   mo_cmbAlmacenDestino.BoundText = mo_farmMovimiento.IdAlmacenDestino
   cmbAlmDestino_Click
   
   mo_cmbConceptos.BoundText = mo_farmMovimiento.idTipoConcepto
   cmbConcepto_Click
   mo_cmbAlmacenOrigen.BoundText = mo_farmMovimiento.IdAlmacenOrigen
   mo_cmbTipoDocum.BoundText = mo_farmMovimiento.DocumentoIdtipo
   txtNdocum.Text = mo_farmMovimiento.DocumentoNumero
   txtObservaciones.Text = mo_farmMovimiento.Observaciones
   txtEstado.Text = mo_ReglasFarmacia.DevuelveEstadoActualDelMovimiento("idEstadoMovimiento=" & mo_farmMovimiento.idEstadoMovimiento)
   txtFregistro.Text = Format(mo_farmMovimiento.fechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
   txtHoraRegistro.Text = Format(mo_farmMovimiento.fechaCreacion, SIGHEntidades.DevuelveHoraSoloFormato_HM)
   ml_idUsuarioCreo = mo_farmMovimiento.idUsuario
   '**************Datos de la tabla FarmMovimientoNotaIngreso *****************
   mo_farmMovimientoNotaIngreso.movNumero = ml_movNumero
   mo_farmMovimientoNotaIngreso.MovTipo = lcConstanteMovimientoEntrada
   If Not mo_ReglasFarmacia.FarmMovimientoNotaIngresoSeleccionarPorId(mo_farmMovimientoNotaIngreso) Then
      MsgBox mo_ReglasFarmacia.MensajeError
      Exit Sub
   End If
   If mo_farmMovimientoNotaIngreso.DocumentoFechaRecepcion <> 0 Then
      txtFrecepcion.Text = Format(mo_farmMovimientoNotaIngreso.DocumentoFechaRecepcion, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
   End If
   If mo_farmMovimientoNotaIngreso.OrigenIdTipo > 0 Then
      mo_cmbTipoDocumO.BoundText = mo_farmMovimientoNotaIngreso.OrigenIdTipo
   End If
   txtNdocO.Text = mo_farmMovimientoNotaIngreso.oRigenNumero
   If mo_farmMovimientoNotaIngreso.OrigenFecha <> 0 Then
      txtFdocO.Text = mo_farmMovimientoNotaIngreso.OrigenFecha
   End If
   If mo_farmMovimientoNotaIngreso.idTipoCompra > 0 Then
      mo_cmbTipoCompra.BoundText = mo_farmMovimientoNotaIngreso.idTipoCompra
   End If
   If mo_farmMovimientoNotaIngreso.idTipoProceso > 0 Then
      mo_cmbTproceso.BoundText = mo_farmMovimientoNotaIngreso.idTipoProceso
   End If
   txtNproceso.Text = mo_farmMovimientoNotaIngreso.NumeroProceso
   txtNcuenta.Text = mo_farmMovimientoNotaIngreso.idCuentaAtencion
   txtNcuenta_LostFocus
   
   Me.txtdocext.Text = mo_farmMovimiento.docExterno 'RHA 12/01/2021 CAMBIO 50
   
   'PAQUETES
   If mo_ReglasFarmacia.LaNIoNSesUnARMADO_PAQUETE(mo_farmMovimiento.IdAlmacenDestino, mo_farmMovimiento.idTipoConcepto, _
                                                  mo_farmMovimiento.DocumentoNumero, True) = True Then
      MsgBox "No puede MODIFICAR/ELIMINAR la Nota de Ingreso, debe de usar la opción ARMADO DE PAQUETES", vbInformation, Me.Caption
      btnAceptar.Enabled = False
   End If
   'proveedor
   ml_IdProveedor = mo_farmMovimientoNotaIngreso.IdProveedor
   If ml_IdProveedor > 0 Then
      oDoProveedores.IdProveedor = ml_IdProveedor
      CargaDatosProveedor
   End If
   'paciente
   ml_IdPaciente = mo_farmMovimientoNotaIngreso.IdPaciente
   If ml_IdPaciente > 0 Then
        mo_DoPaciente.IdPaciente = ml_IdPaciente
        Set mo_DoPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(ml_IdPaciente, oConexion)
        txtNhistoria.Text = mo_DoPaciente.NroHistoriaClinica
        txtNombrePaciente.Text = Trim(mo_DoPaciente.ApellidoPaterno) & " " & Trim(mo_DoPaciente.ApellidoMaterno) & " " & mo_DoPaciente.PrimerNombre
   End If
   ml_IdComprobantePago = mo_farmMovimientoNotaIngreso.IdComprobantePago
   '**************Datos de la tabla FarmMovimientoDetalle *****************
   grdProductos.movNumero = ml_movNumero
   grdProductos.CargaProductosPorMovNumero
   grdProductos.RefrescarDatos
   lnTotalDocumento = grdProductos.DevuelveTotal
   If mo_farmMovimiento.idEstadoMovimiento = 0 Or mo_cmbConceptos.BoundText = "19" Then
      'Si estado=Anulado o Concepto=INVENTARIO INICIAL
      btnAceptar.Enabled = False
   End If
   'Es una DEVOLUCION DE BOLETA/FACTURA y ya tiene NOTA DE CREDITO             kike 2017
   If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
        If Left(grdProductos.DocumentoNumero, 1) = "B" Or Left(grdProductos.DocumentoNumero, 1) = "F" Then
           Dim lcNroSerie98 As String, lcNroDocumento98 As String, lnIdComprobantePago98 As Long
           Dim oRsTmp98 As New Recordset
           lcNroSerie98 = Left(grdProductos.DocumentoNumero, InStr(grdProductos.DocumentoNumero, "-") - 1)
           lcNroDocumento98 = Trim(Mid(grdProductos.DocumentoNumero, InStr(grdProductos.DocumentoNumero, "-") + 1, 100))
           Set oRsTmp98 = mo_ReglasCaja.CajaComprobantesPagoSeleccionarPorNroSerieNroDocumento(lcNroSerie98, lcNroDocumento98)
           If oRsTmp98.RecordCount > 0 Then
              If oRsTmp98!IdEstadoComprobante = 4 Then
                 lnIdComprobantePago98 = oRsTmp98!IdComprobantePago
                 oRsTmp98.Close
                 Set oRsTmp98 = mo_ReglasCaja.NotaCreditoBuscaPorIdComprobante(lnIdComprobantePago98)
                 If oRsTmp98.RecordCount > 0 Then
                    MsgBox "No se podrá MODIFICAR/ELIMINAR porque ya tiene NOTA DE CREDITO", vbInformation, ""
                    btnAceptar.Enabled = False
                 End If
              End If
           End If
           oRsTmp98.Close
           Set oRsTmp98 = Nothing
           If mi_Opcion = sghModificar And btnAceptar.Enabled = True Then
               MsgBox "No se podrá MODIFICAR porque es una DEVOLUCION POR BOLETA/FACTURA" & Chr(13) & _
                      "Deberá ANULAR LA NOTA DE INGRESO y volver a AGREGARLA", vbInformation, ""
               btnAceptar.Enabled = False
           End If
        End If
   End If
   '******Modificar documento con Fecha Anterior a la actual,
   '******siempre y cuando no hubieron SALIDAS
   Dim oRsTmp As New ADODB.Recordset
   Set mRs_Productos = grdProductos.DevuelveProductos
   If mRs_Productos.RecordCount > 0 Then
      mRs_Productos.MoveFirst
      Do While Not mRs_Productos.EOF
         Set oRsTmp = mo_ReglasFarmacia.farmMovimientoDetalleDevuelveSalidasSegunAlmacenProductoLote(mo_farmMovimiento.IdAlmacenDestino, mRs_Productos.Fields!idProducto, mRs_Productos.Fields!Lote, mRs_Productos.Fields!FechaVencimiento)
         If oRsTmp.RecordCount > 0 Then
            If oRsTmp.Fields!fechaCreacion >= CDate(txtFregistro.Text & " " & txtHoraRegistro.Text) Then
                MsgBox "No podrá Modificar/Anular una NI porque ya se despachó el producto: " & Chr(13) & Trim(mRs_Productos.Fields!codigo) & " - " & Trim(mRs_Productos.Fields!nombreProducto) & "   NS: " & oRsTmp.Fields!movNumero, vbExclamation, Me.Caption
                btnAceptar.Enabled = False
                Exit Do
            End If
         End If
         mRs_Productos.MoveNext
      Loop
   End If
   'devolucion de paciente, con cuenta.....que anule y vuelva a AGREGAR
   grdProductos.Visible = True
   Me.grdProductosDevol.Visible = False
   grdConsumoPaciente.Visible = False
   If mo_cmbConceptos.BoundText = LcIdTipoConceptoDevolucionPaciente And txtDatosDeCuenta.Text <> "" Then
       Dim oRsTmp1 As New Recordset
       Dim lnCantidadS As Long, lnPrecioS As Double, ldFechaS As Date
       Dim lnFinanciamientoS As Long, lnUsuarioS As Long
       Dim lnIdOrden As Long
       Set oRsTmp = mo_ReglasFarmacia.FacturacionBienesDevolucionesSeleccionarPorMovNumeroE(mo_farmMovimiento.movNumero, mo_farmMovimiento.MovTipo)
       If oRsTmp.RecordCount > 0 Then
          oRsTmp.MoveFirst
          Do While Not oRsTmp.EOF
                If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(ml_IdTipoFinanciamiento, oConexion) = True Then
                   'pagante
                   Set oRsTmp1 = mo_ReglasFacturacion.FacturacionBienesPagosSeleccionarPorMovNumeroProducto(oRsTmp.Fields!movNumero, "S", oRsTmp.Fields!idProducto, oConexion)
                   If oRsTmp1.RecordCount = 0 Then Exit Sub
                   lnCantidadS = oRsTmp1.Fields!CantidadPagar
                   lnPrecioS = oRsTmp1.Fields!PrecioVenta
                   'ldFechaS = Null
                   lnFinanciamientoS = 0
                   lnUsuarioS = 0
                   lnIdOrden = oRsTmp1.Fields!idOrden
                Else
                   'seguros
                   Set oRsTmp1 = mo_ReglasFacturacion.FacturacionBienesFinanciamientosSeleccionaXProdFinanciam(oRsTmp.Fields!movNumero, "S", oRsTmp.Fields!idProducto, ml_IdTipoFinanciamiento)
                   If oRsTmp1.RecordCount = 0 Then Exit Sub
                   lnCantidadS = oRsTmp1.Fields!CantidadFinanciada
                   lnPrecioS = oRsTmp1.Fields!PrecioFinanciado
                   ldFechaS = oRsTmp1.Fields!fechaAutoriza
                   lnFinanciamientoS = oRsTmp1.Fields!idFuenteFinanciamiento
                   lnUsuarioS = oRsTmp1.Fields!IdUsuarioAutoriza
                   lnIdOrden = 0
                End If
          
                mRs_ProductosDevol.AddNew
                mRs_ProductosDevol.Fields!idProducto = oRsTmp.Fields!idProducto
                mRs_ProductosDevol.Fields!Cantidad = oRsTmp.Fields!CantidadAdevolver
                mRs_ProductosDevol.Fields!movNumeroS = oRsTmp.Fields!movNumero
                mRs_ProductosDevol.Fields!idOrdenS = lnIdOrden
                mRs_ProductosDevol.Fields!IdTipoFinanciamientoS = ml_IdTipoFinanciamiento
                mRs_ProductosDevol.Fields!IdFuenteFinanciamientoS = lnFinanciamientoS
                mRs_ProductosDevol.Fields!cantidadS = lnCantidadS
                mRs_ProductosDevol.Fields!PrecioS = lnPrecioS
                mRs_ProductosDevol.Fields!FechaS = ldFechaS
                mRs_ProductosDevol.Fields!UsuarioS = lnUsuarioS
                mRs_ProductosDevol.Update
                oRsTmp.MoveNext
          Loop
       End If
       Set oRsTmp1 = Nothing
       If mi_Opcion = sghModificar Then
          btnAceptar.Enabled = False
       End If
   End If
   Set oRsTmp = Nothing
   oConexion.Close
   Set oConexion = Nothing
   '******permiso a Modificar documento con Fecha Anterior a la actual
   Dim mo_PermisosFacturacion As New PermisosFacturacion
   Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
   Set mo_PermisosFacturacion = mo_ReglasSeguridad.UsuariosRolesSeleccionarPermisosFacturacion(ml_idUsuario)
   If mo_PermisosFacturacion.ActualizaFechaDocumentoES = False And mi_Opcion <> sghConsultar Then
      If CDate(lcBuscaParametro.RetornaFechaServidorSQL) <> CDate(txtFregistro.Text) Then
         MsgBox "No tiene ACCESO a Modificar/Anular una NI" & Chr(13) & " de una Fecha Registro diferente a la actual", vbExclamation, Me.Caption
         btnAceptar.Enabled = False
      End If
   End If
   Set mo_PermisosFacturacion = Nothing
   Set mo_ReglasSeguridad = Nothing
End Sub


Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   If mo_ReglasFarmacia.LaFarmaciaEstaRegenerandoSaldos(Val(mo_cmbAlmacenDestino.BoundText)) = True Then
      btnCancelar_Click
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
   'SCCQ 12/10/2020 Cambio28 Inicio
    'Antes: If ValidarDatosObligatorios() Then
       If ValidarDatosObligatorios("A") Then
   'SCCQ 12/10/2020 Cambio28 Fin
           CargaDatosAlObjetosDeDatos
            If AgregarDatos() Then
               ml_idUsuarioCreo = ml_idUsuario
               'MsgBox "Se agregó correctamente la Nota de Ingreso N° " + txtNotaIngreso.Text, vbExclamation, Me.Caption
                ImprimeDocumento
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo agregar los datos " + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
       End If
   Case sghModificar
   'SCCQ 12/10/2020 Cambio28 Inicio
    'Antes: If ValidarDatosObligatorios() Then
       If ValidarDatosObligatorios("M") Then
   'SCCQ 12/10/2020 Cambio28 Fin
            CargaDatosAlObjetosDeDatos
            If ModificarDatos() Then
               ml_idUsuarioCreo = ml_idUsuario
               'MsgBox "Se Modificó correctamente la Nota de Ingreso N° " + txtNotaIngreso.Text, vbExclamation, Me.Caption
               ImprimeDocumento
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
       End If
   Case sghEliminar
        If MsgBox("Esta seguro de Anular ?", vbQuestion + vbYesNo, "") = vbYes Then
            CargaDatosAlObjetosDeDatos
            If AnularNI() Then
                MsgBox " Se anuló la Nota de Ingreso N° " + txtNotaIngreso.Text, vbInformation, Me.Caption
                Me.Visible = False
                LimpiarVariablesDeMemoria
            Else
                MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
            End If
        End If
   End Select
End Sub
'SCCQ 12/10/2020 Cambio28 Inicio
    'Antes: Function ValidarDatosObligatorios() As Boolean
Function ValidarDatosObligatorios(modo As String) As Boolean
'SCCQ 12/10/2020 Cambio28 Fin
   Dim lbSigue As Boolean
   Dim lnCantProducto As Long
   ValidarDatosObligatorios = False
   ms_MensajeError = ""
   If mo_cmbAlmacenOrigen.BoundText = mo_cmbAlmacenDestino.BoundText Then
       ms_MensajeError = ms_MensajeError + "El Almacén Origen y Destino deben ser DIFERENTES" + Chr(13)
   ElseIf cmbConcepto.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor elija el Concepto" + Chr(13)
       cmbConcepto.SetFocus
   ElseIf cmbAlmOrigen.Text = "" Then
       ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén Origen" + Chr(13)
       cmbAlmOrigen.SetFocus
   ElseIf txtNdocum.Locked = False And txtNdocum.Text = "" Then
   'SCCQ 12/10/2020 Cambio28 Inicio
    If modo = "M" Then 'Modifica
    'SCCQ 12/10/2020 Cambio28 Fin
          ms_MensajeError = ms_MensajeError + "Por favor ingrese el N° Documento" + Chr(13)
          'txtNdocum.SetFocus -->'Se comentó la línea de código
    'SCCQ 12/10/2020 Cambio28 Inicio
    End If
    'SCCQ 12/10/2020 Cambio28 Fin
   ElseIf txtFrecepcion.Enabled = True And txtFrecepcion.Text = SIGHEntidades.FECHA_VACIA_DMY Then
          ms_MensajeError = ms_MensajeError + "Por favor ingrese la Fecha de Recepción" + Chr(13)
          txtFrecepcion.SetFocus
   ElseIf txtNdocO.Locked = False And txtNdocO.Text = "" Then
          ms_MensajeError = ms_MensajeError + "Por favor ingrese el N° Documento Origen" + Chr(13)
          txtNdocO.SetFocus
   ElseIf txtFdocO.Enabled = True And txtFdocO.Text = SIGHEntidades.FECHA_VACIA_DMY Then
          ms_MensajeError = ms_MensajeError + "Por favor ingrese la Fecha del Documento Origen" + Chr(13)
          txtFdocO.SetFocus
   ElseIf cmbTipoCompra.Locked = False And cmbTipoCompra.Text = "" Then
          ms_MensajeError = ms_MensajeError + "Por favor elija el Tipo de Compra" + Chr(13)
          cmbTipoCompra.SetFocus
   ElseIf cmbTproceso.Locked = False And cmbTproceso.Text = "" Then
          ms_MensajeError = ms_MensajeError + "Por favor elija el Tipo de Proceso" + Chr(13)
          cmbTproceso.SetFocus
   ElseIf txtNproceso.Locked = False And txtNproceso.Text = "" Then
          ms_MensajeError = ms_MensajeError + "Por favor ingrese el N° de Proceso" + Chr(13)
          txtNproceso.SetFocus
   ElseIf txtRuc.Locked = False And Trim(txtProveedor.Text) = "" Then
          ms_MensajeError = ms_MensajeError + "Por favor ingrese el N° de RUC/Razón Social" + Chr(13)
          txtRuc.SetFocus
   ElseIf mo_cmbConceptos.BoundText = "21" And Me.txtNcuenta.Text = "" Then
            If Me.chkPacienteForaneo.Value = 0 Then
                ms_MensajeError = ms_MensajeError + "Por favor ingrese el N° de Cuenta del Paciente que devuelve" + Chr(13)
                Me.txtNcuenta.SetFocus
            End If
   End If
   If SIGHEntidades.EsFecha(txtFrecepcion, "DD/MM/AAAA") Then
      If CDate(txtFrecepcion.Text) > CDate(txtFregistro.Text) Then
         ms_MensajeError = ms_MensajeError + "La Fecha de Recepción no puede ser mayor a la Fecha de Registro" + Chr(13)
         txtFrecepcion.SetFocus
      End If
   End If
   If SIGHEntidades.EsFecha(txtFdocO.Text, "DD/MM/AAAA") Then
      If CDate(txtFdocO.Text) > CDate(txtFregistro.Text) Then
         ms_MensajeError = ms_MensajeError + "La Fecha de Doc.Origen no puede ser mayor a la Fecha de Registro" + Chr(13)
         txtFdocO.SetFocus
      End If
   End If
   If mi_Opcion = sghAgregar And txtNdocum.Text <> "" Then
      Dim oRsTmp As New ADODB.Recordset
      Set oRsTmp = mo_ReglasFarmacia.farmMovimientoSeleccionarPorTipoYnumeroDocumento(txtNdocum.Text, Val(mo_cmbTipoDocum.BoundText))
      oRsTmp.Filter = "idEstadoMovimiento=1"
      If oRsTmp.RecordCount > 0 Then
         ms_MensajeError = ms_MensajeError + "El Número de Documento: " & txtNdocum.Text & " EXISTE en NI: " & Trim(oRsTmp.Fields!movNumero) & "     Fecha: " & oRsTmp.Fields!fechaCreacion & Chr(13)
      End If
      oRsTmp.Close
      Set oRsTmp = Nothing
   End If
   If mo_cmbConceptos.BoundText = LcIdTipoConceptoDevolucionPaciente Then    'se eligio=Devolucion del Paciente
      If ml_idFuenteFinanciamiento = 0 Then
         ml_idFuenteFinanciamiento = 1   'CONTADO =es un Paciente sin Nro Cuenta
         ml_IdTipoFinanciamiento = 1
      End If
   End If
     
   If mo_cmbConceptos.BoundText = LcIdTipoConceptoDevolucionPaciente And txtDatosDeCuenta.Text <> "" Then
        'devolucion de paciente, con cuenta
        Set mRs_Productos = mRs_ProductosDevol.Clone()
   Else
        If txtDatosDeCuenta.Text = "" And Me.chkPacienteForaneo.Value = 1 Then
             'devolución con Boleta      'kike 2017
             Set mRs_Productos = mRs_ProductosDevol.Clone()
             lnTotalDocumento = 0
             If mRs_Productos.RecordCount > 0 Then
                mRs_Productos.MoveFirst
                Do While Not mRs_Productos.EOF
                    mRs_Productos!Total = Round(mRs_Productos!Cantidad * mRs_Productos!Precio, 2)
                    mRs_Productos.Update
                    lnTotalDocumento = lnTotalDocumento + mRs_Productos!Total
                    mRs_Productos.MoveNext
                Loop
             End If
            
            
        Else
            lnTotalDocumento = grdProductos.DevuelveTotal
            Set mRs_Productos = grdProductos.DevuelveProductos
        End If
   End If
   
   If mRs_Productos.RecordCount = 0 Then
       ms_MensajeError = ms_MensajeError + "Por favor Ingrese Productos" + Chr(13)
   Else
        '*******devolucion de paciente, con Cuenta
        If mo_cmbConceptos.BoundText = LcIdTipoConceptoDevolucionPaciente And txtDatosDeCuenta.Text <> "" And mRs_ProductosDevol.RecordCount > 0 Then
              Dim lnIdProd As Long, lcLot As String, ldFV As Date, lcProd As String
              mRs_ProductosDevol.Sort = "idProducto,Lote,FechaVencimiento"
              mRs_ProductosDevol.MoveFirst
              lnTotalDocumento = 0
              Do While Not mRs_ProductosDevol.EOF
                 lnCantProducto = 0
                 lnIdProd = mRs_ProductosDevol.Fields!idProducto
                 lcLot = mRs_ProductosDevol.Fields!Lote
                 ldFV = mRs_ProductosDevol.Fields!FechaVencimiento
                 lcProd = mRs_ProductosDevol.Fields!nombreProducto
                 Do While Not mRs_ProductosDevol.EOF And lnIdProd = mRs_ProductosDevol.Fields!idProducto And lcLot = mRs_ProductosDevol.Fields!Lote And ldFV = mRs_ProductosDevol.Fields!FechaVencimiento
                    mRs_ProductosDevol!Total = Round(mRs_ProductosDevol.Fields!Cantidad * mRs_ProductosDevol!Precio, 2)
                    lnTotalDocumento = lnTotalDocumento + mRs_ProductosDevol!Total
                    mRs_ProductosDevol.MoveNext
                    lnCantProducto = lnCantProducto + 1
                    If mRs_ProductosDevol.EOF Then
                       Exit Do
                    End If
                 Loop
                 If lnCantProducto > 1 Then
                    ms_MensajeError = ms_MensajeError + "El producto " + lcProd + "  (tiene LOTE y FECHA VENCIMIENTO repetidas)" & Chr(13)
                    Exit Do
                 End If
              Loop
        End If
        '
        Dim LdFechaMinimaDespacho As Date
        If mo_cmbConceptos.BoundText = "5" Then
           LdFechaMinimaDespacho = Date - 300  'Devolucion por Vencimiento
        Else
           LdFechaMinimaDespacho = CDate(txtFregistro.Text) + Val(lcBuscaParametro.SeleccionaFilaParametro(224))
        End If
        mRs_Productos.MoveFirst
        Do While Not mRs_Productos.EOF
           If Trim(mRs_Productos.Fields!codigo) = "" Or Trim(mRs_Productos.Fields!nombreProducto) = "" Then
                mRs_Productos.Delete
                mRs_Productos.Update
           Else
                If mRs_Productos.Fields!Cantidad <= 0 Then
                   ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  (Tiene problemas con la CANTIDAD)" + Chr(13)
                End If
                If Trim(mRs_Productos!Lote) = "" Then
                   ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  (Tiene problemas en el LOTE)" + Chr(13)
                ElseIf mRs_Productos!esPaquete = True And Trim(mRs_Productos!Lote) <> WxLOTEpaquete Then
                   ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  (es un PAQUETE el LOTE debe llamarse " & WxLOTEpaquete & ")" + Chr(13)
                End If
                If mRs_Productos!FechaVencimiento <= LdFechaMinimaDespacho Then
                   ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  (Tiene problemas con la FECHA DE VENCIMIENTO, debe ser mayor a " & LdFechaMinimaDespacho & ")" & Chr(13)
                ElseIf mRs_Productos!esPaquete = True And mRs_Productos!FechaVencimiento <> CDate(WxFVENCIMIENTOpaquete) Then
                   ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  (es un PAQUETE la FECHA DE VENCIMIENTO debe ser " & WxFVENCIMIENTOpaquete & ")" & Chr(13)
                End If
                If mRs_Productos!Precio <= 0 Then
                   ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  (Tiene problemas con el Precio)" + Chr(13)
                End If
                If IsNull(mRs_Productos!registroSanitario) Then
                   ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  (Debe ingresar el REGISTRO SANITARIO)" + Chr(13)
                ElseIf mRs_Productos!esPaquete = True And Trim(mRs_Productos!registroSanitario) <> WxREGSANITARIOpaquete Then
                   ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  (es un PAQUETE el REGISTRO SANITARIO debe llamarse " & WxREGSANITARIOpaquete & ")" + Chr(13)
                End If
                '*******devolucion de paciente, con Cuenta
                If mo_cmbConceptos.BoundText = LcIdTipoConceptoDevolucionPaciente And txtDatosDeCuenta.Text <> "" Then
                   If mRs_ConsumoPaciente.RecordCount > 0 Then
                      lnCantProducto = 0
                      lbSigue = True
                      mRs_ConsumoPaciente.MoveFirst
                      Do While Not mRs_ConsumoPaciente.EOF
                         If Trim(mRs_Productos!codigo) = Trim(mRs_ConsumoPaciente.Fields!codigo) And Val(mo_cmbAlmacenDestino.BoundText) = mRs_ConsumoPaciente.Fields!IdAlmacenOrigen And mRs_Productos!movNumeroS = mRs_ConsumoPaciente.Fields!movNumeroS Then
                            lbSigue = False
                            lnCantProducto = lnCantProducto + mRs_ConsumoPaciente!Cantidad
                         End If
                         mRs_ConsumoPaciente.MoveNext
                      Loop
                      If lbSigue = True Then
                         ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  (NO TIENE NINGUN DESPACHO en el ALMACEN:  " & cmbAlmDestino.Text & ") & Chr(13)"
                      ElseIf lnCantProducto < mRs_Productos.Fields!Cantidad Then
                         ms_MensajeError = ms_MensajeError + "El producto " + Trim(mRs_Productos.Fields!codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  (No debe devolver una cantidad mayor a: " & lnCantProducto & ")" & Chr(13)
                      End If
                   Else
                      ms_MensajeError = ms_MensajeError + "No existe ningún DESPACHO" + Chr(13)
                   End If
                End If
           End If
           mRs_Productos.MoveNext
        Loop
   End If
   
   'Es un despacho hacia la FARMACIA UNIDOSIS
   ms_MensajeError = ms_MensajeError & mo_ReglasFarmacia.DevuelveSiSonItemsDeUNIDOSIS(lbLaFarmaciaEsUnidosis, _
                                                         mRs_Productos, oRsItemsUnidosis)
   'Es devolucion de Cta UNIDOSIS hacia FARMACIA XXX, debe elegir la FARMACIA ORIGEN UNIDOSIS
   If lnCuentaUnidosis = Val(Me.txtNcuenta.Text) And mo_cmbConceptos.BoundText = LcIdTipoConceptoDevolucionPaciente And _
                         ms_MensajeError = "" And cmbUnidosis.Visible = True Then
      If cmbUnidosis.Text = "" Then
         ms_MensajeError = ms_MensajeError + "Debe elegir la F.UNIDOSIS (origen)" + Chr(13)
      End If
      '
      Set oRsItemsUNIDOSISpNS = DevuelveItemsDeFarmaciaUNIDOSISconLotes(ms_MensajeError)
   End If
   '
   If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function


Sub CargaDatosAlObjetosDeDatos()
    Select Case mi_Opcion
    Case sghAgregar
        With mo_farmMovimiento
            .DocumentoIdtipo = Val(mo_cmbTipoDocum.BoundText)                      '10
            .DocumentoNumero = txtNdocum.Text                                      '2014-08
            .fechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL          'igual
            .IdAlmacenDestino = Val(mo_cmbAlmacenDestino.BoundText)                '8
            .IdAlmacenOrigen = Val(mo_cmbAlmacenOrigen.BoundText)                  '0
            .idEstadoMovimiento = sghEstadoTabla.sghRegistrado                     'igual
            .idTipoConcepto = Val(mo_cmbConceptos.BoundText)                       '20
            .idUsuario = ml_idUsuario                                              'igual
            .IdUsuarioAuditoria = ml_idUsuario                                     'igual
            .MovTipo = lcConstanteMovimientoEntrada                                'igual
            .Observaciones = txtObservaciones.Text                                 'vacio
            .Total = lnTotalDocumento                                              'sumar
            
            .docExterno = txtdocext.Text 'RHA 12/01/2021 CAMBIO 50
            
        End With
        With mo_farmMovimientoNotaIngreso
            .DocumentoFechaRecepcion = IIf(txtFrecepcion.Text = SIGHEntidades.FECHA_VACIA_DMY, 0, txtFrecepcion.Text)  'hoy
            .IdPaciente = ml_IdPaciente                                                             '0
            .IdComprobantePago = ml_IdComprobantePago                                               '0
            .IdProveedor = ml_IdProveedor                                                           '0
            .idTipoCompra = Val(mo_cmbTipoCompra.BoundText)                                         '1
            .idTipoProceso = Val(mo_cmbTproceso.BoundText)                                          '1
            .IdUsuarioAuditoria = ml_idUsuario                                                      'igual
            .MovTipo = lcConstanteMovimientoEntrada                                                 'igual
            .NumeroProceso = txtNproceso.Text                                                       'vacio
            .OrigenFecha = IIf(txtFdocO.Text = SIGHEntidades.FECHA_VACIA_DMY, 0, txtFdocO.Text)     'vacio
            .OrigenIdTipo = Val(mo_cmbTipoDocumO.BoundText)                                         '22
            .oRigenNumero = txtNdocO.Text                                                           'vacio
            .idCuentaAtencion = Val(txtNcuenta.Text)                                                'vacio
            .idFuenteFinanciamiento = ml_idFuenteFinanciamiento                                     '0
        End With
        If txtProveedor.Locked = False Then   'nuevo proveedor
            If Trim(txtProveedor.Text) <> "" And ml_IdProveedor = 0 Then
                With oDoProveedores
                    .IdProveedor = ml_IdProveedor
                    .razonSocial = txtProveedor.Text
                    .ruc = txtRuc.Text
                    .IdUsuarioAuditoria = ml_idUsuario
                End With
            End If
        End If
   Case sghModificar
        With mo_farmMovimiento
            .DocumentoNumero = txtNdocum.Text
            .Observaciones = txtObservaciones.Text
            .IdUsuarioAuditoria = ml_idUsuario
            .Total = lnTotalDocumento
            '.FechaCreacion = txtFregistro.Text
            .docExterno = txtdocext.Text 'RHA 12/01/2021 CAMBIO 50
            
        End With
        With mo_farmMovimientoNotaIngreso
            .DocumentoFechaRecepcion = IIf(txtFrecepcion.Text = SIGHEntidades.FECHA_VACIA_DMY, 0, txtFrecepcion.Text)
            .IdPaciente = ml_IdPaciente
            .IdComprobantePago = ml_IdComprobantePago
            .IdProveedor = ml_IdProveedor
            .idTipoCompra = Val(mo_cmbTipoCompra.BoundText)
            .idTipoProceso = Val(mo_cmbTproceso.BoundText)
            .IdUsuarioAuditoria = ml_idUsuario
            .MovTipo = lcConstanteMovimientoEntrada
            .NumeroProceso = txtNproceso.Text
            .OrigenFecha = IIf(txtFdocO.Text = SIGHEntidades.FECHA_VACIA_DMY, 0, txtFdocO.Text)
            .OrigenIdTipo = Val(mo_cmbTipoDocumO.BoundText)
            .oRigenNumero = txtNdocO.Text
            .FechaModificacion = lcBuscaParametro.RetornaFechaServidorSQL
            .idUsuarioModifica = ml_idUsuario
            .idCuentaAtencion = Val(txtNcuenta.Text)
            .idFuenteFinanciamiento = ml_idFuenteFinanciamiento
        End With
        If txtProveedor.Locked = False Then   'nuevo proveedor
            If Trim(txtProveedor.Text) <> "" And ml_IdProveedor = 0 Then
                With oDoProveedores
                    .IdProveedor = ml_IdProveedor
                    .razonSocial = txtProveedor.Text
                    .ruc = txtRuc.Text
                    .IdUsuarioAuditoria = ml_idUsuario
                End With
            End If
        End If
   Case sghEliminar
        With mo_farmMovimiento
            .fechaAnulacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
            .idEstadoMovimiento = sghEstadoTabla.sghAnulado    'Anulado
            .IdUsuarioAuditoria = ml_idUsuario
        End With
   End Select
End Sub

Function DevuelveItemsDeFarmaciaUNIDOSISconLotes(ByRef lcMensaje As String) As Recordset
    Dim mRs_Productos As New Recordset
    With mRs_Productos
          .Fields.Append "IdProducto", adInteger, 4
          .Fields.Append "Codigo", adVarChar, 20
          .Fields.Append "NombreProducto", adChar, 300
          .Fields.Append "idTipoSalidaBienInsumo", adInteger
          .Fields.Append "Lote", adVarChar, 15
          .Fields.Append "FechaVencimiento", adDate, , adFldIsNullable
          .Fields.Append "Saldo", adInteger
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Precio", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "RegistroSanitario", adVarChar, 50, adFldIsNullable
          .Fields.Append "NumeroDocumento", adVarChar, 20, adFldIsNullable
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    If mi_Opcion <> sghAgregar Then
       Set DevuelveItemsDeFarmaciaUNIDOSISconLotes = mRs_Productos
       Exit Function
    End If
    Dim oRsConceptos1 As New Recordset
    Dim rs As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim oFarmMovimientoDetalle As New farmMovimientoDetalle
    Dim lcCodigoConPunto As String, lnCantidad1 As Long, lnPrecioUnitario As Double
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 900
    oConexion.Open SIGHEntidades.CadenaConexion
    Set oFarmMovimientoDetalle.Conexion = oConexion
    Set oRsConceptos1 = mo_ReglasFarmacia.FarmTipoConceptosDevuelveParaRegistroDeNiNs(oRsAlmacenDestino.Fields!idTipoLocales, lcConstanteMovimientoEntrada, oRsAlmacenDestino.Fields!idTipoSuministro)
    oRsConceptos1.MoveFirst
    oRsConceptos1.Find "idTipoConcepto=" & lnTipoConceptoAjusteInventario
    mRs_ProductosDevol.MoveFirst
    Do While Not mRs_ProductosDevol.EOF
       lcCodigoConPunto = Trim(mRs_ProductosDevol!codigo) & SIGHEntidades.Pto
       oRsItemsUnidosis.MoveFirst
       oRsItemsUnidosis.Find "codigo='" & Trim(mRs_ProductosDevol!codigo) & "'"
       If oRsItemsUnidosis.EOF Then
          lcMensaje = lcMensaje & "El Código: " & lcCodigoConPunto & " no existe en FARMACIA UNIDOSIS" & Chr(13)
       Else
          lnCantidad1 = mRs_ProductosDevol!Cantidad * Val(oRsItemsUnidosis!convertir)
          Set rs = oFarmMovimientoDetalle.FarmDevuelveSaldosConLotesSegunAlmacenCliente(Val(mo_cmbUnidosis.BoundText), 0, _
                                                                                        lcCodigoConPunto)
          If rs.RecordCount > 0 Then
             rs.MoveFirst
             Do While Not rs.EOF
                lnPrecioUnitario = mo_ReglasFarmacia.DevuelvePrecioSegunTipoConcepto(rs!idProducto, _
                                                                                     oRsConceptos1!TipoPrecioParaNiNs)
                mRs_Productos.AddNew
                mRs_Productos!idProducto = rs!idProducto
                mRs_Productos!codigo = lcCodigoConPunto
                mRs_Productos!nombreProducto = rs!Nombre
                mRs_Productos!idTipoSalidaBienInsumo = rs!idTipoSalidaBienInsumoSaldo
                mRs_Productos!Lote = rs!Lote
                mRs_Productos!FechaVencimiento = rs!FechaVencimiento
                mRs_Productos!saldo = rs!saldo
                If lnCantidad1 <= rs!saldo Then
                   mRs_Productos!Cantidad = lnCantidad1
                   mRs_Productos!Total = Round(lnCantidad1 * lnPrecioUnitario, 2)
                   lnCantidad1 = 0
                Else
                   mRs_Productos!Cantidad = rs!saldo
                   mRs_Productos!Total = Round(rs!saldo * lnPrecioUnitario, 2)
                   lnCantidad1 = lnCantidad1 - rs!saldo
                End If
                mRs_Productos!Precio = lnPrecioUnitario
                'mRs_Productos!RegistroSanitario = rs
                'mRs_Productos!NumeroDocumento = rs
                mRs_Productos.Update
                If lnCantidad1 = 0 Then
                   Exit Do
                End If
                rs.MoveNext
             Loop
             If lnCantidad1 > 0 Then
                lcMensaje = lcMensaje & "El Código: " & lcCodigoConPunto & " no tiene SALDO SUFICIENTE en FARMACIA UNIDOSIS" & Chr(13)
             End If
          End If
       End If
       mRs_ProductosDevol.MoveNext
    Loop
    Set DevuelveItemsDeFarmaciaUNIDOSISconLotes = mRs_Productos
    oConexion.Close
    Set oConexion = Nothing
    Set mRs_Productos = Nothing
    Set rs = Nothing
    Set oFarmMovimientoDetalle = Nothing
    Set oRsConceptos1 = Nothing
End Function

Sub CreaNSaFarmaciaUNIDOSIS()
    If lnCuentaUnidosis > 0 And lnCuentaUnidosis = Val(Me.txtNcuenta.Text) And mo_cmbConceptos.BoundText = LcIdTipoConceptoDevolucionPaciente Then
        Dim lbActualizaDatos As Boolean
        Dim mo_farmMovimiento2 As New farmMovimiento
        Dim oConexion As New Connection
        Dim lnTotal1 As Double
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        If mi_Opcion = sghAgregar Then
           lnTotal1 = 0
           oRsItemsUNIDOSISpNS.MoveFirst
           Do While Not oRsItemsUNIDOSISpNS.EOF
              lnTotal1 = lnTotal1 + oRsItemsUNIDOSISpNS!Total
              oRsItemsUNIDOSISpNS.MoveNext
           Loop
           With mo_farmMovimiento1
                '.movNumero
                .MovTipo = lcConstanteMovimientoSalida
                .IdAlmacenOrigen = Val(mo_cmbUnidosis.BoundText)
                .idTipoConcepto = lnTipoConceptoAjusteInventario
                .DocumentoIdtipo = 10
                .DocumentoNumero = Format(Now, SIGHEntidades.DevuelveFechaSoloFormato_DMYHMS)
                .Total = lnTotal1
                .fechaCreacion = mo_farmMovimiento.fechaCreacion
                .idUsuario = mo_farmMovimiento.IdUsuarioAuditoria
                .idEstadoMovimiento = sghEstadoTabla.sghRegistrado
                .IdUsuarioAuditoria = mo_farmMovimiento.IdUsuarioAuditoria
           End With
           lbActualizaDatos = mo_ReglasFarmacia.AgregaDatosDeNotaSalida(mo_farmMovimiento1, oRsItemsUNIDOSISpNS, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
           mo_farmMovimiento.Observaciones = mo_farmMovimiento1.movNumero
           Set mo_farmMovimiento2.Conexion = oConexion
           If mo_farmMovimiento2.Modificar(mo_farmMovimiento) Then
              MsgBox "Se creó Nota de Salida en FARMACIA UNIDOSIS en forma automática", vbInformation, Me.Caption
           End If
        ElseIf mi_Opcion = sghEliminar Then
           Set mo_farmMovimiento2.Conexion = oConexion
           mo_farmMovimiento1.MovTipo = lcConstanteMovimientoSalida
           If mo_farmMovimiento2.SeleccionarPorId(mo_farmMovimiento1) Then
              With mo_farmMovimiento1
                    .fechaAnulacion = mo_farmMovimiento.fechaCreacion
                    .idEstadoMovimiento = sghEstadoTabla.sghAnulado    'Anulado
                    .IdUsuarioAuditoria = mo_farmMovimiento.IdUsuarioAuditoria
              End With
              lbActualizaDatos = mo_ReglasFarmacia.AnulaNotaSalida(mo_farmMovimiento1, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, 0, 0)
              If lbActualizaDatos = True Then
                 MsgBox "Se Anuló Nota de Salida en FARMACIA UNIDOSIS en forma automática", vbInformation, Me.Caption
              End If
           End If
        End If
        oConexion.Close
        Set mo_farmMovimiento2 = Nothing
        Set oConexion = Nothing
    End If
End Sub

Function AgregarDatos() As Boolean
    'SCCQ 19/10/2020 Cambio28 Inicio
    oRsConceptos.MoveFirst
    oRsConceptos.Find "idTipoConcepto=" & mo_cmbConceptos.BoundText
    If oRsConceptos.Fields!DocumentoEsAutomatico = "S" Then 'Verificamos si el número de documento se genera de forma AUTOMATICA
        AgregarDatos = mo_ReglasFarmacia.AgregaDatosDeNotaIngreso_NumDocAutomatico(oRsAlmacenDestino.Fields!idTipoLocales, oRsAlmacenDestino.Fields!idTipoSuministro, CLng(mo_cmbTipoDocum.BoundText), mo_farmMovimiento, mo_farmMovimientoNotaIngreso, oDoProveedores, mRs_Productos, ml_IdTipoFinanciamiento, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
        txtNdocum.Text = mo_farmMovimiento.DocumentoNumero
    Else
    'SCCQ 19/10/2020 Cambio28 Fin
        AgregarDatos = mo_ReglasFarmacia.AgregaDatosDeNotaIngreso(mo_farmMovimiento, mo_farmMovimientoNotaIngreso, oDoProveedores, mRs_Productos, ml_IdTipoFinanciamiento, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    'SCCQ 19/10/2020 Cambio28 Inicio
    End If
    'SCCQ 19/10/2020 Cambio28 Fin
    txtNotaIngreso.Text = mo_farmMovimiento.movNumero
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
    If Val(Me.txtNcuenta.Text) > 0 Then
       mo_ReglasSISgalenhos.FuaActualizaDespachosEnFarmacia Val(Me.txtNcuenta.Text), wxParametro302, lnIdTipoServicio, ml_idFuenteFinanciamiento
    End If
    
    CreaNSaFarmaciaUNIDOSIS

End Function
Function ModificarDatos() As Boolean
    ModificarDatos = mo_ReglasFarmacia.ModificaDatosDeNotaIngreso(mo_farmMovimiento, mo_farmMovimientoNotaIngreso, oDoProveedores, mRs_Productos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
    If Val(Me.txtNcuenta.Text) > 0 Then
       mo_ReglasSISgalenhos.FuaActualizaDespachosEnFarmacia Val(Me.txtNcuenta.Text), wxParametro302, lnIdTipoServicio, ml_idFuenteFinanciamiento
    End If
End Function
Function AnularNI() As Boolean
    AnularNI = mo_ReglasFarmacia.AnulaNotaIngreso(mo_farmMovimiento, mo_farmMovimientoNotaIngreso, ml_IdTipoFinanciamiento, mRs_ProductosDevol, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc)
    ms_MensajeError = mo_ReglasFarmacia.MensajeError
    If Val(Me.txtNcuenta.Text) > 0 Then
       mo_ReglasSISgalenhos.FuaActualizaDespachosEnFarmacia Val(Me.txtNcuenta.Text), wxParametro302, lnIdTipoServicio, ml_idFuenteFinanciamiento
    End If
    CreaNSaFarmaciaUNIDOSIS
End Function

Private Sub ImprimeDocumento()
'    Dim mo_Imprime As New NotaIngreso
'    mo_Imprime.CrearReporte_excel cmbAlmOrigen.Text, cmbAlmDestino.Text, cmbTipoDocum.Text, txtNdocum.Text, txtFregistro.Text, txtNotaIngreso.Text, lnTotalDocumento
    Dim oDOfarmAlmacen As New DoFarmAlmacen
    Dim oRptClase As New rCrystal
    Dim mo_AdminComun As New ReglasComunes
    Set oDOfarmAlmacen = mo_ReglasFarmacia.FarmAlmacenSeleccionarPorId(Val(mo_cmbAlmacenDestino.BoundText))
    oRptClase.MovTipo = "E"
    oRptClase.Documento = txtNotaIngreso.Text
    oRptClase.TextoDelFiltro = "NOTA DE INGRESO"
    oRptClase.Almacen = "(" & oDOfarmAlmacen.CodigoSismed & ")" & cmbAlmDestino.Text
    oRptClase.AlmacenO = cmbAlmOrigen.Text
    oRptClase.HoraInicio = txtFregistro.Text
    oRptClase.HoraFin = Trim(cmbTipoDocum.Text) & " - " & txtNdocum.Text
    oRptClase.Importe = lnTotalDocumento
    oRptClase.TipoReporte = "NiNs"
    
     'RHA 15/01/2021 CAMBIO 52 INICIO
    If Len(Trim(Me.txtNombrePaciente.Text)) = 0 Then
    oRptClase.Paciente = " "
    Else
    oRptClase.Paciente = "Paciente: " & Me.txtNombrePaciente.Text
    End If
    'RHA 15/01/2021 CAMBIO 52 FIN
        '
    '
    oRptClase.Observaciones = Trim(Me.txtObservaciones.Text)
    oRptClase.Observaciones = Trim(Me.txtObservaciones.Text) & "  (" & Label1.Caption & ":  " & cmbConcepto.Text & ")"    'debb-07/10/2016

    
    If Trim(cmbTipodocumO.Text) <> "" Then
        oRptClase.Proveedor = Label11.Caption & ": " & Trim(cmbTipodocumO.Text) & "/" & Trim(txtNdocO.Text) & " (" & _
                              Label17.Caption & ": " & Trim(txtProveedor.Text) & ")"

                              
    End If
    oRptClase.idUsuario = ml_idUsuarioCreo
    oRptClase.Show vbModal
    Set oRptClase = Nothing
    Set oDOfarmAlmacen = Nothing
    Set mo_AdminComun = Nothing
End Sub

Private Sub btnImprimir_Click()
   ImprimeDocumento
End Sub









Private Sub Form_Unload(Cancel As Integer)
   If SIGHEntidades.ParaAuditoria = "" Then
      LimpiarVariablesDeMemoria
   ElseIf MsgBox("Hubo cambios, desea salir de todas maneras ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
      LimpiarVariablesDeMemoria
      SIGHEntidades.ParaAuditoriaPorCadaDato sghAudLimpiar, ""
   End If
End Sub

Private Sub grdConsumoPaciente_DblClick()
    If Not mRs_ConsumoPaciente.BOF And Not mRs_ConsumoPaciente.EOF Then
        Dim lbEntrar As Boolean
        lbEntrar = False
        
        If cmbUnidosis.Visible = True And lnCuentaUnidosis = Val(txtNcuenta.Text) Then
            If cmbUnidosis.Text = "" Then
                MsgBox "Debe elegir la FARMACIA UNIDOSIS (origen) antes de devolver cada ITEM", vbInformation, ""
                Exit Sub
            End If
        End If
        
        
        If mRs_ProductosDevol.RecordCount > 0 Then
           mRs_ProductosDevol.MoveFirst
           Do While Not mRs_ProductosDevol.EOF
              If mRs_ProductosDevol.Fields!idProducto = mRs_ConsumoPaciente.Fields!idProducto And mRs_ProductosDevol.Fields!movNumeroS = mRs_ConsumoPaciente.Fields!movNumeroS Then
                 MsgBox "Ese Producto/NumeroDocumento ya se eligió", vbInformation, Me.Caption
                 Exit Sub
              End If
              mRs_ProductosDevol.MoveNext
           Loop
           lbEntrar = True
        Else
           lbEntrar = True
        End If
        
        If lbEntrar = True Then
            Dim oRsTmp1 As New Recordset
            Dim oRsTmp2 As New Recordset
            Dim lnCantidadS As Long, lnPrecioS As Double, ldFechaS As Date
            Dim lnFinanciamientoS As Long, lnUsuarioS As Long
            Dim oConexion As New Connection
            Dim lcLote As String, ldFechaVencimiento As Date, lcReg_sanit As String
            oConexion.Open SIGHEntidades.CadenaConexion
            oConexion.CursorLocation = adUseClient
            If mRs_ConsumoPaciente.Fields!idOrdenS > 0 Then
               'pagante
               Set oRsTmp1 = mo_ReglasFacturacion.FacturacionBienesPagosSeleccionarPorMovNumeroProducto(mRs_ConsumoPaciente.Fields!movNumeroS, "S", mRs_ConsumoPaciente.Fields!idProducto, oConexion)
               If oRsTmp1.RecordCount = 0 Then Exit Sub
               lnCantidadS = oRsTmp1.Fields!CantidadPagar
               lnPrecioS = oRsTmp1.Fields!PrecioVenta
               'ldFechaS = Null
               lnFinanciamientoS = 0
               lnUsuarioS = 0
            Else
               'seguros
               Set oRsTmp1 = mo_ReglasFacturacion.FacturacionBienesFinanciamientosSeleccionaXProdFinanciam(mRs_ConsumoPaciente.Fields!movNumeroS, "S", mRs_ConsumoPaciente.Fields!idProducto, mRs_ConsumoPaciente.Fields!IdTipoFinanciamientoS)
               If oRsTmp1.RecordCount = 0 Then Exit Sub
               lnCantidadS = oRsTmp1.Fields!CantidadFinanciada
               lnPrecioS = oRsTmp1.Fields!PrecioFinanciado
               ldFechaS = oRsTmp1.Fields!fechaAutoriza
               lnFinanciamientoS = oRsTmp1.Fields!idFuenteFinanciamiento
               lnUsuarioS = oRsTmp1.Fields!IdUsuarioAutoriza
            End If
            'debb-05-03-2012
            lcLote = ""
            ldFechaVencimiento = Date
            Set oRsTmp2 = mo_ReglasFarmacia.farmMovimientoDetalleSeleccionarSalidasXitem(mRs_ConsumoPaciente.Fields!movNumeroS, mRs_ConsumoPaciente.Fields!idProducto)
            If oRsTmp2.RecordCount > 0 Then
                lcLote = oRsTmp2.Fields!Lote
                ldFechaVencimiento = oRsTmp2.Fields!FechaVencimiento
            End If
            oRsTmp2.Close
            'hra
            Set oRsTmp2 = mo_ReglasFarmacia.ExportaPreciosSismedRegSant(mRs_ConsumoPaciente.Fields!idProducto, oConexion)
            lcReg_sanit = ""
            If oRsTmp2.RecordCount > 0 Then
               lcReg_sanit = Left(oRsTmp2!registroSanitario, 50)
            End If
            oRsTmp2.Close
            '
            mRs_ProductosDevol.AddNew
            mRs_ProductosDevol.Fields!NumeroDocumento = mRs_ConsumoPaciente.Fields!NumeroDocumento
            mRs_ProductosDevol.Fields!idProducto = mRs_ConsumoPaciente.Fields!idProducto
            mRs_ProductosDevol.Fields!codigo = mRs_ConsumoPaciente.Fields!codigo
            mRs_ProductosDevol.Fields!nombreProducto = mRs_ConsumoPaciente.Fields!nombreProducto
            mRs_ProductosDevol.Fields!Lote = lcLote
            mRs_ProductosDevol.Fields!FechaVencimiento = ldFechaVencimiento
            mRs_ProductosDevol.Fields!Cantidad = mRs_ConsumoPaciente.Fields!Cantidad
            mRs_ProductosDevol.Fields!Precio = mRs_ConsumoPaciente.Fields!Precio
            mRs_ProductosDevol.Fields!Total = mRs_ConsumoPaciente.Fields!Total
            mRs_ProductosDevol.Fields!movNumeroS = mRs_ConsumoPaciente.Fields!movNumeroS
            mRs_ProductosDevol.Fields!idOrdenS = mRs_ConsumoPaciente.Fields!idOrdenS
            mRs_ProductosDevol.Fields!IdTipoFinanciamientoS = mRs_ConsumoPaciente.Fields!IdTipoFinanciamientoS
            mRs_ProductosDevol.Fields!IdFuenteFinanciamientoS = lnFinanciamientoS
            mRs_ProductosDevol.Fields!cantidadS = lnCantidadS
            mRs_ProductosDevol.Fields!PrecioS = lnPrecioS
            mRs_ProductosDevol.Fields!FechaS = ldFechaS
            mRs_ProductosDevol.Fields!UsuarioS = lnUsuarioS
            mRs_ProductosDevol.Fields!idTipoSalidaBienInsumo = sghTipoSalidaItemFarmacia.sghSoloVenta
            mRs_ProductosDevol.Fields!registroSanitario = lcReg_sanit
            mRs_ProductosDevol.Update
            Set oRsTmp1 = Nothing
            oConexion.Close
            Set oConexion = Nothing
        End If
    End If
End Sub

Private Sub grdConsumoPaciente_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     
     grdConsumoPaciente.Bands(0).Columns("idProducto").Hidden = True
     grdConsumoPaciente.Bands(0).Columns("FechaDocumento").Width = 1000
     grdConsumoPaciente.Bands(0).Columns("NumeroDocumento").Width = 1000
     grdConsumoPaciente.Bands(0).Columns("Precio").Hidden = True
     grdConsumoPaciente.Bands(0).Columns("Total").Hidden = True
     grdConsumoPaciente.Bands(0).Columns("IdAlmacenOrigen").Hidden = True
     grdConsumoPaciente.Bands(0).Columns("codigo").Width = 700
     grdConsumoPaciente.Bands(0).Columns("codigo").Activation = ssActivationActivateNoEdit
     grdConsumoPaciente.Bands(0).Columns("NombreProducto").Width = 5000
     grdConsumoPaciente.Bands(0).Columns("NombreProducto").Activation = ssActivationActivateNoEdit
     grdConsumoPaciente.Bands(0).Columns("cantidad").Width = 800
     grdConsumoPaciente.Bands(0).Columns("cantidad").Activation = ssActivationActivateNoEdit
     grdConsumoPaciente.Bands(0).Columns("cantidad").Format = "###0"
     grdConsumoPaciente.Bands(0).Columns("FarmaciaDespacho").Header.Caption = "Almacén"
     grdConsumoPaciente.Bands(0).Columns("FarmaciaDespacho").Activation = ssActivationActivateNoEdit
     grdConsumoPaciente.Bands(0).Columns("FarmaciaDespacho").Width = 5000

End Sub

Private Sub grdConsumoPaciente_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdConsumoPaciente_DblClick
    End If
End Sub

Private Sub grdProductosDevol_AfterRowsDeleted()
    Set grdProductosDevol.DataSource = mRs_ProductosDevol
End Sub

Private Sub grdProductosDevol_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     grdProductosDevol.Bands(0).Columns("idProducto").Hidden = True
     grdProductosDevol.Bands(0).Columns("Precio").Hidden = True
     grdProductosDevol.Bands(0).Columns("Total").Hidden = True
     grdProductosDevol.Bands(0).Columns("movNumeroS").Hidden = True
     grdProductosDevol.Bands(0).Columns("idOrdenS").Hidden = True
     grdProductosDevol.Bands(0).Columns("idTipoFinanciamientoS").Hidden = True
     grdProductosDevol.Bands(0).Columns("idFuenteFinanciamientoS").Hidden = True
     grdProductosDevol.Bands(0).Columns("CantidadS").Hidden = True
     grdProductosDevol.Bands(0).Columns("PrecioS").Hidden = True
     grdProductosDevol.Bands(0).Columns("FechaS").Hidden = True
     grdProductosDevol.Bands(0).Columns("UsuarioS").Hidden = True
     grdProductosDevol.Bands(0).Columns("NumeroDocumento").Activation = ssActivationActivateNoEdit
     grdProductosDevol.Bands(0).Columns("codigo").Width = 700
     grdProductosDevol.Bands(0).Columns("codigo").Activation = ssActivationActivateNoEdit
     grdProductosDevol.Bands(0).Columns("NombreProducto").Width = 8200
     grdProductosDevol.Bands(0).Columns("NombreProducto").Activation = ssActivationActivateNoEdit
     grdProductosDevol.Bands(0).Columns("cantidad").Width = 800
     grdProductosDevol.Bands(0).Columns("cantidad").Format = "###0"
     grdProductosDevol.Bands(0).Columns("idTipoSalidaBienInsumo").Hidden = True
End Sub

Private Sub txtFdocO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdocO

End Sub

Private Sub txtFdocO_LostFocus()
    If txtFdocO <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFdocO, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFdocO = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If

End Sub

Private Sub txtFrecepcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFrecepcion

End Sub

Private Sub txtFrecepcion_LostFocus()
    If txtFrecepcion <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFrecepcion, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFrecepcion = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If

End Sub



Private Sub txtNcuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    ElseIf Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
    End If

End Sub

Private Sub txtNcuenta_LostFocus()
    If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
       Dim oRsTmp As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oRsTmp9 As New Recordset
       Dim lbSigue As Boolean
       Dim lnIdOrden As Long
       Dim lcNIunidosisMovNumero As String, lnNIudidosisIdFarmacia As Long
       Dim oConexion As New Connection
       Dim lbContinuar As Boolean
       oConexion.Open SIGHEntidades.CadenaConexion
       oConexion.CursorLocation = adUseClient
       Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
       txtDatosDeCuenta.Text = ""
       ml_IdPaciente = 0
       ml_idFuenteFinanciamiento = 0
       ml_IdTipoFinanciamiento = 0
       txtNombrePaciente.Text = ""
       txtNhistoria.Text = ""
       lbSigue = True
       lblUnidosis.Visible = False
       BlanquedaVariablesUnidosis
       If oRsTmp.RecordCount > 0 Then
          If oRsTmp.Fields!idEstado <> 1 Then
             If mi_Opcion <> sghConsultar Then
                MsgBox "Ese estado de Cuenta no se encuentra ABIERTA", vbInformation, Me.Caption
                If mi_Opcion = sghModificar Or mi_Opcion = sghEliminar Then
                   btnAceptar.Enabled = False
                Else
                   lbSigue = False
                End If
             End If
          End If
          If lbSigue Then
                'unidosis
                lbLaFuenteFinanciamientoUsadoEnFUnidosis = mo_ReglasComunes.FuenteFinanciamientoEsUnidosis(oRsTmp!idFuenteFinanciamiento, lnCuentaUnidosis)
                If lbSigue = True And mi_Opcion = sghAgregar And lbLaFuenteFinanciamientoUsadoEnFUnidosis = True Then
                   If lbSigue = True And lbLaFarmaciaEsUnidosis = False And lnCuentaUnidosis = Val(txtNcuenta.Text) Then
                      cmbUnidosis.Visible = True
                      lblUnidosis.Visible = True
                   End If
                End If
                '
                lnIdTipoServicio = oRsTmp.Fields!IdTipoServicio
                txtDatosDeCuenta.Text = "IAFA Act: " & Trim(oRsTmp.Fields!dFuenteFinanciamiento) & "  F.Ing: " & oRsTmp.Fields!fechaingreso & "- " & IIf(oRsTmp.Fields!IdTipoServicio = 1, "Consultorios Externos", IIf(oRsTmp.Fields!IdTipoServicio = 3, "Hospitalización", "Emergencia")) & "- (Est: " & Trim(oRsTmp.Fields!estadoCta) & ")"
                ml_IdPaciente = oRsTmp.Fields!IdPaciente
                ml_idFuenteFinanciamiento = oRsTmp.Fields!idFuenteFinanciamiento
                ml_IdTipoFinanciamiento = oRsTmp.Fields!IdFormaPago
                txtNombrePaciente.Text = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!PrimerNombre
                txtNhistoria.Text = oRsTmp.Fields!NroHistoriaClinica
                'carga consumos del Paciente
                If mo_cmbConceptos.BoundText = LcIdTipoConceptoDevolucionPaciente Then
                    Dim oRsTmp1 As New Recordset
                    Dim oRsTmp3 As New Recordset 'RHA 07/10/20 Cambio33
                    Dim lnCantidadConsumoQueda As Long
                    If mi_Opcion = sghAgregar Then
                        grdConsumoPaciente.Visible = True
                        grdProductosDevol.Visible = True
                    End If
                    Set rsTmp = mo_ReglasFarmacia.FarmMovimientoVentasDetalleSeleccionarPorCuenta(Val(txtNcuenta.Text), oConexion)
                    rsTmp.Filter = "idEstadoMovimiento=1"
                    If mRs_ConsumoPaciente.RecordCount > 0 Then
                        mRs_ConsumoPaciente.MoveFirst
                        While Not mRs_ConsumoPaciente.EOF
                            mRs_ConsumoPaciente.Delete
                            mRs_ConsumoPaciente.MoveNext
                        Wend
                    End If
                    If rsTmp.RecordCount > 0 Then
                       rsTmp.MoveFirst
                       Do While Not rsTmp.EOF
                          lbContinuar = True
                          Set oRsTmp1 = mo_ReglasFacturacion.FacturacionBienesDevolucionesSeleccionarPorIdProducto(rsTmp.Fields!movNumero, "S", rsTmp.Fields!idProducto, oConexion)
                          lnCantidadConsumoQueda = rsTmp.Fields!Cantidad
                          'RHA 07/10/20 Cambio33 Inicio
                          Set oRsTmp3 = mo_ReglasFacturacion1.usp_FacturacionBienesDevolucionesSeleccionarPorIdProducto(rsTmp.Fields!movNumero, "S", rsTmp.Fields!idProducto, oConexion)
                          If oRsTmp1.RecordCount > 0 Then
                            lnCantidadConsumoQueda = rsTmp.Fields!Cantidad - oRsTmp3.Fields!cantdevolver 'Resta la cantidad con la sumatoria total de devoluciones
                          End If
                          'RHA 07/10/20 Cambio33 Fin
                           
                          If lnCantidadConsumoQueda > 0 Then
                             lnIdOrden = 0
                             If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(ml_IdTipoFinanciamiento, oConexion) = True Then
                                Set oRsTmp2 = mo_ReglasFacturacion.FacturacionBienesPagosSeleccionarPorMovNumeroProducto(rsTmp.Fields!movNumero, "S", rsTmp.Fields!idProducto, oConexion)
                                If oRsTmp2.RecordCount > 0 Then
                                   lnIdOrden = oRsTmp2.Fields!idOrden
                                   If oRsTmp2.Fields!idEstadoFacturacion = 4 Then   'Si ya tiene Boleta no se debe devolver
                                      lbContinuar = False
                                   End If
                                End If
                                oRsTmp2.Close
                             End If
                             If lbContinuar = True Then
                                '
                                lcNIunidosisMovNumero = ""
                                lnNIudidosisIdFarmacia = 0
                                If cmbUnidosis.Visible = True And lnCuentaUnidosis = Val(txtNcuenta.Text) Then
                                   Set oRsTmp9 = mo_ReglasFarmacia.farmMovimientoSeleccionarPorMovNumero(Trim(rsTmp!Observaciones), lcConstanteMovimientoEntrada)
                                   If oRsTmp9.RecordCount > 0 Then
                                      lcNIunidosisMovNumero = oRsTmp9!movNumero
                                      lnNIudidosisIdFarmacia = oRsTmp9!IdAlmacenDestino
                                   End If
                                   oRsTmp9.Close
                                End If
                                '
                                mRs_ConsumoPaciente.AddNew
                                mRs_ConsumoPaciente.Fields!IdAlmacenOrigen = rsTmp.Fields!IdAlmacenOrigen
                                mRs_ConsumoPaciente.Fields!idProducto = rsTmp.Fields!idProducto
                                mRs_ConsumoPaciente.Fields!NumeroDocumento = rsTmp.Fields!DocumentoNumero
                                mRs_ConsumoPaciente.Fields!FechaDocumento = rsTmp.Fields!fechaCreacion
                                mRs_ConsumoPaciente.Fields!codigo = rsTmp.Fields!codigo
                                mRs_ConsumoPaciente.Fields!nombreProducto = rsTmp.Fields!Nombre
                                mRs_ConsumoPaciente.Fields!Cantidad = lnCantidadConsumoQueda
                                mRs_ConsumoPaciente.Fields!FarmaciaDespacho = rsTmp.Fields!dAlmacen
                                mRs_ConsumoPaciente.Fields!Precio = rsTmp.Fields!Precio
                                mRs_ConsumoPaciente.Fields!Total = Round(lnCantidadConsumoQueda * rsTmp.Fields!Precio, 2)
                                mRs_ConsumoPaciente.Fields!movNumeroS = rsTmp.Fields!movNumero
                                mRs_ConsumoPaciente.Fields!IdTipoFinanciamientoS = ml_IdTipoFinanciamiento
                                mRs_ConsumoPaciente.Fields!idOrdenS = lnIdOrden
                                mRs_ConsumoPaciente.Fields!NIunidosisMovNumero = lcNIunidosisMovNumero
                                mRs_ConsumoPaciente.Fields!NIudidosisIdFarmacia = lnNIudidosisIdFarmacia
                                mRs_ConsumoPaciente.Update
                             End If
                          End If
                          oRsTmp1.Close
                          rsTmp.MoveNext
                       Loop
                       FiltraItemsDeLaFarmaciaUnidosisElejida

                    End If
                    Set grdConsumoPaciente.DataSource = mRs_ConsumoPaciente
                    grdConsumoPaciente.Caption = "Consumos de la Cuenta N° " & txtNcuenta.Text
                End If
          End If
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
       oConexion.Close
       Set oConexion = Nothing
       Set oRsTmp9 = Nothing
    End If

End Sub

Sub FiltraItemsDeLaFarmaciaUnidosisElejida()
    On Error Resume Next
    If mo_cmbUnidosis.BoundText <> "" And cmbUnidosis.Visible = True And lnCuentaUnidosis = Val(txtNcuenta.Text) Then
       mRs_ConsumoPaciente.Filter = "NIudidosisIdFarmacia=" & mo_cmbUnidosis.BoundText
    End If
    If mRs_ConsumoPaciente.RecordCount > 0 Then
       mRs_ConsumoPaciente.MoveFirst
    End If
End Sub

Private Sub txtNdocO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNdocO

End Sub

Private Sub txtNdocum_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNdocum

End Sub

Private Sub txtNhistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNhistoria

End Sub

Private Sub btnBuscarPaciente_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New DOPaciente
    Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            ml_IdPaciente = oDOPaciente.IdPaciente
            txtNhistoria.Text = oDOPaciente.NroHistoriaClinica
            txtNombrePaciente.Text = Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + oDOPaciente.PrimerNombre
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub











Private Sub txtNhistoria_LostFocus()
      If txtNhistoria.Text <> "" Then
        ml_idFuenteFinanciamiento = 0
        ml_IdTipoFinanciamiento = 0
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oDOPaciente As New sighComun.DOPaciente
        oDOPaciente.NroHistoriaClinica = txtNhistoria.Text
        Set oRsTmp1 = mo_AdminAdmision.PacientesFiltrar(oDOPaciente, False, False, "")
        If oRsTmp1.RecordCount > 0 Then
           ml_IdPaciente = oRsTmp1.Fields!IdPaciente
           txtNombrePaciente.Text = Trim(oRsTmp1.Fields!ApellidoPaterno) & " " & Trim(oRsTmp1.Fields!ApellidoMaterno) & " " & oRsTmp1.Fields!PrimerNombre
        Else
           ml_IdPaciente = 0
           txtNombrePaciente.Text = ""
        End If
        Set oRsTmp1 = Nothing
        Set oDOPaciente = Nothing
      End If
End Sub

Private Sub txtNproceso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNproceso

End Sub

Private Sub txtNumeroBoleta_KeyDown(KeyCode As Integer, Shift As Integer)
      mo_Teclado.RealizarNavegacion KeyCode, txtNumeroBoleta
End Sub

Private Sub txtNumeroBoleta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    ElseIf Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
    End If
End Sub

Private Sub txtObservaciones_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNproceso
End Sub





Private Sub txtRuc_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtRuc

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
           grdConsumoPaciente.Visible = False
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub



Private Sub txtRuc_LostFocus()
   If txtRuc.Text <> "" Then
       mo_Formulario.HabilitarDeshabilitar Me.txtProveedor, False
       ml_IdProveedor = 0
       If Len(txtRuc.Text) <> 11 Then
          MsgBox "El Número de RUC debe tener 11 dígitos", vbInformation, Me.Caption
       Else
          Dim oRsTmp As New ADODB.Recordset
          Set oRsTmp = mo_ReglasFacturacion.ProveedoresSeleccionarPorRUC(txtRuc.Text)
          If oRsTmp.RecordCount > 0 Then
             txtProveedor.Text = oRsTmp.Fields!razonSocial
             ml_IdProveedor = oRsTmp.Fields!IdProveedor
          Else
             mo_Formulario.HabilitarDeshabilitar Me.txtProveedor, True
             txtProveedor.SetFocus
          End If
          oRsTmp.Close
          Set oRsTmp = Nothing
       End If
   End If
End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta

End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Formulario = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set oRsConceptos = Nothing
    Set oRsAlmacenDestino = Nothing
    Set rsTmp = Nothing
    Set mo_cmbConceptos = Nothing
    Set mo_cmbAlmacenOrigen = Nothing
    Set mo_cmbAlmacenDestino = Nothing
    Set mo_cmbTipoDocum = Nothing
    Set mo_cmbTipoDocumO = Nothing
    Set mo_cmbTipoCompra = Nothing
    Set mo_cmbTproceso = Nothing
    Set lcBuscaParametro = Nothing
    Set mRs_Productos = Nothing
    Set mo_farmMovimiento = Nothing
    Set mo_farmMovimientoNotaIngreso = Nothing
    Set oDoProveedores = Nothing
    Set mo_DoPaciente = Nothing
    Set mo_ReglasSeguridad = Nothing
    Set mo_AdminAdmision = Nothing
    Set mRs_ProductosDevol = Nothing
    Set mRs_ConsumoPaciente = Nothing
End Sub


Sub GenerarRecordsetProductos()
    With mRs_ProductosDevol
          .Fields.Append "IdProducto", adInteger, 4
          .Fields.Append "NumeroDocumento", adChar, 20
          .Fields.Append "Codigo", adChar, 20
          .Fields.Append "NombreProducto", adChar, 300
          .Fields.Append "Lote", adChar, 15, adFldIsNullable
          .Fields.Append "FechaVencimiento", adDate, , adFldIsNullable
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Precio", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "MovNumeroS", adChar, 9
          .Fields.Append "idOrdenS", adInteger, 4               'de tabla:FacturacionBienesPagos
          .Fields.Append "idTipoFinanciamientoS", adInteger, 4  'de tabla:FacturacionBienesFinanciamientos
          .Fields.Append "idFuenteFinanciamientoS", adInteger, 4  'de tabla:FacturacionBienesFinanciamientos
          .Fields.Append "CantidadS", adInteger                 'de tabla:FacturacionBienesFinanciamientos,FacturacionBienesPagos
          .Fields.Append "PrecioS", adDouble                    'de tabla:FacturacionBienesFinanciamientos,FacturacionBienesPagos
          .Fields.Append "FechaS", adDate, , adFldIsNullable    'de tabla:FacturacionBienesFinanciamientos
          .Fields.Append "UsuarioS", adInteger, 4
          .Fields.Append "IdTipoSalidaBienInsumo", adInteger
          .Fields.Append "RegistroSanitario", adChar, 50, adFldIsNullable
          .Fields.Append "esPaquete", adBoolean
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdProductosDevol.DataSource = mRs_ProductosDevol
    gridInfra.ConfigurarFilasBiColores grdProductosDevol, SIGHEntidades.GrillaConFilasBicolor
    '
    With mRs_ConsumoPaciente
          .Fields.Append "IdProducto", adInteger, 4
          .Fields.Append "NumeroDocumento", adChar, 20
          .Fields.Append "FechaDocumento", adDate, , adFldIsNullable
          .Fields.Append "Codigo", adChar, 15
          .Fields.Append "NombreProducto", adChar, 300
          .Fields.Append "Cantidad", adInteger          'sino cancelo
          .Fields.Append "FarmaciaDespacho", adChar, 60
          .Fields.Append "Precio", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "MovNumeroS", adChar, 9        'sino cancelo
          .Fields.Append "IdAlmacenOrigen", adInteger, 4
          .Fields.Append "idOrdenS", adInteger, 4               'de tabla:FacturacionBienesPagos
          .Fields.Append "idTipoFinanciamientoS", adInteger, 4  'de tabla:FacturacionBienesFinanciamientos
          .Fields.Append "NIunidosisMovNumero", adChar, 9, adFldIsNullable   'Nota de Ingreso en farmacia UNIDOSIS
          .Fields.Append "NIudidosisIdFarmacia", adInteger                   'Farmacia Unidosis donde se despachó
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
End Sub


Private Sub txtSerieBoleta_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, txtSerieBoleta
End Sub

Private Sub txtSerieBoleta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'    ElseIf Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'           KeyAscii = 0
'    End If
End Sub

Private Sub txtSerieBoleta_LostFocus()
    If txtSerieBoleta.Text = "" Then
        MsgBox "Ingrese la Serie de la boleta", vbInformation, Me.Caption
        Exit Sub
    End If
    If txtNumeroBoleta.Text <> "" Then   'kike 2017
       BuscarDetalleBoleta
    End If
End Sub

Private Sub txtNumeroBoleta_LostFocus()
    If txtNumeroBoleta.Text = "" Then
        MsgBox "Ingrese el número de la boleta", vbInformation, Me.Caption
        Exit Sub
    End If
    BuscarDetalleBoleta
End Sub

Public Sub BuscarDetalleBoleta()
    If txtSerieBoleta.Text = "" Then
        'MsgBox "Ingrese la Serie de la boleta", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If txtNumeroBoleta.Text = "" Then
        'MsgBox "Ingrese el número de la boleta", vbInformation, Me.Caption
        Exit Sub
    End If
    'kike 2017
    Dim oRsTmp As New Recordset
    If mi_Opcion = sghAgregar Then
        Set oRsTmp = mo_ReglasCaja.NotaCreditoFarmNotaIngreso(txtSerieBoleta.Text & "-" & txtNumeroBoleta.Text)
        If oRsTmp.RecordCount > 0 Then
           MsgBox "Ya tiene registrado una NOTA DE INGRESO: " & oRsTmp!movNumero, vbInformation, ""
           Set oRsTmp = Nothing
           Exit Sub
        End If
    End If
    '
    Dim oRsTmp2 As New Recordset
    Dim lnIdOrden As Long
    Dim oConexion As New Connection
    Dim lbContinuar As Boolean
    Dim lcBoleta As String
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    txtDatosDeCuenta.Text = ""
    ml_IdPaciente = 0
    ml_idFuenteFinanciamiento = 0
    ml_IdTipoFinanciamiento = 0
    txtNombrePaciente.Text = ""
    txtNhistoria.Text = ""
    grdConsumoPaciente.Visible = False
    lcBoleta = Trim(txtSerieBoleta.Text) & "-" & Trim(txtNumeroBoleta.Text)
                 
    If mo_cmbConceptos.BoundText = LcIdTipoConceptoDevolucionPaciente Then
        Dim oRsTmp1 As New Recordset
        Dim lnCantidadConsumoQueda As Long
        
        ml_idFuenteFinanciamiento = 1
        ml_IdTipoFinanciamiento = 1
        
        If mi_Opcion = sghAgregar Then
            grdConsumoPaciente.Visible = True
            grdProductosDevol.Visible = True
        End If
        
        Set rsTmp = mo_ReglasFarmacia.FarmMovimientoVentasDetalleSeleccionarPorNBoleta(lcBoleta, oConexion)
        rsTmp.Filter = "idEstadoMovimiento=1"
        If mRs_ConsumoPaciente.RecordCount > 0 Then
            mRs_ConsumoPaciente.MoveFirst
            While Not mRs_ConsumoPaciente.EOF
                mRs_ConsumoPaciente.Delete
                mRs_ConsumoPaciente.MoveNext
            Wend
        End If
        If rsTmp.RecordCount > 0 Then
           rsTmp.MoveFirst
           Do While Not rsTmp.EOF
              lbContinuar = True
              Set oRsTmp1 = mo_ReglasFacturacion.FacturacionBienesDevolucionesSeleccionarPorIdProducto(rsTmp.Fields!movNumero, "S", rsTmp.Fields!idProducto, oConexion)
              lnCantidadConsumoQueda = rsTmp.Fields!Cantidad
              If oRsTmp1.RecordCount > 0 Then
                 lnCantidadConsumoQueda = rsTmp.Fields!Cantidad - oRsTmp1.Fields!CantidadAdevolver
              End If
              If lnCantidadConsumoQueda > 0 Then
                 lnIdOrden = 0
                 If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(ml_IdTipoFinanciamiento, oConexion) = True Then
                    Set oRsTmp2 = mo_ReglasFacturacion.FacturacionBienesPagosSeleccionarPorMovNumeroProducto(rsTmp.Fields!movNumero, "S", rsTmp.Fields!idProducto, oConexion)
                    If oRsTmp2.RecordCount > 0 Then
                       lnIdOrden = oRsTmp2.Fields!idOrden
                       If oRsTmp2.Fields!idEstadoFacturacion = 4 Then   'Si ya tiene Boleta no se debe devolver
'                                      lbContinuar = False
                       End If
                    End If
                    oRsTmp2.Close
                 End If
                 If lbContinuar = True Then
                    mRs_ConsumoPaciente.AddNew
                    mRs_ConsumoPaciente.Fields!IdAlmacenOrigen = rsTmp.Fields!IdAlmacenOrigen
                    mRs_ConsumoPaciente.Fields!idProducto = rsTmp.Fields!idProducto
                    mRs_ConsumoPaciente.Fields!NumeroDocumento = rsTmp.Fields!DocumentoNumero
                    mRs_ConsumoPaciente.Fields!FechaDocumento = rsTmp.Fields!fechaCreacion
                    mRs_ConsumoPaciente.Fields!codigo = rsTmp.Fields!codigo
                    mRs_ConsumoPaciente.Fields!nombreProducto = rsTmp.Fields!Nombre
                    mRs_ConsumoPaciente.Fields!Cantidad = lnCantidadConsumoQueda
                    mRs_ConsumoPaciente.Fields!FarmaciaDespacho = rsTmp.Fields!dAlmacen
                    mRs_ConsumoPaciente.Fields!Precio = rsTmp.Fields!Precio
                    mRs_ConsumoPaciente.Fields!Total = Round(lnCantidadConsumoQueda * rsTmp.Fields!Precio, 2)
                    mRs_ConsumoPaciente.Fields!movNumeroS = rsTmp.Fields!movNumero
                    mRs_ConsumoPaciente.Fields!IdTipoFinanciamientoS = ml_IdTipoFinanciamiento
                    mRs_ConsumoPaciente.Fields!idOrdenS = lnIdOrden
                    mRs_ConsumoPaciente.Update
                 End If
              End If
              oRsTmp1.Close
              rsTmp.MoveNext
           Loop
           If mRs_ConsumoPaciente.RecordCount > 0 Then
              mRs_ConsumoPaciente.MoveFirst
           End If
        Else
           MsgBox "No se encontró la BOLETA en tabla: FarmMovimiento", vbInformation, ""
        End If
        
        Set grdConsumoPaciente.DataSource = mRs_ConsumoPaciente
        grdConsumoPaciente.Caption = "Consumos del Paciente Foráneo " & txtNcuenta.Text
    End If

'    oRsTmp.Close
'    Set oRsTmp = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub

