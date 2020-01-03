VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form ReembolsosDetalle 
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "ReembolsosDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraReembolso 
      Caption         =   "Reembolso"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7185
      Left            =   30
      TabIndex        =   17
      Top             =   1530
      Width           =   15015
      Begin VB.Frame Frame3 
         Caption         =   "Datos del REEMBOLSO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   5715
         TabIndex        =   49
         Top             =   4050
         Width           =   9195
         Begin VB.CheckBox chkIGV 
            Caption         =   "Genera IGV"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   45
            TabIndex        =   78
            Top             =   2190
            Width           =   1275
         End
         Begin VB.CheckBox chkGrabaDefinitivamente 
            Caption         =   "Graba Definitivamente (si marca no se podrá Modificar)"
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
            Height          =   315
            Left            =   45
            TabIndex        =   65
            Top             =   2700
            Width           =   5445
         End
         Begin VB.TextBox txtMotivoAnulacion 
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
            Left            =   2340
            MaxLength       =   150
            TabIndex        =   63
            Top             =   1035
            Width           =   3240
         End
         Begin VB.TextBox txtSaldoInicial 
            Alignment       =   1  'Right Justify
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
            Left            =   6945
            TabIndex        =   54
            Top             =   225
            Width           =   1335
         End
         Begin VB.TextBox txtDescripcionR 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   2340
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   52
            Text            =   "ReembolsosDetalle.frx":0CCA
            Top             =   1410
            Width           =   6675
         End
         Begin VB.ComboBox cmbMes 
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
            ItemData        =   "ReembolsosDetalle.frx":0CCC
            Left            =   3450
            List            =   "ReembolsosDetalle.frx":0CCE
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   255
            Width           =   2115
         End
         Begin VB.ComboBox cmbAnio 
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
            ItemData        =   "ReembolsosDetalle.frx":0CD0
            Left            =   2340
            List            =   "ReembolsosDetalle.frx":0CD2
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   255
            Width           =   1035
         End
         Begin VB.TextBox txtSaldoFinal 
            Alignment       =   1  'Right Justify
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
            Left            =   6945
            TabIndex        =   51
            Top             =   585
            Width           =   1335
         End
         Begin VB.CommandButton btnRefrescar 
            Caption         =   "Refrescar Importes"
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
            Left            =   6150
            TabIndex        =   50
            Top             =   975
            Width           =   2175
         End
         Begin VB.TextBox txtDctoHospital 
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
            Left            =   2340
            MaxLength       =   100
            TabIndex        =   57
            Top             =   645
            Width           =   3240
         End
         Begin VB.Label lblDescripcionLarga 
            Caption         =   $"ReembolsosDetalle.frx":0CD4
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   2355
            TabIndex        =   69
            Top             =   2130
            Visible         =   0   'False
            Width           =   6795
         End
         Begin VB.Label lblMotivoAnulacion 
            Caption         =   "Motivo de Anulación"
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
            Left            =   90
            TabIndex        =   64
            Top             =   1065
            Width           =   2175
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Saldo Inicial"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5910
            TabIndex        =   61
            Top             =   285
            Width           =   975
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Saldo Final"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5925
            TabIndex        =   60
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Descripción del Reembolso"
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
            Left            =   90
            TabIndex        =   59
            Top             =   1410
            Width           =   2175
         End
         Begin VB.Label Label13 
            Caption         =   "Periodo (Año, Mes)"
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
            Left            =   90
            TabIndex        =   58
            Top             =   285
            Width           =   2145
         End
         Begin VB.Label Label1 
            Caption         =   "Documentos del Hospital"
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
            Left            =   90
            TabIndex        =   56
            Top             =   675
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdEliminaLista 
         Caption         =   "Limpia Lista"
         DisabledPicture =   "ReembolsosDetalle.frx":0D5C
         DownPicture     =   "ReembolsosDetalle.frx":10E7
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   375
         Picture         =   "ReembolsosDetalle.frx":147A
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Elimina todos las Cuentas de la Lista"
         Top             =   3615
         Width           =   2580
      End
      Begin VB.Frame Frame2 
         Caption         =   "Genera comprobante en CAJA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   105
         TabIndex        =   36
         Top             =   4050
         Width           =   5595
         Begin VB.TextBox txtNroSerie 
            Alignment       =   1  'Right Justify
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
            Left            =   3420
            MaxLength       =   4
            TabIndex        =   72
            Top             =   975
            Width           =   585
         End
         Begin VB.TextBox txtNroDocumento 
            Alignment       =   1  'Right Justify
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
            Left            =   4170
            MaxLength       =   8
            TabIndex        =   71
            Top             =   975
            Width           =   1380
         End
         Begin VB.TextBox txtEmailProv 
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
            MaxLength       =   100
            TabIndex        =   47
            Top             =   2310
            Width           =   4230
         End
         Begin VB.TextBox txtDireccionProv 
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
            MaxLength       =   100
            TabIndex        =   48
            Top             =   2670
            Width           =   4230
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
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   46
            Top             =   1950
            Width           =   4230
         End
         Begin VB.ComboBox cmbIdTurno 
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
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   285
            Width           =   4230
         End
         Begin VB.ComboBox cmbIdCaja 
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
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   630
            Width           =   4230
         End
         Begin VB.ComboBox cmbIdTipoComprobante 
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
            Left            =   1320
            TabIndex        =   44
            Text            =   "cmbIdTipoComprobante"
            Top             =   945
            Width           =   2025
         End
         Begin VB.TextBox txtRuc 
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
            Left            =   1320
            MaxLength       =   11
            TabIndex        =   45
            Top             =   1620
            Width           =   4230
         End
         Begin MSMask.MaskEdBox txtFemision 
            Height          =   315
            Left            =   1320
            TabIndex        =   76
            Top             =   1275
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/#### ##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFcobro 
            Height          =   315
            Left            =   3915
            TabIndex        =   77
            Top             =   1290
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/#### ##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Fec.Cobro"
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
            Left            =   3075
            TabIndex        =   75
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "F.Emisión"
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
            TabIndex        =   74
            Top             =   1335
            Width           =   750
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4050
            TabIndex        =   73
            Top             =   1020
            Width           =   75
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Email Prov"
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
            TabIndex        =   68
            Top             =   2370
            Width           =   825
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Dirección Prov"
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
            TabIndex        =   67
            Top             =   2685
            Width           =   1155
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Razón social"
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
            TabIndex        =   62
            Top             =   2010
            Width           =   960
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Turno"
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
            TabIndex        =   41
            Top             =   330
            Width           =   495
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Caja"
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
            TabIndex        =   40
            Top             =   675
            Width           =   330
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Comprobante"
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
            TabIndex        =   38
            Top             =   1020
            Width           =   1110
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "RUC"
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
            TabIndex        =   37
            Top             =   1680
            Width           =   330
         End
      End
      Begin VB.Frame FraNcuenta 
         Height          =   675
         Left            =   11490
         TabIndex        =   29
         Top             =   30
         Visible         =   0   'False
         Width           =   2835
         Begin VB.TextBox txtNcuenta 
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
            Left            =   1080
            MaxLength       =   30
            TabIndex        =   31
            Top             =   210
            Width           =   1245
         End
         Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
            Caption         =   "..."
            Height          =   315
            Left            =   2400
            TabIndex        =   30
            ToolTipText     =   "Busca Cuenta por Apellidos y Nombres"
            Top             =   210
            Width           =   315
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "N° Cuenta"
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
            TabIndex        =   32
            Top             =   225
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdProrratea 
         DisabledPicture =   "ReembolsosDetalle.frx":180B
         DownPicture     =   "ReembolsosDetalle.frx":1BF4
         Height          =   495
         Left            =   14265
         Picture         =   "ReembolsosDetalle.frx":2000
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Proratea en cada Cuenta el Reembolso total (sólo los marcados)"
         Top             =   3360
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox txtTconsumo 
         Alignment       =   1  'Right Justify
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
         Left            =   9885
         TabIndex        =   27
         Top             =   3390
         Width           =   1335
      End
      Begin VB.TextBox txtTporReembolsar 
         Alignment       =   1  'Right Justify
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
         Left            =   11265
         TabIndex        =   26
         Top             =   3390
         Width           =   1335
      End
      Begin VB.TextBox txtTreembolso 
         Alignment       =   1  'Right Justify
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
         Left            =   12645
         TabIndex        =   25
         Top             =   3390
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregar 
         DisabledPicture =   "ReembolsosDetalle.frx":2442
         DownPicture     =   "ReembolsosDetalle.frx":282B
         Height          =   315
         Left            =   14250
         Picture         =   "ReembolsosDetalle.frx":2C37
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Agregar"
         Top             =   270
         Width           =   585
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "ConReembolso/PendienteRembolso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   165
         TabIndex        =   19
         Top             =   3390
         Width           =   3405
      End
      Begin UltraGrid.SSUltraGrid grdReembolsos 
         Height          =   3060
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   270
         Width           =   14025
         _ExtentX        =   24739
         _ExtentY        =   5398
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
         Caption         =   "Cuentas  seleccionadas"
      End
      Begin VB.Label Label3 
         Caption         =   "<F10> = Detalle_Cta         <F11>=Cambia_Dx          <Supr>  = Eliminar "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3705
         TabIndex        =   33
         Top             =   3420
         Width           =   6105
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtros de Cuentas de los Pacientes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   30
      TabIndex        =   16
      Top             =   30
      Width           =   15015
      Begin VB.CheckBox chkSinBuscarCtas 
         Alignment       =   1  'Right Justify
         Caption         =   "Sin buscar Cuentas"
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
         Left            =   8220
         TabIndex        =   66
         Top             =   1020
         Width           =   1995
      End
      Begin VB.ComboBox cmbTipoConsumo 
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
         ItemData        =   "ReembolsosDetalle.frx":3043
         Left            =   7290
         List            =   "ReembolsosDetalle.frx":3045
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   2925
      End
      Begin VB.ComboBox cmbTipoServicio 
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
         ItemData        =   "ReembolsosDetalle.frx":3047
         Left            =   7290
         List            =   "ReembolsosDetalle.frx":3057
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2925
      End
      Begin VB.ComboBox cmbFuenteFinanciamiento 
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
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   3960
      End
      Begin VB.ComboBox cmbAreaTramitaR 
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
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3960
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   13530
         Picture         =   "ReembolsosDetalle.frx":3096
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Haga click en este botón para filtrar las historias solicitadas"
         Top             =   240
         Width           =   1305
      End
      Begin MSMask.MaskEdBox txtFechaIni 
         Height          =   315
         Left            =   1740
         TabIndex        =   2
         Top             =   1020
         Width           =   1395
         _ExtentX        =   2461
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
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   4230
         TabIndex        =   3
         Top             =   1020
         Width           =   1395
         _ExtentX        =   2461
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Consumo"
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
         Left            =   6090
         TabIndex        =   35
         Top             =   660
         Width           =   1170
      End
      Begin VB.Label Departamento 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Servicio"
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
         Left            =   6210
         TabIndex        =   34
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label17 
         Caption         =   "Area que Tramita"
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
         Left            =   150
         TabIndex        =   24
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label lblFAltaMedica 
         Caption         =   "F.Alta Médica"
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
         Left            =   150
         TabIndex        =   23
         Top             =   1050
         Width           =   1350
      End
      Begin VB.Label lblHasta 
         Caption         =   "al"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3750
         TabIndex        =   22
         Top             =   1050
         Width           =   390
      End
      Begin VB.Label Label14 
         Caption         =   "Fte.Financiam/IAFA"
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
         Left            =   180
         TabIndex        =   21
         Top             =   660
         Width           =   1755
      End
   End
   Begin VB.Frame fraDestino 
      Height          =   435
      Left            =   5310
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   5205
      Begin VB.CommandButton btnAgregarDx 
         DisabledPicture =   "ReembolsosDetalle.frx":5CDF
         DownPicture     =   "ReembolsosDetalle.frx":60C8
         Height          =   315
         Left            =   3180
         Picture         =   "ReembolsosDetalle.frx":64D4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   570
         Width           =   1005
      End
      Begin VB.CommandButton btnQuitarDx 
         DisabledPicture =   "ReembolsosDetalle.frx":68E0
         DownPicture     =   "ReembolsosDetalle.frx":6C6B
         Height          =   315
         Left            =   4260
         Picture         =   "ReembolsosDetalle.frx":6FFE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   570
         Width           =   1005
      End
      Begin VB.TextBox txtNombreServicioDestino 
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
         Left            =   3180
         TabIndex        =   12
         Top             =   180
         Width           =   5115
      End
      Begin VB.TextBox txtIdServicioDestino 
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
         Left            =   1590
         TabIndex        =   11
         Top             =   195
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "Servicio destino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   15
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   30
      TabIndex        =   9
      Top             =   8745
      Width           =   15000
      Begin VB.CommandButton btnReImprime 
         Caption         =   "ReImprime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   225
         Picture         =   "ReembolsosDetalle.frx":738F
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   240
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ReembolsosDetalle.frx":7868
         DownPicture     =   "ReembolsosDetalle.frx":7D2C
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
         Left            =   7657
         Picture         =   "ReembolsosDetalle.frx":8218
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ReembolsosDetalle.frx":8704
         DownPicture     =   "ReembolsosDetalle.frx":8B64
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
         Left            =   6075
         Picture         =   "ReembolsosDetalle.frx":8FD9
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "ReembolsosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Registro de Reembolsos
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Formulario As New sighEntidades.Formulario
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Apariencia As New sighEntidades.GridInfragistic
'
Dim mo_cmbAreaTramitaR As New sighEntidades.ListaDespleglable
Dim mo_cmbFuenteFinanciamiento As New sighEntidades.ListaDespleglable
Dim mo_cmbIdCaja As New ListaDespleglable
Dim mo_cmbIdTurno As New ListaDespleglable
Dim mo_cmbIdTipoComprobante As New ListaDespleglable
Dim mo_cmbTipoConsumo As New ListaDespleglable
'
Dim mo_DoFactReembolsos As New DoFactReembolsos
Dim mo_DoFactReembolsosDcto As New DoFactReembolsosDcto
'
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_sighProxies As New SIGHProxies.Procesos
'
Dim mrs_Reembolsos As New ADODB.Recordset
'
Dim mi_Opcion As sghOpciones
Dim ml_IdFactReembolso As Long
Dim ml_idUsuario As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ms_MensajeError As String
'
Dim lc_FuenteFinanciamientoPermitidos As String
Dim lnCuentaActualDelGrid As Long
Dim lnTotalConsumoParaProrrateo As Double
Dim lnIdComprobantePagoActual As Long, lnIdTipoComprobanteActual As Long, lcNroSerieActual As String, lcNroDocumentoActual As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lcIdReembolsoPorCuenta As String
Dim ml_SoloSeIngresaUnaCuenta As Boolean, lbYaTieneFactura As Boolean, lbTieneGrabadoFechaCobro As Boolean
Dim lbTieneLicenciaParaNotaCreditoYsunat As Boolean

Property Let SoloSeIngresaUnaCuenta(lValue As Boolean)
   ml_SoloSeIngresaUnaCuenta = lValue
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let IdFactReembolso(lValue As Long)
   ml_IdFactReembolso = lValue
End Property
Property Get IdFactReembolso() As Long
   IdFactReembolso = ml_IdFactReembolso
End Property



Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   Dim oRsTmp As New Recordset
   ValidarDatosObligatorios = False
   If txtDescripcionR.Text = "" Then
       sMensaje = sMensaje + "Ingrese la Descripción del Reembolso " + Chr(13)
       txtDescripcionR.SetFocus
   End If
   If mrs_Reembolsos.RecordCount = 0 Then
       sMensaje = sMensaje + "Debe haber al menos una Cuenta en la Lista " + Chr(13)
   End If
   If Val(Me.txtTreembolso.Text) <= 0 Then
       sMensaje = sMensaje + "Ingrese el TOTAL REEMBOLSADO" + Chr(13)
       'Me.txtTreembolso.SetFocus
   End If
   If txtNroSerie.Text = "" Then
      sMensaje = sMensaje + "Ingrese el NRO SERIE" + Chr(13)
   End If
   If txtNroDocumento.Text = "" Then
      sMensaje = sMensaje + "Ingrese el NRO DOCUMENTO" + Chr(13)
   End If
   
   If mrs_Reembolsos.RecordCount > 0 Then
      mrs_Reembolsos.MoveFirst
      Do While Not mrs_Reembolsos.EOF
         If mrs_Reembolsos.Fields!seleccionar = True Then
            If Val(mo_cmbAreaTramitaR.BoundText) = 4 Then
'               If mrs_Reembolsos.Fields!nroReferenciaDestino = "" Or IsNull(mrs_Reembolsos.Fields!nroReferenciaDestino) Then
'                  sMensaje = sMensaje + "Falta N° REFERENCIA para el Paciente: " + mrs_Reembolsos.Fields!Paciente + Chr(13)
'               End If
            Else
               'If mrs_Reembolsos.Fields!dxid = 0 And (mrs_Reembolsos.Fields!IdTipoServicio = 2 Or mrs_Reembolsos.Fields!IdTipoServicio = 3 Or mrs_Reembolsos.Fields!IdTipoServicio = 4) Then
               '   sMensaje = sMensaje + "Falta DIAGNOSTICO para el Paciente: " + mrs_Reembolsos.Fields!Paciente + Chr(13)
               'End If
            End If
         End If
         mrs_Reembolsos.MoveNext
      Loop
   End If
   txtNroSerie.Text = Trim(txtNroSerie.Text)
   txtNroDocumento.Text = Trim(txtNroDocumento.Text)
   If txtNroSerie.Text <> "" Or txtNroDocumento.Text <> "" Then
      If Left(txtNroSerie.Text, 1) <> "B" And Left(txtNroSerie.Text, 1) <> "F" Then
           sMensaje = sMensaje + "El NUMERO DE SERIE debe empezar con la letra   B   o    F " + Chr(13)
      End If
      If cmbIdTurno.Text = "" Then
         sMensaje = sMensaje + "Elija el TURNO para Generar Documento en CAJA" + Chr(13)
         cmbIdTurno.SetFocus
      End If
      If cmbIdCaja.Text = "" Then
         sMensaje = sMensaje + "Elija la CAJA para Generar Documento en CAJA" + Chr(13)
         cmbIdCaja.SetFocus
      End If
      If cmbIdTipoComprobante.Text = "" Then
         sMensaje = sMensaje + "Elija el TIPO DE COMPROBANTE para Generar Documento en CAJA" + Chr(13)
         cmbIdTipoComprobante.SetFocus
      End If
      If txtNroSerie.Text = "" Then
         sMensaje = sMensaje + "Ingrese el  NUMERO DE SERIE para Generar Documento en CAJA" + Chr(13)
         txtNroSerie.SetFocus
      End If
      If txtNroDocumento.Text = "" Then
         sMensaje = sMensaje + "Ingrese el  NUMERO DE DOCUMENTO para Generar Documento en CAJA" + Chr(13)
         txtNroDocumento.SetFocus
      End If
      If txtNroSerie.Text <> "" And txtNroDocumento.Text <> "" Then
         Select Case mi_Opcion
         Case sghAgregar
            lnIdComprobantePagoActual = 0
            Set oRsTmp = mo_AdminCaja.CajaComprobantesPagoSeleccionarPorNroSerieNroDocumento(txtNroSerie.Text, txtNroDocumento.Text)
            oRsTmp.Filter = "idTipoComprobante=" & mo_cmbIdTipoComprobante.BoundText
            If oRsTmp.RecordCount > 0 Then
                If MsgBox("Ese NUMERO DE SERIE + DOCUMENTO + COMPROBANTE " & IIf(oRsTmp.Fields!idEstadoComprobante = 9, "(anulado)", "(pagado)") & " ya fué emitido el: " & Format(oRsTmp.Fields!FechaCobranza, "dd/mm/yyyy") & Chr(13) & "Esta seguro de Continuar ?", vbQuestion + vbYesNo, "") = vbYes Then
                    lnIdComprobantePagoActual = oRsTmp.Fields!IdComprobantePago
                Else
                    sMensaje = sMensaje + "No se pudo Grabar" + Chr(13)
                    txtNroDocumento.SetFocus
                End If
            End If
         Case sghModificar
            If lcNroSerieActual <> txtNroSerie.Text Or lcNroDocumentoActual <> txtNroDocumento.Text Or lnIdTipoComprobanteActual <> Val(mo_cmbIdTipoComprobante.BoundText) Then
                Set oRsTmp = mo_AdminCaja.CajaComprobantesPagoSeleccionarPorNroSerieNroDocumento(txtNroSerie.Text, txtNroDocumento.Text)
                oRsTmp.Filter = "idTipoComprobante=" & Val(mo_cmbIdTipoComprobante.BoundText)
                If oRsTmp.RecordCount > 0 Then
                   sMensaje = sMensaje + "Ese NUMERO DE SERIE + DOCUMENTO + COMPROBANTE " + IIf(oRsTmp.Fields!idEstadoComprobante = 9, "(anulado)", "(pagado)") + " ya fué emitido el: " + Format(oRsTmp.Fields!FechaCobranza, "dd/mm/yyyy") + Chr(13)
                   txtNroDocumento.SetFocus
                End If
            End If
         End Select
         If mo_cmbIdTipoComprobante.BoundText = "2" Then    'Factura
            If txtRuc.Text = "" Then
                sMensaje = sMensaje + "Ingrese el Número  de RUC" + Chr(13)
                txtRuc.SetFocus
            End If
            If txtRazonSocial.Text = "" Then
                sMensaje = sMensaje + "Ingrese la RAZON SOCIAL" + Chr(13)
                txtRazonSocial.SetFocus
            End If
         End If
      End If
   End If
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   Set oRsTmp = Nothing
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   Dim sMensaje As String
   ValidarReglas = False
   sMensaje = ""
   If mi_Opcion = sghEliminar And Me.txtMotivoAnulacion.Text = "" Then
       sMensaje = sMensaje + "Debe ingresar el Motivo de la Anulación" + Chr(13)
       Me.txtMotivoAnulacion.SetFocus
   End If
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
    ValidarReglas = True
End Function

Sub ActualizaTablaFacturacionCuentasAtencionPtos()
    If mrs_Reembolsos.RecordCount > 0 Then
       mrs_Reembolsos.MoveFirst
       Do While Not mrs_Reembolsos.EOF
          mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar mrs_Reembolsos.Fields!nrocuenta, False, 0
          mrs_Reembolsos.MoveNext
       Loop
    End If
End Sub

Sub ActualizaEstadoDeCuentaApagada()
    On Error GoTo errActEsCta
    Dim lnConsumoS As Double, lnConsumoF As Double, lnIdEstadoCuentaNuevo As Long, lnConsumoFSenReembolso As Double
    Dim oRsTmp8 As New Recordset
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    If mrs_Reembolsos.RecordCount > 0 Then
       mrs_Reembolsos.MoveFirst
       Do While Not mrs_Reembolsos.EOF
          lnConsumoS = mo_ReglasFacturacion.RetornaConsumoPacienteServiciosConSeguroPorNroCuenta(mrs_Reembolsos!nrocuenta, True)
          lnConsumoF = mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(mrs_Reembolsos!nrocuenta)
          Set oRsTmp8 = mo_ReglasFacturacion.FacturacionReembolsosSeleccionarPorCuenta(mrs_Reembolsos!nrocuenta, oConexion)
          oRsTmp8.Filter = "grabaDefinitivamente=1"
          lnIdEstadoCuentaNuevo = sghEstadoCuenta.sghPendientePagoSeguros
          If oRsTmp8.RecordCount > 0 Then
             lnConsumoFSenReembolso = 0
             oRsTmp8.MoveFirst
             Do While Not oRsTmp8.EOF
                lnConsumoFSenReembolso = lnConsumoFSenReembolso + oRsTmp8!consumo
                oRsTmp8.MoveNext
             Loop
             If lnConsumoFSenReembolso >= (lnConsumoS + lnConsumoF) Then
                lnIdEstadoCuentaNuevo = sghEstadoCuenta.sghPagado
             Else
                lnIdEstadoCuentaNuevo = sghEstadoCuenta.sghReembolsoParcial
             End If
          Else
             oRsTmp8.Filter = "grabaDefinitivamente=0"
             If oRsTmp8.RecordCount > 0 Then
                lnIdEstadoCuentaNuevo = sghEstadoCuenta.sghReembolsoParcial
             End If
          End If
          oRsTmp8.Close
          mo_ReglasFacturacion.FacturacionCuentasAtencionActualizarIdEstado oConexion, lnIdEstadoCuentaNuevo, _
                                                                            mrs_Reembolsos!nrocuenta
          
          mrs_Reembolsos.MoveNext
       Loop
    End If
errActEsCta:
    oConexion.Close
    Set oConexion = Nothing
    Set oRsTmp8 = Nothing
    Exit Sub
    Resume
End Sub

Function AgregarDatos() As Boolean
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    mo_ReglasFacturacion.ActualizaRsTemporalConReembolsosPagadosAdelantosConsumos mo_DoFactReembolsos.IdFactReembolso, mo_DoFactReembolsos.idAreaTramitaSeguro, mo_DoFactReembolsos.IdFuenteFinanciamiento, mo_DoFactReembolsos.idTipoConsumo, mrs_Reembolsos, oConexion
    oConexion.Close
    Set oConexion = Nothing
    AgregarDatos = ActualizaDatosDeCaja
    If AgregarDatos = True Then
        AgregarDatos = mo_ReglasFacturacion.ReembolsosAgregar(mo_DoFactReembolsos, mo_DoFactReembolsosDcto, mrs_Reembolsos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")
        ms_MensajeError = mo_ReglasFacturacion.MensajeError
    End If
    ActualizaTablaFacturacionCuentasAtencionPtos
    'kike 2017
    If ms_MensajeError = "" Then
        If lbTieneLicenciaParaNotaCreditoYsunat = True Then
            Dim oExportar As New SIGHProxies.Procesos
            oExportar.ExportarFacturasBoletas "", "", txtNroSerie.Text, txtNroDocumento.Text, False
            Set oExportar = Nothing
        End If
        '
        ActualizaProveedor
        ImprimeFactura
        ActualizaEstadoDeCuentaApagada
    End If
    
End Function
Function ModificarDatos() As Boolean
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    mo_ReglasFacturacion.ActualizaRsTemporalConReembolsosPagadosAdelantosConsumos mo_DoFactReembolsos.IdFactReembolso, mo_DoFactReembolsos.idAreaTramitaSeguro, mo_DoFactReembolsos.IdFuenteFinanciamiento, mo_DoFactReembolsos.idTipoConsumo, mrs_Reembolsos, oConexion
    oConexion.Close
    Set oConexion = Nothing
    ModificarDatos = ActualizaDatosDeCaja
    If ModificarDatos = True Then
        ModificarDatos = mo_ReglasFacturacion.ReembolsosModificar(mo_DoFactReembolsos, mo_DoFactReembolsosDcto, _
                                                       mrs_Reembolsos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")
        ms_MensajeError = mo_ReglasFacturacion.MensajeError
    End If
    ActualizaTablaFacturacionCuentasAtencionPtos
    'kike 2017
'    If chkGrabaDefinitivamente.Value = 1 And lbYaTieneFactura = False And ms_MensajeError = "" Then
'        Dim oExportar As New SIGHProxies.Procesos
'        oExportar.ExportarFacturasBoletas "", "", txtNroSerie.Text, txtNroDocumento.Text
'        Set oExportar = Nothing
'        ActualizaProveedor
'        ImprimeFactura
'    End If
    ActualizaProveedor
    ActualizaEstadoDeCuentaApagada
    'ImprimeFactura
End Function

Sub InicilizarParametros()
        wxParametro102 = lcBuscaParametro.SeleccionaFilaParametro(102)
        wxParametro205 = lcBuscaParametro.SeleccionaFilaParametro(205)
        wxParametro206 = lcBuscaParametro.SeleccionaFilaParametro(206)
        wxParametro208 = lcBuscaParametro.SeleccionaFilaParametro(208)
        wxParametro207 = lcBuscaParametro.SeleccionaFilaParametro(207)
        wxParametro211 = lcBuscaParametro.SeleccionaFilaParametro(211)
        wxParametro221 = lcBuscaParametro.SeleccionaFilaParametro(221)
        wxParametro208 = lcBuscaParametro.SeleccionaFilaParametro(208)
        wxParametro285 = lcBuscaParametro.SeleccionaFilaParametro(285)
        wxParametro280 = lcBuscaParametro.SeleccionaFilaParametro(280)
        wxParametro286 = lcBuscaParametro.SeleccionaFilaParametro(286)
        wxParametro288 = lcBuscaParametro.SeleccionaFilaParametro(288)
        wxParametro339 = lcBuscaParametro.SeleccionaFilaParametro(339)
        wxParametro346 = lcBuscaParametro.SeleccionaFilaParametro(346)
        wxParametro379 = lcBuscaParametro.SeleccionaFilaParametro(379) 'sunat
        wxParametro381 = lcBuscaParametro.SeleccionaFilaParametro(381)
        wxParametro386 = lcBuscaParametro.SeleccionaFilaParametro(386)
        wxParametro387 = lcBuscaParametro.SeleccionaFilaParametro(387)
        wxParametro377 = lcBuscaParametro.SeleccionaFilaParametro(377)
        wxParametro500 = UCase(lcBuscaParametro.SeleccionaFilaParametro(500))   'debb-18/05/2016
        wxParametro501 = lcBuscaParametro.SeleccionaFilaParametro(501)  'debb-18/05/2016
        lcParametro523 = lcBuscaParametro.SeleccionaFilaParametro(523)
        lcParametro524 = lcBuscaParametro.SeleccionaFilaParametro(524)
        wxParametro527 = lcBuscaParametro.SeleccionaFilaParametro(527)
        wxParametro533 = lcBuscaParametro.SeleccionaFilaParametro(533)
        wxParametro538 = lcBuscaParametro.SeleccionaFilaParametro(538)
        wxParametro543 = lcBuscaParametro.SeleccionaFilaParametro(543)   'S
        wxParametro548 = lcBuscaParametro.SeleccionaFilaParametro(548)
        wxParametro549 = lcBuscaParametro.SeleccionaFilaParametro(549)
        
        
End Sub
Sub ImprimeFactura()
    On Error GoTo ErrImpF
    If lbTieneGrabadoFechaCobro = False Then
       If mo_AdminCaja.CajaComprobantesPagoActualizaFechaCobranza(lnIdComprobantePagoActual, CDate(txtFemision.Text)) = False Then
          MsgBox "Primero GRABE", vbInformation, ""
          Exit Sub
       End If
    End If
    
    Dim oDOCajaCaja1 As New DOCajaCaja
    Set oDOCajaCaja1 = mo_AdminCaja.CajaSeleccionarPorId(Val(mo_cmbIdCaja.BoundText))
    CargaSetup_Caja App.Path & "\archivos", oDOCajaCaja1.IdTipoComprobante, False
    Set oDOCajaCaja1 = Nothing
    
    InicilizarParametros

    Dim oImprimeBoletaContinua As New RptCaja
    Dim oRsBusquedaRecibos As New Recordset
    Dim lcRuc99 As String
    Set oRsBusquedaRecibos = mo_AdminCaja.CajaComprobantePagoSeleccionarPorFechaOdocumento(txtNroSerie.Text, txtNroDocumento.Text, 0, 0)
    If oRsBusquedaRecibos.RecordCount > 0 Then
       'wxParametro527 = "S"
       wxParametro288 = "S"
       lcRuc99 = IIf(IsNull(oRsBusquedaRecibos!ruc), "", oRsBusquedaRecibos!ruc)
       oImprimeBoletaContinua.ImpresionBoletaEnDosTYPE True, DevuelveRUCyDIRECCIONProveedor(False, lcRuc99), _
                                                    "EFECTIVO:       VUELTO:", _
                                                    oRsBusquedaRecibos!nroSerie, _
                                                    oRsBusquedaRecibos!nrodocumento, _
                                                    False, _
                                                    IIf(oRsBusquedaRecibos!IdTipoOrden = 1, _
                                                           sighEntidades.sghServicio, sighEntidades.sghbien), _
                                                    True, True
    End If
    oRsBusquedaRecibos.Close
    Set oImprimeBoletaContinua = Nothing
    Set oRsBusquedaRecibos = Nothing
    'wxParametro527 = lcBuscaParametro.SeleccionaFilaParametro(527)
    wxParametro288 = lcBuscaParametro.SeleccionaFilaParametro(288)
ErrImpF:
    If lbTieneGrabadoFechaCobro = False Then
       If mo_AdminCaja.CajaComprobantesPagoActualizaFechaCobranza(lnIdComprobantePagoActual, 0) = True Then
       End If
    End If
End Sub

Function DevuelveRUCyDIRECCIONProveedor(lbDesdeBotonAceptar As Boolean, lcRuc As String) As String
    Dim lcDireccion1 As String, lcRuc1 As String
    lcDireccion1 = "": lcRuc1 = ""
    If lbDesdeBotonAceptar = True Then
'        If txtRuc.Text <> "" Then
'           lcRuc1 = txtRuc.Text
'           lcDireccion1 = txtDireccionProv.Text
'        End If
    ElseIf lcRuc <> "" Then
          Dim oRsTmp As New ADODB.Recordset
          Set oRsTmp = mo_ReglasFacturacion.ProveedoresSeleccionarPorRUC(lcRuc)
          If oRsTmp.RecordCount > 0 Then
             lcDireccion1 = IIf(IsNull(oRsTmp!Direccion), "", oRsTmp!Direccion) & IIf(IsNull(oRsTmp!Email), "", "   (EMAIL: " & oRsTmp!Email & ")")
             lcRuc1 = lcRuc
          End If
          oRsTmp.Close
          Set oRsTmp = Nothing
    End If
    If lcRuc1 = "" Then
       DevuelveRUCyDIRECCIONProveedor = ""
    Else
       DevuelveRUCyDIRECCIONProveedor = "RUC: " & lcRuc1 & "    DIRECCION: " & lcDireccion1
    End If
End Function

Function EliminarDatos() As Boolean
    EliminarDatos = ActualizaDatosDeCaja
    If EliminarDatos = True Then
        mo_DoFactReembolsos.IdUsuarioAuditoria = ml_idUsuario
        mo_DoFactReembolsosDcto.MotivoAnulacion = Me.txtMotivoAnulacion.Text
        EliminarDatos = mo_ReglasFacturacion.ReembolsosAnular(mo_DoFactReembolsos, mo_DoFactReembolsosDcto, mrs_Reembolsos, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")
        ms_MensajeError = mo_ReglasFacturacion.MensajeError
    End If
    ActualizaTablaFacturacionCuentasAtencionPtos
    MsgBox "Tiene que hacer una NOTA DE CREDITO en CAJA", vbInformation, ""
'    If lbTieneLicenciaParaNotaCreditoYsunat = True Then
'        Dim oExportar As New SIGHProxies.Procesos
'        oExportar.ExportarFacturasBoletas "", "", txtNroSerie.Text, txtNroDocumento.Text, False
'        Set oExportar = Nothing
'    End If
    ActualizaEstadoDeCuentaApagada
End Function
Function ActualizaDatosDeCaja() As Boolean
    If mi_Opcion = sghModificar And chkGrabaDefinitivamente.Value = 1 Then
       If mo_AdminCaja.CajaComprobantesPagoActualizaFechaCobranza(lnIdComprobantePagoActual, CDate(txtFcobro.Text)) = True Then
          If mo_AdminCaja.CajaComprobantesPagoActualizaTieneCredito(lnIdComprobantePagoActual, "") Then
                ActualizaDatosDeCaja = True
                lbTieneGrabadoFechaCobro = True
          Else
                MsgBox mo_AdminCaja.MensajeError, vbInformation, ""
          End If
       Else
          MsgBox mo_AdminCaja.MensajeError, vbInformation, ""
       End If
       Exit Function
    End If
    
    Dim mo_DOComprobantePago As New DOCajaComprobantesPago
    Dim mo_doCajaGestion As New DOCajaGestion
    Dim mo_DoFactOrdenServPagos As New DoFactOrdenServPagos
    Dim mo_DoAtencion As New DOAtencion
    Dim oRsFacturacionProductos As New Recordset
    Dim oRsCajaComprobantePago As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp5 As New Recordset
    Dim oDllFactUCGestionCaja As New SighFacturacion.dllFactUCGestionCaja
    Dim lnSubTotal As Double, lnTotal As Double
    Dim lnIGV As Double, lnIdProducto As Long
    Dim lcSql As String, lnIdOrden As Long, lnIdOrdenPago As Long
    Dim lnReembolsoPagadoServicio As Double, lnReembolsoPagadoFarmacia As Double
    Dim lnIdTipoConsumoUltimo As Long
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    If txtNroSerie.Text <> "" And txtNroDocumento.Text <> "" Then
        With oRsFacturacionProductos
              .Fields.Append "IdFacturacionProducto", adInteger
              .Fields.Append "IdProducto", adInteger
              .Fields.Append "Codigo", adVarChar, 255, adFldIsNullable
              .Fields.Append "NombreProducto", adVarChar, 255, adFldIsNullable
              .Fields.Append "IdTipoFinanciamiento", adInteger
              .Fields.Append "IdFuenteFinanciamiento", adInteger, , adFldIsNullable
              .Fields.Append "Poliza", adVarChar, 255
              .Fields.Append "TipoFinanciamiento", adVarChar, 255
              .Fields.Append "Cantidad", adInteger
              .Fields.Append "PrecioUnitario", adCurrency
              .Fields.Append "TotalPorPagar", adCurrency
              .Fields.Append "IdEstadoFacturacion", adInteger
              .Fields.Append "IdPuntoCarga", adInteger
              .Fields.Append "IdAtencion", adInteger, , adFldIsNullable
              .Fields.Append "IdCajero", adInteger, , adFldIsNullable
              .Fields.Append "FechaAutorizaPendiente", adDBTimeStamp, , adFldIsNullable
              .Fields.Append "FechaAutorizaSeguro", adDBTimeStamp, , adFldIsNullable
              .Fields.Append "FechaAutorizaDevolucion", adDBTimeStamp, , adFldIsNullable
              .Fields.Append "FechaCajero", adDBTimeStamp, , adFldIsNullable
              .Fields.Append "IdUsuarioAutorizaPendiente", adInteger, , adFldIsNullable
              .Fields.Append "IdUsuarioAutorizaSeguro", adInteger, , adFldIsNullable
              .Fields.Append "IdUsuarioAutorizaDevolucion", adInteger, , adFldIsNullable
              .Fields.Append "IdServicioInternamiento", adInteger, , adFldIsNullable
              .Fields.Append "IdUsuarioAuditoria", adInteger, , adFldIsNullable
              .Fields.Append "EstadoLocal", adVarChar, 1
              .Fields.Append "IdComprobantePago", adInteger, , adFldIsNullable
              .Fields.Append "IdComprobantePagoDevolucion", adInteger, , adFldIsNullable
              .Fields.Append "IdOrden", adInteger
              .Fields.Append "movTipo", adVarChar, 1, adFldIsNullable
              .Fields.Append "movNumero", adVarChar, 9, adFldIsNullable
              .Fields.Append "SeUsaSinPrecio", adBoolean
              .CursorType = adOpenDynamic
              .LockType = adLockOptimistic
              .Open
        End With
        lnSubTotal = 0
        lnIGV = 0
        lnTotal = CCur(txtTreembolso.Text)
        Select Case mi_Opcion
        Case sghAgregar, sghModificar
             If mi_Opcion = sghAgregar Then
                 With mo_DOComprobantePago
                    .IdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
                    .nroSerie = Trim(txtNroSerie.Text)
                    .nrodocumento = Trim(txtNroDocumento.Text)
                    .idCuentaAtencion = 0
                    .razonSocial = Me.txtRazonSocial.Text
                    .Observaciones = IIf(lblDescripcionLarga.Visible = True, txtDescripcionR.Text, "CUENTA_R:" & Trim(Str(lnCuentaActualDelGrid)))
                    .IdGestionCaja = 0   'mo_doCajaGestion.IdGestionCaja
                    .IdUsuarioAuditoria = ml_idUsuario
                    .ruc = txtRuc.Text
                    .Subtotal = lnSubTotal
                    .IGV = lnIGV
                    .Total = lnTotal
                    If mi_Opcion = sghAgregar And Me.chkGrabaDefinitivamente.Value = 1 Then
                       .FechaCobranza = CDate(Me.txtFcobro.Text)
                       lbTieneGrabadoFechaCobro = True
                    End If
                    .fechaEmision = lcBuscaParametro.RetornaFechaHoraServidorSQL
                    .IdComprobantePago = 0
                    .IdTipoPago = 1 'Orden de pago
                    .idPaciente = 0
                    .IdFormaPago = 0
                    .idFarmacia = 0
                    .IdCaja = Val(mo_cmbIdCaja.BoundText)
                    .IdTurno = Val(mo_cmbIdTurno.BoundText)
                    .IdCajero = ml_idUsuario
                    .IdUsuarioAuditoria = ml_idUsuario
                    .Dctos = 0
                    .exoneraciones = 0
                    .Adelantos = 0
                    .idTipoFinanciamiento = 1
                    If chkGrabaDefinitivamente.Value = 0 Then
                       .TieneCredito = "C"
                    End If
                 End With
                 If mo_cmbIdTipoComprobante.BoundText = 2 Then
                     'Es una Factura:
                     If chkIGV.Value = 1 Then
                        lnIGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221)) / 100
                        mo_DOComprobantePago.Subtotal = Round(mo_DOComprobantePago.Total / (lnIGV + 1), 2)
                        mo_DOComprobantePago.IGV = Round(mo_DOComprobantePago.Total * lnIGV / (lnIGV + 1), 2)
                     Else
                        mo_DOComprobantePago.Subtotal = mo_DOComprobantePago.Total
                        mo_DOComprobantePago.IGV = 0
                     End If
'                     If mo_cmbTipoConsumo.BoundText = "1" Then
'                        'Sólo se calcula IGV en Farmacia/
'                        lnIGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221)) / 100
'                        mo_DOComprobantePago.Subtotal = Round(mo_DOComprobantePago.Total / (lnIGV + 1), 2)
'                        mo_DOComprobantePago.IGV = Round(mo_DOComprobantePago.Total * lnIGV / (lnIGV + 1), 2)
'                     Else
'                        'No hay IGV en FACTURA DE SERVICIOS
'                        mo_DOComprobantePago.Subtotal = mo_DOComprobantePago.Total
'                        mo_DOComprobantePago.IGV = 0
'                     End If
                 Else
                     mo_DOComprobantePago.ruc = ""
                 End If
                 '
                 With mo_DoFactOrdenServPagos
                     .fechacreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
                     .idestadofacturacion = 4 'Pagado
                     .IdOrden = 0
                     .idUsuario = ml_idUsuario
                     .IdUsuarioAuditoria = ml_idUsuario
                     .IdOrdenPago = 0
                 End With
                 '
                 Select Case mo_cmbTipoConsumo.BoundText
                 Case "2"
                   'servicio
                   If lblDescripcionLarga.Visible = True Then
                      'Caja con DESCRIPCION LARGA (SIN CPT, SOLO TEXTO)
                      lnIdProducto = Val(lcBuscaParametro.SeleccionaFilaParametro(549))
                   Else
                      lnIdProducto = Val(lcBuscaParametro.SeleccionaFilaParametro(251))
                   End If
                   oRsFacturacionProductos.AddNew
                   oRsFacturacionProductos.Fields!idProducto = lnIdProducto
                   oRsFacturacionProductos.Fields!Cantidad = 1
                   oRsFacturacionProductos.Fields!PrecioUnitario = CCur(txtTreembolso.Text)
                   oRsFacturacionProductos.Fields!TotalPorPagar = CCur(txtTreembolso.Text)
                   oRsFacturacionProductos.Update
                 Case "1"
                   'farmacia
                   If lblDescripcionLarga.Visible = True Then
                      'Caja con DESCRIPCION LARGA (SIN CPT, SOLO TEXTO)
                      lnIdProducto = Val(lcBuscaParametro.SeleccionaFilaParametro(549))
                   Else
                      lnIdProducto = Val(lcBuscaParametro.SeleccionaFilaParametro(252))
                   End If
                   oRsFacturacionProductos.AddNew
                   oRsFacturacionProductos.Fields!idProducto = lnIdProducto
                   oRsFacturacionProductos.Fields!Cantidad = 1
                   oRsFacturacionProductos.Fields!PrecioUnitario = CCur(txtTreembolso.Text)
                   oRsFacturacionProductos.Fields!TotalPorPagar = CCur(txtTreembolso.Text)
                   oRsFacturacionProductos.Update
                Case Else
                   If lblDescripcionLarga.Visible = True Then
                        'Caja con DESCRIPCION LARGA (SIN CPT, SOLO TEXTO)
                        lnIdProducto = Val(lcBuscaParametro.SeleccionaFilaParametro(549))
                        oRsFacturacionProductos.AddNew
                        oRsFacturacionProductos.Fields!idProducto = lnIdProducto
                        oRsFacturacionProductos.Fields!Cantidad = 1
                        oRsFacturacionProductos.Fields!PrecioUnitario = CCur(txtTreembolso.Text)
                        oRsFacturacionProductos.Fields!TotalPorPagar = CCur(txtTreembolso.Text)
                        oRsFacturacionProductos.Update
                        lnReembolsoPagadoFarmacia = 0
                        lnReembolsoPagadoServicio = CCur(txtTreembolso.Text)
                   Else
                        lnReembolsoPagadoFarmacia = 0
                        lnReembolsoPagadoServicio = 0
                        mrs_Reembolsos.MoveFirst
                        Do While Not mrs_Reembolsos.EOF
                           lnReembolsoPagadoFarmacia = lnReembolsoPagadoFarmacia + mrs_Reembolsos.Fields!ReembolsoPagadoFarmacia
                           lnReembolsoPagadoServicio = lnReembolsoPagadoServicio + mrs_Reembolsos.Fields!ReembolsoPagadoServicio
                           mrs_Reembolsos.MoveNext
                        Loop
                        'servicio
                        lnIdProducto = Val(lcBuscaParametro.SeleccionaFilaParametro(251))
                        oRsFacturacionProductos.AddNew
                        oRsFacturacionProductos.Fields!idProducto = lnIdProducto
                        oRsFacturacionProductos.Fields!Cantidad = 1
                        oRsFacturacionProductos.Fields!PrecioUnitario = lnReembolsoPagadoServicio
                        oRsFacturacionProductos.Fields!TotalPorPagar = lnReembolsoPagadoServicio
                        oRsFacturacionProductos.Update
                        'farmacia
                        lnIdProducto = Val(lcBuscaParametro.SeleccionaFilaParametro(252))
                        oRsFacturacionProductos.AddNew
                        oRsFacturacionProductos.Fields!idProducto = lnIdProducto
                        oRsFacturacionProductos.Fields!Cantidad = 1
                        oRsFacturacionProductos.Fields!PrecioUnitario = lnReembolsoPagadoFarmacia
                        oRsFacturacionProductos.Fields!TotalPorPagar = lnReembolsoPagadoFarmacia
                        oRsFacturacionProductos.Update
                   End If
                 End Select
                 '
                 If mi_Opcion = sghAgregar Then
                   If lnIdComprobantePagoActual = 0 Then
                       '----Agrega documento nuevo
                       '
                       'ActualizaDatosDeCaja = mo_AdminCaja.CajaComprobantePagoServicioAgregar(mo_DOComprobantePago, mo_doCajaGestion, mo_DoFactOrdenServPagos, oRsFacturacionProductos, ml_idUsuario, mo_DOAtencion, 99, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, 0)
                       ActualizaDatosDeCaja = oDllFactUCGestionCaja.CajaComprobantePagoServicioAgregar(mo_DOComprobantePago, mo_doCajaGestion, mo_DoFactOrdenServPagos, oRsFacturacionProductos, ml_idUsuario, mo_DoAtencion, 99, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, 0, False)
                       '
                       ms_MensajeError = mo_AdminCaja.MensajeError
                       mo_DoFactReembolsos.IdTipoComprobante = mo_DOComprobantePago.IdTipoComprobante
                       mo_DoFactReembolsosDcto.IdComprobantePago = mo_DOComprobantePago.IdComprobantePago
                       lnIdComprobantePagoActual = mo_DOComprobantePago.IdComprobantePago
                       ActualizaDatosDeCaja = True
                   Else
                       'Si existe el documento----actualiza importe, estado=pagado, fecha de hoy
                       mo_DoFactReembolsosDcto.IdComprobantePago = lnIdComprobantePagoActual
                       mo_DoFactReembolsos.IdTipoComprobante = mo_DOComprobantePago.IdTipoComprobante
                       mo_AdminCaja.CajaActualizaDatosDeFactura mo_DOComprobantePago.Subtotal, txtRuc.Text, _
                                    txtRazonSocial.Text, mo_DOComprobantePago.IGV, mo_DOComprobantePago.Total, _
                                    Format(mo_DOComprobantePago.FechaCobranza, "dd/mm/yyyy hh:mm:ss"), _
                                    lnIdComprobantePagoActual, IIf(chkGrabaDefinitivamente.Value = 1, "", "C")
                       Set oRsTmp1 = mo_AdminCaja.FactOrdenServicioPagosSeleccionarXidComprobantePago(lnIdComprobantePagoActual)
                       If oRsTmp1.RecordCount > 0 Then
                          lnIdOrden = oRsTmp1.Fields!IdOrden
                          lnIdOrdenPago = oRsTmp1.Fields!IdOrdenPago
                          mo_AdminCaja.CajaActualizaDatosDeServicioYServicioPago lnIdComprobantePagoActual, lnIdOrden
                          Select Case mo_cmbTipoConsumo.BoundText
                          Case "3"     'farmacia+servicios (ambos)
                               'chequea el TIPO DE CONSUMO del anterior reembolso
                               Set oRsTmp1 = mo_AdminCaja.CajaReembolsosAnterioresXfactura(Me.txtNroSerie.Text, Me.txtNroDocumento.Text, Val(mo_cmbIdTipoComprobante.BoundText))
                               If oRsTmp1.RecordCount = 0 Then
                                  lnIdTipoConsumoUltimo = 3
                               Else
                                  lnIdTipoConsumoUltimo = oRsTmp1.Fields!idTipoConsumo
                               End If
                               oRsTmp1.Close
                               '
                               Select Case lnIdTipoConsumoUltimo
                               Case 1  'ultimo reembolso registrado=solo farmacia
                                   mo_AdminCaja.CajaReembolsoActualizaFarmaciaAddServicio lnReembolsoPagadoFarmacia, lnIdOrdenPago, lnIdOrden, lnReembolsoPagadoServicio
                               Case 2  'ultimo reembolso registrado=solo servicio
                                   mo_AdminCaja.CajaReembolsoActualizaServicioAddFarmacia lnIdOrdenPago, lnIdOrden, lnReembolsoPagadoServicio
                                   
                               Case 3  'ultimo reembolso registrando=ambos
                                   mo_AdminCaja.CajaReembolsoActualizaFarmaciaActualizaServicio lnReembolsoPagadoFarmacia, lnIdOrdenPago, lnIdOrden, lnReembolsoPagadoServicio
                               End Select
                          Case "1"   'Farmacia
                               'chequea el TIPO DE CONSUMO del anterior reembolso
                               Set oRsTmp1 = mo_AdminCaja.CajaReembolsosAnterioresXfactura(Me.txtNroSerie.Text, Me.txtNroDocumento.Text, Val(mo_cmbIdTipoComprobante.BoundText))
                               If oRsTmp1.RecordCount = 0 Then
                                  lnIdTipoConsumoUltimo = 3
                               Else
                                  lnIdTipoConsumoUltimo = oRsTmp1.Fields!idTipoConsumo
                               End If
                               oRsTmp1.Close
                               If lnIdTipoConsumoUltimo = 2 Then
                                   'solo tubo Reeembolso de SErvicio
                                    mo_AdminCaja.CajaReembolsoAddFarmacia lnIdOrdenPago, lnIdOrden, mo_DOComprobantePago.Total
                               Else
                                   'tubo Reembolso de farmacia o ambos (servicio+farmacia)
                                   mo_AdminCaja.CajaReembolsoActualizaServiciosYfarmacia lnIdOrdenPago, lnIdOrden, mo_DOComprobantePago.Total
                               End If
                               mo_AdminCaja.CajaReembolsoEliminaServicio lnIdOrdenPago, lnIdOrden
                          Case "2"   'Servicio
                               'chequea el TIPO DE CONSUMO del anterior reembolso
                               Set oRsTmp1 = mo_AdminCaja.CajaReembolsosAnterioresXfactura(Me.txtNroSerie.Text, Me.txtNroDocumento.Text, Val(mo_cmbIdTipoComprobante.BoundText))
                               If oRsTmp1.RecordCount = 0 Then
                                  lnIdTipoConsumoUltimo = 3
                               Else
                                  lnIdTipoConsumoUltimo = oRsTmp1.Fields!idTipoConsumo
                               End If
                               oRsTmp1.Close
                               If lnIdTipoConsumoUltimo = 1 Then
                                   'solo tubo Reembolso de Farmacia
                                   mo_AdminCaja.CajaReembolsoAddServicio lnIdOrdenPago, lnIdOrden, mo_DOComprobantePago.Total
                               Else
                                   'tubo reembolso de farmacia o Ambos (faramcia+servicio)
                                   mo_AdminCaja.CajaReembolsoActualizaServiciosYfarmacia lnIdOrdenPago, lnIdOrden, mo_DOComprobantePago.Total
                               End If
                               mo_AdminCaja.CajaReembolsoEliminaFarmacia lnIdOrdenPago, lnIdOrden
                          End Select
                       Else
                          oRsTmp1.Close
                       End If
                       ActualizaDatosDeCaja = True
                   End If
                Else
                   '------mi_opcion=modificar
                   If lcNroSerieActual <> txtNroSerie.Text Or lcNroDocumentoActual <> txtNroDocumento.Text Or lnIdTipoComprobanteActual <> Val(mo_cmbIdTipoComprobante.BoundText) Then
                       'Si es nuevo Documento
                       '----Agrega documento nuevo
                       '
                       'ActualizaDatosDeCaja = mo_AdminCaja.CajaComprobantePagoServicioAgregar(mo_DOComprobantePago, mo_doCajaGestion, mo_DoFactOrdenServPagos, oRsFacturacionProductos, ml_idUsuario, mo_DOAtencion, 99, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, 0)
                       ActualizaDatosDeCaja = oDllFactUCGestionCaja.CajaComprobantePagoServicioAgregar(mo_DOComprobantePago, mo_doCajaGestion, mo_DoFactOrdenServPagos, oRsFacturacionProductos, ml_idUsuario, mo_DoAtencion, 99, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, 0, False)
                       '
                       ms_MensajeError = mo_AdminCaja.MensajeError
                       If ActualizaDatosDeCaja = True Then
                           mo_DoFactReembolsos.IdTipoComprobante = mo_DOComprobantePago.IdTipoComprobante
                           mo_DoFactReembolsosDcto.IdComprobantePago = mo_DOComprobantePago.IdComprobantePago
                           '----Anula documento actual
                           Set oRsTmp5 = mo_AdminCaja.FactOrdenServicioPagosSeleccionarXidComprobantePago(lnIdComprobantePagoActual)
                           If oRsTmp5.RecordCount > 0 Then
                             lnIdOrden = oRsTmp5.Fields!IdOrden
                             lnIdOrdenPago = oRsTmp5.Fields!IdOrdenPago
                           End If
                           oRsTmp5.Close
                           Set mo_DOComprobantePago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(lnIdComprobantePagoActual, oConexion)
                           ActualizaDatosDeCaja = mo_AdminCaja.CajaComprobantePagoServicioAnulaBoleta(mo_DOComprobantePago, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lnIdOrden)
                           ms_MensajeError = mo_AdminCaja.MensajeError
                       End If
                   Else
                       'Si existe el documento----actualiza importe
                       mo_DoFactReembolsosDcto.IdComprobantePago = lnIdComprobantePagoActual
                       mo_AdminCaja.CajaActualizaDatosDeFactura mo_DOComprobantePago.Subtotal, txtRuc.Text, _
                                    txtRazonSocial.Text, mo_DOComprobantePago.IGV, mo_DOComprobantePago.Total, _
                                    Format(mo_DOComprobantePago.FechaCobranza, "dd/mm/yyyy hh:mm:ss"), _
                                    lnIdComprobantePagoActual, IIf(chkGrabaDefinitivamente.Value = 1, "", "C")
                       Set oRsTmp1 = mo_AdminCaja.FactOrdenServicioPagosSeleccionarXidComprobantePago(lnIdComprobantePagoActual)
                       If oRsTmp1.RecordCount > 0 Then
                          lnIdOrden = oRsTmp1.Fields!IdOrden
                          lnIdOrdenPago = oRsTmp1.Fields!IdOrdenPago
                          oRsTmp1.Close
                          If mo_cmbTipoConsumo.BoundText = 3 Then
                               mo_AdminCaja.CajaReembolsoActualizaFarmaciaActualizaServicio lnReembolsoPagadoFarmacia, lnIdOrdenPago, lnIdOrden, lnReembolsoPagadoServicio
                          Else
                               'Farmacia o Servicio
                               mo_AdminCaja.CajaReembolsoActualizaServiciosYfarmacia lnIdOrdenPago, lnIdOrden, mo_DOComprobantePago.Total
                          End If
                       Else
                          oRsTmp1.Close
                       End If
                       ActualizaDatosDeCaja = True
                   End If
                End If
            Else
                ActualizaDatosDeCaja = True
            End If
        Case sghEliminar
             Set oRsTmp5 = mo_AdminCaja.FactOrdenServicioPagosSeleccionarXidComprobantePago(lnIdComprobantePagoActual)
             If oRsTmp5.RecordCount > 0 Then
               lnIdOrden = oRsTmp5.Fields!IdOrden
               lnIdOrdenPago = oRsTmp5.Fields!IdOrdenPago
             End If
             oRsTmp5.Close
             Set mo_DOComprobantePago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(lnIdComprobantePagoActual, oConexion)
             ActualizaDatosDeCaja = mo_AdminCaja.CajaComprobantePagoServicioAnulaBoleta(mo_DOComprobantePago, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lnIdOrden)
             ms_MensajeError = mo_AdminCaja.MensajeError
            ActualizaDatosDeCaja = True
        End Select
    ElseIf lnIdComprobantePagoActual > 0 And (mi_Opcion = sghModificar Or mi_Opcion = sghEliminar) Then
        Set oRsTmp5 = mo_AdminCaja.FactOrdenServicioPagosSeleccionarXidComprobantePago(lnIdComprobantePagoActual)
        If oRsTmp5.RecordCount > 0 Then
          lnIdOrden = oRsTmp5.Fields!IdOrden
          lnIdOrdenPago = oRsTmp5.Fields!IdOrdenPago
        End If
        oRsTmp5.Close
        Set mo_DOComprobantePago = mo_AdminCaja.ComprobantePagoSeleccionarPorId(lnIdComprobantePagoActual, oConexion)
        ActualizaDatosDeCaja = mo_AdminCaja.CajaComprobantePagoServicioAnulaBoleta(mo_DOComprobantePago, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lnIdOrden)
        ms_MensajeError = mo_AdminCaja.MensajeError
    Else
        ActualizaDatosDeCaja = True
    End If
    Set mo_DOComprobantePago = Nothing
    Set mo_doCajaGestion = Nothing
    Set mo_DoFactOrdenServPagos = Nothing
    Set mo_DoAtencion = Nothing
    Set oRsFacturacionProductos = Nothing
    Set oRsCajaComprobantePago = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp5 = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oDllFactUCGestionCaja = Nothing
End Function



Private Sub btnBuscar_Click()
    If cmbFuenteFinanciamiento.Text = "" Then
       MsgBox "Tiene que elejir la Fuente Financiamiento/IAFA", vbInformation, Me.Caption
       Me.cmbFuenteFinanciamiento.SetFocus
       Exit Sub
    End If
    If Me.cmbTipoConsumo.Text = "" Then
       MsgBox "Tiene que elejir el Tipo de Consumo", vbInformation, Me.Caption
       Me.cmbTipoConsumo.SetFocus
       Exit Sub
    End If
    LimpiarGrilla
    LimpiarDatos
    lc_FuenteFinanciamientoPermitidos = mo_cmbFuenteFinanciamiento.BoundText
    BusquedaPorPlan
    SumaTotales False
    If mo_cmbTipoConsumo.BoundText = "1" Then
       Me.chkIGV.Value = 1
       Me.chkIGV.Enabled = False
    ElseIf mo_cmbTipoConsumo.BoundText = "2" Then
       Me.chkIGV.Value = 0
       Me.chkIGV.Enabled = False
    End If
End Sub

Sub BusquedaPorPlan()
    Dim oRsBusqueda As New Recordset
    Dim oRsBuscaXcta As New Recordset
    Dim lnConsumo As Double
    Dim oDODiagnostico As New DODiagnostico
    'Hospitalizacion y Emergencia
    If cmbTipoServicio.ListIndex <> 0 Or cmbTipoServicio.ListIndex = 3 Then
        Select Case cmbTipoServicio.ListIndex
        Case 3
           'Hosp y Emergencia
           Set oRsBusqueda = mo_ReglasFacturacion.AtencionesSeleccionarHospEmergPorFechaAltaMedicaYplan(CDate(txtFechaIni.Text), CDate(txtFechaFin.Text), Val(mo_cmbFuenteFinanciamiento.BoundText), 0)
        Case 2
           'solo Hospitalizacion
           Set oRsBusqueda = mo_ReglasFacturacion.AtencionesSeleccionarHospEmergPorFechaAltaMedicaYplan(CDate(txtFechaIni.Text), CDate(txtFechaFin.Text), Val(mo_cmbFuenteFinanciamiento.BoundText), 3)
        Case Else
           'solo emergencia
           Set oRsBusqueda = mo_ReglasFacturacion.AtencionesSeleccionarHospEmergPorFechaAltaMedicaYplan(CDate(txtFechaIni.Text), CDate(txtFechaFin.Text), Val(mo_cmbFuenteFinanciamiento.BoundText), 2)
        End Select
        oRsBusqueda.Filter = "idEstado<>4"
'         'referidos/contraReferidos
'        If Val(mo_cmbAreaTramitaR.BoundText) = 4 Then
'           oRsBusqueda.Filter = "NroReferenciaDestino<>null"
'        End If
        If oRsBusqueda.RecordCount > 0 Then
           oRsBusqueda.MoveFirst
           Do While Not oRsBusqueda.EOF
                If Val(mo_cmbAreaTramitaR.BoundText) = 4 Then
                      'referidos/contraReferidos
                      lnConsumo = 0
                Else
                      lnConsumo = RetornaConsumoPorCuenta(oRsBusqueda.Fields!idCuentaAtencion, Val(mo_cmbAreaTramitaR.BoundText), Val(mo_cmbFuenteFinanciamiento.BoundText), Val(mo_cmbTipoConsumo.BoundText))
                End If
                Set oDODiagnostico = mo_ReglasFacturacion.DevuelveDxAltaMedica(oRsBusqueda.Fields!idAtencion, oRsBusqueda.Fields!idTipoServicio)
                Set oRsBuscaXcta = mo_AdminAdmision.AtencionesDatosAdicionalesSeleccionarPorIdCuenta(oRsBusqueda.Fields!idCuentaAtencion)
                mrs_Reembolsos.AddNew
                mrs_Reembolsos.Fields!seleccionar = False
                mrs_Reembolsos.Fields!nrocuenta = oRsBusqueda.Fields!idCuentaAtencion
                mrs_Reembolsos.Fields!EstadoCuenta = oRsBusqueda.Fields!dEstadoCuenta
                mrs_Reembolsos.Fields!NroHistoria = oRsBusqueda.Fields!NroHistoriaClinica
                mrs_Reembolsos.Fields!Paciente = oRsBusqueda.Fields!Paciente
                mrs_Reembolsos.Fields!Servicio = oRsBusqueda.Fields!DServicio
                mrs_Reembolsos.Fields!fAltaMedica = oRsBusqueda.Fields!fechaEgreso
                mrs_Reembolsos.Fields!consumo = lnConsumo
                mrs_Reembolsos.Fields!Reembolsado = 0
                mrs_Reembolsos.Fields!porReembolsar = 0
                mrs_Reembolsos.Fields!dxid = oDODiagnostico.idDiagnostico
                mrs_Reembolsos.Fields!DxAltaMedica = Left(Trim(oDODiagnostico.CodigoCIE2004) & " " & Trim(oDODiagnostico.descripcion), 100)
                mrs_Reembolsos.Fields!NroReferenciaDestino = IIf(oRsBuscaXcta.RecordCount > 0, oRsBuscaXcta.Fields!NroReferenciaDestino, 0)
                mrs_Reembolsos.Fields!idAtencion = oRsBusqueda.Fields!idAtencion
                mrs_Reembolsos.Fields!idTipoServicio = oRsBusqueda.Fields!idTipoServicio
                mrs_Reembolsos.Fields!UltReembolsoPorCuenta = lcIdReembolsoPorCuenta
                mrs_Reembolsos.Update
                oRsBusqueda.MoveNext
           Loop
        End If
    End If
    'Consulta externa
    If cmbTipoServicio.ListIndex = 0 Or cmbTipoServicio.ListIndex = 3 Then
        Set oRsBusqueda = mo_ReglasFacturacion.AtencionesSeleccionarCEPorFechaIngresoYplan(CDate(txtFechaIni.Text), CDate(txtFechaFin.Text), Val(mo_cmbFuenteFinanciamiento.BoundText))
        oRsBusqueda.Filter = "idEstado<>4"
'        If Val(mo_cmbAreaTramitaR.BoundText) = 4 Then
'           'referidos/contraReferidos
'           oRsBusqueda.Filter = "NroReferenciaDestino<>null"
'        End If
        If oRsBusqueda.RecordCount > 0 Then
           oRsBusqueda.MoveFirst
           Do While Not oRsBusqueda.EOF
              Set oDODiagnostico = mo_ReglasFacturacion.DevuelveDxAltaMedica(oRsBusqueda.Fields!idAtencion, oRsBusqueda.Fields!idTipoServicio)
              Set oRsBuscaXcta = mo_AdminAdmision.AtencionesDatosAdicionalesSeleccionarPorIdCuenta(oRsBusqueda.Fields!idCuentaAtencion)
              lnConsumo = RetornaConsumoPorCuenta(oRsBusqueda.Fields!idCuentaAtencion, Val(mo_cmbAreaTramitaR.BoundText), Val(mo_cmbFuenteFinanciamiento.BoundText), Val(mo_cmbTipoConsumo.BoundText))
              mrs_Reembolsos.AddNew
              mrs_Reembolsos.Fields!seleccionar = False
              mrs_Reembolsos.Fields!nrocuenta = oRsBusqueda.Fields!idCuentaAtencion
              mrs_Reembolsos.Fields!EstadoCuenta = oRsBusqueda.Fields!dEstadoCuenta
              mrs_Reembolsos.Fields!NroHistoria = oRsBusqueda.Fields!NroHistoriaClinica
              mrs_Reembolsos.Fields!Paciente = oRsBusqueda.Fields!Paciente
              mrs_Reembolsos.Fields!Servicio = oRsBusqueda.Fields!DServicio
              mrs_Reembolsos.Fields!fAltaMedica = oRsBusqueda.Fields!FechaIngreso
              mrs_Reembolsos.Fields!consumo = lnConsumo
              mrs_Reembolsos.Fields!Reembolsado = 0
              mrs_Reembolsos.Fields!porReembolsar = 0
              mrs_Reembolsos.Fields!dxid = oDODiagnostico.idDiagnostico
              mrs_Reembolsos.Fields!DxAltaMedica = Left(Trim(oDODiagnostico.CodigoCIE2004) & " " & Trim(oDODiagnostico.descripcion), 100)
              mrs_Reembolsos.Fields!NroReferenciaDestino = IIf(oRsBuscaXcta.RecordCount > 0, oRsBuscaXcta.Fields!NroReferenciaDestino, 0)
              mrs_Reembolsos.Fields!idTipoServicio = 1
              mrs_Reembolsos.Fields!UltReembolsoPorCuenta = lcIdReembolsoPorCuenta
              mrs_Reembolsos.Update
              oRsBusqueda.MoveNext
           Loop
        End If
    End If
    SumaTotales True
    '
    Set oRsBusqueda = Nothing
    Set oRsBuscaXcta = Nothing
    '
    fraFiltro.Enabled = False
    FraReembolso.Enabled = True
End Sub


Function RetornaConsumoPorCuenta(lnIdCuentaAtencion As Long, lnAreaTramitaReembolso As Long, lnFuenteFinanciamiento As Long, lnTipoConsumo As Long) As Double
        Dim oRsBuscaXcta As New Recordset
        Dim lnReembolsosPagados As Double
        Dim lnAdelantos As Double
        Dim lnRetornaConsumoPorCuenta As Double
        Dim lnFarmacia As Double
        Dim oConexion As New Connection
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        If Val(mo_cmbTipoConsumo.BoundText) = 3 Then
           '****Busca Reembolsos anteriores: farmacia + Servicios + ambos
           Set oRsBuscaXcta = mo_ReglasFacturacion.FacturacionReembolsosSeleccionarPorCuentaPlanAreats(lnIdCuentaAtencion, lnAreaTramitaReembolso, lnFuenteFinanciamiento, 0)
        Else
           '***"Busca Reembolsos anteriores solo de Farmacia" ó "Busca Reembolsos anteriores solo de Servicios"
           Set oRsBuscaXcta = mo_ReglasFacturacion.FacturacionReembolsosSeleccionarPorCuentaPlanAreats(lnIdCuentaAtencion, lnAreaTramitaReembolso, lnFuenteFinanciamiento, lnTipoConsumo)
        End If
        lcIdReembolsoPorCuenta = ""
        lnAdelantos = mo_ReglasFacturacion.RetornaImporteDePagosAdelantadosPorNroCuenta(lnIdCuentaAtencion, oConexion)
        If oRsBuscaXcta.RecordCount > 0 Then
            '***************************segundo,tercero... Reembolso
            lnReembolsosPagados = 0
            oRsBuscaXcta.MoveFirst
            Do While Not oRsBuscaXcta.EOF
               lnReembolsosPagados = lnReembolsosPagados + oRsBuscaXcta.Fields!ReembolsoPagadoFarmacia + oRsBuscaXcta.Fields!ReembolsoPagadoServicio
               lcIdReembolsoPorCuenta = lcIdReembolsoPorCuenta & Trim(Str(oRsBuscaXcta.Fields!IdFactReembolso)) & ", "
               oRsBuscaXcta.MoveNext
            Loop
            lcIdReembolsoPorCuenta = Left(lcIdReembolsoPorCuenta, 100)
            Select Case Val(mo_cmbTipoConsumo.BoundText)
            Case 2  'Servicios
               lnRetornaConsumoPorCuenta = mo_ReglasFacturacion.RetornaConsumoPacienteServiciosConSeguroPorNroCuenta(lnIdCuentaAtencion)
               lnFarmacia = mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(lnIdCuentaAtencion)
               If lnAdelantos > lnFarmacia Then
                  lnAdelantos = lnAdelantos - lnFarmacia
               Else
                  lnAdelantos = 0
               End If
               RetornaConsumoPorCuenta = lnRetornaConsumoPorCuenta - lnAdelantos - lnReembolsosPagados
            Case 1  'Farmacia
               lnRetornaConsumoPorCuenta = mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(lnIdCuentaAtencion)
               If lnAdelantos > lnRetornaConsumoPorCuenta Then
                  lnRetornaConsumoPorCuenta = 0
               Else
                  lnRetornaConsumoPorCuenta = lnRetornaConsumoPorCuenta - lnAdelantos
               End If
               RetornaConsumoPorCuenta = lnRetornaConsumoPorCuenta - lnReembolsosPagados
            Case 3  'ambos
               RetornaConsumoPorCuenta = mo_ReglasFacturacion.RetornaConsumoFarmaciaServiciosPorNroCuenta(lnIdCuentaAtencion, oConexion) - lnReembolsosPagados
            End Select
        Else
            '***************************primer Reembolso
            Select Case Val(mo_cmbTipoConsumo.BoundText)
            Case 2  'Servicios
               lnRetornaConsumoPorCuenta = mo_ReglasFacturacion.RetornaConsumoPacienteServiciosConSeguroPorNroCuenta(lnIdCuentaAtencion, True)
               lnFarmacia = mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(lnIdCuentaAtencion)
               If lnAdelantos > lnFarmacia Then
                  lnAdelantos = lnAdelantos - lnFarmacia
               Else
                  lnAdelantos = 0
               End If
               RetornaConsumoPorCuenta = lnRetornaConsumoPorCuenta - lnAdelantos
            Case 1  'Farmacia
               lnRetornaConsumoPorCuenta = mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(lnIdCuentaAtencion)
               If lnAdelantos > lnRetornaConsumoPorCuenta Then
                  RetornaConsumoPorCuenta = 0
               Else
                  RetornaConsumoPorCuenta = lnRetornaConsumoPorCuenta - lnAdelantos
               End If
            Case 3  'ambos
               RetornaConsumoPorCuenta = mo_ReglasFacturacion.RetornaConsumoFarmaciaServiciosPorNroCuenta(lnIdCuentaAtencion, oConexion, True)
            End Select
        End If
        Set oRsBuscaXcta = Nothing
        oConexion.Close
        Set oConexion = Nothing
End Function



Private Sub btnRefrescar_Click()
    SumaTotales True
End Sub

Private Sub btnReImprime_Click()
    Dim oConsultaCta As New ReembolsosCta
    oConsultaCta.idCuentaAtencion = lnCuentaActualDelGrid
    oConsultaCta.GrabaConsumosConsolidados = True
    oConsultaCta.Show 1
    Set oConsultaCta = Nothing
    '
    ImprimeFactura
End Sub

Private Sub chkGrabaDefinitivamente_Click()
    If chkGrabaDefinitivamente.Value = 1 Then
       txtFcobro.Text = lcBuscaParametro.RetornaFechaHoraServidorSQL
    Else
       txtFcobro.Text = sighEntidades.FECHA_VACIA_DMY_HM
    End If
End Sub

Private Sub chkSinBuscarCtas_Click()
   If Me.chkSinBuscarCtas.Value = 1 Then
        If Me.cmbTipoConsumo.Text = "" Then
           Me.chkSinBuscarCtas.Value = 0
           MsgBox "Tiene que elejir el Tipo de Consumo", vbInformation, Me.Caption
           Me.cmbTipoConsumo.SetFocus
           Exit Sub
        End If
        btnBuscar.Visible = False
        LimpiarGrilla
        LimpiarDatos
        lc_FuenteFinanciamientoPermitidos = mo_cmbFuenteFinanciamiento.BoundText
        SumaTotales False
        fraFiltro.Enabled = False
        FraReembolso.Enabled = True
        If mo_cmbTipoConsumo.BoundText = "1" Then
           Me.chkIGV.Value = 1
        End If
   Else
        btnBuscar.Visible = True
   End If
End Sub

Private Sub chkTodos_Click()
    If mrs_Reembolsos.RecordCount > 0 Then
       mrs_Reembolsos.MoveFirst
       Do While Not mrs_Reembolsos.EOF
          If chkTodos.Value = 1 Then
             mrs_Reembolsos.Fields!seleccionar = True
             mrs_Reembolsos.Fields!porReembolsar = 0
             mrs_Reembolsos.Fields!Reembolsado = mrs_Reembolsos.Fields!consumo
          Else
             mrs_Reembolsos.Fields!seleccionar = False
             mrs_Reembolsos.Fields!porReembolsar = mrs_Reembolsos.Fields!consumo
             mrs_Reembolsos.Fields!Reembolsado = 0
          End If
          mrs_Reembolsos.Update
          mrs_Reembolsos.MoveNext
       Loop
       SumaTotales True
    End If
    
End Sub



Private Sub cmbAnio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbAnio
   AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbAreaTramitaR_Click()
       mo_cmbFuenteFinanciamiento.ListField = "Descripcion"
       mo_cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"
       If Val(mo_cmbAreaTramitaR.BoundText) = 4 Then
          '**************Referencia************
          Set mo_cmbFuenteFinanciamiento.RowSource = mo_ReglasFacturacion.FuentesFinanciamientoDevuelveTodosSegunFiltro(" utilizadoEn=3 ")
       Else
          Set mo_cmbFuenteFinanciamiento.RowSource = mo_ReglasFacturacion.FuentesFinanciamientoDevuelveTodosSegunFiltro(" idAreaTramitaSeguros= " & mo_cmbAreaTramitaR.BoundText)
       End If

End Sub

Private Sub cmbAreaTramitaR_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbAreaTramitaR
   AdministrarKeyPreview KeyCode

End Sub



Private Sub cmbFuenteFinanciamiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbFuenteFinanciamiento
   AdministrarKeyPreview KeyCode

End Sub




Private Sub cmbIdCaja_Click()
    Dim lIdTipoComprobante As Long
    Dim oCajaNroDocumento As New DOCajaNroDocumento
    txtNroSerie.Text = ""
    txtNroDocumento.Text = ""
    lIdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
    If lIdTipoComprobante > 0 Then
        Set oCajaNroDocumento = mo_AdminCaja.NroDocumentoSeleccionarPorIdCajaYTipoComprobante(Val(mo_cmbIdCaja.BoundText), _
                                                                                              lIdTipoComprobante)
        txtNroSerie.Text = Trim(oCajaNroDocumento.nroSerie)
        txtNroDocumento.Text = Right("00000" & Trim(oCajaNroDocumento.nrodocumento), 8)
        'ms_MensajeError = Right("00000" & Trim(oCajaNroDocumento.nrodocumento), 8)
        lblDescripcionLarga.Visible = DevuelveDescripcionLargaDeCaja(mo_cmbIdCaja.BoundText)
        '
    End If
    Set oCajaNroDocumento = Nothing
    
End Sub

Function DevuelveDescripcionLargaDeCaja(lnIdCaja As Long) As Boolean
        Dim oRsTmp9 As New Recordset
        DevuelveDescripcionLargaDeCaja = False
        Set oRsTmp9 = mo_AdminCaja.CajaCajaSegunFiltro("idCaja=" & Trim(Str(lnIdCaja)))
        If oRsTmp9.RecordCount > 0 Then
           If Not IsNull(oRsTmp9!idPartida) Then
              If oRsTmp9!idPartida > 0 Then
                 DevuelveDescripcionLargaDeCaja = True
              End If
           End If
        End If
        oRsTmp9.Close
        Set oRsTmp9 = Nothing
End Function

Private Sub cmbIdTipoComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoComprobante
   AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbMes_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbMes
   AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbTipoServicio
   AdministrarKeyPreview KeyCode

End Sub

Private Sub cmdAgregar_Click()
    FraNcuenta.Visible = True
    txtNcuenta.Text = ""
    txtNcuenta.SetFocus
End Sub

Private Sub cmdBuscaCuentaPorApellidos_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New doPaciente
    Dim lbNuevo As Boolean
    Dim lnConsumo As Double
    Dim oDODiagnostico As New DODiagnostico
    Dim oRsBuscaXcta As New Recordset
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            Dim oRsTmp As New Recordset
            Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(oBusqueda.idRegistroSeleccionado, oConexion, True)
            oRsTmp.Filter = "idEstadoAtencion<>0"
            If oRsTmp.RecordCount > 0 Then
                If ValidaElAgregarCuenta(oRsTmp) = True Then
                    Set oDODiagnostico = mo_ReglasFacturacion.DevuelveDxAltaMedica(oRsTmp.Fields!idAtencion, oRsTmp.Fields!idTipoServicio)
                    Set oRsBuscaXcta = mo_ReglasFacturacion.FacturacionReembolsosSeleccionarPorCuentaPlanAreats(oRsTmp.Fields!idCuentaAtencion, Val(mo_cmbAreaTramitaR.BoundText), Val(mo_cmbFuenteFinanciamiento.BoundText), Val(mo_cmbTipoConsumo.BoundText))
                    If oRsBuscaXcta.RecordCount > 0 Then
                       lnConsumo = oRsBuscaXcta.Fields!ReembolsoPorPagar
                    Else
                       lnConsumo = mo_ReglasFacturacion.RetornaConsumoFarmaciaServiciosPorNroCuenta(oRsTmp.Fields!idCuentaAtencion, oConexion, True)
                    End If
                    mrs_Reembolsos.AddNew
                    mrs_Reembolsos.Fields!seleccionar = True
                    mrs_Reembolsos.Fields!nrocuenta = oRsTmp.Fields!idCuentaAtencion
                    mrs_Reembolsos.Fields!EstadoCuenta = oRsTmp.Fields!dEstadoCuenta
                    mrs_Reembolsos.Fields!NroHistoria = oDOPaciente.NroHistoriaClinica
                    mrs_Reembolsos.Fields!Paciente = Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + oDOPaciente.PrimerNombre
                    mrs_Reembolsos.Fields!Servicio = oRsTmp.Fields!dTipoServicio
                    If (oRsTmp.Fields!idTipoServicio <> 2 And oRsTmp.Fields!idTipoServicio <> 3 And oRsTmp.Fields!idTipoServicio <> 4) Or oRsTmp.Fields!EsPacienteExterno = True Then
                       mrs_Reembolsos.Fields!fAltaMedica = oRsTmp.Fields!FechaIngreso
                    Else
                       mrs_Reembolsos.Fields!fAltaMedica = oRsTmp.Fields!fechaEgreso
                    End If
                    mrs_Reembolsos.Fields!consumo = lnConsumo
                    mrs_Reembolsos.Fields!Reembolsado = 0
                    mrs_Reembolsos.Fields!porReembolsar = 0
                    mrs_Reembolsos.Fields!dxid = oDODiagnostico.idDiagnostico
                    mrs_Reembolsos.Fields!DxAltaMedica = Left(Trim(oDODiagnostico.CodigoCIE2004) & " " & Trim(oDODiagnostico.descripcion), 100)
                    mrs_Reembolsos.Fields!NroReferenciaDestino = oRsTmp.Fields!NroReferenciaDestino
                    mrs_Reembolsos.Fields!idAtencion = oRsTmp.Fields!idAtencion
                    mrs_Reembolsos.Update
                    SumaTotales False
                End If
                FraNcuenta.Visible = False
            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
        End If
    End If
    Set oDOPaciente = Nothing
    Set oBusqueda = Nothing
    Set oRsBuscaXcta = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub cmdEliminaLista_Click()
    On Error Resume Next
    If mrs_Reembolsos.RecordCount > 0 Then
        If MsgBox("Esta seguro de Eliminar Cuentas de la Lista", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
           mrs_Reembolsos.MoveFirst
           Do While Not mrs_Reembolsos.EOF
              mrs_Reembolsos.Delete
              mrs_Reembolsos.Update
              mrs_Reembolsos.MoveNext
           Loop
           SumaTotales True
        End If
    End If
End Sub



Private Sub cmdProrratea_Click()
    If lnTotalConsumoParaProrrateo <= 0 Then
       MsgBox "El total de la columna CONSUMOS deberá ser  mayor a CERO", vbInformation, Me.Caption
       Exit Sub
    End If
    Dim lnTreembolso As Double, lnPorc As Double, lnTconsumo As Double, lnTconsumo1 As Double
    lnTreembolso = Val(txtTreembolso.Text)
    If mrs_Reembolsos.RecordCount > 0 And lnTreembolso > 0 Then
       'lnTconsumo = Val(txtTconsumo.Text) + Val(Me.txtTporReembolsar.Text)
       lnTconsumo = lnTotalConsumoParaProrrateo
       If lnTreembolso >= lnTconsumo Then
          lnPorc = Round((lnTreembolso * 100 - lnTconsumo * 100) / lnTconsumo, 2)
       Else
          lnPorc = 100 - Round(lnTreembolso * 100 / lnTconsumo, 2)
       End If
       mrs_Reembolsos.MoveFirst
       Do While Not mrs_Reembolsos.EOF
          If mrs_Reembolsos.Fields!seleccionar = True Then
             lnTconsumo1 = mrs_Reembolsos.Fields!consumo + mrs_Reembolsos.Fields!porReembolsar
             If lnTreembolso >= lnTconsumo Then
                 mrs_Reembolsos.Fields!Reembolsado = lnTconsumo1 + Round(lnTconsumo1 * lnPorc / 100, 2)
             Else
                 mrs_Reembolsos.Fields!Reembolsado = lnTconsumo1 - Round(lnTconsumo1 * lnPorc / 100, 2)
             End If
          Else
             mrs_Reembolsos.Fields!Reembolsado = 0
          End If
          mrs_Reembolsos.Update
          mrs_Reembolsos.MoveNext
       Loop
       SumaTotales True
    End If
End Sub





Private Sub Form_Activate()
       If ml_SoloSeIngresaUnaCuenta = True Then
          ml_SoloSeIngresaUnaCuenta = False
          chkSinBuscarCtas.Value = 1
          chkSinBuscarCtas_Click
          cmdAgregar_Click
       End If
End Sub

Private Sub Form_Initialize()
    Set mo_cmbFuenteFinanciamiento.MiComboBox = cmbFuenteFinanciamiento
    Set mo_cmbAreaTramitaR.MiComboBox = cmbAreaTramitaR
    Set mo_cmbIdCaja.MiComboBox = cmbIdCaja
    Set mo_cmbIdTurno.MiComboBox = cmbIdTurno
    Set mo_cmbIdTipoComprobante.MiComboBox = cmbIdTipoComprobante
    Set mo_cmbTipoConsumo.MiComboBox = cmbTipoConsumo
End Sub

Sub LicenciaSunat()
        Dim lcMensajeLicencia As String
        lbTieneLicenciaParaNotaCreditoYsunat = True

End Sub
Private Sub Form_Load()
       LicenciaSunat
       lbTieneGrabadoFechaCobro = False
       InicilizarParametros
       GenerarRecordsetTemporal
       Select Case mi_Opcion
       Case sghAgregar
           Me.txtFemision.Text = lcBuscaParametro.RetornaFechaHoraServidorSQL
           Me.Caption = "Agregar Reembolso"
           Me.lblMotivoAnulacion.Visible = False
           Me.txtMotivoAnulacion.Visible = False
       Case sghModificar
           Me.Caption = "Modificar Reembolso"
           Me.lblMotivoAnulacion.Enabled = False
           Me.txtMotivoAnulacion.Enabled = False
       Case sghConsultar
           Me.Caption = "Consultar Reembolso"
       Case sghEliminar
           Me.Caption = "Eliminar Reembolso"
           Me.lblMotivoAnulacion.Enabled = True
           Me.txtMotivoAnulacion.Enabled = True
       End Select
       CargarComboBoxes
       CargarDatosAlFormulario
       
       
End Sub

Sub AsignaMesAnioAlComboSegunRangoFechas()
    Dim lnMes1 As Integer, lnAnio1 As Integer, lnFor As Integer
    lnAnio1 = Year(CDate(Me.txtFechaFin.Text))
    lnMes1 = Month(CDate(Me.txtFechaFin.Text))
    cmbMes.ListIndex = lnMes1 - 1
    For lnFor = 0 To cmbAnio.ListCount - 1
        cmbAnio.ListIndex = lnFor
        If Val(cmbAnio.Text) = lnAnio1 Then
           Exit For
        End If
    Next
End Sub

Sub CargarDatosAlFormulario()
    Me.txtFechaIni.Text = sighEntidades.PrimerFechaDDMMYYDelMesAnterior
    Me.txtFechaFin.Text = sighEntidades.UltimaFechaDDMMYYDelMesAnterior
    '
    mo_Formulario.LlenaComboConAnios cmbAnio
    mo_Formulario.LlenaComboConMeses cmbMes
    '
    AsignaMesAnioAlComboSegunRangoFechas
    '
    mo_Formulario.HabilitarDeshabilitar txtFcobro, False
    mo_Formulario.HabilitarDeshabilitar txtFemision, False
    mo_Formulario.HabilitarDeshabilitar txtTporReembolsar, False
    mo_Formulario.HabilitarDeshabilitar txtTconsumo, False
    mo_Formulario.HabilitarDeshabilitar txtSaldoFinal, False
    mo_Formulario.HabilitarDeshabilitar txtTreembolso, False

    Select Case mi_Opcion
     Case sghAgregar
         
     Case sghModificar
         CargarDatosAlosControles
     Case sghConsultar
         CargarDatosAlosControles
     Case sghEliminar
         CargarDatosAlosControles
     End Select
End Sub

Sub CargarDatosAlosControles()
    btnReImprime.Visible = True
    Dim oRsTmp As New Recordset
    fraFiltro.Enabled = False
    Me.FraReembolso.Enabled = True
    lblFAltaMedica.Visible = False
    txtFechaIni.Visible = False
    lblHasta.Visible = False
    txtFechaFin.Visible = False
    '
    Set mo_DoFactReembolsos = mo_ReglasFacturacion.FactReembolsosSeleccionarPorId(ml_IdFactReembolso)
    With mo_DoFactReembolsos
        mo_cmbAreaTramitaR.BoundText = .idAreaTramitaSeguro
        mo_cmbFuenteFinanciamiento.BoundText = .IdFuenteFinanciamiento
        txtDescripcionR.Text = .descripcion
        txtDctoHospital.Text = .Documentos
        cmbAnio.Text = .Anio
        cmbMes.ListIndex = .Mes - 1
        txtTconsumo.Text = Format(.ConsumoPorReembolsar, "####,###,##0.00")
        txtTporReembolsar.Text = Format(.ReembolsoPorPagar, "####,###,##0.00")
        txtTreembolso.Text = Format(.ReembolsoPagado, "####,###,##0.00")
        txtSaldoInicial.Text = Format(.SaldoInicial, "####,###,##0.00")
        txtSaldoFinal.Text = Format(.SaldoFinal, "####,###,##0.00")
        mo_cmbTipoConsumo.BoundText = .idTipoConsumo
        lc_FuenteFinanciamientoPermitidos = mo_cmbFuenteFinanciamiento.BoundText
        chkGrabaDefinitivamente.Value = IIf(.GrabaDefinitivamente = True, 1, 0)
        lbYaTieneFactura = .GrabaDefinitivamente
    End With
    '
    Set mo_DoFactReembolsosDcto = mo_ReglasFacturacion.FactReembolsosDocumentosSeleccionarPorId(ml_IdFactReembolso)
    lnIdComprobantePagoActual = IIf(IsNull(mo_DoFactReembolsosDcto.IdComprobantePago), 0, mo_DoFactReembolsosDcto.IdComprobantePago)
    Me.txtMotivoAnulacion.Text = IIf(IsNull(mo_DoFactReembolsosDcto.MotivoAnulacion), "", mo_DoFactReembolsosDcto.MotivoAnulacion)
    'Hospitalizacion/emergencia
    Set oRsTmp = mo_ReglasFacturacion.FacturacionReembolsosSeleccionarHospEmergPorIdFactReembolso(ml_IdFactReembolso)
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
              mrs_Reembolsos.AddNew
              If oRsTmp.Fields!ReembolsoPagadoFarmacia + oRsTmp.Fields!ReembolsoPagadoServicio > 0 Then
                 mrs_Reembolsos.Fields!seleccionar = True
              Else
                 mrs_Reembolsos.Fields!seleccionar = False
              End If
              mrs_Reembolsos.Fields!nrocuenta = oRsTmp.Fields!idCuentaAtencion
              mrs_Reembolsos.Fields!EstadoCuenta = oRsTmp.Fields!dEstadoCuenta
              mrs_Reembolsos.Fields!NroHistoria = oRsTmp.Fields!NroHistoriaClinica
              mrs_Reembolsos.Fields!Paciente = oRsTmp.Fields!Paciente
              mrs_Reembolsos.Fields!Servicio = IIf(IsNull(oRsTmp.Fields!DServicio), "", oRsTmp.Fields!DServicio)
              If Not IsNull(oRsTmp.Fields!fechaEgreso) Then
                 mrs_Reembolsos.Fields!fAltaMedica = oRsTmp.Fields!fechaEgreso
              End If
              mrs_Reembolsos.Fields!consumo = oRsTmp.Fields!ConsumoPorReembolsar
              mrs_Reembolsos.Fields!Reembolsado = oRsTmp.Fields!ReembolsoPagadoFarmacia + oRsTmp.Fields!ReembolsoPagadoServicio
              mrs_Reembolsos.Fields!porReembolsar = oRsTmp.Fields!ReembolsoPorPagar
              If Not IsNull(oRsTmp.Fields!idDiagnostico) Then
                 mrs_Reembolsos.Fields!DxAltaMedica = Left(Trim(oRsTmp.Fields!CodigoCIE2004) & " " & oRsTmp.Fields!DDx, 100)
                 mrs_Reembolsos.Fields!dxid = oRsTmp.Fields!idDiagnostico
              End If
              If Not IsNull(oRsTmp.Fields!NroReferenciaDestino) Then
                 mrs_Reembolsos.Fields!NroReferenciaDestino = oRsTmp.Fields!NroReferenciaDestino
              End If
              mrs_Reembolsos.Fields!idAtencion = oRsTmp.Fields!idAtencion
              mrs_Reembolsos.Fields!idTipoServicio = oRsTmp.Fields!idTipoServicio
              mrs_Reembolsos.Fields!UltReembolsoPorCuenta = oRsTmp.Fields!IdReembolsosAnteriores
              mrs_Reembolsos.Update
              oRsTmp.MoveNext
       Loop
    End If
    'CE
    Set oRsTmp = mo_ReglasFacturacion.FacturacionReembolsosSeleccionarCEPorIdFactReembolso(ml_IdFactReembolso)
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
              mrs_Reembolsos.AddNew
              If oRsTmp.Fields!ReembolsoPorPagar > 0 Then
                 mrs_Reembolsos.Fields!seleccionar = False
              Else
                 mrs_Reembolsos.Fields!seleccionar = True
              End If
              mrs_Reembolsos.Fields!nrocuenta = oRsTmp.Fields!idCuentaAtencion
              mrs_Reembolsos.Fields!EstadoCuenta = oRsTmp.Fields!dEstadoCuenta
              mrs_Reembolsos.Fields!NroHistoria = IIf(IsNull(oRsTmp.Fields!NroHistoriaClinica), 0, oRsTmp.Fields!NroHistoriaClinica) 'Adams
              mrs_Reembolsos.Fields!Paciente = oRsTmp.Fields!Paciente
              mrs_Reembolsos.Fields!Servicio = oRsTmp.Fields!DServicio
              mrs_Reembolsos.Fields!fAltaMedica = oRsTmp.Fields!FechaIngreso
              mrs_Reembolsos.Fields!consumo = oRsTmp.Fields!ConsumoPorReembolsar
              mrs_Reembolsos.Fields!Reembolsado = IIf(IsNull(oRsTmp.Fields!ReembolsoPagadoFarmacia), 0, oRsTmp.Fields!ReembolsoPagadoFarmacia) + IIf(IsNull(oRsTmp.Fields!ReembolsoPagadoServicio), 0, oRsTmp.Fields!ReembolsoPagadoServicio)
              mrs_Reembolsos.Fields!porReembolsar = oRsTmp.Fields!ReembolsoPorPagar
              If Not IsNull(oRsTmp.Fields!idDiagnostico) Then
                 mrs_Reembolsos.Fields!DxAltaMedica = Left(Trim(oRsTmp.Fields!CodigoCIE2004) & " " & oRsTmp.Fields!DDx, 100)
                 mrs_Reembolsos.Fields!dxid = oRsTmp.Fields!idDiagnostico
              End If
              If Not IsNull(oRsTmp.Fields!NroReferenciaDestino) Then
                 mrs_Reembolsos.Fields!NroReferenciaDestino = oRsTmp.Fields!NroReferenciaDestino
              End If
              mrs_Reembolsos.Fields!idAtencion = oRsTmp.Fields!idAtencion
              mrs_Reembolsos.Fields!idTipoServicio = oRsTmp.Fields!idTipoServicio
              mrs_Reembolsos.Fields!UltReembolsoPorCuenta = oRsTmp.Fields!IdReembolsosAnteriores
              mrs_Reembolsos.Update
              oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    '
'    If chkGrabaDefinitivamente.Value = 1 Then
        If lnIdComprobantePagoActual > 0 Then
            If chkGrabaDefinitivamente.Value = 0 Then
                If mo_AdminCaja.CajaComprobantesPagoActualizaFechaCobranza(lnIdComprobantePagoActual, 0) = True Then
                End If
            End If
            Set oRsTmp = mo_AdminCaja.CajaComprobantesSeleccionarPorId(lnIdComprobantePagoActual)
            If oRsTmp.RecordCount > 0 Then
               If Not IsNull(oRsTmp!FechaCobranza) Then
                    txtFcobro.Text = Format(oRsTmp!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY_HM)
                    lbTieneGrabadoFechaCobro = True
               End If
               Me.txtFemision.Text = Format(oRsTmp!fechaEmision, sighEntidades.DevuelveFechaSoloFormato_DMY_HM)

               
               mo_cmbIdTurno.BoundText = oRsTmp.Fields!IdTurno
               mo_cmbIdCaja.BoundText = oRsTmp.Fields!IdCaja
               mo_cmbIdTipoComprobante.BoundText = oRsTmp.Fields!IdTipoComprobante
               txtNroSerie.Text = oRsTmp.Fields!nroSerie
               txtNroDocumento.Text = oRsTmp.Fields!nrodocumento
               txtRuc.Text = IIf(IsNull(oRsTmp.Fields!ruc), "", oRsTmp.Fields!ruc)
               Me.txtRazonSocial.Text = IIf(IsNull(oRsTmp.Fields!razonSocial), "", oRsTmp.Fields!razonSocial)
               lcNroSerieActual = Trim(oRsTmp.Fields!nroSerie)
               lcNroDocumentoActual = Trim(oRsTmp.Fields!nrodocumento)
               lnIdTipoComprobanteActual = oRsTmp.Fields!IdTipoComprobante
               If oRsTmp!IGV > 0 Then
                  chkIGV.Value = 1
               End If
               chkIGV.Enabled = False
            End If
        End If
'    Else
'        If Val(mo_DoFactReembolsosDcto.nrodocumento) > 0 Then
'            txtNroDocumento.Text = mo_DoFactReembolsosDcto.nrodocumento
'            txtNroSerie.Text = mo_DoFactReembolsosDcto.nroSerie
'            mo_cmbIdTipoComprobante.BoundText = mo_DoFactReembolsos.IdTipoComprobante
'        End If
'        Me.txtMotivoAnulacion.Text = mo_DoFactReembolsosDcto.MotivoAnulacion
'    End If
    '
    Set oRsTmp = Nothing
    If Val(mo_cmbAreaTramitaR.BoundText) = 4 Then
        '**************Referencia************
        grdReembolsos.Bands(0).Columns("Consumo").Activation = ssActivationAllowEdit
        grdReembolsos.Bands(0).Columns("Consumo").Header.Appearance.ForeColor = vbWhite
        grdReembolsos.Bands(0).Columns("Consumo").Header.Appearance.BackColor = vbRed
    End If
    SumaTotales False
    If mo_DoFactReembolsos.idEstadoReembolso = 0 Then
       Me.btnAceptar.Enabled = False
    ElseIf mi_Opcion = sghConsultar Then
       Me.btnAceptar.Enabled = False
    ElseIf chkGrabaDefinitivamente.Value = 1 And mi_Opcion <> sghEliminar Then
       Me.btnAceptar.Enabled = False
    End If
    If mo_cmbIdCaja.BoundText <> "" Then
       lblDescripcionLarga.Visible = DevuelveDescripcionLargaDeCaja(mo_cmbIdCaja.BoundText)
    End If
    CargaDatosDelProveedor
End Sub

Sub CargarComboBoxes()
       Dim oRsTmp As New Recordset
       '
       mo_cmbAreaTramitaR.ListField = "Descripcion"
       mo_cmbAreaTramitaR.BoundColumn = "idAreaTramitaSeguros"
       Set mo_cmbAreaTramitaR.RowSource = mo_ReglasFacturacion.AreaTramitaSegurosDevuelveTodosSegunFiltro("")
       Dim rsPuntoCargaDondeLabora As Recordset
       Set rsPuntoCargaDondeLabora = mo_ReglasComunes.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAreaTramitaSeguros, ml_idUsuario)
       If rsPuntoCargaDondeLabora.RecordCount > 0 Then
           mo_cmbAreaTramitaR.BoundText = rsPuntoCargaDondeLabora.Fields!idLaboraSubArea
           mo_Formulario.HabilitarDeshabilitar cmbAreaTramitaR, False
           cmbAreaTramitaR_Click
       End If
       '
       mo_cmbIdTurno.BoundColumn = "IdTurno"
       mo_cmbIdTurno.ListField = "Descripcion"
       Set mo_cmbIdTurno.RowSource = mo_AdminCaja.TurnosSeleccionarTodosParaLista()
       mo_cmbIdTurno.BoundText = "1"  'general
       '
       mo_cmbIdCaja.BoundColumn = "IdCaja"
       mo_cmbIdCaja.ListField = "Descripcion"
       Set mo_cmbIdCaja.RowSource = mo_AdminCaja.CajaSeleccionarTodosParaLista()
       Set oRsTmp = mo_AdminCaja.CajaCajaSeleccionarPorNombrePC(mo_lcNombrePc)
       If oRsTmp.RecordCount > 0 Then
           mo_cmbIdCaja.BoundText = oRsTmp.Fields!IdCaja
           cmbIdCaja.Enabled = False
       End If
       '
       mo_cmbIdTipoComprobante.BoundColumn = "IdTipoComprobante"
       mo_cmbIdTipoComprobante.ListField = "Descripcion"
       Set mo_cmbIdTipoComprobante.RowSource = mo_AdminCaja.TiposComprobanteSeleccionarTodos()
       mo_cmbIdTipoComprobante.BoundText = "2"  'factura
      ' mo_Formulario.HabilitarDeshabilitar cmbIdTipoComprobante, False
       '
       mo_cmbTipoConsumo.ListField = "Descripcion"
       mo_cmbTipoConsumo.BoundColumn = "idTipoConsumo"
       Set mo_cmbTipoConsumo.RowSource = mo_ReglasFacturacion.TiposConsumosDevuelveTodos("")
       mo_cmbTipoConsumo.BoundText = "3"
       '
       cmbTipoServicio.ListIndex = 3
       '
       Set oRsTmp = Nothing
End Sub

Sub GenerarRecordsetTemporal()
    With mrs_Reembolsos
          .Fields.Append "Seleccionar", adBoolean
          .Fields.Append "NroCuenta", adInteger
          .Fields.Append "EstadoCuenta", adVarChar, 30
          .Fields.Append "NroHistoria", adInteger
          .Fields.Append "Paciente", adVarChar, 160
          .Fields.Append "Servicio", adVarChar, 100
          .Fields.Append "FAltaMedica", adDate
          .Fields.Append "UltReembolsoPorCuenta", adVarChar, 100, adFldIsNullable
          .Fields.Append "Consumo", adDouble
          .Fields.Append "PorReembolsar", adDouble
          .Fields.Append "Reembolsado", adDouble
          .Fields.Append "DxAltaMedica", adVarChar, 100, adFldIsNullable
          .Fields.Append "DxId", adInteger
          .Fields.Append "IdAtencion", adInteger
          .Fields.Append "NroReferenciaDestino", adVarChar, 20, adFldIsNullable
          .Fields.Append "IdTipoServicio", adInteger
          .Fields.Append "adelantosF", adDouble                'solo se usará dentro de la Funcion al Agregar/modificar
          .Fields.Append "ReembolsosPagadosF", adDouble        'solo se usará dentro de la Funcion al Agregar/modificar
          .Fields.Append "ConsumoF", adDouble                  'solo se usará dentro de la Funcion al Agregar/modificar
          .Fields.Append "ReembolsoPagadoFarmacia", adDouble                  'solo se usará dentro de la Funcion al Agregar/modificar
          .Fields.Append "ReembolsoPagadoServicio", adDouble                  'solo se usará dentro de la Funcion al Agregar/modificar
          .Fields.Append "ConsumoS", adDouble                  'solo se usará dentro de la Funcion al Agregar/modificar
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdReembolsos.DataSource = mrs_Reembolsos
    mo_Apariencia.ConfigurarFilasBiColores Me.grdReembolsos, sighEntidades.GrillaConFilasBicolor
End Sub



Private Sub grdReembolsos_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
'    SumaTotales True
End Sub

Private Sub grdReembolsos_AfterRowActivate()
    Dim rsRecordset As ADODB.Recordset
    lnCuentaActualDelGrid = -1
    Set rsRecordset = grdReembolsos.DataSource
    On Error Resume Next
    lnCuentaActualDelGrid = rsRecordset.Fields!nrocuenta
End Sub

Private Sub grdReembolsos_AfterRowsDeleted()
'    SumaTotales True
End Sub

Private Sub grdReembolsos_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
'    On Error Resume Next
'    Dim oRow As SSRow
'    Set oRow = grdReembolsos.ActiveCell.Row
'    Select Case grdReembolsos.ActiveCell.Column.Key
'    Case "Seleccionar"
'        If oRow.Cells("Seleccionar").Value = True Then
'           mrs_Reembolsos.Fields!porReembolsar = 0
'           mrs_Reembolsos.Fields!Reembolsado = mrs_Reembolsos.Fields!consumo
'        Else
'           mrs_Reembolsos.Fields!porReembolsar = mrs_Reembolsos.Fields!consumo
'           mrs_Reembolsos.Fields!Reembolsado = 0
'        End If
'        mrs_Reembolsos.Update
'    Case "PorReembolsar"
'        mrs_Reembolsos.Fields!Reembolsado = mrs_Reembolsos.Fields!consumo - mrs_Reembolsos.Fields!porReembolsar
'        mrs_Reembolsos.Fields!seleccionar = IIf(mrs_Reembolsos.Fields!Reembolsado > 0, True, False)
'        mrs_Reembolsos.Update
'    Case "Reembolsado"
'        mrs_Reembolsos.Fields!porReembolsar = mrs_Reembolsos.Fields!consumo - mrs_Reembolsos.Fields!Reembolsado
'        mrs_Reembolsos.Fields!seleccionar = IIf(mrs_Reembolsos.Fields!Reembolsado > 0, True, False)
'        mrs_Reembolsos.Update
'    End Select
    
End Sub






Private Sub grdReembolsos_Click()
'            On Error Resume Next
'            Select Case grdReembolsos.ActiveCell.Column.Key
'            Case "Seleccionar"
'                SendKeys "{Tab}"
'            End Select

End Sub

Private Sub grdReembolsos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdReembolsos.Bands(0).Columns("Seleccionar").Width = 500
    grdReembolsos.Bands(0).Columns("NroCuenta").Width = 900
    grdReembolsos.Bands(0).Columns("NroCuenta").Activation = ssActivationActivateNoEdit
    grdReembolsos.Bands(0).Columns("EstadoCuenta").Width = 1500
    grdReembolsos.Bands(0).Columns("EstadoCuenta").Activation = ssActivationActivateNoEdit
    grdReembolsos.Bands(0).Columns("NroHistoria").Width = 700
    grdReembolsos.Bands(0).Columns("NroHistoria").Activation = ssActivationActivateNoEdit
    grdReembolsos.Bands(0).Columns("Paciente").Width = 2500
    grdReembolsos.Bands(0).Columns("Paciente").Activation = ssActivationActivateNoEdit
    grdReembolsos.Bands(0).Columns("Servicio").Width = 1500
    grdReembolsos.Bands(0).Columns("Servicio").Activation = ssActivationActivateNoEdit
    grdReembolsos.Bands(0).Columns("FAltaMedica").Width = 1000
    grdReembolsos.Bands(0).Columns("FAltaMedica").Activation = ssActivationActivateNoEdit
    grdReembolsos.Bands(0).Columns("FAltaMedica").Format = "dd/mm/yyyy"
    grdReembolsos.Bands(0).Columns("Consumo").Width = 1300
    grdReembolsos.Bands(0).Columns("Consumo").Activation = ssActivationActivateNoEdit
    grdReembolsos.Bands(0).Columns("Consumo").Format = "#0.000"
    grdReembolsos.Bands(0).Columns("Reembolsado").Width = 1300
    grdReembolsos.Bands(0).Columns("Reembolsado").Format = "#0.000"
    grdReembolsos.Bands(0).Columns("Reembolsado").Header.Appearance.ForeColor = vbWhite
    grdReembolsos.Bands(0).Columns("Reembolsado").Header.Appearance.BackColor = vbRed
    grdReembolsos.Bands(0).Columns("PorReembolsar").Width = 1300
    grdReembolsos.Bands(0).Columns("PorReembolsar").Format = "#0.000"
    grdReembolsos.Bands(0).Columns("PorReembolsar").Header.Appearance.ForeColor = vbWhite
    grdReembolsos.Bands(0).Columns("PorReembolsar").Header.Appearance.BackColor = vbRed
    grdReembolsos.Bands(0).Columns("DxAltaMedica").Width = 11300
    grdReembolsos.Bands(0).Columns("DxAltaMedica").Activation = ssActivationActivateNoEdit
    grdReembolsos.Bands(0).Columns("DxId").Hidden = True
    '
    grdReembolsos.Bands(0).Columns("Consumo").Activation = ssActivationAllowEdit
    grdReembolsos.Bands(0).Columns("Consumo").Header.Appearance.ForeColor = vbWhite
    grdReembolsos.Bands(0).Columns("Consumo").Header.Appearance.BackColor = vbRed
    grdReembolsos.Bands(0).Columns("ultReembolsoPorCuenta").Width = 800
    grdReembolsos.Bands(0).Columns("ultReembolsoPorCuenta").Activation = ssActivationActivateNoEdit
    '
    grdReembolsos.Bands(0).Columns("adelantosF").Hidden = True
    grdReembolsos.Bands(0).Columns("ReembolsosPagadosF").Hidden = True
    grdReembolsos.Bands(0).Columns("ConsumoF").Hidden = True
    grdReembolsos.Bands(0).Columns("ConsumoS").Hidden = True

End Sub


Sub LimpiarGrilla()
        If mrs_Reembolsos Is Nothing Then
            Exit Sub
        End If
        If mrs_Reembolsos.RecordCount > 0 Then
            mrs_Reembolsos.MoveFirst
            Do While Not mrs_Reembolsos.EOF
                mrs_Reembolsos.Delete
                mrs_Reembolsos.Update
                mrs_Reembolsos.MoveNext
            Loop
        End If
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Sub SumaTotales(lbSumaColumnaREEMBOLSADO As Boolean)
    Dim lnTporReembolsar As Double, lnTconsumo As Double, lnSaldofinal As Double
    Dim lnTReembolsado As Double
    On Error Resume Next
    lnTporReembolsar = 0: lnTconsumo = 0: lnSaldofinal = 0: lnTReembolsado = 0
    lnTotalConsumoParaProrrateo = 0
    If mrs_Reembolsos.RecordCount > 0 Then
       mrs_Reembolsos.MoveFirst
       Do While Not mrs_Reembolsos.EOF
          mrs_Reembolsos!porReembolsar = mrs_Reembolsos!consumo - mrs_Reembolsos.Fields!Reembolsado
          
          lnTporReembolsar = lnTporReembolsar + mrs_Reembolsos.Fields!porReembolsar
          lnTconsumo = lnTconsumo + mrs_Reembolsos.Fields!consumo
          lnTReembolsado = lnTReembolsado + mrs_Reembolsos.Fields!Reembolsado
          If mrs_Reembolsos.Fields!seleccionar = True Then
             lnTotalConsumoParaProrrateo = lnTotalConsumoParaProrrateo + mrs_Reembolsos.Fields!consumo
          Else
             'lnTotalConsumoParaProrrateo = lnTotalConsumoParaProrrateo - mrs_Reembolsos.Fields!Consumo
          End If
          mrs_Reembolsos.MoveNext
       Loop
    End If
    If lbSumaColumnaREEMBOLSADO = True Then
       Me.txtTreembolso.Text = Format(lnTReembolsado, "####,###,##0.00")
    End If
    txtTporReembolsar.Text = Format(lnTporReembolsar, "####,###,##0.00")
    txtTconsumo.Text = Format(lnTconsumo, "####,###,##0.00")
    txtSaldoFinal.Text = Format(Val(txtSaldoInicial.Text) + lnTconsumo - Val(txtTreembolso.Text), "####,###,##0.00")
    Set Me.grdReembolsos.DataSource = mrs_Reembolsos
End Sub





Private Sub grdReembolsos_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    Dim lnKeyCode As Integer
    lnKeyCode = KeyCode
    AdministrarKeyPreview lnKeyCode
End Sub

Private Sub grdReembolsos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
'        On Error Resume Next
'        If KeyAscii = 13 Then
'            Select Case grdReembolsos.ActiveCell.Column.Key
'            Case "Consumo", "PorReembolsar", "Reembolsado"
'                SendKeys "{Tab}"
'            End Select
'        End If
End Sub

Private Sub grdReembolsos_LostFocus()
    SumaTotales True
End Sub











Private Sub txtDctoHospital_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDctoHospital
   AdministrarKeyPreview KeyCode

End Sub

Private Sub txtDescripcionR_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDescripcionR
   AdministrarKeyPreview KeyCode

End Sub



Private Sub txtDireccionProv_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDireccionProv
   AdministrarKeyPreview KeyCode

End Sub

Private Sub txtEmailProv_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtEmailProv
   AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaFin
   AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFechaFin_LostFocus()
    If Not EsFecha(txtFechaFin.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
        txtFechaFin.Text = sighEntidades.FECHA_VACIA_DMY
        Exit Sub
    End If
    AsignaMesAnioAlComboSegunRangoFechas
End Sub

Private Sub txtFechaIni_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaIni
   AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFechaIni_LostFocus()
    If Not EsFecha(txtFechaIni.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
        txtFechaIni.Text = sighEntidades.FECHA_VACIA_DMY
        Exit Sub
    End If

End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNcuenta_LostFocus()
    
    If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
        Dim oRsTmp As New Recordset
        Dim lbNuevo As Boolean
        Dim lnConsumo As Double
        Dim oDODiagnostico As New DODiagnostico
        Dim oConexion As New Connection
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(Val(txtNcuenta.Text), oConexion)
        oRsTmp.Filter = "idEstadoAtencion<>0"
        If oRsTmp.RecordCount > 0 Then
            If mrs_Reembolsos.RecordCount = 0 And chkSinBuscarCtas.Value = 1 Then
               mo_cmbAreaTramitaR.BoundText = oRsTmp.Fields!idAreaTramitaSeguros
               mo_cmbFuenteFinanciamiento.BoundText = oRsTmp.Fields!IdFuenteFinanciamiento
               lc_FuenteFinanciamientoPermitidos = mo_cmbFuenteFinanciamiento.BoundText
            End If
            If ValidaElAgregarCuenta(oRsTmp) = True Then
                Set oDODiagnostico = mo_ReglasFacturacion.DevuelveDxAltaMedica(oRsTmp.Fields!idAtencion, oRsTmp.Fields!idTipoServicio)
                lnConsumo = RetornaConsumoPorCuenta(oRsTmp.Fields!idCuentaAtencion, Val(mo_cmbAreaTramitaR.BoundText), Val(mo_cmbFuenteFinanciamiento.BoundText), Val(mo_cmbTipoConsumo.BoundText))
                mrs_Reembolsos.AddNew
                mrs_Reembolsos.Fields!seleccionar = True
                mrs_Reembolsos.Fields!nrocuenta = Val(txtNcuenta.Text)
                mrs_Reembolsos.Fields!EstadoCuenta = oRsTmp.Fields!estadoCta
                mrs_Reembolsos.Fields!NroHistoria = oRsTmp.Fields!NroHistoriaClinica
                mrs_Reembolsos.Fields!Paciente = Trim(oRsTmp.Fields!ApellidoPaterno) + " " + Trim(oRsTmp.Fields!ApellidoMaterno) + " " + oRsTmp.Fields!PrimerNombre
                mrs_Reembolsos.Fields!Servicio = oRsTmp.Fields!dTipoServicio
                If (oRsTmp.Fields!idTipoServicio <> 2 And oRsTmp.Fields!idTipoServicio <> 3 And oRsTmp.Fields!idTipoServicio <> 4) Or oRsTmp.Fields!EsPacienteExterno = True Then
                   mrs_Reembolsos.Fields!fAltaMedica = oRsTmp.Fields!FechaIngreso
                Else
                   mrs_Reembolsos.Fields!fAltaMedica = oRsTmp.Fields!fechaEgreso
                End If
                mrs_Reembolsos.Fields!consumo = lnConsumo
                mrs_Reembolsos.Fields!Reembolsado = lnConsumo
                mrs_Reembolsos.Fields!porReembolsar = 0
                mrs_Reembolsos.Fields!dxid = oDODiagnostico.idDiagnostico
                mrs_Reembolsos.Fields!DxAltaMedica = Left(Trim(oDODiagnostico.CodigoCIE2004) & " " & Trim(oDODiagnostico.descripcion), 100)
                mrs_Reembolsos.Fields!NroReferenciaDestino = oRsTmp.Fields!NroReferenciaDestino
                mrs_Reembolsos.Fields!idAtencion = oRsTmp.Fields!idAtencion
                mrs_Reembolsos.Fields!idTipoServicio = oRsTmp.Fields!idTipoServicio
                mrs_Reembolsos.Fields!UltReembolsoPorCuenta = lcIdReembolsoPorCuenta
                mrs_Reembolsos.Update
                SumaTotales False
                lnCuentaActualDelGrid = Val(txtNcuenta.Text)
            End If
            FraNcuenta.Visible = False
        End If
        Set oRsTmp = Nothing
        oConexion.Close
        Set oConexion = Nothing
        SumaTotales True
    End If
End Sub

Function ValidaElAgregarCuenta(oRsCuenta As Recordset) As Boolean
        Dim lbNuevo As Boolean
        lbNuevo = True
        If oRsCuenta.Fields!idTipoServicio > 1 And IsNull(oRsCuenta.Fields!fechaEgreso) Then
            If oRsCuenta.Fields!EsPacienteExterno = False Then
                MsgBox "La Cuenta N° " & oRsCuenta.Fields!idCuentaAtencion & Chr(13) & "  no tiene ALTA MEDICA", vbInformation, Me.Caption
                lbNuevo = False
            End If
        ElseIf lc_FuenteFinanciamientoPermitidos <> "" And InStr(lc_FuenteFinanciamientoPermitidos, Trim(Str(oRsCuenta.Fields!IdFuenteFinanciamiento))) = 0 Then
            MsgBox "La Cuenta N° " & oRsCuenta.Fields!idCuentaAtencion & Chr(13) & "  tiene otro PLAN DE ATENCION", vbInformation, Me.Caption
            lbNuevo = False
        ElseIf oRsCuenta.Fields!idEstado = 4 Then
            MsgBox "La Cuenta N° " & oRsCuenta.Fields!idCuentaAtencion & Chr(13) & "  tiene estado=PAGADA", vbInformation, Me.Caption
            lbNuevo = False
        End If
        If mrs_Reembolsos.RecordCount > 0 And lbNuevo = True Then
           mrs_Reembolsos.MoveFirst
           mrs_Reembolsos.Find "nroCuenta=" & oRsCuenta.Fields!idCuentaAtencion
           If Not mrs_Reembolsos.EOF Then
               MsgBox "La Cuenta N° " & oRsCuenta.Fields!idCuentaAtencion & Chr(13) & "  ya EXISTEN EN LA LISTA", vbInformation, Me.Caption
               lbNuevo = False
           End If
        End If
        ValidaElAgregarCuenta = lbNuevo
End Function







Private Sub txtNroDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroDocumento
   AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNroSerie_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroSerie
   AdministrarKeyPreview KeyCode

End Sub





Private Sub txtRazonSocial_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtRazonSocial
   AdministrarKeyPreview KeyCode

End Sub

Private Sub txtRuc_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtRuc
   AdministrarKeyPreview KeyCode

End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub



Private Sub txtRuc_LostFocus()
   If txtRuc.Text <> "" Then
       mo_Formulario.HabilitarDeshabilitar txtRazonSocial, False
       mo_Formulario.HabilitarDeshabilitar txtEmailProv, True
       mo_Formulario.HabilitarDeshabilitar txtDireccionProv, True
       If Len(txtRuc.Text) <> 11 Then
          MsgBox "El Número de RUC debe tener 11 dígitos", vbInformation, Me.Caption
       Else
          CargaDatosDelProveedor
       End If
   End If
End Sub

Sub CargaDatosDelProveedor()
          Dim oRsTmp As New ADODB.Recordset
          Set oRsTmp = mo_ReglasFacturacion.ProveedoresSeleccionarPorRUC(txtRuc.Text)
          If oRsTmp.RecordCount > 0 Then
             txtRuc.Tag = oRsTmp!idProveedor
             txtRazonSocial.Text = oRsTmp.Fields!razonSocial
             txtEmailProv.Text = IIf(IsNull(oRsTmp.Fields!Email), "", oRsTmp.Fields!Email)
             txtDireccionProv.Text = IIf(IsNull(oRsTmp!Direccion), "", oRsTmp!Direccion)
          Else
             txtRuc.Tag = 0
             txtRazonSocial.Text = ""
             txtEmailProv.Text = ""
             txtDireccionProv.Text = ""
             mo_Formulario.HabilitarDeshabilitar txtRazonSocial, True
             On Error Resume Next
             txtRazonSocial.SetFocus
          End If
          oRsTmp.Close
          Set oRsTmp = Nothing

End Sub

Private Sub txtSaldoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSaldoInicial
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtSaldoInicial_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtSaldoInicial_LostFocus()
    SumaTotales False
End Sub

Private Sub txtTreembolso_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtTreembolso_LostFocus()
     SumaTotales False
End Sub

Sub LimpiarDatos()
    FraNcuenta.Visible = False
    txtNcuenta.Text = ""
    txtSaldoInicial.Text = ""
    txtTporReembolsar.Text = ""
    txtTconsumo.Text = ""
    txtTreembolso.Text = ""
    txtSaldoFinal.Text = ""
    txtDescripcionR.Text = ""
    txtRuc.Text = ""
    txtRazonSocial.Text = ""
    Me.txtNroSerie.Text = ""
    Me.txtNroDocumento.Text = ""
    lnIdComprobantePagoActual = 0
    lcNroSerieActual = ""
    lcNroDocumentoActual = ""
    lnIdTipoComprobanteActual = 0
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        btnCancelar_Click
    Case vbKeyF2
        btnAceptar_Click
     Case vbKeyF10
        Dim oConsultaCta As New ReembolsosCta
        oConsultaCta.idCuentaAtencion = lnCuentaActualDelGrid
        oConsultaCta.Show 1
        Set oConsultaCta = Nothing
     Case vbKeyF11
        Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
        Dim oDODiagnostico As DODiagnostico
        oBusqueda.MostrarFormulario
        If oBusqueda.BotonPresionado = sghAceptar Then
            Set oDODiagnostico = mo_ReglasComunes.DiagnosticosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
            If Not oDODiagnostico Is Nothing Then
               mrs_Reembolsos.Fields!DxAltaMedica = Left(oDODiagnostico.CodigoCIE2004 & " " & oDODiagnostico.descripcion, 100)
               mrs_Reembolsos.Fields!dxid = oDODiagnostico.idDiagnostico
               mrs_Reembolsos.Update
            End If
        End If
        Set oBusqueda = Nothing
        Set oDODiagnostico = Nothing
    End Select
       
End Sub


Function DevuelveCodigosSegunIdDelPlan(lnIdFuenteFinanciamiento As Long) As String
    Dim oRsTmp9 As New Recordset
    Dim lcCodigos As String
    lcCodigos = ""
    Set oRsTmp9 = mo_ReglasFacturacion.FuentesFinanciamientoDevuelveTodosSegunFiltro("")
    If oRsTmp9.RecordCount > 0 Then
       oRsTmp9.MoveFirst
       Do While Not oRsTmp9.EOF
           lcCodigos = lcCodigos & "." & Trim(Str(oRsTmp9.Fields!IdFuenteFinanciamiento))
          oRsTmp9.MoveNext
       Loop
    End If
    oRsTmp9.Close
    Set oRsTmp9 = Nothing
    DevuelveCodigosSegunIdDelPlan = lcCodigos
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub


Sub CargaDatosAlObjetosDeDatos()
    With mo_DoFactReembolsos
        .Anio = Val(cmbAnio.Text)
        .ConsumoPorReembolsar = CCur(txtTconsumo.Text)
        .descripcion = txtDescripcionR.Text
        .Documentos = txtDctoHospital.Text
        .idAreaTramitaSeguro = Val(mo_cmbAreaTramitaR.BoundText)
        .IdFuenteFinanciamiento = Val(mo_cmbFuenteFinanciamiento.BoundText)
        .IdUsuarioAuditoria = ml_idUsuario
        .Mes = cmbMes.ListIndex + 1
        If txtTreembolso.Text <> "" Then
           .ReembolsoPagado = CCur(txtTreembolso.Text)
        End If
        If txtTporReembolsar.Text <> "" Then
           .ReembolsoPorPagar = CCur(txtTporReembolsar.Text)
        End If
        .SaldoFinal = CCur(txtSaldoFinal.Text)
        If txtSaldoInicial.Text <> "" Then
           .SaldoInicial = CCur(txtSaldoInicial.Text)
        End If
        .idEstadoReembolso = 1
        .idTipoConsumo = Val(mo_cmbTipoConsumo.BoundText)
        .IdTipoComprobante = Val(mo_cmbIdTipoComprobante.BoundText)
        .GrabaDefinitivamente = IIf(chkGrabaDefinitivamente.Value = 1, True, False)
    End With
    With mo_DoFactReembolsosDcto
       .IdComprobantePago = lnIdComprobantePagoActual       'lnIdTipoComprobanteActual
       .IdFactReembolso = mo_DoFactReembolsos.IdFactReembolso
       .IdUsuarioAuditoria = ml_idUsuario
       .nrodocumento = txtNroDocumento.Text
       .nroSerie = txtNroSerie.Text
       .MotivoAnulacion = txtMotivoAnulacion.Text
    End With
End Sub


Sub ActualizaProveedor()
        On Error GoTo errActPrv
        Dim oProveedores As New Proveedores
        Dim oDoProveedores As New DoProveedores
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        Set oProveedores.Conexion = oConexion
        oDoProveedores.idProveedor = txtRuc.Tag
        oDoProveedores.IdUsuarioAuditoria = sighEntidades.Usuario
        oDoProveedores.ruc = txtRuc.Text
        If txtRuc.Tag > 0 Then
            oDoProveedores.Email = txtEmailProv.Text
            oDoProveedores.Direccion = txtDireccionProv.Text
            oDoProveedores.razonSocial = txtRazonSocial.Text
            If Not oProveedores.Modificar(oDoProveedores) Then
            End If
        Else
            oDoProveedores.ruc = txtRuc.Text
            oDoProveedores.Email = txtEmailProv.Text
            oDoProveedores.Direccion = txtDireccionProv.Text
            oDoProveedores.razonSocial = txtRazonSocial.Text
            If Not oProveedores.Insertar(oDoProveedores) Then
            End If
        End If
        oConexion.Close
errActPrv:
        Set oProveedores = Nothing
        Set oDoProveedores = Nothing
        Set oConexion = Nothing

End Sub


