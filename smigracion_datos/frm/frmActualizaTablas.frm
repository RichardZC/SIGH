VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmActualizaTablas 
   Caption         =   "Actualizar, buscar"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   9165
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9105
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   16060
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tablas"
      TabPicture(0)   =   "frmActualizaTablas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Facturacion Farmacia"
      TabPicture(1)   =   "frmActualizaTablas.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "grdFacturacionBienesFinanciamiento"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtNroCuenta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command19"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command14"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Facturacion Servicios"
      TabPicture(2)   =   "frmActualizaTablas.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "grdFacturacionServicioFinanciamientos"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame5"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtCuentaS"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Command10"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Command24"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "HBT- Actualizar Datos"
      TabPicture(3)   =   "frmActualizaTablas.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblProcesando"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame10"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmbConsideraciones"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame6"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame7"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Frame8"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Frame9"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "vacio"
      TabPicture(4)   =   "frmActualizaTablas.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame16"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "txtProblemas"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frameweb"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "txtNuievasCuentas"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "cmdCambiaCPT"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "fracta"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "SIS - Central "
      TabPicture(5)   =   "frmActualizaTablas.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame15"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame14"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Frame13"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Frame12"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Frame11"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Frame39"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).ControlCount=   6
      Begin VB.Frame fracta 
         Caption         =   "Cambia una CUENTA por OTRA para un despacho"
         Height          =   2490
         Left            =   -66105
         TabIndex        =   151
         Top             =   480
         Width           =   4470
         Begin VB.CommandButton cmdCambiaCuenta 
            Caption         =   "Cambia la CUENTA para N°DCTO"
            Height          =   435
            Left            =   120
            TabIndex        =   161
            Top             =   1965
            Width           =   3825
         End
         Begin VB.TextBox txtCuentaD 
            Height          =   360
            Left            =   1380
            TabIndex        =   160
            Top             =   1485
            Width           =   1875
         End
         Begin VB.TextBox txtCuentaO 
            Height          =   360
            Left            =   1230
            TabIndex        =   158
            Top             =   225
            Width           =   1875
         End
         Begin VB.TextBox txtDcto 
            Height          =   360
            Left            =   930
            TabIndex        =   157
            Top             =   975
            Width           =   1875
         End
         Begin Threed.SSOption optFarmacia 
            Height          =   300
            Left            =   150
            TabIndex        =   154
            Top             =   570
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "Farmacia"
            Value           =   -1
         End
         Begin Threed.SSOption optLaboratorio 
            Height          =   300
            Left            =   1260
            TabIndex        =   155
            Top             =   600
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "Laboratorio"
         End
         Begin Threed.SSOption optImagen 
            Height          =   300
            Left            =   2475
            TabIndex        =   156
            Top             =   615
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "Imágen"
         End
         Begin Threed.SSOption optOtros 
            Height          =   300
            Left            =   3540
            TabIndex        =   162
            Top             =   645
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "Otros"
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta DESTINO"
            Height          =   195
            Left            =   75
            TabIndex        =   159
            Top             =   1575
            Width           =   1275
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "N° DCTO"
            Height          =   195
            Left            =   210
            TabIndex        =   153
            Top             =   1035
            Width           =   675
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta actual"
            Height          =   195
            Left            =   120
            TabIndex        =   152
            Top             =   315
            Width           =   990
         End
      End
      Begin VB.CommandButton cmdCambiaCPT 
         Caption         =   "Cambia el CPT 80061 por 80064"
         Height          =   435
         Left            =   -70155
         TabIndex        =   148
         Top             =   2595
         Width           =   3990
      End
      Begin VB.TextBox txtNuievasCuentas 
         Height          =   1695
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   147
         Text            =   "frmActualizaTablas.frx":00A8
         Top             =   7185
         Width           =   12990
      End
      Begin VB.Frame Frameweb 
         BackColor       =   &H8000000D&
         Caption         =   "Actualiza datos SIGHweb"
         Height          =   2070
         Left            =   -70170
         TabIndex        =   139
         Top             =   495
         Width           =   4020
         Begin VB.CommandButton cmdActWeb 
            Caption         =   "Actualiza datos de atenciones en la WEB Galenhos"
            Height          =   555
            Left            =   135
            TabIndex        =   146
            Top             =   1230
            Width           =   2640
         End
         Begin MSMask.MaskEdBox txtFweb1 
            Height          =   345
            Left            =   720
            TabIndex        =   140
            Top             =   660
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFweb2 
            Height          =   345
            Left            =   2685
            TabIndex        =   141
            Top             =   675
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblHoraFin 
            AutoSize        =   -1  'True
            Caption         =   "..."
            Height          =   195
            Left            =   2670
            TabIndex        =   145
            Top             =   300
            Width           =   135
         End
         Begin VB.Label lblHoraInicio 
            AutoSize        =   -1  'True
            Caption         =   "..."
            Height          =   195
            Left            =   720
            TabIndex        =   144
            Top             =   315
            Width           =   135
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "hasta"
            Height          =   195
            Left            =   2190
            TabIndex        =   143
            Top             =   735
            Width           =   390
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Fechas"
            Height          =   195
            Left            =   165
            TabIndex        =   142
            Top             =   705
            Width           =   525
         End
      End
      Begin VB.TextBox txtProblemas 
         Height          =   2730
         Left            =   -74865
         MultiLine       =   -1  'True
         TabIndex        =   138
         Text            =   "frmActualizaTablas.frx":00B5
         Top             =   4305
         Width           =   12990
      End
      Begin VB.Frame Frame16 
         Height          =   2685
         Left            =   -74970
         TabIndex        =   129
         Top             =   420
         Width           =   4785
         Begin VB.TextBox TxtCuenta 
            Height          =   360
            Left            =   1335
            TabIndex        =   150
            Top             =   1455
            Width           =   1875
         End
         Begin VB.TextBox txtMDB 
            Height          =   345
            Left            =   945
            TabIndex        =   131
            Text            =   "atenciones     (sql, que apunte a bd atenciones)"
            Top             =   240
            Width           =   3720
         End
         Begin VB.CommandButton cmdAgregaAtencionCE 
            Caption         =   "Agrega atenciones del Médico que falta"
            Height          =   435
            Left            =   60
            TabIndex        =   130
            Top             =   2070
            Width           =   4605
         End
         Begin MSMask.MaskEdBox txtF1 
            Height          =   345
            Left            =   945
            TabIndex        =   132
            Top             =   645
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtF2 
            Height          =   345
            Left            =   3480
            TabIndex        =   133
            Top             =   660
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Solo la cuenta"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   149
            Top             =   1545
            Width           =   1020
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "* tabla mdb: RecetasdetalleItem..DocumentoDespacho texto(50)"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   90
            TabIndex        =   137
            Top             =   1140
            Width           =   4560
         End
         Begin VB.Label Label23 
            Caption         =   "odbc"
            Height          =   285
            Left            =   90
            TabIndex        =   136
            Top             =   270
            Width           =   1515
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "F.Citas"
            Height          =   195
            Left            =   90
            TabIndex        =   135
            Top             =   720
            Width           =   480
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "hasta"
            Height          =   195
            Left            =   3000
            TabIndex        =   134
            Top             =   720
            Width           =   390
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1185
         Left            =   -66960
         TabIndex        =   121
         Top             =   3840
         Width           =   4785
         Begin VB.CheckBox chkHistorias 
            Caption         =   "Actualiza los datos de Movimientos de Historias ya grabadas?"
            Height          =   345
            Left            =   120
            TabIndex        =   123
            Top             =   270
            Width           =   4575
         End
         Begin VB.CommandButton cmdHistorias 
            Caption         =   "3) Agrega los Movimiento de Historias que faltan"
            Enabled         =   0   'False
            Height          =   405
            Left            =   120
            TabIndex        =   122
            Top             =   660
            Width           =   4575
         End
      End
      Begin VB.Frame Frame8 
         Height          =   975
         Left            =   -66960
         TabIndex        =   118
         Top             =   2790
         Width           =   4785
         Begin VB.CommandButton cmdProgramacion 
            Caption         =   "2) Agrega los Programación que faltan"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   120
            Top             =   510
            Width           =   4575
         End
         Begin VB.CheckBox chkProgramacion 
            Caption         =   "Actualiza los datos de Programació ya grabados?"
            Height          =   345
            Left            =   90
            TabIndex        =   119
            Top             =   180
            Width           =   3975
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1545
         Left            =   -66960
         TabIndex        =   110
         Top             =   5070
         Width           =   4785
         Begin VB.CommandButton cmdProcesaAtenciones 
            Caption         =   "4) Agrega Movimiento de Atenciones"
            Height          =   435
            Left            =   60
            TabIndex        =   112
            Top             =   1020
            Width           =   4605
         End
         Begin VB.TextBox txtOdbc 
            Height          =   345
            Left            =   1650
            TabIndex        =   111
            Text            =   "GalenhosHBT"
            Top             =   240
            Width           =   3015
         End
         Begin MSMask.MaskEdBox txtFinicial 
            Height          =   345
            Left            =   1650
            TabIndex        =   113
            Top             =   660
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFfinal 
            Height          =   345
            Left            =   3480
            TabIndex        =   114
            Top             =   660
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "hasta"
            Height          =   195
            Left            =   3000
            TabIndex        =   117
            Top             =   720
            Width           =   390
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "F.tabla Atenciones"
            Height          =   195
            Left            =   90
            TabIndex        =   116
            Top             =   720
            Width           =   1320
         End
         Begin VB.Label Label3 
            Caption         =   "ODBC del Servidor:"
            Height          =   285
            Left            =   90
            TabIndex        =   115
            Top             =   270
            Width           =   1515
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1425
         Left            =   -66960
         TabIndex        =   106
         Top             =   1350
         Width           =   4785
         Begin VB.OptionButton optPacienteAdd 
            Caption         =   "Adiciona nuevos Pacientes"
            Height          =   285
            Left            =   120
            TabIndex        =   109
            Top             =   210
            Value           =   -1  'True
            Width           =   4545
         End
         Begin VB.OptionButton optPacienteAct 
            Caption         =   "Actualiza los datos de Pacientes ya grabados y Adiciona Nuevos"
            Height          =   345
            Left            =   120
            TabIndex        =   108
            Top             =   540
            Width           =   4545
         End
         Begin VB.CommandButton cmdPacientes 
            Caption         =   "1) Agrega los Pacientes que faltan"
            Enabled         =   0   'False
            Height          =   375
            Left            =   90
            TabIndex        =   107
            Top             =   960
            Width           =   4575
         End
      End
      Begin VB.ListBox cmbConsideraciones 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   7410
         Left            =   -74940
         TabIndex        =   105
         Top             =   1410
         Width           =   7905
      End
      Begin VB.CommandButton Command14 
         Caption         =   "New"
         Height          =   255
         Left            =   -62640
         TabIndex        =   104
         Top             =   2520
         Width           =   465
      End
      Begin VB.CommandButton Command24 
         Caption         =   "New"
         Height          =   255
         Left            =   -62460
         TabIndex        =   103
         Top             =   2490
         Width           =   465
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Del"
         Height          =   255
         Left            =   -62640
         TabIndex        =   102
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Del"
         Height          =   255
         Left            =   -62430
         TabIndex        =   101
         Top             =   2820
         Width           =   375
      End
      Begin VB.TextBox txtCuentaS 
         Height          =   315
         Left            =   -73830
         TabIndex        =   100
         Top             =   8670
         Width           =   1515
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1755
         Left            =   -74910
         TabIndex        =   96
         Top             =   420
         Width           =   12855
         Begin VB.CommandButton Command23 
            Caption         =   "New"
            Height          =   255
            Left            =   12360
            TabIndex        =   98
            Top             =   270
            Width           =   465
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Del"
            Height          =   255
            Left            =   12360
            TabIndex        =   97
            Top             =   600
            Width           =   375
         End
         Begin MSDataGridLib.DataGrid grdFactOrdenServicio 
            Height          =   1485
            Left            =   60
            TabIndex        =   99
            Top             =   180
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   2619
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Cabecera Despacho (FACTORDENSERVICIO)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   4395
         Left            =   -74940
         TabIndex        =   86
         Top             =   4110
         Width           =   12855
         Begin VB.CommandButton Command27 
            Caption         =   "New"
            Height          =   255
            Left            =   12330
            TabIndex        =   92
            Top             =   3180
            Width           =   465
         End
         Begin VB.CommandButton Command26 
            Caption         =   "New"
            Height          =   255
            Left            =   12330
            TabIndex        =   91
            Top             =   1470
            Width           =   465
         End
         Begin VB.CommandButton Command25 
            Caption         =   "New"
            Height          =   255
            Left            =   12330
            TabIndex        =   90
            Top             =   180
            Width           =   465
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Del"
            Height          =   255
            Left            =   12360
            TabIndex        =   89
            Top             =   3510
            Width           =   375
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Del"
            Height          =   255
            Left            =   12420
            TabIndex        =   88
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Del"
            Height          =   255
            Left            =   12330
            TabIndex        =   87
            Top             =   480
            Width           =   375
         End
         Begin MSDataGridLib.DataGrid grdFactOrdenServicioPagos 
            Height          =   1245
            Left            =   60
            TabIndex        =   93
            Top             =   150
            Width           =   12195
            _ExtentX        =   21511
            _ExtentY        =   2196
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Cabecera Pagos (FactOrdenesServicioPagos)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid grdFacturacionServicioPagos 
            Height          =   1665
            Left            =   60
            TabIndex        =   94
            Top             =   1410
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   2937
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Detalle Pagos (FacturacionServicioPagos)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid grdCajaComprobantesPagoS 
            Height          =   1245
            Left            =   60
            TabIndex        =   95
            Top             =   3120
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   2196
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Recibos Servicios (CajaComprobantePago)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   4305
         Left            =   -74940
         TabIndex        =   76
         Top             =   4260
         Width           =   12855
         Begin VB.CommandButton Command17 
            Caption         =   "New"
            Height          =   255
            Left            =   12360
            TabIndex        =   82
            Top             =   3210
            Width           =   465
         End
         Begin VB.CommandButton Command16 
            Caption         =   "New"
            Height          =   255
            Left            =   12330
            TabIndex        =   81
            Top             =   1530
            Width           =   465
         End
         Begin VB.CommandButton Command15 
            Caption         =   "New"
            Height          =   255
            Left            =   12330
            TabIndex        =   80
            Top             =   180
            Width           =   465
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Del"
            Height          =   255
            Left            =   12360
            TabIndex        =   79
            Top             =   3510
            Width           =   375
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Del"
            Height          =   255
            Left            =   12420
            TabIndex        =   78
            Top             =   1830
            Width           =   375
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Del"
            Height          =   255
            Left            =   12360
            TabIndex        =   77
            Top             =   480
            Width           =   375
         End
         Begin MSDataGridLib.DataGrid grdFactOrdenesBienes 
            Height          =   1245
            Left            =   60
            TabIndex        =   83
            Top             =   150
            Width           =   12225
            _ExtentX        =   21564
            _ExtentY        =   2196
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Cabecera Pagos (FactOrdenesBienes)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid grdFacturacionBienesPagos 
            Height          =   1575
            Left            =   60
            TabIndex        =   84
            Top             =   1500
            Width           =   12225
            _ExtentX        =   21564
            _ExtentY        =   2778
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Detalle Pagos (FacturacionBienesPagos)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid grdCajaComprobantesPago 
            Height          =   1095
            Left            =   60
            TabIndex        =   85
            Top             =   3180
            Width           =   12285
            _ExtentX        =   21669
            _ExtentY        =   1931
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Recibos Farmacia (CajaComprobantePago)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1905
         Left            =   -74910
         TabIndex        =   72
         Top             =   420
         Width           =   12855
         Begin VB.CommandButton Command13 
            Caption         =   "New"
            Height          =   255
            Left            =   12330
            TabIndex        =   74
            Top             =   180
            Width           =   465
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Del"
            Height          =   255
            Left            =   12330
            TabIndex        =   73
            Top             =   480
            Width           =   375
         End
         Begin MSDataGridLib.DataGrid grdFarmMovimientoVentas 
            Height          =   1635
            Left            =   60
            TabIndex        =   75
            Top             =   180
            Width           =   12225
            _ExtentX        =   21564
            _ExtentY        =   2884
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Cabecera Despacho (FARMMOVIMIENTOVENTAS)"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtNroCuenta 
         Height          =   315
         Left            =   -73830
         TabIndex        =   71
         Top             =   8640
         Width           =   1515
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Tablas: GalenHos"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   7875
         Index           =   0
         Left            =   60
         TabIndex        =   58
         Top             =   570
         Width           =   12945
         Begin VB.CommandButton Command29 
            Caption         =   "..."
            Height          =   225
            Left            =   6930
            TabIndex        =   69
            ToolTipText     =   "PREVENTAS: Genera Punto de Carga si no existe (incluye IdServicio) y lo asocia a cada  Procedimiento CPT"
            Top             =   6750
            Width           =   405
         End
         Begin VB.CommandButton Command28 
            Caption         =   "..."
            Height          =   195
            Left            =   5040
            TabIndex        =   68
            ToolTipText     =   "Actualiza 'Convenios' para 'Clinica hospitalizados' (despachos en farmacia)"
            Top             =   6840
            Width           =   255
         End
         Begin VB.CommandButton Command4 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   12420
            TabIndex        =   67
            ToolTipText     =   "New"
            Top             =   300
            Width           =   435
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Del"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   12420
            TabIndex        =   66
            Top             =   570
            Width           =   435
         End
         Begin VB.TextBox txtGalenHos 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   65
            Top             =   7020
            Width           =   12255
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   225
            Left            =   5850
            TabIndex        =   64
            ToolTipText     =   "Actualiza Precio Venta Farmacia para los nuevos Tarifarios"
            Top             =   6780
            Width           =   345
         End
         Begin VB.CommandButton Command5 
            Caption         =   "..."
            Height          =   195
            Left            =   240
            TabIndex        =   63
            ToolTipText     =   "DesEncriptar texto"
            Top             =   6870
            Width           =   255
         End
         Begin VB.CommandButton Command3 
            Caption         =   "..."
            Height          =   195
            Left            =   1770
            TabIndex        =   62
            ToolTipText     =   "Total Consumo Servicios, segun Nro Cuenta"
            Top             =   6840
            Width           =   255
         End
         Begin VB.CommandButton Command6 
            Caption         =   "..."
            Height          =   195
            Left            =   2160
            TabIndex        =   61
            ToolTipText     =   "Total Consumo Farmacia, segun Nro Cuenta"
            Top             =   6840
            Width           =   255
         End
         Begin VB.CommandButton Command7 
            Caption         =   "..."
            Height          =   195
            Left            =   2790
            TabIndex        =   60
            ToolTipText     =   "Cambia Nro Historia Clinica (ARCHIVO CLINICO)"
            Top             =   6840
            Width           =   255
         End
         Begin VB.CommandButton Command30 
            Caption         =   "..."
            Height          =   405
            Left            =   12060
            TabIndex        =   59
            Top             =   5730
            Width           =   585
         End
         Begin MSDataGridLib.DataGrid grdGalenHos 
            Height          =   6585
            Left            =   120
            TabIndex        =   70
            Top             =   300
            Width           =   12225
            _ExtentX        =   21564
            _ExtentY        =   11615
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame10 
         Height          =   1995
         Left            =   -66960
         TabIndex        =   50
         Top             =   6780
         Width           =   4755
         Begin VB.CommandButton cmdPreciosSismedv2 
            Caption         =   "Actualiza Medicamentos y Precios desde el SISMEDV2"
            Enabled         =   0   'False
            Height          =   495
            Left            =   60
            TabIndex        =   54
            Top             =   1410
            Width           =   4575
         End
         Begin VB.TextBox txtSismedv2 
            Height          =   345
            Left            =   1560
            TabIndex        =   53
            Text            =   "Sismedv2"
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox txtIdCentroCosto 
            Height          =   315
            Left            =   1290
            TabIndex        =   52
            Text            =   "999"
            Top             =   180
            Width           =   645
         End
         Begin VB.TextBox txtPartida 
            Height          =   315
            Left            =   3930
            TabIndex        =   51
            Text            =   "999"
            Top             =   210
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "ODBC del Servidor:"
            Height          =   285
            Left            =   120
            TabIndex        =   57
            Top             =   630
            Width           =   1515
         End
         Begin VB.Label Label11 
            Caption         =   "Id Centro Costo:"
            Height          =   225
            Left            =   90
            TabIndex        =   56
            Top             =   210
            Width           =   1245
         End
         Begin VB.Label Label12 
            Caption         =   "Id Partida:"
            Height          =   225
            Left            =   3090
            TabIndex        =   55
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.Frame Frame39 
         Caption         =   "AS: Graba datos Citas/Archivos/ en MDB"
         ForeColor       =   &H000000FF&
         Height          =   4635
         Left            =   -74850
         TabIndex        =   35
         Top             =   510
         Width           =   6195
         Begin VB.CheckBox chkSolo 
            Caption         =   "Solo Citas"
            Height          =   225
            Left            =   2970
            TabIndex        =   43
            Top             =   1410
            Value           =   1  'Checked
            Width           =   3045
         End
         Begin VB.TextBox txtMesMaximo 
            Height          =   315
            Left            =   1860
            TabIndex        =   42
            Text            =   "12"
            Top             =   1380
            Width           =   405
         End
         Begin VB.TextBox txtNivelEESS 
            Height          =   315
            Left            =   5160
            TabIndex        =   41
            Top             =   960
            Width           =   885
         End
         Begin VB.TextBox txtRenaes 
            Height          =   315
            Left            =   2910
            TabIndex        =   40
            Top             =   960
            Width           =   1035
         End
         Begin VB.TextBox txtAnio 
            Height          =   315
            Left            =   570
            TabIndex        =   39
            Text            =   "2013"
            Top             =   960
            Width           =   585
         End
         Begin VB.CommandButton cmdCitas 
            Caption         =   "procesa  (rCitas, rCitasFa, rCitasDe, rProgCab, rProgDet, rProgSer, rRRHH)"
            Height          =   525
            Left            =   90
            TabIndex        =   38
            Top             =   2280
            Width           =   5925
         End
         Begin VB.CommandButton cmdSoloFarmacia 
            Caption         =   "procesa SALIDAS (genera  'HojaLibre1.xls') para archivos de FARMACIA"
            Height          =   525
            Left            =   90
            TabIndex        =   37
            Top             =   2970
            Width           =   5925
         End
         Begin VB.CommandButton cmdStock 
            Caption         =   "procesa STOCK  (genera  'HojaLibre1.xls') para archivos de FARMACIA"
            Height          =   525
            Left            =   90
            TabIndex        =   36
            Top             =   3660
            Width           =   5925
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            Caption         =   "Procesar hasta el mes"
            Height          =   195
            Left            =   210
            TabIndex        =   49
            Top             =   1410
            Width           =   1560
         End
         Begin VB.Label lblProceso 
            Caption         =   "...."
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1890
            Width           =   2745
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            Caption         =   "Nivel EESS"
            Height          =   195
            Left            =   4320
            TabIndex        =   47
            Top             =   1020
            Width           =   825
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            Caption         =   "Renaes"
            Height          =   195
            Left            =   2280
            TabIndex        =   46
            Top             =   990
            Width           =   555
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            Caption         =   "Año"
            Height          =   195
            Left            =   210
            TabIndex        =   45
            Top             =   1020
            Width           =   285
         End
         Begin VB.Label Label70 
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmActualizaTablas.frx":00C2
            Height          =   555
            Left            =   90
            TabIndex        =   44
            Top             =   360
            Width           =   5985
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FF8080&
         Caption         =   "AS: Graba datos personales de los PACIENTES desde RENIEC"
         ForeColor       =   &H000000FF&
         Height          =   2640
         Left            =   -68640
         TabIndex        =   24
         Top             =   510
         Width           =   5445
         Begin VB.CommandButton cmdBuscaDNIenRENIEC 
            Caption         =   "Busca los DNI en la WEB RENIEC y actualiza los datos"
            Height          =   435
            Left            =   90
            TabIndex        =   29
            Top             =   2055
            Width           =   5265
         End
         Begin VB.TextBox txtDNIb 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1230
            TabIndex        =   28
            Top             =   1590
            Width           =   945
         End
         Begin VB.TextBox txtProcesados 
            Height          =   315
            Left            =   3390
            TabIndex        =   27
            Top             =   1530
            Width           =   675
         End
         Begin VB.TextBox txtTotalReg 
            Height          =   315
            Left            =   3000
            TabIndex        =   26
            Top             =   1170
            Width           =   1035
         End
         Begin VB.TextBox txtDNIprocesados 
            Height          =   315
            Left            =   1260
            TabIndex        =   25
            Top             =   1170
            Width           =   705
         End
         Begin VB.Label Label7 
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmActualizaTablas.frx":0155
            Height          =   705
            Left            =   90
            TabIndex        =   34
            Top             =   360
            Width           =   5265
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DNI a buscar"
            Height          =   195
            Left            =   270
            TabIndex        =   33
            Top             =   1650
            Width           =   945
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            Height          =   195
            Left            =   2550
            TabIndex        =   32
            Top             =   1260
            Width           =   360
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DNI Buscados"
            Height          =   195
            Left            =   2310
            TabIndex        =   31
            Top             =   1590
            Width           =   1035
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° DNI hallados"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   1260
            Width           =   1140
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "AS: Graba datos en columnas Peso1,Peso2, ...Peso45 de tabla SIP en MDB"
         ForeColor       =   &H000000FF&
         Height          =   2415
         Left            =   -74820
         TabIndex        =   18
         Top             =   5250
         Width           =   6285
         Begin VB.CommandButton CommandCuatro 
            Caption         =   "procesar (desde dbf)"
            Enabled         =   0   'False
            Height          =   705
            Left            =   90
            TabIndex        =   21
            Top             =   1380
            Width           =   1665
         End
         Begin VB.CommandButton cmdProcesaSip 
            Caption         =   "Procesar2 (desde 2 tablas mdb)"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2820
            TabIndex        =   20
            Top             =   1320
            Width           =   3255
         End
         Begin VB.CommandButton cmdProcesaSip2000 
            Caption         =   "Procesar1"
            Enabled         =   0   'False
            Height          =   345
            Left            =   3900
            TabIndex        =   19
            Top             =   1920
            Visible         =   0   'False
            Width           =   2265
         End
         Begin VB.Label Label55 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Este proceso lee el archivo ........\archivos\percentiles.xls"
            Height          =   315
            Left            =   90
            TabIndex        =   23
            Top             =   930
            Width           =   6105
         End
         Begin VB.Label Label14 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe tener  el archivo .....\tablasYpa.mdb (previamente se importó parte01, parte02 sin la columna FECHA"
            Height          =   555
            Left            =   90
            TabIndex        =   22
            Top             =   360
            Width           =   6105
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00FF8080&
         Caption         =   "Percentiles"
         Height          =   1260
         Left            =   -74790
         TabIndex        =   14
         Top             =   7725
         Width           =   3465
         Begin VB.TextBox txtHoja 
            Height          =   375
            Left            =   1860
            TabIndex        =   16
            Text            =   "JEPELACIO"
            Top             =   240
            Width           =   1410
         End
         Begin VB.CommandButton cmdActualizaPercentil 
            Caption         =   "Agrega 3 columnas de Percentiles"
            Height          =   510
            Left            =   120
            TabIndex        =   15
            Top             =   630
            Width           =   3195
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Hoja de PADRON.XLS"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   330
            Width           =   1635
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Percentiles"
         Height          =   2010
         Left            =   -68610
         TabIndex        =   8
         Top             =   3195
         Width           =   4485
         Begin VB.CommandButton cmdPercentilSM 
            Caption         =   "Agrega 1 columna de Percentil"
            Height          =   510
            Left            =   135
            TabIndex        =   10
            Top             =   1245
            Width           =   3195
         End
         Begin VB.TextBox txtExcelSM 
            Height          =   375
            Left            =   720
            TabIndex        =   9
            Text            =   "C:\total_sm.XLSx"
            Top             =   810
            Width           =   1575
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Excel"
            Height          =   195
            Left            =   195
            TabIndex        =   13
            Top             =   900
            Width           =   390
         End
         Begin VB.Label Label16 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Este proceso lee el archivo ........\archivos\percentiles.xls"
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   330
            Width           =   4215
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "la hoja es: TOTALFB"
            Height          =   195
            Left            =   2430
            TabIndex        =   11
            Top             =   900
            Width           =   1485
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00FF8080&
         Caption         =   "Percentiles"
         Height          =   2010
         Left            =   -68490
         TabIndex        =   2
         Top             =   5325
         Width           =   4485
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   675
            TabIndex        =   4
            Text            =   "C:\total_sm.XLSX"
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton cmdVariaColumPerc 
            Caption         =   "Agrega  columna 0 al 45 de Percentil de c/madre"
            Height          =   510
            Left            =   90
            TabIndex        =   3
            Top             =   1395
            Width           =   4230
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "la hoja es: hoja1"
            Height          =   195
            Left            =   2385
            TabIndex        =   7
            Top             =   1050
            Width           =   1155
         End
         Begin VB.Label Label19 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Este proceso lee el archivo ........\archivos\percentiles.xls   *pasar 'total_sm.xlsx' a la bd 'sigh_externa'"
            Height          =   555
            Left            =   120
            TabIndex        =   6
            Top             =   330
            Width           =   4215
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Excel"
            Height          =   195
            Left            =   150
            TabIndex        =   5
            Top             =   1050
            Width           =   390
         End
      End
      Begin MSDataGridLib.DataGrid grdFacturacionBienesFinanciamiento 
         Height          =   1725
         Left            =   -74910
         TabIndex        =   124
         Top             =   2460
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   3043
         _Version        =   393216
         BackColor       =   8454143
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Detalle Seguros (FacturacionBienesFinanciamientos)"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grdFacturacionServicioFinanciamientos 
         Height          =   1725
         Left            =   -74940
         TabIndex        =   125
         Top             =   2310
         Width           =   12435
         _ExtentX        =   21934
         _ExtentY        =   3043
         _Version        =   393216
         BackColor       =   8454143
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Detalle Seguros (FacturacionServicioFinanciamientos)"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblProcesando 
         Caption         =   "........"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   128
         Top             =   8850
         Width           =   7635
      End
      Begin VB.Label Label2 
         Caption         =   "Nro Cuenta:"
         Height          =   255
         Index           =   0
         Left            =   -74910
         TabIndex        =   127
         Top             =   8700
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Cuenta:"
         Height          =   255
         Left            =   -74910
         TabIndex        =   126
         Top             =   8670
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmActualizaTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Consulta Cuentas
'        Programado por: Barrantes D
'        Fecha: Diciembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit
Dim wrs_Gal As New ADODB.Recordset
Dim oRsFarmMovimientoVentas As New ADODB.Recordset
Dim oRsCajaComprobantePago As New ADODB.Recordset
Dim oRsFacturacionBienesFinanciamiento As New ADODB.Recordset
Dim oRsFactOrdenesBienes As New ADODB.Recordset
Dim oRsFacturacionBienesPagos As New ADODB.Recordset
Dim oRsFactOrdenServicio As New ADODB.Recordset
Dim oRsCajaComprobantePagoS As New ADODB.Recordset
Dim oRsFacturacionServicioFinanciamientos As New ADODB.Recordset
Dim oRsFactOrdenServicioPagos As New ADODB.Recordset
Dim oRsFacturacionServicioPagos As New ADODB.Recordset
Dim oRsPatologia As New Recordset
Dim oRsFarmacia As New Recordset
Dim lcSql As String
Dim oRsUltCodigo As Long
Const lnIdUsuario As Long = 738
Const lnIdTipoFinanciamiento As Long = 1
Const lnIdFuenteFinanciamiento As Long = 1
Const ln2020 As Long = 9999999
Const lcVacio As String = "(VACIO)"
Dim mo_conexion As ADODB.Connection
Dim lnErrCA As Long
Dim ml_Errores As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_Procesos As New SIGHProxies.Procesos



Private Sub cmdActWeb_Click()
        Dim lcMensaje As String, lbSeTerminaSistema As Boolean, oRsTmp1 As Recordset
        lblHoraInicio.Caption = Format(Now, "hh:mm:ss")
        'SCCQ 14/02/2020 Cambio3 Inicio
        ''mo_Procesos.SomeeActualizaDatos 2, lcMensaje, "", "", CDate(Me.txtFweb1.Text), CDate(Me.txtFweb2.Text), lbSeTerminaSistema, oRsTmp1
        'SCCQ 14/02/2020 Cambio3 Fin
        lblHoraFin.Caption = Format(Now, "hh:mm:ss")
        Set oRsTmp1 = Nothing
        MsgBox lcMensaje, vbInformation, Me.Caption
End Sub


Function ConvierteEnDias(lcEdad As String) As Long
    If InStr(lcEdad, "h") > 0 Then
       ConvierteEnDias = 1
    ElseIf InStr(lcEdad, "días") > 0 Or InStr(lcEdad, "d") > 0 Then
       ConvierteEnDias = Val(Left(lcEdad, InStr(lcEdad, "d") - 1))
    ElseIf InStr(lcEdad, "meses") > 0 Then
       ConvierteEnDias = (Val(Left(lcEdad, InStr(lcEdad, "meses") - 1)) * 30) + 29
    Else
       ConvierteEnDias = (Val(Left(lcEdad, InStr(lcEdad, "a") - 1)) * 365) + 364
    End If
End Function








Private Sub cmdCambiaCPT_Click()
     Me.MousePointer = 11
     Dim oRsTmp1 As New Recordset
     Dim oRsTmp2 As New Recordset
     Dim oConexion As New Connection
     Dim lcSql As String
     Dim lnIdProductoSale As Long, lnIdProductoQueda As Long
     lnIdProductoSale = 3278
     lnIdProductoQueda = 51554
     oConexion.CursorLocation = adUseClient
     oConexion.CommandTimeout = 300
     oConexion.Open SIGHEntidades.CadenaConexion
     lcSql = "SELECT     dbo.LabMovimientoLaboratorio.IdCuentaAtencion, dbo.LabMovimientoLaboratorio.IdComprobantePago, dbo.LabMovimiento.IdLabEstado" & _
"                      dbo.LabMovimiento.IdMovimiento, dbo.LabMovimiento.Fecha, dbo.LabMovimientoCPT.idProductoCPT," & _
"                      dbo.LabMovimientoLaboratorio.IdOrden" & _
" FROM         dbo.LabMovimientoCPT INNER JOIN" & _
"                      dbo.LabMovimientoLaboratorio ON dbo.LabMovimientoCPT.idMovimiento = dbo.LabMovimientoLaboratorio.IdMovimiento INNER JOIN" & _
"                      dbo.LabMovimiento ON dbo.LabMovimientoCPT.idMovimiento = dbo.LabMovimiento.IdMovimiento" & _
" where not (3278 in (select  dbo.LabResultadoPorItems.idProductoCPT from dbo.LabResultadoPorItems where" & _
"                    dbo.LabResultadoPorItems.idOrden=dbo.LabMovimientoLaboratorio.IdOrden and" & _
"                    dbo.LabResultadoPorItems.idProductoCpt = dbo.LabMovimientoCPT.idProductoCPT ) )" & _
"      and (dbo.LabMovimiento.IdLabEstado<>0) and dbo.LabMovimientoCPT.idProductoCPT  =3278"
     oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
     If oRsTmp1.RecordCount > 0 Then
        Do While Not oRsTmp1.EOF
            lcSql = "update FacturacionServicioDespacho set IdProducto =" & lnIdProductoQueda & _
                    " where idProducto =" & lnIdProductoSale & " and idOrden=" & oRsTmp1!IdOrden
            If oRsTmp2.State = 1 Then oRsTmp2.Close
            oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
            lcSql = "update FacturacionServicioFinanciamientos set IdProducto =" & lnIdProductoQueda & _
                    " where idProducto =" & lnIdProductoSale & " and idOrden=" & oRsTmp1!IdOrden
            If oRsTmp2.State = 1 Then oRsTmp2.Close
            oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
            lcSql = "update LabMovimientoCPT set idProductoCPT =" & lnIdProductoQueda & _
                    " where idProductoCPT=" & lnIdProductoSale & "and idMovimiento="" & oRsTmp1!IdMovimiento"
            If oRsTmp2.State = 1 Then oRsTmp2.Close
            oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
            lcSql = "update FacturacionServicioPagos set idProducto =" & lnIdProductoQueda & _
                    " where idProducto =" & lnIdProductoSale & " and idOrdenPago =" & _
                    "   (select * from dbo.FactOrdenServicioPagos where idOrden =" & oRsTmp1!IdOrden & " )"
            If oRsTmp2.State = 1 Then oRsTmp2.Close
            oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
            oRsTmp1.MoveNext
        Loop
     End If
     oRsTmp1.Close
     Set oRsTmp1 = Nothing
     Set oRsTmp2 = Nothing
     Me.MousePointer = 1
     Unload Me
End Sub

Private Sub cmdCambiaCuenta_Click()
    If Val(txtCuentaO.Text) <= 0 Then
       MsgBox "La CUENTA ORIGEN debe ser mayor a CERO", vbInformation, ""
       Exit Sub
    End If
    If txtDcto.Text = "" Then
       MsgBox "Debe ingresar el N°DCTO", vbInformation, ""
       Exit Sub
    End If
    If Val(txtCuentaD.Text) <= 0 Then
       MsgBox "La CUENTA DESTINO debe ser mayor a CERO", vbInformation, ""
       Exit Sub
    End If
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        Dim oConexion As New Connection
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim oRsTmp3 As New Recordset
        Dim oRsTmp4 As New Recordset
        Dim oRsTmp5 As New Recordset
        Dim lcSql As String
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        
        If Me.optFarmacia.Value = True Then
           lcSql = "SELECT     dbo.farmMovimiento.DocumentoNumero, dbo.farmMovimientoVentas.idCuentaAtencion, dbo.Atenciones.IdFormaPago, " & _
                   "   dbo.Atenciones.idFuenteFinanciamiento,dbo.farmMovimiento.movNumero,dbo.farmMovimiento.movTipo" & _
                   " FROM         dbo.farmMovimiento INNER JOIN" & _
                   "   dbo.farmMovimientoVentas ON dbo.farmMovimiento.MovNumero = dbo.farmMovimientoVentas.movNumero AND" & _
                   "   dbo.farmMovimiento.MovTipo = dbo.farmMovimientoVentas.movTipo INNER JOIN" & _
                   "   dbo.Atenciones ON dbo.farmMovimientoVentas.idCuentaAtencion = dbo.Atenciones.IdCuentaAtencion" & _
                   " WHERE dbo.farmMovimiento.DocumentoNumero='" & txtDcto.Text & "'"
           oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
           If oRsTmp1.RecordCount = 0 Then
              MsgBox "Ese DOCUMENTO no existe"
           Else
              If oRsTmp1!idCuentaAtencion <> Val(Me.txtCuentaO.Text) Then
                 MsgBox "Esa CUENTA ORIGEN no pertenece al N°DCTO de despacho en FARMACIA"
              Else
                 lcSql = "select * from atenciones where idcuentaatencion=" & Me.txtCuentaD.Text
                 oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                 If oRsTmp2.RecordCount = 0 Then
                    MsgBox "La CUENTA DESTINO no existe"
                 Else
                    If oRsTmp1!IdFormaPago <> oRsTmp2!IdFormaPago Or oRsTmp1!idFuenteFinanciamiento <> oRsTmp2!idFuenteFinanciamiento Then
                       MsgBox "La FUENTE FIANACIAMIENTO o TARIFAS son diferentes"
                    Else
                       lcSql = "select * from FactOrdenesBienes where movNumero='" & oRsTmp1!movNumero & "' and movTipo='" & oRsTmp1!movTipo & "'"
                       oRsTmp3.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                       If oRsTmp3.RecordCount > 0 Then
                          oRsTmp3.MoveFirst
                          Do While Not oRsTmp3.EOF
                            If Not IsNull(oRsTmp3!idComprobantePago) Then
                               lcSql = "select * from CajaComprobantesPago where idComprobantePago=" & oRsTmp3!idComprobantePago
                               oRsTmp4.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                               If oRsTmp4.RecordCount > 0 Then
                                  oRsTmp4!idCuentaAtencion = Val(Me.txtCuentaD.Text)
                                  oRsTmp4.Update
                               End If
                               oRsTmp4.Close
                            End If
                            oRsTmp3!idCuentaAtencion = Val(Me.txtCuentaD.Text)
                            oRsTmp3.Update
                            oRsTmp3.MoveNext
                          Loop
                       End If
                       oRsTmp3.Close
                       lcSql = "update farmMovimientoVentas set idCuentaAtencion=" & Me.txtCuentaD.Text & _
                             " where movNumero='" & oRsTmp1!movNumero & "' and movTipo='" & oRsTmp1!movTipo & "'"
                       If oRsTmp3.State = 1 Then oRsTmp3.Close
                       oRsTmp3.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                    End If
                 End If
              End If
           End If
           oRsTmp1.Close
        Else
           If Me.optImagen.Value = True Then
              lcSql = "SELECT  * From ImagMovimientoImagenes where idMovimiento=" & Me.txtDcto.Text
           ElseIf Me.optLaboratorio.Value = True Then
              lcSql = "SELECT  * From LabMovimientoLaboratorio where idMovimiento=" & Me.txtDcto.Text
           Else
              lcSql = "select * from FactOrdenServicio where idOrden=" & Me.txtDcto.Text
           End If
           oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
           If oRsTmp1.RecordCount = 0 Then
              MsgBox "Ese DOCUMENTO no existe"
           Else
              If oRsTmp1!idCuentaAtencion <> Val(Me.txtCuentaO.Text) Then
                 MsgBox "Esa CUENTA ORIGEN no pertenece al N°DCTO de despacho en " & _
                        IIf(Me.optLaboratorio.Value, "Laboratorio", IIf(Me.optImagen.Value, "Imagen", "Otros"))
              Else
                 lcSql = "select * from atenciones where idcuentaatencion=" & Me.txtCuentaD.Text
                 oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                 If oRsTmp2.RecordCount = 0 Then
                    MsgBox "La CUENTA DESTINO no existe"
                 Else
                    lcSql = "select * from FactOrdenServicio where idOrden=" & oRsTmp1!IdOrden
                    oRsTmp3.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                    If oRsTmp3.RecordCount > 0 Then
                        If oRsTmp3!IdTipoFinanciamiento <> oRsTmp2!IdFormaPago Or oRsTmp3!idFuenteFinanciamiento <> oRsTmp2!idFuenteFinanciamiento Then
                           MsgBox "La FUENTE FIANACIAMIENTO o TARIFAS son diferentes"
                        Else
                           lcSql = "select * from factOrdenServicioPagos where idOrden=" & oRsTmp1!IdOrden
                           oRsTmp4.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                           If oRsTmp4.RecordCount > 0 Then
                              oRsTmp4.MoveFirst
                              Do While Not oRsTmp4.EOF
                                If Not IsNull(oRsTmp4!idComprobantePago) Then
                                  lcSql = "update CajaComprobantesPago set idCuentaAtencion=" & Me.txtCuentaD.Text & _
                                          " where idComprobantePago=" & oRsTmp4!idComprobantePago
                                  If oRsTmp5.State = 1 Then oRsTmp5.Close
                                  oRsTmp5.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                                End If
                                oRsTmp4.MoveNext
                              Loop
                           End If
                           oRsTmp4.Close
                           oRsTmp3!idCuentaAtencion = Val(Me.txtCuentaD.Text)
                           oRsTmp3.Update
                           If Me.optOtros = False Then
                                oRsTmp1!idCuentaAtencion = Val(Me.txtCuentaD.Text)
                                oRsTmp1.Update
                           End If
                        End If
                    End If
                    oRsTmp3.Close
                 End If
                 oRsTmp2.Close
              End If
              oRsTmp1.Close
           End If
        End If
        oConexion.Close
        Set oConexion = Nothing
        Set oRsTmp1 = Nothing
        Set oRsTmp2 = Nothing
        Set oRsTmp3 = Nothing
        Set oRsTmp4 = Nothing
        Set oRsTmp5 = Nothing
        Unload Me
    End If
End Sub

Private Sub cmdCitas_Click()
     If txtNivelEESS.Text = "" Then
        MsgBox "Averiguar NIVEL ESTABLECIMIENTO", vbCritical, ""
        Exit Sub
     End If
     Dim oConexionMDB As New Connection
     Dim oConexion As New Connection
     Dim oRsTmp1 As New Recordset
     Dim oRsTmp2 As New Recordset
     Dim oRsCitas As New Recordset
     Dim oRsCitasFa As New Recordset
     Dim oRsCitasDe As New Recordset
     Dim oRsProgCab As New Recordset
     Dim oRsProgDet As New Recordset
     Dim oRsProgServ As New Recordset
     Dim oRsRRHH As New Recordset
     
     Dim mo_ReglasFacturacion As New ReglasFacturacion
     Dim mo_ReglasFarmacia As New ReglasFarmacia
     Dim lnMonto As Double, ldFechaIngreso As Date
     Dim lnRegistros As Long, lcRenaes As String, lnHorasProgr As Long, lnHorasCit As Long
     Dim lnIdPaciente As Long, lnMes As Long, lnMontoFarmacia As Double, lnFFSis As Double
     Dim lnFFSoat As Double, lnFFParticular As Double, lnFFConvenio As Double
     Dim lnIdEspecialidad As Long, lnMontoFarmXesp As Double, lcPaciente As String, lcHistoria As String
     Dim lcDNI As String, ldFechaNa As Date, lcSexo As String, lcDpto As String, lcEspecialidad As String
     Dim lcProv As String, lcDist As String, lcEducacio As String, lcIdioma As String
     Dim lcEtnia As String, lnIdMedico As Long, lcMedico As String, lcColegiatura As String, lcCondTrab As String
     Dim lnNroConsultorios As Long, lnThorasProgamadas As Long, lnTlnHorasCitas As Long, lcTipoProf As String
     Dim lnFFSis1 As Integer, lnFFSoat1 As Integer, lnFFParticular1 As Integer, lnFFConvenio1 As Integer
     Dim lcEESSnivel As String, lcEESSnombre As String, lnHorasProgramadas As Long, lnHorasCitas As Long
     Dim lnIdServicio As Long, lcConsultorio As String, lnHorasProg As Integer
     Const lnHorasProgramEmerg As Integer = 24
     Me.MousePointer = 11
     '
     oConexion.CursorLocation = adUseClient
     oConexion.CommandTimeout = 300
     oConexion.Open SIGHEntidades.CadenaConexion
     oConexionMDB.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\tablasYpa.mdb;"
     '
     lcRenaes = txtRenaes.Text
     lcEESSnivel = txtNivelEESS.Text
     lcEESSnombre = lcBuscaParametro.SeleccionaFilaParametro(205)
     '**************************************************** RRHH ********************************************************
     If chkSolo.Value <> 1 Then
        lblProceso.Caption = "(2/4) RRHH"
        'Elimina los movimientos del establecimiento/año
        lcSql = "delete from rRRHH where anio='" & Me.txtAnio.Text & "' and eess='" & lcRenaes & "'"
        oRsRRHH.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
        lcSql = "select * from rRRHH"
        oRsRRHH.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
        'Barre movimientos y agrega a tabla MDB
        lcSql = "SELECT     dbo.ProgramacionMedica.*, LTRIM(dbo.Empleados.ApellidoPaterno) + ' ' + LTRIM(dbo.Empleados.ApellidoMaterno) " & _
               "                      + ' ' + dbo.Empleados.Nombres AS Medico, dbo.Empleados.DNI, dbo.Medicos.Colegiatura, dbo.TiposOcupacion.descripcion AS Ocupacion," & _
               "                      dbo.TiposCondicionTrabajo.Descripcion AS CondTrabaj, dbo.Especialidades.Nombre AS EspecialidadS" & _
               " FROM         dbo.Especialidades RIGHT OUTER JOIN" & _
               "                      dbo.ProgramacionMedica ON dbo.Especialidades.IdEspecialidad = dbo.ProgramacionMedica.IdEspecialidad LEFT OUTER JOIN" & _
               "                      dbo.Empleados INNER JOIN" & _
               "                      dbo.Medicos ON dbo.Empleados.IdEmpleado = dbo.Medicos.IdEmpleado LEFT OUTER JOIN" & _
               "                      dbo.TiposOcupacion ON dbo.Empleados.IdTipoEmpleado = dbo.TiposOcupacion.IdTipoOcupacion LEFT OUTER JOIN" & _
               "                      dbo.TiposCondicionTrabajo ON dbo.Empleados.IdCondicionTrabajo = dbo.TiposCondicionTrabajo.IdCondicionTrabajo ON" & _
               "                      dbo.ProgramacionMedica.IdMedico = dbo.Medicos.IdMedico" & _
               " Where (dbo.ProgramacionMedica.IdTipoServicio = 1) and  Year(dbo.ProgramacionMedica.Fecha) = " & Me.txtAnio.Text & _
               "        and month(dbo.ProgramacionMedica.Fecha) <= " & Me.txtMesMaximo.Text & _
               " order by month(dbo.ProgramacionMedica.Fecha),dbo.ProgramacionMedica.idMedico,dbo.ProgramacionMedica.idEspecialidad"
        oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
        lnRegistros = oRsTmp1.RecordCount
        If lnRegistros > 0 Then
           ProgressBar1.Min = 0
           ProgressBar1.Max = lnRegistros
           ProgressBar1.Value = 0
           Do While Not oRsTmp1.EOF
              lnMes = Month(oRsTmp1.Fields!Fecha)
              Do While Not oRsTmp1.EOF And lnMes = Month(oRsTmp1.Fields!Fecha)
                   lnIdMedico = oRsTmp1.Fields!idMedico
                   lcMedico = oRsTmp1.Fields!medico
                   lcDNI = IIf(IsNull(oRsTmp1.Fields!DNI), "", Left(oRsTmp1.Fields!DNI, 8))
                   lcColegiatura = IIf(IsNull(oRsTmp1.Fields!Colegiatura), "", oRsTmp1.Fields!Colegiatura)
                   lcTipoProf = IIf(IsNull(oRsTmp1.Fields!ocupacion), "", oRsTmp1.Fields!ocupacion)
                   lcCondTrab = IIf(IsNull(oRsTmp1.Fields!CondTrabaj), "", oRsTmp1.Fields!CondTrabaj)
                   
                   Do While Not oRsTmp1.EOF And lnMes = Month(oRsTmp1.Fields!Fecha) And lnIdMedico = oRsTmp1.Fields!idMedico
                      lnHorasProgramadas = 0
                      lnHorasCitas = 0
                      lnMonto = 0
                      lnIdEspecialidad = oRsTmp1.Fields!IdEspecialidad
                      lcEspecialidad = IIf(IsNull(oRsTmp1.Fields!EspecialidadS), "", oRsTmp1.Fields!EspecialidadS)
                      Do While Not oRsTmp1.EOF And lnMes = Month(oRsTmp1.Fields!Fecha) And lnIdMedico = oRsTmp1.Fields!idMedico And lnIdEspecialidad = oRsTmp1.Fields!IdEspecialidad
                           lnHorasProgramadas = lnHorasProgramadas + DateDiff("h", CDate(oRsTmp1.Fields!HoraInicio), CDate(oRsTmp1.Fields!HoraFin))
                           lcSql = "SELECT     dbo.Citas.Fecha, dbo.Servicios.Nombre AS Consultorio, dbo.Citas.HoraInicio, dbo.Citas.HoraFin, dbo.Citas.IdServicio, dbo.Servicios.IdTipoServicio, " & _
                                   "           dbo.Atenciones.IdCuentaAtencion,dbo.Atenciones.IdFormaPago" & _
                                   " FROM         dbo.Citas LEFT OUTER JOIN" & _
                                   "                      dbo.Atenciones ON dbo.Citas.IdAtencion = dbo.Atenciones.IdAtencion LEFT OUTER JOIN" & _
                                   "                      dbo.Servicios ON dbo.Citas.IdServicio = dbo.Servicios.IdServicio" & _
                                   "  Where  dbo.Citas.idProgramacion=" & oRsTmp1.Fields!IdProgramacion
                           oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                           If oRsTmp2.RecordCount > 0 Then
                              oRsTmp2.MoveFirst
                              Do While Not oRsTmp2.EOF
                                   If oRsTmp2.Fields!IdFormaPago = 1 Then
                                       'Es pagante
                                       lnMonto = lnMonto + mo_ReglasFacturacion.RetornaTotalConsumoFarmaciaParaPagantePorNroCuenta(oRsTmp2.Fields!idCuentaAtencion)
                                   Else
                                       'tiene seguros
                                       lnMonto = lnMonto + mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(oRsTmp2.Fields!idCuentaAtencion)
                                   End If
                                   
                                   lnHorasCitas = lnHorasCitas + DateDiff("n", CDate(oRsTmp2.Fields!HoraInicio), CDate(oRsTmp2.Fields!HoraFin))
                                   oRsTmp2.MoveNext
                              Loop
                           End If
                           oRsTmp2.Close
                           
                           DoEvents: ProgressBar1.Value = ProgressBar1.Value + 1: Me.Refresh
                           oRsTmp1.MoveNext
                           If oRsTmp1.EOF Then
                              Exit Do
                           End If
                       Loop
                       lnHorasCitas = Round(lnHorasCitas / 60, 0)
                       oRsRRHH.AddNew
                       oRsRRHH.Fields!Anio = Me.txtAnio.Text
                       oRsRRHH.Fields!Mes = Trim(Str(lnMes))
                       oRsRRHH.Fields!eess = lcRenaes
                       oRsRRHH.Fields!eessnomb = lcEESSnombre
                       oRsRRHH.Fields!Personal = lcMedico
                       oRsRRHH.Fields!DNI = lcDNI
                       oRsRRHH.Fields!Colegiat = lcColegiatura
                       oRsRRHH.Fields!TipoProf = lcTipoProf
                       oRsRRHH.Fields!condTrab = lcCondTrab
                       oRsRRHH.Fields!especial = lcEspecialidad
                       oRsRRHH.Fields!HrProgra = lnHorasProgramadas
                       oRsRRHH.Fields!HrSinCit = lnHorasProgramadas - lnHorasCitas
                       If lnHorasProgramadas - lnHorasCitas = 0 Then
                          oRsRRHH.Fields!Promedio = 0
                       Else
                          oRsRRHH.Fields!Promedio = Round(lnHorasProgramadas / (lnHorasProgramadas - lnHorasCitas), 2)
                       End If
                       oRsRRHH.Fields!Farmacia = lnMonto
                       oRsRRHH.Update
                       If oRsTmp1.EOF Then
                          Exit Do
                       End If
                   Loop
                   If oRsTmp1.EOF Then
                      Exit Do
                   End If
              Loop
              If oRsTmp1.EOF Then
                  Exit Do
              End If
           Loop
        End If
        oRsTmp1.Close
     End If
     '**************************************************** PROGRAMACION ********************************************************
     If chkSolo.Value <> 1 Then
         lblProceso.Caption = "(3/4) Programacion"
         'Elimina los movimientos del establecimiento/año
         lcSql = "delete from rProgCab where anio='" & Me.txtAnio.Text & "' and eess='" & lcRenaes & "'"
         oRsTmp1.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
         lcSql = "delete from rProgDet where anio='" & Me.txtAnio.Text & "' and eess='" & lcRenaes & "'"
         oRsTmp1.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
         lcSql = "select * from rProgCab"
         oRsProgCab.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
         lcSql = "select * from rProgDet"
         oRsProgDet.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
         
         lcSql = "delete from rProgSer where year(fecha)=" & Me.txtAnio.Text & " and eess='" & lcRenaes & "'"
         oRsProgServ.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
         lcSql = "select *  from rProgSer "
         oRsProgServ.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
         'Barre movimientos y agrega a tabla MDB
         lcSql = "SELECT     dbo.Servicios.Nombre AS Consultorio, dbo.Servicios.IdTipoServicio, dbo.ProgramacionMedica.*" & _
                " FROM         dbo.Servicios RIGHT OUTER JOIN" & _
                "                      dbo.ProgramacionMedica ON dbo.Servicios.IdServicio = dbo.ProgramacionMedica.IdServicio" & _
                " Where (dbo.Servicios.IdTipoServicio = 1) and  Year(dbo.ProgramacionMedica.Fecha) = " & Me.txtAnio.Text & _
                "        and month(dbo.ProgramacionMedica.Fecha) <= " & Me.txtMesMaximo.Text & _
                " order by month(dbo.ProgramacionMedica.Fecha),dbo.ProgramacionMedica.idServicio,dbo.ProgramacionMedica.Fecha"
         oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
         lnRegistros = oRsTmp1.RecordCount
         If lnRegistros > 0 Then
            ProgressBar1.Min = 0
            ProgressBar1.Max = lnRegistros
            ProgressBar1.Value = 0
            Do While Not oRsTmp1.EOF
               lnMes = Month(oRsTmp1.Fields!Fecha)
               lnNroConsultorios = 0
               lnThorasProgamadas = 0
               lnTlnHorasCitas = 0
               Do While Not oRsTmp1.EOF And lnMes = Month(oRsTmp1.Fields!Fecha)
                    lnIdServicio = oRsTmp1.Fields!IdServicio
                    lcConsultorio = oRsTmp1.Fields!consultorio
                    lnHorasProgramadas = 0
                    lnHorasCitas = 0
                    Do While Not oRsTmp1.EOF And lnMes = Month(oRsTmp1.Fields!Fecha) And lnIdServicio = oRsTmp1.Fields!IdServicio
                        lnHorasProg = DateDiff("h", CDate(oRsTmp1.Fields!HoraInicio), CDate(oRsTmp1.Fields!HoraFin))
                        lnHorasProgramadas = lnHorasProgramadas + lnHorasProg
                        ldFechaIngreso = oRsTmp1.Fields!Fecha
                        
                        lcSql = "SELECT     dbo.Citas.Fecha, dbo.Servicios.Nombre AS Consultorio, dbo.Citas.HoraInicio, dbo.Citas.HoraFin, dbo.Citas.IdServicio, " & _
                                 "                      dbo.Servicios.IdTipoServicio" & _
                                 " FROM         dbo.Citas LEFT OUTER JOIN" & _
                                 "                      dbo.Servicios ON dbo.Citas.IdServicio = dbo.Servicios.IdServicio" & _
                                 " Where (dbo.Servicios.IdTipoServicio = 1) and dbo.Citas.idProgramacion=" & oRsTmp1.Fields!IdProgramacion
                        oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                        lnHorasCit = 0
                        If oRsTmp2.RecordCount > 0 Then
                           oRsTmp2.MoveFirst
                           Do While Not oRsTmp2.EOF
                              lcSql = Format(oRsTmp2.Fields!Fecha)
                              lnHorasCit = lnHorasCit + DateDiff("n", CDate(oRsTmp2.Fields!HoraInicio), CDate(oRsTmp2.Fields!HoraFin))
                              oRsTmp2.MoveNext
                           Loop
                           lnHorasCitas = lnHorasCitas + lnHorasCit
                        End If
                        oRsTmp2.Close
                        lnHorasCitas = Round(lnHorasCit / 60, 0)
                        oRsProgServ.AddNew
                        oRsProgServ.Fields!Fecha = ldFechaIngreso
                        oRsProgServ.Fields!eess = lcRenaes
                        oRsProgServ.Fields!eessnomb = lcEESSnombre
                        oRsProgServ.Fields!Consultor = lcConsultorio
                        oRsProgServ.Fields!ConsHrPr = lnHorasProg
                        oRsProgServ.Fields!ConsHrCi = lnHorasCitas
                        oRsProgServ.Fields!ConsHrDi = lnHorasProg - lnHorasCitas
                        oRsProgServ.Update
                        
                        DoEvents: ProgressBar1.Value = ProgressBar1.Value + 1: Me.Refresh
                        oRsTmp1.MoveNext
                        If oRsTmp1.EOF Then
                           Exit Do
                        End If
                    Loop
                    lnHorasCitas = Round(lnHorasCitas / 60, 0)
                    oRsProgDet.AddNew
                    oRsProgDet.Fields!Anio = Me.txtAnio.Text
                    oRsProgDet.Fields!Mes = Trim(Str(lnMes))
                    oRsProgDet.Fields!eess = lcRenaes
                    oRsProgDet.Fields!eessnomb = lcEESSnombre
                    oRsProgDet.Fields!Consultor = lcConsultorio
                    oRsProgDet.Fields!ConsHrPr = lnHorasProgramadas
                    oRsProgDet.Fields!ConsHrCi = lnHorasCitas
                    oRsProgDet.Fields!ConsHrDi = lnHorasProgramadas - lnHorasCitas
                    oRsProgDet.Update
                    lnNroConsultorios = lnNroConsultorios + 1
                    lnThorasProgamadas = lnThorasProgamadas + lnHorasProgramadas
                    lnTlnHorasCitas = lnTlnHorasCitas + lnHorasCitas
                    If oRsTmp1.EOF Then
                       Exit Do
                    End If
               Loop
                oRsProgCab.AddNew
                oRsProgCab.Fields!Anio = Me.txtAnio.Text
                oRsProgCab.Fields!Mes = Trim(Str(lnMes))
                oRsProgCab.Fields!eess = lcRenaes
                oRsProgCab.Fields!consNro = lnNroConsultorios
                oRsProgCab.Fields!eessnomb = lcEESSnombre
                oRsProgCab.Fields!ConsProg = lnNroConsultorios
                oRsProgCab.Fields!ConsHrPr = lnThorasProgamadas
                oRsProgCab.Fields!ConsHrCi = lnTlnHorasCitas
                oRsProgCab.Fields!ConsHrDi = lnThorasProgamadas - lnTlnHorasCitas
                oRsProgCab.Update
               If oRsTmp1.EOF Then
                   Exit Do
               End If
            Loop
         End If
         oRsTmp1.Close
         'Programacion de Servicios en Emergencia, solo para tabla: "rProgSer"
         lcSql = "SELECT     dbo.Atenciones.IdServicioEgreso as idServicio, dbo.Servicios.Nombre AS Consultorio, dbo.Atenciones.IdTipoServicio, dbo.Atenciones.FechaIngreso, " & _
                "                      dbo.Atenciones.HoraIngreso, dbo.Atenciones.FechaEgreso, dbo.Atenciones.HoraEgreso" & _
                " FROM         dbo.Atenciones LEFT OUTER JOIN" & _
                "                      dbo.Servicios ON dbo.Atenciones.IdServicioIngreso = dbo.Servicios.IdServicio" & _
                " Where (dbo.Atenciones.IdTipoServicio = 2) and not (dbo.Atenciones.FechaEgreso is null) " & _
                "       and dbo.Atenciones.idEstadoAtencion<>0 " & _
                "       and  Year(dbo.Atenciones.FechaIngreso) = " & Me.txtAnio.Text & _
                "       and month(dbo.Atenciones.FechaIngreso) <= " & Me.txtMesMaximo.Text & _
                " order by dbo.Atenciones.IdServicioIngreso,dbo.Atenciones.FechaIngreso"
         oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
        ' oRstmp1.Filter = " Consultorio='Emergencia Ginecología'"
         lnRegistros = oRsTmp1.RecordCount
         If lnRegistros > 0 Then
            ProgressBar1.Min = 0
            ProgressBar1.Max = lnRegistros
            ProgressBar1.Value = 0
            Do While Not oRsTmp1.EOF
               lnMes = Month(oRsTmp1.Fields!FechaIngreso)
               lnNroConsultorios = 0
               lnThorasProgamadas = 0
               lnTlnHorasCitas = 0
               Do While Not oRsTmp1.EOF
                    lnIdServicio = oRsTmp1.Fields!IdServicio
                    lcConsultorio = oRsTmp1.Fields!consultorio
                    ldFechaIngreso = oRsTmp1.Fields!FechaIngreso
                    lnHorasProgramadas = 0
                    lnHorasCitas = 0
                    Do While Not oRsTmp1.EOF And ldFechaIngreso = oRsTmp1.Fields!FechaIngreso And lnIdServicio = oRsTmp1.Fields!IdServicio
                        lnHorasCit = DateDiff("h", CDate(oRsTmp1.Fields!FechaIngreso & " " & oRsTmp1.Fields!HoraIngreso), _
                                                    CDate(oRsTmp1.Fields!FechaEgreso & " " & oRsTmp1.Fields!HoraEgreso))
                        If lnHorasCit = 0 Then
                           lnHorasCit = 1
                        End If
                                                    
                        lnHorasCitas = lnHorasCitas + lnHorasCit
                        
                        DoEvents: ProgressBar1.Value = ProgressBar1.Value + 1: Me.Refresh
                        oRsTmp1.MoveNext
                        If oRsTmp1.EOF Then
                           Exit Do
                        End If
                    Loop
                    lnHorasCit = lnHorasCitas  ' Round(lnHorasCitas / 60, 0)
                    oRsProgServ.AddNew
                    oRsProgServ.Fields!Fecha = ldFechaIngreso
                    oRsProgServ.Fields!eess = lcRenaes
                    oRsProgServ.Fields!eessnomb = lcEESSnombre
                    oRsProgServ.Fields!Consultor = lcConsultorio
                    oRsProgServ.Fields!ConsHrPr = lnHorasProgramEmerg
                    oRsProgServ.Fields!ConsHrCi = lnHorasCit
                    oRsProgServ.Fields!ConsHrDi = lnHorasProgramEmerg - lnHorasCit
                    oRsProgServ.Update
                    If oRsTmp1.EOF Then
                       Exit Do
                    End If
                Loop
            Loop
         End If
         oRsTmp1.Close
     End If
     '**************************************************** CITAS ********************************************************
'     If chkSolo.Value <> 1 Then
         lblProceso.Caption = "(4/4) Citas"
         'Elimina los movimientos del establecimiento/año
         lcSql = "delete from rCitas where anio='" & Me.txtAnio.Text & "' and eess='" & lcRenaes & "'"
         oRsTmp1.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
         lcSql = "delete from rCitasFa where anio='" & Me.txtAnio.Text & "' and eess='" & lcRenaes & "'"
         oRsTmp1.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
         lcSql = "delete from rCitasDe where anio='" & Me.txtAnio.Text & "' and eess='" & lcRenaes & "'"
         oRsTmp1.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
         lcSql = "select * from rCitasFa"
         oRsCitasFa.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
         lcSql = "select * from rCitas"
         oRsCitas.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
         lcSql = "select * from rCitasDe"
         oRsCitasDe.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
         'Barre movimientos y agrega a tabla MDB
         lcSql = "SELECT     dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre, dbo.Pacientes.SegundoNombre, " & _
    "                      dbo.Pacientes.NroDocumento, dbo.Pacientes.NroHistoriaClinica, dbo.Pacientes.FechaNacimiento, dbo.Pacientes.IdTipoSexo," & _
    "                      dbo.Departamentos.Nombre AS Dpto, dbo.Provincias.Nombre AS Prov, dbo.Distritos.Nombre AS Dist, dbo.Atenciones.IdFormaPago," & _
    "                      dbo.TiposGradoInstruccion.Descripcion AS Educacion, dbo.TiposIdiomas.Lengua, dbo.HIS_tabetnia.etnias, dbo.Atenciones.FechaIngreso," & _
    "                      dbo.Pacientes.IdPaciente, dbo.Especialidades.Nombre AS Especialidad, dbo.Especialidades.IdEspecialidad," & _
    "                      dbo.Atenciones.idCuentaAtencion" & _
    " FROM         dbo.Especialidades RIGHT OUTER JOIN" & _
    "                      dbo.TiposIdiomas RIGHT OUTER JOIN" & _
    "                      dbo.HIS_tabetnia RIGHT OUTER JOIN" & _
    "                      dbo.Pacientes ON dbo.HIS_tabetnia.codetni = dbo.Pacientes.IdEtnia ON dbo.TiposIdiomas.IdIdioma = dbo.Pacientes.IdIdioma LEFT OUTER JOIN" & _
    "                      dbo.TiposGradoInstruccion ON dbo.Pacientes.IdGradoInstruccion = dbo.TiposGradoInstruccion.IdGradoInstruccion LEFT OUTER JOIN" & _
    "                      dbo.Distritos ON dbo.Pacientes.IdDistritoDomicilio = dbo.Distritos.IdDistrito RIGHT OUTER JOIN" & _
    "                      dbo.Atenciones LEFT OUTER JOIN" & _
    "                      dbo.Servicios ON dbo.Atenciones.IdServicioIngreso = dbo.Servicios.IdServicio ON" & _
    "                      dbo.Pacientes.IdPaciente = dbo.Atenciones.IdPaciente RIGHT OUTER JOIN" & _
    "                      dbo.Provincias ON dbo.Distritos.IdProvincia = dbo.Provincias.IdProvincia LEFT OUTER JOIN" & _
    "                      dbo.Departamentos ON dbo.Provincias.IdDepartamento = dbo.Departamentos.IdDepartamento ON" & _
    "                      dbo.Especialidades.IdEspecialidad = dbo.Servicios.IdEspecialidad" & _
    " Where dbo.Atenciones.idEstadoAtencion<>0 and dbo.Atenciones.idTipoServicio=1 and year(dbo.Atenciones.FechaIngreso)=" & Me.txtAnio.Text & _
    "        and month(dbo.Atenciones.FechaIngreso) <= " & Me.txtMesMaximo.Text & _
    " Order by dbo.Pacientes.IdPaciente,month(dbo.Atenciones.FechaIngreso),dbo.Especialidades.IdEspecialidad"
         oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
    
         lnRegistros = oRsTmp1.RecordCount
         If lnRegistros > 0 Then
            ProgressBar1.Min = 0
            ProgressBar1.Max = lnRegistros
            ProgressBar1.Value = 0
            Do While Not oRsTmp1.EOF
               lcPaciente = Trim(oRsTmp1.Fields!ApellidoPaterno) & " " & Trim(oRsTmp1.Fields!ApellidoMaterno) & " " & _
                            Trim(oRsTmp1.Fields!PrimerNombre) & " " & IIf(IsNull(oRsTmp1.Fields!SegundoNombre), "", Trim(oRsTmp1.Fields!SegundoNombre))
               If IsNull(oRsTmp1.Fields!NroHistoriaClinica) Then
                  lcHistoria = ""
               Else
                  lcHistoria = Trim(Str(oRsTmp1.Fields!NroHistoriaClinica))
               End If
               lcDNI = IIf(IsNull(oRsTmp1.Fields!NroDocumento), "", Left(oRsTmp1.Fields!NroDocumento, 8))
               ldFechaNa = IIf(IsNull(oRsTmp1.Fields!FechaNacimiento), 0, oRsTmp1.Fields!FechaNacimiento)
               lcSexo = IIf(oRsTmp1.Fields!idTipoSexo = 1, "Masculino", "Femenino")
               lcDpto = IIf(IsNull(oRsTmp1.Fields!dpto), "", oRsTmp1.Fields!dpto)
               lcProv = IIf(IsNull(oRsTmp1.Fields!prov), "", oRsTmp1.Fields!prov)
               lcDist = IIf(IsNull(oRsTmp1.Fields!dist), "", oRsTmp1.Fields!dist)
               lcEducacio = IIf(IsNull(oRsTmp1.Fields!educacion), "", Trim(oRsTmp1.Fields!educacion))
               lcIdioma = IIf(IsNull(oRsTmp1.Fields!lengua), "Español", Trim(oRsTmp1.Fields!lengua))
               lcEtnia = IIf(IsNull(oRsTmp1.Fields!etnias), "Mestizo", Trim(oRsTmp1.Fields!etnias))
               
               lnIdPaciente = oRsTmp1.Fields!idPaciente
               lnMes = Month(oRsTmp1.Fields!FechaIngreso)
               lnMontoFarmacia = 0
               lnFFSis = 0: lnFFSoat = 0: lnFFParticular = 0: lnFFConvenio = 0
               Do While Not oRsTmp1.EOF And lnIdPaciente = oRsTmp1.Fields!idPaciente And lnMes = Month(oRsTmp1.Fields!FechaIngreso)
                    lnIdEspecialidad = oRsTmp1.Fields!IdEspecialidad
                    lcEspecialidad = oRsTmp1.Fields!Especialidad
                    lnMontoFarmXesp = 0
                    
                    Do While Not oRsTmp1.EOF And lnIdPaciente = oRsTmp1.Fields!idPaciente And lnMes = Month(oRsTmp1.Fields!FechaIngreso) And lnIdEspecialidad = oRsTmp1.Fields!IdEspecialidad
                        DoEvents: ProgressBar1.Value = ProgressBar1.Value + 1: Me.Refresh
                        lnFFSis1 = 0: lnFFSoat1 = 0: lnFFParticular1 = 0: lnFFConvenio1 = 0
                        Select Case oRsTmp1.Fields!IdFormaPago
                        Case 1
                            'Es pagante
                            lnMonto = mo_ReglasFacturacion.RetornaTotalConsumoFarmaciaParaPagantePorNroCuenta(oRsTmp1.Fields!idCuentaAtencion)
                            lnFFParticular = lnFFParticular + 1
                            lnFFParticular1 = 1
                        Case Else
                            'tiene seguros
                            lnMonto = mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(oRsTmp1.Fields!idCuentaAtencion)
                            If oRsTmp1.Fields!IdFormaPago = 2 Then
                               lnFFSis = lnFFSis + 1
                               lnFFSis1 = 1
                            ElseIf oRsTmp1.Fields!IdFormaPago = 3 Then
                               lnFFSoat = lnFFSoat + 1
                               lnFFSoat1 = 1
                            Else
                               lnFFConvenio = lnFFConvenio + 1
                               lnFFConvenio1 = 1
                            End If
                        End Select
                        lnMontoFarmacia = lnMontoFarmacia + lnMonto
                        lnMontoFarmXesp = lnMontoFarmXesp + lnMonto
                        
                        oRsCitasDe.AddNew
                        oRsCitasDe.Fields!Anio = Me.txtAnio.Text
                        oRsCitasDe.Fields!Mes = Trim(Str(lnMes))
                        oRsCitasDe.Fields!eess = lcRenaes
                        oRsCitasDe.Fields!eessnomb = lcEESSnombre
                        oRsCitasDe.Fields!eessnive = lcEESSnivel
                        oRsCitasDe.Fields!Fecha = oRsTmp1.Fields!FechaIngreso
                        oRsCitasDe.Fields!Paciente = lcPaciente
                        oRsCitasDe.Fields!Historia = lcHistoria
                        oRsCitasDe.Fields!DNI = lcDNI
                        oRsCitasDe.Fields!fnacimie = ldFechaNa
                        oRsCitasDe.Fields!sexo = lcSexo
                        oRsCitasDe.Fields!dpto = lcDpto
                        oRsCitasDe.Fields!provinc = lcProv
                        oRsCitasDe.Fields!Distrito = lcDist
                        oRsCitasDe.Fields!ffsis = lnFFSis1
                        oRsCitasDe.Fields!ffpartic = lnFFParticular1
                        oRsCitasDe.Fields!ffsoat = lnFFSoat1
                        oRsCitasDe.Fields!ffconven = lnFFConvenio1
                        oRsCitasDe.Fields!educacio = lcEducacio
                        oRsCitasDe.Fields!idioma = lcIdioma
                        oRsCitasDe.Fields!etnia = lcEtnia
                        oRsCitasDe.Fields!mfarmaci = lnMonto
                        oRsCitasDe.Update
                        
                        oRsTmp1.MoveNext
                        If oRsTmp1.EOF Then
                           Exit Do
                        End If
                    Loop
                    If lnMontoFarmXesp > 0 Then
                        oRsCitasFa.AddNew
                        oRsCitasFa.Fields!Anio = Me.txtAnio.Text
                        oRsCitasFa.Fields!Mes = Trim(Str(lnMes))
                        oRsCitasFa.Fields!eess = lcRenaes
                        oRsCitasFa.Fields!eessnomb = lcEESSnombre
                        oRsCitasFa.Fields!Paciente = lcPaciente
                        oRsCitasFa.Fields!Historia = lcHistoria
                        oRsCitasFa.Fields!especiali = lcEspecialidad
                        oRsCitasFa.Fields!mfarmaci = lnMontoFarmXesp
                        oRsCitasFa.Update
                    End If
                    If oRsTmp1.EOF Then
                       Exit Do
                    End If
               Loop
               oRsCitas.AddNew
               oRsCitas.Fields!Anio = Me.txtAnio.Text
               oRsCitas.Fields!Mes = Trim(Str(lnMes))
               oRsCitas.Fields!eess = lcRenaes
               oRsCitas.Fields!eessnomb = lcEESSnombre
               oRsCitas.Fields!eessnive = lcEESSnivel
               oRsCitas.Fields!Paciente = lcPaciente
               oRsCitas.Fields!Historia = lcHistoria
               oRsCitas.Fields!DNI = lcDNI
               oRsCitas.Fields!fnacimie = ldFechaNa
               oRsCitas.Fields!sexo = lcSexo
               oRsCitas.Fields!dpto = lcDpto
               oRsCitas.Fields!provinc = lcProv
               oRsCitas.Fields!Distrito = lcDist
               oRsCitas.Fields!ffsis = lnFFSis           'IIf(lnFFSis = 0, 0, Round(lnFFSis * 100 / (lnFFSis + lnFFSoat + lnFFParticular + lnFFConvenio), 2))
               oRsCitas.Fields!ffpartic = lnFFParticular 'IIf(lnFFParticular = 0, 0, Round(lnFFParticular * 100 / (lnFFSis + lnFFSoat + lnFFParticular + lnFFConvenio), 2))
               oRsCitas.Fields!ffsoat = lnFFSoat         'IIf(lnFFSoat = 0, 0, Round(lnFFSoat * 100 / (lnFFSis + lnFFSoat + lnFFParticular + lnFFConvenio), 2))
               oRsCitas.Fields!ffconven = lnFFConvenio   'IIf(lnFFConvenio = 0, 0, Round(lnFFConvenio * 100 / (lnFFSis + lnFFSoat + lnFFParticular + lnFFConvenio), 2))
               oRsCitas.Fields!educacio = lcEducacio
               oRsCitas.Fields!idioma = lcIdioma
               oRsCitas.Fields!etnia = lcEtnia
               oRsCitas.Fields!mfarmaci = lnMontoFarmacia
               oRsCitas.Update
               If oRsTmp1.EOF Then
                   Exit Do
               End If
            Loop
         End If
         oRsTmp1.Close
 '    End If
     '******************************************************************************************
     Me.MousePointer = 1
     Unload Me

End Sub

Private Sub cmdHistorias_Click()
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Dim oConexHBT As New Connection
       Dim oConexion As New Connection
       Dim oRsTmpHBT1 As New Recordset
       Dim oRsTmpHBT2 As New Recordset
       Dim oRsTmpHBT3 As New Recordset
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oRsTmp3 As New Recordset
       Dim oDOMovimientoHistoriaClinica  As New DOMovimientoHistoriaClinica
       Dim oMovimientosHistoriaClinica As New MovimientosHistoriaClinica
       Dim lcSql As String, lnCant As Long, lnTotal As Long
       Dim lnUltimoId As Long
       Dim ms_MensajeError As String
       On Error GoTo Terminar
       Me.MousePointer = 11
       ms_MensajeError = ""
       oConexHBT.Open "dsn=" & txtOdbc.Text
       oConexion.Open SIGHEntidades.CadenaConexion
       oConexion.BeginTrans
       '
       Set oMovimientosHistoriaClinica.Conexion = oConexion
       Set mo_conexion = oConexion
       '
       lnUltimoId = 0
       lcSql = "select * from MovimientosHistoriaClinica order by idMovimiento desc"
       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       If oRsTmp1.RecordCount > 0 Then
           lnUltimoId = oRsTmp1.Fields!idMovimiento
       End If
       oRsTmp1.Close
       lcSql = "select * from MovimientosHistoriaClinica order  by idMovimiento"
       oRsTmpHBT1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
       lnTotal = oRsTmpHBT1.RecordCount
       If lnTotal > 0 Then
          ProgressBar1.Min = 0
          ProgressBar1.Max = lnTotal
          lnCant = 1
          oRsTmpHBT1.MoveFirst
          Do While Not oRsTmpHBT1.EOF
'If lnCant > 100 Then
'Exit Do
'End If
            ProgressBar1.Value = lnCant
            lnCant = lnCant + 1
            'ProgramacionMedica
            With oDOMovimientoHistoriaClinica
                .FechaMovimiento = oRsTmpHBT1.Fields!FechaMovimiento
                '.idAtencion = oRsTmpHBT1.Fields!
                .IdEmpleadoArchivo = oRsTmpHBT1.Fields!IdEmpleadoArchivo
                .IdEmpleadoRecepcion = oRsTmpHBT1.Fields!IdEmpleadoRecepcion
                .IdEmpleadoTransporte = oRsTmpHBT1.Fields!IdEmpleadoTransporte
                .IdGrupoMovimiento = oRsTmpHBT1.Fields!IdGrupoMovimiento
                .idMotivo = oRsTmpHBT1.Fields!idMotivo
                .idMovimiento = oRsTmpHBT1.Fields!idMovimiento
                .idPaciente = oRsTmpHBT1.Fields!idPaciente
                .idServicioDestino = IIf(IsNull(oRsTmpHBT1.Fields!idServicioDestino), 0, oRsTmpHBT1.Fields!idServicioDestino)
                .IdServicioOrigen = IIf(IsNull(oRsTmpHBT1.Fields!IdServicioOrigen), 0, oRsTmpHBT1.Fields!IdServicioOrigen)
                .IdUsuarioAuditoria = lnIdUsuario
                .NroFolios = IIf(IsNull(oRsTmpHBT1.Fields!NroFolios), 0, oRsTmpHBT1.Fields!NroFolios)
                .Observacion = IIf(IsNull(oRsTmpHBT1.Fields!Observacion), "", oRsTmpHBT1.Fields!Observacion)
            End With
            If lnUltimoId < oDOMovimientoHistoriaClinica.idMovimiento Then
                If Not InsertarDebbMovimientoHistoriaClinica(oDOMovimientoHistoriaClinica) Then
                      GoTo Terminar
                End If
            ElseIf Me.chkHistorias.Value = 1 Then
                If Not oMovimientosHistoriaClinica.Modificar(oDOMovimientoHistoriaClinica) Then
                     ms_MensajeError = oMovimientosHistoriaClinica.MensajeError: GoTo Terminar
                End If
            End If
            '
            oRsTmpHBT1.MoveNext
          Loop
       End If
       oRsTmpHBT1.Close
       '
       oConexion.CommitTrans
       Me.MousePointer = 1
       Unload Me
    End If
    Exit Sub
            
Terminar:
    oConexion.RollbackTrans
    MsgBox ms_MensajeError
    Me.MousePointer = 1
    Resume

End Sub











Private Sub cmdPacientes_Click()
     If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Dim oConexHBT As New Connection
       Dim oConexion As New Connection
       Dim oRsTmpHBT1 As New Recordset
       Dim oRsTmpHBT2 As New Recordset
       Dim oRsTmpHBT3 As New Recordset
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oRsTmp3 As New Recordset
       Dim oPaciente As New Pacientes
       Dim oDOPaciente As New DOPaciente
       Dim oDOHistoria As New DOHistoriaClinica
       Dim oHistoria As New HistoriasClinicas
       Dim lcSql As String, lnCant As Long, lnTotal As Long
       Dim lnUltimoId As Long
       Dim ms_MensajeError As String
       Dim oExcel As Excel.Application
       Dim oSheet As Excel.Worksheet
       Dim j As Integer
       On Error GoTo Terminar
       Me.MousePointer = 11
       ms_MensajeError = ""
       ml_Errores = ""
       oConexHBT.Open "dsn=" & txtOdbc.Text
       oConexion.Open SIGHEntidades.CadenaConexion
       oConexion.BeginTrans
       '
       Set oPaciente.Conexion = oConexion
       Set mo_conexion = oConexion
       Set oHistoria.Conexion = oConexion
       '
       Set oExcel = New Excel.Application
       oExcel.Visible = True
       oExcel.Workbooks.Add
       Set oSheet = oExcel.ActiveSheet
       oSheet.Cells(1, 1).Value = "Error"
       oSheet.Cells(1, 6).Value = "Observación"
       j = 3
       '
       lnUltimoId = 0
       lcSql = "select * from Pacientes order by idPaciente desc"
       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       If oRsTmp1.RecordCount > 0 Then
           lnUltimoId = oRsTmp1.Fields!idPaciente
       End If
       oRsTmp1.Close
       If Me.optPacienteAdd.Value = True Then
          lcSql = "select * from Pacientes where idPaciente>" & lnUltimoId & " order by idPaciente"
       Else
          lcSql = "select * from Pacientes order by idPaciente"
       End If
       oRsTmpHBT1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
       lnTotal = oRsTmpHBT1.RecordCount
       If lnTotal > 0 Then
          ProgressBar1.Min = 0
          ProgressBar1.Max = lnTotal
          lnCant = 1
          oRsTmpHBT1.MoveFirst
          Do While Not oRsTmpHBT1.EOF
             ProgressBar1.Value = lnCant
             lnCant = lnCant + 1

             'Tabla: Pacientes
             oDOPaciente.ApellidoMaterno = oRsTmpHBT1.Fields!ApellidoMaterno
             oDOPaciente.ApellidoPaterno = oRsTmpHBT1.Fields!ApellidoPaterno
             oDOPaciente.Autogenerado = oRsTmpHBT1.Fields!Autogenerado
             oDOPaciente.DireccionDomicilio = IIf(IsNull(oRsTmpHBT1.Fields!DireccionDomicilio), "", oRsTmpHBT1.Fields!DireccionDomicilio)
             oDOPaciente.FechaNacimiento = IIf(IsNull(oRsTmpHBT1.Fields!FechaNacimiento), 0, oRsTmpHBT1.Fields!FechaNacimiento)
             oDOPaciente.IdCentroPobladoDomicilio = IIf(IsNull(oRsTmpHBT1.Fields!IdCentroPobladoDomicilio), 0, oRsTmpHBT1.Fields!IdCentroPobladoDomicilio)
             oDOPaciente.IdCentroPobladoNacimiento = IIf(IsNull(oRsTmpHBT1.Fields!IdCentroPobladoNacimiento), 0, oRsTmpHBT1.Fields!IdCentroPobladoNacimiento)
             oDOPaciente.IdCentroPobladoProcedencia = IIf(IsNull(oRsTmpHBT1.Fields!IdCentroPobladoProcedencia), 0, oRsTmpHBT1.Fields!IdCentroPobladoProcedencia)
             oDOPaciente.IdDistritoDomicilio = IIf(IsNull(oRsTmpHBT1.Fields!IdDistritoDomicilio), 0, oRsTmpHBT1.Fields!IdDistritoDomicilio)
             oDOPaciente.IdDistritoNacimiento = IIf(IsNull(oRsTmpHBT1.Fields!IdDistritoNacimiento), 0, oRsTmpHBT1.Fields!IdDistritoNacimiento)
             oDOPaciente.IdDistritoProcedencia = IIf(IsNull(oRsTmpHBT1.Fields!IdDistritoProcedencia), 0, oRsTmpHBT1.Fields!IdDistritoProcedencia)
             oDOPaciente.IdDocIdentidad = IIf(IsNull(oRsTmpHBT1.Fields!IdDocIdentidad), 0, oRsTmpHBT1.Fields!IdDocIdentidad)
             oDOPaciente.IdEstadoCivil = IIf(IsNull(oRsTmpHBT1.Fields!IdEstadoCivil), 0, oRsTmpHBT1.Fields!IdEstadoCivil)
             oDOPaciente.IdGradoInstruccion = IIf(IsNull(oRsTmpHBT1.Fields!IdGradoInstruccion), 0, oRsTmpHBT1.Fields!IdGradoInstruccion)
             oDOPaciente.idPaciente = oRsTmpHBT1.Fields!idPaciente
             oDOPaciente.IdPaisDomicilio = IIf(IsNull(oRsTmpHBT1.Fields!IdPaisDomicilio), 0, oRsTmpHBT1.Fields!IdPaisDomicilio)
             oDOPaciente.IdPaisNacimiento = IIf(IsNull(oRsTmpHBT1.Fields!IdPaisNacimiento), 0, oRsTmpHBT1.Fields!IdPaisNacimiento)
             oDOPaciente.IdPaisProcedencia = IIf(IsNull(oRsTmpHBT1.Fields!IdPaisProcedencia), 0, oRsTmpHBT1.Fields!IdPaisProcedencia)
             oDOPaciente.IdProcedencia = IIf(IsNull(oRsTmpHBT1.Fields!IdProcedencia), 0, oRsTmpHBT1.Fields!IdProcedencia)
             oDOPaciente.IdTipoNumeracion = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoNumeracion), 0, oRsTmpHBT1.Fields!IdTipoNumeracion)
             oDOPaciente.IdTipoOcupacion = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoOcupacion), 0, oRsTmpHBT1.Fields!IdTipoOcupacion)
             oDOPaciente.idTipoSexo = IIf(IsNull(oRsTmpHBT1.Fields!idTipoSexo), 1, oRsTmpHBT1.Fields!idTipoSexo)
             oDOPaciente.IdUsuarioAuditoria = lnIdUsuario
             oDOPaciente.NombreMadre = IIf(IsNull(oRsTmpHBT1.Fields!NombreMadre), "", oRsTmpHBT1.Fields!NombreMadre)
             oDOPaciente.NombrePadre = IIf(IsNull(oRsTmpHBT1.Fields!NombrePadre), "", oRsTmpHBT1.Fields!NombrePadre)
             oDOPaciente.NroDocumento = IIf(IsNull(oRsTmpHBT1.Fields!NroDocumento), "", oRsTmpHBT1.Fields!NroDocumento)
             oDOPaciente.NroHistoriaClinica = IIf(IsNull(oRsTmpHBT1.Fields!NroHistoriaClinica), 0, oRsTmpHBT1.Fields!NroHistoriaClinica)
             oDOPaciente.Observacion = IIf(IsNull(oRsTmpHBT1.Fields!Observacion), "", oRsTmpHBT1.Fields!Observacion)
             oDOPaciente.PrimerNombre = oRsTmpHBT1.Fields!PrimerNombre
             oDOPaciente.SegundoNombre = IIf(IsNull(oRsTmpHBT1.Fields!SegundoNombre), "", oRsTmpHBT1.Fields!SegundoNombre)
             oDOPaciente.Telefono = IIf(IsNull(oRsTmpHBT1.Fields!Telefono), "", oRsTmpHBT1.Fields!Telefono)
             oDOPaciente.TercerNombre = IIf(IsNull(oRsTmpHBT1.Fields!TercerNombre), "", oRsTmpHBT1.Fields!TercerNombre)
             If lnUltimoId < oDOPaciente.idPaciente Then
                If Not InsertarTmpPacientesAgregar(oDOPaciente) Then
                      GoTo Terminar
                End If
             ElseIf Me.optPacienteAct.Value = True Then
                If Not oPaciente.Modificar(oDOPaciente, False) Then
                     ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
                End If
             End If
             'Tabla: HistoriasClinicas
             lcSql = "select * from HistoriasClinicas where idPaciente=" & oDOPaciente.idPaciente
             oRsTmpHBT2.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
             If oRsTmpHBT2.RecordCount = 0 Then
'                If oRsTmpHBT1.Fields!IdTipoNumeracion < 4 Then
'                    oSheet.Cells(j, 1).Value = "..<<falta de dato>>...IdPaciente: " & oDOPaciente.idPaciente & " en tabla HistoriasClinicas (Tipo Numeración:" & Trim(Str(oRsTmpHBT1.Fields!IdTipoNumeracion)) & ")"
'                    oSheet.Cells(j, 6).Value = ""
'                    j = j + 1
'
'                End If
                If oRsTmpHBT1.Fields!IdTipoNumeracion < 4 Then
                    oDOHistoria.FechaCreacion = CDate("01/01/2010")
                    oDOHistoria.FechaPasoAPasivo = 0
                    oDOHistoria.IdEstadoHistoria = 1
                    oDOHistoria.idPaciente = oDOPaciente.idPaciente
                    oDOHistoria.IdTipoHistoria = 1
                    oDOHistoria.IdTipoNumeracion = oDOPaciente.IdTipoNumeracion
                    oDOHistoria.IdTipoNumeracionAnterior = oDOPaciente.IdTipoNumeracion
                    oDOHistoria.IdUsuarioAuditoria = lnIdUsuario
                    oDOHistoria.NroHistoriaClinica = oDOPaciente.NroHistoriaClinica
                    oDOHistoria.NroHistoriaClinicaAnterior = 0
                    If Not InsertarDebbHistorias(oDOHistoria) Then
                       If Val(Left(ml_Errores, 11)) = -2147217873 Then
                            oDOHistoria.NroHistoriaClinica = 2000000 + oDOPaciente.NroHistoriaClinica
                            oDOHistoria.NroHistoriaClinicaAnterior = oDOPaciente.NroHistoriaClinica
                            If Not InsertarDebbHistorias(oDOHistoria) Then
                               GoTo Terminar
                            End If
                            oDOPaciente.NroHistoriaClinica = 2000000 + oDOPaciente.NroHistoriaClinica
                            If Not oPaciente.Modificar(oDOPaciente, False) Then
                                 ms_MensajeError = oPaciente.MensajeError: GoTo Terminar
                            End If
                            oSheet.Cells(j, 1).Value = oDOPaciente.NroHistoriaClinica & "..<<historia duplicada>>...IdPaciente: " & oDOPaciente.idPaciente & "....tipo numero=" & oDOPaciente.IdTipoNumeracion
                            oSheet.Cells(j, 6).Value = ""
                            j = j + 1
                       Else
                         GoTo Terminar
                       End If
                    End If
                End If
             Else
                oDOHistoria.FechaCreacion = oRsTmpHBT2.Fields!FechaCreacion
                oDOHistoria.FechaPasoAPasivo = IIf(IsNull(oRsTmpHBT2.Fields!FechaPasoAPasivo), 0, oRsTmpHBT2.Fields!FechaPasoAPasivo)
                oDOHistoria.IdEstadoHistoria = oRsTmpHBT2.Fields!IdEstadoHistoria
                oDOHistoria.idPaciente = oRsTmpHBT2.Fields!idPaciente
                oDOHistoria.IdTipoHistoria = IIf(IsNull(oRsTmpHBT2.Fields!IdTipoHistoria), 0, oRsTmpHBT2.Fields!IdTipoHistoria)
                oDOHistoria.IdTipoNumeracion = oRsTmpHBT2.Fields!IdTipoNumeracion
                oDOHistoria.IdTipoNumeracionAnterior = IIf(IsNull(oRsTmpHBT2.Fields!IdTipoNumeracionAnterior), 0, oRsTmpHBT2.Fields!IdTipoNumeracionAnterior)
                oDOHistoria.IdUsuarioAuditoria = lnIdUsuario
                oDOHistoria.NroHistoriaClinica = oRsTmpHBT2.Fields!NroHistoriaClinica
                oDOHistoria.NroHistoriaClinicaAnterior = IIf(IsNull(oRsTmpHBT2.Fields!NroHistoriaClinicaAnterior), 0, oRsTmpHBT2.Fields!NroHistoriaClinicaAnterior)
                If lnUltimoId < oDOPaciente.idPaciente Then
                    If Not InsertarDebbHistorias(oDOHistoria) Then
                          GoTo Terminar
                    End If
                Else
                    If Not oHistoria.Modificar(oDOHistoria) Then
                        ms_MensajeError = oHistoria.MensajeError: GoTo Terminar
                    End If
                End If
             End If
             oRsTmpHBT2.Close
             '
             oRsTmpHBT1.MoveNext
          Loop
       End If
       oRsTmpHBT1.Close
       '
       oConexion.CommitTrans
       Me.MousePointer = 1
'       oSheet.SaveAs "c:\estructura.xls"
 '      MsgBox "Se grabó c:\estructura.xls"
       Unload Me
    End If
    Exit Sub
            
Terminar:
    oConexion.RollbackTrans
    MsgBox ms_MensajeError & ml_Errores
    Me.MousePointer = 1
    oSheet.SaveAs "c:\estructura.xls"
    MsgBox "Se grabó c:\estructura.xls"
    Resume
End Sub

Private Sub cmdPercentilSM_Click()

    'On Error GoTo ErrRptHuelga
    On Error Resume Next
    Dim ml_EdadEnMeses As Long
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Dim s As Excel.Worksheet
    Dim W1 As Excel.Workbook
    Dim s1 As Excel.Worksheet
    Dim oRsTmp1 As New Recordset
    Dim oFila As Long, ldFecha As Date, lbNuevo As Boolean
    Dim ldFechaInicialHist As Date, ldFechaFinalHist As Date
    Dim lnNroConsultas As Long, lcFecha As String, lcHoraAtencion As String, lcTexto As String
    Dim oConexion As New Connection
    Dim ml_idTipoSexo As Integer, ldFechaNacimiento As Date, ldFechaAtencion As Date
    Dim lnPeso As Double, lnTalla As Double, lnEdadGest As Integer
    Dim lnEdadEnMesesMasPuntoCinco As Double, lnMinimo As Double, lnMaximo As Double, lnIMC As Double
    Dim lnTallaEnCmMasPuntoCinco As Double
    Dim lnPercentilPE As Double, lnPercentilTE As Double, lnPercentilPT As Double
    Dim lnPercentilIMC As Double, lcPercentilIMC As String

    Const lnPercentilNull As Long = 0
    '
    Set W = EXL.Workbooks.Open(App.Path & "\archivos\percentiles.xls")
    Set s = W.Sheets("IMC")

    '
    Set W1 = EXL.Workbooks.Open(txtExcelSM.Text)
    Set s1 = W1.Sheets("TOTALFB")
    s1.Cells(1, 19).Value = "Percentil"
    oFila = 2
    Do While True
'    If oFila > 50 Then
'         Exit Do
'    End If

             lcSql = Trim(s1.Cells(oFila, 1).Value)
             If Len(lcSql) = 0 Then
                  Exit Do
             End If
             DoEvents
             txtExcelSM.Text = oFila
             Me.Refresh

             lcSql = "1"
             lnPercentilIMC = 0
             lcPercentilIMC = "ERR"
             lnPeso = Val(s1.Cells(oFila, 7).Value)
             lnTalla = Val(s1.Cells(oFila, 8).Value)
             lnEdadGest = Val(s1.Cells(oFila, 11).Value)
             If lnPeso > 0 And lnTalla > 0 Then
                s.Cells(203, 6).Value = lnPeso
                s.Cells(205, 6).Value = Round(lnTalla / 100, 2)
                s.Cells(209, 6).Value = lnEdadGest
                lcSql = "percentil"
                lcPercentilIMC = s.Cells(211, 6).Value
                lcSql = ".."
                lnPercentilIMC = IIf(UCase(Left(lcPercentilIMC, 3)) = "ERR", 0, Val(lcPercentilIMC))
                s1.Cells(oFila, 19).Value = lnPercentilIMC
             End If
             oFila = oFila + 1
    Loop
    '
    'W.Close True
    EXL.Visible = True
    W1.PrintPreview
    Set s = Nothing
    Set s1 = Nothing
    Set W = Nothing
    Set W1 = Nothing
    Set EXL = Nothing
    MsgBox "procesó sin problemas"
    Exit Sub
'ErrRptHuelga:
'    MsgBox Err.Description
'    Resume

'Dim lcLlave As String
'Dim lnImpSubTot As Double: Dim lntImpSubTot As Double
'Dim lnImpAnul As Double: Dim lntImpAnul As Double
'Dim lnImpExo As Double: Dim lntImpExo As Double
'Dim lnImpDevol As Double: Dim lntImpDevol As Double
'Dim lnImpPagCta As Double: Dim lntImpPagCta As Double
'Dim lnImpTot As Double: Dim lntImpTot As Double
'Dim impigv As Double: Dim IGV As Double
'Dim impbruto As Double: Dim inbruto As Double
'Dim lnDctos As Double, lnImpRedondeo As Double, lntImpRedondeo As Double
'Dim iFila As Long
'Dim lRecordCount As Long, lcCadenaConexion As String
'Dim rsReporte As New Recordset
'Dim lcBuscaParametro As New SIGHDatos.Parametros
'Dim mo_ReglasReportes As New ReglasReportes
'Dim CantidadSOAT As Long: Dim PrecioSOAT As Double
'Dim lbEsOpenOffice As Boolean
'Dim lcSql As String
'Dim oConexion As New Connection, lbTienePagoAcuenta As Boolean, lnRedondeo As Double
'
'
'On Error GoTo ManejadorError
'
'
'        Dim oExcel As Excel.Application
'        Dim oWorkSheet As Worksheet

'        Dim oWorkBookPlantilla As Workbook
'        Dim oWorkBook As Workbook

'        Dim oRange As range
'        Dim range As Excel.range
'        Dim borders As Excel.borders
'
'    Set EXL = New Excel.Application
'    Dim W As Excel.Workbook
'    Dim s As Excel.Worksheet
'    Set W = EXL.Workbooks.Open(App.Path & "\archivos\percentiles.xls")
'    Set s = W.Sheets("IMC")
'
'
'            'Crea nueva hoja
'            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
'            Set oWorkBook = oExcel.Workbooks.Add
'            'Abre, copia y cierra la plantilla
'            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ECajaConsolidadoFarm.xls")
'            oWorkBookPlantilla.Worksheets("CajaConsolidadoFarm").Copy Before:=oWorkBook.Sheets(1)
'            oWorkBookPlantilla.Close
'            'Activa la primera hoja
'            Set oWorkSheet = oWorkBook.Sheets(1)
'            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
'            oWorkSheet.Cells(2, 5).Value = "CONSOLIDADO DE FARMACIA"
'            oWorkSheet.Cells(2, 14).Value = lcBuscaParametro.RetornaFechaHoraServidorSQL
'            oWorkSheet.Cells(4, 3).Value = ml_TextoDelFiltro
'
'        iFila = 7
'
'        lRecordCount = 0: lntImpSubTot = 0: lntImpAnul = 0: lntImpExo = 0: lntImpDevol = 0
'        lntImpPagCta = 0: lntImpTot = 0: IGV = 0: impigv = 0: impbruto = 0: inbruto = 0: lntImpRedondeo = 0
'        rsReporte.MoveFirst
'        Do While Not rsReporte.EOF
'                oWorkSheet.Cells(iFila, COL_FECHA).Value = "'" & Format(rsReporte.Fields!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY)
'                oWorkSheet.Cells(iFila, COL_USUARIO).Value = mo_ReglasCaja.SeleccionaDatosCajero(rsReporte.Fields!IdCajero, sghIniciales)
'                oWorkSheet.Cells(iFila, COL_BOLETA).Value = rsReporte!NroSerie + " - " + rsReporte!NroDocumento
'                oWorkSheet.Cells(iFila, COL_NRO_HISTORIA).Value = mo_ReporteUtil.NullToVacio(rsReporte!NroHistoriaClinica)
'                oWorkSheet.Cells(iFila, COL_RAZON_SOCIAL).Value = mo_ReporteUtil.NullToVacio(rsReporte!RazonSocial)
'
'            iFila = iFila + 1
'        Loop
'        iFila = iFila + 1
'
'            oWorkSheet.Cells(iFila, 2).Value = "Cantidad de Documentos: " + Trim(Str(lRecordCount))
'            oWorkSheet.Cells(iFila, COL_SUBTOTAL).Value = lntImpSubTot
'            oWorkSheet.Cells(iFila, COL_REDONDEO).Value = lntImpRedondeo
'            oWorkSheet.Cells(iFila, COL_EXONERADO).Value = lntImpExo
'            oWorkSheet.Cells(iFila, COL_ANULADO).Value = lntImpAnul
'            oWorkSheet.Cells(iFila, COL_DEVOLUCION).Value = lntImpDevol
'            oWorkSheet.Cells(iFila, COL_PAGOCTA).Value = lntImpPagCta
'            oWorkSheet.Cells(iFila, COL_TOTAL_BRUTO).Value = impbruto
'            oWorkSheet.Cells(iFila, COL_IGV).Value = impigv
'            oWorkSheet.Cells(iFila, COL_TOTAL_NETO).Value = lntImpTot
'
'                oExcel.Visible = True
'                oWorkSheet.PrintPreview
'        Dim oExcel As Excel.Application
'        Dim oWorkSheet As Worksheet
'
'        'liberar memoria
'        Set oExcel = Nothing
'        Set oWorkBookPlantilla = Nothing
'        Set oWorkBook = Nothing
'        Set oWorkSheet = Nothing
'
'    Set oConexion = Nothing
'    Set rsReporte = Nothing
'    Set lcBuscaParametro = Nothing
'    Set mo_ReglasReportes = Nothing
'    Exit Sub
'ManejadorError:
'    Select Case Err.Number
'    Case 1004
'        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
'    Case Else
'        MsgBox Err.Description
'    End Select
'    Exit Sub


End Sub

Private Sub cmdPreciosSismedv2_Click()
    Dim oConexionFox As New ADODB.Connection
    Dim oConexion As New ADODB.Connection
    Dim oRsFoxProd As New Recordset
    Dim oRsSeguros As New Recordset
    Dim oRsCatBienes As New Recordset
    Dim lcCodigo As String, lcNombre As String, lcMedTip As String
    Dim lnTipo As Long, lnIdProducto As Long
    Dim lnPv As Double, lnPc As Double, lnPd As Double
    Dim lbSigue As Boolean
    On Error GoTo Terminar
    Me.MousePointer = 11
    oConexionFox.CommandTimeout = 150
    oConexionFox.CursorLocation = adUseServer
    oConexionFox.Open "dsn=" & Me.txtSismedv2.Text
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    '
    oRsSeguros.Open "select * from TiposFinanciamiento where seIngresPrecios=1  and idTipoFinanciamiento<>0 and idTipoFinanciamiento<>1000", oConexion, adOpenKeyset, adLockOptimistic
    oRsFoxProd.Open "SELECT * from xprodu ", oConexionFox, adOpenKeyset, adLockOptimistic
    If oRsFoxProd.RecordCount > 0 Then
       oRsFoxProd.MoveFirst
       Do While Not oRsFoxProd.EOF
             lcCodigo = Trim(oRsFoxProd.Fields!medCod)
             lnTipo = IIf(oRsFoxProd.Fields!medEst = "S", 3, IIf(oRsFoxProd.Fields!medEst = "E", 2, 1))
             lcNombre = Left(Trim(oRsFoxProd.Fields!medNom) & " " & Trim(oRsFoxProd.Fields!medPres) & " " & Trim(oRsFoxProd.Fields!medcnc), 290) & " " & Trim(oRsFoxProd.Fields!medFF)
             lcMedTip = oRsFoxProd.Fields!medTip
             lnPv = oRsFoxProd.Fields!prdPreOpe
             lnPc = oRsFoxProd.Fields!prdPreAdq
             lnPd = oRsFoxProd.Fields!prdPreDist
             'Actualiza Catalogo productos y precios
             oRsCatBienes.Open "select * from factCatalogoBienesInsumos where codigo='" & Trim(lcCodigo) & "'", oConexion, adOpenKeyset, adLockOptimistic
             If oRsCatBienes.RecordCount = 0 Then
                 oRsCatBienes.AddNew
                 oRsCatBienes.Fields!Codigo = lcCodigo
                 oRsCatBienes.Fields!NombreComercial = ""
                 oRsCatBienes.Fields!IdGrupoFarmacologico = 999
                 oRsCatBienes.Fields!IdSubGrupoFarmacologico = 999
             End If
             oRsCatBienes.Fields!nombre = lcNombre
             oRsCatBienes.Fields!IdPartida = Val(txtPartida.Text)
             oRsCatBienes.Fields!IdCentroCosto = Val(Me.txtIdCentroCosto.Text)
             oRsCatBienes.Fields!PrecioCompra = lnPc
             oRsCatBienes.Fields!PrecioDistribucion = lnPd
             oRsCatBienes.Fields!idTipoSalidaBienInsumo = lnTipo
             oRsCatBienes.Fields!TipoProducto = IIf(UCase(lcMedTip) = "M", 0, 1)
             oRsCatBienes.Update
             lnIdProducto = oRsCatBienes.Fields!idProducto
             oRsCatBienes.Close
             'Actualiza SEGUROS
             oRsCatBienes.Open "select * from factCatalogoBienesInsumosHosp where idProducto=" & lnIdProducto, oConexion, adOpenKeyset, adLockOptimistic
             oRsSeguros.MoveFirst
             Do While Not oRsSeguros.EOF
                lbSigue = True
                If oRsCatBienes.RecordCount > 0 Then
                   oRsCatBienes.MoveFirst
                   oRsCatBienes.Find "idTipoFinanciamiento=" & oRsSeguros.Fields!IdTipoFinanciamiento
                   If Not oRsCatBienes.EOF Then
                      lbSigue = False
                   End If
                End If
                If lbSigue = True Then
                    oRsCatBienes.AddNew
                    oRsCatBienes.Fields!idProducto = lnIdProducto
                    oRsCatBienes.Fields!IdTipoFinanciamiento = oRsSeguros.Fields!IdTipoFinanciamiento
                End If
                oRsCatBienes.Fields!PrecioUnitario = lnPv
                oRsCatBienes.Fields!Activo = 1
                oRsCatBienes.Update
                oRsSeguros.MoveNext
             Loop
             oRsCatBienes.Close
             oRsFoxProd.MoveNext
       Loop
    End If
    oConexion.CommitTrans
    Unload Me
    Exit Sub
Terminar:
    MsgBox Err.Description
    oConexion.RollbackTrans
End Sub

Private Sub cmdProcesaAtenciones_Click()
'    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
'       Dim oConexHBT As New Connection
'       Dim oConexion As New Connection
'       Dim oRsTmpHBT1 As New Recordset
'       Dim oRsTmpHBT2 As New Recordset
'       Dim oRsTmpHBT3 As New Recordset
'       Dim oRstmp1 As New Recordset
'       Dim oRsTmp2 As New Recordset
'       Dim oRsTmp3 As New Recordset
'       Dim oDOAtencion As New DOAtencion
'       Dim oAtenciones As New Atenciones
'       Dim oDOCuentaAtencion As New DOCuentaAtencion
'       Dim oCuentaAtencion As New CuentasAtencion
'       Dim oDOAtencionDiagnostico As New DOAtencionDiagnostico
'       Dim oAtencionesDiagnosticos As New AtencionesDiagnosticos
'       Dim oAtencionesEmergencia As New AtencionesEmergencia
'       Dim oDOAtencionEmergencia As New DOAtencionEmergencia
'       Dim oAtencionesEstanciaHosp As New AtencionesEstanciaHosp
'       Dim oDOEstanciaHospitalaria As New DOEstanciaHospitalaria
'       Dim oAtencionesNacimientos As New AtencionesNacimientos
'       Dim oDOAtencionNacimiento As New DOAtencionNacimiento
'       Dim oCitas As New Citas
'       Dim oDoCita As New DOCita
'       Dim lcSql As String, lnCant As Long, lnTotal As Long, lcEstoyEn As String
'       Dim ms_MensajeError As String
'       Dim oExcel As Excel.Application
'       Dim oSheet As Excel.Worksheet
'       Dim j As Integer
'       Dim lnIdCuentaAtencion2020 As Long
'       On Error GoTo Terminar
'       Me.MousePointer = 11
'       ms_MensajeError = ""
'       oConexHBT.Open "dsn=" & txtOdbc.Text
'       oConexion.Open SIGHEntidades.CadenaConexion
'       oConexion.BeginTrans
'       '
'       Set oAtenciones.Conexion = oConexion
'       Set oCuentaAtencion.Conexion = oConexion
'       Set oAtencionesDiagnosticos.Conexion = oConexion
'       Set oAtencionesEmergencia.Conexion = oConexion
'       Set oAtencionesEstanciaHosp.Conexion = oConexion
'       Set oAtencionesNacimientos.Conexion = oConexion
'       Set oCitas.Conexion = oConexion
'       Set mo_conexion = oConexion
'       '
'       lcEstoyEn = ""
'       lcSql = "select * from atenciones where FechaIngreso>='" & txtFinicial.Text & "' and fechaIngreso<='" & txtFfinal.Text & "' order by FechaIngreso"
'       oRsTmpHBT1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
'       lnTotal = oRsTmpHBT1.RecordCount
'       If lnTotal > 0 Then
'          ProgressBar1.Min = 0
'          lnCant = 1
'          'Elimina datos anteriores
'          lcEstoyEn = "Elimina datos anteriores"
'          Me.lblProcesando.Caption = "...Eliminando Datos"
'          lcSql = "select * from atenciones where FechaIngreso>='" & txtFinicial.Text & "' and fechaIngreso<='" & txtFfinal.Text & "'"
'          oRstmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'
'          If oRstmp1.RecordCount > 0 Then
'              ProgressBar1.Max = oRstmp1.RecordCount
'             oRstmp1.MoveFirst
'             Do While Not oRstmp1.EOF
'
'                ProgressBar1.Value = lnCant
'                lnCant = lnCant + 1
'                lcSql = "delete from FacturacionCuentasAtencion where idCuentaAtencion=" & oRstmp1.Fields!idCuentaAtencion
'                oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'                lcSql = "delete from AtencionesDiagnosticos where idAtencion=" & oRstmp1.Fields!idAtencion
'                oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'                lcSql = "delete from AtencionesEmergencia where idAtencion=" & oRstmp1.Fields!idAtencion
'                oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'                lcSql = "delete from AtencionesEstanciaHospitalaria where idAtencion=" & oRstmp1.Fields!idAtencion
'                oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'                lcSql = "delete from AtencionesNacimientos where idAtencion=" & oRstmp1.Fields!idAtencion
'                oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'                lcSql = "delete from Citas where idAtencion=" & oRstmp1.Fields!idAtencion
'                oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'                oRstmp1.Delete
'                oRstmp1.Update
'                oRstmp1.MoveNext
'             Loop
'          End If
'          '
'          Me.lblProcesando.Caption = "...Procesando Datos"
'          ProgressBar1.Max = lnTotal
'          lnCant = 1
'          '
'          Set oExcel = New Excel.Application
'          oExcel.Visible = True
'          oExcel.Workbooks.Add
'          Set oSheet = oExcel.ActiveSheet
'          oSheet.Cells(1, 1).Value = "Error"
'          oSheet.Cells(1, 6).Value = "Observación"
'          j = 3
'          '
'          oRsTmpHBT1.MoveFirst
'          Do While Not oRsTmpHBT1.EOF
'
'             ProgressBar1.Value = lnCant
'             lnCant = lnCant + 1
'             lcSql = "select idPaciente from Pacientes where idPaciente=" & oRsTmpHBT1.Fields!idPaciente
'             oRsTmp3.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'             If oRsTmp3.RecordCount = 0 Then
'                oSheet.Cells(j, 1).Value = "..<<falta de dato>>...IdPaciente: " & oRsTmpHBT1.Fields!idPaciente & " en tabla Pacientes (idAtencion=" & Trim(Str(oRsTmpHBT1.Fields!idAtencion)) & ") (idTipoServicio: " & Trim(Str(oRsTmpHBT1.Fields!IdTipoServicio)) & ")"
'                oSheet.Cells(j, 6).Value = ""
'                j = j + 1
'                oRsTmp3.Close
'             Else
'                 oRsTmp3.Close
'                 'tabla: Cuentas de Atencion
'                 lcEstoyEn = "Cuentas de Atencion"
'                 lcSql = "select * from FacturacionCuentasAtencion where idCuentaAtencion=" & oRsTmpHBT1.Fields!idCuentaAtencion
'                 oRsTmpHBT2.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
'                 If oRsTmpHBT2.RecordCount = 0 Then
''                    oSheet.Cells(j, 1).Value = "..<<falta de dato>>...IdCuentaAtencion: " & oRsTmpHBT1.Fields!idCuentaAtencion & " en tabla FacturacionCuentasAtencion (idTipoServicio: " & Trim(Str(oRsTmpHBT1.Fields!IdTipoServicio)) & ")"
''                    oSheet.Cells(j, 6).Value = ""
''                    j = j + 1
'                    oDOCuentaAtencion.FechaApertura = oRsTmpHBT1.Fields!FechaIngreso
'                    oDOCuentaAtencion.FechaCierre = 0
'                    oDOCuentaAtencion.FechaCreacion = oRsTmpHBT1.Fields!FechaIngreso
'                    oDOCuentaAtencion.HoraApertura = oRsTmpHBT1.Fields!HoraIngreso
'                    oDOCuentaAtencion.HoraCierre = ""
'                    oDOCuentaAtencion.idCuentaAtencion = oRsTmpHBT1.Fields!idAtencion
'                    oDOCuentaAtencion.idEstado = 0
'                    oDOCuentaAtencion.idPaciente = oRsTmpHBT1.Fields!idPaciente
'                    oDOCuentaAtencion.IdUsuarioAuditoria = lnIdUsuario
'                    oDOCuentaAtencion.TotalAsegurado = 0
'                    oDOCuentaAtencion.TotalExonerado = 0
'                    oDOCuentaAtencion.TotalPagado = 0
'                    oDOCuentaAtencion.TotalPorPagar = 0
'                 Else
'                    oDOCuentaAtencion.FechaApertura = IIf(IsNull(oRsTmpHBT2.Fields!FechaApertura), 0, oRsTmpHBT2.Fields!FechaApertura)
'                    oDOCuentaAtencion.FechaCierre = IIf(IsNull(oRsTmpHBT2.Fields!FechaCierre), 0, oRsTmpHBT2.Fields!FechaCierre)
'                    oDOCuentaAtencion.FechaCreacion = IIf(IsNull(oRsTmpHBT2.Fields!FechaCreacion), 0, oRsTmpHBT2.Fields!FechaCreacion)
'                    oDOCuentaAtencion.HoraApertura = IIf(IsNull(oRsTmpHBT2.Fields!HoraApertura), "", oRsTmpHBT2.Fields!HoraApertura)
'                    oDOCuentaAtencion.HoraCierre = IIf(IsNull(oRsTmpHBT2.Fields!HoraCierre), "", oRsTmpHBT2.Fields!HoraCierre)
'                    oDOCuentaAtencion.idCuentaAtencion = oRsTmpHBT1.Fields!idAtencion
'                    oDOCuentaAtencion.idEstado = IIf(IsNull(oRsTmpHBT2.Fields!idEstado), 0, oRsTmpHBT2.Fields!idEstado)
'                    oDOCuentaAtencion.idPaciente = oRsTmpHBT2.Fields!idPaciente
'                    oDOCuentaAtencion.IdUsuarioAuditoria = lnIdUsuario
'                    oDOCuentaAtencion.TotalAsegurado = IIf(IsNull(oRsTmpHBT2.Fields!TotalAsegurado), 0, oRsTmpHBT2.Fields!TotalAsegurado)
'                    oDOCuentaAtencion.TotalExonerado = IIf(IsNull(oRsTmpHBT2.Fields!TotalExonerado), 0, oRsTmpHBT2.Fields!TotalExonerado)
'                    oDOCuentaAtencion.TotalPagado = IIf(IsNull(oRsTmpHBT2.Fields!TotalPagado), 0, oRsTmpHBT2.Fields!TotalPagado)
'                    oDOCuentaAtencion.TotalPorPagar = IIf(IsNull(oRsTmpHBT2.Fields!TotalPorPagar), 0, oRsTmpHBT2.Fields!TotalPorPagar)
'                End If
'                lnErrCA = 0
'                If Not InsertarDebbCuentaAtencion(oDOCuentaAtencion) Then
'                        oSheet.Cells(j, 1).Value = "..<<ya existe>>...IdCuentaAtencion: " & oRsTmpHBT1.Fields!idCuentaAtencion & " en tabla FacturacionCuentasAtencion (idAtencion=" & Trim(Str(oRsTmpHBT1.Fields!idAtencion)) & ") (FechaIngreso: " & oRsTmpHBT1.Fields!FechaIngreso & ")"
'                        oSheet.Cells(j, 6).Value = ""
'                        j = j + 1
'                        GoTo Terminar
'                End If
'                If lnErrCA = 0 Then
'                    'tabla: Atenciones
'                    lcEstoyEn = "Atencion"
'                    With oDOAtencion
'                        '.DireccionDomicilio = IIf(IsNull(oRsTmpHBT1.Fields!DireccionDomicilio), "", oRsTmpHBT1.Fields!DireccionDomicilio)
'                        .Edad = IIf(IsNull(oRsTmpHBT1.Fields!Edad), 0, oRsTmpHBT1.Fields!Edad)
'                        .EsPacienteExterno = False
'                        .FechaEgreso = IIf(IsNull(oRsTmpHBT1.Fields!FechaEgreso), 0, oRsTmpHBT1.Fields!FechaEgreso)
'                        .FechaEgresoAdministrativo = IIf(IsNull(oRsTmpHBT1.Fields!FechaEgresoAdministrativo), 0, oRsTmpHBT1.Fields!FechaEgresoAdministrativo)
'                        .FechaIngreso = oRsTmpHBT1.Fields!FechaIngreso
'                        .HoraEgreso = IIf(IsNull(oRsTmpHBT1.Fields!HoraEgreso), 0, oRsTmpHBT1.Fields!HoraEgreso)
'                        .HoraEgresoAdministrativo = IIf(IsNull(oRsTmpHBT1.Fields!HoraEgresoAdministrativo), 0, oRsTmpHBT1.Fields!HoraEgresoAdministrativo)
'                        .HoraIngreso = oRsTmpHBT1.Fields!HoraIngreso
'                        .HuboInfeccionIntraHospitalaria = IIf(IsNull(oRsTmpHBT1.Fields!HuboInfeccionIntraHospitalaria), 0, oRsTmpHBT1.Fields!HuboInfeccionIntraHospitalaria)
'                        .idAtencion = oRsTmpHBT1.Fields!idAtencion
'                        If IsNull(oRsTmpHBT1.Fields!IdCamaEgreso) Then
'                          If oDOAtencion.IdTipoServicio = 3 Then
'                             .IdCamaEgreso = .IdCamaIngreso
'                          End If
'                        Else
'                          .IdCamaEgreso = oRsTmpHBT1.Fields!IdCamaEgreso
'                        End If
'                        .IdCamaIngreso = IIf(IsNull(oRsTmpHBT1.Fields!IdCamaIngreso), 0, oRsTmpHBT1.Fields!IdCamaIngreso)
'                        .IdCondicionAlta = IIf(IsNull(oRsTmpHBT1.Fields!IdCondicionAlta), 0, oRsTmpHBT1.Fields!IdCondicionAlta)
'                        .idCuentaAtencion = oRsTmpHBT1.Fields!idAtencion
'                        .IdDestinoAtencion = IIf(IsNull(oRsTmpHBT1.Fields!IdDestinoAtencion), 0, oRsTmpHBT1.Fields!IdDestinoAtencion)
'                        .IdEspecialidadMedico = IIf(IsNull(oRsTmpHBT1.Fields!IdEspecialidadMedico), 0, oRsTmpHBT1.Fields!IdEspecialidadMedico)
'                        .IdEstablecimientoDestino = IIf(IsNull(oRsTmpHBT1.Fields!IdEstablecimientoDestino), 0, oRsTmpHBT1.Fields!IdEstablecimientoDestino)
'                        .IdEstablecimientoNoMinsaDestino = IIf(IsNull(oRsTmpHBT1.Fields!IdEstablecimientoNoMinsaDestino), 0, oRsTmpHBT1.Fields!IdEstablecimientoNoMinsaDestino)
'                        .IdEstablecimientoNoMinsaOrigen = IIf(IsNull(oRsTmpHBT1.Fields!IdEstablecimientoNoMinsaOrigen), 0, oRsTmpHBT1.Fields!IdEstablecimientoNoMinsaOrigen)
'                        .IdEstablecimientoOrigen = IIf(IsNull(oRsTmpHBT1.Fields!IdEstablecimientoOrigen), 0, oRsTmpHBT1.Fields!IdEstablecimientoOrigen)
'                        .IdEstadoAtencion = 1  'registrado
'                        .IdFormaPago = lnIdTipoFinanciamiento
'                        .idFuenteFinanciamiento = lnIdFuenteFinanciamiento
'                        If IsNull(oRsTmpHBT1.Fields!IdMedicoEgreso) Then
'                           .IdMedicoEgreso = IIf(IsNull(oRsTmpHBT1.Fields!IdServicioEgreso), 0, oRsTmpHBT1.Fields!IdMedicoIngreso)
'                        Else
'                           .IdMedicoEgreso = oRsTmpHBT1.Fields!IdMedicoEgreso
'                        End If
'                        .IdMedicoIngreso = oRsTmpHBT1.Fields!IdMedicoIngreso
'                        .IdMedicoRespNacimiento = IIf(IsNull(oRsTmpHBT1.Fields!IdMedicoRespNacimiento), 0, oRsTmpHBT1.Fields!IdMedicoRespNacimiento)
'                        .IdOrigenAtencion = IIf(IsNull(oRsTmpHBT1.Fields!IdOrigenAtencion), 0, oRsTmpHBT1.Fields!IdOrigenAtencion)
'                        .idPaciente = oRsTmpHBT1.Fields!idPaciente
'                        .IdServicioEgreso = IIf(IsNull(oRsTmpHBT1.Fields!IdServicioEgreso), 0, oRsTmpHBT1.Fields!IdServicioEgreso)
'                        .IdServicioIngreso = IIf(IsNull(oRsTmpHBT1.Fields!IdServicioIngreso), 0, oRsTmpHBT1.Fields!IdServicioIngreso)
'                        .IdTipoAlta = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoAlta), 0, oRsTmpHBT1.Fields!IdTipoAlta)
'                        .IdTipoCondicionALEstab = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoCondicionALEstab), 0, oRsTmpHBT1.Fields!IdTipoCondicionALEstab)
'                        .IdTipoCondicionAlServicio = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoCondicionAlServicio), 0, oRsTmpHBT1.Fields!IdTipoCondicionAlServicio)
'                        .IdTipoEdad = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoEdad), 0, oRsTmpHBT1.Fields!IdTipoEdad)
'                        .IdTipoGravedad = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoGravedad), 0, oRsTmpHBT1.Fields!IdTipoGravedad)
'                        .IdTipoReferenciaDestino = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoReferenciaDestino), 0, oRsTmpHBT1.Fields!IdTipoReferenciaDestino)
'                        .IdTipoReferenciaOrigen = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoReferenciaOrigen), 0, oRsTmpHBT1.Fields!IdTipoReferenciaOrigen)
'                        .IdTipoServicio = oRsTmpHBT1.Fields!IdTipoServicio
'                        .IdUsuarioAuditoria = lnIdUsuario
'                        '.NombreAcompaniante = IIf(IsNull(oRsTmpHBT1.Fields!NombreAcompaniante), "", oRsTmpHBT1.Fields!NombreAcompaniante)
'                        .NroReferenciaDestino = ""
'                        .NroReferenciaOrigen = ""
'                        '.Observacion = IIf(IsNull(oRsTmpHBT1.Fields!Observacion), "", oRsTmpHBT1.Fields!Observacion)
'                        .PisoDomicilio = IIf(IsNull(oRsTmpHBT1.Fields!PisoDomicilio), "", oRsTmpHBT1.Fields!PisoDomicilio)
'                        .RecienNacido = IIf(IsNull(oRsTmpHBT1.Fields!RecienNacido), 0, oRsTmpHBT1.Fields!RecienNacido)
'                        .TieneNecropsia = IIf(IsNull(oRsTmpHBT1.Fields!TieneNecropsia), 0, oRsTmpHBT1.Fields!TieneNecropsia)
'                    End With
'                    If Not InsertarDebbAtenciones(oDOAtencion) Then
'                          GoTo Terminar
'                    End If
'                    'AtencionesDiagnosticos
'                    lcEstoyEn = "AtencionesDiagnosticos"
'                    lcSql = "select * from AtencionesDiagnosticos where idAtencion=" & oRsTmpHBT1.Fields!idAtencion
'                    oRsTmpHBT3.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
'                    If oRsTmpHBT3.RecordCount > 0 Then
'                       oRsTmpHBT3.MoveFirst
'                       Do While Not oRsTmpHBT3.EOF
'                            With oDOAtencionDiagnostico
'                                .idAtencion = oRsTmpHBT3.Fields!idAtencion
'                                .IdAtencionDiagnostico = oRsTmpHBT3.Fields!IdAtencionDiagnostico
'                                .IdClasificacionDx = IIf(IsNull(oRsTmpHBT3.Fields!IdClasificacionDx), 0, oRsTmpHBT3.Fields!IdClasificacionDx)
'                                .idDiagnostico = IIf(IsNull(oRsTmpHBT3.Fields!idDiagnostico), 0, oRsTmpHBT3.Fields!idDiagnostico)
'                                .IdSubClasificacionDX = IIf(IsNull(oRsTmpHBT3.Fields!IdSubClasificacionDX), 0, oRsTmpHBT3.Fields!IdSubClasificacionDX)
'                                .IdUsuarioAuditoria = lnIdUsuario
'                                .labConfHIS = ""
'                            End With
'                            If Not InsertarDebbAtencionDiagnostico(oDOAtencionDiagnostico) Then
'                                 GoTo Terminar
'                            End If
'                            oRsTmpHBT3.MoveNext
'                       Loop
'                    End If
'                    oRsTmpHBT3.Close
'                    If oDOAtencion.IdTipoServicio > 1 Then
'                        'atencionesEmergencia
'                        lcEstoyEn = "atencionesEmergencia"
'                        lcSql = "select * from AtencionesEmergencia where idAtencion=" & oRsTmpHBT1.Fields!idAtencion
'                        oRsTmpHBT3.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
'                        If oRsTmpHBT3.RecordCount > 0 Then
'                           oRsTmpHBT3.MoveFirst
'                           Do While Not oRsTmpHBT3.EOF
'                                With oDOAtencionEmergencia
'                                    .idAtencion = oRsTmpHBT3.Fields!idAtencion
'                                    .IdAtencionEmergencia = oRsTmpHBT3.Fields!IdAtencionEmergencia
'                                    .IdCausaExternaMorbilidad = IIf(IsNull(oRsTmpHBT3.Fields!IdCausaExternaMorbilidad), 0, oRsTmpHBT3.Fields!IdCausaExternaMorbilidad)
'                                    .IdClaseAccidente = IIf(IsNull(oRsTmpHBT3.Fields!IdClaseAccidente), 0, oRsTmpHBT3.Fields!IdClaseAccidente)
'                                    .IdGrupoOcupacionalALAB = IIf(IsNull(oRsTmpHBT3.Fields!IdGrupoOcupacionalALAB), 0, oRsTmpHBT3.Fields!IdGrupoOcupacionalALAB)
'                                    .IdLugarEvento = IIf(IsNull(oRsTmpHBT3.Fields!IdLugarEvento), 0, oRsTmpHBT3.Fields!IdLugarEvento)
'                                    .IdPosicionLesionadoALAB = IIf(IsNull(oRsTmpHBT3.Fields!IdPosicionLesionadoALAB), 0, oRsTmpHBT3.Fields!IdPosicionLesionadoALAB)
'                                    .IdRelacionAgresorVictima = IIf(IsNull(oRsTmpHBT3.Fields!IdRelacionAgresorVictima), 0, oRsTmpHBT3.Fields!IdRelacionAgresorVictima)
'                                    .IdSeguridad = IIf(IsNull(oRsTmpHBT3.Fields!IdSeguridad), 0, oRsTmpHBT3.Fields!IdSeguridad)
'                                    .IdTipoAgenteAGAN = IIf(IsNull(oRsTmpHBT3.Fields!IdTipoAgenteAGAN), 0, oRsTmpHBT3.Fields!IdTipoAgenteAGAN)
'                                    .IdTipoEvento = IIf(IsNull(oRsTmpHBT3.Fields!IdTipoEvento), 0, oRsTmpHBT3.Fields!IdTipoEvento)
'                                    .IdTipoTransporte = IIf(IsNull(oRsTmpHBT3.Fields!IdTipoTransporte), 0, oRsTmpHBT3.Fields!IdTipoTransporte)
'                                    .IdTipoVehiculo = IIf(IsNull(oRsTmpHBT3.Fields!IdTipoVehiculo), 0, oRsTmpHBT3.Fields!IdTipoVehiculo)
'                                    .IdUbicacionLesionado = IIf(IsNull(oRsTmpHBT3.Fields!IdUbicacionLesionado), 0, oRsTmpHBT3.Fields!IdUbicacionLesionado)
'                                    .IdUsuarioAuditoria = lnIdUsuario
'                                End With
'                                If Not InsertarDebbAtencionEmergencia(oDOAtencionEmergencia) Then
'                                      GoTo Terminar
'                                End If
'                                oRsTmpHBT3.MoveNext
'                           Loop
'                        End If
'                        oRsTmpHBT3.Close
'                        'atencionesEstanciaHospitalaria
'                        lcEstoyEn = "atencionesEstanciaHospitalaria"
'                        lcSql = "select * from AtencionesEstanciaHospitalaria where idAtencion=" & oRsTmpHBT1.Fields!idAtencion
'                        oRsTmpHBT3.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
'                        If oRsTmpHBT3.RecordCount > 0 Then
'                           oRsTmpHBT3.MoveFirst
'                           Do While Not oRsTmpHBT3.EOF
'                                With oDOEstanciaHospitalaria
'                                    .DiasEstancia = IIf(IsNull(oRsTmpHBT3.Fields!DiasEstancia), 0, oRsTmpHBT3.Fields!DiasEstancia)
'                                    .FechaDesocupacion = IIf(IsNull(oRsTmpHBT3.Fields!FechaDesocupacion), 0, oRsTmpHBT3.Fields!FechaDesocupacion)
'                                    .FechaOcupacion = IIf(IsNull(oRsTmpHBT3.Fields!FechaOcupacion), 0, oRsTmpHBT3.Fields!FechaOcupacion)
'                                    .HoraDesocupacion = IIf(IsNull(oRsTmpHBT3.Fields!HoraDesocupacion), "", oRsTmpHBT3.Fields!HoraDesocupacion)
'                                    .HoraOcupacion = IIf(IsNull(oRsTmpHBT3.Fields!HoraOcupacion), "", oRsTmpHBT3.Fields!HoraOcupacion)
'                                    .idAtencion = oRsTmpHBT3.Fields!idAtencion
'                                    .IdCama = IIf(IsNull(oRsTmpHBT3.Fields!IdCama), 0, oRsTmpHBT3.Fields!IdCama)
'                                    .IdEstanciaHospitalaria = oRsTmpHBT3.Fields!IdEstanciaHospitalaria
'                                    .IdFacturacionServicio = IIf(IsNull(oRsTmpHBT3.Fields!IdFacturacionServicio), 0, oRsTmpHBT3.Fields!IdFacturacionServicio)
'                                    .IdMedicoOrdena = IIf(IsNull(oRsTmpHBT3.Fields!IdMedicoOrdena), 0, oRsTmpHBT3.Fields!IdMedicoOrdena)
'                                    .idProducto = 4590
'                                    .idServicio = IIf(IsNull(oRsTmpHBT3.Fields!idServicio), 0, oRsTmpHBT3.Fields!idServicio)
'                                    .IdUsuarioAuditoria = lnIdUsuario
'                                    .LlegoAlServicio = 1
'                                    .Secuencia = oRsTmpHBT3.Fields!Secuencia
'                                End With
'                                If Not InsertarDebbEstanciaHospitalaria(oDOEstanciaHospitalaria) Then
'                                     GoTo Terminar
'                                End If
'                                oRsTmpHBT3.MoveNext
'                           Loop
'                        End If
'                        oRsTmpHBT3.Close
'                        'atencionesNacimientos
'                        lcEstoyEn = "atencionesNacimientos"
'                        lcSql = "select * from AtencionesNacimientos where idAtencion=" & oRsTmpHBT1.Fields!idAtencion
'                        oRsTmpHBT3.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
'                        If oRsTmpHBT3.RecordCount > 0 Then
'                           oRsTmpHBT3.MoveFirst
'                           Do While Not oRsTmpHBT3.EOF
'                                With oDOAtencionNacimiento
'                                    .EdadSemanas = IIf(IsNull(oRsTmpHBT3.Fields!EdadSemanas), 0, oRsTmpHBT3.Fields!EdadSemanas)
'                                    .FechaNacimiento = IIf(IsNull(oRsTmpHBT3.Fields!FechaNacimiento), 0, oRsTmpHBT3.Fields!FechaNacimiento)
'                                    .idAtencion = oRsTmpHBT3.Fields!idAtencion
'                                    .IdCondicionRN = IIf(IsNull(oRsTmpHBT3.Fields!IdCondicionRN), 0, oRsTmpHBT3.Fields!IdCondicionRN)
'                                    .IdNacimiento = oRsTmpHBT3.Fields!IdNacimiento
'                                    .IdTipoSexo = IIf(IsNull(oRsTmpHBT3.Fields!IdTipoSexo), 0, oRsTmpHBT3.Fields!IdTipoSexo)
'                                    .IdUsuarioAuditoria = lnIdUsuario
'                                    .Peso = IIf(IsNull(oRsTmpHBT3.Fields!Peso), 0, oRsTmpHBT3.Fields!Peso)
'                                    .Talla = IIf(IsNull(oRsTmpHBT3.Fields!Talla), 0, oRsTmpHBT3.Fields!Talla)
'                                End With
'                                If Not InsertarDebbAtencionNacimiento(oDOAtencionNacimiento) Then
'                                      GoTo Terminar
'                                End If
'                                oRsTmpHBT3.MoveNext
'                           Loop
'                        End If
'                        oRsTmpHBT3.Close
'                    Else
'                        'citas
'                        lcEstoyEn = "citas"
'                        lcSql = "select * from Citas where idAtencion=" & oRsTmpHBT1.Fields!idAtencion
'                        oRsTmpHBT3.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
'                        If oRsTmpHBT3.RecordCount > 0 Then
'                            With oDoCita
'                                .Fecha = oRsTmpHBT3.Fields!Fecha
'                                .FechaSolicitud = IIf(IsNull(oRsTmpHBT3.Fields!FechaSolicitud), 0, oRsTmpHBT3.Fields!FechaSolicitud)
'                                .HoraFin = oRsTmpHBT3.Fields!HoraFin
'                                .HoraInicio = oRsTmpHBT3.Fields!HoraInicio
'                                .HoraSolicitud = IIf(IsNull(oRsTmpHBT3.Fields!HoraSolicitud), "", oRsTmpHBT3.Fields!HoraSolicitud)
'                                .idAtencion = oRsTmpHBT3.Fields!idAtencion
'                                .IdCita = oRsTmpHBT3.Fields!IdCita
'                                .IdEspecialidad = IIf(IsNull(oRsTmpHBT3.Fields!IdEspecialidad), 0, oRsTmpHBT3.Fields!IdEspecialidad)
'                                .IdEstadoCita = IIf(IsNull(oRsTmpHBT3.Fields!IdEstadoCita), 0, oRsTmpHBT3.Fields!IdEstadoCita)
'                                .IdMedico = oRsTmpHBT3.Fields!IdMedico
'                                .idPaciente = oRsTmpHBT3.Fields!idPaciente
'                                .idProducto = IIf(IsNull(oRsTmpHBT3.Fields!idProducto), 0, oRsTmpHBT3.Fields!idProducto)
'                                .IdProgramacion = oRsTmpHBT3.Fields!IdProgramacion
'                                .idServicio = oRsTmpHBT3.Fields!idServicio
'                                .IdUsuarioAuditoria = lnIdUsuario
'                            End With
'                            If Not InsertarDebbCita(oDoCita) Then
'                                  GoTo Terminar
'                            End If
'                        End If
'                        oRsTmpHBT3.Close
'                    End If
'                End If
'                '
'                oRsTmpHBT2.Close
'             End If
'             '
'             oRsTmpHBT1.MoveNext
'          Loop
'       End If
'       oRsTmpHBT1.Close
'       '
'       lcEstoyEn = "..Todo OK..."
'       oConexion.CommitTrans
'       Me.MousePointer = 1
'       'oSheet.SaveAs "c:\estructura.xls"
'       'MsgBox "Se grabó c:\estructura.xls"
'       Unload Me
'    End If
'    Exit Sub
'Terminar:
'    If ms_MensajeError = "" Then
'       ms_MensajeError = Err.Description
'    End If
'    If MsgBox(ms_MensajeError & Chr(13) & "iva a Grabar en:" & lcEstoyEn & Chr(13) & Chr(13) & "Desea grabar la información hasta el " & Format((oRsTmpHBT1.Fields!FechaIngreso - 1), "dd/mm/yyyy") & " ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
'       oConexion.CommitTrans
'       Me.MousePointer = 1
'       Unload Me
'    Else
'       oConexion.RollbackTrans
'    End If
'    Me.MousePointer = 1
'   ' Resume
End Sub

Private Sub cmdProgramacion_Click()
    If MsgBox("Está seguro ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Dim oConexHBT As New Connection
       Dim oConexion As New Connection
       Dim oRsTmpHBT1 As New Recordset
       Dim oRsTmpHBT2 As New Recordset
       Dim oRsTmpHBT3 As New Recordset
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim oRsTmp3 As New Recordset
       Dim oDOProgramacionMedica  As New DOProgramacionMedica
       Dim oProgramacionMedica As New ProgramacionMedica
       Dim lcSql As String, lnCant As Long, lnTotal As Long
       Dim lnUltimoId As Long
       Dim ms_MensajeError As String
       On Error GoTo Terminar
       Me.MousePointer = 11
       ms_MensajeError = ""
       oConexHBT.Open "dsn=" & txtOdbc.Text
       oConexion.Open SIGHEntidades.CadenaConexion
       oConexion.BeginTrans
       '
       Set oProgramacionMedica.Conexion = oConexion
       Set mo_conexion = oConexion
       '
       lnUltimoId = 0
       lcSql = "select * from ProgramacionMedica order by idProgramacion desc"
       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       If oRsTmp1.RecordCount > 0 Then
           lnUltimoId = oRsTmp1.Fields!IdProgramacion
       End If
       oRsTmp1.Close
       lcSql = "select * from ProgramacionMedica order  by idProgramacion"
       oRsTmpHBT1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
       lnTotal = oRsTmpHBT1.RecordCount
       If lnTotal > 0 Then
          ProgressBar1.Min = 0
          ProgressBar1.Max = lnTotal
          lnCant = 1
          oRsTmpHBT1.MoveFirst
          Do While Not oRsTmpHBT1.EOF
If lnCant > 100000 Then
Exit Do
End If
            ProgressBar1.Value = lnCant
            lnCant = lnCant + 1
            'ProgramacionMedica
            With oDOProgramacionMedica
                .Color = IIf(IsNull(oRsTmpHBT1.Fields!Color), 0, oRsTmpHBT1.Fields!Color)
                .Descripcion = IIf(IsNull(oRsTmpHBT1.Fields!Descripcion), "", oRsTmpHBT1.Fields!Descripcion)
                .Fecha = oRsTmpHBT1.Fields!Fecha
                .HoraFin = oRsTmpHBT1.Fields!HoraFin
                .HoraInicio = oRsTmpHBT1.Fields!HoraInicio
                .IdDepartamento = IIf(IsNull(oRsTmpHBT1.Fields!IdDepartamento), 0, oRsTmpHBT1.Fields!IdDepartamento)
                .IdEspecialidad = IIf(IsNull(oRsTmpHBT1.Fields!IdEspecialidad), 0, oRsTmpHBT1.Fields!IdEspecialidad)
                .idMedico = oRsTmpHBT1.Fields!idMedico
                .IdProgramacion = oRsTmpHBT1.Fields!IdProgramacion
                .IdServicio = IIf(IsNull(oRsTmpHBT1.Fields!IdServicio), 0, oRsTmpHBT1.Fields!IdServicio)
                .IdTipoProgramacion = IIf(IsNull(oRsTmpHBT1.Fields!IdTipoProgramacion), 0, oRsTmpHBT1.Fields!IdTipoProgramacion)
                .IdTipoServicio = oRsTmpHBT1.Fields!IdTipoServicio
                .IdTurno = IIf(IsNull(oRsTmpHBT1.Fields!IdTurno), 0, oRsTmpHBT1.Fields!IdTurno)
                .IdUsuarioAuditoria = lnIdUsuario
            End With
            If lnUltimoId < oDOProgramacionMedica.IdProgramacion Then
                If Not InsertarDebbProgramacionMedicaAgregar(oDOProgramacionMedica) Then
                      GoTo Terminar
                End If
            ElseIf Me.chkProgramacion.Value = 1 Then
                If Not oProgramacionMedica.Modificar(oDOProgramacionMedica) Then
                     ms_MensajeError = oProgramacionMedica.MensajeError: GoTo Terminar
                End If
            End If
            '
            oRsTmpHBT1.MoveNext
          Loop
       End If
       oRsTmpHBT1.Close
       '
       oConexion.CommitTrans
       Me.MousePointer = 1
       Unload Me
    End If
    Exit Sub
            
Terminar:
    oConexion.RollbackTrans
    MsgBox ms_MensajeError
    Me.MousePointer = 1
    Resume
    
End Sub

Private Sub cmdRx_Click()

End Sub




Private Sub cmdSoloFarmacia_Click()

     Dim oConexionMDB As New Connection
     Dim oConexion As New Connection
     Dim oRsTmp1 As New Recordset
     Dim oRsTmp2 As New Recordset
     Dim oRsCitas As New Recordset
     Dim oRsCitasFa As New Recordset
     Dim oRsCitasDe As New Recordset
     Dim oRsProgCab As New Recordset
     Dim oRsProgDet As New Recordset
     Dim oRsProgServ As New Recordset
     Dim oRsRRHH As New Recordset
     
     Dim mo_ReglasFacturacion As New ReglasFacturacion
     Dim mo_ReglasFarmacia As New ReglasFarmacia
     Dim lnMonto As Double, ldFechaIngreso As Date
     Dim lnRegistros As Long, lcRenaes As String, lnHorasProgr As Long, lnHorasCit As Long
     Dim lnIdPaciente As Long, lnMes As Long, lnMontoFarmacia As Double, lnFFSis As Double
     Dim lnFFSoat As Double, lnFFParticular As Double, lnFFConvenio As Double
     Dim lnIdEspecialidad As Long, lnMontoFarmXesp As Double, lcPaciente As String, lcHistoria As String
     Dim lcDNI As String, ldFechaNa As Date, lcSexo As String, lcDpto As String, lcEspecialidad As String
     Dim lcProv As String, lcDist As String, lcEducacio As String, lcIdioma As String
     Dim lcEtnia As String, lnIdMedico As Long, lcMedico As String, lcColegiatura As String, lcCondTrab As String
     Dim lnNroConsultorios As Long, lnThorasProgamadas As Long, lnTlnHorasCitas As Long, lcTipoProf As String
     Dim lnFFSis1 As Integer, lnFFSoat1 As Integer, lnFFParticular1 As Integer, lnFFConvenio1 As Integer
     Dim lcEESSnivel As String, lcEESSnombre As String, lnHorasProgramadas As Long, lnHorasCitas As Long
     Dim lnIdServicio As Long, lcConsultorio As String, lnHorasProg As Integer
     Dim lnEne As Long, lnFeb As Long, lnMar As Long, lnAbr As Long, lnMay As Long, lnJun As Long
     Dim lnJul As Long, lnAgo As Long, lnSet As Long, lnOct As Long, lnNov As Long, lnDic As Long
     Dim lnEne1 As Double, lnFeb1 As Double, lnMar1 As Double, lnAbr1 As Double, lnMay1 As Double, lnJun1 As Double
     Dim lnJul1 As Double, lnAgo1 As Double, lnSet1 As Double, lnOct1 As Double, lnNov1 As Double, lnDic1 As Double
     Dim ldFecha As Date, lcCodigo As String, lnTotal As Double
     Const lnHorasProgramEmerg As Integer = 24
     '
     FileCopy App.Path & "\HojaLibre.xls", App.Path & "\HojaLibre1.xls"

     '
     Dim EXL As Excel.Application
     Set EXL = New Excel.Application
     Dim W As Excel.Workbook
     Set W = EXL.Workbooks.Open(App.Path & "\HojaLibre1.xls")
     Dim s As Excel.Worksheet
     Set s = W.Sheets("Hoja1")
     Dim lnFor As Long, lnFila As Integer, lcRango As String
     '
     
     
     Me.MousePointer = 11
     '
     oConexion.CursorLocation = adUseClient
     oConexion.CommandTimeout = 300
     oConexion.Open SIGHEntidades.CadenaConexion
     oConexionMDB.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\tablasYpa.mdb;"
     '
     lcRenaes = txtRenaes.Text
     lcEESSnivel = txtNivelEESS.Text
     lcEESSnombre = lcBuscaParametro.SeleccionaFilaParametro(205)
     '**************************************************** FARMACIA ********************************************************
        lblProceso.Caption = "(1/1) FARMACIA"
        'Barre movimientos y agrega a tabla MDB
        lcSql = "SELECT   dbo.FactCatalogoBienesInsumos.Codigo, dbo.FactCatalogoBienesInsumos.Nombre, dbo.farmMovimientoDetalle.Cantidad, " & _
"                      dbo.farmMovimientoDetalle.Total , dbo.farmMovimiento.FechaCreacion, dbo.farmMovimiento.idEstadoMovimiento, dbo.farmMovimiento.movTipo, dbo.farmMovimientoDetalle.Precio" & _
" FROM         dbo.farmMovimientoDetalle INNER JOIN" & _
"                      dbo.FactCatalogoBienesInsumos ON dbo.farmMovimientoDetalle.idProducto = dbo.FactCatalogoBienesInsumos.IdProducto INNER JOIN" & _
"                      dbo.farmMovimiento ON dbo.farmMovimientoDetalle.MovNumero = dbo.farmMovimiento.MovNumero AND" & _
"                      dbo.farmMovimientoDetalle.MovTipo = dbo.farmMovimiento.MovTipo LEFT OUTER JOIN" & _
"                      dbo.farmAlmacen ON dbo.farmMovimiento.idAlmacenOrigen = dbo.farmAlmacen.idAlmacen" & _
" WHERE     (dbo.farmMovimiento.idEstadoMovimiento <> 0) AND (dbo.farmMovimiento.MovTipo = 'S')" & _
"         and  dbo.farmAlmacen.idTipoLocales='F' and (YEAR(dbo.farmMovimiento.fechaCreacion) = " & txtAnio.Text & ")" & _
" ORDER BY dbo.FactCatalogoBienesInsumos.Codigo, dbo.farmMovimiento.fechaCreacion"
        
        oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
        lnRegistros = oRsTmp1.RecordCount
        If lnRegistros > 0 Then
           lnFila = 1
           lcRango = "A" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Año: " & Me.txtAnio.Text & "  Establecimiento: " & lcEESSnombre
           lnFila = lnFila + 2
           lcRango = "C" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Ventas (Suma de Cantidades) de todas las Farmacias"
           lcRango = "Q" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Ventas (Suma de Importes) de todas las Farmacias"
           lnFila = lnFila + 2
           lcRango = "A" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Código"
           lcRango = "B" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Medicamento/Insumo"
           lcRango = "C" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Enero"
           lcRango = "D" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Febrero"
           lcRango = "E" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Marzo"
           lcRango = "F" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Abril"
           lcRango = "G" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Mayo"
           lcRango = "H" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Junio"
           lcRango = "I" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Julio"
           lcRango = "J" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Agosto"
           lcRango = "K" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Setiembre"
           lcRango = "L" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Octubre"
           lcRango = "M" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Noviembre"
           lcRango = "N" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Diciembre"
           lcRango = "O" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Total"
           lcRango = "Q" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Enero"
           lcRango = "R" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Febrero"
           lcRango = "S" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Marzo"
           lcRango = "T" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Abril"
           lcRango = "U" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Mayo"
           lcRango = "V" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Junio"
           lcRango = "W" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Julio"
           lcRango = "X" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Agosto"
           lcRango = "Y" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Setiembre"
           lcRango = "Z" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Octubre"
           lcRango = "AA" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Noviembre"
           lcRango = "AB" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Diciembre"
           lcRango = "AC" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Total"
           lnFila = lnFila + 1
           
           ProgressBar1.Min = 0
           ProgressBar1.Max = lnRegistros
           ProgressBar1.Value = 0
           Do While Not oRsTmp1.EOF
              ldFecha = oRsTmp1.Fields!FechaCreacion
              lcCodigo = oRsTmp1.Fields!Codigo
              lcDNI = oRsTmp1.Fields!nombre
              lnEne = 0: lnFeb = 0: lnMar = 0: lnAbr = 0: lnMay = 0: lnJun = 0: lnJul = 0: lnAgo = 0: lnSet = 0: lnOct = 0: lnNov = 0: lnDic = 0
              lnEne1 = 0: lnFeb1 = 0: lnMar1 = 0: lnAbr1 = 0: lnMay1 = 0: lnJun1 = 0: lnJul1 = 0: lnAgo1 = 0: lnSet1 = 0: lnOct1 = 0: lnNov1 = 0: lnDic1 = 0
              Do While Not oRsTmp1.EOF And lcCodigo = oRsTmp1.Fields!Codigo
                    Select Case Month(oRsTmp1.Fields!FechaCreacion)
                    Case 1
                        lnEne = lnEne + oRsTmp1.Fields!cantidad
                        lnEne1 = lnEne1 + oRsTmp1.Fields!Total
                    Case 2
                        lnFeb = lnFeb + oRsTmp1.Fields!cantidad
                        lnFeb1 = lnFeb1 + oRsTmp1.Fields!Total
                    Case 3
                        lnMar = lnMar + oRsTmp1.Fields!cantidad
                        lnMar1 = lnMar1 + oRsTmp1.Fields!Total
                    Case 4
                        lnAbr = lnAbr + oRsTmp1.Fields!cantidad
                        lnAbr1 = lnAbr1 + oRsTmp1.Fields!Total
                    Case 5
                        lnMay = lnMay + oRsTmp1.Fields!cantidad
                        lnMay1 = lnMay1 + oRsTmp1.Fields!Total
                    Case 6
                        lnJun = lnJun + oRsTmp1.Fields!cantidad
                        lnJun1 = lnJun1 + oRsTmp1.Fields!Total
                    Case 7
                        lnJul = lnJul + oRsTmp1.Fields!cantidad
                        lnJul1 = lnJul1 + oRsTmp1.Fields!Total
                    Case 8
                        lnAgo = lnAgo + oRsTmp1.Fields!cantidad
                        lnAgo1 = lnAgo1 + oRsTmp1.Fields!Total
                    Case 9
                        lnSet = lnSet + oRsTmp1.Fields!cantidad
                        lnSet1 = lnSet1 + oRsTmp1.Fields!Total
                    Case 10
                        lnOct = lnOct + oRsTmp1.Fields!cantidad
                        lnOct1 = lnOct1 + oRsTmp1.Fields!Total
                    Case 11
                        lnNov = lnNov + oRsTmp1.Fields!cantidad
                        lnNov1 = lnNov1 + oRsTmp1.Fields!Total
                    Case 12
                        lnDic = lnDic + oRsTmp1.Fields!cantidad
                        lnDic1 = lnDic1 + oRsTmp1.Fields!Total
                    End Select
                    DoEvents: ProgressBar1.Value = ProgressBar1.Value + 1: Me.Refresh
                    oRsTmp1.MoveNext
                    If oRsTmp1.EOF Then
                       Exit Do
                    End If
              Loop
              '
              lcRango = "A" + Trim(Str(lnFila)): s.Range(lcRango).Value = "'" & lcCodigo
              lcRango = "B" + Trim(Str(lnFila)): s.Range(lcRango).Value = lcDNI
              lcRango = "C" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnEne
              lcRango = "D" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnFeb
              lcRango = "E" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnMar
              lcRango = "F" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnAbr
              lcRango = "G" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnMay
              lcRango = "H" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnJun
              lcRango = "I" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnJul
              lcRango = "J" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnAgo
              lcRango = "K" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnSet
              lcRango = "L" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnOct
              lcRango = "M" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnNov
              lcRango = "N" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnDic
              lnTotal = lnEne + lnFeb + lnMar + lnAbr + lnMay + lnJun + lnJul + lnAgo + lnSet + lnOct + lnNov + lnDic
              lcRango = "O" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnTotal
              
              lcRango = "Q" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnEne1
              lcRango = "R" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnFeb1
              lcRango = "S" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnMar1
              lcRango = "T" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnAbr1
              lcRango = "U" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnMay1
              lcRango = "V" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnJun1
              lcRango = "W" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnJul1
              lcRango = "X" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnAgo1
              lcRango = "Y" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnSet1
              lcRango = "Z" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnOct1
              lcRango = "AA" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnNov1
              lcRango = "AB" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnDic1
              lnTotal = lnEne1 + lnFeb1 + lnMar1 + lnAbr1 + lnMay1 + lnJun1 + lnJul1 + lnAgo1 + lnSet1 + lnOct1 + lnNov1 + lnDic1
              lcRango = "AC" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnTotal
              '
              lnFila = lnFila + 1
           Loop
        End If
        oRsTmp1.Close
        Set s = Nothing
        W.Save
        W.Close
        Set W = Nothing
        Set EXL = Nothing
        
    '
    Me.MousePointer = 1
    Unload Me
End Sub

Private Sub cmdStock_Click()

     Dim oConexionMDB As New Connection
     Dim oConexion As New Connection
     Dim oRsTmp1 As New Recordset
     Dim oRsTmp2 As New Recordset
     Dim oRsCitas As New Recordset
     Dim oRsCitasFa As New Recordset
     Dim oRsCitasDe As New Recordset
     Dim oRsProgCab As New Recordset
     Dim oRsProgDet As New Recordset
     Dim oRsProgServ As New Recordset
     Dim oRsRRHH As New Recordset
     
     Dim mo_ReglasFacturacion As New ReglasFacturacion
     Dim mo_ReglasFarmacia As New ReglasFarmacia
     Dim lnMonto As Double, ldFechaIngreso As Date
     Dim lnRegistros As Long, lcRenaes As String, lnHorasProgr As Long, lnHorasCit As Long
     Dim lnIdPaciente As Long, lnMes As Long, lnMontoFarmacia As Double, lnFFSis As Double
     Dim lnFFSoat As Double, lnFFParticular As Double, lnFFConvenio As Double
     Dim lnIdEspecialidad As Long, lnMontoFarmXesp As Double, lcPaciente As String, lcHistoria As String
     Dim lcDNI As String, ldFechaNa As Date, lcSexo As String, lcDpto As String, lcEspecialidad As String
     Dim lcProv As String, lcDist As String, lcEducacio As String, lcIdioma As String
     Dim lcEtnia As String, lnIdMedico As Long, lcMedico As String, lcColegiatura As String, lcCondTrab As String
     Dim lnNroConsultorios As Long, lnThorasProgamadas As Long, lnTlnHorasCitas As Long, lcTipoProf As String
     Dim lnFFSis1 As Integer, lnFFSoat1 As Integer, lnFFParticular1 As Integer, lnFFConvenio1 As Integer
     Dim lcEESSnivel As String, lcEESSnombre As String, lnHorasProgramadas As Long, lnHorasCitas As Long
     Dim lnIdServicio As Long, lcConsultorio As String, lnHorasProg As Integer
     Dim lnEne As Long, lnFeb As Long, lnMar As Long, lnAbr As Long, lnMay As Long, lnJun As Long
     Dim lnJul As Long, lnAgo As Long, lnSet As Long, lnOct As Long, lnNov As Long, lnDic As Long
     Dim lnEne1 As Double, lnFeb1 As Double, lnMar1 As Double, lnAbr1 As Double, lnMay1 As Double, lnJun1 As Double
     Dim lnJul1 As Double, lnAgo1 As Double, lnSet1 As Double, lnOct1 As Double, lnNov1 As Double, lnDic1 As Double
     Dim ldFecha As Date, lcCodigo As String, lnTotal As Double
     Const lnHorasProgramEmerg As Integer = 24
     '
     FileCopy App.Path & "\HojaLibre.xls", App.Path & "\HojaLibre1.xls"
     '
     Dim EXL As Excel.Application
     Set EXL = New Excel.Application
     Dim W As Excel.Workbook
     Set W = EXL.Workbooks.Open(App.Path & "\HojaLibre1.xls")
     Dim s As Excel.Worksheet
     Set s = W.Sheets("Hoja1")
     Dim lnFor As Long, lnFila As Integer, lcRango As String
     '
     
     
     Me.MousePointer = 11
     '
     oConexion.CursorLocation = adUseClient
     oConexion.CommandTimeout = 300
     oConexion.Open SIGHEntidades.CadenaConexion
     oConexionMDB.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\tablasYpa.mdb;"
     '
     lcRenaes = txtRenaes.Text
     lcEESSnivel = txtNivelEESS.Text
     lcEESSnombre = lcBuscaParametro.SeleccionaFilaParametro(205)
     '**************************************************** FARMACIA ********************************************************
        lblProceso.Caption = "(1/1) FARMACIA"
        'Barre movimientos y agrega a tabla MDB
        lcSql = "SELECT    dbo.FactCatalogoBienesInsumos.Codigo, dbo.FactCatalogoBienesInsumos.Nombre, dbo.farmSaldoMensual.SaldoFecha, " & _
                "                      dbo.farmSaldoMensual.Saldo , dbo.FarmAlmacen.idTipoLocales" & _
                " FROM         dbo.farmSaldoMensual LEFT OUTER JOIN" & _
                "                      dbo.farmAlmacen ON dbo.farmSaldoMensual.idAlmacen = dbo.farmAlmacen.idAlmacen LEFT OUTER JOIN" & _
                "                      dbo.FactCatalogoBienesInsumos ON dbo.farmSaldoMensual.idProducto = dbo.FactCatalogoBienesInsumos.IdProducto" & _
                " WHERE     (dbo.farmAlmacen.idTipoLocales = 'F') AND YEAR(dbo.farmSaldoMensual.SaldoFecha) =" & Me.txtAnio.Text & _
                " ORDER BY dbo.FactCatalogoBienesInsumos.Codigo"
        oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
        lnRegistros = oRsTmp1.RecordCount
        If lnRegistros > 0 Then
           lnFila = 1
           lcRango = "A" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Año: " & Me.txtAnio.Text & "  Establecimiento: " & lcEESSnombre
           lnFila = lnFila + 2
           lcRango = "C" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Stock Total (Suma de Saldos) de todas las Farmacias"
           lnFila = lnFila + 2
           lcRango = "A" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Código"
           lcRango = "B" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Medicamento/Insumo"
           lcRango = "C" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Enero"
           lcRango = "D" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Febrero"
           lcRango = "E" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Marzo"
           lcRango = "F" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Abril"
           lcRango = "G" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Mayo"
           lcRango = "H" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Junio"
           lcRango = "I" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Julio"
           lcRango = "J" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Agosto"
           lcRango = "K" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Setiembre"
           lcRango = "L" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Octubre"
           lcRango = "M" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Noviembre"
           lcRango = "N" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Diciembre"
           lcRango = "O" + Trim(Str(lnFila)): s.Range(lcRango).Value = "Total"
           lnFila = lnFila + 1
           
           ProgressBar1.Min = 0
           ProgressBar1.Max = lnRegistros
           ProgressBar1.Value = 0
           Do While Not oRsTmp1.EOF
              lcCodigo = oRsTmp1.Fields!Codigo
              lcDNI = oRsTmp1.Fields!nombre
              lnEne = 0: lnFeb = 0: lnMar = 0: lnAbr = 0: lnMay = 0: lnJun = 0: lnJul = 0: lnAgo = 0: lnSet = 0: lnOct = 0: lnNov = 0: lnDic = 0
              lnEne1 = 0: lnFeb1 = 0: lnMar1 = 0: lnAbr1 = 0: lnMay1 = 0: lnJun1 = 0: lnJul1 = 0: lnAgo1 = 0: lnSet1 = 0: lnOct1 = 0: lnNov1 = 0: lnDic1 = 0
              Do While Not oRsTmp1.EOF And lcCodigo = oRsTmp1.Fields!Codigo
                    Select Case Month(oRsTmp1.Fields!saldoFecha)
                    Case 1
                        lnEne = lnEne + oRsTmp1.Fields!SALDO
                    Case 2
                        lnFeb = lnFeb + oRsTmp1.Fields!SALDO
                    Case 3
                        lnMar = lnMar + oRsTmp1.Fields!SALDO
                    Case 4
                        lnAbr = lnAbr + oRsTmp1.Fields!SALDO
                    Case 5
                        lnMay = lnMay + oRsTmp1.Fields!SALDO
                    Case 6
                        lnJun = lnJun + oRsTmp1.Fields!SALDO
                    Case 7
                        lnJul = lnJul + oRsTmp1.Fields!SALDO
                    Case 8
                        lnAgo = lnAgo + oRsTmp1.Fields!SALDO
                    Case 9
                        lnSet = lnSet + oRsTmp1.Fields!SALDO
                    Case 10
                        lnOct = lnOct + oRsTmp1.Fields!SALDO
                    Case 11
                        lnNov = lnNov + oRsTmp1.Fields!SALDO
                    Case 12
                        lnDic = lnDic + oRsTmp1.Fields!SALDO
                    End Select
                    DoEvents: ProgressBar1.Value = ProgressBar1.Value + 1: Me.Refresh
                    oRsTmp1.MoveNext
                    If oRsTmp1.EOF Then
                       Exit Do
                    End If
              Loop
              '
              lcRango = "A" + Trim(Str(lnFila)): s.Range(lcRango).Value = "'" & lcCodigo
              lcRango = "B" + Trim(Str(lnFila)): s.Range(lcRango).Value = lcDNI
              lcRango = "C" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnEne
              lcRango = "D" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnFeb
              lcRango = "E" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnMar
              lcRango = "F" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnAbr
              lcRango = "G" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnMay
              lcRango = "H" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnJun
              lcRango = "I" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnJul
              lcRango = "J" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnAgo
              lcRango = "K" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnSet
              lcRango = "L" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnOct
              lcRango = "M" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnNov
              lcRango = "N" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnDic
              lnTotal = lnEne + lnFeb + lnMar + lnAbr + lnMay + lnJun + lnJul + lnAgo + lnSet + lnOct + lnNov + lnDic
              lcRango = "O" + Trim(Str(lnFila)): s.Range(lcRango).Value = lnTotal
              
              '
              lnFila = lnFila + 1
           Loop
        End If
        oRsTmp1.Close
        Set s = Nothing
        W.Save
        W.Close
        Set W = Nothing
        Set EXL = Nothing
        
    '
    Me.MousePointer = 1
    Unload Me

End Sub

Private Sub cmdVariaColumPerc_Click()
    On Error GoTo ErrRptHuelga
    'On Error Resume Next
    Dim ml_EdadEnMeses As Long
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Dim s As Excel.Worksheet
    Dim W1 As Excel.Workbook
    Dim s1 As Excel.Worksheet
    Dim oRsTmp1 As New Recordset
    Dim oConexionExterna As New Connection
    Dim oFila As Long, ldFecha As Date, lbNuevo As Boolean
    Dim ldFechaInicialHist As Date, ldFechaFinalHist As Date
    Dim lnNroConsultas As Long, lcFecha As String, lcHoraAtencion As String, lcTexto As String
    Dim oConexion As New Connection
    Dim ml_idTipoSexo As Integer, ldFechaNacimiento As Date, ldFechaAtencion As Date
    Dim lnPeso As Double, lnTalla As Double, lnEdadGest As Integer
    Dim lnEdadEnMesesMasPuntoCinco As Double, lnMinimo As Double, lnMaximo As Double, lnIMC As Double
    Dim lnTallaEnCmMasPuntoCinco As Double
    Dim lnPercentilPE As Double, lnPercentilTE As Double, lnPercentilPT As Double
    Dim lnPercentilIMC As Double, lcPercentilIMC As String
    Dim oCol As Long

    Const lnPercentilNull As Long = 0
    '
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.Open "dsn=GalenhosExterna"
    '
    Set W = EXL.Workbooks.Open(App.Path & "\archivos\percentiles.xls")
    Set s = W.Sheets("IMC")

    '
    Set W1 = EXL.Workbooks.Open(txtExcelSM.Text)
    Set s1 = W1.Sheets("hoja1")
    
    oRsTmp1.Open "select * from totalfb order by idMadre,fecha", oConexionExterna, adOpenKeyset, adLockOptimistic
    
    If oRsTmp1.RecordCount = 0 Then
       MsgBox "No existe informacion"
    Else
       ProgressBar1.Max = oRsTmp1.RecordCount + 1
       ProgressBar1.Min = 0
       oFila = 2
       s1.Cells(oFila, 1).Value = "idMadre"
       s1.Cells(oFila, 2).Value = "EDAD"
       s1.Cells(oFila, 3).Value = "DISTRITO"
       s1.Cells(oFila, 4).Value = "ESTUDIOS"
       For oCol = 0 To 45
           s1.Cells(oFila, oCol + 5).Value = "'" & Trim(Str(oCol))
       Next
       oFila = oFila + 2

       
       Do While Not oRsTmp1.EOF
          s1.Cells(oFila, 1).Value = oRsTmp1!IDMADRE
          s1.Cells(oFila, 2).Value = IIf(IsNull(oRsTmp1!Edad), "", oRsTmp1!Edad)
          s1.Cells(oFila, 3).Value = IIf(IsNull(oRsTmp1!Distrito), "", oRsTmp1!Distrito)
          s1.Cells(oFila, 4).Value = IIf(IsNull(oRsTmp1!ESTUDIOS), "", oRsTmp1!ESTUDIOS)
          lcTexto = oRsTmp1!IDMADRE
          Do While Not oRsTmp1.EOF And lcTexto = oRsTmp1!IDMADRE
                DoEvents: ProgressBar1.Value = ProgressBar1.Value + 1: Me.Refresh
                
                lcSql = "1"
                lnPercentilIMC = 0
                lcPercentilIMC = "ERR"
                lnPeso = oRsTmp1!PESO_HABIT
                lnTalla = oRsTmp1!Talla
                lnEdadGest = oRsTmp1!EDAD_GESTA
                oCol = oRsTmp1!EDAD_GESTA
'                If lnPeso > 0 And lnTalla > 0 Then
'                   s.Cells(203, 6).Value = lnPeso
'                   s.Cells(205, 6).Value = Round(lnTalla / 100, 2)
'                   s.Cells(209, 6).Value = lnEdadGest
'                   lcSql = "percentil"
'                   lcPercentilIMC = s.Cells(211, 6).Value
'                   lcSql = ".."
'                   oCol = IIf(UCase(Left(lcPercentilIMC, 3)) = "ERR", 0, Val(lcPercentilIMC))
'If oCol > 0 Then
'lcSql = "."
'End If
'                End If
                s1.Cells(oFila, oCol + 5).Value = "X"
                oRsTmp1.MoveNext
                If oRsTmp1.EOF Then
                   Exit Do
                End If
          Loop
          oFila = oFila + 1
'If oFila > 100 Then
'Exit Do
'End If
       Loop
       
    End If
    
    EXL.Visible = True
    W1.PrintPreview
    Set s = Nothing
    Set s1 = Nothing
    Set W = Nothing
    Set W1 = Nothing
    Set EXL = Nothing
    MsgBox "procesó sin problemas"
    Exit Sub
ErrRptHuelga:
    MsgBox Err.Description
    Resume
End Sub

Private Sub Command10_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFacturacionServicioFinanciamientos.Delete
        oRsFacturacionServicioFinanciamientos.Update
    End If
ErrElim:

End Sub

Private Sub Command11_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFactOrdenServicioPagos.Delete
        oRsFactOrdenServicioPagos.Update
    End If
ErrElim:


End Sub

Private Sub Command12_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFacturacionServicioPagos.Delete
        oRsFacturacionServicioPagos.Update
    End If
ErrElim:

End Sub

Private Sub Command13_Click()
    oRsFarmMovimientoVentas.AddNew
End Sub

Private Sub Command14_Click()
     oRsFacturacionBienesFinanciamiento.AddNew
End Sub

Private Sub Command15_Click()
   oRsFactOrdenesBienes.AddNew
End Sub

Private Sub Command16_Click()
    oRsFacturacionBienesPagos.AddNew
End Sub

Private Sub Command17_Click()
    oRsCajaComprobantePago.AddNew
End Sub

Private Sub Command18_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFarmMovimientoVentas.Delete
        oRsFarmMovimientoVentas.Update
    End If
ErrElim:

End Sub

Private Sub Command19_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFacturacionBienesFinanciamiento.Delete
        oRsFacturacionBienesFinanciamiento.Update
    End If
ErrElim:

End Sub

Private Sub Command2_Click()
    Dim oRs1 As New Recordset
    Dim oRs2 As New Recordset
    Dim oRs3 As New Recordset
    Dim oRs4 As New Recordset
    'INICIO-Actualiza Precio Venta Farmacia para los nuevos Tarifarios,
    '       en base a la tarifa idTipoFinanciamiento=1
    oRs4.Open "select * from FactCatalogoBienesInsumosHosp where idTipoFinanciamiento=1 and activo=1", SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    oRs1.Open "select * from TiposFinanciamiento where SeIngresPrecios=1 and idTipoFinanciamiento>0", SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    oRs4.MoveFirst
    Do While Not oRs4.EOF
       oRs1.MoveFirst
       Do While Not oRs1.EOF
            lcSql = "select * from FactCatalogoBienesInsumosHosp where idProducto=" & oRs4.Fields!idProducto & " and idTipoFinanciamiento=" & oRs1.Fields!IdTipoFinanciamiento
            oRs2.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
            If oRs2.RecordCount > 0 Then
               oRs2.Fields!PrecioUnitario = oRs4.Fields!PrecioUnitario
               oRs2.Update
            Else
               oRs2.AddNew
               oRs2.Fields!idProducto = oRs4.Fields!idProducto
               oRs2.Fields!IdTipoFinanciamiento = oRs1.Fields!IdTipoFinanciamiento
               oRs2.Fields!Activo = 1
               oRs2.Fields!PrecioUnitario = oRs4.Fields!PrecioUnitario
               oRs2.Update
            End If
            oRs2.Close
            oRs1.MoveNext
       Loop
       oRs4.MoveNext
   Loop
   Unload Me
End Sub



Private Sub Command20_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFactOrdenesBienes.Delete
        oRsFactOrdenesBienes.Update
    End If
ErrElim:

End Sub

Private Sub Command21_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFacturacionBienesPagos.Delete
        oRsFacturacionBienesPagos.Update
    End If
ErrElim:

End Sub

Private Sub Command22_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsCajaComprobantePagoS.Delete
        oRsCajaComprobantePagoS.Update
    End If
ErrElim:

End Sub

Private Sub Command23_Click()
    oRsFactOrdenServicio.AddNew
End Sub

Private Sub Command24_Click()
    oRsFacturacionServicioFinanciamientos.AddNew
End Sub

Private Sub Command25_Click()
    oRsFactOrdenServicioPagos.AddNew
End Sub

Private Sub Command26_Click()
    oRsFacturacionServicioPagos.AddNew
End Sub

Private Sub Command27_Click()
    oRsCajaComprobantePagoS.AddNew
End Sub

Private Sub Command28_Click()
   Dim oRsTmp0 As New Recordset
   Dim oRsTmp1 As New Recordset
   oRsTmp0.Open "select * from FarmMovimientoVentas where idfuenteFinanciamiento=14", SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
   If oRsTmp0.RecordCount > 0 Then
      oRsTmp0.MoveFirst
      Do While Not oRsTmp0.EOF
         lcSql = "update FarmMovimiento set idTipoConcepto=23 where movNumero='" & oRsTmp0.Fields!movNumero & "' and movTipo='" & oRsTmp0.Fields!movTipo & "'"
         oRsTmp1.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
         oRsTmp0.MoveNext
      Loop
   End If
   oRsTmp0.Close
   oRsTmp1.Open "update FuentesFinanciamiento set idTipoConceptoFarmacia=23 where idFuenteFinanciamiento=14", SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
   Unload Me
End Sub

Sub GeneraOactualizaPuntoDeCargaPorCadaCPT(lnIdPuntoCarga As Long, lnIdServicioPuntoCarga As Long)
   Dim lcFiltraCPT As String
   Dim oRsFiltraCPT As New Recordset
   Dim oRsTmp1 As New Recordset
   Dim oRsTmp2 As New Recordset
   Dim DServicio As String
   'Genera nuevo Servicio
   If lnIdServicioPuntoCarga = 0 And (lnIdPuntoCarga = 32 Or lnIdPuntoCarga = 2 Or lnIdPuntoCarga = 31 Or lnIdPuntoCarga = 33 Or lnIdPuntoCarga = 34 Or lnIdPuntoCarga = 35 Or lnIdPuntoCarga = 36 Or lnIdPuntoCarga = 37) Then
        If lnIdPuntoCarga = 32 Then
           DServicio = "Anatomía Patológica"
        End If
        If lnIdPuntoCarga = 2 Then
           DServicio = "Patología Clínica"
        End If
        If lnIdPuntoCarga = 31 Then
           DServicio = "Citología"
        End If
        If lnIdPuntoCarga = 33 Then
           DServicio = "Microbiología"
        End If
        If lnIdPuntoCarga = 34 Then
           DServicio = "Hematología"
        End If
        If lnIdPuntoCarga = 35 Then
           DServicio = "Inmunoserología"
        End If
        If lnIdPuntoCarga = 36 Then
           DServicio = "Urianálisis y Parasitología"
        End If
        If lnIdPuntoCarga = 37 Then
           DServicio = "Bioquímica"
        End If
        lcSql = "select * from Servicios where nombre='" & DServicio & "'"
        oRsFiltraCPT.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
        If oRsFiltraCPT.RecordCount = 0 Then
           oRsFiltraCPT.AddNew
           oRsFiltraCPT.Fields!nombre = DServicio
           oRsFiltraCPT.Fields!IdEspecialidad = 59
           oRsFiltraCPT.Fields!IdTipoServicio = 5
           oRsFiltraCPT.Fields!Codigo = Trim(Str(oRsUltCodigo))
           oRsFiltraCPT.Fields!soloTipoSexo = 0
           oRsFiltraCPT.Fields!maximaEdad = 54750
           oRsFiltraCPT.Fields!idEstado = 1
           oRsFiltraCPT.Update
           oRsUltCodigo = oRsUltCodigo + 1
        End If
        lnIdServicioPuntoCarga = oRsFiltraCPT.Fields!IdServicio
        oRsFiltraCPT.Close
   End If
   'Actualiza IdServicio en tabla 'FactPuntoCarga'
   lcSql = "update FactPuntosCarga set idServicio=" & lnIdServicioPuntoCarga & " where idPuntoCarga=" & lnIdPuntoCarga
   oRsTmp1.Open lcSql, SIGHEntidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
   '
   If lnIdServicioPuntoCarga > 0 And lnIdPuntoCarga > 0 Then
        lcFiltraCPT = "SELECT      dbo.FactCatalogoServicios.IdProducto, dbo.FactCatalogoServicios.Codigo, dbo.FactCatalogoServicios.Nombre, " & _
                 "                      dbo.FactCatalogoServiciosHosp.PrecioUnitario, dbo.FactCatalogoServiciosHosp.Activo, dbo.FactCatalogoServiciosPtos.idPuntoCarga," & _
                 "                      dbo.FactCatalogoServiciosHosp.SeUsaSinPrecio, dbo.FactCatalogoServicios.Nombre AS NombreProducto" & _
                 " FROM         dbo.FactCatalogoServicios RIGHT OUTER JOIN" & _
                 "                      dbo.FactCatalogoServiciosPtos ON dbo.FactCatalogoServicios.IdProducto = dbo.FactCatalogoServiciosPtos.idProducto RIGHT OUTER JOIN" & _
                 "                      dbo.FactCatalogoServiciosHosp ON dbo.FactCatalogoServicios.IdProducto = dbo.FactCatalogoServiciosHosp.IdProducto" & _
                 " WHERE     (dbo.FactCatalogoServiciosPtos.idPuntoCarga = " & lnIdPuntoCarga & ") AND (dbo.FactCatalogoServiciosHosp.IdTipoFinanciamiento = 1) AND" & _
                 "                      (dbo.FactCatalogoServicios.EsCPT = 1)" & _
                 " ORDER BY dbo.FactCatalogoServicios.Nombre"
         oRsTmp1.Open lcFiltraCPT, SIGHEntidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
         If oRsTmp1.RecordCount > 0 Then
            Do While Not oRsTmp1.EOF
               lcSql = "select * from FactCatalogoServiciosPtos where idPuntoCarga=" & lnIdPuntoCarga & " and idProducto=" & oRsTmp1.Fields!idProducto
               oRsTmp2.Open lcSql, SIGHEntidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
               If oRsTmp2.RecordCount > 0 Then
                  oRsTmp2.Fields!EsPreVenta = 1
                  oRsTmp2.Update
               Else
                  oRsTmp2.AddNew
                  oRsTmp2.Fields!idPuntoCarga = lnIdPuntoCarga
                  oRsTmp2.Fields!idProducto = oRsTmp1.Fields!idProducto
                  oRsTmp2.Fields!EsPreVenta = 1
                  oRsTmp2.Update
               End If
               oRsTmp2.Close
               oRsTmp1.MoveNext
            Loop
         End If
         oRsTmp1.Close
    End If
End Sub

Private Sub Command29_Click()
   Dim oRsFiltraCPT As New Recordset
   Dim lnIdServicio As Long, lnIdPuntoCarga As Long, lnIdServicioPuntoCarga As Long
   oRsUltCodigo = 999991
   'Agregar Servicio de "Estadistica" y actualizar en tabla "parametros"
   lcSql = "select * from Servicios where nombre='Estadística'"
   oRsFiltraCPT.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
   If oRsFiltraCPT.RecordCount = 0 Then
      oRsFiltraCPT.AddNew
      oRsFiltraCPT.Fields!nombre = "Estadística"
      oRsFiltraCPT.Fields!IdEspecialidad = 93
      oRsFiltraCPT.Fields!IdTipoServicio = 1
      oRsFiltraCPT.Fields!Codigo = Trim(Str(oRsUltCodigo))
      oRsFiltraCPT.Fields!soloTipoSexo = 3
      oRsFiltraCPT.Fields!maximaEdad = 54750
      oRsFiltraCPT.Fields!idEstado = 1
      oRsFiltraCPT.Update
   End If
   lnIdServicio = oRsFiltraCPT.Fields!IdServicio
   oRsFiltraCPT.Close
   lcSql = "update parametros set valorTexto='" & Trim(Str(lnIdServicio)) & "' where idParametro=256"
   oRsFiltraCPT.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
   oRsUltCodigo = oRsUltCodigo + 1
   'Rx
   lnIdPuntoCarga = 21
   lnIdServicioPuntoCarga = 23
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Tomografia
   lnIdPuntoCarga = 22
   lnIdServicioPuntoCarga = 22
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Ecog General
   lnIdPuntoCarga = 20
   lnIdServicioPuntoCarga = 24
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Ecog.Obstetrica
   lnIdPuntoCarga = 23
   lnIdServicioPuntoCarga = 95
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Banco Sangre
   lnIdPuntoCarga = 38
   lnIdServicioPuntoCarga = 19
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Anat.Patologica
   lnIdPuntoCarga = 32
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Patologia clinica
   lnIdPuntoCarga = 2
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Citologia
   lnIdPuntoCarga = 31
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Microbiologia
   lnIdPuntoCarga = 33
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Hematologia
   lnIdPuntoCarga = 34
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'InmunoSerologia
   lnIdPuntoCarga = 35
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Urianalisis y parasitologia
   lnIdPuntoCarga = 36
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   'Bioquimica
   lnIdPuntoCarga = 37
   lnIdServicioPuntoCarga = 0
   GeneraOactualizaPuntoDeCargaPorCadaCPT lnIdPuntoCarga, lnIdServicioPuntoCarga
   '
   Unload Me
ErrorC7:
End Sub

Private Sub Command3_Click()
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    txtGalenHos.Text = Trim(Str(mo_ReglasFacturacion.RetornaConsumoPacienteServiciosConSeguroPorNroCuenta(Val(txtGalenHos.Text))))
End Sub

Private Sub Command30_Click()
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim lcSerie As String, lcDcto As String, ldFecha As Date
    Dim lbProceso As Boolean
    lcSql = "SELECT      dbo.CajaComprobantesPago.NroSerie, dbo.CajaComprobantesPago.NroDocumento, dbo.FactOrdenesBienes.idPuntoCarga, " & _
            "                      dbo.FacturacionBienesPagos.IdOrden, dbo.FacturacionBienesPagos.IdProducto, dbo.FacturacionBienesPagos.CantidadPagar," & _
            "                      dbo.FacturacionBienesPagos.PrecioVenta, dbo.FacturacionBienesPagos.TotalPagar, dbo.FactOrdenesBienes.idOrden," & _
            "                      dbo.FactOrdenesBienes.idCuentaAtencion, dbo.FactOrdenesBienes.idPreventa, dbo.CajaComprobantesPago.IdTipoOrden," & _
            "                      dbo.CajaComprobantesPago.Total , dbo.CajaComprobantesPago.IdEstadoComprobante, dbo.CajaComprobantesPago.FechaCobranza" & _
            " FROM         dbo.FactOrdenesBienes LEFT OUTER JOIN" & _
            "                      dbo.CajaComprobantesPago ON dbo.FactOrdenesBienes.idComprobantePago = dbo.CajaComprobantesPago.IdComprobantePago LEFT OUTER JOIN" & _
            "                      dbo.FacturacionBienesPagos ON dbo.FactOrdenesBienes.idOrden = dbo.FacturacionBienesPagos.IdOrden" & _
            " Where (dbo.CajaComprobantesPago.IdEstadoComprobante = 9)" & _
            " ORDER BY dbo.CajaComprobantesPago.NroSerie, dbo.CajaComprobantesPago.NroDocumento"
     oRsTmp.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
     If oRsTmp.RecordCount > 0 Then
        With wrs_Gal
            .Fields.Append "Serie", adVarChar, 5, adFldIsNullable
            .Fields.Append "Documento", adVarChar, 20, adFldIsNullable
            .Fields.Append "Fecha", adDate
            .LockType = adLockOptimistic
            .Open
        End With
        oRsTmp.MoveFirst
        Do While Not oRsTmp.EOF
           lcSerie = oRsTmp.Fields!NroSerie
           lcDcto = oRsTmp.Fields!NroDocumento
           ldFecha = oRsTmp.Fields!FechaCobranza
           lbProceso = False
           Do While Not oRsTmp.EOF And lcSerie = oRsTmp.Fields!NroSerie And lcDcto = oRsTmp.Fields!NroDocumento
              If oRsTmp.Fields!idPreventa > 0 Then
                 lcSql = "select * from FarmPreventa where idPreventa=" & oRsTmp.Fields!idPreventa
                 oRsTmp1.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
                 If oRsTmp1.RecordCount > 0 And oRsTmp1.Fields!idEstadoPreventa = 1 Then
                    lbProceso = True
                    lcSql = "update FactORdenesBienes set idComprobantePago=null where idOrden=" & oRsTmp.Fields!IdOrden
                    oRsTmp2.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
                 End If
                 oRsTmp1.Close
              End If
              oRsTmp.MoveNext
              If oRsTmp.EOF Then
                 Exit Do
              End If
           Loop
           If lbProceso = True Then
              wrs_Gal.AddNew
              wrs_Gal.Fields!Serie = lcSerie
              wrs_Gal.Fields!Documento = lcDcto
              wrs_Gal.Fields!Fecha = ldFecha
              wrs_Gal.Update
           End If
        Loop
     End If
     oRsTmp.Close
     Set grdGalenHos.DataSource = wrs_Gal
     MsgBox "Comprueba que la lista de BOLETAS estan bien reparadas"

End Sub

Private Sub Command4_Click()
  wrs_Gal.AddNew
'  wrs_Gal.Update
End Sub


Private Sub Command1_Click()
   If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        wrs_Gal.Delete
        wrs_Gal.Update
        wrs_Gal.Requery
    End If
End Sub

Private Sub Command5_Click()
      Dim oCrypKey As New CrypKey.Util
      MsgBox oCrypKey.DecryptString(txtGalenHos.Text)
End Sub

Private Sub Command6_Click()
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    txtGalenHos.Text = Trim(Str(mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(Val(txtGalenHos.Text))))

End Sub

Private Sub Command7_Click()
    Dim lcNroHistoriaNew As String
    lcNroHistoriaNew = InputBox("ingrese N° Historia NUEVA: ")
    If Val(txtGalenHos.Text) = 0 Then
       MsgBox "ingrese el N° Historia ACTUAL en texto SQL"
       Exit Sub
    End If
    If Val(lcNroHistoriaNew) = 0 Then
       MsgBox "ingrese el N° Historia NUEVA"
       Exit Sub
    End If
    On Error GoTo ErrorC7
    Dim oRsTmp As New Recordset
    Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.BeginTrans
    oRsTmp.Open "update historiasClinicas set NroHistoriaClinica=" & lcNroHistoriaNew & " where nroHistoriaClinica=" & txtGalenHos.Text, oConexion, adOpenKeyset, adLockOptimistic
    oRsTmp.Open "update pacientes set NroHistoriaClinica=" & lcNroHistoriaNew & " where nroHistoriaClinica=" & txtGalenHos.Text, oConexion, adOpenKeyset, adLockOptimistic
    oConexion.CommitTrans
    Exit Sub
ErrorC7:
    oConexion.RollbackTrans
    MsgBox Err.Description
End Sub

Private Sub Command8_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsFactOrdenServicio.Delete
        oRsFactOrdenServicio.Update
    End If
ErrElim:
End Sub

Private Sub Command9_Click()
    If MsgBox("Esta seguro de ELIMINAR", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        On Error GoTo ErrElim
        oRsCajaComprobantePagoS.Delete
        oRsCajaComprobantePagoS.Update
    End If
ErrElim:

End Sub

Private Sub Form_Load()
    On Error Resume Next
    txtGalenHos.Text = "SELECT      dbo.Diagnosticos.CodigoCIE2004, dbo.Diagnosticos.Descripcion, dbo.Atenciones.IdCuentaAtencion, dbo.Pacientes.NroHistoriaClinica, " & _
"                         dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre, dbo.TiposServicio.Descripcion AS TSERVICIO, " & _
"                         dbo.Atenciones.FechaIngreso , dbo.Atenciones.FechaEgreso " & _
" FROM            dbo.Diagnosticos RIGHT OUTER JOIN " & _
"                         dbo.AtencionesDiagnosticos ON dbo.Diagnosticos.IdDiagnostico = dbo.AtencionesDiagnosticos.IdDiagnostico LEFT OUTER JOIN " & _
"                         dbo.Pacientes INNER JOIN " & _
"                         dbo.Atenciones ON dbo.Pacientes.IdPaciente = dbo.Atenciones.IdPaciente RIGHT OUTER JOIN " & _
"                         dbo.TiposServicio ON dbo.Atenciones.IdTipoServicio = dbo.TiposServicio.IdTipoServicio ON dbo.AtencionesDiagnosticos.IdAtencion = dbo.Atenciones.IdAtencion" & _
" where dbo.Atenciones.idEstadoAtencion <>0 AND left(dbo.Diagnosticos.CodigoCIE2004,2)='C9'"
    
    Set grdFarmMovimientoVentas.DataSource = oRsFarmMovimientoVentas
    Set grdCajaComprobantesPago.DataSource = oRsCajaComprobantePago
    Set grdFacturacionBienesFinanciamiento.DataSource = oRsFacturacionBienesFinanciamiento
    Set grdFactOrdenesBienes.DataSource = oRsFactOrdenesBienes
    Set grdFacturacionBienesPagos.DataSource = oRsFacturacionBienesPagos
    
    Set grdFactOrdenServicio.DataSource = oRsFactOrdenServicio
    Set grdCajaComprobantesPagoS.DataSource = oRsCajaComprobantePagoS
    Set grdFacturacionServicioFinanciamientos.DataSource = oRsFacturacionServicioFinanciamientos
    Set grdFactOrdenServicioPagos.DataSource = oRsFactOrdenServicioPagos
    Set grdFacturacionServicioPagos.DataSource = oRsFacturacionServicioPagos
    '
    cmbConsideraciones.AddItem "Consideraciones:"
    cmbConsideraciones.AddItem ""
    cmbConsideraciones.AddItem "1-Se elimina las Atenciones de esas Fechas,"
    cmbConsideraciones.AddItem "  se agrega las Atenciones de esas Fechas. "
    cmbConsideraciones.AddItem "2-Deberá tener actualizado las tablas:     "
    cmbConsideraciones.AddItem "  empleados,Medicos,Especialidades,MedicosEspecialidad,"
    cmbConsideraciones.AddItem "  EstablecimientosNoMinsa,Servicios,"
    cmbConsideraciones.AddItem "  camas,.."
    cmbConsideraciones.AddItem "  FactCatalogoBienesInsumos,FactCatalogoServicios,"
    cmbConsideraciones.AddItem "  FuentesFinanciamiento, FuentesFinanciamientoTarifas,"
    cmbConsideraciones.AddItem "  Turnos"
    cmbConsideraciones.AddItem ""
    cmbConsideraciones.AddItem "3-Quitar Autogenerado de: Atenciones, atencionesEmergencia,,"
    cmbConsideraciones.AddItem "  AtencionesDiagnosticos , atencionesEstanciaHospitalaria"
    cmbConsideraciones.AddItem "  atencionesNacimientos, citas, camas,  EstablecimientosNoMinsa "
    cmbConsideraciones.AddItem "  empleados,    "
    cmbConsideraciones.AddItem "  FacturacionCuentasAtencion, "
    cmbConsideraciones.AddItem "  HistoriasSolicitadas, MovimientosHistoriaClinica, "
    cmbConsideraciones.AddItem "  Medicos, MedicosEspecialidad,"
    cmbConsideraciones.AddItem "  ProgramacionMedica, Pacientes,"
    cmbConsideraciones.AddItem ""
    cmbConsideraciones.AddItem "4-Agrega nuevos Insumos o actualiza sus Datos y precios"
    cmbConsideraciones.AddItem "  * crear ODBC Sismedv2 que apunte a carpeta c:\barrantes"
    cmbConsideraciones.AddItem "  * debe existir el archivo 'c:\barrantes\xprodu.dbf'"
    cmbConsideraciones.AddItem ""
    cmbConsideraciones.AddItem "5-Cuando ya se termine totalmente de migrar:"
    cmbConsideraciones.AddItem "  * ejecutar 'actualiza AUTOGENERADOS.SQL'"
    cmbConsideraciones.AddItem "  * eliminar Proced.Almac. que empiezen con DEBB..."
    cmbConsideraciones.AddItem "  * actualizar correlativos en: GeneradorNroHistoriaClinica"
    '
    CreaTemporal
    '
    txtRenaes.Text = lcBuscaParametro.SeleccionaFilaParametro(280)
    txtNivelEESS.Text = ""
    Me.Caption = Me.Caption & "      EESS= " & lcBuscaParametro.SeleccionaFilaParametro(205)
    If wxVersionSQL = sghVersionBD.sighSql2000 Then
       Me.Caption = Me.Caption & " (BD SQL2000)"
    Else
       Me.Caption = Me.Caption & " (BD SQL2008)"
    End If
    '
    txtF1.Text = Format(Date - 1, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    txtF2.Text = Format(Date - 1, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    '
    txtFweb1.Text = Date
    txtFweb2.Text = Date
End Sub

Sub CreaTemporal()
    On Error GoTo ErrCrea
    Dim oRsTmp As New Recordset, oRsTmp1 As New Recordset, oRsTmp2 As New Recordset, lcDescripcion As String
    Dim oConexODBC As New Connection
    Dim oConexExterna As New Connection
    oConexODBC.Open "dsn=Galenhos"
    oConexExterna.Open "dsn=GalenhosExterna"
    If oRsPatologia.State = 1 Then Set oRsPatologia = Nothing
    With oRsPatologia
          .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Procedimiento", adVarChar, 255, adFldIsNullable
          .Fields.Append "idPuntoCarga", adInteger
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    '
    If oRsFarmacia.State = 1 Then Set oRsFarmacia = Nothing
    With oRsFarmacia
          .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "MedicInsumo", adVarChar, 255, adFldIsNullable
          .Fields.Append "idPuntoCarga", adInteger
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    'llenado
    lcSql = "select * from FuaDefaultsCptFarmacia"
    oRsTmp.Open lcSql, oConexExterna, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          If UCase(Trim(oRsTmp.Fields!Tipo)) = "CPT" Then
             lcSql = "select * from FactCatalogoServicios where codigo='" & oRsTmp.Fields!Codigo & "'"
          Else
             lcSql = "select * from FactCatalogoBienesInsumos where codigo='" & oRsTmp.Fields!Codigo & "'"
          End If
          If oRsTmp1.State = 1 Then oRsTmp1.Close
          oRsTmp1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
          lcDescripcion = lcVacio
          If oRsTmp1.RecordCount > 0 Then
             lcDescripcion = oRsTmp1.Fields!nombre
          End If
          If UCase(Trim(oRsTmp.Fields!Tipo)) = "CPT" Then
             oRsPatologia.AddNew
             oRsPatologia.Fields!Codigo = oRsTmp.Fields!Codigo
             oRsPatologia.Fields!procedimiento = lcDescripcion
             oRsPatologia.Fields!idPuntoCarga = oRsTmp.Fields!idPuntoCarga
             oRsPatologia.Update
          Else
             oRsFarmacia.AddNew
             oRsFarmacia.Fields!Codigo = oRsTmp.Fields!Codigo
             oRsFarmacia.Fields!MedicInsumo = lcDescripcion
             oRsFarmacia.Fields!idPuntoCarga = oRsTmp.Fields!idPuntoCarga
             oRsFarmacia.Update
          End If
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    oRsFarmacia.Sort = "MedicInsumo"
    oRsPatologia.Sort = "Procedimiento"
ErrCrea:
End Sub

Sub LimpiaGrid()
    Set grdFarmMovimientoVentas.DataSource = Nothing
    Set grdCajaComprobantesPago.DataSource = Nothing
    Set grdFacturacionBienesFinanciamiento.DataSource = Nothing
    Set grdFactOrdenesBienes.DataSource = Nothing
    Set grdFacturacionBienesPagos.DataSource = Nothing
End Sub

Sub LimpiaGridS()
    Set grdFactOrdenServicio.DataSource = Nothing
    Set grdCajaComprobantesPagoS.DataSource = Nothing
    Set grdFacturacionServicioFinanciamientos.DataSource = Nothing
    Set grdFactOrdenServicioPagos.DataSource = Nothing
    Set grdFacturacionServicioPagos.DataSource = Nothing

End Sub

Private Sub grdFactOrdenesBienes_DblClick()
        On Error GoTo ErrFOB
        Dim lnLinea As Integer
        lnLinea = 1
        lcSql = "select * from FacturacionBienesPagos where idOrden=" & oRsFactOrdenesBienes.Fields!IdOrden
        oRsFacturacionBienesPagos.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
        Set grdFacturacionBienesPagos.DataSource = oRsFacturacionBienesPagos
        lnLinea = 2
       If oRsFactOrdenesBienes.Fields!idComprobantePago > 0 Then
            lcSql = "select * from CajaComprobantesPago where idComprobantePago=" & oRsFactOrdenesBienes.Fields!idComprobantePago
            oRsCajaComprobantePago.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
            Set grdCajaComprobantesPago.DataSource = oRsCajaComprobantePago
       Else
            Set grdCajaComprobantesPago.DataSource = Nothing
       End If
    Exit Sub
ErrFOB:
   If Err.Number = 3705 Then
      Select Case lnLinea
      Case 1
           oRsFacturacionBienesPagos.Close
      Case 2
           oRsCajaComprobantePago.Close
      End Select
      Resume
   End If
End Sub

Private Sub grdFactOrdenServicio_DblClick()
       On Error GoTo ErrFMV
       Dim lnLinea As Integer
       lnLinea = 1
       lcSql = "select * from FactOrdenServicioPagos where idOrden=" & oRsFactOrdenServicio.Fields!IdOrden
       oRsFactOrdenServicioPagos.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdFactOrdenServicioPagos.DataSource = oRsFactOrdenServicioPagos
       lnLinea = 2
       lcSql = "select * from FacturacionServicioFinanciamientos where idOrden=" & oRsFactOrdenServicio.Fields!IdOrden
       oRsFacturacionServicioFinanciamientos.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdFacturacionServicioFinanciamientos.DataSource = oRsFacturacionServicioFinanciamientos
       
       Set grdCajaComprobantesPagoS.DataSource = Nothing
       Set grdFacturacionServicioPagos.DataSource = Nothing
       
    Exit Sub
ErrFMV:
   If Err.Number = 3705 Then
      Select Case lnLinea
      Case 1
          oRsFactOrdenServicioPagos.Close
      Case 2
          oRsFacturacionServicioFinanciamientos.Close
      End Select
      Resume
   End If
End Sub

Private Sub grdFactOrdenServicioPagos_DblClick()
       On Error GoTo ErrFOB
       Dim lnLinea As Integer
       lnLinea = 1
       lcSql = "select * from FacturacionServicioPagos where idOrdenPago=" & oRsFactOrdenServicioPagos.Fields!idOrdenPago
       oRsFacturacionServicioPagos.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdFacturacionServicioPagos.DataSource = oRsFacturacionServicioPagos
       lnLinea = 2
       If oRsFactOrdenServicioPagos.Fields!idComprobantePago > 0 Then
            lcSql = "select * from CajaComprobantesPago where idComprobantePago=" & oRsFactOrdenServicioPagos.Fields!idComprobantePago
            oRsCajaComprobantePagoS.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
            Set grdCajaComprobantesPagoS.DataSource = oRsCajaComprobantePagoS
       Else
            Set grdCajaComprobantesPagoS.DataSource = Nothing
       End If
    Exit Sub
ErrFOB:
   If Err.Number = 3705 Then
      Select Case lnLinea
      Case 1
           oRsFacturacionServicioPagos.Close
      Case 2
           oRsCajaComprobantePagoS.Close
      End Select
      Resume
   End If

End Sub

Private Sub grdFarmMovimientoVentas_DblClick()
       On Error GoTo ErrFMV
       Dim lnLinea As Integer
       lnLinea = 1
       lcSql = "select * from FactOrdenesBienes where movNumero='" & oRsFarmMovimientoVentas.Fields!movNumero & "' and movTipo='" & oRsFarmMovimientoVentas.Fields!movTipo & "'"
       oRsFactOrdenesBienes.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdFactOrdenesBienes.DataSource = oRsFactOrdenesBienes
       lnLinea = 2
       lcSql = "select * from FacturacionBienesFinanciamientos where movNumero='" & oRsFarmMovimientoVentas.Fields!movNumero & "' and movTipo='" & oRsFarmMovimientoVentas.Fields!movTipo & "'"
       oRsFacturacionBienesFinanciamiento.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdFacturacionBienesFinanciamiento.DataSource = oRsFacturacionBienesFinanciamiento
       
       Set grdCajaComprobantesPago.DataSource = Nothing
       Set grdFacturacionBienesPagos.DataSource = Nothing
    Exit Sub
ErrFMV:
   If Err.Number = 3705 Then
      Select Case lnLinea
      Case 1
          oRsFactOrdenesBienes.Close
      Case 2
          oRsFacturacionBienesFinanciamiento.Close
      End Select
      Resume
   End If
End Sub





Private Sub Text1_Change()

End Sub

Private Sub txtCuentaS_KeyPress(KeyAscii As Integer)
    On Error GoTo errCta
    If KeyAscii = 13 And Val(txtCuentaS.Text) > 0 Then
       LimpiaGridS
       Dim lnLinea As Integer
       lnLinea = 1
       lcSql = "select * from FactOrdenServicio where idCuentaAtencion=" & txtCuentaS.Text
       oRsFactOrdenServicio.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdFactOrdenServicio.DataSource = oRsFactOrdenServicio
       lnLinea = 2
       lcSql = "select * from CajaComprobantesPago where IdTipoOrden=1 and idCuentaAtencion=" & txtCuentaS.Text
       oRsCajaComprobantePagoS.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdCajaComprobantesPagoS.DataSource = oRsCajaComprobantePagoS
    End If
    Exit Sub
errCta:
   If Err.Number = 3705 Then
      Select Case lnLinea
      Case 1
           oRsFactOrdenServicio.Close
      Case 2
           oRsCajaComprobantePagoS.Close
      End Select
      Resume
   End If

End Sub

Private Sub txtGalenHos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       usarSelectGalenHos txtGalenHos.Text
    End If

End Sub

Sub usarSelectGalenHos(txt As String)
    Dim wrs_Prg As New ADODB.Recordset
    On Error GoTo eRRCarga2
    wrs_Gal.Open txt, SIGHEntidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
    Set grdGalenHos.DataSource = wrs_Gal
    Exit Sub
eRRCarga2:
    If Err.Number = 3705 Then
       wrs_Gal.Close
       Resume
    End If
End Sub



Private Sub txtNroCuenta_KeyPress(KeyAscii As Integer)
    On Error GoTo errCta
    If KeyAscii = 13 And Val(txtNroCuenta.Text) > 0 Then
       LimpiaGrid
       Dim lnLinea As Integer
       lnLinea = 1
       Set grdFarmMovimientoVentas.DataSource = Nothing
       lcSql = "select * from FarmMovimientoVentas where idCuentaAtencion=" & txtNroCuenta.Text
       oRsFarmMovimientoVentas.Open lcSql, SIGHEntidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
       Set grdFarmMovimientoVentas.DataSource = oRsFarmMovimientoVentas
       lnLinea = 2
       lcSql = "select * from CajaComprobantesPago where IdTipoOrden<>1 and idCuentaAtencion=" & txtCuentaS.Text
       oRsCajaComprobantePago.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
       Set grdCajaComprobantesPago.DataSource = oRsCajaComprobantePago
    End If
    Exit Sub
errCta:
   If Err.Number = 3705 Then
      Select Case lnLinea
      Case 1
           oRsFarmMovimientoVentas.Close
      Case 2
           oRsCajaComprobantePago.Close
      End Select
      Resume
   End If
End Sub




Function InsertarTmpPacientesAgregar(ByVal oTabla As DOPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarTmpPacientesAgregar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbPacientesAgregar"
           Set oParameter = .CreateParameter("@IdPaisNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdPaisNacimiento = 0, Null, oTabla.IdPaisNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, IIf(oTabla.ApellidoMaterno = "", Null, oTabla.ApellidoMaterno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@DireccionDomicilio", adVarChar, adParamInput, 100, IIf(oTabla.DireccionDomicilio = "", Null, oTabla.DireccionDomicilio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 150, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNumeracion = 0, Null, oTabla.IdTipoNumeracion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaisProcedencia", adInteger, adParamInput, 0, IIf(oTabla.IdPaisProcedencia = 0, Null, oTabla.IdPaisProcedencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oTabla.idPaciente): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, IIf(oTabla.ApellidoPaterno = "", Null, oTabla.ApellidoPaterno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, IIf(oTabla.PrimerNombre = "", Null, oTabla.PrimerNombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SegundoNombre", adVarChar, adParamInput, 40, IIf(oTabla.SegundoNombre = "", Null, oTabla.SegundoNombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TercerNombre", adVarChar, adParamInput, 40, IIf(oTabla.TercerNombre = "", Null, oTabla.TercerNombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 12, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Telefono", adVarChar, adParamInput, 10, IIf(oTabla.Telefono = "", Null, oTabla.Telefono)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Autogenerado", adVarChar, adParamInput, 20, IIf(oTabla.Autogenerado = "", Null, oTabla.Autogenerado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 4, IIf(oTabla.idTipoSexo = 0, Null, oTabla.idTipoSexo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProcedencia", adInteger, adParamInput, 4, IIf(oTabla.IdProcedencia = 0, Null, oTabla.IdProcedencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGradoInstruccion", adInteger, adParamInput, 4, IIf(oTabla.IdGradoInstruccion = 0, Null, oTabla.IdGradoInstruccion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoCivil", adInteger, adParamInput, 4, IIf(oTabla.IdEstadoCivil = 0, Null, oTabla.IdEstadoCivil)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 4, IIf(oTabla.IdDocIdentidad = 0, Null, oTabla.IdDocIdentidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoOcupacion", adInteger, adParamInput, 4, IIf(oTabla.IdTipoOcupacion = 0, Null, oTabla.IdTipoOcupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroPobladoDomicilio", adInteger, adParamInput, 4, IIf(oTabla.IdCentroPobladoDomicilio = 0, Null, oTabla.IdCentroPobladoDomicilio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NombrePadre", adVarChar, adParamInput, 20, IIf(oTabla.NombrePadre = "", Null, oTabla.NombrePadre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NombreMadre", adVarChar, adParamInput, 20, IIf(oTabla.NombreMadre = "", Null, oTabla.NombreMadre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaisDomicilio", adInteger, adParamInput, 4, IIf(oTabla.IdPaisDomicilio = 0, Null, oTabla.IdPaisDomicilio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 4, IIf(oTabla.NroHistoriaClinica = 0, Null, oTabla.NroHistoriaClinica)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroPobladoNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdCentroPobladoNacimiento = 0, Null, oTabla.IdCentroPobladoNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroPobladoProcedencia", adInteger, adParamInput, 0, IIf(oTabla.IdCentroPobladoProcedencia = 0, Null, oTabla.IdCentroPobladoProcedencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDistritoProcedencia", adInteger, adParamInput, 0, IIf(oTabla.IdDistritoProcedencia = 0, Null, oTabla.IdDistritoProcedencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDistritoDomicilio", adInteger, adParamInput, 0, IIf(oTabla.IdDistritoDomicilio = 0, Null, oTabla.IdDistritoDomicilio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDistritoNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdDistritoNacimiento = 0, Null, oTabla.IdDistritoNacimiento)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
   InsertarTmpPacientesAgregar = True
 Exit Function
ManejadorDeError:
   MsgBox Err.Number & " " + Err.Description
Exit Function
End Function

Function InsertarDebbProgramacionMedicaAgregar(ByVal oTabla As DOProgramacionMedica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbProgramacionMedicaAgregar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbProgramacionMedicaAgregar"
           Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidad = 0, Null, oTabla.IdEspecialidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Color", adInteger, adParamInput, 0, IIf(oTabla.Color = 0, Null, oTabla.Color)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, oTabla.IdProgramacion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IIf(oTabla.idMedico = 0, Null, oTabla.idMedico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, IIf(oTabla.IdDepartamento = 0, Null, oTabla.IdDepartamento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoServicio = 0, Null, oTabla.IdTipoServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 100, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoProgramacion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoProgramacion = 0, Null, oTabla.IdTipoProgramacion)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
   InsertarDebbProgramacionMedicaAgregar = True
Exit Function
ManejadorDeError:
   MsgBox Err.Number & " " + Err.Description
Exit Function
End Function


Function InsertarDebbCuentaAtencion(ByVal oTabla As DOCuentaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbCuentaAtencion = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbFacturacionCuentasAtencionAgregar"
           Set oParameter = .CreateParameter("@TotalPorPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPorPagar = 0, Null, oTabla.TotalPorPagar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, IIf(oTabla.idEstado = 0, Null, oTabla.idEstado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPagado", adCurrency, adParamInput, 0, IIf(oTabla.TotalPagado = 0, Null, oTabla.TotalPagado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalAsegurado", adCurrency, adParamInput, 0, IIf(oTabla.TotalAsegurado = 0, Null, oTabla.TotalAsegurado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalExonerado", adCurrency, adParamInput, 0, IIf(oTabla.TotalExonerado = 0, Null, oTabla.TotalExonerado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraCierre", adChar, adParamInput, 5, IIf(oTabla.HoraCierre = "", Null, oTabla.HoraCierre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCierre", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCierre = 0, Null, oTabla.FechaCierre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraApertura", adChar, adParamInput, 5, IIf(oTabla.HoraApertura = "", Null, oTabla.HoraApertura)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaApertura", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaApertura = 0, Null, oTabla.FechaApertura)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.idPaciente = 0, Null, oTabla.idPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oTabla.idCuentaAtencion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
   InsertarDebbCuentaAtencion = True
Exit Function
ManejadorDeError:
   lnErrCA = Err.Number
   If lnErrCA <> -2147217873 Then
      MsgBox Err.Number & " " + Err.Description
   End If
Exit Function
End Function



'Function InsertarDebbAtenciones(ByVal oTabla As DOAtencion) As Boolean
'On Error GoTo ManejadorDeError
'Dim oCommand As New ADODB.Command
'Dim oParameter As ADODB.Parameter
'   InsertarDebbAtenciones = False
'   With oCommand
'       .CommandType = adCmdStoredProc
'       Set .ActiveConnection = mo_conexion
'       .CommandText = "debbAtencionesAgregar"
'           Set oParameter = .CreateParameter("@IdTipoReferenciaDestino", adInteger, adParamInput, 0, IIf(oTabla.IdTipoReferenciaDestino = 0, Null, oTabla.IdTipoReferenciaDestino)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdTipoReferenciaOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdTipoReferenciaOrigen = 0, Null, oTabla.IdTipoReferenciaOrigen)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdEstablecimientoDestino", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoDestino = 0, Null, oTabla.IdEstablecimientoDestino)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdEstablecimientoOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoOrigen = 0, Null, oTabla.IdEstablecimientoOrigen)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@HoraIngreso", adChar, adParamInput, 5, IIf(oTabla.HoraIngreso = "", Null, oTabla.HoraIngreso)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@FechaIngreso", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaIngreso = 0, Null, oTabla.FechaIngreso)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoServicio = 0, Null, oTabla.IdTipoServicio)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.idPaciente = 0, Null, oTabla.idPaciente)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oTabla.idAtencion): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdTipoCondicionALEstab", adInteger, adParamInput, 0, IIf(oTabla.IdTipoCondicionALEstab = 0, Null, oTabla.IdTipoCondicionALEstab)): .Parameters.Append oParameter
'           'Set oParameter = .CreateParameter("@DireccionDomicilio", adVarChar, adParamInput, 50, IIf(oTabla.DireccionDomicilio = "", Null, oTabla.DireccionDomicilio)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@FechaEgresoAdministrativo", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEgresoAdministrativo = 0, Null, oTabla.FechaEgresoAdministrativo)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdMedicoRespNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoRespNacimiento = 0, Null, oTabla.IdMedicoRespNacimiento)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdCamaEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdCamaEgreso = 0, Null, oTabla.IdCamaEgreso)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdCamaIngreso", adInteger, adParamInput, 0, IIf(oTabla.IdCamaIngreso = 0, Null, oTabla.IdCamaIngreso)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdServicioEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdServicioEgreso = 0, Null, oTabla.IdServicioEgreso)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdTipoAlta", adInteger, adParamInput, 0, IIf(oTabla.IdTipoAlta = 0, Null, oTabla.IdTipoAlta)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdCondicionAlta", adInteger, adParamInput, 0, IIf(oTabla.IdCondicionAlta = 0, Null, oTabla.IdCondicionAlta)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdTipoEdad", adInteger, adParamInput, 0, IIf(oTabla.IdTipoEdad = 0, Null, oTabla.IdTipoEdad)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@RecienNacido", adBoolean, adParamInput, 0, IIf(oTabla.RecienNacido = 0, Null, oTabla.RecienNacido)): .Parameters.Append oParameter
'           'Set oParameter = .CreateParameter("@NombreAcompaniante", adVarChar, adParamInput, 30, IIf(oTabla.NombreAcompaniante = "", Null, oTabla.NombreAcompaniante)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdOrigenAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdOrigenAtencion = 0, Null, oTabla.IdOrigenAtencion)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@TieneNecropsia", adBoolean, adParamInput, 0, IIf(oTabla.TieneNecropsia = 0, Null, oTabla.TieneNecropsia)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdDestinoAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdDestinoAtencion = 0, Null, oTabla.IdDestinoAtencion)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@HoraEgresoAdministrativo", adChar, adParamInput, 5, IIf(oTabla.HoraEgresoAdministrativo = "", Null, oTabla.HoraEgresoAdministrativo)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdTipoCondicionAlServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoCondicionAlServicio = 0, Null, oTabla.IdTipoCondicionAlServicio)): .Parameters.Append oParameter
'           'Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 200, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@HoraEgreso", adChar, adParamInput, 5, IIf(oTabla.HoraEgreso = "", Null, oTabla.HoraEgreso)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@FechaEgreso", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEgreso = 0, Null, oTabla.FechaEgreso)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdMedicoEgreso", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoEgreso = 0, Null, oTabla.IdMedicoEgreso)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdEstablecimientoNoMinsaDestino", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoNoMinsaDestino = 0, Null, oTabla.IdEstablecimientoNoMinsaDestino)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdEstablecimientoNoMinsaOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoNoMinsaOrigen = 0, Null, oTabla.IdEstablecimientoNoMinsaOrigen)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@Edad", adInteger, adParamInput, 0, IIf(oTabla.Edad = 0, Null, oTabla.Edad)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdEspecialidadMedico", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidadMedico = 0, Null, oTabla.IdEspecialidadMedico)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdMedicoIngreso", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoIngreso = 0, Null, oTabla.IdMedicoIngreso)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdServicioIngreso", adInteger, adParamInput, 0, IIf(oTabla.IdServicioIngreso = 0, Null, oTabla.IdServicioIngreso)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdTipoGravedad", adInteger, adParamInput, 0, IIf(oTabla.IdTipoGravedad = 0, Null, oTabla.IdTipoGravedad)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.idCuentaAtencion = 0, Null, oTabla.idCuentaAtencion)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@HuboInfeccionIntraHospitalaria", adBoolean, adParamInput, 0, IIf(oTabla.HuboInfeccionIntraHospitalaria = 0, Null, oTabla.HuboInfeccionIntraHospitalaria)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@idFormaPago", adInteger, adParamInput, 4, IIf(oTabla.IdFormaPago = 0, Null, oTabla.IdFormaPago)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 4, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@idEstadoAtencion", adInteger, adParamInput, 4, oTabla.IdEstadoAtencion): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@NroReferenciaOrigen", adVarChar, adParamInput, 20, IIf(oTabla.NroReferenciaOrigen = "", Null, oTabla.NroReferenciaOrigen)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@NroReferenciaDestino", adVarChar, adParamInput, 20, IIf(oTabla.NroReferenciaDestino = "", Null, oTabla.NroReferenciaDestino)): .Parameters.Append oParameter
'           Set oParameter = .CreateParameter("@EsPacienteExterno", adBoolean, adParamInput, 0, IIf(oTabla.EsPacienteExterno = True, 1, 0)): .Parameters.Append oParameter
'       .Execute
'   End With
'   InsertarDebbAtenciones = True
'Exit Function
'ManejadorDeError:
'      MsgBox Err.Number & " " + Err.Description
'Exit Function
'End Function

Function InsertarDebbAtencionDiagnostico(ByVal oTabla As DOAtencionDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbAtencionDiagnostico = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbAtencionesDiagnosticosAgregar"
           Set oParameter = .CreateParameter("@IdSubclasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdSubClasificacionDX = 0, Null, oTabla.IdSubClasificacionDX)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdClasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdClasificacionDx = 0, Null, oTabla.IdClasificacionDx)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionDiagnostico", adInteger, adParamInput, 0, oTabla.IdAtencionDiagnostico): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@labConfHIS", adVarChar, adParamInput, 3, IIf(oTabla.labConfHIS = "", Null, oTabla.labConfHIS)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
   InsertarDebbAtencionDiagnostico = True
Exit Function
ManejadorDeError:
      MsgBox Err.Number & " " + Err.Description
Exit Function
End Function


Function InsertarDebbAtencionEmergencia(ByVal oTabla As DOAtencionEmergencia) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbAtencionEmergencia = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbAtencionesEmergenciaAgregar"
           Set oParameter = .CreateParameter("@IdTipoAgenteAGAN", adInteger, adParamInput, 0, IIf(oTabla.IdTipoAgenteAGAN = 0, Null, oTabla.IdTipoAgenteAGAN)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGrupoOcupacionalALAB", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoOcupacionalALAB = 0, Null, oTabla.IdGrupoOcupacionalALAB)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPosicionLesionadoALAB", adInteger, adParamInput, 0, IIf(oTabla.IdPosicionLesionadoALAB = 0, Null, oTabla.IdPosicionLesionadoALAB)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUbicacionLesionado", adInteger, adParamInput, 0, IIf(oTabla.IdUbicacionLesionado = 0, Null, oTabla.IdUbicacionLesionado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoTransporte", adInteger, adParamInput, 0, IIf(oTabla.IdTipoTransporte = 0, Null, oTabla.IdTipoTransporte)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoVehiculo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoVehiculo = 0, Null, oTabla.IdTipoVehiculo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdClaseAccidente", adInteger, adParamInput, 0, IIf(oTabla.IdClaseAccidente = 0, Null, oTabla.IdClaseAccidente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdRelacionAgresorVictima", adInteger, adParamInput, 0, IIf(oTabla.IdRelacionAgresorVictima = 0, Null, oTabla.IdRelacionAgresorVictima)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdSeguridad", adInteger, adParamInput, 0, IIf(oTabla.IdSeguridad = 0, Null, oTabla.IdSeguridad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoEvento", adInteger, adParamInput, 0, IIf(oTabla.IdTipoEvento = 0, Null, oTabla.IdTipoEvento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdLugarEvento", adInteger, adParamInput, 0, IIf(oTabla.IdLugarEvento = 0, Null, oTabla.IdLugarEvento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCausaExternaMorbilidad", adInteger, adParamInput, 0, IIf(oTabla.IdCausaExternaMorbilidad = 0, Null, oTabla.IdCausaExternaMorbilidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionEmergencia", adInteger, adParamInput, 0, oTabla.IdAtencionEmergencia): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
   InsertarDebbAtencionEmergencia = True
Exit Function
ManejadorDeError:
      MsgBox Err.Number & " " + Err.Description
Exit Function
End Function


Function InsertarDebbEstanciaHospitalaria(ByVal oTabla As DOEstanciaHospitalaria) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbEstanciaHospitalaria = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbAtencionesEstanciaHospitalariaAgregar"
           Set oParameter = .CreateParameter("@DiasEstancia", adDecimal, adParamInput, 5, IIf(oTabla.DiasEstancia = 0, Null, oTabla.DiasEstancia)):
           oParameter.Precision = 8
           oParameter.NumericScale = 2
           .Parameters.Append oParameter
           
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionServicio", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionServicio = 0, Null, oTabla.IdFacturacionServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoOrdena = 0, Null, oTabla.IdMedicoOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCama", adInteger, adParamInput, 0, IIf(oTabla.IdCama = 0, Null, oTabla.IdCama)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraDesocupacion", adChar, adParamInput, 5, IIf(oTabla.HoraDesocupacion = "", Null, oTabla.HoraDesocupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaDesocupacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaDesocupacion = 0, Null, oTabla.FechaDesocupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraOcupacion", adChar, adParamInput, 5, IIf(oTabla.HoraOcupacion = "", Null, oTabla.HoraOcupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaOcupacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaOcupacion = 0, Null, oTabla.FechaOcupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Secuencia", adInteger, adParamInput, 0, IIf(oTabla.Secuencia = 0, Null, oTabla.Secuencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstanciaHospitalaria", adInteger, adParamInput, 0, oTabla.IdEstanciaHospitalaria): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@LlegoAlServicio", adInteger, adParamInput, 0, oTabla.LlegoAlServicio): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
   InsertarDebbEstanciaHospitalaria = True
Exit Function
ManejadorDeError:
      MsgBox Err.Number & " " + Err.Description
Exit Function
End Function


Function InsertarDebbAtencionNacimiento(ByVal oTabla As DOAtencionNacimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbAtencionNacimiento = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbAtencionesNacimientosAgregar"
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCondicionRN", adInteger, adParamInput, 0, IIf(oTabla.IdCondicionRN = 0, Null, oTabla.IdCondicionRN)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.idTipoSexo = 0, Null, oTabla.idTipoSexo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Peso", adDouble, adParamInput, 0, IIf(oTabla.Peso = 0, Null, oTabla.Peso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Talla", adDouble, adParamInput, 0, IIf(oTabla.Talla = 0, Null, oTabla.Talla)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EdadSemanas", adInteger, adParamInput, 0, IIf(oTabla.EdadSemanas = 0, Null, oTabla.EdadSemanas)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdNacimiento", adInteger, adParamInput, 0, oTabla.IdNacimiento): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdNacimiento = .Parameters("@IdNacimiento")
   End With
   InsertarDebbAtencionNacimiento = True
Exit Function
ManejadorDeError:
     MsgBox Err.Number & " " + Err.Description
Exit Function
End Function



Function InsertarDebbCita(ByVal oTabla As DOCita) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbCita = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbCitasAgregar"
           Set oParameter = .CreateParameter("@HoraSolicitud", adChar, adParamInput, 5, IIf(oTabla.HoraSolicitud = "", Null, oTabla.HoraSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaSolicitud", adDBTimeStamp, adParamInput, 8, IIf(oTabla.FechaSolicitud = 0, Null, oTabla.FechaSolicitud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProgramacion", adInteger, adParamInput, 0, IIf(oTabla.IdProgramacion = 0, Null, oTabla.IdProgramacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 4, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(oTabla.HoraFin = "", Null, oTabla.HoraFin)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(oTabla.HoraInicio = "", Null, oTabla.HoraInicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCita", adInteger, adParamInput, 0, oTabla.IdCita): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 4, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoCita", adInteger, adParamInput, 4, IIf(oTabla.IdEstadoCita = 0, Null, oTabla.IdEstadoCita)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 4, IIf(oTabla.idMedico = 0, Null, oTabla.idMedico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 4, IIf(oTabla.IdEspecialidad = 0, Null, oTabla.IdEspecialidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 4, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 4, IIf(oTabla.idPaciente = 0, Null, oTabla.idPaciente)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
   InsertarDebbCita = True
Exit Function
ManejadorDeError:
   MsgBox Err.Number & " " + Err.Description
Exit Function
End Function


Function InsertarDebbMovimientoHistoriaClinica(ByVal oTabla As DOMovimientoHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   InsertarDebbMovimientoHistoriaClinica = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "debbMovimientosHistoriaClinicaAgregar"
           Set oParameter = .CreateParameter("@NroFolios", adInteger, adParamInput, 0, IIf(oTabla.NroFolios = 0, Null, oTabla.NroFolios)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioDestino", adInteger, adParamInput, 0, IIf(oTabla.idServicioDestino = 0, Null, oTabla.idServicioDestino)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdServicioOrigen = 0, Null, oTabla.IdServicioOrigen)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 100, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMotivo", adInteger, adParamInput, 0, IIf(oTabla.idMotivo = 0, Null, oTabla.idMotivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaMovimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaMovimiento = 0, Null, oTabla.FechaMovimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.idPaciente = 0, Null, oTabla.idPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oTabla.idMovimiento): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoRecepcion", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoRecepcion = 0, Null, oTabla.IdEmpleadoRecepcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoTransporte", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoTransporte = 0, Null, oTabla.IdEmpleadoTransporte)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoArchivo", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoArchivo = 0, Null, oTabla.IdEmpleadoArchivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGrupoMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoMovimiento = 0, Null, oTabla.IdGrupoMovimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter

       .Execute
   End With
   InsertarDebbMovimientoHistoriaClinica = True
Exit Function
ManejadorDeError:
   MsgBox Err.Number & " " + Err.Description
Exit Function
End Function

Function InsertarDebbHistorias(ByVal oTabla As DOHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   InsertarDebbHistorias = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_conexion
       .CommandText = "HistoriasClinicasAgregar"
           Set oParameter = .CreateParameter("@IdTipoNumeracionAnterior", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNumeracionAnterior = 0, Null, oTabla.IdTipoNumeracionAnterior)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroHistoriaClinicaAnterior", adInteger, adParamInput, 0, IIf(oTabla.NroHistoriaClinicaAnterior = 0, Null, oTabla.NroHistoriaClinicaAnterior)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNumeracion = 0, Null, oTabla.IdTipoNumeracion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, IIf(oTabla.NroHistoriaClinica = 0, Null, oTabla.NroHistoriaClinica)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaPasoAPasivo", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaPasoAPasivo = 0, Null, oTabla.FechaPasoAPasivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoHistoria", adInteger, adParamInput, 4, IIf(oTabla.IdTipoHistoria = 0, Null, oTabla.IdTipoHistoria)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoHistoria", adInteger, adParamInput, 4, IIf(oTabla.IdEstadoHistoria = 0, Null, oTabla.IdEstadoHistoria)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 4, IIf(oTabla.idPaciente = 0, Null, oTabla.idPaciente)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
          ' oTabla.NroHistoriaClinica = .Parameters("@NroHistoriaClinica")
   End With
 
   InsertarDebbHistorias = True
 
Exit Function
ManejadorDeError:
   ml_Errores = Err.Number & " " & Err.Description
Exit Function
End Function


Sub EliminaProcedAlmacenados()
    Dim oProcedimiento As ADOX.Procedure
    Dim oCatalogo As New ADOX.Catalog
    Dim oRsTmp As New Recordset
    Dim lcSql As String, sNombre As String
    Dim oConexODBC1 As New Connection
    On Error GoTo ErrEPA
    'PA de sigh
    oConexODBC1.Open "dsn=GALENHOS"
    oCatalogo.ActiveConnection = oConexODBC1
    For Each oProcedimiento In oCatalogo.Procedures
       If Left(oProcedimiento.Name, 3) <> "dt_" Then
            sNombre = Left(oProcedimiento.Name, InStr(oProcedimiento.Name, ";") - 1)
            lcSql = "DROP PROCEDURE " & sNombre
            oRsTmp.Open lcSql, oConexODBC1, adOpenKeyset, adLockOptimistic
       End If
    Next
    'PA de sigh_externa
    oConexODBC1.Close
    oConexODBC1.Open "dsn=GalenhosExterna"
    oCatalogo.ActiveConnection = oConexODBC1
    For Each oProcedimiento In oCatalogo.Procedures
       If Left(oProcedimiento.Name, 3) <> "dt_" Then
            sNombre = Left(oProcedimiento.Name, InStr(oProcedimiento.Name, ";") - 1)
            lcSql = "DROP PROCEDURE " & sNombre
            oRsTmp.Open lcSql, oConexODBC1, adOpenKeyset, adLockOptimistic
       End If
    Next
    oConexODBC1.Close
    Set oConexODBC1 = Nothing
    '
    Exit Sub
ErrEPA:
    'MsgBox "Error al ELIMINAR PROCEDIMIENTOS ALMACENADOS" & Chr(13) & Err.Number & " - " & Err.Description
    Resume Next
End Sub



Private Sub CommandCuatro_Click()
    On Error GoTo ErSip2000
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Dim s As Excel.Worksheet
    Set W = EXL.Workbooks.Open(App.Path & "\archivos\percentiles.xls")       'usa
    Set s = W.Sheets("IMC")
    '
    Dim oConexionMDB As New Connection, oRsMDB As New Recordset
    oConexionMDB.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\tablasYpa.mdb;"
    '
    Dim lcImadre As String, lcHC As String, lcEstablecim As String, lcEdad As String, lcDistrito As String
    Dim lcEstudios As String, lcTalla As String, lcFecha As String, lcTalla_rn As String, lcPeso_rn As String
    Dim Peso1 As String, Peso2 As String, Peso3 As String, Peso4 As String, Peso5 As String
    Dim Peso6 As String, Peso7 As String, Peso8 As String, Peso9 As String, Peso10 As String
    Dim Peso11 As String, Peso12 As String, Peso13 As String, Peso14 As String, Peso15 As String
    Dim Peso16 As String, Peso17 As String, Peso18 As String, Peso19 As String, Peso20 As String
    Dim Peso21 As String, Peso22 As String, Peso23 As String, Peso24 As String, Peso25 As String
    Dim Peso26 As String, Peso27 As String, Peso28 As String, Peso29 As String, Peso30 As String
    Dim Peso31 As String, Peso32 As String, Peso33 As String, Peso34 As String, Peso35 As String
    Dim Peso36 As String, Peso37 As String, Peso38 As String, Peso39 As String, Peso40 As String
    Dim Peso41 As String, Peso42 As String, Peso43 As String, Peso44 As String, Peso45 As String
    Dim lbNuevo As Boolean, lnFor1 As Integer
    Dim lnFor As Long, lnFila As Long, lcRango As String, lnFilaFinal As Long
    Dim oRsFox1 As New Recordset, lnPercentilIMC As Double, lcPercentilIMC As String
    Dim lcEdadG As String, ldFecha As Date, ldFecEmbarazo As Date
    '
    Dim oConexionFox As New Connection
    oConexionFox.CommandTimeout = 300
    oConexionFox.Open "DSN=his"
    
    
    lcSql = "delete from sip"
    If oRsFox1.State = 1 Then oRsFox1.Close
    oRsFox1.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    '
         lcSql = "select * from princip"
         oRsMDB.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
         '
         lnFila = 1
         lnFilaFinal = oRsMDB.RecordCount
         oRsMDB.MoveFirst
         ProgressBar1.Min = 0: ProgressBar1.Max = lnFilaFinal
         For lnFor = lnFila To lnFilaFinal
             ProgressBar1.Value = lnFor
             lcImadre = Left(oRsMDB.Fields!IDMADRE & Space(30), 30)
If Trim(lcImadre) = "000X2AWB4763" Then
lcSql = ""
End If
             If Trim(lcImadre) = "" Then
                Exit For
             End If
             lcHC = oRsMDB.Fields!HC
             lcEstablecim = oRsMDB.Fields!ESTABLECIM
             lcEdad = oRsMDB.Fields!Edad
             lcDistrito = IIf(IsNull(oRsMDB.Fields!Distrito), "", oRsMDB.Fields!Distrito)
             lcEstudios = oRsMDB.Fields!ESTUDIOS
             lcTalla = oRsMDB.Fields!Talla
             lcTalla_rn = oRsMDB.Fields!talla_ra
             lcPeso_rn = oRsMDB.Fields!peso_rn
             lcEdadG = oRsMDB.Fields!EDAD_GESTA
             If SIGHEntidades.EsFecha(oRsMDB.Fields!Fecha, "DD/MM/AAAA") = True Then
                ldFecha = oRsMDB.Fields!Fecha
             End If
             Peso1 = 0: Peso2 = 0: Peso3 = 0: Peso4 = 0: Peso5 = 0
             Peso6 = 0: Peso7 = 0: Peso8 = 0: Peso9 = 0: Peso10 = 0
             Peso11 = 0: Peso12 = 0: Peso13 = 0: Peso14 = 0: Peso15 = 0
             Peso16 = 0: Peso17 = 0: Peso18 = 0: Peso19 = 0: Peso20 = 0
             Peso21 = 0: Peso22 = 0: Peso23 = 0: Peso24 = 0: Peso25 = 0
             Peso26 = 0: Peso27 = 0: Peso28 = 0: Peso29 = 0: Peso30 = 0
             Peso31 = 0: Peso32 = 0: Peso33 = 0: Peso34 = 0: Peso35 = 0
             Peso36 = 0: Peso37 = 0: Peso38 = 0: Peso39 = 0: Peso40 = 0
             Peso41 = 0: Peso42 = 0: Peso43 = 0: Peso44 = 0: Peso45 = 0
             '
             lcSql = ".."
             lnPercentilIMC = 0
             lcPercentilIMC = "ERR"
             If oRsMDB.Fields!Peso > 0 And oRsMDB.Fields!Talla > 0 Then
                s.Cells(203, 6).Value = oRsMDB.Fields!Peso
                s.Cells(205, 6).Value = Round(oRsMDB.Fields!Talla / 100, 2)
                s.Cells(209, 6).Value = lcEdadG
                lcSql = "percentil"
                lcPercentilIMC = s.Cells(211, 6).Value
                lcSql = ".."
                lnPercentilIMC = IIf(UCase(Left(lcPercentilIMC, 3)) = "ERR", 0, Val(lcPercentilIMC))
             End If
             '
             ldFecEmbarazo = ldFecha - (Val(lcEdadG) * 7)
             Select Case lcEdadG
             Case "1"
                     Peso1 = lnPercentilIMC         'oRsMDB.Fields!Peso
                     
             Case "2"
                     Peso2 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "3"
                     Peso3 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "4"
                     Peso4 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "5"
                     Peso5 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "6"
                     Peso6 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "7"
                     Peso7 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "8"
                     Peso8 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "9"
                     Peso9 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "10"
                     Peso10 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "11"
                     Peso11 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "12"
                     Peso12 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "13"
                     Peso13 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "14"
                     Peso14 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "15"
                     Peso15 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "16"
                     Peso16 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "17"
                     Peso17 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "18"
                     Peso18 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "19"
                     Peso19 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "20"
                     Peso20 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "21"
                     Peso21 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "22"
                     Peso22 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "23"
                     Peso23 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "24"
                     Peso24 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "25"
                     Peso25 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "26"
                     Peso26 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "27"
                     Peso27 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "28"
                     Peso28 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "29"
                     Peso29 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "30"
                     Peso30 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "31"
                     Peso31 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "32"
                     Peso32 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "33"
                     Peso33 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "34"
                     Peso34 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "35"
                     Peso35 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "36"
                     Peso36 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "37"
                     Peso37 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "38"
                     Peso38 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "39"
                     Peso39 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "40"
                     Peso40 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "41"
                     Peso41 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "42"
                     Peso42 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "43"
                     Peso43 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "44"
                     Peso44 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "45"
                     Peso45 = lnPercentilIMC         'oRsMDB.Fields!Peso
             End Select
If Trim(lcImadre) = "1513KULW7812" Then
lcSql = ""
End If
             lbNuevo = False
             lcSql = "select * from sip where iMadre='" & lcImadre & "'"
             If oRsFox1.State = 1 Then oRsFox1.Close
             oRsFox1.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
             If oRsFox1.RecordCount = 0 Then
                lbNuevo = True
             End If
            
             If lbNuevo = True Then
                oRsFox1.AddNew
                oRsFox1.Fields!iMadre = Left(lcImadre, 30)
                oRsFox1.Fields!HC = Left(lcHC, 20)
                oRsFox1.Fields!ESTABLECIM = Left(lcEstablecim, 100)
                oRsFox1.Fields!Edad = Val(lcEdad)
                oRsFox1.Fields!Distrito = Left(lcDistrito, 10)
                oRsFox1.Fields!ESTUDIOS = Left(lcEstudios, 60)
                oRsFox1.Fields!Talla = Val(lcTalla)
              '  oRsFox1.Fields!Fecha = CDate(lcFecha)
                oRsFox1.Fields!talla_rn = Val(lcTalla_rn)
                oRsFox1.Fields!peso_rn = CDbl(lcPeso_rn)
                If Year(ldFecEmbarazo) > 1980 Then
                   oRsFox1.Fields!Fembarazo = ldFecEmbarazo
                End If
             Else
                If ldFecEmbarazo < oRsFox1.Fields!Fembarazo Then
                   If Year(ldFecEmbarazo) > 1980 Then
                      oRsFox1.Fields!Fembarazo = ldFecEmbarazo
                   End If
                End If
             End If
             If Peso1 > 0 Then
                oRsFox1.Fields!Peso1 = CDbl(Peso1)
             End If
             If Peso1 > 0 Then
                oRsFox1.Fields!Peso2 = CDbl(Peso2)
             End If
             If Peso3 > 0 Then
                oRsFox1.Fields!Peso3 = CDbl(Peso3)
             End If
             If Peso4 > 0 Then
                oRsFox1.Fields!Peso4 = CDbl(Peso4)
             End If
             If Peso5 > 0 Then
                oRsFox1.Fields!Peso5 = CDbl(Peso5)
             End If
             If Peso6 > 0 Then
                oRsFox1.Fields!Peso6 = CDbl(Peso6)
             End If
             If Peso7 > 0 Then
                oRsFox1.Fields!Peso7 = CDbl(Peso7)
             End If
             If Peso8 > 0 Then
                oRsFox1.Fields!Peso8 = CDbl(Peso8)
             End If
             If Peso9 > 0 Then
                oRsFox1.Fields!Peso9 = CDbl(Peso9)
             End If
             If Peso10 > 0 Then
                oRsFox1.Fields!Peso10 = CDbl(Peso10)
             End If
             If Peso11 > 0 Then
                oRsFox1.Fields!Peso11 = CDbl(Peso11)
             End If
             If Peso12 > 0 Then
                oRsFox1.Fields!Peso12 = CDbl(Peso12)
             End If
             If Peso13 > 0 Then
                oRsFox1.Fields!Peso13 = CDbl(Peso13)
             End If
             If Peso14 > 0 Then
                oRsFox1.Fields!Peso14 = CDbl(Peso14)
             End If
             If Peso15 > 0 Then
                oRsFox1.Fields!Peso15 = CDbl(Peso15)
             End If
             If Peso16 > 0 Then
                oRsFox1.Fields!Peso16 = CDbl(Peso16)
             End If
             If Peso17 > 0 Then
                oRsFox1.Fields!Peso17 = CDbl(Peso17)
             End If
             If Peso18 > 0 Then
                oRsFox1.Fields!Peso18 = CDbl(Peso18)
             End If
             If Peso19 > 0 Then
                oRsFox1.Fields!Peso19 = CDbl(Peso19)
             End If
             If Peso20 > 0 Then
                oRsFox1.Fields!Peso20 = CDbl(Peso20)
             End If
             If Peso21 > 0 Then
                oRsFox1.Fields!Peso21 = CDbl(Peso21)
             End If
             If Peso22 > 0 Then
                oRsFox1.Fields!Peso22 = CDbl(Peso22)
             End If
             If Peso23 > 0 Then
                oRsFox1.Fields!Peso23 = CDbl(Peso23)
             End If
             If Peso24 > 0 Then
                oRsFox1.Fields!Peso24 = CDbl(Peso24)
             End If
             If Peso25 > 0 Then
                oRsFox1.Fields!Peso25 = CDbl(Peso25)
             End If
             If Peso26 > 0 Then
                oRsFox1.Fields!Peso26 = CDbl(Peso26)
             End If
             If Peso27 > 0 Then
                oRsFox1.Fields!Peso27 = CDbl(Peso27)
             End If
             If Peso28 > 0 Then
                oRsFox1.Fields!Peso28 = CDbl(Peso28)
             End If
             If Peso29 > 0 Then
                oRsFox1.Fields!Peso29 = CDbl(Peso29)
             End If
             If Peso30 > 0 Then
                oRsFox1.Fields!Peso30 = CDbl(Peso30)
             End If
             If Peso31 > 0 Then
                oRsFox1.Fields!Peso31 = CDbl(Peso31)
             End If
             If Peso32 > 0 Then
                oRsFox1.Fields!Peso32 = CDbl(Peso32)
             End If
             If Peso33 > 0 Then
                oRsFox1.Fields!Peso33 = CDbl(Peso33)
             End If
             If Peso34 > 0 Then
                oRsFox1.Fields!Peso34 = CDbl(Peso34)
             End If
             If Peso35 > 0 Then
                oRsFox1.Fields!Peso35 = CDbl(Peso35)
             End If
             If Peso36 > 0 Then
                oRsFox1.Fields!Peso36 = CDbl(Peso36)
             End If
             If Peso37 > 0 Then
                oRsFox1.Fields!Peso37 = CDbl(Peso37)
             End If
             If Peso38 > 0 Then
                oRsFox1.Fields!Peso38 = CDbl(Peso38)
             End If
             If Peso39 > 0 Then
                oRsFox1.Fields!Peso39 = CDbl(Peso39)
             End If
             If Peso40 > 0 Then
                oRsFox1.Fields!Peso40 = CDbl(Peso40)
             End If
             If Peso41 > 0 Then
                oRsFox1.Fields!Peso41 = CDbl(Peso41)
             End If
             If Peso42 > 0 Then
                oRsFox1.Fields!Peso42 = CDbl(Peso42)
             End If
             If Peso43 > 0 Then
                oRsFox1.Fields!Peso43 = CDbl(Peso43)
             End If
             If Peso44 > 0 Then
                oRsFox1.Fields!Peso44 = CDbl(Peso44)
             End If
             If Peso45 > 0 Then
                 oRsFox1.Fields!Peso45 = CDbl(Peso45)
             End If
             oRsFox1.Update
             oRsMDB.MoveNext
        Next
   Unload Me
   Exit Sub
ErSip2000:
   If Err.Number = 13 And lcSql = "percentil" Then
      Resume Next
   Else
      MsgBox Err.Description
      Resume
   End If

End Sub

Private Sub cmdProcesaSip_Click()
    On Error GoTo ErSip2000
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Dim s As Excel.Worksheet
    Set W = EXL.Workbooks.Open(App.Path & "\archivos\percentiles.xls")       'usa
    Set s = W.Sheets("IMC")
    '
    Dim oConexionMDB As New Connection, oRsMDB As New Recordset
    oConexionMDB.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\tablasYpa.mdb;"
    '
    Dim lcImadre As String, lcHC As String, lcEstablecim As String, lcEdad As String, lcDistrito As String
    Dim lcEstudios As String, lcTalla As String, lcFecha As String, lcTalla_rn As String, lcPeso_rn As String
    Dim Peso1 As String, Peso2 As String, Peso3 As String, Peso4 As String, Peso5 As String
    Dim Peso6 As String, Peso7 As String, Peso8 As String, Peso9 As String, Peso10 As String
    Dim Peso11 As String, Peso12 As String, Peso13 As String, Peso14 As String, Peso15 As String
    Dim Peso16 As String, Peso17 As String, Peso18 As String, Peso19 As String, Peso20 As String
    Dim Peso21 As String, Peso22 As String, Peso23 As String, Peso24 As String, Peso25 As String
    Dim Peso26 As String, Peso27 As String, Peso28 As String, Peso29 As String, Peso30 As String
    Dim Peso31 As String, Peso32 As String, Peso33 As String, Peso34 As String, Peso35 As String
    Dim Peso36 As String, Peso37 As String, Peso38 As String, Peso39 As String, Peso40 As String
    Dim Peso41 As String, Peso42 As String, Peso43 As String, Peso44 As String, Peso45 As String
    Dim lbNuevo As Boolean, lnFor1 As Integer
    Dim lnFor As Long, lnFila As Long, lcRango As String, lnFilaFinal As Long
    Dim oRsFox1 As New Recordset, lnPercentilIMC As Double, lcPercentilIMC As String
    Dim lcEdadG As String, ldFecha As Date, ldFecEmbarazo As Date
    
    lcSql = "delete from sip"
    If oRsFox1.State = 1 Then oRsFox1.Close
    oRsFox1.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    '
    For lnFor1 = 1 To 2
         If lnFor1 = 1 Then
            lcSql = "select * from parte01"
         Else
            lcSql = "select * from parte02"
         End If
         If oRsMDB.State = 1 Then oRsMDB.Close
         oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
         lnFila = 1
         lnFilaFinal = oRsMDB.RecordCount
         oRsMDB.MoveFirst
         ProgressBar1.Min = 0: ProgressBar1.Max = lnFilaFinal
         For lnFor = lnFila To lnFilaFinal
             ProgressBar1.Value = lnFor
             lcImadre = Left(oRsMDB.Fields!IDMADRE & Space(30), 30)
If Trim(lcImadre) = "000X2AWB4763" Then
lcSql = ""
End If
             If Trim(lcImadre) = "" Then
                Exit For
             End If
             lcHC = oRsMDB.Fields!HC
             lcEstablecim = oRsMDB.Fields!Establecimiento
             lcEdad = oRsMDB.Fields!Edad
             lcDistrito = IIf(IsNull(oRsMDB.Fields!Distrito), "", oRsMDB.Fields!Distrito)
             lcEstudios = oRsMDB.Fields!ESTUDIOS
             lcTalla = oRsMDB.Fields!Talla
             lcTalla_rn = oRsMDB.Fields!talla_ra
             lcPeso_rn = oRsMDB.Fields!peso_rn
             lcEdadG = oRsMDB.Fields!edad_gestacional
             ldFecha = oRsMDB.Fields!Fecha
             Peso1 = 0: Peso2 = 0: Peso3 = 0: Peso4 = 0: Peso5 = 0
             Peso6 = 0: Peso7 = 0: Peso8 = 0: Peso9 = 0: Peso10 = 0
             Peso11 = 0: Peso12 = 0: Peso13 = 0: Peso14 = 0: Peso15 = 0
             Peso16 = 0: Peso17 = 0: Peso18 = 0: Peso19 = 0: Peso20 = 0
             Peso21 = 0: Peso22 = 0: Peso23 = 0: Peso24 = 0: Peso25 = 0
             Peso26 = 0: Peso27 = 0: Peso28 = 0: Peso29 = 0: Peso30 = 0
             Peso31 = 0: Peso32 = 0: Peso33 = 0: Peso34 = 0: Peso35 = 0
             Peso36 = 0: Peso37 = 0: Peso38 = 0: Peso39 = 0: Peso40 = 0
             Peso41 = 0: Peso42 = 0: Peso43 = 0: Peso44 = 0: Peso45 = 0
             '
             lcSql = ".."
             lnPercentilIMC = 0
             lcPercentilIMC = "ERR"
             If oRsMDB.Fields!Peso > 0 And oRsMDB.Fields!Talla > 0 Then
                s.Cells(203, 6).Value = oRsMDB.Fields!Peso
                s.Cells(205, 6).Value = Round(oRsMDB.Fields!Talla / 100, 2)
                s.Cells(209, 6).Value = oRsMDB.Fields!edad_gestacional
                lcSql = "percentil"
                lcPercentilIMC = s.Cells(211, 6).Value
                lcSql = ".."
                lnPercentilIMC = IIf(UCase(Left(lcPercentilIMC, 3)) = "ERR", 0, Val(lcPercentilIMC))
             End If
             '
             ldFecEmbarazo = ldFecha - (Val(lcEdadG) * 7)
             Select Case lcEdadG
             Case "1"
                     Peso1 = lnPercentilIMC         'oRsMDB.Fields!Peso
                     
             Case "2"
                     Peso2 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "3"
                     Peso3 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "4"
                     Peso4 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "5"
                     Peso5 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "6"
                     Peso6 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "7"
                     Peso7 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "8"
                     Peso8 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "9"
                     Peso9 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "10"
                     Peso10 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "11"
                     Peso11 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "12"
                     Peso12 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "13"
                     Peso13 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "14"
                     Peso14 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "15"
                     Peso15 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "16"
                     Peso16 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "17"
                     Peso17 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "18"
                     Peso18 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "19"
                     Peso19 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "20"
                     Peso20 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "21"
                     Peso21 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "22"
                     Peso22 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "23"
                     Peso23 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "24"
                     Peso24 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "25"
                     Peso25 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "26"
                     Peso26 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "27"
                     Peso27 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "28"
                     Peso28 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "29"
                     Peso29 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "30"
                     Peso30 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "31"
                     Peso31 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "32"
                     Peso32 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "33"
                     Peso33 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "34"
                     Peso34 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "35"
                     Peso35 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "36"
                     Peso36 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "37"
                     Peso37 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "38"
                     Peso38 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "39"
                     Peso39 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "40"
                     Peso40 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "41"
                     Peso41 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "42"
                     Peso42 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "43"
                     Peso43 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "44"
                     Peso44 = lnPercentilIMC         'oRsMDB.Fields!Peso
             Case "45"
                     Peso45 = lnPercentilIMC         'oRsMDB.Fields!Peso
             End Select
If Trim(lcImadre) = "1513KULW7812" Then
lcSql = ""
End If
             lbNuevo = False
             lcSql = "select * from sip where iMadre='" & lcImadre & "'"
             If oRsFox1.State = 1 Then oRsFox1.Close
             oRsFox1.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
             If oRsFox1.RecordCount = 0 Then
                lbNuevo = True
             End If
            
             If lbNuevo = True Then
                oRsFox1.AddNew
                oRsFox1.Fields!iMadre = lcImadre
                oRsFox1.Fields!HC = Left(lcHC, 20)
                oRsFox1.Fields!ESTABLECIM = Left(lcEstablecim, 100)
                oRsFox1.Fields!Edad = Val(lcEdad)
                oRsFox1.Fields!Distrito = Left(lcDistrito, 10)
                oRsFox1.Fields!ESTUDIOS = Left(lcEstudios, 60)
                oRsFox1.Fields!Talla = Val(lcTalla)
              '  oRsFox1.Fields!Fecha = CDate(lcFecha)
                oRsFox1.Fields!talla_rn = Val(lcTalla_rn)
                oRsFox1.Fields!peso_rn = CDbl(lcPeso_rn)
                oRsFox1.Fields!Fembarazo = ldFecEmbarazo
             Else
                If ldFecEmbarazo < oRsFox1.Fields!Fembarazo Then
                   oRsFox1.Fields!Fembarazo = ldFecEmbarazo
                End If
             End If
             If Peso1 > 0 Then
                oRsFox1.Fields!Peso1 = CDbl(Peso1)
             End If
             If Peso1 > 0 Then
                oRsFox1.Fields!Peso2 = CDbl(Peso2)
             End If
             If Peso3 > 0 Then
                oRsFox1.Fields!Peso3 = CDbl(Peso3)
             End If
             If Peso4 > 0 Then
                oRsFox1.Fields!Peso4 = CDbl(Peso4)
             End If
             If Peso5 > 0 Then
                oRsFox1.Fields!Peso5 = CDbl(Peso5)
             End If
             If Peso6 > 0 Then
                oRsFox1.Fields!Peso6 = CDbl(Peso6)
             End If
             If Peso7 > 0 Then
                oRsFox1.Fields!Peso7 = CDbl(Peso7)
             End If
             If Peso8 > 0 Then
                oRsFox1.Fields!Peso8 = CDbl(Peso8)
             End If
             If Peso9 > 0 Then
                oRsFox1.Fields!Peso9 = CDbl(Peso9)
             End If
             If Peso10 > 0 Then
                oRsFox1.Fields!Peso10 = CDbl(Peso10)
             End If
             If Peso11 > 0 Then
                oRsFox1.Fields!Peso11 = CDbl(Peso11)
             End If
             If Peso12 > 0 Then
                oRsFox1.Fields!Peso12 = CDbl(Peso12)
             End If
             If Peso13 > 0 Then
                oRsFox1.Fields!Peso13 = CDbl(Peso13)
             End If
             If Peso14 > 0 Then
                oRsFox1.Fields!Peso14 = CDbl(Peso14)
             End If
             If Peso15 > 0 Then
                oRsFox1.Fields!Peso15 = CDbl(Peso15)
             End If
             If Peso16 > 0 Then
                oRsFox1.Fields!Peso16 = CDbl(Peso16)
             End If
             If Peso17 > 0 Then
                oRsFox1.Fields!Peso17 = CDbl(Peso17)
             End If
             If Peso18 > 0 Then
                oRsFox1.Fields!Peso18 = CDbl(Peso18)
             End If
             If Peso19 > 0 Then
                oRsFox1.Fields!Peso19 = CDbl(Peso19)
             End If
             If Peso20 > 0 Then
                oRsFox1.Fields!Peso20 = CDbl(Peso20)
             End If
             If Peso21 > 0 Then
                oRsFox1.Fields!Peso21 = CDbl(Peso21)
             End If
             If Peso22 > 0 Then
                oRsFox1.Fields!Peso22 = CDbl(Peso22)
             End If
             If Peso23 > 0 Then
                oRsFox1.Fields!Peso23 = CDbl(Peso23)
             End If
             If Peso24 > 0 Then
                oRsFox1.Fields!Peso24 = CDbl(Peso24)
             End If
             If Peso25 > 0 Then
                oRsFox1.Fields!Peso25 = CDbl(Peso25)
             End If
             If Peso26 > 0 Then
                oRsFox1.Fields!Peso26 = CDbl(Peso26)
             End If
             If Peso27 > 0 Then
                oRsFox1.Fields!Peso27 = CDbl(Peso27)
             End If
             If Peso28 > 0 Then
                oRsFox1.Fields!Peso28 = CDbl(Peso28)
             End If
             If Peso29 > 0 Then
                oRsFox1.Fields!Peso29 = CDbl(Peso29)
             End If
             If Peso30 > 0 Then
                oRsFox1.Fields!Peso30 = CDbl(Peso30)
             End If
             If Peso31 > 0 Then
                oRsFox1.Fields!Peso31 = CDbl(Peso31)
             End If
             If Peso32 > 0 Then
                oRsFox1.Fields!Peso32 = CDbl(Peso32)
             End If
             If Peso33 > 0 Then
                oRsFox1.Fields!Peso33 = CDbl(Peso33)
             End If
             If Peso34 > 0 Then
                oRsFox1.Fields!Peso34 = CDbl(Peso34)
             End If
             If Peso35 > 0 Then
                oRsFox1.Fields!Peso35 = CDbl(Peso35)
             End If
             If Peso36 > 0 Then
                oRsFox1.Fields!Peso36 = CDbl(Peso36)
             End If
             If Peso37 > 0 Then
                oRsFox1.Fields!Peso37 = CDbl(Peso37)
             End If
             If Peso38 > 0 Then
                oRsFox1.Fields!Peso38 = CDbl(Peso38)
             End If
             If Peso39 > 0 Then
                oRsFox1.Fields!Peso39 = CDbl(Peso39)
             End If
             If Peso40 > 0 Then
                oRsFox1.Fields!Peso40 = CDbl(Peso40)
             End If
             If Peso41 > 0 Then
                oRsFox1.Fields!Peso41 = CDbl(Peso41)
             End If
             If Peso42 > 0 Then
                oRsFox1.Fields!Peso42 = CDbl(Peso42)
             End If
             If Peso43 > 0 Then
                oRsFox1.Fields!Peso43 = CDbl(Peso43)
             End If
             If Peso44 > 0 Then
                oRsFox1.Fields!Peso44 = CDbl(Peso44)
             End If
             If Peso45 > 0 Then
                 oRsFox1.Fields!Peso45 = CDbl(Peso45)
             End If
             oRsFox1.Update
             oRsMDB.MoveNext
        Next
   Next
   Unload Me
   Exit Sub
ErSip2000:
   If Err.Number = 13 And lcSql = "percentil" Then
      Resume Next
   Else
      MsgBox Err.Description
      Resume
   End If
   


End Sub



Private Sub cmdProcesaSip2000_Click()
    On Error GoTo ErrSip2000
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Set W = EXL.Workbooks.Open("d:\barrantes\sisMt.xls")
    Dim s As Excel.Worksheet
    Dim lcImadre As String, lcHC As String, lcEstablecim As String, lcEdad As String, lcDistrito As String
    Dim lcEstudios As String, lcTalla As String, lcFecha As String, lcTalla_rn As String, lcPeso_rn As String
    Dim Peso1 As String, Peso2 As String, Peso3 As String, Peso4 As String, Peso5 As String
    Dim Peso6 As String, Peso7 As String, Peso8 As String, Peso9 As String, Peso10 As String
    Dim Peso11 As String, Peso12 As String, Peso13 As String, Peso14 As String, Peso15 As String
    Dim Peso16 As String, Peso17 As String, Peso18 As String, Peso19 As String, Peso20 As String
    Dim Peso21 As String, Peso22 As String, Peso23 As String, Peso24 As String, Peso25 As String
    Dim Peso26 As String, Peso27 As String, Peso28 As String, Peso29 As String, Peso30 As String
    Dim Peso31 As String, Peso32 As String, Peso33 As String, Peso34 As String, Peso35 As String
    Dim Peso36 As String, Peso37 As String, Peso38 As String, Peso39 As String, Peso40 As String
    Dim Peso41 As String, Peso42 As String, Peso43 As String, Peso44 As String, Peso45 As String
    Dim lbNuevo As Boolean, lnFor1 As Integer
    Dim lnFor As Long, lnFila As Long, lcRango As String, lnFilaFinal As Long
    Dim oConexionFox As New ADODB.Connection, oRsFox1 As New Recordset, lcEdadG As String
    
    
    oConexionFox.CommandTimeout = 300
    oConexionFox.Open "DSN=his"
    '
    For lnFor1 = 1 To 2
         If lnFor1 = 1 Then
            Set s = W.Sheets("Parte 01")
         Else
            Set s = W.Sheets("Parte 02")
         End If
         lnFila = 2
         lnFilaFinal = 65500
         ProgressBar1.Min = 0: ProgressBar1.Max = lnFilaFinal
         For lnFor = lnFila To lnFilaFinal
             ProgressBar1.Value = lnFor
             lcRango = "C" + Trim(Str(lnFor))
             lcImadre = Left(Trim(s.Range(lcRango).Value) & Space(100), 30)
             If Trim(lcImadre) = "" Then
                Exit For
             End If
             lcRango = "D" + Trim(Str(lnFor))
             lcHC = Trim(s.Range(lcRango).Value)
             lcRango = "E" + Trim(Str(lnFor))
             lcEstablecim = Trim(s.Range(lcRango).Value)
             lcRango = "F" + Trim(Str(lnFor))
             lcEdad = Trim(s.Range(lcRango).Value)
             lcRango = "G" + Trim(Str(lnFor))
             lcDistrito = Trim(s.Range(lcRango).Value)
             lcRango = "H" + Trim(Str(lnFor))
             lcEstudios = Trim(s.Range(lcRango).Value)
             lcRango = "J" + Trim(Str(lnFor))
             lcTalla = Trim(s.Range(lcRango).Value)
             'lcRango = "L" + Trim(Str(lnFor))
             'lcFecha = Trim(s.Range(lcRango).Value)
             lcRango = "R" + Trim(Str(lnFor))
             lcTalla_rn = Trim(s.Range(lcRango).Value)
             lcRango = "S" + Trim(Str(lnFor))
             lcPeso_rn = Trim(s.Range(lcRango).Value)
             lcRango = "M" + Trim(Str(lnFor))
             lcEdadG = Trim(s.Range(lcRango).Value)
             Peso1 = 0: Peso2 = 0: Peso3 = 0: Peso4 = 0: Peso5 = 0
             Peso6 = 0: Peso7 = 0: Peso8 = 0: Peso9 = 0: Peso10 = 0
             Peso11 = 0: Peso12 = 0: Peso13 = 0: Peso14 = 0: Peso15 = 0
             Peso16 = 0: Peso17 = 0: Peso18 = 0: Peso19 = 0: Peso20 = 0
             Peso21 = 0: Peso22 = 0: Peso23 = 0: Peso24 = 0: Peso25 = 0
             Peso26 = 0: Peso27 = 0: Peso28 = 0: Peso29 = 0: Peso30 = 0
             Peso31 = 0: Peso32 = 0: Peso33 = 0: Peso34 = 0: Peso35 = 0
             Peso36 = 0: Peso37 = 0: Peso38 = 0: Peso39 = 0: Peso40 = 0
             Peso41 = 0: Peso42 = 0: Peso43 = 0: Peso44 = 0: Peso45 = 0
             Select Case lcEdadG
             Case "1"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso1 = Trim(s.Range(lcRango).Value)
             Case "2"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso2 = Trim(s.Range(lcRango).Value)
             Case "3"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso3 = Trim(s.Range(lcRango).Value)
             Case "4"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso4 = Trim(s.Range(lcRango).Value)
             Case "5"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso5 = Trim(s.Range(lcRango).Value)
             Case "6"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso6 = Trim(s.Range(lcRango).Value)
             Case "7"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso7 = Trim(s.Range(lcRango).Value)
             Case "8"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso8 = Trim(s.Range(lcRango).Value)
             Case "9"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso9 = Trim(s.Range(lcRango).Value)
             Case "10"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso10 = Trim(s.Range(lcRango).Value)
             Case "11"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso11 = Trim(s.Range(lcRango).Value)
             Case "12"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso12 = Trim(s.Range(lcRango).Value)
             Case "13"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso13 = Trim(s.Range(lcRango).Value)
             Case "14"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso14 = Trim(s.Range(lcRango).Value)
             Case "15"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso15 = Trim(s.Range(lcRango).Value)
             Case "16"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso16 = Trim(s.Range(lcRango).Value)
             Case "17"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso17 = Trim(s.Range(lcRango).Value)
             Case "18"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso18 = Trim(s.Range(lcRango).Value)
             Case "19"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso19 = Trim(s.Range(lcRango).Value)
             Case "20"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso20 = Trim(s.Range(lcRango).Value)
             Case "21"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso21 = Trim(s.Range(lcRango).Value)
             Case "22"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso22 = Trim(s.Range(lcRango).Value)
             Case "23"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso23 = Trim(s.Range(lcRango).Value)
             Case "24"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso24 = Trim(s.Range(lcRango).Value)
             Case "25"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso25 = Trim(s.Range(lcRango).Value)
             Case "26"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso26 = Trim(s.Range(lcRango).Value)
             Case "27"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso27 = Trim(s.Range(lcRango).Value)
             Case "28"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso28 = Trim(s.Range(lcRango).Value)
             Case "29"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso29 = Trim(s.Range(lcRango).Value)
             Case "30"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso30 = Trim(s.Range(lcRango).Value)
             Case "31"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso31 = Trim(s.Range(lcRango).Value)
             Case "32"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso32 = Trim(s.Range(lcRango).Value)
             Case "33"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso33 = Trim(s.Range(lcRango).Value)
             Case "34"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso34 = Trim(s.Range(lcRango).Value)
             Case "35"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso35 = Trim(s.Range(lcRango).Value)
             Case "36"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso36 = Trim(s.Range(lcRango).Value)
             Case "37"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso37 = Trim(s.Range(lcRango).Value)
             Case "38"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso38 = Trim(s.Range(lcRango).Value)
             Case "39"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso39 = Trim(s.Range(lcRango).Value)
             Case "40"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso40 = Trim(s.Range(lcRango).Value)
             Case "41"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso41 = Trim(s.Range(lcRango).Value)
             Case "42"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso42 = Trim(s.Range(lcRango).Value)
             Case "43"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso43 = Trim(s.Range(lcRango).Value)
             Case "44"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso44 = Trim(s.Range(lcRango).Value)
             Case "45"
                     lcRango = "N" + Trim(Str(lnFor))
                     Peso45 = Trim(s.Range(lcRango).Value)
             End Select
If Trim(lcImadre) = "M00KJO1E9935" Then
lcSql = ""
End If
             lbNuevo = False
             lcSql = "select * from sip2000 where iMadre='" & lcImadre & "'"
             If oRsFox1.State = 1 Then oRsFox1.Close
             oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
             If oRsFox1.RecordCount = 0 Then
                lbNuevo = True
             End If
            
             If lbNuevo = True Then
                oRsFox1.AddNew
                oRsFox1.Fields!iMadre = lcImadre
                oRsFox1.Fields!HC = lcHC
                oRsFox1.Fields!ESTABLECIM = lcEstablecim
                oRsFox1.Fields!Edad = Val(lcEdad)
                oRsFox1.Fields!Distrito = lcDistrito
                oRsFox1.Fields!ESTUDIOS = lcEstudios
                oRsFox1.Fields!Talla = Val(lcTalla)
              '  oRsFox1.Fields!Fecha = CDate(lcFecha)
                oRsFox1.Fields!talla_rn = Val(lcTalla_rn)
                oRsFox1.Fields!peso_rn = CDbl(lcPeso_rn)
             End If
             oRsFox1.Fields!Peso1 = CDbl(Peso1)
             oRsFox1.Fields!Peso2 = CDbl(Peso2)
             oRsFox1.Fields!Peso3 = CDbl(Peso3)
             oRsFox1.Fields!Peso4 = CDbl(Peso4)
             oRsFox1.Fields!Peso5 = CDbl(Peso5)
             oRsFox1.Fields!Peso6 = CDbl(Peso6)
             oRsFox1.Fields!Peso7 = CDbl(Peso7)
             oRsFox1.Fields!Peso8 = CDbl(Peso8)
             oRsFox1.Fields!Peso9 = CDbl(Peso9)
             oRsFox1.Fields!Peso10 = CDbl(Peso10)
             oRsFox1.Fields!Peso11 = CDbl(Peso11)
             oRsFox1.Fields!Peso12 = CDbl(Peso12)
             oRsFox1.Fields!Peso13 = CDbl(Peso13)
             oRsFox1.Fields!Peso14 = CDbl(Peso14)
             oRsFox1.Fields!Peso15 = CDbl(Peso15)
             oRsFox1.Fields!Peso16 = CDbl(Peso16)
             oRsFox1.Fields!Peso17 = CDbl(Peso17)
             oRsFox1.Fields!Peso18 = CDbl(Peso18)
             oRsFox1.Fields!Peso19 = CDbl(Peso19)
             oRsFox1.Fields!Peso20 = CDbl(Peso20)
             oRsFox1.Fields!Peso21 = CDbl(Peso21)
             oRsFox1.Fields!Peso22 = CDbl(Peso22)
             oRsFox1.Fields!Peso23 = CDbl(Peso23)
             oRsFox1.Fields!Peso24 = CDbl(Peso24)
             oRsFox1.Fields!Peso25 = CDbl(Peso25)
             oRsFox1.Fields!Peso26 = CDbl(Peso26)
             oRsFox1.Fields!Peso27 = CDbl(Peso27)
             oRsFox1.Fields!Peso28 = CDbl(Peso28)
             oRsFox1.Fields!Peso29 = CDbl(Peso29)
             oRsFox1.Fields!Peso30 = CDbl(Peso30)
             oRsFox1.Fields!Peso31 = CDbl(Peso31)
             oRsFox1.Fields!Peso32 = CDbl(Peso32)
             oRsFox1.Fields!Peso33 = CDbl(Peso33)
             oRsFox1.Fields!Peso34 = CDbl(Peso34)
             oRsFox1.Fields!Peso35 = CDbl(Peso35)
             oRsFox1.Fields!Peso36 = CDbl(Peso36)
             oRsFox1.Fields!Peso37 = CDbl(Peso37)
             oRsFox1.Fields!Peso38 = CDbl(Peso38)
             oRsFox1.Fields!Peso39 = CDbl(Peso39)
             oRsFox1.Fields!Peso40 = CDbl(Peso40)
             oRsFox1.Fields!Peso41 = CDbl(Peso41)
             oRsFox1.Fields!Peso42 = CDbl(Peso42)
             oRsFox1.Fields!Peso43 = CDbl(Peso43)
             oRsFox1.Fields!Peso44 = CDbl(Peso44)
             oRsFox1.Fields!Peso45 = CDbl(Peso45)
             oRsFox1.Update
        Next
   Next
   Unload Me
   Exit Sub
ErrSip2000:
   MsgBox Err.Description
   Resume
End Sub



Private Sub cmdActualizaPercentil_Click()

    'On Error GoTo ErrRptHuelga
    On Error Resume Next
    Dim ml_EdadEnMeses As Long
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Dim s As Excel.Worksheet
    Dim W1 As Excel.Workbook
    Dim s1 As Excel.Worksheet
    Dim oRsTmp1 As New Recordset
    Dim oFila As Long, ldFecha As Date, lbNuevo As Boolean
    Dim ldFechaInicialHist As Date, ldFechaFinalHist As Date
    Dim lnNroConsultas As Long, lcFecha As String, lcHoraAtencion As String, lcTexto As String
    Dim oConexion As New Connection
    Dim ml_idTipoSexo As Integer, ldFechaNacimiento As Date, ldFechaAtencion As Date
    Dim lnPeso As Double, lnTalla As Double
    Dim lnEdadEnMesesMasPuntoCinco As Double, lnMinimo As Double, lnMaximo As Double, lnIMC As Double
    Dim lnTallaEnCmMasPuntoCinco As Double
    Dim lnPercentilPE As Double, lnPercentilTE As Double, lnPercentilPT As Double
    Const lnPercentilNull As Long = 0
    '
    Set W = EXL.Workbooks.Open(App.Path & "\Plantillas\cred.xls")
    '
    Set W1 = EXL.Workbooks.Open(App.Path & "\plantillas\padron.xls")
    Set s1 = W1.Sheets(txtHoja.Text)
    s1.Cells(1, 30).Value = "Perc.PesoTalla"
    s1.Cells(1, 31).Value = "Perc.TalleEdad"
    s1.Cells(1, 32).Value = "Perc.PesoEdad"
    oFila = 2
    Do While True
               ml_idTipoSexo = Val(s1.Cells(oFila, 9).Value)
               If ml_idTipoSexo = 0 Then
                  Exit Do
               End If
               ldFechaNacimiento = CDate(s1.Cells(oFila, 11).Value)
               ldFechaAtencion = CDate(s1.Cells(oFila, 15).Value)
               lnPeso = Val(s1.Cells(oFila, 16).Value)
               lnTalla = Val(s1.Cells(oFila, 17).Value)
               If lnPeso > 0 And lnTalla > 0 And IsDate(ldFechaNacimiento) And IsDate(ldFechaAtencion) Then
                    lnPercentilPE = lnPercentilNull: lnPercentilTE = lnPercentilNull: lnPercentilPT = lnPercentilNull
                    ml_EdadEnMeses = SIGHEntidades.DevuelveEdadEnMeses(ldFechaNacimiento, ldFechaAtencion)
                    lnEdadEnMesesMasPuntoCinco = ml_EdadEnMeses + 0.5
                    lnTallaEnCmMasPuntoCinco = lnTalla + 0.5
                    'Peso Edad
                    Set s = W.Sheets("P-E")
                    lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
                    lnMaximo = s.Cells(243, IIf(ml_idTipoSexo = 1, 2, 18)).Value
                    lnPercentilPE = lnPercentilNull
                    s.Cells(246, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnEdadEnMesesMasPuntoCinco
                    s.Cells(247, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnPeso
                    lnPercentilPE = s.Cells(254, IIf(ml_idTipoSexo = 1, 3, 19)).Value

               
                    'Talla Edad
                    Set s = W.Sheets("T-E")
                    lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
                    lnMaximo = s.Cells(243, IIf(ml_idTipoSexo = 1, 2, 18)).Value
                    lnPercentilTE = lnPercentilNull
                    s.Cells(246, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnEdadEnMesesMasPuntoCinco
                    s.Cells(247, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnTalla
                    lnPercentilTE = s.Cells(254, IIf(ml_idTipoSexo = 1, 3, 19)).Value
               
                    'Peso Talla
                    Set s = W.Sheets("P-T")
                    lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
                    lnMaximo = s.Cells(61, IIf(ml_idTipoSexo = 1, 2, 18)).Value
                    lnPercentilPT = lnPercentilNull
                    s.Cells(64, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnTallaEnCmMasPuntoCinco
                    s.Cells(65, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnPeso
                    lnPercentilPT = s.Cells(72, IIf(ml_idTipoSexo = 1, 3, 19)).Value
                    
               
                    s1.Cells(oFila, 30).Value = Format(lnPercentilPT, "####,###.####")
                    s1.Cells(oFila, 31).Value = Format(lnPercentilTE, "####,###.####")
                    s1.Cells(oFila, 32).Value = Format(lnPercentilPE, "####,###.####")
                    
               End If
               oFila = oFila + 1
    Loop
    '
    W.Close True
    W1.Close True
    Set s = Nothing
    Set s1 = Nothing
    Set W = Nothing
    Set W1 = Nothing
    Set EXL = Nothing
    MsgBox "procesó sin problemas"
'    Exit Sub
'ErrRptHuelga:
'    MsgBox Err.Description
'    Resume
End Sub

Private Sub cmdAgregaAtencionCE_Click()
    On Error GoTo ErrAgAtCE
    Dim oConexMDB As New Connection
    Dim oConexion As New Connection
    Dim oConexJamo As New Connection
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim rsDiagnosticos As New Recordset
    Dim oRsMDB1 As New Recordset
    Dim oRsMDB2 As New Recordset
    Dim oRsMDB3 As New Recordset
    Dim oRsMDB4 As New Recordset
    Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
    Dim oFactOrdenServicio As New FactOrdenServicio
    Dim oDoFactOrdenServ As New DoFactOrdenServ
    Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
    Dim oDoFacturacionServicioDespacho As New DoFacturacionServicioDespacho
    Dim mo_AdminAdmision As New ReglasAdmision
    Dim oDiagnostico As New DOAtencionDiagnostico
    Dim mo_paciente As New DOPaciente, oRecetaCabecera As New RecetaCabecera
    Dim mo_atenciones As New DOAtencion, mo_Diagnosticos As New Collection
    Dim mo_DoAtencionDatosAdicionales As New DoAtencionDatosAdicionales
    Dim oRsDevuelveRayosX As New Recordset, oRsDevuelveEcografiaO As New Recordset
    Dim oRsDevuelveEcografiaG As New Recordset, oRsDevuelveTomografia As New Recordset
    Dim oRsDevuelveAnatomia As New Recordset, oRsDevuelvePatologia As New Recordset
    Dim oRsDevuelveBancoSangre As New Recordset, oRsDevuelveFarmacia As New Recordset
    Dim oRsDevuelveRecetaAntesDeImprimir As New Recordset
    Dim oRsDevuelveDx As New Recordset
    Dim mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String
    Dim ml_ldFechaIngreso As Date, mo_cita As New DOCita, ml_idCuentaAtencion As Long
    Dim ml_idUsuario As Long, lnRecetaRayosX As Long, lnRecetaEcografiaO As Long
    Dim lnRecetaEcografiaG As Long, lnRecetaTomografia As Long, lnRecetaAnatomiaP As Long
    Dim lnRecetaPatologiaC As Long, lnRecetaBancoS As Long, lnRecetaFarmacia As Long
    Dim ml_FechaReceta As Date, ms_NombrePaciente As String, wxParametro302 As String, ml_lcServicio As String
    Dim lcSql As String, lbEsNuevaCita As Boolean, mb_ExistenDatos As Boolean, lnIdAtencionNueva As Long
    Dim ldFechaHoy As Date, lcHoraHoy As String, lnHistoriaCreada As Long, lnIdCuentaAtencionNueva As Long
    Dim lcHoraFinal As String, lnIdProgramacionMedica As Long, lbEsCitaAdicional As Boolean
    Dim lbPacienteExistePeroEsDiferente As Boolean
    Dim lnIdPacienteDif As Long    'debb-03/03
    Dim oRsTmp3 As New Recordset    'debb-03/03
    Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes 'debb-03/03
    Dim oRsPacientesDiferentes As New Recordset
    Dim lcHoraInicioAtencion As String
    Dim lnIdCuentaPorProcesar As Long
    
    mo_lnIdTablaLISTBARITEMS = 103
    mo_lcNombrePc = SIGHEntidades.RetornaNombrePC
    wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
    
    oConexMDB.CommandTimeout = 300
    lcSql = "Driver=Microsoft Access Driver (*.mdb);DBQ=" & txtMDB.Text & ";"
    If UCase(Right(txtMDB.Text, 3)) <> "MDB" Then
       lcSql = "dsn=atenciones"
    End If
    oConexMDB.Open lcSql
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexJamo.CommandTimeout = 300
    oConexJamo.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    
    Set oFactOrdenServicio.Conexion = oConexion
    Set oFacturacionServicioDespacho.Conexion = oConexion
    
'    lcSql = "select* from atenciones where not(horaEgreso is null) and idTipoServicio=1 and idEstadoAtencion<>0 order by fechaIngreso,horaIngreso"
'    oRsMDB1.Open lcSql, oConexMDB, adOpenKeyset, adLockOptimistic
'    lcSql = "fechaIngreso>='" & txtF1.Text & "' and fechaIngreso<='" & txtF2.Text & "'"
'    oRsMDB1.Filter = lcSql
    If Val(TxtCuenta.Text) > 0 Then
        lcSql = "select* from atenciones where not(horaEgreso is null) and idTipoServicio=1 and idEstadoAtencion<>0 and idCuentaAtencion=" & TxtCuenta.Text & "  order by fechaIngreso,horaIngreso"
        oRsMDB1.Open lcSql, oConexMDB, adOpenKeyset, adLockOptimistic
    Else
         lcSql = "select* from atenciones where not(horaEgreso is null) and idTipoServicio=1 and idEstadoAtencion<>0 order by fechaIngreso,horaIngreso"
        oRsMDB1.Open lcSql, oConexMDB, adOpenKeyset, adLockOptimistic
        lcSql = "fechaIngreso>='" & txtF1.Text & "' and fechaIngreso<='" & txtF2.Text & "'"
        oRsMDB1.Filter = lcSql
    End If
    '
    ProgressBar1.Min = 0
    txtProblemas.Text = ""
    txtNuievasCuentas.Text = ""
    '
    If oRsMDB1.RecordCount = 0 Then
       MsgBox "No hay datos"
    Else
       
       
       With oRsPacientesDiferentes
            .Fields.Append "IdAtencion", adInteger
            .Fields.Append "IdCuentaAtencion", adInteger
            .Fields.Append "IdPacienteNew", adInteger
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open
       End With
    
       
       
       ProgressBar1.Max = oRsMDB1.RecordCount
       ProgressBar1.Max = ProgressBar1.Max + 2
       ProgressBar1.Value = 0
       
       
       
       
       
       
       
       Set oRecetaCabecera.Conexion = oConexMDB
       
       
       Dim lnRegistro11 As Long, lnPorcentaje As Double
       lnRegistro11 = 1000
       If InStr(TxtCuenta.Text, "*") > 0 Then
            lnPorcentaje = Val(Mid(TxtCuenta.Text, 2, 10))
            lnRegistro11 = Round(ProgressBar1.Max * lnPorcentaje, 0)
            oRsMDB1.MoveFirst
            Do While Not oRsMDB1.EOF
               DoEvents: ProgressBar1.Value = ProgressBar1.Value + 1: Me.Refresh
               oRsMDB1.MoveNext
               lnRegistro11 = lnRegistro11 - 1
               If lnRegistro11 < 0 Then
                 Exit Do
               End If
            Loop
       Else
          oRsMDB1.MoveFirst
       End If
       Do While Not oRsMDB1.EOF
          lnIdCuentaPorProcesar = oRsMDB1!idCuentaAtencion
          DoEvents: ProgressBar1.Value = ProgressBar1.Value + 1: Me.Refresh
          
          lnIdPacienteDif = 0   'debb-03/03
          
          lcHoraInicioAtencion = ""
          If Not IsNull(oRsMDB1!HoraInicioAtencion) Then
             lcHoraInicioAtencion = oRsMDB1!HoraInicioAtencion
          End If
          
If oRsMDB1!idCuentaAtencion = 214842 Then
lcSql = ""
End If
          '
          lnIdAtencionNueva = oRsMDB1!idAtencion
          lnIdCuentaAtencionNueva = oRsMDB1!idCuentaAtencion
          lbEsNuevaCita = False
          lcSql = "select * from atenciones where idTipoServicio=1 and idPaciente=" & oRsMDB1!idPaciente & " and idServicioIngreso=" & _
                    oRsMDB1!IdServicioIngreso & " and IdMedicoIngreso=" & oRsMDB1!IdMedicoIngreso & " and  HoraIngreso='" & _
                    oRsMDB1!HoraIngreso & "'"
          If oRsTmp1.State = 1 Then oRsTmp1.Close
          oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
          'oRsTmp1.Filter = "fechaIngreso='" & oRsMDB1!FechaIngreso & "'"
          If Val(TxtCuenta.Text) > 0 Then
             oRsTmp1.Filter = "idCuentaAtencion=" & lnIdCuentaAtencionNueva
          Else
             oRsTmp1.Filter = "fechaIngreso='" & oRsMDB1!FechaIngreso & "'"
          End If
          If oRsTmp1.RecordCount = 0 Then
             '***** es un paciente con CITA ADICIONAL, SE CREO lA CiTA en el EESS (inicio)***
             lbEsNuevaCita = True
             lcSql = "select * from Citas where idAtencion=" & oRsMDB1!idAtencion
             If oRsMDB2.State = 1 Then oRsMDB2.Close
             oRsMDB2.Open lcSql, oConexMDB, adOpenKeyset, adLockOptimistic
             If oRsMDB2.RecordCount > 0 Then
             
                lcSql = "select fechaCreacion from HistoriasClinicas where idPaciente=" & oRsMDB2!idPaciente
                If oRsMDB4.State = 1 Then oRsMDB4.Close
                oRsMDB4.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                lbPacienteExistePeroEsDiferente = False
                If oRsMDB4.RecordCount > 0 Then
                   If CDate(Format(oRsMDB4!FechaCreacion, "dd/mm/yyyy")) >= CDate(Format(oRsMDB2!Fecha, "dd/mm/yyyy")) Then
                      lbPacienteExistePeroEsDiferente = True
                   End If
                End If
                
                lbEsCitaAdicional = IIf(oRsMDB2!EsCitaAdicional = True, True, False)
                lcHoraFinal = oRsMDB2!HoraFin
                lnIdProgramacionMedica = oRsMDB2!IdProgramacion
                lcSql = "select * from programacionMedica where idProgramacion=" & lnIdProgramacionMedica
                If oRsMDB2.State = 1 Then oRsMDB2.Close
                oRsMDB2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                If oRsMDB2.RecordCount > 0 Then
                    If lcHoraFinal > oRsMDB2!HoraFin Then
                       lcSql = "update programacionMedica set horaFin='" & lcHoraFinal & "' where idProgramacion=" & lnIdProgramacionMedica
                       If oRsTmp1.State = 1 Then oRsTmp1.Close
                       oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                    End If
                    ldFechaHoy = oRsMDB1!FechaIngreso
                    lcHoraHoy = oRsMDB1!HoraIngreso
                    If mo_ReglasDeProgMedica.CrearCitaAutomatica(".", oRsMDB1!FechaIngreso, oRsMDB1!HoraIngreso, lcHoraFinal, _
                                      oRsMDB1!IdServicioIngreso, oRsMDB1!IdMedicoIngreso, ".", _
                                      ".", Date, 0, 1, ".", _
                                      ".", oRsMDB1!idFuenteFinanciamiento, ldFechaHoy, lcHoraHoy, _
                                      lnHistoriaCreada, oConexion, SIGHEntidades.USUARIO, 500 + 183, mo_lcNombrePc, _
                                      lbEsCitaAdicional, oRsMDB1!idPaciente, lnIdAtencionNueva, 0, "", "", 0, "") = True Then
                          Set mo_atenciones = mo_AdminAdmision.AtencionesSeleccionarPorId(lnIdAtencionNueva, oConexion)
                          lnIdCuentaAtencionNueva = mo_atenciones.idCuentaAtencion
                          lbEsNuevaCita = False
                          txtNuievasCuentas.Text = txtNuievasCuentas.Text & " <<CUENTA NUEVA: " & Trim(Str(lnIdCuentaAtencionNueva)) & ">> "
                          If lbPacienteExistePeroEsDiferente = True Then
                             txtNuievasCuentas.Text = txtNuievasCuentas.Text & " <<CUENTA NUEVA: " & Trim(Str(lnIdCuentaAtencionNueva)) & " con PACIENTE DIFERENTE>> "
                          End If
                    Else
                         txtProblemas.Text = txtProblemas.Text & "No se pudo registrar CITA para la cuenta: " & Trim(Str(oRsMDB1!idCuentaAtencion)) & Chr(13) & Chr(10)
                    End If
                End If
             End If
             '***** es un paciente con CITA ADICIONAL, SE CREO lA CiTA en el EESS (fin)***
          Else
             '***debb-06/06
             If oRsTmp1.RecordCount > 1 Then
                oRsTmp1.Filter = "idCuentaAtencion=" & lnIdCuentaAtencionNueva
                If oRsTmp1.RecordCount = 0 Then
                   txtProblemas.Text = txtProblemas.Text & "<<la cuenta: " & Trim(Str(lnIdCuentaAtencionNueva)) & " no existe en la tabla SIGH.ATENCIONES (Importe por CUENTA)>>" & Chr(13) & Chr(10)
                   lbEsNuevaCita = True
                End If
             End If
             If lbEsNuevaCita = False Then
                Set mo_atenciones = mo_AdminAdmision.AtencionesSeleccionarPorId(oRsTmp1!idAtencion, oConexion)
                lnIdAtencionNueva = mo_atenciones.idAtencion
                lnIdCuentaAtencionNueva = mo_atenciones.idCuentaAtencion
             End If
             '***debb-06/06
'             If oRsTmp1.RecordCount > 1 Then
'                oRsTmp1.Filter = "idCuentaAtencion=" & lnIdCuentaAtencionNueva
'             End If
'             Set mo_atenciones = mo_AdminAdmision.AtencionesSeleccionarPorId(oRsTmp1!idAtencion, oConexion)
'             lnIdAtencionNueva = mo_atenciones.idAtencion
'             lnIdCuentaAtencionNueva = mo_atenciones.IdCuentaAtencion
          End If
          If lbEsNuevaCita = False Then
                With mo_atenciones
                        .HoraInicioAtencion = lcHoraInicioAtencion
                        .IdDestinoAtencion = oRsMDB1!IdDestinoAtencion
                        .IdMedicoEgreso = 0
                        .HoraEgreso = ""
                        .FechaEgreso = 0
                        .IdTipoGravedad = 0
                        .IdUsuarioAuditoria = SIGHEntidades.USUARIO
                        .IdEstadoAtencion = sghEstadoTabla.sghRegistrado
                        .IdTipoCondicionALEstab = IIf(IsNull(oRsMDB1!IdTipoCondicionALEstab), 1, oRsMDB1!IdTipoCondicionALEstab)
                        .IdTipoCondicionAlServicio = IIf(IsNull(oRsMDB1!IdTipoCondicionAlServicio), 1, oRsMDB1!IdTipoCondicionAlServicio)
                End With
                '
                lcSql = "select * from AtencionesDatosAdicionales where idAtencion=" & oRsMDB1!idAtencion
                If oRsMDB2.State = 1 Then oRsMDB2.Close
                oRsMDB2.Open lcSql, oConexMDB, adOpenKeyset, adLockOptimistic
                Set mo_DoAtencionDatosAdicionales = mo_AdminAdmision.AtencionesDatosAdicionalesSeleccionarPorId(lnIdAtencionNueva, oConexion)
                With mo_DoAtencionDatosAdicionales
                       .IdTipoReferenciaDestino = IIf(IsNull(oRsMDB2!IdTipoReferenciaDestino), 0, oRsMDB2!IdTipoReferenciaDestino)
                       If oRsMDB2!IdTipoReferenciaDestino = 1 Then
                            .IdEstablecimientoDestino = IIf(IsNull(oRsMDB2!IdEstablecimientoDestino), 0, oRsMDB2!IdEstablecimientoDestino)
                            .IdEstablecimientoNoMinsaDestino = 0
                       Else
                            .IdEstablecimientoDestino = 0
                            .IdEstablecimientoNoMinsaDestino = IIf(IsNull(oRsMDB2!IdEstablecimientoNoMinsaDestino), 0, oRsMDB2!IdEstablecimientoNoMinsaDestino)
                       End If
                       .NroReferenciaDestino = IIf(IsNull(oRsMDB2!NroReferenciaDestino), "", oRsMDB2!NroReferenciaDestino)
                       .NumeroDeHijos = IIf(IsNull(oRsMDB2!NumeroDeHijos), 0, oRsMDB2!NumeroDeHijos)
                       .ProximaCita = IIf(IsNull(oRsMDB2!ProximaCita), 0, oRsMDB2!ProximaCita)
                       .referenciaDservicio = IIf(IsNull(oRsMDB2!referenciaDservicio), "", oRsMDB2!referenciaDservicio)
                       .referenciaDfextension = IIf(IsNull(oRsMDB2!referenciaDfextension), 0, oRsMDB2!referenciaDfextension)
                       .referenciaDftramite = IIf(IsNull(oRsMDB2!referenciaDftramite), 0, oRsMDB2!referenciaDftramite)
                End With
                '
                Set mo_Diagnosticos = Nothing
                If rsDiagnosticos.State = 1 Then rsDiagnosticos.Close
                Set rsDiagnosticos = AtencionesDiagnosticosSeleccionarPorAtencion(oRsMDB1!idAtencion, sghAtencionConsultaExterna, oConexMDB)
                If rsDiagnosticos.RecordCount > 0 Then
                      rsDiagnosticos.MoveFirst
                      Do While Not rsDiagnosticos.EOF
                          Set oDiagnostico = New DOAtencionDiagnostico
                          oDiagnostico.IdAtencionDiagnostico = 0
                          oDiagnostico.idAtencion = lnIdAtencionNueva
                          oDiagnostico.IdDiagnostico = rsDiagnosticos!IdDiagnostico
                          oDiagnostico.IdClasificacionDx = rsDiagnosticos!IdClasificacionDx
                          oDiagnostico.IdSubClasificacionDX = rsDiagnosticos!IdSubClasificacionDX
                          oDiagnostico.IdUsuarioAuditoria = mo_atenciones.IdUsuarioAuditoria
                          oDiagnostico.labConfHIS = IIf(IsNull(rsDiagnosticos!labConfHIS), "", rsDiagnosticos!labConfHIS)
                          mo_Diagnosticos.Add oDiagnostico
                          rsDiagnosticos.MoveNext
                      Loop
                End If
                rsDiagnosticos.Close
                '
                If oRsDevuelveDx.State = 1 Then Set oRsDevuelveDx = Nothing
                With oRsDevuelveDx
                    .Fields.Append "IdCuentaAtencion", adInteger
                    .Fields.Append "IdTipoDiagnostico", adInteger, 4, adFldIsNullable
                    .Fields.Append "DescripcionTipoDx", adVarChar, 100, adFldIsNullable
                    .Fields.Append "IdDiagnostico", adInteger
                    .Fields.Append "CodigoCIE2004", adVarChar, 10
                    .Fields.Append "Descripcion", adVarChar, 255
                    .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable
                    .Fields.Append "Fua", adInteger
                    .Fields.Append "FuaCodigoPrestacion", adVarChar, 3, adFldIsNullable
                    .Fields.Append "Consultorio", adVarChar, 100, adFldIsNullable
                    .Fields.Append "idServicio", adInteger
                    .Fields.Append "grupo", adInteger
                    .Fields.Append "subgrupo", adInteger
                    .CursorType = adOpenKeyset
                    .LockType = adLockOptimistic
                    .Open
                End With
                Set rsDiagnosticos = AtencionesDiagnosticosSeleccionarPorAtencion(oRsMDB1!idAtencion, sghAtencionConsultaExterna, oConexMDB)
                If rsDiagnosticos.RecordCount > 0 Then
                      rsDiagnosticos.MoveFirst
                      Do While Not rsDiagnosticos.EOF
                        With oRsDevuelveDx
                            .AddNew
                            .Fields!IdTipoDiagnostico = rsDiagnosticos!IdSubClasificacionDX
                            .Fields!DescripcionTipoDx = "."
                            .Fields!IdDiagnostico = rsDiagnosticos!IdDiagnostico
                            .Fields!CodigoCIE2004 = "."
                            .Fields!Descripcion = "."
                            .Fields!labConfHIS = IIf(IsNull(rsDiagnosticos!labConfHIS), "", rsDiagnosticos!labConfHIS)
                            .Fields!Grupo = IIf(IsNull(rsDiagnosticos!GrupoHIS), 0, rsDiagnosticos!GrupoHIS)
                            .Fields!SubGrupo = IIf(IsNull(rsDiagnosticos!SubGrupoHIS), 0, rsDiagnosticos!SubGrupoHIS)
                            .Fields!IdServicio = oRsMDB1!idCuentaAtencion
                            'If ml_AScorrelativo > 0 Then
                               .Fields!fua = 1
                            'End If
                        End With
                        rsDiagnosticos.MoveNext
                      Loop
                End If
                rsDiagnosticos.Close
                '
                 ms_NombrePaciente = ""
'                lcSql = "select * from Citas where idAtencion=" & lnIdAtencionNueva
'                If oRsTmp1.State = 1 Then oRsTmp1.Close
'                oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                Set oRsTmp1 = mo_AdminAdmision.CitasSeleccionarPorNroCuenta(lnIdCuentaAtencionNueva)
                If oRsTmp1.RecordCount > 0 Then
                      lnIdPacienteDif = oRsTmp1!idPaciente  'debb-03/03
                      mb_ExistenDatos = mo_AdminAdmision.CitasSeleccionarPorId(oRsTmp1!IdCita, mo_cita, mo_paciente, oConexion)
                      With mo_cita
                      End With
                      ms_NombrePaciente = mo_paciente.ApellidoPaterno & " " & mo_paciente.ApellidoMaterno & " " & mo_paciente.PrimerNombre
                End If
                If mo_paciente.ApellidoPaterno = "" Then
                   txtProblemas.Text = txtProblemas.Text & "No hay HISTORIA para la cuenta: " & Trim(Str(oRsMDB1!idCuentaAtencion)) & Chr(13) & Chr(10)
                Else
                    CreaYCargaTemporales oRsDevuelveFarmacia, oRsDevuelveBancoSangre, oRsDevuelvePatologia, oRsDevuelveAnatomia, _
                                   oRsDevuelveTomografia, oRsDevuelveEcografiaG, _
                                   oRsDevuelveEcografiaO, oRsDevuelveRayosX, oRsMDB1!idCuentaAtencion, lnRecetaRayosX, _
                                   lnRecetaEcografiaO, lnRecetaEcografiaG, _
                                   lnRecetaTomografia, lnRecetaAnatomiaP, _
                                   lnRecetaPatologiaC, lnRecetaBancoS, _
                                   lnRecetaFarmacia, oConexMDB, ml_FechaReceta, oRecetaCabecera
                    '
                    lcSql = "select * from atencionesCE where idAtencion=" & oRsMDB1!idAtencion
                    If oRsMDB2.State = 1 Then oRsMDB2.Close
                    oRsMDB2.Open lcSql, oConexMDB, adOpenKeyset, adLockOptimistic
                    If oRsMDB2.RecordCount > 0 Then
                       'Chequea que el PACIENTE sea el MISMO    'debb-03/03
                       If lnIdPacienteDif > 0 Then
                            lcSql = "select idPaciente from Pacientes where nroHistoriaClinica=" & oRsMDB2!NroHistoriaClinica
                            If oRsTmp3.State = 1 Then oRsTmp3.Close
                            oRsTmp3.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                            If oRsTmp3.RecordCount > 0 Then
                               If oRsTmp3!idPaciente <> lnIdPacienteDif Then
                                  oRsPacientesDiferentes.AddNew
                                  oRsPacientesDiferentes!idAtencion = oRsMDB1!idAtencion
                                  oRsPacientesDiferentes!idCuentaAtencion = lnIdCuentaAtencionNueva
                                  oRsPacientesDiferentes!idPacienteNew = oRsTmp3!idPaciente
                                  oRsPacientesDiferentes.Update
                               End If
                            End If
                            oRsTmp3.Close
                       End If
                       '
                       If IsNull(oRsMDB2!CitaServicioJamo) Then
                          txtProblemas.Text = txtProblemas.Text & "Tiene CITA, pero No hay ATENCION DEL MEDICO para la cuenta: " & Trim(Str(oRsMDB1!idCuentaAtencion)) & Chr(13) & Chr(10)
                       Else
                            ml_lcServicio = oRsMDB2!CitaServicioJamo
                            ModificarDatos mo_atenciones, mo_Diagnosticos, _
                                      mo_DoAtencionDatosAdicionales, _
                                      mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                      mo_atenciones.FechaIngreso, mo_cita, lnIdCuentaAtencionNueva, _
                                      mo_atenciones.IdUsuarioAuditoria, lnRecetaRayosX, lnRecetaEcografiaO, _
                                      lnRecetaEcografiaG, lnRecetaTomografia, lnRecetaAnatomiaP, _
                                      lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, _
                                      oRsDevuelveRayosX, oRsDevuelveEcografiaO, _
                                      oRsDevuelveEcografiaG, oRsDevuelveTomografia, _
                                      oRsDevuelveAnatomia, oRsDevuelvePatologia, _
                                      oRsDevuelveBancoSangre, oRsDevuelveFarmacia, _
                                      ml_FechaReceta, ms_NombrePaciente, wxParametro302, _
                                      oRsDevuelveDx, oRsDevuelveRecetaAntesDeImprimir, ml_lcServicio, oRsMDB2, mo_paciente, _
                                      oConexMDB, oConexion
                        End If
                    End If
                    'CPT registrados en la atencion
                    lcSql = "select * from FactOrdenServicio where idPuntoCarga=1 and idCuentaAtencion=" & lnIdCuentaAtencionNueva
                    If oRsTmp2.State = 1 Then oRsTmp2.Close
                    oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                    If oRsTmp2.RecordCount = 0 Then
                        lcSql = "select * from FactOrdenServicio where idPuntoCarga=1 and idCuentaAtencion=" & oRsMDB1!idCuentaAtencion
                        If oRsMDB2.State = 1 Then oRsMDB2.Close
                        oRsMDB2.Open lcSql, oConexMDB, adOpenKeyset, adLockOptimistic
                        If oRsMDB2.RecordCount > 0 Then
                            oRsMDB2.MoveFirst
                            Do While Not oRsMDB2.EOF
                                With oDoFactOrdenServ
                                     .FechaCreacion = oRsMDB2!FechaCreacion
                                     .FechaDespacho = oRsMDB2!FechaDespacho
                                     .FechaHoraRealizaCpt = oRsMDB2!FechaHoraRealizaCpt
                                     .idCuentaAtencion = lnIdCuentaAtencionNueva
                                     .IdEstadoFacturacion = oRsMDB2!IdEstadoFacturacion
                                     .idFuenteFinanciamiento = oRsMDB2!idFuenteFinanciamiento
                                     '.IdOrden = oRsMDB2!IdOrden
                                     .idPaciente = oRsMDB2!idPaciente
                                     .idPuntoCarga = oRsMDB2!idPuntoCarga
                                     .IdServicioPaciente = oRsMDB2!IdServicioPaciente
                                     .IdTipoFinanciamiento = oRsMDB2!IdTipoFinanciamiento
                                     .idUsuario = oRsMDB2!idUsuario
                                     .IdUsuarioAuditoria = mo_atenciones.IdUsuarioAuditoria
                                     .IdUsuarioDespacho = oRsMDB2!IdUsuarioDespacho
                                End With
                                If oFactOrdenServicio.Insertar(oDoFactOrdenServ) = True Then
                                    lcSql = "select * from FacturacionServicioDespacho where idOrden=" & oRsMDB2!IdOrden
                                    If oRsMDB3.State = 1 Then oRsMDB3.Close
                                    oRsMDB3.Open lcSql, oConexMDB, adOpenKeyset, adLockOptimistic
                                    If oRsMDB3.RecordCount > 0 Then
                                       oRsMDB3.MoveFirst
                                       Do While Not oRsMDB3.EOF
                                            With oDoFacturacionServicioDespacho
                                                 .cantidad = oRsMDB3!cantidad
                                                 .GrupoHIS = oRsMDB3!GrupoHIS
                                                 .IdOrden = oDoFactOrdenServ.IdOrden
                                                 .idProducto = oRsMDB3!idProducto
                                                 .IdUsuarioAuditoria = mo_atenciones.IdUsuarioAuditoria
                                                 .labConfHIS = oRsMDB3!labConfHIS
                                                 .precio = oRsMDB3!precio
                                                 .SubGrupoHIS = oRsMDB3!SubGrupoHIS
                                                 .Total = oRsMDB3!Total
                                            End With
                                            If oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDespacho) = True Then
                                               lcSql = oFacturacionServicioDespacho.MensajeError
                                            End If
                                            oRsMDB3.MoveNext
                                       Loop
                                    End If
                                End If
                                oRsMDB2.MoveNext
                            Loop
                        End If
                    End If
                End If
          End If
'          Exit Do
          '
          oRsMDB1.MoveNext
       Loop
       'debb-03/03
       If oRsPacientesDiferentes.RecordCount > 0 Then
          oRsPacientesDiferentes.MoveFirst
          Do While Not oRsPacientesDiferentes.EOF
             mo_ReglasComunes.ActualizaIdPacienteEnTodasLasTablasSegunNroCuenta oRsPacientesDiferentes!idPacienteNew, _
                              oRsPacientesDiferentes!idCuentaAtencion, oRsPacientesDiferentes!idAtencion, _
                              0, SIGHEntidades.USUARIO, "", ""
             oRsPacientesDiferentes.MoveNext
          Loop
       End If
       '
       
    End If

    oRsMDB1.Close
    oConexion.Close
    oConexMDB.Close
    oConexJamo.Close
    
    Set oConexMDB = Nothing
    Set oConexion = Nothing
    Set oConexJamo = Nothing
    Set oRsTmp1 = Nothing
    Set oRsMDB1 = Nothing
    Set mo_atenciones = Nothing
    Set mo_Diagnosticos = Nothing
    Set mo_DoAtencionDatosAdicionales = Nothing
    Set oRsDevuelveRayosX = Nothing
    Set oRsDevuelveEcografiaO = Nothing
    Set oRsDevuelveEcografiaG = Nothing
    Set oRsDevuelveTomografia = Nothing
    Set oRsDevuelveAnatomia = Nothing
    Set oRsDevuelvePatologia = Nothing
    Set oRsDevuelveBancoSangre = Nothing
    Set oRsDevuelveFarmacia = Nothing
    Set oRsDevuelveRecetaAntesDeImprimir = Nothing
    Set oRsDevuelveDx = Nothing
    Set mo_AdminAdmision = Nothing
    Set rsDiagnosticos = Nothing
    Set mo_paciente = Nothing
    Set oRecetaCabecera = Nothing
    Set oDiagnostico = Nothing
    Set oRsMDB3 = Nothing
    Set oFactOrdenServicio = Nothing
    Set oDoFactOrdenServ = Nothing
    Set oFacturacionServicioDespacho = Nothing
    Set oDoFacturacionServicioDespacho = Nothing
    Set oRsTmp3 = Nothing
    Set mo_ReglasComunes = Nothing
    Set oRsPacientesDiferentes = Nothing
    
    MsgBox "Terminó de ProcEsAr TODO"
    Exit Sub
ErrAgAtCE:
    MsgBox "Cuenta de Atenciones.Atenciones: " & Trim(Str(lnIdCuentaPorProcesar)) & Chr(13) & Err.Description
    Exit Sub
    Resume

End Sub

Sub CreaYCargaTemporales(ByRef oRsFarmacia As Recordset, ByRef oRsBanco As Recordset, ByRef oRsPatologia As Recordset, _
                         ByRef oRsAnatomia As Recordset, ByRef oRsTomografia As Recordset, ByRef oRsEcografiaG As Recordset, _
                         ByRef oRsEcografiaO As Recordset, ByRef oRsRayosX As Recordset, lnIdCuentaAtencion As Long, _
                         ByRef lnRecetaRayosX As Long, _
                         ByRef lnRecetaEcografiaO As Long, ByRef lnRecetaEcografiaG As Long, _
                         ByRef lnRecetaTomografia As Long, ByRef lnRecetaAnatomiaP As Long, _
                         ByRef lnRecetaPatologiaC As Long, ByRef lnRecetaBancoS As Long, _
                         ByRef lnRecetaFarmacia As Long, oConexMDB As Connection, ml_FechaReceta As Date, _
                         oRecetaCabecera As RecetaCabecera)
    If oRsRayosX.State = 1 Then
       If oRsRayosX.RecordCount > 0 Then
            oRsRayosX.MoveFirst
            Do While Not oRsRayosX.EOF
               oRsRayosX.Delete
               oRsRayosX.Update
               oRsRayosX.MoveNext
            Loop
       End If
    Else
    With oRsRayosX
          .Fields.Append "Fua", adInteger
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
          .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    End If
    '
    If oRsEcografiaO.State = 1 Then
       If oRsEcografiaO.RecordCount > 0 Then
            oRsEcografiaO.MoveFirst
            Do While Not oRsEcografiaO.EOF
               oRsEcografiaO.Delete
               oRsEcografiaO.Update
               oRsEcografiaO.MoveNext
            Loop
       End If
    Else
    With oRsEcografiaO
          .Fields.Append "Fua", adInteger
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
          .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    End If
    '
    If oRsEcografiaG.State = 1 Then
       If oRsEcografiaG.RecordCount > 0 Then
            oRsEcografiaG.MoveFirst
            Do While Not oRsEcografiaG.EOF
               oRsEcografiaG.Delete
               oRsEcografiaG.Update
               oRsEcografiaG.MoveNext
            Loop
       End If
    Else
    With oRsEcografiaG
          .Fields.Append "Fua", adInteger
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
          .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    End If
    '
    If oRsTomografia.State = 1 Then
       If oRsTomografia.RecordCount > 0 Then
            oRsTomografia.MoveFirst
            Do While Not oRsTomografia.EOF
               oRsTomografia.Delete
               oRsTomografia.Update
               oRsTomografia.MoveNext
            Loop
       End If
    Else
    With oRsTomografia
          .Fields.Append "Fua", adInteger
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
          .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    End If
    '
    If oRsAnatomia.State = 1 Then
       If oRsAnatomia.RecordCount > 0 Then
            oRsAnatomia.MoveFirst
            Do While Not oRsAnatomia.EOF
               oRsAnatomia.Delete
               oRsAnatomia.Update
               oRsAnatomia.MoveNext
            Loop
       End If
    Else
    With oRsAnatomia
          .Fields.Append "Fua", adInteger
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
          .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    End If
    '
    If oRsPatologia.State = 1 Then
       If oRsPatologia.RecordCount > 0 Then
            oRsPatologia.MoveFirst
            Do While Not oRsPatologia.EOF
               oRsPatologia.Delete
               oRsPatologia.Update
               oRsPatologia.MoveNext
            Loop
       End If
    Else
    With oRsPatologia
          .Fields.Append "Fua", adInteger
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
          .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    End If
    '
    If oRsBanco.State = 1 Then
       If oRsBanco.RecordCount > 0 Then
            oRsBanco.MoveFirst
            Do While Not oRsBanco.EOF
               oRsBanco.Delete
               oRsBanco.Update
               oRsBanco.MoveNext
            Loop
       End If
    Else
    With oRsBanco
          .Fields.Append "Fua", adInteger
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
          .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    End If
    '
    If oRsFarmacia.State = 1 Then
       If oRsFarmacia.RecordCount > 0 Then
            oRsFarmacia.MoveFirst
            Do While Not oRsFarmacia.EOF
               oRsFarmacia.Delete
               oRsFarmacia.Update
               oRsFarmacia.MoveNext
            Loop
       End If
    Else
    With oRsFarmacia
          .Fields.Append "Fua", adInteger
          .Fields.Append "Id", adInteger
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
          .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    End If
    '***carga Recetas
        
    
       Dim oRsDetalleReceta As New Recordset, oRsTmp1 As New Recordset, oRsCabeceraRecetas As New Recordset
       lcSql = "SELECT     RecetaCabecera.*, " & _
"                          (SELECT     TOP 1 DocumentoDespacho" & _
"                            From RecetaDetalleItem" & _
"                            WHERE      idReceta = recetacabecera.idreceta) AS DocumentoDespacho" & _
" From RecetaCabecera" & _
" where RecetaCabecera.idCuentaAtencion =" & Trim(Str(lnIdCuentaAtencion))
       oRsCabeceraRecetas.Open lcSql, oConexMDB, adOpenKeyset, adLockOptimistic
       'Set oRsCabeceraRecetas = oRecetaCabecera.SeleccionarPorIdCuentaAtencion(lnIdCuentaAtencion)
       lnRecetaRayosX = 0
       lnRecetaEcografiaO = 0
       lnRecetaEcografiaG = 0
       lnRecetaTomografia = 0
       lnRecetaAnatomiaP = 0
       lnRecetaPatologiaC = 0
       lnRecetaBancoS = 0
       lnRecetaFarmacia = 0
       If oRsCabeceraRecetas.RecordCount = 0 Then
          Exit Sub
       End If
       oRsCabeceraRecetas.MoveFirst
       Do While Not oRsCabeceraRecetas.EOF
          Select Case oRsCabeceraRecetas.Fields!idPuntoCarga
          Case sghPtoCargaRayosX
               lnRecetaRayosX = oRsCabeceraRecetas.Fields!IdReceta
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  lnRecetaRayosX = -100
               End If
               Set oRsDetalleReceta = RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!IdReceta, _
                                      oRsCabeceraRecetas.Fields!idPuntoCarga, oConexMDB)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsRayosX.AddNew
                     oRsRayosX.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsRayosX.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsRayosX.Fields!cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsRayosX.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsRayosX.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     oRsRayosX.Fields!Receta = oRsCabeceraRecetas.Fields!IdReceta
                     oRsRayosX.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsRayosX.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsRayosX.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     oRsRayosX.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsRayosX.Fields!dx = IIf(IsNull(oRsDetalleReceta.Fields!dx), "", oRsDetalleReceta.Fields!dx)
                     If oRsDetalleReceta.Fields!precio > 0 Then
                        oRsRayosX.Fields!hayCpt = True
                     End If
                     oRsRayosX.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          Case sghPtoCargaEcogObstetrica
               lnRecetaEcografiaO = oRsCabeceraRecetas.Fields!IdReceta
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  lnRecetaEcografiaO = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                  End Select
               End If
               Set oRsDetalleReceta = RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!IdReceta, oRsCabeceraRecetas.Fields!idPuntoCarga, oConexMDB)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsEcografiaO.AddNew
                     oRsEcografiaO.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsEcografiaO.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsEcografiaO.Fields!cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsEcografiaO.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsEcografiaO.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     If oRsDetalleReceta.Fields!precio > 0 Then
                        oRsEcografiaO.Fields!hayCpt = True
                     End If
                     oRsEcografiaO.Fields!Receta = oRsCabeceraRecetas.Fields!IdReceta
                     oRsEcografiaO.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsEcografiaO.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsEcografiaO.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     oRsEcografiaO.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsEcografiaO.Fields!dx = IIf(IsNull(oRsDetalleReceta.Fields!dx), "", oRsDetalleReceta.Fields!dx)
                     oRsEcografiaO.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          Case sghPtoCargaEcogGeneral
               lnRecetaEcografiaG = oRsCabeceraRecetas.Fields!IdReceta
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  lnRecetaEcografiaG = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                  End Select
               End If
               Set oRsDetalleReceta = RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!IdReceta, oRsCabeceraRecetas.Fields!idPuntoCarga, oConexMDB)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsEcografiaG.AddNew
                     oRsEcografiaG.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsEcografiaG.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsEcografiaG.Fields!cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsEcografiaG.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsEcografiaG.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     If oRsDetalleReceta.Fields!precio > 0 Then
                        oRsEcografiaG.Fields!hayCpt = True
                     End If
                     oRsEcografiaG.Fields!Receta = oRsCabeceraRecetas.Fields!IdReceta
                     oRsEcografiaG.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsEcografiaG.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsEcografiaG.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     oRsEcografiaG.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsEcografiaG.Fields!dx = IIf(IsNull(oRsDetalleReceta.Fields!dx), "", oRsDetalleReceta.Fields!dx)
                     oRsEcografiaG.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          Case sghPtoCargaTomografia
               lnRecetaTomografia = oRsCabeceraRecetas.Fields!IdReceta
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  lnRecetaTomografia = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                  End Select
               End If
               Set oRsDetalleReceta = RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!IdReceta, oRsCabeceraRecetas.Fields!idPuntoCarga, oConexMDB)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsTomografia.AddNew
                     oRsTomografia.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsTomografia.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsTomografia.Fields!cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsTomografia.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsTomografia.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     If oRsDetalleReceta.Fields!precio > 0 Then
                        oRsTomografia.Fields!hayCpt = True
                     End If
                     oRsTomografia.Fields!Receta = oRsCabeceraRecetas.Fields!IdReceta
                     oRsTomografia.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsTomografia.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsTomografia.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     oRsTomografia.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsTomografia.Fields!dx = IIf(IsNull(oRsDetalleReceta.Fields!dx), "", oRsDetalleReceta.Fields!dx)
                     oRsTomografia.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          Case sghPtoCargaPatologiaClinica
               lnRecetaPatologiaC = oRsCabeceraRecetas.Fields!IdReceta
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  lnRecetaPatologiaC = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                  End Select
               End If
               Set oRsDetalleReceta = RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!IdReceta, oRsCabeceraRecetas.Fields!idPuntoCarga, oConexMDB)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsPatologia.AddNew
                     oRsPatologia.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsPatologia.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsPatologia.Fields!cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsPatologia.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsPatologia.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     If oRsDetalleReceta.Fields!precio > 0 Then
                        oRsPatologia.Fields!hayCpt = True
                     End If
                     oRsPatologia.Fields!Receta = oRsCabeceraRecetas.Fields!IdReceta
                     oRsPatologia.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsPatologia.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsPatologia.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     oRsPatologia.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsPatologia.Fields!dx = IIf(IsNull(oRsDetalleReceta.Fields!dx), "", oRsDetalleReceta.Fields!dx)
                     oRsPatologia.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          Case sghPtoCargaAnatomiaPatologica1
               lnRecetaAnatomiaP = oRsCabeceraRecetas.Fields!IdReceta
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  lnRecetaAnatomiaP = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                  End Select
               End If
               Set oRsDetalleReceta = RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!IdReceta, oRsCabeceraRecetas.Fields!idPuntoCarga, oConexMDB)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsAnatomia.AddNew
                     oRsAnatomia.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsAnatomia.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsAnatomia.Fields!cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsAnatomia.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsAnatomia.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     If oRsDetalleReceta.Fields!precio > 0 Then
                        oRsAnatomia.Fields!hayCpt = True
                     End If
                     oRsAnatomia.Fields!Receta = oRsCabeceraRecetas.Fields!IdReceta
                     oRsAnatomia.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsAnatomia.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsAnatomia.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     oRsAnatomia.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsAnatomia.Fields!dx = IIf(IsNull(oRsDetalleReceta.Fields!dx), "", oRsDetalleReceta.Fields!dx)
                     oRsAnatomia.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          Case sghPtoCargaBancoSangre1
               lnRecetaBancoS = oRsCabeceraRecetas.Fields!IdReceta
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  lnRecetaBancoS = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                  End Select
               End If
               Set oRsDetalleReceta = RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!IdReceta, oRsCabeceraRecetas.Fields!idPuntoCarga, oConexMDB)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsBanco.AddNew
                     oRsBanco.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsBanco.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsBanco.Fields!cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsBanco.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsBanco.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     If oRsDetalleReceta.Fields!precio > 0 Then
                        oRsBanco.Fields!hayCpt = True
                     End If
                     oRsBanco.Fields!Receta = oRsCabeceraRecetas.Fields!IdReceta
                     oRsBanco.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsBanco.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsBanco.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     oRsBanco.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsBanco.Fields!dx = IIf(IsNull(oRsDetalleReceta.Fields!dx), "", oRsDetalleReceta.Fields!dx)
                     oRsBanco.Update
                     oRsDetalleReceta.MoveNext
                  Loop
               End If
               oRsDetalleReceta.Close
          Case sghPtoCargaFarmacia
               'debb-24/06/2015
               ml_FechaReceta = oRsCabeceraRecetas.Fields!FechaVigencia
               '
               lnRecetaFarmacia = oRsCabeceraRecetas.Fields!IdReceta
               If oRsCabeceraRecetas.Fields!idEstado <> sghRecetaEstados.sighRecetaRegistrada Then
                  lnRecetaFarmacia = -100
                  Select Case oRsCabeceraRecetas.Fields!idEstado
                  Case sghRecetaEstados.sighRecetaDespachada, sghRecetaEstados.sighRecetaConBoleta
                  End Select
               End If
               Set oRsDetalleReceta = RecetasDevuelveDatosDelDetalle(oRsCabeceraRecetas.Fields!IdReceta, oRsCabeceraRecetas.Fields!idPuntoCarga, oConexMDB)
               If oRsDetalleReceta.RecordCount > 0 Then
                  oRsDetalleReceta.MoveFirst
                  Do While Not oRsDetalleReceta.EOF
                     oRsFarmacia.AddNew
                     oRsFarmacia.Fields!ID = oRsDetalleReceta.Fields!idItem
                     oRsFarmacia.Fields!procedimiento = oRsDetalleReceta.Fields!Producto
                     oRsFarmacia.Fields!cantidad = oRsDetalleReceta.Fields!CantidadPedida
                     oRsFarmacia.Fields!precio = oRsDetalleReceta.Fields!precio
                     oRsFarmacia.Fields!saldoActual = oRsDetalleReceta.Fields!SaldoEnRegistroReceta
                     If oRsDetalleReceta.Fields!SaldoEnRegistroReceta > 0 Then
                        oRsFarmacia.Fields!haySaldo = True
                     End If
                     oRsFarmacia.Fields!Receta = oRsCabeceraRecetas.Fields!IdReceta
                     oRsFarmacia.Fields!idDosisRecetada = IIf(IsNull(oRsDetalleReceta.Fields!idDosisRecetada), 0, oRsDetalleReceta.Fields!idDosisRecetada)
                     oRsFarmacia.Fields!idEstadoDetalle = IIf(IsNull(oRsDetalleReceta.Fields!idEstadoDetalle), 0, oRsDetalleReceta.Fields!idEstadoDetalle)
                     oRsFarmacia.Fields!MotivoAnulacionMedico = IIf(IsNull(oRsDetalleReceta.Fields!MotivoAnulacionMedico), "", oRsDetalleReceta.Fields!MotivoAnulacionMedico)
                     
                     oRsFarmacia.Fields!IdViaAdministracion = IIf(IsNull(oRsDetalleReceta.Fields!IdViaAdministracion), 0, oRsDetalleReceta.Fields!IdViaAdministracion) 'Actualizado 26092014
                     
                     oRsFarmacia.Fields!Observaciones = IIf(IsNull(oRsDetalleReceta.Fields!Observaciones), "", oRsDetalleReceta.Fields!Observaciones)
                     oRsFarmacia.Fields!FechaVigencia = ml_FechaReceta
                     oRsFarmacia.Fields!dx = IIf(IsNull(oRsDetalleReceta.Fields!dx), "", oRsDetalleReceta.Fields!dx)
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
       Set oRsCabeceraRecetas = Nothing
    
End Sub


Function ModificarDatos(mo_atenciones As DOAtencion, mo_Diagnosticos As Collection, _
                        mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales, _
                        mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, _
                        ml_ldFechaIngreso As Date, mo_cita As DOCita, ml_idCuentaAtencion As Long, _
                        ml_idUsuario As Long, lnRecetaRayosX As Long, lnRecetaEcografiaO As Long, _
                        lnRecetaEcografiaG As Long, lnRecetaTomografia As Long, lnRecetaAnatomiaP As Long, _
                        lnRecetaPatologiaC As Long, lnRecetaBancoS As Long, lnRecetaFarmacia As Long, _
                        oRsDevuelveRayosX As Recordset, oRsDevuelveEcografiaO As Recordset, _
                        oRsDevuelveEcografiaG As Recordset, oRsDevuelveTomografia As Recordset, _
                        oRsDevuelveAnatomia As Recordset, oRsDevuelvePatologia As Recordset, _
                        oRsDevuelveBancoSangre As Recordset, oRsDevuelveFarmacia As Recordset, _
                        ml_FechaReceta As Date, ms_NombrePaciente As String, wxParametro302 As String, _
                        oRsDevuelveDx As Recordset, oRsDevuelveRecetaAntesDeImprimir As Recordset, _
                        ml_lcServicio As String, oRsAtencionesCE As Recordset, mo_paciente As DOPaciente, _
                        oConexMDB As Connection, oConexion As Connection) As Boolean
                        
        Dim oRsTmp999 As New Recordset
        Dim oRsSoloLabActividades As New Recordset
        Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
        Dim oEpisodioClinico As EpisodioClinico
        Dim ml_AScorrelativo As Long, ms_MensajeError As String
        Dim oRsOtrosCpt As New Recordset, lnRecetaOtrosCpt As Long
        
        lcSql = "select idAtencion from Atenciones where idAtencion=0"
        oRsSoloLabActividades.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
        
        ms_MensajeError = ""
        ml_AScorrelativo = 0
        oEpisodioClinico = EpisodioClinicoDevuelveDatos
        '
'        If mi_Opcion = sghEliminar Then
'           mo_Atenciones.HoraEgreso = "99:99"
'        End If
        '
        ModificarDatos = mo_AdminAdmision.AdmisionCEModificarAM(mo_atenciones, mo_Diagnosticos, _
                                           mo_DoAtencionDatosAdicionales, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                           "", oEpisodioClinico, ml_ldFechaIngreso, mo_cita.IdCita, oRsSoloLabActividades, True)
        ms_MensajeError = mo_AdminAdmision.MensajeError
        If ms_MensajeError <> "" Then
           ms_MensajeError = ms_MensajeError
           txtProblemas.Text = txtProblemas.Text & "No se pudo registrar ATENCION para la cuenta: " & Trim(Str(mo_atenciones.idCuentaAtencion)) & Chr(13) & Chr(10)
        Else

            'If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
                If GrabaAtencionJamo(mo_lnIdTablaLISTBARITEMS, ml_lcServicio, mo_atenciones.idAtencion, _
                                     mo_lcNombrePc, mo_atenciones, oRsAtencionesCE, mo_paciente, oConexMDB) = True Then
'                   GrabaAtencionPerinatal
'                   GrabaAtencionPerinatalAS   'debb-09/06/2016
'                   GrabaAtencionProgramaMaterno
                End If
                If ml_FechaReceta = 0 Then
                   ml_FechaReceta = lcBuscaParametro.RetornaFechaHoraServidorSQL
                End If
                '
                lnRecetaRayosX = 0: lnRecetaEcografiaO = 0: lnRecetaEcografiaG = 0: lnRecetaTomografia = 0
                lnRecetaAnatomiaP = 0: lnRecetaPatologiaC = 0: lnRecetaBancoS = 0: lnRecetaFarmacia = 0
                lcSql = "select * from RecetaCabecera where idCuentaAtencion=" & ml_idCuentaAtencion
                oRsTmp999.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                If oRsTmp999.RecordCount > 0 Then
                   oRsTmp999.Filter = "idPuntoCarga=" & sghPtoCargaRayosX
                   If oRsTmp999.RecordCount > 0 Then
                      lnRecetaRayosX = oRsTmp999!IdReceta
                   End If
                   oRsTmp999.Filter = "idPuntoCarga=" & sghPtoCargaEcogObstetrica
                   If oRsTmp999.RecordCount > 0 Then
                      lnRecetaEcografiaO = oRsTmp999!IdReceta
                   End If
                   oRsTmp999.Filter = "idPuntoCarga=" & sghPtoCargaEcogGeneral
                   If oRsTmp999.RecordCount > 0 Then
                      lnRecetaEcografiaG = oRsTmp999!IdReceta
                   End If
                   oRsTmp999.Filter = "idPuntoCarga=" & sghPtoCargaTomografia
                   If oRsTmp999.RecordCount > 0 Then
                      lnRecetaTomografia = oRsTmp999!IdReceta
                   End If
                   oRsTmp999.Filter = "idPuntoCarga=" & sghPtoCargaAnatomiaPatologica1
                   If oRsTmp999.RecordCount > 0 Then
                      lnRecetaAnatomiaP = oRsTmp999!IdReceta
                   End If
                   oRsTmp999.Filter = "idPuntoCarga=" & sghPtoCargaBancoSangre1
                   If oRsTmp999.RecordCount > 0 Then
                      lnRecetaBancoS = oRsTmp999!IdReceta
                   End If
                   oRsTmp999.Filter = "idPuntoCarga=" & sghPtoCargaFarmacia
                   If oRsTmp999.RecordCount > 0 Then
                      lnRecetaFarmacia = oRsTmp999!IdReceta
                   End If
                   oRsTmp999.Filter = "idPuntoCarga=" & sghPtoCargaPatologiaClinica
                   If oRsTmp999.RecordCount > 0 Then
                      lnRecetaPatologiaC = oRsTmp999!IdReceta
                   End If
                End If
                oRsTmp999.Close
                '
                ModificarDatos = mo_AdminAdmision.RecetaModificar(ml_idCuentaAtencion, mo_atenciones.IdServicioIngreso, ml_idUsuario, _
                                                 lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, lnRecetaTomografia, _
                                                 lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, _
                                                 oRsDevuelveRayosX, oRsDevuelveEcografiaO, _
                                                 oRsDevuelveEcografiaG, oRsDevuelveTomografia, _
                                                 oRsDevuelveAnatomia, oRsDevuelvePatologia, _
                                                 oRsDevuelveBancoSangre, oRsDevuelveFarmacia, ml_FechaReceta, _
                                                 mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Paciente : " & ms_NombrePaciente, mo_atenciones.IdMedicoIngreso, False, oRsOtrosCpt, lnRecetaOtrosCpt)
                ms_MensajeError = mo_AdminAdmision.MensajeError
 '               If mi_Opcion <> sghEliminar Then
'                If lnRecetaRayosX > 0 Or lnRecetaEcografiaO > 0 Or lnRecetaEcografiaG > 0 Or lnRecetaTomografia > 0 Or _
'                         lnRecetaAnatomiaP > 0 Or lnRecetaPatologiaC > 0 Or lnRecetaBancoS > 0 Or lnRecetaFarmacia > 0 Then
'                   Me.UcRecetas1.Tratamiento = Trim(TxtCitaTratamiento.Text)
'                   Me.UcRecetas1.CargaNumeroDeRecetaEimprime lnRecetaRayosX, lnRecetaEcografiaO, lnRecetaEcografiaG, lnRecetaTomografia, _
'                                                             lnRecetaAnatomiaP, lnRecetaPatologiaC, lnRecetaBancoS, lnRecetaFarmacia, True
'                End If
'                End If
        '    End If
            '
            'no se considera los PAGANTES, porque se espera que vaya a CAJA, allí si se considera este proceso
            If mo_atenciones.IdFormaPago > 1 Then
                Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
                mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar ml_idCuentaAtencion, False, 0
                Set mo_ReglasFacturacion = Nothing
            End If
'            If wxParametro302 = "S" And mo_atenciones.idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
'               mo_ReglasSISgalenhos.SisFuaAtencionActualizaDatosDesdeHospEmegCE mo_atenciones.IdCuentaAtencion, _
'                                                                      mo_atenciones.IdTipoServicio, mo_atenciones.IdAtencion, _
'                                                                      mo_lnIdTablaLISTBARITEMS, ml_idUsuario
'            End If
            'debb-27/05/2015
'            Set mo_Diagnosticos = Nothing
'            Me.UcDiagnosticoDetalle1.CargarDiagnosticosAlObjetoDatosMenosCtaActual mo_Diagnosticos
'            If mo_AdminAdmision.AtencionesEnOtrosConsultoriosAlMismoTiempo(mo_paciente, mo_Atenciones, _
'                                                                           mo_DoAtencionDatosAdicionales, mo_CuentasAtencion, _
'                                                                           mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, _
'                                                                           mo_lcNombrePc, mi_Opcion, _
'                                                                           Me.UcDiagnosticoDetalle1.DevuelveDx, _
'                                                                           Me.UcDiagnosticoDetalle1.TipoDiagnostico, _
'                                                                           oRsGrdOtrosCpt, False, _
'                                                                           ml_AScorrelativo) = False Then
'            End If
'            If wxParametro302 = "S" And mo_atenciones.idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
'                If mo_AdminAdmision.ActualizaVariosFUASdeVariosConsultorios(mo_atenciones.IdAtencion, _
'                                                     orsDevuelveDx, oRsGrdOtrosCpt, _
'                                                     oRsDevuelveRayosX, oRsDevuelveEcografiaO, _
'                                                     oRsDevuelveEcografiaG, oRsDevuelveTomografia, _
'                                                     oRsDevuelveAnatomia, oRsDevuelvePatologia, _
'                                                     oRsDevuelveBancoSangre, oRsDevuelveFarmacia, _
'                                                     ml_AScorrelativo) Then
'                End If
'            End If
            '
       End If
       Set oRsTmp999 = Nothing
       Set mo_AdminAdmision = Nothing
       Set oRsSoloLabActividades = Nothing
End Function
Function EpisodioClinicoDevuelveDatos() As EpisodioClinico
        Dim oEpisodioClinico As EpisodioClinico
        oEpisodioClinico.idEpisodio = 1
        oEpisodioClinico.lbCierreEpisodio = False
        oEpisodioClinico.lbNuevoEpisodio = True
        EpisodioClinicoDevuelveDatos = oEpisodioClinico
End Function

Function GrabaAtencionJamo(mo_lnIdTablaLISTBARITEMS As Long, ml_lcServicio As String, _
                           ml_idAtencion As Long, mo_lcNombrePc As String, mo_atenciones As DOAtencion, _
                           oRsAtencionesCE As Recordset, mo_paciente As DOPaciente, oConexMDB As Connection) As Boolean
    If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then
        Dim oRsJAMO1 As New Recordset
        Dim oRsTmpBuscaAtencion As New Recordset
        Dim mo_DOAtencionesCE As New DOAtencionesCE
        Dim mo_AdminAdmision As New ReglasAdmision
        Dim txtCitaExClinicos As String
        txtCitaExClinicos = "...."   'falta llenar del MDB
        
        'cargar del MDB tabla Paciente
    
        
        
'        Select Case mi_Opcion
'        Case sghAgregar
'             GrabaAtencionJamo = True
'        Case sghModificar
             CargaDatosAtencionJamo mo_DOAtencionesCE, mo_atenciones, ml_lcServicio, mo_paciente, oRsAtencionesCE
             Set oRsTmpBuscaAtencion = mo_AdminAdmision.AtencionCESeleccionarPorIdAtencion(ml_idAtencion)
             mo_DOAtencionesCE.idAtencion = ml_idAtencion
             If oRsTmpBuscaAtencion.RecordCount = 0 Then
                GrabaAtencionJamo = mo_AdminAdmision.AtencionCEAgregar(mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, _
                                             mo_lcNombrePc, "IdAtencion: " & Trim(Str(ml_idAtencion)) & "(desde Atención)")
             Else
                GrabaAtencionJamo = mo_AdminAdmision.AtencionCEModificar(mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, _
                                             mo_lcNombrePc, "IdAtencion: " & Trim(Str(ml_idAtencion)) & "(desde Atención)")
             End If
             PacienteDatosAdicionalesGrabar mo_atenciones.idPaciente, mo_atenciones.IdUsuarioAuditoria, oConexMDB
             'debb-2/3/2015**inicio
             If Val(mo_DOAtencionesCE.TriajeTalla) > 0 And Val(mo_DOAtencionesCE.TriajePeso) > 0 Then
                Dim oConexion As New Connection
                Dim lnIdDxNutricional As Long, lnGrafXedadEnMeses As Long, lnGrafYpercentilTE As Long, lnGrafYpercentilPT As Long
                Dim lnGrafYpercentilPE As Long, lnZetaPT As Double, lnZetaTE As Double, lnZetaPE As Double
                Dim lnPercentilIMC As Double, lnPercentilIMC_Z As Double
                Dim lnPesoKg As Double, lnTallaCM As Long, ml_EdadEnMeses As Long, lnEdadEnAniosEnAtencion As Integer
                Dim ml_idTipoSexo As Long
                oConexion.CommandTimeout = 300
                oConexion.CursorLocation = adUseClient
                oConexion.Open SIGHEntidades.CadenaConexion
                lnPesoKg = Val(mo_DOAtencionesCE.TriajePeso)
                lnTallaCM = Val(mo_DOAtencionesCE.TriajeTalla)
                ml_EdadEnMeses = SIGHEntidades.DevuelveEdadEnMeses(mo_paciente.FechaNacimiento, mo_atenciones.FechaIngreso)
                lnEdadEnAniosEnAtencion = IIf(mo_atenciones.IdTipoEdad = 1, mo_atenciones.Edad, 0)
                ml_idTipoSexo = mo_paciente.idTipoSexo
                lnIdDxNutricional = 0
                lnGrafXedadEnMeses = ml_EdadEnMeses
                lnGrafYpercentilTE = 0
                lnGrafYpercentilPT = 0
                lnGrafYpercentilPE = 0
                lnPercentilIMC = 0
                lnZetaPT = 0
                lnZetaTE = 0
                lnZetaPE = 0
                lnPercentilIMC_Z = 0
                
                'Dim oProcesos As New Procesos
                'oProcesos.CalculaPercentiles lnPesoKg, lnTallaCM, ml_EdadEnMeses, _
                                             ml_idTipoSexo, lnEdadEnAniosEnAtencion, _
                                             lnGrafYpercentilPE, lnGrafYpercentilTE, lnGrafYpercentilPT, lnPercentilIMC, _
                                             lnZetaPE, lnZetaTE, lnZetaPT, lnPercentilIMC_Z
                'Set oProcesos = Nothing
                CalculaPercentiles lnPesoKg, lnTallaCM, ml_EdadEnMeses, _
                                             ml_idTipoSexo, lnEdadEnAniosEnAtencion, _
                                             lnGrafYpercentilPE, lnGrafYpercentilTE, lnGrafYpercentilPT, lnPercentilIMC, _
                                             lnZetaPE, lnZetaTE, lnZetaPT, lnPercentilIMC_Z
                
                '
                
                If lnZetaPE >= -1 And lnZetaPE <= 1 Then
                   lnIdDxNutricional = 1
                ElseIf lnZetaPE >= -2 And lnZetaPE < -1 Then
                   lnIdDxNutricional = 2
                ElseIf lnZetaPE < -2 Then
                   lnIdDxNutricional = 4
                ElseIf lnZetaPE >= 1 And lnZetaPE <= 2 Then
                   lnIdDxNutricional = 1
                ElseIf lnZetaPE > 2 And lnZetaPE <= 3 Then
                   lnIdDxNutricional = 5
                ElseIf lnZetaPE > 3 Then
                   lnIdDxNutricional = 6
                End If
                mo_AdminServiciosComunes.ActualizaTablaPacientesMovimientos oConexion, mo_atenciones.FechaIngreso, _
                                                   mo_atenciones.idCuentaAtencion, sghRegistroAtencionCE, False, Val(mo_DOAtencionesCE.TriajePeso), _
                                                   Val(mo_DOAtencionesCE.TriajeTalla), lnIdDxNutricional, lnGrafXedadEnMeses, _
                                                   lnGrafYpercentilTE, lnGrafYpercentilPT, lnGrafYpercentilPE, lnZetaPT, lnZetaTE, _
                                                   lnZetaPE
                oConexion.Close
                Set oConexion = Nothing
             End If
             'debb-2/3/2015****final
'        Case sghEliminar
'             CargaDatosAtencionJamo
'             GrabaAtencionJamo = mo_AdminAdmision.AtencionCEeliminar(mo_DOAtencionesCE, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "IdAtencion: " & Trim(Str(ml_idAtencion)) & "(desde Atención)")
'
'        End Select
        If GrabaAtencionJamo = False Then
           MsgBox "Falló al Grabar DATOS JAMO" & Chr(13) & mo_AdminAdmision.MensajeError
        End If
        Set oRsTmpBuscaAtencion = Nothing
    End If
    Set oRsTmpBuscaAtencion = Nothing
    Set mo_DOAtencionesCE = Nothing
    Set mo_paciente = Nothing
    Set mo_AdminAdmision = Nothing
End Function

Sub CargaDatosAtencionJamo(mo_DOAtencionesCE As DOAtencionesCE, mo_atenciones As DOAtencion, _
                           ml_lcServicio As String, mo_paciente As DOPaciente, oRsAtencionesCE As Recordset)
    Dim mo_AdminFacturacion As New ReglasFacturacion
    Dim oDOAtencionesCETriaje As DOAtencionesCE
    Dim oDOAtencionCE As New DOAtencionesCE
    Dim lcLineaChar As String
    Dim txtCitaDxMedico As String, _
        txtCitaExamenClinico As String, txtCitaMotivo As String, txtCitaObservaciones As String, _
        TxtCitaTratamiento As String, txtCitaAntecedente As String, ml_lnEdadEnDias As Long
    ml_lnEdadEnDias = 0        'hay que buscarlo
    lcLineaChar = "¨"
    With mo_DOAtencionesCE
        .idAtencion = mo_atenciones.idAtencion
        .CitaDiagMed = IIf(IsNull(oRsAtencionesCE!CitaDiagMed), "", oRsAtencionesCE!CitaDiagMed) 'Left(mo_AdminFacturacion.DevuelveDxAltaMedicaTodosDx(mo_atenciones.IdAtencion, 1), 600) & Chr(13) & Chr(10) & lcLineaChar & txtCitaDxMedico
        .CitaDniMedicoJamo = ""
        .CitaExamenClinico = IIf(IsNull(oRsAtencionesCE!CitaExamenClinico), "", oRsAtencionesCE!CitaExamenClinico) 'txtCitaExamenClinico
        .CitaExClinicos = "" 'Usado por compatibilidad con Hospital Jmo(Datos Importados de un sistema anterior)
        .CitaFecha = oRsAtencionesCE!CitaFecha 'CDate(mo_atenciones.FechaIngreso & " " & mo_atenciones.HoraIngreso)
        .CitaFechaAtencion = oRsAtencionesCE!CitaFechaAtencion   ' CDate(mo_atenciones.FechaIngreso & " " & mo_atenciones.HoraIngreso)
        .CitaIdServicio = oRsAtencionesCE!CitaIdServicio    ' mo_atenciones.IdServicioIngreso
        .CitaIdUsuario = oRsAtencionesCE!CitaIdUsuario  '  mo_atenciones.IdUsuarioAuditoria
        .CitaMedico = oRsAtencionesCE!CitaMedico   ' mo_atenciones.IdMedicoIngreso
        .CitaMotivo = IIf(IsNull(oRsAtencionesCE!CitaMotivo), "", oRsAtencionesCE!CitaMotivo) '  txtCitaMotivo
        .CitaObservaciones = IIf(IsNull(oRsAtencionesCE!CitaObservaciones), "", oRsAtencionesCE!CitaObservaciones) 'txtCitaObservaciones
        .CitaServicioJamo = oRsAtencionesCE!CitaServicioJamo   ' ml_lcServicio
        .CitaTratamiento = IIf(IsNull(oRsAtencionesCE!CitaTratamiento), "", oRsAtencionesCE!CitaTratamiento) ' TxtCitaTratamiento
        .IdUsuarioAuditoria = mo_atenciones.IdUsuarioAuditoria
        .NroHistoriaClinica = oRsAtencionesCE!NroHistoriaClinica   ' mo_paciente.NroHistoriaClinica
        .TriajeEdad = IIf(IsNull(oRsAtencionesCE!TriajeEdad), "", oRsAtencionesCE!TriajeEdad)
        'If IsNull(.IdAtencion) Then
           .TriajeFecha = oRsAtencionesCE!TriajeFecha   '  CDate(mo_atenciones.FechaIngreso & " " & mo_atenciones.HoraIngreso)
           .TriajeIdUsuario = mo_atenciones.IdUsuarioAuditoria
        'End If
        '
        'Call mo_AdminServiciosComunes.cargarDatosTriajeAObjetoDatos(mo_DOAtencionesCE, oDOAtencionCE)
        .TriajeFrecCardiaca = IIf(IsNull(oRsAtencionesCE!TriajeFrecCardiaca), 0, oRsAtencionesCE!TriajeFrecCardiaca) ' oDOAtencionesCETriaje.TriajeFrecCardiaca
        .TriajeFrecRespiratoria = IIf(IsNull(oRsAtencionesCE!TriajeFrecRespiratoria), 0, oRsAtencionesCE!TriajeFrecRespiratoria) '  Val(oDOAtencionesCETriaje.TriajeFrecRespiratoria)
        .TriajeOrigen = IIf(IsNull(oRsAtencionesCE!TriajeOrigen), 0, oRsAtencionesCE!TriajeOrigen) ' oDOAtencionesCETriaje.TriajeOrigen
        .TriajePerimCefalico = IIf(IsNull(oRsAtencionesCE!TriajePerimCefalico), 0, oRsAtencionesCE!TriajePerimCefalico) ' oDOAtencionesCETriaje.TriajePerimCefalico
        .TriajePeso = IIf(IsNull(oRsAtencionesCE!TriajePeso), "", oRsAtencionesCE!TriajePeso) ' oDOAtencionesCETriaje.TriajePeso
        .TriajePresion = IIf(IsNull(oRsAtencionesCE!TriajePresion), "", oRsAtencionesCE!TriajePresion) ' oDOAtencionesCETriaje.TriajePresion
        .TriajePulso = IIf(IsNull(oRsAtencionesCE!TriajePulso), 0, oRsAtencionesCE!TriajePulso) ' Val(oDOAtencionesCETriaje.TriajePulso)
        .TriajeTalla = IIf(IsNull(oRsAtencionesCE!TriajeTalla), "", oRsAtencionesCE!TriajeTalla) ' oDOAtencionesCETriaje.TriajeTalla
        .TriajeTemperatura = IIf(IsNull(oRsAtencionesCE!TriajeTemperatura), "", oRsAtencionesCE!TriajeTemperatura) ' oDOAtencionesCETriaje.TriajeTemperatura
        '
        .CitaAntecedente = IIf(IsNull(oRsAtencionesCE!CitaAntecedente), "", oRsAtencionesCE!CitaAntecedente) ' txtCitaAntecedente
    End With
    Set mo_AdminFacturacion = Nothing
    Set oDOAtencionesCETriaje = Nothing
    Set oDOAtencionCE = Nothing
End Sub

Sub PacienteDatosAdicionalesGrabar(ml_idPaciente As Long, ml_idUsuario As Long, oConexMDB As Connection)
    
    Dim oPacientesDatosAdd As New PacientesDatosAdd
    Dim oConexion As New Connection
    Dim oDoPacienteDatosAdd As New DoPacienteDatosAdd
    Dim txtAntecedentes As String, txtantecedAlergico As String, txtantecedObstetrico As String
    Dim txtantecedQuirurgico As String, txtantecedFamiliar As String, txtantecedPatologico As String
    Dim lbPacienteDatosAdicionalesEsNuevo As Boolean
    Dim oRsTmp9 As New Recordset
    
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    'buscar pacientesdAtosAdicionales (mdb)
    lcSql = "select * from PacientesDAtosAdicionales where idPaciente=" & Trim(Str(ml_idPaciente))
    oRsTmp9.Open lcSql, oConexMDB, adOpenKeyset, adLockOptimistic
    If oRsTmp9.RecordCount > 0 Then
        lbPacienteDatosAdicionalesEsNuevo = False
        oDoPacienteDatosAdd.idPaciente = ml_idPaciente
        oDoPacienteDatosAdd.IdUsuarioAuditoria = ml_idUsuario
        Set oPacientesDatosAdd.Conexion = oConexion
        If oPacientesDatosAdd.SeleccionarPorId(oDoPacienteDatosAdd) = False Then
            oDoPacienteDatosAdd.idPaciente = ml_idPaciente
            oDoPacienteDatosAdd.antecedentes = IIf(IsNull(oRsTmp9!antecedentes), "", oRsTmp9!antecedentes)
            oDoPacienteDatosAdd.antecedAlergico = IIf(IsNull(oRsTmp9!antecedAlergico), "", oRsTmp9!antecedAlergico)
            oDoPacienteDatosAdd.antecedObstetrico = IIf(IsNull(oRsTmp9!antecedObstetrico), "", oRsTmp9!antecedObstetrico)
            oDoPacienteDatosAdd.antecedQuirurgico = IIf(IsNull(oRsTmp9!antecedQuirurgico), "", oRsTmp9!antecedQuirurgico)
            oDoPacienteDatosAdd.antecedFamiliar = IIf(IsNull(oRsTmp9!antecedFamiliar), "", oRsTmp9!antecedFamiliar)
            oDoPacienteDatosAdd.antecedPatologico = IIf(IsNull(oRsTmp9!antecedPatologico), "", oRsTmp9!antecedPatologico)
            If oPacientesDatosAdd.Insertar(oDoPacienteDatosAdd) = False Then
              MsgBox oPacientesDatosAdd.MensajeError, vbInformation, "CE"
            End If
        Else
            oDoPacienteDatosAdd.antecedentes = IIf(IsNull(oRsTmp9!antecedentes), "", oRsTmp9!antecedentes)
            oDoPacienteDatosAdd.antecedAlergico = IIf(IsNull(oRsTmp9!antecedAlergico), "", oRsTmp9!antecedAlergico)
            oDoPacienteDatosAdd.antecedObstetrico = IIf(IsNull(oRsTmp9!antecedObstetrico), "", oRsTmp9!antecedObstetrico)
            oDoPacienteDatosAdd.antecedQuirurgico = IIf(IsNull(oRsTmp9!antecedQuirurgico), "", oRsTmp9!antecedQuirurgico)
            oDoPacienteDatosAdd.antecedFamiliar = IIf(IsNull(oRsTmp9!antecedFamiliar), "", oRsTmp9!antecedFamiliar)
            oDoPacienteDatosAdd.antecedPatologico = IIf(IsNull(oRsTmp9!antecedPatologico), "", oRsTmp9!antecedPatologico)
            If oPacientesDatosAdd.Modificar(oDoPacienteDatosAdd) = False Then
              MsgBox oPacientesDatosAdd.MensajeError, vbInformation, "CE"
            End If
        End If
    End If
    oRsTmp9.Close
    oConexion.Close
    Set oConexion = Nothing
    Set oPacientesDatosAdd = Nothing
    Set oDoPacienteDatosAdd = Nothing
    Set oRsTmp9 = Nothing
End Sub


Function RecetasDevuelveDatosDelDetalle(lnIdReceta As Long, lnIdPuntoCarga As Long, oConexMDB As Connection) As Recordset
      Dim oRsTmp As New Recordset
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      
      
      If lnIdPuntoCarga = 5 Then
         lcSql = "SELECT      RecetaDetalle.*, 'xx' as Producto " & _
                  "   FROM          RecetaDetalle " & _
                  "                       " & _
                  "   Where         RecetaDetalle.idReceta=" & Trim(Str(lnIdReceta))
      Else
         lcSql = "SELECT      RecetaDetalle.*, 'xx' as Producto " & _
               "      FROM          RecetaDetalle " & _
               "                          " & _
               "      Where         RecetaDetalle.idReceta=" & Trim(Str(lnIdReceta))
      End If
      oRsTmp.Open lcSql, oConexMDB, adOpenKeyset, adLockOptimistic
      Set RecetasDevuelveDatosDelDetalle = oRsTmp

End Function


Function AtencionesDiagnosticosSeleccionarPorAtencion(lIdAtencion As Long, lIdTipoDiagnostico As Long, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcSql As String
   lcSql = "select " & _
" AtencionesDiagnosticos.IdDiagnostico," & _
" AtencionesDiagnosticos.idClasificacionDx," & _
" " & _
" " & _
" AtencionesDiagnosticos.labConfHIS," & _
" AtencionesDiagnosticos.IdSubClasificacionDx," & _
" " & _
" AtencionesDiagnosticos.GrupoHIS , AtencionesDiagnosticos.SubGrupoHIS" & _
" from AtencionesDiagnosticos " & _
" " & _
" " & _
" where AtencionesDiagnosticos.IdAtencion =  " & Trim(Str(lIdAtencion)) & _
" and AtencionesDiagnosticos.IdClasificacionDx = " & Trim(Str(lIdTipoDiagnostico))
oRecordset.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
   Set AtencionesDiagnosticosSeleccionarPorAtencion = oRecordset
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
   Exit Function
End Function

