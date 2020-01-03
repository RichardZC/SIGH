VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CFFE0A60-8E3A-11D3-BCC0-00104B9E0792}#1.0#0"; "ssInput1.ocx"
Object = "{0002E558-0000-0000-C000-000000000046}#1.1#0"; "OWC11.DLL"
Begin VB.UserControl UcPrograma 
   BackColor       =   &H00FF8080&
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11685
   ScaleHeight     =   7110
   ScaleWidth      =   11685
   Begin TabDlg.SSTab TabPrograma 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Primera Entrevista"
      TabPicture(0)   =   "UcPrograma.ctx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FraCabeceraPrograma"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Control"
      TabPicture(1)   =   "UcPrograma.ctx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FrameControl"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FraHistoricoControles"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frmGrafico"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame frmGrafico 
         Height          =   2295
         Left            =   8280
         TabIndex        =   22
         Top             =   4680
         Width           =   3315
         Begin OWC11.ChartSpace ChartSpace1 
            Height          =   2205
            Left            =   30
            OleObjectBlob   =   "UcPrograma.ctx":0038
            TabIndex        =   23
            Top             =   30
            Width           =   3225
         End
      End
      Begin VB.Frame FraHistoricoControles 
         Caption         =   "Histórico Controles"
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
         Height          =   4275
         Left            =   8280
         TabIndex        =   20
         Top             =   360
         Width           =   3255
         Begin UltraGrid.SSUltraGrid grdHistoricoControles 
            Height          =   3435
            Left            =   120
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "Dar doble clic para seleccionar el control actual"
            Top             =   240
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   6059
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            RowConnectorColor=   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "grdHistoricoControles"
         End
         Begin VB.Label lblMensajeAyuda 
            Caption         =   "Dar doble clic para seleccionar el control actual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   27
            Top             =   3720
            Width           =   3015
         End
      End
      Begin VB.Frame FrameControl 
         Caption         =   "Control Nº 1"
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
         Height          =   6675
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   8120
         Begin VB.Frame FraTratamiento 
            Caption         =   "Medicamento/Insumo"
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
            Left            =   120
            TabIndex        =   24
            Top             =   4680
            Width           =   7920
            Begin VB.CommandButton btnAgregarInsumo 
               DisabledPicture =   "UcPrograma.ctx":0C2C
               DownPicture     =   "UcPrograma.ctx":1015
               Height          =   315
               Left            =   6240
               Picture         =   "UcPrograma.ctx":1421
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton btnQuitarInsumo 
               DisabledPicture =   "UcPrograma.ctx":182D
               DownPicture     =   "UcPrograma.ctx":1BB8
               Height          =   315
               Left            =   6615
               Picture         =   "UcPrograma.ctx":1F4B
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "Elimina todos los CPT"
               Top             =   240
               Width           =   405
            End
            Begin VB.ComboBox cmbTtoProgama 
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
               ItemData        =   "UcPrograma.ctx":22DC
               Left            =   960
               List            =   "UcPrograma.ctx":22DE
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   240
               Width           =   5295
            End
            Begin VB.TextBox txtFiltroInsumos 
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
               Left            =   50
               TabIndex        =   34
               ToolTipText     =   "Filtrar listado de diagnósticos"
               Top             =   240
               Width           =   855
            End
            Begin UltraGrid.SSUltraGrid grdTratamientos 
               Height          =   1245
               Left            =   45
               TabIndex        =   25
               Top             =   600
               Width           =   7800
               _ExtentX        =   13758
               _ExtentY        =   2196
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
               Caption         =   "Medicamento/Insumo"
            End
         End
         Begin VB.Frame fraDiagnosticos 
            Caption         =   "Diagnósticos"
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
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   7905
            Begin VB.ComboBox cmbDxPrograma 
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
               ItemData        =   "UcPrograma.ctx":22E0
               Left            =   960
               List            =   "UcPrograma.ctx":22E2
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   240
               Width           =   5685
            End
            Begin VB.CommandButton btnBusquedaDiagnostico 
               Caption         =   "..."
               Height          =   315
               Left            =   7515
               TabIndex        =   13
               TabStop         =   0   'False
               ToolTipText     =   "Otros Diagnósticos"
               Top             =   240
               Width           =   315
            End
            Begin VB.TextBox txtFiltroDx 
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
               Left            =   50
               TabIndex        =   12
               ToolTipText     =   "Filtrar listado de diagnósticos"
               Top             =   240
               Width           =   855
            End
            Begin VB.CommandButton btnQuitar 
               DisabledPicture =   "UcPrograma.ctx":22E4
               DownPicture     =   "UcPrograma.ctx":266F
               Height          =   315
               Left            =   7095
               Picture         =   "UcPrograma.ctx":2A02
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Elimina el Dx"
               Top             =   225
               Width           =   360
            End
            Begin VB.CommandButton btnAgregar 
               DisabledPicture =   "UcPrograma.ctx":2D93
               DownPicture     =   "UcPrograma.ctx":317C
               Height          =   315
               Left            =   6675
               Picture         =   "UcPrograma.ctx":3588
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   225
               Width           =   390
            End
            Begin UltraGrid.SSUltraGrid grdDiagnosticos 
               Height          =   1215
               Left            =   45
               TabIndex        =   14
               Top             =   600
               Width           =   7770
               _ExtentX        =   13705
               _ExtentY        =   2143
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
               Caption         =   "grdDiagnosticos"
            End
            Begin PVCOMBOLibCtl.PVComboBox cmbLabHisDx 
               Height          =   330
               Left            =   6675
               TabIndex        =   38
               Top             =   240
               Visible         =   0   'False
               Width           =   1185
               _Version        =   524288
               _cx             =   2090
               _cy             =   582
               Appearance      =   1
               Enabled         =   -1  'True
               BackColor       =   16777215
               ForeColor       =   0
               Locked          =   0   'False
               Style           =   0
               Sorted          =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ShowPictures    =   0   'False
               ColumnHeaders   =   -1  'True
               PrimaryColumn   =   1
               VisibleItems    =   10
               ColumnHeaderHeight=   20
               ListMember      =   ""
               ColumnHeaderForeColor=   0
               ColumnHeaderBackColor=   13160660
               SelectedForeColor=   16777215
               SelectedBackColor=   6956042
               AlternateBackColor=   16777215
               ItemLabelStyle  =   1
               ItemLabelType   =   0
               ItemLabelWidth  =   20
               ItemLabelForeColor=   0
               ItemLabelBackColor=   13160660
               ColumnHeaderStyle=   0
               VerticalGridLines=   -1  'True
               HorizontalGridLines=   -1  'True
               ColumnResize    =   0   'False
               ItemLabelResize =   0   'False
               AllowDBAutoConfig=   0   'False
               GridLineColor   =   13421772
               List            =   ""
               NullString      =   "[NULL]"
               DropShadow      =   -1  'True
               Text            =   ""
               SortOnColumnHeaderClick=   0   'False
               DropEffect      =   1
               ColumnCount     =   3
               Column0.Heading =   "Id"
               Column0.Width   =   10
               Column0.Alignment=   0
               Column0.Hidden  =   -1  'True
               Column0.Name    =   "IdHisSituacio"
               Column0.Format  =   ""
               Column0.Bound   =   -1  'True
               Column0.Locked  =   0   'False
               Column0.HeaderAlignment=   0
               Column1.Heading =   "Valores"
               Column1.Width   =   35
               Column1.Alignment=   0
               Column1.Hidden  =   0   'False
               Column1.Name    =   "valores"
               Column1.Format  =   ""
               Column1.Bound   =   -1  'True
               Column1.Locked  =   0   'False
               Column1.HeaderAlignment=   0
               Column2.Heading =   "Descripción"
               Column2.Width   =   100
               Column2.Alignment=   0
               Column2.Hidden  =   0   'False
               Column2.Name    =   "descripcio"
               Column2.Format  =   ""
               Column2.Bound   =   -1  'True
               Column2.Locked  =   0   'False
               Column2.HeaderAlignment=   0
               SortKey1.Column =   -1
               SortKey1.Ascending=   -1  'True
               SortKey1.CaseInsensitive=   -1  'True
               SortKey2.Column =   -1
               SortKey2.Ascending=   -1  'True
               SortKey2.CaseInsensitive=   -1  'True
               SortKey3.Column =   -1
               SortKey3.Ascending=   -1  'True
               SortKey3.CaseInsensitive=   -1  'True
               BoundColumn     =   ""
               Border          =   -1  'True
               VertAlign       =   1
               Format          =   ""
            End
         End
         Begin VB.Frame FraProcedimientos 
            Caption         =   "Procedimientos"
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
            Height          =   2130
            Left            =   120
            TabIndex        =   7
            Top             =   2520
            Width           =   7905
            Begin VB.TextBox txtFiltraCPT 
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
               Left            =   50
               TabIndex        =   33
               ToolTipText     =   "Filtrar listado de diagnósticos"
               Top             =   240
               Width           =   855
            End
            Begin VB.ComboBox cmbProcPrograma 
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
               ItemData        =   "UcPrograma.ctx":3994
               Left            =   960
               List            =   "UcPrograma.ctx":3996
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   240
               Width           =   5685
            End
            Begin VB.CommandButton btnBuscarProc 
               Caption         =   "..."
               Height          =   315
               Left            =   7515
               TabIndex        =   30
               TabStop         =   0   'False
               ToolTipText     =   "Otros Diagnósticos"
               Top             =   240
               Width           =   315
            End
            Begin VB.CommandButton btnQuitaOtrosProcedimientos 
               DisabledPicture =   "UcPrograma.ctx":3998
               DownPicture     =   "UcPrograma.ctx":3D23
               Height          =   315
               Left            =   7095
               Picture         =   "UcPrograma.ctx":40B6
               Style           =   1  'Graphical
               TabIndex        =   29
               ToolTipText     =   "Elimina todos los CPT"
               Top             =   240
               Width           =   405
            End
            Begin VB.CommandButton btnAgregarProcedminiento 
               DisabledPicture =   "UcPrograma.ctx":4447
               DownPicture     =   "UcPrograma.ctx":4830
               Height          =   315
               Left            =   6675
               Picture         =   "UcPrograma.ctx":4C3C
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   240
               Width           =   390
            End
            Begin UltraGrid.SSUltraGrid grdProcedimientos 
               Height          =   1455
               Left            =   45
               TabIndex        =   8
               Top             =   600
               Width           =   7770
               _ExtentX        =   13705
               _ExtentY        =   2566
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
               Caption         =   "grdProcedimientos"
            End
            Begin PVCOMBOLibCtl.PVComboBox cmbLabHisCpt 
               Height          =   330
               Left            =   6675
               TabIndex        =   39
               Top             =   240
               Visible         =   0   'False
               Width           =   1185
               _Version        =   524288
               _cx             =   2090
               _cy             =   582
               Appearance      =   1
               Enabled         =   -1  'True
               BackColor       =   16777215
               ForeColor       =   0
               Locked          =   0   'False
               Style           =   0
               Sorted          =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ShowPictures    =   0   'False
               ColumnHeaders   =   -1  'True
               PrimaryColumn   =   1
               VisibleItems    =   10
               ColumnHeaderHeight=   20
               ListMember      =   ""
               ColumnHeaderForeColor=   0
               ColumnHeaderBackColor=   13160660
               SelectedForeColor=   16777215
               SelectedBackColor=   6956042
               AlternateBackColor=   16777215
               ItemLabelStyle  =   1
               ItemLabelType   =   0
               ItemLabelWidth  =   20
               ItemLabelForeColor=   0
               ItemLabelBackColor=   13160660
               ColumnHeaderStyle=   0
               VerticalGridLines=   -1  'True
               HorizontalGridLines=   -1  'True
               ColumnResize    =   0   'False
               ItemLabelResize =   0   'False
               AllowDBAutoConfig=   0   'False
               GridLineColor   =   13421772
               List            =   ""
               NullString      =   "[NULL]"
               DropShadow      =   -1  'True
               Text            =   ""
               SortOnColumnHeaderClick=   0   'False
               DropEffect      =   1
               ColumnCount     =   3
               Column0.Heading =   "Id"
               Column0.Width   =   10
               Column0.Alignment=   0
               Column0.Hidden  =   -1  'True
               Column0.Name    =   "IdHisSituacio"
               Column0.Format  =   ""
               Column0.Bound   =   -1  'True
               Column0.Locked  =   0   'False
               Column0.HeaderAlignment=   0
               Column1.Heading =   "Valores"
               Column1.Width   =   35
               Column1.Alignment=   0
               Column1.Hidden  =   0   'False
               Column1.Name    =   "valores"
               Column1.Format  =   ""
               Column1.Bound   =   -1  'True
               Column1.Locked  =   0   'False
               Column1.HeaderAlignment=   0
               Column2.Heading =   "Descripción"
               Column2.Width   =   100
               Column2.Alignment=   0
               Column2.Hidden  =   0   'False
               Column2.Name    =   "descripcio"
               Column2.Format  =   ""
               Column2.Bound   =   -1  'True
               Column2.Locked  =   0   'False
               Column2.HeaderAlignment=   0
               SortKey1.Column =   -1
               SortKey1.Ascending=   -1  'True
               SortKey1.CaseInsensitive=   -1  'True
               SortKey2.Column =   -1
               SortKey2.Ascending=   -1  'True
               SortKey2.CaseInsensitive=   -1  'True
               SortKey3.Column =   -1
               SortKey3.Ascending=   -1  'True
               SortKey3.CaseInsensitive=   -1  'True
               BoundColumn     =   ""
               Border          =   -1  'True
               VertAlign       =   1
               Format          =   ""
            End
         End
         Begin MSMask.MaskEdBox txtFechaControl 
            Height          =   315
            Left            =   2280
            TabIndex        =   15
            Tag             =   "__/__/____"
            Top             =   255
            Width           =   1425
            _ExtentX        =   2514
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
         Begin MSMask.MaskEdBox txtDatoControl 
            Height          =   315
            Index           =   0
            Left            =   6000
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin ActiveInput.SSComboBoxEx cmbDiagnosticos 
            Height          =   345
            Left            =   6600
            TabIndex        =   17
            Top             =   1320
            Visible         =   0   'False
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   609
            _Version        =   65536
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "SSComboBoxEx1"
            CheckList       =   -1  'True
            MultiSelect     =   -1  'True
            OverrideText    =   "(multiple items selected)"
            Separator       =   " | "
         End
         Begin VB.Label lblDatoControl 
            Caption         =   "ValorTexto"
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
            Index           =   0
            Left            =   6480
            TabIndex        =   19
            Top             =   480
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label lblFechaControl 
            Caption         =   "Fecha de control"
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
            Left            =   120
            TabIndex        =   18
            Top             =   255
            Width           =   2085
         End
      End
      Begin VB.Frame FraCabeceraPrograma 
         Caption         =   "Primera Entrevista Gestante"
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
         Height          =   6675
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   11475
         Begin VB.CommandButton btnInicializarControles 
            DisabledPicture =   "UcPrograma.ctx":5048
            DownPicture     =   "UcPrograma.ctx":5431
            Height          =   315
            Left            =   11160
            Picture         =   "UcPrograma.ctx":583D
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Iniciar los controles"
            Top             =   120
            Width           =   390
         End
         Begin VB.CheckBox chkDatoCabecera 
            Height          =   315
            Index           =   0
            Left            =   10080
            TabIndex        =   2
            Top             =   480
            Visible         =   0   'False
            Width           =   1425
         End
         Begin MSMask.MaskEdBox txtDatoCabecera 
            Height          =   315
            Index           =   0
            Left            =   10080
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label lblUnidades 
            Caption         =   "Kg."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   10080
            TabIndex        =   26
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblDatoCabecera 
            Caption         =   "ValorTexto"
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
            Index           =   0
            Left            =   10560
            TabIndex        =   5
            Top             =   360
            Visible         =   0   'False
            Width           =   1020
         End
      End
   End
   Begin VB.Menu mnuControlesEESS 
      Caption         =   "mnuControlesEESS"
      Begin VB.Menu mnuAgregarControl 
         Caption         =   "Agregar Control"
      End
      Begin VB.Menu mnuModificarControl 
         Caption         =   "Modificar Control"
      End
      Begin VB.Menu mnuEliminarControl 
         Caption         =   "Eliminar Control"
      End
   End
End
Attribute VB_Name = "UcPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para Programa Materno
'        Programado por: Cachay F
'        Fecha: Enero 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lnIdModulo As sighentidades.sghPerinatalModulos
Dim ml_IdPrograma As Long
Dim ml_IdProCabecera As Long
Dim ml_IdControl As Long
Dim ml_IdControlSeleccionado As Long
Dim ml_IdUltimoControl As Long
Dim ml_idAtencion As Long
Dim ml_idTipoSexo As Long
Dim ml_IdPaciente As Long
Dim ml_FechaAtencion As Date
Dim ml_idUsuario As Long
Dim ml_IdTabCtrl As Integer
Dim ml_peso As Double
Dim ml_presion As String
Dim ml_talla As Long
Dim ml_EdadEnDias As Long
Dim mb_ConsultaControl As Boolean
Dim mc_MensajeValidacion As String
Dim mo_cmbDiagnostico As New sighentidades.ListaDespleglable
Dim oRsProCabeceraConfigDatos As New Recordset
Dim oRsProControlConfigDatos As New Recordset
Dim oRsProCatalogoDiagnosticos As New Recordset
Dim oRsProCatalogoProcedimientos As New Recordset
Dim oRsProCatalogoTratamientos As New Recordset
Dim oRsProCatalogoControles As New Recordset
Dim oRsPercentil As New Recordset
Const lcCombo As String = "o"
Const ml_ColorCorrecto As Long = &HFFFFFF
Const ml_ColorError As Long = &HFF6347
Dim ml_SeteoUnaSolaVesCheckBox As Boolean
Dim mo_cmbDxPrograma As New sighentidades.ListaDespleglable
Dim mo_cmbCptPrograma As New sighentidades.ListaDespleglable
Dim mo_cmbTtoPrograma As New sighentidades.ListaDespleglable

Const lnPercentilNull As Long = 0
Dim lnPercentilPE As Long, lnPercentilTE As Long, lnPercentilPT As Long, lnPercentilIMC As Double
Dim lnPercentilPE_Z As Double, lnPercentilTE_Z As Double, lnPercentilPT_Z As Double, lnPercentilIMC_Z As Double

Dim xValues As Variant, yValues As Variant, yValues2 As Variant, yValues3 As Variant, yValues4 As Variant, yValues5 As Variant
Dim owcChart As OWC11.ChChart
Dim owcSeries As OWC11.ChSeries
Dim lnNroPuntosGraficos As Integer
Dim ml_YaCargoUnaSolaVez As Boolean
Dim mb_ControlNuevo As Boolean 'Actualizado 27102014
Dim mo_RsLabHis As ADODB.Recordset
Dim ml_IdPuntoCargaHosp As Long
Dim ml_IdServicioIngreso As Long
Dim ml_idCuentaAtencion As Long
Dim ml_IdFormaPago As Long
'Dim mo_RsLabHis As ADODB.Recordset
'
Property Let IdPrograma(lValue As Long)
   ml_IdPrograma = lValue
End Property
Property Let idAtencion(lValue As Long)
   ml_idAtencion = lValue
End Property
Property Let idTipoSexo(lValue As Long)
   ml_idTipoSexo = lValue
End Property
Property Let idPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Let FechaAtencion(lValue As Date)
   ml_FechaAtencion = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get MensajeValidacion() As String
    MensajeValidacion = mc_MensajeValidacion
End Property

Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property

Property Let IdFormaPago(lValue As Long)
   ml_IdFormaPago = lValue
End Property

Property Let IdServicioIngreso(lValue As Long)
   ml_IdServicioIngreso = lValue
   ml_IdPuntoCargaHosp = getIdPuntoCargaHospitalizacion(lValue)
End Property

'Actualizado 27102014
Property Let ControlNuevo(lValue As Boolean)
   mb_ControlNuevo = lValue
End Property

Public Sub Inicializar()
    If ml_YaCargoUnaSolaVez = False Then
        ml_YaCargoUnaSolaVez = True
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        
        ml_IdTabCtrl = 1 'Poniendo tab a los controles del formulario
        
        'Crea temporal de configuracion Cabecera y Controles
        CreaTemporalProControlConfigDatos
        CreaTemporaloRsProCabeceraConfigDatos
        
        'Configura Cabecera del programa
        ConfiguraCabeceraPrograma oConexion
        
        'Configura Control del Programa
        ConfiguraControlesPrograma oConexion
        
        'Inicializa CatalogoControles
        CreaTemporalHistoricoControles
        InicializarLaGrilla grdHistoricoControles
        CargaCatalogoControles oConexion
        
        'Inicializa Diagnosticos
        CargaComboProDiagnosticos oConexion
        CargaListaTiposDiagnosticos oConexion
        CreaTemporalDiagnosticos
        InicializarLaGrilla grdDiagnosticos
        LimpiaDiagnosticosFrecuentes False
        
        'Inicializa Procedimientos
        CargaComboProProcedimientos oConexion
        CargaListaResultadosDeProcedimientos oConexion
        CreaTemporalProcedimientos
        InicializarLaGrilla grdProcedimientos
        
        'Inicializa Tratamientos
        CargaComboProTratamientos oConexion
        CreaTemporalTratamientos
        InicializarLaGrilla grdTratamientos
'        CargaCatalogoTratamientos oConexion
        
        oConexion.Close
        Set oConexion = Nothing
    End If
End Sub

Private Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
    Select Case oGrilla.Name
    Case "grdHistoricoControles"
        oGrilla.Bands(0).Columns("idcontrol").Hidden = True
        oGrilla.Bands(0).Columns("descripcion").Header.Caption = "Controles"
        oGrilla.Bands(0).Columns("descripcion").Width = 950
        oGrilla.Bands(0).Columns("descripcion").Activation = ssActivationActivateNoEdit
        oGrilla.Bands(0).Columns("FechaControl").Header.Caption = "Fecha de Control"
        oGrilla.Bands(0).Columns("FechaControl").Width = 1700
        oGrilla.Bands(0).Columns("FechaControl").Activation = ssActivationActivateNoEdit
        oGrilla.Bands(0).Columns("ControlOtroEESS").Hidden = True
        oGrilla.Bands(0).Columns("IdEstablecimiento").Hidden = True
    Case "grdDiagnosticos"
         oGrilla.Bands(0).Columns("IdDiagnostico").Hidden = True
         oGrilla.Bands(0).Columns("CodigoCIE10").Hidden = True
         'mgaray201412a
         oGrilla.Bands(0).Columns("labConfHIS").Hidden = True
'         oGrilla.Bands(0).Columns("CodigoCIE10").Header.Caption = "Dx"
'         oGrilla.Bands(0).Columns("CodigoCIE10").Width = 800
'         oGrilla.Bands(0).Columns("CodigoCIE10").Activation = ssActivationActivateNoEdit
         oGrilla.Bands(0).Columns("IdSubClasificacionDX").Header.Caption = "Tipo"
         oGrilla.Bands(0).Columns("IdSubClasificacionDX").Width = 600
         oGrilla.Bands(0).Columns("IdSubClasificacionDX").ValueList = "TipoDiagnostico"
         oGrilla.Bands(0).Columns("IdSubClasificacionDX").Style = ssStyleDropDownValidate
         oGrilla.Bands(0).Columns("Descripcion").Header.Caption = "Diagnosticos"
         oGrilla.Bands(0).Columns("Descripcion").Width = 5700 '4900
         oGrilla.Bands(0).Columns("Descripcion").Activation = ssActivationActivateNoEdit
         oGrilla.Bands(0).Columns("labConfHIS").Header.Caption = "Lab"
         oGrilla.Bands(0).Columns("labConfHIS").Width = 800
         oGrilla.Bands(0).Columns("labConfHIS").Activation = ssActivationAllowEdit
         oGrilla.Bands(0).Columns("Principal").Header.Caption = "Principal"
         oGrilla.Bands(0).Columns("Principal").Width = 800
         Call AsignarListaDeLabsEnGridaDiagnosticos(oGrilla, "labConfHIS")
    Case "grdProcedimientos"
         oGrilla.Bands(0).Columns("IdDiagnostico").Hidden = True
         oGrilla.Bands(0).Columns("IdProducto").Hidden = True
         'mgaray201412a
         oGrilla.Bands(0).Columns("labConfHIS").Hidden = True
'         oGrilla.Bands(0).Columns("Procedimiento").Width = 3620
         oGrilla.Bands(0).Columns("Procedimiento").Width = 6000 '5200
         oGrilla.Bands(0).Columns("Procedimiento").Header.Caption = "Procedimiento"
         oGrilla.Bands(0).Columns("Procedimiento").Activation = ssActivationActivateNoEdit
'         oGrilla.Bands(0).Columns("CodigoCIE10").Width = 900
'         oGrilla.Bands(0).Columns("CodigoCIE10").Header.Caption = "Dx"
'         oGrilla.Bands(0).Columns("CodigoCIE10").Activation = ssActivationActivateNoEdit
         oGrilla.Bands(0).Columns("CodigoCIE10").Hidden = True
         oGrilla.Bands(0).Columns("labConfHIS").Header.Caption = "Lab"
         oGrilla.Bands(0).Columns("labConfHIS").Width = 800
         oGrilla.Bands(0).Columns("labConfHIS").Activation = ssActivationAllowEdit
         oGrilla.Bands(0).Columns("IdResultado").Header.Caption = "Resultados"
         oGrilla.Bands(0).Columns("IdResultado").Width = 1200
         oGrilla.Bands(0).Columns("IdResultado").ValueList = "TipoResultado"
         oGrilla.Bands(0).Columns("IdResultado").Style = ssStyleDropDownValidate
         oGrilla.Bands(0).Columns("seleccionar").Width = 800
         oGrilla.Bands(0).Columns("seleccionar").Header.Caption = "Seleccionar"
         oGrilla.Bands(0).Columns("seleccionar").Hidden = True
         Call AsignarListaDeLabsEnGridaDiagnosticos(oGrilla, "labConfHIS")
    Case "grdTratamientos"
         oGrilla.Bands(0).Columns("IdProducto").Hidden = True
         oGrilla.Bands(0).Columns("IdProducto").Hidden = True
         oGrilla.Bands(0).Columns("Tratamientos").Header.Caption = "Insumo"
         oGrilla.Bands(0).Columns("Tratamientos").Width = 7150
         oGrilla.Bands(0).Columns("Tratamientos").Activation = ssActivationActivateNoEdit
         oGrilla.Bands(0).Columns("seleccionar").Width = 1200
         oGrilla.Bands(0).Columns("seleccionar").Hidden = True
    End Select
End Sub

Sub LimpiaDiagnosticos()
    CreaTemporalDiagnosticos
End Sub
Sub LimpiaHistoricoControles()
    CreaTemporalHistoricoControles
End Sub
Sub LimpiaProcedimientos()
    CreaTemporalProcedimientos
End Sub
Sub LimpiaTratamientos()
    CreaTemporalTratamientos
End Sub

Sub CreaTemporalHistoricoControles()
    If oRsProCatalogoControles.State = 1 Then
       Set oRsProCatalogoControles = Nothing
    End If
    With oRsProCatalogoControles
          .Fields.Append "idcontrol", adInteger, 4, adFldIsNullable
          .Fields.Append "descripcion", adVarChar, 100, adFldIsNullable
          .Fields.Append "FechaControl", adVarChar, 100, adFldIsNullable
          .Fields.Append "ControlOtroEESS", adBoolean
          .Fields.Append "IdEstablecimiento", adInteger, 0, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdHistoricoControles.DataSource = oRsProCatalogoControles
    mo_Apariencia.ConfigurarFilasBiColores grdHistoricoControles, sighentidades.GrillaConFilasBicolor
End Sub

Sub CreaTemporalDiagnosticos()
    If oRsProCatalogoDiagnosticos.State = 1 Then
       Set oRsProCatalogoDiagnosticos = Nothing
    End If
    With oRsProCatalogoDiagnosticos
          .Fields.Append "IdSubClasificacionDX", adInteger
          .Fields.Append "IdDiagnostico", adInteger
          .Fields.Append "CodigoCIE10", adVarChar, 255, adFldIsNullable
          .Fields.Append "Descripcion", adVarChar, 255, adFldIsNullable
          .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable + adFldUpdatable
          .Fields.Append "Principal", adBoolean
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdDiagnosticos.DataSource = oRsProCatalogoDiagnosticos
    mo_Apariencia.ConfigurarFilasBiColores grdDiagnosticos, sighentidades.GrillaConFilasBicolor
End Sub

Sub CargaListaResultadosDeProcedimientos(oConexion As Connection)
    'Listados de resultados para la grilla de procedimientos
    Dim oRcs_Lista As New Recordset
    'Codigo de Tipos de Resultados
    Set oRcs_Lista = mo_reglasComunes.TiposResultadoServInterm(oConexion)
    oRcs_Lista.MoveFirst
    grdProcedimientos.ValueLists.Add ("TipoResultado")
    While Not oRcs_Lista.EOF
        grdProcedimientos.ValueLists("TipoResultado").ValueListItems.Add CInt(oRcs_Lista!IDRESULTADOSI), CStr(oRcs_Lista!Descripcion)
        oRcs_Lista.MoveNext
    Wend
    
    oRcs_Lista.Close
    Set oRcs_Lista = Nothing
End Sub

Sub CargaListaTiposDiagnosticos(oConexion As Connection)
    'Listados de resultados para la grilla de procedimientos
    Dim oRcs_Lista As New Recordset
      
    'Codigo de Tipos de Diagnosticos
    Set oRcs_Lista = mo_ReglasHIS.ListaTiposDiagnosticos
    oRcs_Lista.MoveFirst
    grdDiagnosticos.ValueLists.Add ("TipoDiagnostico")
    While Not oRcs_Lista.EOF
        grdDiagnosticos.ValueLists("TipoDiagnostico").ValueListItems.Add CInt(oRcs_Lista!IdSubclasificacionDx), CStr(oRcs_Lista!DescripcionLarga)
        oRcs_Lista.MoveNext
    Wend
    
    oRcs_Lista.Close
    Set oRcs_Lista = Nothing
End Sub

Sub CreaTemporalProcedimientos()
    'Crea recordset temporal de procedimientos
    If oRsProCatalogoProcedimientos.State = 1 Then
       Set oRsProCatalogoProcedimientos = Nothing
    End If
    With oRsProCatalogoProcedimientos
          .Fields.Append "IdDiagnostico", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdProducto", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "Procedimiento", adVarChar, 255, adFldIsNullable + adFldUpdatable
          .Fields.Append "CodigoCIE10", adVarChar, 255, adFldIsNullable + adFldUpdatable
          .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdResultado", adInteger, 0, adFldIsNullable
          .Fields.Append "Seleccionar", adBoolean
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdProcedimientos.DataSource = oRsProCatalogoProcedimientos
    mo_Apariencia.ConfigurarFilasBiColores grdProcedimientos, sighentidades.GrillaConFilasBicolor
End Sub

Sub CreaTemporalTratamientos()
    If oRsProCatalogoTratamientos.State = 1 Then
       Set oRsProCatalogoTratamientos = Nothing
    End If
    With oRsProCatalogoTratamientos
          .Fields.Append "IdProducto", adInteger
          .Fields.Append "Tratamientos", adVarChar, 255, adFldIsNullable
          .Fields.Append "Seleccionar", adBoolean
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdTratamientos.DataSource = oRsProCatalogoTratamientos
    mo_Apariencia.ConfigurarFilasBiColores grdTratamientos, sighentidades.GrillaConFilasBicolor
End Sub

Sub CreaTemporalProControlConfigDatos()
    If oRsProControlConfigDatos.State = 1 Then
       Set oRsProControlConfigDatos = Nothing
    End If
    With oRsProControlConfigDatos
          .Fields.Append "IdControlDato", adInteger
          .Fields.Append "Control_Texto", adVarChar, 255, adFldIsNullable
          .Fields.Append "Control_Tipo", adVarChar, 255, adFldIsNullable
          .Fields.Append "Control_Ancho", adInteger
          .Fields.Append "Control_EsDatoObligatorio", adBoolean
          .Fields.Append "Control_TextoToolTip", adVarChar, 255, adFldIsNullable
          .Fields.Append "Control_EsPresion", adBoolean
          .Fields.Append "Control_EsPeso", adBoolean
          .Fields.Append "Control_EsTalla", adBoolean
          .Fields.Append "Control_EsDatoCalculado", adBoolean
          .Fields.Append "Control_FormulaCalculaValor", adVarChar, 255, adFldIsNullable
          .Fields.Append "Control_EsDatoGrafico", adBoolean
          .Fields.Append "Control_EsGraficoEjeX", adBoolean
          .Fields.Append "Control_Fila", adInteger
          .Fields.Append "Control_Columna", adInteger
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

Sub CreaTemporaloRsProCabeceraConfigDatos()
    If oRsProCabeceraConfigDatos.State = 1 Then
       Set oRsProCabeceraConfigDatos = Nothing
    End If
    With oRsProCabeceraConfigDatos
          .Fields.Append "IdCabDato", adInteger
          .Fields.Append "Cab_Texto", adVarChar, 255, adFldIsNullable
          .Fields.Append "Cab_Tipo", adVarChar, 255, adFldIsNullable
          .Fields.Append "Cab_Ancho", adInteger
          .Fields.Append "Cab_EsDatoObligatorio", adInteger
          .Fields.Append "Cab_TextoToolTip", adVarChar, 255, adFldIsNullable
          .Fields.Append "Cab_EsDatoCalculado", adBoolean
          .Fields.Append "Cab_FormulaCalculaValor", adVarChar, 255, adFldIsNullable
          .Fields.Append "Cab_EsDatoCalculador", adBoolean
          .Fields.Append "Cab_RangoInicial", adInteger
          .Fields.Append "Cab_RangoFinal", adInteger
          .Fields.Append "Cab_Fila", adInteger
          .Fields.Append "Cab_Columna", adInteger
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

Sub LimpiaDiagnosticosFrecuentes(lbSoloDxCombo As Boolean)
    On Error GoTo errLimp
    Dim ml_IdDiagnostico As Long
    Dim lnFor As Integer, lnFor1 As Integer
    If lbSoloDxCombo Then
        For lnFor1 = 1 To 3
            For lnFor = 0 To cmbDiagnosticos.ListCount - 1
                ml_IdDiagnostico = Val(Mid(cmbDiagnosticos.ListItems.Item(lnFor).Key, 2, 100))
                With oRsProCatalogoDiagnosticos
                    If .RecordCount > 0 Then
                       .MoveFirst
                       Do While Not .EOF
                          If .Fields!idDiagnostico = ml_IdDiagnostico Then
                            .Delete
                            .Update
                          End If
                          .MoveNext
                       Loop
                    End If
                End With
            Next
        Next
    Else
        With oRsProCatalogoDiagnosticos
            If .RecordCount > 0 Then
               .MoveFirst
               Do While Not .EOF
                  .Delete
                  .Update
                  .MoveNext
               Loop
            End If
        End With
    End If

    With oRsProCatalogoProcedimientos
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              .Delete
              .Update
              .MoveNext
           Loop
        End If
    End With
errLimp:
End Sub

Private Sub btnAgregar_Click()
    Dim ml_IdDiagnostico As Long
    If mo_cmbDxPrograma.BoundText = "" Then Exit Sub
    ml_IdDiagnostico = mo_cmbDxPrograma.BoundText
    CargaDiagnosticosYProcedimientos ml_IdDiagnostico, ObtieneCodigoDescripcionDiagnostico(cmbDxPrograma.Text, "CODIGO"), ObtieneCodigoDescripcionDiagnostico(cmbDxPrograma.Text, "DESCRIPCION"), False, Right(Trim(cmbLabHisDx.Text), 3), 102
    cmbDxPrograma.ListIndex = -1
    txtFiltroDx.Text = ""
    oRsProCatalogoDiagnosticos.MoveFirst
End Sub

Private Sub btnAgregarProcedminiento_Click()
    Dim ml_IdCpt As Long
    If mo_cmbCptPrograma.BoundText = "" Then Exit Sub
    ml_IdCpt = mo_cmbCptPrograma.BoundText
    CargaProcedimientos ml_IdCpt, ObtieneCodigoDescripcionDiagnostico(cmbProcPrograma.Text, "CODIGO"), ObtieneCodigoDescripcionDiagnostico(cmbProcPrograma.Text, "DESCRIPCION"), 1, Space(3), False
    cmbProcPrograma.ListIndex = -1
    txtFiltroDx.Text = ""
End Sub

Private Sub btnAgregarInsumo_Click()
    Dim ml_IdTto As Long
    If mo_cmbTtoPrograma.BoundText = "" Then Exit Sub
    ml_IdTto = mo_cmbTtoPrograma.BoundText
    CargaTratamiento ml_IdTto, "", cmbTtoProgama.Text, False
    cmbTtoProgama.ListIndex = -1
    txtFiltroInsumos.Text = ""
End Sub

Private Sub btnBusquedaDiagnostico_Click()
    BusquedaDx ""
End Sub

Private Sub btnBuscarProc_Click()
'    Dim oBusqueda As New SIGHNegocios.BuscaCatalogoServiciosHosp
    Dim oBusqueda As New SIGHNegocios.BuscaServicio
    Dim oDoFactCatalogoServicio As DOCatalogoServicio

'    oBusqueda.idPuntoCarga = ml_IdPuntoCargaHosp '560
'    oBusqueda.idTipoFinanciamiento = ml_IdFormaPago '2
'    oBusqueda.TipoServicioOfrecido = 1
    oBusqueda.IdTipoCatalogo = 1
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDoFactCatalogoServicio = mo_reglasComunes.CatalogoServiciosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDoFactCatalogoServicio Is Nothing Then
'            If AgregarProcedimientosSeleccionado(oDoFactCatalogoServicio) = False Then
''                MsgBox "Procedimiento ya ha sido agregado", vbInformation, "Módulo Perinatal"
'            Else
'                cmbProcedimientosFrecuentes.Text = ""
'            End If
            Dim ml_IdCpt As Long
            ml_IdCpt = oDoFactCatalogoServicio.idProducto
            CargaProcedimientos oDoFactCatalogoServicio.idProducto, oDoFactCatalogoServicio.Codigo, oDoFactCatalogoServicio.nombre, 1, Space(3), False
            cmbProcPrograma.ListIndex = -1
            txtFiltroDx.Text = ""

        End If
    End If
    Set oBusqueda = Nothing
    
End Sub

Sub BusquedaDx(lcCodigoDx As String)
    Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
    Dim oDODiagnostico As DODiagnostico
    oBusqueda.SoloMuestraDxGalenHos = False
    oBusqueda.CodigoDx = lcCodigoDx
    oBusqueda.MostrarFormulario
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_reglasComunes.DiagnosticosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            CargaDiagnosticosYProcedimientos oDODiagnostico.idDiagnostico, Trim(oDODiagnostico.CodigoCiE10), Trim(oDODiagnostico.Descripcion), False, Space(3), 102
        End If
    End If
    Set oDODiagnostico = Nothing
    Set oBusqueda = Nothing
End Sub

Private Sub btnInicializarControles_Click()
    If MsgBox("¿Desea inializar los controles del programa?", vbOKCancel, "Mensaje") = vbOK Then
            Inicializar_Control_Programa
    End If
End Sub

Private Sub btnQuitaOtrosProcedimientos_Click()
    EliminarProcedimientosSeleccionado
End Sub

Private Sub btnQuitar_Click()
'    If MsgBox("¿Desea eliminar todos los diagnosticos seleccionados?", vbYesNo, "Mensaje") = vbYes Then
'        LimpiaDiagnosticosFrecuentes False
'        Dim lnFor As Integer, lnFor1 As Integer
'        On Error Resume Next
'        Do While True
'            For lnFor1 = 1 To 3
'                For lnFor = 0 To cmbDiagnosticos.ListCount - 1
'                    cmbDiagnosticos.SelectedItems(lnFor).Selected = False
'                Next
'            Next
'            If cmbDiagnosticos.SelectedItems.Count = 0 Then
'               Exit Do
'            End If
'        Loop
'    End If

    EliminarDiagnosticoSeleccionado
    
    
End Sub

Public Sub EliminarDiagnosticoSeleccionado()
    If MsgBox("¿Desea eliminar el diagnóstico seleccionado?", vbYesNo, "Eliminar diagnósticos") = vbYes Then
        If oRsProCatalogoDiagnosticos.RecordCount > 0 Then
            With oRsProCatalogoDiagnosticos
                If Not .EOF And Not .BOF Then
                   .Delete
                   .Update
                End If
            End With
            If oRsProCatalogoDiagnosticos.RecordCount > 0 Then oRsProCatalogoDiagnosticos.MoveFirst
'            CargarProcedimientosDiagnosticos
        End If
    End If
End Sub

Public Sub EliminarProcedimientosSeleccionado()
    If MsgBox("¿Desea eliminar el procedimiento seleccionado?", vbYesNo, "Eliminar Procedimientos") = vbYes Then
        If oRsProCatalogoProcedimientos.RecordCount > 0 Then
            With oRsProCatalogoProcedimientos
                If Not .EOF And Not .BOF Then
                   .Delete
                   .Update
                End If
            End With
            If oRsProCatalogoProcedimientos.RecordCount > 0 Then oRsProCatalogoProcedimientos.MoveFirst
'            CargarProcedimientosDiagnosticos
        End If
    End If
End Sub

Sub CargaDiagnosticosYProcedimientos(lnIdDiagnostico As Long, lcCodigoCiE10 As String, lcDescripcion As String, lbPrincipal As Boolean, lcLabHis As String, lnIdSubClasificacionDx As Long)
    If oRsProCatalogoDiagnosticos.RecordCount = 0 Then
        oRsProCatalogoDiagnosticos.AddNew
        oRsProCatalogoDiagnosticos.Fields!IdSubclasificacionDx = lnIdSubClasificacionDx
        oRsProCatalogoDiagnosticos.Fields!idDiagnostico = lnIdDiagnostico
        oRsProCatalogoDiagnosticos.Fields!CodigoCiE10 = lcCodigoCiE10
        oRsProCatalogoDiagnosticos.Fields!Descripcion = lcCodigoCiE10 + " - " + lcDescripcion
        oRsProCatalogoDiagnosticos.Fields!Principal = lbPrincipal
        oRsProCatalogoDiagnosticos.Fields!labConfHIS = lcLabHis
        oRsProCatalogoDiagnosticos.Update
    Else
        If ValidarDiagnosticosMaterno(lnIdDiagnostico, lcLabHis) = False Then
            Exit Sub
        End If
        oRsProCatalogoDiagnosticos.AddNew
        oRsProCatalogoDiagnosticos.Fields!IdSubclasificacionDx = lnIdSubClasificacionDx
        oRsProCatalogoDiagnosticos.Fields!idDiagnostico = lnIdDiagnostico
        oRsProCatalogoDiagnosticos.Fields!CodigoCiE10 = lcCodigoCiE10
        oRsProCatalogoDiagnosticos.Fields!Descripcion = lcCodigoCiE10 + " - " + lcDescripcion
        oRsProCatalogoDiagnosticos.Fields!Principal = lbPrincipal
        oRsProCatalogoDiagnosticos.Fields!labConfHIS = lcLabHis
        oRsProCatalogoDiagnosticos.Update
        
'        oRsProCatalogoDiagnosticos.MoveFirst
'        oRsProCatalogoDiagnosticos.Find "IdDiagnostico=" & lnIdDiagnostico
'        If Not oRsProCatalogoDiagnosticos.EOF Then
'            If IIf(IsNull(oRsProCatalogoDiagnosticos!labConfHIS) = True, Space(3), oRsProCatalogoDiagnosticos!labConfHIS) = lcLabHis Then
'                MsgBox "El diagnóstico con el mismo codigo lab ya fué registrado", vbInformation, "Diagnósticos"
'                oRsProCatalogoDiagnosticos.MoveFirst
'                Exit Sub
'            Else
'                oRsProCatalogoDiagnosticos.AddNew
'                oRsProCatalogoDiagnosticos.Fields!IdSubclasificacionDx = lnIdSubClasificacionDx
'                oRsProCatalogoDiagnosticos.Fields!IdDiagnostico = lnIdDiagnostico
'                oRsProCatalogoDiagnosticos.Fields!CodigoCiE10 = lcCodigoCiE10
'                oRsProCatalogoDiagnosticos.Fields!Descripcion = lcCodigoCiE10 + " - " + lcDescripcion
'                oRsProCatalogoDiagnosticos.Fields!Principal = lbPrincipal
'                oRsProCatalogoDiagnosticos.Fields!labConfHIS = lcLabHis
'                oRsProCatalogoDiagnosticos.Update
'            End If
'        Else
'            oRsProCatalogoDiagnosticos.AddNew
'            oRsProCatalogoDiagnosticos.Fields!IdSubclasificacionDx = lnIdSubClasificacionDx
'            oRsProCatalogoDiagnosticos.Fields!IdDiagnostico = lnIdDiagnostico
'            oRsProCatalogoDiagnosticos.Fields!CodigoCiE10 = lcCodigoCiE10
'            oRsProCatalogoDiagnosticos.Fields!Descripcion = lcCodigoCiE10 + " - " + lcDescripcion
'            oRsProCatalogoDiagnosticos.Fields!Principal = lbPrincipal
'            oRsProCatalogoDiagnosticos.Fields!labConfHIS = lcLabHis
'            oRsProCatalogoDiagnosticos.Update
'        End If
    End If
End Sub

Sub CargaProcedimientos(lnIdDxProc As Long, lcCodigoCPT As String, lcDescripcion As String, lnIdResultado As Integer, lcLabConfHIS As String, lbSeleccionar As Boolean)
    If oRsProCatalogoProcedimientos.RecordCount = 0 Then
        oRsProCatalogoProcedimientos.AddNew
        oRsProCatalogoProcedimientos.Fields!idDiagnostico = 1
        oRsProCatalogoProcedimientos.Fields!idProducto = lnIdDxProc
        oRsProCatalogoProcedimientos.Fields!procedimiento = lcCodigoCPT & " = " & lcDescripcion
        oRsProCatalogoProcedimientos.Fields!CodigoCiE10 = ""
        oRsProCatalogoProcedimientos.Fields!labConfHIS = lcLabConfHIS
        oRsProCatalogoProcedimientos.Fields!IDRESULTADO = lnIdResultado
        oRsProCatalogoProcedimientos.Fields!seleccionar = lbSeleccionar
        oRsProCatalogoProcedimientos.Update
    Else
        oRsProCatalogoProcedimientos.MoveFirst
        oRsProCatalogoProcedimientos.Find "idProducto=" & lnIdDxProc
        If Not oRsProCatalogoProcedimientos.EOF Then
            MsgBox "El CPT ya fue ingresado", vbInformation, "Procedimientos"
            oRsProCatalogoProcedimientos.MoveFirst
            Exit Sub
        Else
            oRsProCatalogoProcedimientos.AddNew
            oRsProCatalogoProcedimientos.Fields!idDiagnostico = 1
            oRsProCatalogoProcedimientos.Fields!idProducto = lnIdDxProc
            oRsProCatalogoProcedimientos.Fields!procedimiento = lcCodigoCPT & " = " & lcDescripcion
            oRsProCatalogoProcedimientos.Fields!CodigoCiE10 = ""
            oRsProCatalogoProcedimientos.Fields!labConfHIS = lcLabConfHIS
            oRsProCatalogoProcedimientos.Fields!IDRESULTADO = lnIdResultado
            oRsProCatalogoProcedimientos.Fields!seleccionar = lbSeleccionar
            oRsProCatalogoProcedimientos.Update
        End If
    End If

End Sub

Sub CargaTratamiento(lnIdTtoPro As Long, lcCodigoTto As String, lcDescripcion As String, lbSeleccionar As Boolean)
    If oRsProCatalogoTratamientos.RecordCount = 0 Then
        oRsProCatalogoTratamientos.AddNew
        oRsProCatalogoTratamientos.Fields!idProducto = lnIdTtoPro
        oRsProCatalogoTratamientos.Fields!Tratamientos = lcDescripcion
        oRsProCatalogoTratamientos.Fields!seleccionar = lbSeleccionar
        oRsProCatalogoTratamientos.Update
    Else
        oRsProCatalogoTratamientos.MoveFirst
        oRsProCatalogoTratamientos.Find "idProducto=" & lnIdTtoPro
        If Not oRsProCatalogoTratamientos.EOF Then
            MsgBox "El Tratamiento ya fue ingresado", vbInformation, "Insumos"
            oRsProCatalogoTratamientos.MoveFirst
            Exit Sub
        Else
            oRsProCatalogoTratamientos.AddNew
            oRsProCatalogoTratamientos.Fields!idProducto = lnIdTtoPro
            oRsProCatalogoTratamientos.Fields!Tratamientos = lcDescripcion
            oRsProCatalogoTratamientos.Fields!seleccionar = lbSeleccionar
            oRsProCatalogoTratamientos.Update
        End If
    End If
End Sub


Public Sub CargarProcedimientosDiagnosticos()
    'Limpiar Procedimientos
    With oRsProCatalogoProcedimientos
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              .Delete
              .Update
              .MoveNext
           Loop
        End If
    End With
    
    'Cargar Procedimientos de los diagnosticos
    Dim lbExisteProcedimiento As Boolean
    Dim oRsTmp1 As New Recordset
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    If oRsProCatalogoDiagnosticos.RecordCount > 0 Then
        oRsProCatalogoDiagnosticos.MoveFirst
        Do While Not oRsProCatalogoDiagnosticos.EOF
            Set oRsTmp1 = mo_reglasComunes.ProCatalogoProcedimientosSelecionarPorIdProgramaIdDiagnostico(ml_IdPrograma, oRsProCatalogoDiagnosticos.Fields!idDiagnostico, oConexion)
              If oRsTmp1.RecordCount > 0 Then
                 oRsTmp1.MoveFirst
                 Do While Not oRsTmp1.EOF
                    lbExisteProcedimiento = False
                    If oRsProCatalogoProcedimientos.RecordCount > 0 Then
                        oRsProCatalogoProcedimientos.MoveFirst
                        Do While Not oRsProCatalogoProcedimientos.EOF
                            If oRsProCatalogoProcedimientos.Fields!idProducto = oRsTmp1.Fields!idProducto Then
                                lbExisteProcedimiento = True
                                Exit Do
                            End If
                            oRsProCatalogoProcedimientos.MoveNext
                        Loop
                    End If
                    If lbExisteProcedimiento = False Then
                        oRsProCatalogoProcedimientos.AddNew
                        oRsProCatalogoProcedimientos.Fields!idDiagnostico = oRsTmp1.Fields!idDiagnostico
                        oRsProCatalogoProcedimientos.Fields!idProducto = oRsTmp1.Fields!idProducto
                        oRsProCatalogoProcedimientos.Fields!procedimiento = oRsTmp1.Fields!nombre
                        oRsProCatalogoProcedimientos.Fields!CodigoCiE10 = oRsTmp1.Fields!CodigoCiE10
                        oRsProCatalogoProcedimientos.Fields!CodigoCiE10 = ""
                        oRsProCatalogoProcedimientos.Fields!labConfHIS = Space(3)
                        oRsProCatalogoProcedimientos.Fields!IDRESULTADO = 1
                        oRsProCatalogoProcedimientos.Fields!seleccionar = False
                        oRsProCatalogoProcedimientos.Update
                    End If
                    oRsTmp1.MoveNext
                 Loop
                 oRsProCatalogoProcedimientos.MoveFirst
            End If
            oRsTmp1.Close
            Set oRsTmp1 = Nothing
            oRsProCatalogoDiagnosticos.MoveNext
        Loop
    End If

    oConexion.Close
    Set oConexion = Nothing
End Sub


Public Function ObtieneCodigoDescripcionDiagnostico(ByVal mcDiagnosticoLargo As String, ByVal mcTipo As String) As String
    Dim TestPos As Integer
    ObtieneCodigoDescripcionDiagnostico = ""
    TestPos = InStr(1, mcDiagnosticoLargo, "=")
    If TestPos <> 0 Then
        Select Case mcTipo
        Case "CODIGO"
           ObtieneCodigoDescripcionDiagnostico = Mid(mcDiagnosticoLargo, 1, TestPos - 1)
        Case "DESCRIPCION"
           ObtieneCodigoDescripcionDiagnostico = Mid(mcDiagnosticoLargo, TestPos + 2, Len(mcDiagnosticoLargo) - TestPos - 1)
        End Select
    End If
End Function

Public Sub EliminarTratamientoSeleccionado()
    If MsgBox("¿Desea eliminar el insumo seleccionado?", vbYesNo, "Eliminar Insumod") = vbYes Then
        If oRsProCatalogoTratamientos.RecordCount > 0 Then
            With oRsProCatalogoTratamientos
                If Not .EOF And Not .BOF Then
                   .Delete
                   .Update
                End If
            End With
            If oRsProCatalogoTratamientos.RecordCount > 0 Then oRsProCatalogoTratamientos.MoveFirst
'            CargarProcedimientosDiagnosticos
        End If
    End If
End Sub

Private Sub btnQuitarInsumo_Click()
    EliminarTratamientoSeleccionado
End Sub

Private Sub chkDatoCabecera_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDatoCabecera(Index)
End Sub

Private Sub cmbLabHisCpt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then btnAgregarProcedminiento.SetFocus
End Sub

Private Sub cmbLabHisDx_keyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then btnAgregar.SetFocus
End Sub

Private Sub cmbProcPrograma_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmbLabHisCpt.SetFocus
End Sub

Private Sub grdDiagnosticos_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
    If ml_IdControlSeleccionado = ml_IdControl Then
        EliminarDiagnosticoSeleccionado
    End If
End Sub

Private Sub grdDiagnosticos_Error(ByVal ErrorInfo As UltraGrid.SSError)
    If ErrorInfo.Code = 16389 And ErrorInfo.DataError.Cell.Column.Key = "IdSubClasificacionDX" Then
        ErrorInfo.DisplayErrorDialog = False
        MsgBox "El valor del tipo de diagnóstico no puede ser vacio", vbInformation, "Validación"
    End If
End Sub

Private Sub grdDiagnosticos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdDiagnosticos
End Sub

Private Sub grdHistoricoControles_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdHistoricoControles_Click()
'Actualizado 27102014
'    If grdHistoricoControles.Selected.Rows.Count > 0 Then
'        ml_IdControlSeleccionado = oRsProCatalogoControles.Fields!idcontrol
''        ml_IdControlSeleccionado = grdHistoricoControles.Selected.Rows(0).Cells("idcontrol").Value
''        CargaControlSeleccionado
'    End If
End Sub

Private Sub mnuAgregarControl_Click()
    If grdHistoricoControles.Selected.Rows.Count > 0 Then
        If oRsProCatalogoControles.Fields!ControlOtroEESS = False Then
            If oRsProCatalogoControles.Fields!IdControl < ml_IdControl Then
                If oRsProCatalogoControles.Fields!FechaControl = "" Then
                    Dim mo_ProEstablecimientos As New frmProEstablecimientos
                    mo_ProEstablecimientos.Opcion = sghAgregar
                    mo_ProEstablecimientos.IdPrograma = ml_IdPrograma
                    mo_ProEstablecimientos.IdProCabecera = ml_IdProCabecera
                    mo_ProEstablecimientos.IdControl = oRsProCatalogoControles.Fields!IdControl
                    mo_ProEstablecimientos.Show 1
                    If mo_ProEstablecimientos.BotonPresionado = sghAceptar Then
                        If CDate(mo_ProEstablecimientos.FechaControl) > CDate(txtFechaControl.Text) Then
                            MsgBox "La fecha registrada no puede ser mayor o igual que la fecha de control actual", vbInformation, "Agregar control"
                            Set mo_ProEstablecimientos = Nothing
                            Exit Sub
                        End If
                        If txtDatoCabecera(1).Text <> sighentidades.FECHA_VACIA_DMY Then
                            If CDate(mo_ProEstablecimientos.FechaControl) < CDate(txtDatoCabecera(1).Text) Then
                                MsgBox "La fecha registrada no puede ser menor que la fecha FUM", vbInformation, "Agregar control"
                                Set mo_ProEstablecimientos = Nothing
                                Exit Sub
                            End If
                        End If
                        MsgBox "Los datos fuerón agregados satisfactoriamente.", vbInformation, "Agregar control"
                        oRsProCatalogoControles.Fields!ControlOtroEESS = True
                        oRsProCatalogoControles.Fields!IdEstablecimiento = mo_ProEstablecimientos.IdEstablecimiento
                        oRsProCatalogoControles.Fields!FechaControl = mo_ProEstablecimientos.FechaControl
                    End If
                    Set mo_ProEstablecimientos = Nothing
                    Exit Sub
                End If
            End If
        End If
    Else
        MsgBox "Seleccione el número de control", vbInformation, "Agregar control"
    End If
End Sub

Private Sub mnuModificarControl_Click()
    If grdHistoricoControles.Selected.Rows.Count > 0 Then
        If oRsProCatalogoControles.Fields!ControlOtroEESS = True Then
            Dim mo_ProEstablecimientos As New frmProEstablecimientos
            mo_ProEstablecimientos.Opcion = sghModificar
            mo_ProEstablecimientos.IdEstablecimiento = oRsProCatalogoControles.Fields!IdEstablecimiento
            mo_ProEstablecimientos.FechaControl = oRsProCatalogoControles.Fields!FechaControl
            mo_ProEstablecimientos.IdControl = oRsProCatalogoControles.Fields!IdControl
            mo_ProEstablecimientos.Show 1
            If mo_ProEstablecimientos.BotonPresionado = sghAceptar Then
                If CDate(mo_ProEstablecimientos.FechaControl) >= CDate(txtFechaControl.Text) Then
                    MsgBox "La fecha registrada no puede ser mayor o igual que la fecha de control actual", vbInformation, "Agregar control"
                    Set mo_ProEstablecimientos = Nothing
                    Exit Sub
                End If
                If txtDatoCabecera(1).Text <> sighentidades.FECHA_VACIA_DMY Then
                    If CDate(mo_ProEstablecimientos.FechaControl) < CDate(txtDatoCabecera(1).Text) Then
                        MsgBox "La fecha registrada no puede ser menor que la fecha FUM", vbInformation, "Agregar control"
                        Set mo_ProEstablecimientos = Nothing
                        Exit Sub
                    End If
                End If
                MsgBox "Los datos fuerón modificados satisfactoriamente.", vbInformation, "Agregar control"
                oRsProCatalogoControles.Fields!ControlOtroEESS = True
                oRsProCatalogoControles.Fields!IdEstablecimiento = mo_ProEstablecimientos.IdEstablecimiento
                oRsProCatalogoControles.Fields!FechaControl = mo_ProEstablecimientos.FechaControl
            End If
            Set mo_ProEstablecimientos = Nothing
            Exit Sub
        End If
    Else
        MsgBox "Seleccione el número de control", vbInformation, "Modificar control"
    End If
End Sub

Private Sub mnuEliminarControl_Click()
    If grdHistoricoControles.Selected.Rows.Count > 0 Then
        If oRsProCatalogoControles.Fields!ControlOtroEESS = True Then
            Dim mo_ProEstablecimientos As New frmProEstablecimientos
            mo_ProEstablecimientos.Opcion = sghEliminar
            mo_ProEstablecimientos.IdEstablecimiento = oRsProCatalogoControles.Fields!IdEstablecimiento
            mo_ProEstablecimientos.FechaControl = oRsProCatalogoControles.Fields!FechaControl
            mo_ProEstablecimientos.IdControl = oRsProCatalogoControles.Fields!IdControl
            mo_ProEstablecimientos.Show 1
            If mo_ProEstablecimientos.BotonPresionado = sghAceptar Then
                oRsProCatalogoControles.Fields!ControlOtroEESS = False
                oRsProCatalogoControles.Fields!IdEstablecimiento = 0
                oRsProCatalogoControles.Fields!FechaControl = ""
            End If
            Set mo_ProEstablecimientos = Nothing
            Exit Sub
        End If
    Else
        MsgBox "Seleccione el número de control", vbInformation, "Modificar control"
    End If
End Sub

Private Sub grdHistoricoControles_DblClick()
    If grdHistoricoControles.Selected.Rows.Count > 0 Then
        If oRsProCatalogoControles.Fields!ControlOtroEESS = False Then
            ml_IdControlSeleccionado = oRsProCatalogoControles.Fields!IdControl
            If oRsProCatalogoControles.Fields!FechaControl <> "" Then
'                If ml_IdControlSeleccionado >= ml_IdUltimoControl Then
'                    If mb_ControlNuevo = True Then ml_IdControl = ml_IdControlSeleccionado
'                End If
                CargaControlSeleccionado
            Else
                If ml_IdControlSeleccionado >= ml_IdUltimoControl Then
                    If mb_ControlNuevo = True Then ml_IdControl = ml_IdControlSeleccionado
                    CargaControlSeleccionado
                End If
            End If
        Else
            'MsgBox "CONSULTA PARA INGRESO DE EESS"
            mnuModificarControl_Click
        End If
    End If
End Sub

Private Sub grdHistoricoControles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuControlesEESS
    End If
End Sub

Private Sub grdHistoricoControles_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdHistoricoControles
End Sub

Private Sub grdProcedimientos_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
    If ml_IdControlSeleccionado = ml_IdControl Then
        EliminarProcedimientosSeleccionado
    End If
End Sub

Private Sub grdProcedimientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdProcedimientos
End Sub

Private Sub grdTratamientos_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
    If ml_IdControlSeleccionado = ml_IdControl Then
        EliminarTratamientoSeleccionado
    End If
End Sub

Private Sub grdTratamientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdTratamientos
End Sub

Public Sub Inicializar_Control_Programa()
    ml_IdProCabecera = 0
    ml_IdControl = 1
    ml_IdUltimoControl = 1
    ml_IdControlSeleccionado = 1
    mb_ConsultaControl = False
    'Limpia Datos de cabecera
'    If oRsProCabeceraConfigDatos.RecordCount > 0 Then
'        oRsProCabeceraConfigDatos.MoveFirst
'        Do While Not oRsProCabeceraConfigDatos.EOF
'           txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Text = txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Tag
'           oRsProCabeceraConfigDatos.MoveNext
'        Loop
'    End If
    
    If oRsProCabeceraConfigDatos.RecordCount > 0 Then
        oRsProCabeceraConfigDatos.MoveFirst
        Do While Not oRsProCabeceraConfigDatos.EOF
           Select Case oRsProCabeceraConfigDatos.Fields!cab_tipo
           Case "ValorEntero", "ValorTexto", "ValorFecha", "ValorDouble"
                txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Text = txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Tag
           Case "ValorCheck"
                chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Value = 0
           End Select
           oRsProCabeceraConfigDatos.MoveNext
        Loop
    End If
             
    CargarDatosDefectoControl
    CargaGraficoChartSpace
    
    If oRsProCatalogoControles.RecordCount > 0 Then
        oRsProCatalogoControles.MoveFirst
        Do While Not oRsProCatalogoControles.EOF
           If oRsProCatalogoControles.Fields!IdControl = ml_IdControl Then
'               oRsProCatalogoControles.Fields!FechaControl = "Control Actual"
                oRsProCatalogoControles.Fields!FechaControl = ""
           Else
                oRsProCatalogoControles.Fields!FechaControl = ""
           End If
           oRsProCatalogoControles.Update
           oRsProCatalogoControles.MoveNext
        Loop
        oRsProCatalogoControles.MoveFirst
    End If
End Sub

Public Sub CargaDatosAcontroles(ml_peso_triaje As Double, ml_presion_triaje As String, ml_Talla_triaje As Long, ml_lnEdadEnDias As Long, oConexion As Connection)
    'Consulta si tiene controles en el programa
    Dim oRsTmp1 As New Recordset
    Dim mc_CompletarCeroEntero As String
    Dim i As Integer
    ml_IdProCabecera = 0
    ml_IdControl = 1
    ml_IdControlSeleccionado = 1
    ml_IdUltimoControl = 1
    mb_ConsultaControl = False
    ml_peso = ml_peso_triaje
    ml_presion = ml_presion_triaje
    ml_talla = ml_Talla_triaje
    ml_EdadEnDias = ml_lnEdadEnDias
    ml_SeteoUnaSolaVesCheckBox = False
    TabPrograma.Tab = 0

    Set oRsTmp1 = mo_reglasComunes.ProConsultarCabeceraControlPorIdPacienteIdAtencion(ml_IdPrograma, ml_IdPaciente, ml_idAtencion, oConexion)
    If oRsTmp1.RecordCount > 0 Then
        ml_IdProCabecera = oRsTmp1.Fields!IdProCabecera
        ml_IdControl = oRsTmp1.Fields!IdControl
        ml_IdUltimoControl = ml_IdControl
        ml_IdControlSeleccionado = ml_IdControl
        mb_ConsultaControl = True
    Else
        If oRsTmp1.State = 1 Then oRsTmp1.Close
        Set oRsTmp1 = mo_reglasComunes.ProConsultarActualProgramaPaciente(ml_IdPrograma, ml_IdPaciente, oConexion)
        If oRsTmp1.RecordCount > 0 Then
            ml_IdProCabecera = oRsTmp1.Fields!IdProCabecera
        End If
        txtFechaControl.Text = ml_FechaAtencion
        mo_Formulario.HabilitarDeshabilitar txtFechaControl, False
    End If
    
    'Limpia Datos de cabecera
    If oRsProCabeceraConfigDatos.RecordCount > 0 Then
        oRsProCabeceraConfigDatos.MoveFirst
        Do While Not oRsProCabeceraConfigDatos.EOF
           Select Case oRsProCabeceraConfigDatos.Fields!cab_tipo
           Case "ValorEntero", "ValorTexto", "ValorFecha", "ValorDouble"
                txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Text = txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Tag
           Case "ValorCheck"
                chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Value = 0
           End Select
           oRsProCabeceraConfigDatos.MoveNext
        Loop
    End If
    
    'Limpia Datos de Control
    If oRsProCatalogoControles.RecordCount > 0 Then
        oRsProCatalogoControles.MoveFirst
        Do While Not oRsProCatalogoControles.EOF
           oRsProCatalogoControles.Fields!FechaControl = ""
           oRsProCatalogoControles.Fields!IdEstablecimiento = 0
           oRsProCatalogoControles.Fields!ControlOtroEESS = False
           oRsProCatalogoControles.Update
           oRsProCatalogoControles.MoveNext
        Loop
        oRsProCatalogoControles.MoveFirst
    End If
    
    If ml_IdProCabecera <> 0 Then
        'Consultar datos de cabecera
        If oRsTmp1.State = 1 Then oRsTmp1.Close
        Set oRsTmp1 = mo_reglasComunes.ProConsultarDatosCabecera(ml_IdPrograma, ml_IdProCabecera, oConexion)
        'Cargar Datos de cabecera del programa
        If oRsTmp1.RecordCount > 0 Then
            oRsTmp1.MoveFirst
            Do While Not oRsTmp1.EOF
               If oRsTmp1.Fields!CabDato <> "" Then
                    Select Case oRsTmp1.Fields!cab_tipo
                    Case "ValorEntero", "ValorTexto", "ValorFecha", "ValorDouble"
                        If ml_IdPrograma = 1 Then
                            If oRsTmp1.Fields!IdCabDato = 13 And Len(oRsTmp1.Fields!CabDato) < 6 Then
                                txtDatoCabecera(oRsTmp1.Fields!IdCabDato).Text = txtDatoCabecera(oRsTmp1.Fields!IdCabDato).Tag
                            Else
                                txtDatoCabecera(oRsTmp1.Fields!IdCabDato).Text = oRsTmp1.Fields!CabDato
                            End If
                        Else
                            txtDatoCabecera(oRsTmp1.Fields!IdCabDato).Text = oRsTmp1.Fields!CabDato
                        End If

'                        txtDatoCabecera(oRsTmp1.Fields!IdCabDato).Text = oRsTmp1.Fields!CabDato
                    Case "ValorCheck"
                        ml_SeteoUnaSolaVesCheckBox = True
                        chkDatoCabecera(oRsTmp1.Fields!IdCabDato).Value = Val(oRsTmp1.Fields!CabDato)
                        ml_SeteoUnaSolaVesCheckBox = False
                    End Select
               End If
               oRsTmp1.MoveNext
            Loop
         End If
        'Consultar Historial de Controles
        If oRsTmp1.State = 1 Then oRsTmp1.Close
        Set oRsTmp1 = mo_reglasComunes.ProConsultarControles(ml_IdPrograma, ml_IdProCabecera, oConexion)
        'Cargar Historial de Controles
        If mb_ConsultaControl = False Then
           If ml_IdControl = 1 Then
                If oRsTmp1.RecordCount > 0 Then
                    oRsTmp1.MoveLast
                    ml_IdControl = oRsTmp1.Fields!IdControl + 1 'Control Actual
                    ml_IdUltimoControl = ml_IdControl
                Else
                    ml_IdControl = 1 'Control Actual
                    ml_IdUltimoControl = 1
                End If
                'ml_IdControl = oRsTmp1.RecordCount + 1 'Control Actual
           End If
           'Cuando el control llego al limite se inializa los controles
           If ml_IdControl > oRsProCatalogoControles.RecordCount Then
                Inicializar_Control_Programa
                Exit Sub
           End If
        End If
        ml_IdControlSeleccionado = ml_IdControl 'Control Seleccionado Actual
        If oRsProCatalogoControles.RecordCount > 0 Then
              If oRsTmp1.RecordCount > 0 Then
                 oRsTmp1.MoveFirst
                 Do While Not oRsTmp1.EOF
                    oRsProCatalogoControles.MoveFirst
                    oRsProCatalogoControles.Find "IdControl=" & oRsTmp1.Fields!IdControl
                    If Not oRsProCatalogoControles.EOF Then
                       oRsProCatalogoControles.Fields!FechaControl = oRsTmp1.Fields!FechaControl
                       oRsProCatalogoControles.Fields!ControlOtroEESS = IIf(IsNull(oRsTmp1.Fields!ControlOtroEESS), False, oRsTmp1.Fields!ControlOtroEESS)
                       oRsProCatalogoControles.Fields!IdEstablecimiento = oRsTmp1.Fields!IdEstablecimiento
                       oRsProCatalogoControles.Update
                    End If
                    oRsTmp1.MoveNext
                 Loop
              End If
              oRsTmp1.Close
         End If
         If mb_ConsultaControl Then
            ml_IdControlSeleccionado = ml_IdControl
            CargaControlSeleccionado
         End If
    End If
    If mb_ConsultaControl = False Then
        btnInicializarControles.Enabled = True
        If ml_IdControl = ml_IdControlSeleccionado Then
            CargarDatosDefectoControl
            CargaGraficoChartSpace
            'Frank 03092014
            If ml_IdPrograma = 1 Then 'Solo para el Programa Materno
                If ml_IdControl = 1 Then
                    Cargar_DiagnosticoPorDefecto_PrimerControl
                End If
            End If
        End If
    End If
    oRsProCatalogoControles.MoveFirst
    oRsProCatalogoControles.Find "IdControl=" & ml_IdControl
    If Not oRsProCatalogoControles.EOF Then
'        oRsProCatalogoControles.Fields!FechaControl = "Control Actual"
'        oRsProCatalogoControles.Fields!FechaControl = ""
        oRsProCatalogoControles.Update
        FrameControl.Caption = "Control Nº " + CStr(ml_IdControl) + " (Control Actual)"
    End If
    Set oRsTmp1 = Nothing
End Sub

'Frank 02092014
Public Sub Cargar_DiagnosticoPorDefecto_PrimerControl()
    'ml_EdadEnDias
    'Cargar Diagnostico Embarazo Confirmado
    CargarDiagnosticoPorDefectoPorIdDx 16061
    If ml_EdadEnDias > 35 Then 'Cargar Diagnostico Gestante con Riesgo
        CargarDiagnosticoPorDefectoPorIdDx 16071
    End If
End Sub

'Frank 02092014
Public Sub CargarDiagnosticoPorDefectoPorIdDx(ByVal mlIdDiagnostico As Long)
    Dim lnFor1 As Integer
    Dim lnFor As Integer
    Dim ml_IdDiagnostico As Long
    Dim CBLI As SSCBListItem

    For lnFor1 = 1 To 3
        For lnFor = 0 To cmbDiagnosticos.ListCount - 1
            If Val(Mid(cmbDiagnosticos.ListItems.Item(lnFor).Key, 2, 100)) = mlIdDiagnostico Then
                cmbDiagnosticos.ListItems.Item(lnFor).Selected = True
            End If
        Next
    Next
    
    LimpiaDiagnosticosFrecuentes True
    If cmbDiagnosticos.SelectedItems.Count > 0 Then
        For Each CBLI In cmbDiagnosticos.SelectedItems
            mlIdDiagnostico = Val(Mid(CBLI.Key, 2, 100))
            CargaDiagnosticosYProcedimientos mlIdDiagnostico, ObtieneCodigoDescripcionDiagnostico(CBLI.Text, "CODIGO"), ObtieneCodigoDescripcionDiagnostico(CBLI.Text, "DESCRIPCION"), False, Space(3), 102
        Next CBLI
        On Error Resume Next
        oRsProCatalogoDiagnosticos.MoveFirst
    End If
End Sub


Public Sub CargarDatosDefectoControl()
    LimpiarControl
    'Ingresa datos por defecto del control
    oRsProControlConfigDatos.MoveFirst
    txtFechaControl.Text = ml_FechaAtencion
    Do While Not oRsProControlConfigDatos.EOF
        If oRsProControlConfigDatos.Fields!Control_Tipo <> "ValorFecha" Then txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Text = ""
        If oRsProControlConfigDatos.Fields!Control_EsPeso Then txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Text = ml_peso
        If oRsProControlConfigDatos.Fields!Control_EsPresion Then txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Text = ml_presion
        If oRsProControlConfigDatos.Fields!Control_EsTalla Then txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Text = ml_talla
        If oRsProControlConfigDatos.Fields!Control_EsDatoCalculado Then
            ActualizaDatosControlCalculados oRsProControlConfigDatos.Fields!idcontroldato, oRsProControlConfigDatos.Fields!Control_FormulaCalculaValor
        End If
        oRsProControlConfigDatos.MoveNext
    Loop
End Sub

'Actuaizar esta región para las nuevas formulas de calculo en el control
Public Function ActualizaDatosControlCalculados(ml_IdControlDato As Long, TextoFormulaCalcularValor As String) As String
    ActualizaDatosControlCalculados = ""
    If ml_IdControl = ml_IdControlSeleccionado Then
        Select Case TextoFormulaCalcularValor
           Case "ProMat_DevuelveEdadGestacional"
                If txtDatoCabecera(1).Text <> sighentidades.FECHA_VACIA_DMY Then
                    If sighentidades.EsFecha(txtDatoCabecera(1).Text, "DD/MM/AAAA", False) Then
                        If sighentidades.EsFecha(txtFechaControl.Text, "DD/MM/AAAA", False) Then
                            If CDate(txtDatoCabecera(1).Text) > CDate(txtFechaControl.Text) Then
                                ActualizaDatosControlCalculados = "La Fecha '" & lblDatoCabecera(1).Caption & "' no debe ser mayor que la fecha de control"
                                txtDatoCabecera(1).SetFocus
                                Exit Function
                            Else
                                txtDatoControl(ml_IdControlDato).Text = DevuelveEdadGestacional(CDate(txtDatoCabecera(1).Text), CDate(txtFechaControl.Text))
                            End If
                        End If
                    End If
                End If
           Case "ProMat_CalculaPercentilIMC"
                If txtDatoControl(4).Text <> "" Then
                    txtDatoControl(ml_IdControlDato).Text = CalculaPercentilIMC(ml_peso, ml_talla, CLng(txtDatoControl(4).Text))
                End If
           Case "Programa_Formula2" 'Agregar las funciones en ProControlConfigDatos.Control_FormulaCalculaValor
           Case "Programa_Formula3" 'Agregar las funciones en ProControlConfigDatos.Control_FormulaCalculaValor
           Case "Programa_Formula4" 'Agregar las funciones en ProControlConfigDatos.Control_FormulaCalculaValor ... etc
        End Select
    End If
End Function

Sub LimpiarControl()
    Dim lnFor As Integer, lnFor1 As Integer
    'Limpiar datos del control
    txtFechaControl.Text = txtFechaControl.Tag
    
    If oRsProControlConfigDatos.RecordCount > 0 Then
       oRsProControlConfigDatos.MoveFirst
       Do While Not oRsProControlConfigDatos.EOF
          txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Text = txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Tag
          oRsProControlConfigDatos.MoveNext
       Loop
    End If
    
    'Limpiar diagnosticos y procedimientos
    cmbDiagnosticos.Enabled = True
    mo_Formulario.HabilitarDeshabilitar cmbDiagnosticos, True
    mo_Formulario.HabilitarDeshabilitar btnBusquedaDiagnostico, True
    mo_Formulario.HabilitarDeshabilitar btnQuitar, True
    mo_Formulario.HabilitarDeshabilitar cmbDxPrograma, True
    mo_Formulario.HabilitarDeshabilitar txtFiltroDx, True
    mo_Formulario.HabilitarDeshabilitar btnAgregar, True
    mo_Formulario.HabilitarDeshabilitar btnBusquedaDiagnostico, True
    
    mo_Formulario.HabilitarDeshabilitar txtFiltraCPT, True
    mo_Formulario.HabilitarDeshabilitar cmbProcPrograma, True
    mo_Formulario.HabilitarDeshabilitar btnAgregarProcedminiento, True
    mo_Formulario.HabilitarDeshabilitar btnQuitaOtrosProcedimientos, True
    mo_Formulario.HabilitarDeshabilitar btnBuscarProc, True
    
    mo_Formulario.HabilitarDeshabilitar txtFiltraCPT, True
    mo_Formulario.HabilitarDeshabilitar cmbProcPrograma, True
    mo_Formulario.HabilitarDeshabilitar btnAgregarProcedminiento, True
    mo_Formulario.HabilitarDeshabilitar btnQuitaOtrosProcedimientos, True
    mo_Formulario.HabilitarDeshabilitar btnBuscarProc, True
'    LimpiaDiagnosticosFrecuentes False
    
    LimpiaDiagnosticosFrecuentes False
    On Error Resume Next
    Do While True
        For lnFor1 = 1 To 3
            For lnFor = 0 To cmbDiagnosticos.ListCount - 1
                cmbDiagnosticos.SelectedItems(lnFor).Selected = False
            Next
        Next
        If cmbDiagnosticos.SelectedItems.Count = 0 Then
            Exit Do
        End If
    Loop
    
    'Limpiar Tratamientos
    If oRsProCatalogoTratamientos.RecordCount > 0 Then
       oRsProCatalogoTratamientos.MoveFirst
       Do While Not oRsProCatalogoTratamientos.EOF
            oRsProCatalogoTratamientos.Fields!seleccionar = False
            oRsProCatalogoTratamientos.MoveNext
       Loop
       oRsProCatalogoTratamientos.MoveFirst
    End If
End Sub

Sub CargaControlSeleccionado()
    Dim oRsTmp1 As New ADODB.Recordset
    Dim oRsTmp2 As New ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim mc_CompletarCeroEntero As String
    Dim lnFor As Integer, lnFor1 As Integer
    Dim i As Integer
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion

    LimpiarControl
    
    'Valida que control seleccionado no sea mayor que el control actual
    If ml_IdControlSeleccionado > ml_IdControl Then
        If mb_ControlNuevo = False Then
            MsgBox "No puede elegir un control mayor al control actual", vbInformation, "Historico de controles"
            ml_IdControlSeleccionado = ml_IdControl
            oRsProCatalogoControles.MoveFirst
            oRsProCatalogoControles.Find "IdControl=" & ml_IdControl
            If Not oRsProCatalogoControles.EOF Then
'               oRsProCatalogoControles.Fields!FechaControl = "Control Actual"
'               oRsProCatalogoControles.Fields!FechaControl = ""
               oRsProCatalogoControles.Update
               FrameControl.Caption = "Control Nº " + CStr(ml_IdControl) + " (Control Actual)"
            End If
        End If
    Else
'        If mb_ControlNuevo = True Then
            If ml_IdControlSeleccionado = ml_IdControl Then
                FrameControl.Caption = "Control Nº " + CStr(ml_IdControlSeleccionado) + " (Control Actual)"
            Else
                FrameControl.Caption = "Control Nº " + CStr(ml_IdControlSeleccionado)
            End If
'        Else
'            FrameControl.Caption = "Control Nº " + CStr(ml_IdControlSeleccionado)
'        End If
    End If

    If mb_ConsultaControl = False Then
        If ml_IdControl = ml_IdControlSeleccionado Then
            CargarDatosDefectoControl
            
            'Frank 03092014
            If ml_IdPrograma = 1 Then 'Solo para el Programa Materno
                If ml_IdControl = 1 Then
                    Cargar_DiagnosticoPorDefecto_PrimerControl
                End If
            End If
            
        End If
    Else
        If ml_IdControl = ml_IdControlSeleccionado Then
            Actualiza_Peso_Presion ml_peso, ml_presion, ml_talla
        End If
    End If

    If (ml_IdControlSeleccionado < ml_IdControl) Or _
           (ml_IdControlSeleccionado = ml_IdControl And mb_ConsultaControl = True) Then
            'Carga Fecha Control
            Set oRsTmp1 = mo_reglasComunes.ProControlesSeleccionarPorId(ml_IdPrograma, ml_IdProCabecera, ml_IdControlSeleccionado, oConexion)
            If oRsTmp1.RecordCount > 0 Then
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                   txtFechaControl.Text = oRsTmp1.Fields!FechaControl
'                   txtFechaControl.Enabled = False
                   mo_Formulario.HabilitarDeshabilitar txtFechaControl, False
                   oRsTmp1.MoveNext
                Loop
            End If
            'Cargar Datos dinamicos de control
            Set oRsTmp1 = mo_reglasComunes.ProControlDatoSeleccionarPorId(ml_IdPrograma, ml_IdProCabecera, ml_IdControlSeleccionado, oConexion)
            If oRsTmp1.RecordCount > 0 Then
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                   txtDatoControl(oRsTmp1.Fields!idcontroldato).Text = oRsTmp1.Fields!controldato
                   oRsTmp1.MoveNext
                 Loop
             End If
             'Cargar Diagnosticos
             If oRsTmp1.State = 1 Then oRsTmp1.Close
             Set oRsTmp1 = mo_reglasComunes.ProDiagnosticosSeleccionarPorIdControl(ml_IdPrograma, ml_IdProCabecera, ml_IdControlSeleccionado, oConexion)
             If oRsTmp1.RecordCount > 0 Then
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                    For lnFor1 = 1 To 3
                        For lnFor = 0 To cmbDiagnosticos.ListCount - 1
                            If Val(Mid(cmbDiagnosticos.ListItems.Item(lnFor).Key, 2, 100)) = oRsTmp1.Fields!idDiagnostico Then
                                cmbDiagnosticos.ListItems.Item(lnFor).Selected = True
                            End If
                        Next
                    Next
                    oRsProCatalogoDiagnosticos.AddNew
                    oRsProCatalogoDiagnosticos.Fields!idDiagnostico = oRsTmp1.Fields!idDiagnostico
                    oRsProCatalogoDiagnosticos.Fields!CodigoCiE10 = oRsTmp1.Fields!CodigoCiE10
                    oRsProCatalogoDiagnosticos.Fields!Descripcion = oRsTmp1.Fields!CodigoCiE10 + " - " + oRsTmp1.Fields!Descripcion
                    oRsProCatalogoDiagnosticos.Fields!Principal = oRsTmp1.Fields!Principal
                    oRsProCatalogoDiagnosticos.Fields!labConfHIS = IIf(IsNull(oRsTmp1.Fields!labConfHIS), Space(3), oRsTmp1.Fields!labConfHIS)
                    oRsProCatalogoDiagnosticos.Fields!IdSubclasificacionDx = IIf(IsNull(oRsTmp1.Fields!IdSubclasificacionDx), 102, oRsTmp1.Fields!IdSubclasificacionDx)
                    oRsProCatalogoDiagnosticos.Update
                    oRsTmp1.MoveNext
                Loop
                oRsProCatalogoDiagnosticos.MoveFirst
             End If
            'Cargar Procedimientos
            If oRsTmp1.State = 1 Then oRsTmp1.Close
            Set oRsTmp1 = mo_reglasComunes.ProProcedimientosSeleccionarPorIdControl(ml_IdPrograma, ml_IdProCabecera, ml_IdControlSeleccionado, oConexion)
            If oRsTmp1.RecordCount > 0 Then
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                      If oRsTmp1.Fields!Seleccionado = 1 Then
                        oRsProCatalogoProcedimientos.AddNew
                        oRsProCatalogoProcedimientos.Fields!idDiagnostico = oRsTmp1.Fields!idDiagnostico
                        oRsProCatalogoProcedimientos.Fields!idProducto = oRsTmp1.Fields!idProducto
                        oRsProCatalogoProcedimientos.Fields!procedimiento = oRsTmp1.Fields!nombre
                        oRsProCatalogoProcedimientos.Fields!CodigoCiE10 = oRsTmp1.Fields!CodigoCiE10
                        oRsProCatalogoProcedimientos.Fields!labConfHIS = oRsTmp1.Fields!labConfHIS
                        oRsProCatalogoProcedimientos.Fields!IDRESULTADO = oRsTmp1.Fields!IDRESULTADO
'                        If oRsTmp1.Fields!Seleccionado = 0 Then
'                              oRsProCatalogoProcedimientos.Fields!seleccionar = False
'                        Else
                        oRsProCatalogoProcedimientos.Fields!seleccionar = True
'                        End If
                        oRsProCatalogoProcedimientos.Update
                      End If
                      oRsTmp1.MoveNext
                Loop
               oRsProCatalogoProcedimientos.MoveFirst
            End If
            'Cargar Tratamientos
            With oRsProCatalogoTratamientos
                If .RecordCount > 0 Then
                   .MoveFirst
                   Do While Not .EOF
                      .Delete
                      .Update
                      .MoveNext
                   Loop
                End If
            End With
    
            If oRsTmp1.State = 1 Then oRsTmp1.Close
            Set oRsTmp1 = mo_reglasComunes.ProTratamientosSeleccionarPorIdControl(ml_IdPrograma, ml_IdProCabecera, ml_IdControlSeleccionado, oConexion)
            If oRsTmp1.RecordCount > 0 Then
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                    oRsProCatalogoTratamientos.AddNew
                    oRsProCatalogoTratamientos.Fields!idProducto = oRsTmp1.Fields!idProducto
                    oRsProCatalogoTratamientos.Fields!Tratamientos = oRsTmp1.Fields!Tratamientos
                    If oRsTmp1.Fields!Seleccionado = 0 Then
                         oRsProCatalogoTratamientos.Fields!seleccionar = False
                    Else
                         oRsProCatalogoTratamientos.Fields!seleccionar = True
                    End If
                    oRsProCatalogoTratamientos.Update
                    oRsTmp1.MoveNext
                Loop
               oRsProCatalogoTratamientos.MoveFirst
            End If
            
            'Cargar grafico con el ultimo control seleccionado
            
            
    End If
    
    grdDiagnosticos.Bands(0).Columns("principal").Activation = ssActivationAllowEdit
    grdDiagnosticos.Bands(0).Columns("labConfHIS").Activation = ssActivationAllowEdit
    grdProcedimientos.Bands(0).Columns("seleccionar").Activation = ssActivationAllowEdit
    grdProcedimientos.Bands(0).Columns("IdResultado").Activation = ssActivationAllowEdit
    grdTratamientos.Bands(0).Columns("seleccionar").Activation = ssActivationAllowEdit

    oRsProControlConfigDatos.MoveFirst
    Do While Not oRsProControlConfigDatos.EOF
        mo_Formulario.HabilitarDeshabilitar txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato), True
        If oRsProControlConfigDatos.Fields!Control_EsDatoCalculado Then
            mo_Formulario.HabilitarDeshabilitar txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato), False
        End If
        If oRsProControlConfigDatos.Fields!Control_EsPresion Or _
           oRsProControlConfigDatos.Fields!Control_EsPeso Or _
           oRsProControlConfigDatos.Fields!Control_EsTalla Then
            mo_Formulario.HabilitarDeshabilitar txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato), False
        End If
        oRsProControlConfigDatos.MoveNext
    Loop
    
    'Desactivar edicion para un control anterior
    If ml_IdControlSeleccionado < ml_IdControl Then
        cmbDiagnosticos.Enabled = False
        mo_Formulario.HabilitarDeshabilitar btnQuitar, False
        mo_Formulario.HabilitarDeshabilitar btnBusquedaDiagnostico, False
        
        mo_Formulario.HabilitarDeshabilitar cmbDxPrograma, False
        mo_Formulario.HabilitarDeshabilitar txtFiltroDx, False
        mo_Formulario.HabilitarDeshabilitar btnAgregar, False
        
        mo_Formulario.HabilitarDeshabilitar txtFiltraCPT, False
        mo_Formulario.HabilitarDeshabilitar cmbProcPrograma, False
        mo_Formulario.HabilitarDeshabilitar btnAgregarProcedminiento, False
        mo_Formulario.HabilitarDeshabilitar btnQuitaOtrosProcedimientos, False
        mo_Formulario.HabilitarDeshabilitar btnBuscarProc, False
        
        mo_Formulario.HabilitarDeshabilitar txtFiltraCPT, False
        mo_Formulario.HabilitarDeshabilitar cmbProcPrograma, False
        mo_Formulario.HabilitarDeshabilitar btnAgregarProcedminiento, False
        mo_Formulario.HabilitarDeshabilitar btnQuitaOtrosProcedimientos, False
        mo_Formulario.HabilitarDeshabilitar btnBuscarProc, False
    
'        grdDiagnosticos.Bands(0).Columns("principal").Activation = ssActivationActivateNoEdit
'        grdProcedimientos.Bands(0).Columns("seleccionar").Activation = ssActivationActivateNoEdit
'        grdTratamientos.Bands(0).Columns("seleccionar").Activation = ssActivationActivateNoEdit
        
        grdDiagnosticos.Bands(0).Columns("principal").Activation = ssActivationActivateNoEdit
        grdDiagnosticos.Bands(0).Columns("labConfHIS").Activation = ssActivationActivateNoEdit
        grdProcedimientos.Bands(0).Columns("seleccionar").Activation = ssActivationActivateNoEdit
        grdProcedimientos.Bands(0).Columns("IdResultado").Activation = ssActivationActivateNoEdit
        grdTratamientos.Bands(0).Columns("seleccionar").Activation = ssActivationActivateNoEdit
    
        mo_Formulario.HabilitarDeshabilitar txtFechaControl, False
        oRsProControlConfigDatos.MoveFirst
        Do While Not oRsProControlConfigDatos.EOF
            mo_Formulario.HabilitarDeshabilitar txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato), False
            oRsProControlConfigDatos.MoveNext
        Loop
    End If
    CargaGraficoChartSpace

    oConexion.Close
    Set oConexion = Nothing
End Sub

Sub CargaComboProDiagnosticos(oConexion As Connection)
      Dim lnIdListItem1 As Integer
      Dim oRsTmp1 As New Recordset
      Set oRsTmp1 = mo_reglasComunes.ProCatalogoDiagnosticosSeleccionarPorIdPrograma(ml_IdPrograma, oConexion)
      If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
            cmbDiagnosticos.ListItems.Add lnIdListItem1, lcCombo & Trim(Str(oRsTmp1.Fields!idDiagnostico)), oRsTmp1.Fields!DIAGNOSTICO
            lnIdListItem1 = lnIdListItem1 + 1
            oRsTmp1.MoveNext
        Loop
      End If
      
      If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Set mo_cmbDxPrograma.MiComboBox = cmbDxPrograma
        mo_cmbDxPrograma.BoundColumn = "IdDiagnostico"
        mo_cmbDxPrograma.ListField = "DIAGNOSTICO"
        Set mo_cmbDxPrograma.RowSource = oRsTmp1
      End If
      
      Set mo_RsLabHis = mo_reglasComunes.DevuelveHIS_SITUACIOporDescripcion()
      Set cmbLabHisDx.ListSource = mo_RsLabHis
      oRsTmp1.Close
      Set oRsTmp1 = Nothing
End Sub

Sub CargaComboProProcedimientos(oConexion As Connection)
      Dim lnIdListItem1 As Integer
      Dim oRsTmp1 As New Recordset
      Set oRsTmp1 = mo_reglasComunes.ProCatalogoCptSeleccionarPorIdPrograma(ml_IdPrograma, oConexion)
     
      If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Set mo_cmbCptPrograma.MiComboBox = cmbProcPrograma
        mo_cmbCptPrograma.BoundColumn = "IdProducto"
        mo_cmbCptPrograma.ListField = "PROCEDIMIENTO"
        Set mo_cmbCptPrograma.RowSource = oRsTmp1
      End If
      Set cmbLabHisCpt.ListSource = mo_RsLabHis
      oRsTmp1.Close
      Set oRsTmp1 = Nothing
End Sub

Sub CargaComboProTratamientos(oConexion As Connection)
      Dim lnIdListItem1 As Integer
      Dim oRsTmp1 As New Recordset
      Set oRsTmp1 = mo_reglasComunes.ProCatalogoTratamientosSeleccionarPorIdPrograma(ml_IdPrograma, oConexion)
     
      If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Set mo_cmbTtoPrograma.MiComboBox = cmbTtoProgama
        mo_cmbTtoPrograma.BoundColumn = "IdProducto"
        mo_cmbTtoPrograma.ListField = "Tratamientos"
        Set mo_cmbTtoPrograma.RowSource = oRsTmp1
      End If
      oRsTmp1.Close
      Set oRsTmp1 = Nothing
End Sub

Sub CargaCatalogoControles(oConexion As Connection)
     Dim oRsTmp1 As New Recordset
     Set oRsTmp1 = mo_reglasComunes.ProCatalogoControlesPorIdPrograma(ml_IdPrograma, oConexion)
     If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
           oRsProCatalogoControles.AddNew
           oRsProCatalogoControles.Fields!IdControl = oRsTmp1.Fields!IdControl
           oRsProCatalogoControles.Fields!Descripcion = oRsTmp1.Fields!Descripcion
           oRsProCatalogoControles.Fields!FechaControl = ""
           oRsProCatalogoControles.Update
           oRsTmp1.MoveNext
        Loop
        oRsProCatalogoControles.MoveFirst
     End If
     oRsTmp1.Close
     Set oRsTmp1 = Nothing
End Sub

Sub CargaCatalogoTratamientos(oConexion As Connection)
     Dim oRsTmp1 As New Recordset
     Set oRsTmp1 = mo_reglasComunes.ProCatalogoTratamientosSeleccionarPorIdPrograma(ml_IdPrograma, oConexion)
     If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
           oRsProCatalogoTratamientos.AddNew
           oRsProCatalogoTratamientos.Fields!idProducto = oRsTmp1.Fields!idProducto
           oRsProCatalogoTratamientos.Fields!Tratamientos = oRsTmp1.Fields!Tratamientos
           oRsProCatalogoTratamientos.Fields!seleccionar = False
           oRsProCatalogoTratamientos.Update
           oRsTmp1.MoveNext
        Loop
        oRsProCatalogoTratamientos.MoveFirst
     End If
     oRsTmp1.Close
     Set oRsTmp1 = Nothing
End Sub

Sub ConfiguraCabeceraPrograma(oConexion As Connection)
   Dim ml_FilaDatoCabecera As Integer
   Dim ml_ColumnaDatoCabecera As Integer
   Dim ml_numero_ctrlCabecera As Integer
   Dim oRsTmp1 As New Recordset
   'Configura controles para los datos de cabecera
    Set oRsTmp1 = mo_reglasComunes.ProCabeceraConfigDatosSeleccionarPorIdPrograma(ml_IdPrograma, oConexion)
    If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
              oRsProCabeceraConfigDatos.AddNew
              oRsProCabeceraConfigDatos.Fields!IdCabDato = oRsTmp1.Fields!IdCabDato
              oRsProCabeceraConfigDatos.Fields!Cab_Texto = oRsTmp1.Fields!Cab_Texto
              oRsProCabeceraConfigDatos.Fields!cab_tipo = oRsTmp1.Fields!cab_tipo
              oRsProCabeceraConfigDatos.Fields!cab_ancho = oRsTmp1.Fields!cab_ancho
              oRsProCabeceraConfigDatos.Fields!Cab_EsDatoObligatorio = oRsTmp1.Fields!Cab_EsDatoObligatorio
              oRsProCabeceraConfigDatos.Fields!Cab_TextoToolTip = oRsTmp1.Fields!Cab_TextoToolTip
              oRsProCabeceraConfigDatos.Fields!Cab_EsDatoCalculado = oRsTmp1.Fields!Cab_EsDatoCalculado
              oRsProCabeceraConfigDatos.Fields!Cab_FormulaCalculaValor = oRsTmp1.Fields!Cab_FormulaCalculaValor
              oRsProCabeceraConfigDatos.Fields!Cab_EsDatoCalculador = oRsTmp1.Fields!Cab_EsDatoCalculador
              oRsProCabeceraConfigDatos.Fields!Cab_RangoInicial = IIf(IsNull(oRsTmp1.Fields!Cab_RangoInicial), 0, oRsTmp1.Fields!Cab_RangoInicial)
              oRsProCabeceraConfigDatos.Fields!Cab_RangoFinal = IIf(IsNull(oRsTmp1.Fields!Cab_RangoFinal), 0, oRsTmp1.Fields!Cab_RangoFinal)
              oRsProCabeceraConfigDatos.Fields!Cab_Fila = IIf(IsNull(oRsTmp1.Fields!Cab_Fila), 0, oRsTmp1.Fields!Cab_Fila)
              oRsProCabeceraConfigDatos.Fields!Cab_Columna = IIf(IsNull(oRsTmp1.Fields!Cab_Columna), 0, oRsTmp1.Fields!Cab_Columna)
              oRsProCabeceraConfigDatos.Update
              oRsTmp1.MoveNext
        Loop
      oRsProCabeceraConfigDatos.MoveFirst
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
    
    ml_numero_ctrlCabecera = 1
    Do While Not oRsProCabeceraConfigDatos.EOF
        'Calcula coordenadas datos de cabecera
'        ml_ColumnaDatoCabecera = IIf(Round(ml_numero_ctrlCabecera / 3) < ml_numero_ctrlCabecera / 3, Round(ml_numero_ctrlCabecera / 3) + 1, Round(ml_numero_ctrlCabecera / 3))
'        ml_FilaDatoCabecera = IIf(ml_numero_ctrlCabecera Mod 3 = 0, 3, ml_numero_ctrlCabecera Mod 3)

        If oRsProCabeceraConfigDatos.Fields!IdCabDato = 8 Or oRsProCabeceraConfigDatos.Fields!IdCabDato = 12 Then
            If ml_IdPrograma <> 1 Then
                ml_ColumnaDatoCabecera = oRsProCabeceraConfigDatos.Fields!Cab_Columna
                ml_FilaDatoCabecera = oRsProCabeceraConfigDatos.Fields!Cab_Fila
            End If
        Else
            ml_ColumnaDatoCabecera = oRsProCabeceraConfigDatos.Fields!Cab_Columna
            ml_FilaDatoCabecera = oRsProCabeceraConfigDatos.Fields!Cab_Fila
        End If
        
        'Visualiza los label para el ingreso de  datos de cabecera
        If oRsProCabeceraConfigDatos.Fields!IdCabDato = 8 Or oRsProCabeceraConfigDatos.Fields!IdCabDato = 12 Then
            Load lblDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato)
            If ml_IdPrograma <> 1 Then
                lblDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Visible = True
            Else
                lblDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Visible = False
            End If
            lblDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Left = (ml_ColumnaDatoCabecera - 1) * 3700 + 300
            lblDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Width = 2000
            lblDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Caption = oRsProCabeceraConfigDatos.Fields!Cab_Texto
            lblDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Top = (ml_FilaDatoCabecera - 1) * 320 + 260
        Else
            Load lblDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato)
            lblDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Visible = True
            lblDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Left = (ml_ColumnaDatoCabecera - 1) * 3700 + 300
            lblDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Width = 2000
            lblDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Caption = oRsProCabeceraConfigDatos.Fields!Cab_Texto
            lblDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Top = (ml_FilaDatoCabecera - 1) * 320 + 260
        End If

        
        'Visualiza los textbox para el ingreso de datos de cabecera
        Select Case oRsProCabeceraConfigDatos.Fields!cab_tipo
        Case "ValorEntero", "ValorFecha", "ValorTexto", "ValorDouble"
            If oRsProCabeceraConfigDatos.Fields!IdCabDato = 8 Or oRsProCabeceraConfigDatos.Fields!IdCabDato = 12 Then
                If ml_IdPrograma = 1 Then
                    Load txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato)
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Visible = True
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Left = (ml_ColumnaDatoCabecera - 1) * 3700 + 3020
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Width = 705
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Top = (ml_FilaDatoCabecera - 1) * 320 + 260
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).TabIndex = ml_IdTabCtrl
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).ToolTipText = oRsProCabeceraConfigDatos.Fields!Cab_TextoToolTip
                    If oRsProCabeceraConfigDatos.Fields!Cab_EsDatoCalculado Then
                        mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato), False
                    End If
                    MasKTextoDato txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato), oRsProCabeceraConfigDatos.Fields!cab_tipo, oRsProCabeceraConfigDatos.Fields!cab_ancho
                Else
                    Load txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato)
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Visible = True
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Left = (ml_ColumnaDatoCabecera - 1) * 3700 + 2300
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Width = 1425
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Top = (ml_FilaDatoCabecera - 1) * 320 + 260
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).TabIndex = ml_IdTabCtrl
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).ToolTipText = oRsProCabeceraConfigDatos.Fields!Cab_TextoToolTip
                    If oRsProCabeceraConfigDatos.Fields!Cab_EsDatoCalculado Then
                        mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato), False
                    End If
                    MasKTextoDato txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato), oRsProCabeceraConfigDatos.Fields!cab_tipo, oRsProCabeceraConfigDatos.Fields!cab_ancho
                End If
            Else
                If oRsProCabeceraConfigDatos.Fields!IdCabDato = 7 Or oRsProCabeceraConfigDatos.Fields!IdCabDato = 11 Then
                    If ml_IdPrograma = 1 Then
                        Load txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato)
                        txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Visible = True
                        txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Left = (ml_ColumnaDatoCabecera - 1) * 3700 + 2300
                        txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Width = 705
                        txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Top = (ml_FilaDatoCabecera - 1) * 320 + 260
                        txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).TabIndex = ml_IdTabCtrl
                        txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).ToolTipText = oRsProCabeceraConfigDatos.Fields!Cab_TextoToolTip
                        If oRsProCabeceraConfigDatos.Fields!Cab_EsDatoCalculado Then
                            mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato), False
                        End If
                        MasKTextoDato txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato), oRsProCabeceraConfigDatos.Fields!cab_tipo, oRsProCabeceraConfigDatos.Fields!cab_ancho
                    Else
                        Load txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato)
                        txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Visible = True
                        txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Left = (ml_ColumnaDatoCabecera - 1) * 3700 + 2300
                        txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Width = 1425
                        txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Top = (ml_FilaDatoCabecera - 1) * 320 + 260
                        txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).TabIndex = ml_IdTabCtrl
                        txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).ToolTipText = oRsProCabeceraConfigDatos.Fields!Cab_TextoToolTip
                        If oRsProCabeceraConfigDatos.Fields!Cab_EsDatoCalculado Then
                            mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato), False
                        End If
                        MasKTextoDato txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato), oRsProCabeceraConfigDatos.Fields!cab_tipo, oRsProCabeceraConfigDatos.Fields!cab_ancho
                    End If
                Else
                    Load txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato)
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Visible = True
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Left = (ml_ColumnaDatoCabecera - 1) * 3700 + 2300
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Width = 1425
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Top = (ml_FilaDatoCabecera - 1) * 320 + 260
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).TabIndex = ml_IdTabCtrl
                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).ToolTipText = oRsProCabeceraConfigDatos.Fields!Cab_TextoToolTip
                    If oRsProCabeceraConfigDatos.Fields!Cab_EsDatoCalculado Then
                        mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato), False
                    End If
                    MasKTextoDato txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato), oRsProCabeceraConfigDatos.Fields!cab_tipo, oRsProCabeceraConfigDatos.Fields!cab_ancho
                    
                    If ml_IdPrograma = 1 Then
                        If oRsProCabeceraConfigDatos.Fields!IdCabDato = 13 Then
                            lblUnidades.Top = (ml_FilaDatoCabecera - 1) * 320 + 280
                            lblUnidades.Left = (ml_ColumnaDatoCabecera - 1) * 3700 + 3870
                        End If
                    End If
                End If
           End If
        Case "ValorCheck"
            If oRsProCabeceraConfigDatos.Fields!IdCabDato = 7 Or oRsProCabeceraConfigDatos.Fields!IdCabDato = 11 Then
                If ml_IdPrograma = 1 Then
                    Load chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato)
                    chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Visible = True
                    chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Left = (ml_ColumnaDatoCabecera - 1) * 3700 + 2300
                    chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Width = 705
                    chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Top = (ml_FilaDatoCabecera - 1) * 320 + 260
                    chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).TabIndex = ml_IdTabCtrl
        '            chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).ToolTipText = oRsProCabeceraConfigDatos.Fields!Cab_TextoToolTip
                    chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Caption = oRsProCabeceraConfigDatos.Fields!Cab_TextoToolTip
                Else
                    Load chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato)
                    chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Visible = True
                    chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Left = (ml_ColumnaDatoCabecera - 1) * 3700 + 2300
                    chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Width = 1425
                    chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Top = (ml_FilaDatoCabecera - 1) * 320 + 260
                    chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).TabIndex = ml_IdTabCtrl
        '            chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).ToolTipText = oRsProCabeceraConfigDatos.Fields!Cab_TextoToolTip
                    chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Caption = oRsProCabeceraConfigDatos.Fields!Cab_TextoToolTip
                End If
            Else
                Load chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato)
                chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Visible = True
                chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Left = (ml_ColumnaDatoCabecera - 1) * 3700 + 2300
                chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Width = 1425
                chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Top = (ml_FilaDatoCabecera - 1) * 320 + 260
                chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).TabIndex = ml_IdTabCtrl
    '            chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).ToolTipText = oRsProCabeceraConfigDatos.Fields!Cab_TextoToolTip
                chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Caption = oRsProCabeceraConfigDatos.Fields!Cab_TextoToolTip
            End If
        End Select
        
        ml_numero_ctrlCabecera = ml_numero_ctrlCabecera + 1
        ml_IdTabCtrl = ml_IdTabCtrl + 1 'TabIndex
        oRsProCabeceraConfigDatos.MoveNext
    Loop
    
    If ml_IdPrograma = 1 Then 'Valores bloqueado inicialmente - Modulo Materno
        txtDatoCabecera(8).Text = "0"
        mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(8), False
        
        txtDatoCabecera(10).Text = txtDatoCabecera(10).Tag
        mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(10), False
        
        txtDatoCabecera(11).Text = ""
        mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(11), False
        
        txtDatoCabecera(12).Text = ""
        mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(12), False
    Else
    
    End If
End Sub

Sub ConfiguraControlesPrograma(oConexion As Connection)
    Dim ml_FilaDatoControl As Integer
    Dim ml_ColumnaDatoControl As Integer
    Dim ml_numero_ctrlControl As Integer
    Dim oRsTmp1 As New Recordset
    
    'Configura controles para los datos de control
    'Primero la fecha del control
     lblFechaControl.Left = 300
     lblFechaControl.Width = 2000
     lblFechaControl.Top = 280
     txtFechaControl.Left = 2300
     txtFechaControl.Width = 1425
     txtFechaControl.Top = 280
     txtFechaControl.TabIndex = ml_IdTabCtrl
     ml_IdTabCtrl = ml_IdTabCtrl + 1
     
    'Segundo los controles dinamicos
     Set oRsTmp1 = mo_reglasComunes.ProControlConfigDatoSeleccionarPorIdPrograma(ml_IdPrograma, oConexion)
     If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
           oRsProControlConfigDatos.AddNew
           oRsProControlConfigDatos.Fields!idcontroldato = oRsTmp1.Fields!idcontroldato
           oRsProControlConfigDatos.Fields!control_texto = oRsTmp1.Fields!control_texto
           oRsProControlConfigDatos.Fields!Control_Tipo = oRsTmp1.Fields!Control_Tipo
           oRsProControlConfigDatos.Fields!control_ancho = oRsTmp1.Fields!control_ancho
           oRsProControlConfigDatos.Fields!Control_EsDatoObligatorio = oRsTmp1.Fields!Control_EsDatoObligatorio
           oRsProControlConfigDatos.Fields!Control_TextoToolTip = oRsTmp1.Fields!Control_TextoToolTip
           oRsProControlConfigDatos.Fields!Control_EsPresion = oRsTmp1.Fields!Control_EsPresion
           oRsProControlConfigDatos.Fields!Control_EsPeso = oRsTmp1.Fields!Control_EsPeso
           oRsProControlConfigDatos.Fields!Control_EsTalla = oRsTmp1.Fields!Control_EsTalla
           oRsProControlConfigDatos.Fields!Control_EsDatoCalculado = oRsTmp1.Fields!Control_EsDatoCalculado
           oRsProControlConfigDatos.Fields!Control_FormulaCalculaValor = oRsTmp1.Fields!Control_FormulaCalculaValor
           oRsProControlConfigDatos.Fields!Control_EsDatoGrafico = oRsTmp1.Fields!Control_EsDatoGrafico
           oRsProControlConfigDatos.Fields!control_esgraficoejex = oRsTmp1.Fields!control_esgraficoejex
           oRsProControlConfigDatos.Fields!control_fila = oRsTmp1.Fields!control_fila
           oRsProControlConfigDatos.Fields!control_columna = oRsTmp1.Fields!control_columna
           oRsProControlConfigDatos.Update
           oRsTmp1.MoveNext
        Loop
        oRsProControlConfigDatos.MoveFirst
     End If
     oRsTmp1.Close
     Set oRsTmp1 = Nothing
    
     ml_numero_ctrlControl = 2
     Do While Not oRsProControlConfigDatos.EOF
        If oRsProControlConfigDatos.Fields!Control_EsDatoGrafico Then
            Load lblDatoControl(oRsProControlConfigDatos.Fields!idcontroldato)
            Load txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato)
        Else
            'Calcula coordenadas datos de control
'            ml_ColumnaDatoControl = IIf(Round(ml_numero_ctrlControl / 3) < ml_numero_ctrlControl / 3, Round(ml_numero_ctrlControl / 3) + 1, Round(ml_numero_ctrlControl / 3))
'            ml_FilaDatoControl = IIf(ml_numero_ctrlControl Mod 3 = 0, 3, ml_numero_ctrlControl Mod 3)

            ml_ColumnaDatoControl = oRsProControlConfigDatos.Fields!control_columna
            ml_FilaDatoControl = oRsProControlConfigDatos.Fields!control_fila
            
            'Visualiza los label para el ingreso de  datos de control
            Load lblDatoControl(oRsProControlConfigDatos.Fields!idcontroldato)
            lblDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Visible = True
            lblDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Left = (ml_ColumnaDatoControl - 1) * 3700 + 300
            lblDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Width = 2000
            lblDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Caption = oRsProControlConfigDatos.Fields!control_texto
            lblDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Top = (ml_FilaDatoControl - 1) * 315 + 280
            
            'Visualiza los textbox para el ingreso de  datos de control
            Load txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato)
            txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Visible = True
            txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Left = (ml_ColumnaDatoControl - 1) * 3700 + 2300
            txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Width = 1425
            txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Top = (ml_FilaDatoControl - 1) * 315 + 280
            txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).TabIndex = ml_IdTabCtrl
            txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).ToolTipText = oRsProControlConfigDatos.Fields!Control_TextoToolTip
            If oRsProControlConfigDatos.Fields!Control_EsDatoCalculado Then mo_Formulario.HabilitarDeshabilitar txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato), False
            If oRsProControlConfigDatos.Fields!Control_EsPresion Or _
               oRsProControlConfigDatos.Fields!Control_EsPeso Or _
               oRsProControlConfigDatos.Fields!Control_EsTalla Then
                    mo_Formulario.HabilitarDeshabilitar txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato), False
            End If
            MasKTextoDato txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato), oRsProControlConfigDatos.Fields!Control_Tipo, oRsProControlConfigDatos.Fields!control_ancho
            ml_numero_ctrlControl = ml_numero_ctrlControl + 1
        End If
        ml_IdTabCtrl = ml_IdTabCtrl + 1 'TabIndex
        oRsProControlConfigDatos.MoveNext
    Loop
End Sub

Sub MasKTextoDato(oTexto As MaskEdBox, Lctipo As String, LnAncho As Integer)
    Select Case Lctipo
        Case "ValorFecha"
            oTexto.Mask = "##/##/####"
            oTexto.Tag = sighentidades.FECHA_VACIA_DMY
        Case "ValorEntero", "ValorTexto" '"ValorDouble"
            oTexto.MaxLength = LnAncho
        Case "ValorDouble"
            oTexto.Mask = "###.##"
            oTexto.Tag = "___.__"
    End Select
End Sub

Public Function DevuelveEsControlParaActualizar() As Boolean
    DevuelveEsControlParaActualizar = mb_ConsultaControl
End Function

Public Function EsControlActual() As Boolean
    EsControlActual = False
    If ml_IdControl = ml_IdControlSeleccionado Then
        EsControlActual = True
    End If
End Function

Public Function DevuelveProCabecera() As DoProCabecera
    Dim oDoProCabecera As New DoProCabecera
    With oDoProCabecera
        .IdPrograma = ml_IdPrograma
        .idPaciente = ml_IdPaciente
        .IdProCabecera = ml_IdProCabecera
        .IdUsuarioAuditoria = ml_idUsuario
    End With
    Set DevuelveProCabecera = oDoProCabecera
End Function

Public Function DevuelveDatosProCabecera() As ADODB.Recordset
Dim oRsProCabeceraDato As New Recordset

Set DevuelveDatosProCabecera = Nothing
    'Crea recordset DatosProCabecera
    If oRsProCabeceraDato.State = 1 Then
       Set oRsProCabeceraDato = Nothing
    End If
    With oRsProCabeceraDato
          .Fields.Append "IdPrograma", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdProCabecera", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdCabDato", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "CabDato", adVarChar, 255, adFldIsNullable + adFldUpdatable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    
    'Lee los datos dinamicos de cabecera
'    Set oRsProCabeceraConfigDatos = mo_ReglasComunes.ProCabeceraConfigDatosSeleccionarPorIdPrograma(ml_IdPrograma, oConexion)
     If oRsProCabeceraConfigDatos.RecordCount > 0 Then
        oRsProCabeceraConfigDatos.MoveFirst
        Do While Not oRsProCabeceraConfigDatos.EOF
           oRsProCabeceraDato.AddNew
           oRsProCabeceraDato.Fields!IdPrograma = ml_IdPrograma
           oRsProCabeceraDato.Fields!IdProCabecera = ml_IdProCabecera
           oRsProCabeceraDato.Fields!IdCabDato = oRsProCabeceraConfigDatos.Fields!IdCabDato
           Select Case oRsProCabeceraConfigDatos.Fields!cab_tipo
           Case "ValorEntero", "ValorFecha", "ValorTexto", "ValorDouble"
                oRsProCabeceraDato.Fields!CabDato = txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Text
           Case "ValorCheck"
                oRsProCabeceraDato.Fields!CabDato = chkDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Value
           End Select
           oRsProCabeceraDato.Update
           oRsProCabeceraConfigDatos.MoveNext
        Loop
        oRsProCabeceraDato.MoveFirst
     End If
Set DevuelveDatosProCabecera = oRsProCabeceraDato
End Function

Public Function DevuelveProControles() As DOProControles
    Dim oDOProControles As New DOProControles
    With oDOProControles
        .IdPrograma = ml_IdPrograma
        .IdProCabecera = ml_IdProCabecera
        .IdControl = ml_IdControl
        .idAtencion = ml_idAtencion
        If sighentidades.EsFecha(txtFechaControl.Text, "DD/MM/AAAA", False) = True Then
            .FechaControl = txtFechaControl.Text
        Else
            .FechaControl = ml_FechaAtencion
        End If
        .IdUsuarioAuditoria = ml_idUsuario
    End With
    Set DevuelveProControles = oDOProControles
End Function

Public Function DevuelveDatosProControles() As ADODB.Recordset
Dim oRsProControlDato As New Recordset
Dim LnDiferencia As Long

Set DevuelveDatosProControles = Nothing
    'Crea recordset DatosProCabecera
    If oRsProControlDato.State = 1 Then
       Set oRsProControlDato = Nothing
    End If
    With oRsProControlDato
          .Fields.Append "IdPrograma", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdProCabecera", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdControl", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdControlDato", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "ControlDato", adVarChar, 255, adFldIsNullable + adFldUpdatable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    
    'Lee los datos dinamicos de control
     If oRsProControlConfigDatos.RecordCount > 0 Then
        oRsProControlConfigDatos.MoveFirst
        Do While Not oRsProControlConfigDatos.EOF
           oRsProControlDato.AddNew
           oRsProControlDato.Fields!IdPrograma = ml_IdPrograma
           oRsProControlDato.Fields!IdProCabecera = ml_IdProCabecera
           oRsProControlDato.Fields!IdControl = ml_IdControl
           oRsProControlDato.Fields!idcontroldato = oRsProControlConfigDatos.Fields!idcontroldato
           oRsProControlDato.Fields!controldato = txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Text
           oRsProControlDato.Update
           oRsProControlConfigDatos.MoveNext
        Loop
        oRsProControlDato.MoveFirst
      End If
Set DevuelveDatosProControles = oRsProControlDato
End Function

Public Function DevuelveProDiagnosticos() As ADODB.Recordset
Dim oRsProDiagnosticos As New Recordset

Set DevuelveProDiagnosticos = Nothing
    'Crea recordset ProDiagnosticos
    If oRsProDiagnosticos.State = 1 Then
       Set oRsProDiagnosticos = Nothing
    End If
    With oRsProDiagnosticos
          .Fields.Append "IdPrograma", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdProCabecera", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdControl", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdDiagnostico", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "Principal", adBoolean
          .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdSubClasificacionDX", adVarChar, 3, adFldIsNullable + adFldUpdatable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    
    If oRsProCatalogoDiagnosticos.RecordCount > 0 Then
        oRsProCatalogoDiagnosticos.MoveFirst
        Do While Not oRsProCatalogoDiagnosticos.EOF
           oRsProDiagnosticos.AddNew
           oRsProDiagnosticos.Fields!IdPrograma = ml_IdPrograma
           oRsProDiagnosticos.Fields!IdProCabecera = ml_IdProCabecera
           oRsProDiagnosticos.Fields!IdControl = ml_IdControl
           oRsProDiagnosticos.Fields!idDiagnostico = oRsProCatalogoDiagnosticos.Fields!idDiagnostico
           oRsProDiagnosticos.Fields!Principal = oRsProCatalogoDiagnosticos.Fields!Principal
           oRsProDiagnosticos.Fields!labConfHIS = IIf(IsNull(oRsProCatalogoDiagnosticos.Fields!labConfHIS), Space(3), oRsProCatalogoDiagnosticos.Fields!labConfHIS)
           oRsProDiagnosticos.Fields!IdSubclasificacionDx = oRsProCatalogoDiagnosticos.Fields!IdSubclasificacionDx
           oRsProDiagnosticos.Update
           oRsProCatalogoDiagnosticos.MoveNext
        Loop
        oRsProCatalogoDiagnosticos.MoveFirst
        If oRsProDiagnosticos.RecordCount > 0 Then
            oRsProDiagnosticos.MoveFirst
        End If
      End If
Set DevuelveProDiagnosticos = oRsProDiagnosticos
End Function

Public Function DevuelveProProcedimientos() As ADODB.Recordset
Dim oRsProProcedimientos As New Recordset

Set DevuelveProProcedimientos = Nothing
    'Crea recordset ProProcedimientos
    If oRsProProcedimientos.State = 1 Then
       Set oRsProProcedimientos = Nothing
    End If
    With oRsProProcedimientos
          .Fields.Append "IdPrograma", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdProCabecera", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdControl", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdDiagnostico", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdProducto", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdResultado", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "Procedimiento", adVarChar, 250, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    grdProcedimientos.Update
    If oRsProCatalogoProcedimientos.RecordCount > 0 Then
        oRsProCatalogoProcedimientos.MoveFirst
        Do While Not oRsProCatalogoProcedimientos.EOF
'           If oRsProCatalogoProcedimientos.Fields!seleccionar = True Then
                oRsProProcedimientos.AddNew
                oRsProProcedimientos.Fields!IdPrograma = ml_IdPrograma
                oRsProProcedimientos.Fields!IdProCabecera = ml_IdProCabecera
                oRsProProcedimientos.Fields!IdControl = ml_IdControl
                oRsProProcedimientos.Fields!idDiagnostico = oRsProCatalogoProcedimientos.Fields!idDiagnostico
                oRsProProcedimientos.Fields!idProducto = oRsProCatalogoProcedimientos.Fields!idProducto
                oRsProProcedimientos.Fields!labConfHIS = IIf(IsNull(oRsProCatalogoProcedimientos.Fields!labConfHIS), Space(3), oRsProCatalogoProcedimientos.Fields!labConfHIS)
                oRsProProcedimientos.Fields!IDRESULTADO = oRsProCatalogoProcedimientos.Fields!IDRESULTADO
                oRsProProcedimientos.Fields!procedimiento = oRsProCatalogoProcedimientos.Fields!procedimiento
                oRsProProcedimientos.Update
'           End If
           oRsProCatalogoProcedimientos.MoveNext
        Loop
        If oRsProProcedimientos.RecordCount > 0 Then
            oRsProProcedimientos.MoveFirst
        End If
        If oRsProCatalogoProcedimientos.RecordCount > 0 Then
            oRsProCatalogoProcedimientos.MoveFirst
        End If
     End If
Set DevuelveProProcedimientos = oRsProProcedimientos
End Function

Public Function DevuelveProTratamientos() As ADODB.Recordset
Dim oRsProTratamientos As New Recordset

Set DevuelveProTratamientos = Nothing
    'Crea recordset ProTratamientos
    If oRsProTratamientos.State = 1 Then
       Set oRsProTratamientos = Nothing
    End If
    With oRsProTratamientos
          .Fields.Append "IdPrograma", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdProCabecera", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdControl", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdProducto", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "Procedimiento", adVarChar, 250, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    
    If oRsProCatalogoTratamientos.RecordCount > 0 Then
        oRsProCatalogoTratamientos.MoveFirst
        Do While Not oRsProCatalogoTratamientos.EOF
'           If oRsProCatalogoTratamientos.Fields!seleccionar = True Then
            oRsProTratamientos.AddNew
            oRsProTratamientos.Fields!IdPrograma = ml_IdPrograma
            oRsProTratamientos.Fields!IdProCabecera = ml_IdProCabecera
            oRsProTratamientos.Fields!IdControl = ml_IdControl
            oRsProTratamientos.Fields!idProducto = oRsProCatalogoTratamientos.Fields!idProducto
            oRsProTratamientos.Fields!procedimiento = oRsProCatalogoTratamientos.Fields!Tratamientos
            oRsProTratamientos.Update
'           End If
           oRsProCatalogoTratamientos.MoveNext
        Loop
        oRsProCatalogoTratamientos.MoveFirst
        If oRsProTratamientos.RecordCount > 0 Then
            oRsProTratamientos.MoveFirst
        End If
     End If
Set DevuelveProTratamientos = oRsProTratamientos
End Function



Private Sub txtDatoCabecera_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'mo_Teclado.RealizarNavegacion KeyCode, txtDatoCabecera(Index)
    Dim lc_MensajeValidacion As String
    
    If KeyCode = vbKeyReturn Then
        If Index = 6 Then
            TabPrograma.Tab = 1
            txtFiltroDx.SetFocus
            Exit Sub
        Else
            If ml_IdPrograma = 1 Then 'Modulo Materno
                If Index = 10 Or Index = 11 Or Index = 12 Then
                    If txtDatoCabecera(10).Text <> sighentidades.FECHA_VACIA_DMY Then
                        If sighentidades.EsFecha(txtDatoCabecera(10).Text, "DD/MM/AAAA", False) = False Then
                            MsgBox "El valor de " & lblDatoCabecera(10).Caption & " no tiene el formato correcto", vbInformation, "Validación"
                            txtDatoCabecera(10).SetFocus
                            Exit Sub
                        Else
                            If Index = 10 And CDate(txtDatoCabecera(10).Text) > CDate(txtFechaControl.Text) Then
                                MsgBox "La fecha de la ecografía es incorrecta, no puede ser mayor que la fecha de control", vbInformation, "Validación"
                                txtDatoCabecera(10).SetFocus
                                Exit Sub
                            End If
                            If Index = 11 And txtDatoCabecera(11).Text = "" Then
                                MsgBox "Las semanas y dias de gestación del resultado de la ecografía no puedes estar vacías", vbInformation, "Validación"
                                txtDatoCabecera(11).SetFocus
                                Exit Sub
                            End If
                            If Index = 12 And txtDatoCabecera(12).Text = "" Then
                                MsgBox "Las semanas y dias de gestación del resultado de la ecografía no puedes estar vacías", vbInformation, "Validación"
                                txtDatoCabecera(12).SetFocus
                                Exit Sub
                            End If
                            
                            txtDatoCabecera(1).Text = Devuelve_Fecha_FUM_Ecografia(txtDatoCabecera(10).Text, IIf(txtDatoCabecera(11).Text = "", 0, txtDatoCabecera(11).Text), IIf(txtDatoCabecera(12).Text = "", 0, txtDatoCabecera(12).Text))
                                                
                            txtDatoCabecera(2).Text = txtDatoCabecera(2).Tag
                            If CDate(Devuelve_Fecha_Posible_Parto(txtDatoCabecera(1).Text)) < CDate(txtFechaControl.Text) Then
                                If txtDatoCabecera(1).Enabled = True Then
                                    txtDatoCabecera(1).SetFocus
                                Else
                                    MsgBox "La FUM calculada es incorrecta, la FPP (" & Devuelve_Fecha_Posible_Parto(txtDatoCabecera(1).Text) & ") no puede ser menor que la fecha de control." + vbCrLf + "Revise los datos de la ecografía", vbInformation, "Validación"
                                    txtDatoCabecera(10).SetFocus
                                End If
                                Exit Sub
                            End If
                            txtDatoCabecera(2).Text = Devuelve_Fecha_Posible_Parto(txtDatoCabecera(1).Text)
                            
                            If ml_IdControl = ml_IdControlSeleccionado Then
                                lc_MensajeValidacion = ActualizaDatosControlCalculados(4, "ProMat_DevuelveEdadGestacional")
                                If lc_MensajeValidacion <> "" Then
                                    MsgBox lc_MensajeValidacion, vbInformation, "Validación"
                                    Exit Sub
                                End If
                                lc_MensajeValidacion = ActualizaDatosControlCalculados(5, "ProMat_DevuelveEdadGestacional")
                                If lc_MensajeValidacion <> "" Then
                                    MsgBox lc_MensajeValidacion, vbInformation, "Validación"
                                    Exit Sub
                                End If
                                If txtDatoControl(4).Text <> "" Then
                                    txtDatoControl(6).Text = CalculaPercentilIMC(ml_peso, ml_talla, CLng(txtDatoControl(4).Text))
                                End If
                                CargaGraficoChartSpace
                            End If
                        End If
                    Else
                        MsgBox "La fecha de ecografía no puede estar vacía", vbInformation, "Validación"
                        txtDatoCabecera(10).SetFocus
                        Exit Sub
                    End If
                End If
            End If
        
        End If
    End If
    mo_Teclado.RealizarNavegacion KeyCode, txtDatoCabecera(Index)
End Sub

Private Sub txtFechaControl_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaControl
End Sub

Private Sub txtDatoControl_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDatoControl(Index)
End Sub

Private Sub cmbDiagnosticos_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbDiagnosticos
End Sub

Private Sub txtDatoCabecera_KeyPress(Index As Integer, KeyAscii As Integer)
    oRsProCabeceraConfigDatos.MoveFirst
    Do While Not oRsProCabeceraConfigDatos.EOF
        If oRsProCabeceraConfigDatos.Fields!IdCabDato = Index Then
            If oRsProCabeceraConfigDatos.Fields!cab_tipo = "ValorEntero" Then
                If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
                    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                        KeyAscii = 0
                    End If
                End If
                Exit Do
            Else
                If oRsProCabeceraConfigDatos.Fields!cab_tipo = "ValorDouble" Then
                    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
                        If Not (mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Or KeyAscii = 46) Then
                            KeyAscii = 0
                        End If
                    End If
                    Exit Do
                End If
            End If
        End If
        oRsProCabeceraConfigDatos.MoveNext
    Loop
    
End Sub

Private Sub chkDatoCabecera_Click(Index As Integer)
    If ml_IdPrograma = 1 Then
        'Casos Especiales - Modulo Materno
        Select Case Index
        Case 7
            If chkDatoCabecera(7).Value = 1 Then
                If ml_SeteoUnaSolaVesCheckBox = False Then txtDatoCabecera(8).Text = ""
                mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(8), True
            Else
                If ml_SeteoUnaSolaVesCheckBox = False Then txtDatoCabecera(8).Text = "0"
                mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(8), False
            End If
        Case 9
            If chkDatoCabecera(9).Value = 1 Then
                If ml_SeteoUnaSolaVesCheckBox = False Then txtDatoCabecera(1).Text = txtDatoCabecera(1).Tag
                mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(1), False
                
                If ml_SeteoUnaSolaVesCheckBox = False Then txtDatoCabecera(2).Text = txtDatoCabecera(2).Tag
                mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(2), False
            
                If ml_SeteoUnaSolaVesCheckBox = False Then txtDatoCabecera(10).Text = txtDatoCabecera(10).Tag
                mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(10), True
                
                If ml_SeteoUnaSolaVesCheckBox = False Then txtDatoCabecera(11).Text = ""
                mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(11), True
                
                If ml_SeteoUnaSolaVesCheckBox = False Then txtDatoCabecera(12).Text = ""
                mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(12), True
            Else
                If ml_SeteoUnaSolaVesCheckBox = False Then txtDatoCabecera(1).Text = txtDatoCabecera(1).Tag
                mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(1), True
                
                If ml_SeteoUnaSolaVesCheckBox = False Then txtDatoCabecera(2).Text = txtDatoCabecera(2).Tag
                mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(2), False
                
                If ml_SeteoUnaSolaVesCheckBox = False Then txtDatoCabecera(10).Text = txtDatoCabecera(10).Tag
                mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(10), False
                
                If ml_SeteoUnaSolaVesCheckBox = False Then txtDatoCabecera(11).Text = ""
                mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(11), False
                
                If ml_SeteoUnaSolaVesCheckBox = False Then txtDatoCabecera(12).Text = ""
                mo_Formulario.HabilitarDeshabilitar txtDatoCabecera(12), False
            End If
        End Select
    End If
End Sub

Private Sub txtDatoCabecera_LostFocus(Index As Integer)
    Dim lc_MensajeValidacion As String
    Dim lb_EsCalculador As Boolean
    'Valida datos de cabecera segun formato
    oRsProCabeceraConfigDatos.MoveFirst
    Do While Not oRsProCabeceraConfigDatos.EOF
        If oRsProCabeceraConfigDatos.Fields!IdCabDato = Index Then
            Select Case oRsProCabeceraConfigDatos.Fields!cab_tipo
                Case "ValorFecha"
                    If txtDatoCabecera(Index).Text <> sighentidades.FECHA_VACIA_DMY Then
                        If sighentidades.EsFecha(txtDatoCabecera(Index).Text, "DD/MM/AAAA", False) = False Then
                            MsgBox "El valor de " & lblDatoCabecera(Index).Caption & " no tiene el formato correcto", vbInformation, "Validación"
                            txtDatoCabecera(Index).SetFocus
                            Exit Sub
                        End If
                    End If
                Case "ValorEntero", "ValorDouble"
                    If Not (IsNull(oRsProCabeceraConfigDatos.Fields!Cab_RangoInicial) Or oRsProCabeceraConfigDatos.Fields!Cab_RangoInicial = 0) Then
                        If txtDatoCabecera(Index).Text <> "" Then
                            If Val(txtDatoCabecera(Index).Text) < oRsProCabeceraConfigDatos.Fields!Cab_RangoInicial Then
                                MsgBox "El valor de " & lblDatoCabecera(Index).Caption & " no debe ser menor a " & oRsProCabeceraConfigDatos.Fields!Cab_RangoInicial, vbInformation, "Validación"
                                txtDatoCabecera(Index).SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                    If Not (IsNull(oRsProCabeceraConfigDatos.Fields!Cab_RangoFinal) Or oRsProCabeceraConfigDatos.Fields!Cab_RangoFinal = 0) Then
                        If txtDatoCabecera(Index).Text <> "" Then
                            If Val(txtDatoCabecera(Index).Text) > oRsProCabeceraConfigDatos.Fields!Cab_RangoFinal Then
                                MsgBox "El valor de " & lblDatoCabecera(Index).Caption & " no debe ser mayor a " & oRsProCabeceraConfigDatos.Fields!Cab_RangoFinal, vbInformation, "Validación"
                                txtDatoCabecera(Index).SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                    'CASO ESPECIAL
                    If ml_IdPrograma = 1 Then 'Modulo Materno
                        If Index = 3 Then
                            If Val(txtDatoCabecera(3).Text) < Val(txtDatoCabecera(6).Text) Then
                                MsgBox "El valor de " & lblDatoCabecera(6).Caption & " no debe ser mayor al valor de " & lblDatoCabecera(3).Caption, vbInformation, "Validación"
                                txtDatoCabecera(3).SetFocus
                                Exit Sub
                            End If
                        End If
                        If Index = 6 Then
                            If Val(txtDatoCabecera(3).Text) < Val(txtDatoCabecera(6).Text) Then
                                MsgBox "El valor de '" & lblDatoCabecera(6).Caption & "' no debe ser mayor al valor de " & lblDatoCabecera(3).Caption, vbInformation, "Validación"
                                txtDatoCabecera(6).SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
            End Select
        End If
        oRsProCabeceraConfigDatos.MoveNext
    Loop
        
            If ml_IdPrograma = 1 Then 'Modulo Materno
                If Index = 10 Or Index = 11 Or Index = 12 Then
                    If txtDatoCabecera(10).Text <> sighentidades.FECHA_VACIA_DMY Then
                        If sighentidades.EsFecha(txtDatoCabecera(10).Text, "DD/MM/AAAA", False) = True Then
                            If CDate(txtDatoCabecera(10).Text) > CDate(txtFechaControl.Text) Then
                                MsgBox "La fecha de la ecografía es incorrecta, no puede ser mayor que la fecha de control", vbInformation, "Validación"
                                txtDatoCabecera(10).SetFocus
                                Exit Sub
                            End If
'                            If Index = 11 And txtDatoCabecera(11).Text = "" Then
'                                MsgBox "Las semanas y dias de gestación del resultado de la ecografía no puedes estar vacías", vbInformation, "Validación"
'                                txtDatoCabecera(11).SetFocus
'                                Exit Sub
'                            End If
'                            If Index = 12 And txtDatoCabecera(12).Text = "" Then
'                                MsgBox "Las semanas y dias de gestación del resultado de la ecografía no puedes estar vacías", vbInformation, "Validación"
'                                txtDatoCabecera(12).SetFocus
'                                Exit Sub
'                            End If
                            
                            txtDatoCabecera(1).Text = Devuelve_Fecha_FUM_Ecografia(txtDatoCabecera(10).Text, IIf(txtDatoCabecera(11).Text = "", 0, txtDatoCabecera(11).Text), IIf(txtDatoCabecera(12).Text = "", 0, txtDatoCabecera(12).Text))
                                                
                            txtDatoCabecera(2).Text = txtDatoCabecera(2).Tag
                            If CDate(Devuelve_Fecha_Posible_Parto(txtDatoCabecera(1).Text)) < CDate(txtFechaControl.Text) Then
                                If txtDatoCabecera(1).Enabled = True Then
                                    txtDatoCabecera(1).SetFocus
                                Else
                                    If Index = 10 Then
                                        MsgBox "La FUM calculada es incorrecta, la FPP (" & Devuelve_Fecha_Posible_Parto(txtDatoCabecera(1).Text) & ") no puede ser menor que la fecha de control." + vbCrLf + "Revise los datos de la ecografía", vbInformation, "Validación"
                                        txtDatoCabecera(10).SetFocus
                                    End If
                                End If
                                Exit Sub
                            End If
                            txtDatoCabecera(2).Text = Devuelve_Fecha_Posible_Parto(txtDatoCabecera(1).Text)
                            
                            If ml_IdControl = ml_IdControlSeleccionado Then
                                lc_MensajeValidacion = ActualizaDatosControlCalculados(4, "ProMat_DevuelveEdadGestacional")
                                If lc_MensajeValidacion <> "" Then
                                    MsgBox lc_MensajeValidacion, vbInformation, "Validación"
                                    Exit Sub
                                End If
                                lc_MensajeValidacion = ActualizaDatosControlCalculados(5, "ProMat_DevuelveEdadGestacional")
                                If lc_MensajeValidacion <> "" Then
                                    MsgBox lc_MensajeValidacion, vbInformation, "Validación"
                                    Exit Sub
                                End If
                                If txtDatoControl(4).Text <> "" Then
                                    txtDatoControl(6).Text = CalculaPercentilIMC(ml_peso, ml_talla, CLng(txtDatoControl(4).Text))
                                End If
                                CargaGraficoChartSpace
                            End If
                        End If
                    End If
                End If
            End If
        
    'Consullta si el dato de cabecera es principal para el calculo de otros datos
    oRsProCabeceraConfigDatos.MoveFirst
    oRsProCabeceraConfigDatos.Find "IdCabDato=" & Index
    If Not oRsProCabeceraConfigDatos.EOF Then
        lb_EsCalculador = oRsProCabeceraConfigDatos.Fields!Cab_EsDatoCalculador
    End If
    
    If lb_EsCalculador = True Then
        'Genera datos calculado de cabecera a partir de otro dato de cabecera
        oRsProCabeceraConfigDatos.MoveFirst
        Do While Not oRsProCabeceraConfigDatos.EOF
            If oRsProCabeceraConfigDatos.Fields!Cab_EsDatoCalculado = True Then
                 Select Case oRsProCabeceraConfigDatos.Fields!Cab_FormulaCalculaValor
                     Case "ProMat_Devuelve_FPP"
                            If Index = 1 Then
                                If txtDatoCabecera(Index).Text <> sighentidades.FECHA_VACIA_DMY Then
                                    If CDate(Devuelve_Fecha_Posible_Parto(txtDatoCabecera(Index).Text)) < CDate(txtFechaControl.Text) Then
                                        MsgBox "La fecha '" & lblDatoCabecera(Index).Caption & "' es incorrecta, la FPP no puede ser menor que la fecha de control", vbInformation, "Validación"
                                        txtDatoCabecera(Index).SetFocus
                                        Exit Sub
                                    End If
                                    txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Text = Devuelve_Fecha_Posible_Parto(txtDatoCabecera(Index).Text)
                                End If
                            End If
                     Case "Registrar Calcula Formula 2 (ProCabeceraConfigDatos)"
                     Case "Registrar Calcula Formula 3 (ProCabeceraConfigDatos)"
                     Case "Registrar Calcula Formula 3 (ProCabeceraConfigDatos)"
                 End Select
            End If
            oRsProCabeceraConfigDatos.MoveNext
        Loop
        
        If ml_IdControl = ml_IdControlSeleccionado Then
            'Genera datos calculado de control a partir de un dato de cabecera
            oRsProControlConfigDatos.MoveFirst
            Do While Not oRsProControlConfigDatos.EOF
                If oRsProControlConfigDatos.Fields!Control_EsDatoCalculado Then
                    lc_MensajeValidacion = ActualizaDatosControlCalculados(oRsProControlConfigDatos.Fields!idcontroldato, oRsProControlConfigDatos.Fields!Control_FormulaCalculaValor)
                    If lc_MensajeValidacion <> "" Then
                        MsgBox lc_MensajeValidacion, vbInformation, "Validación"
                        Exit Sub
                    End If
                End If
                oRsProControlConfigDatos.MoveNext
            Loop
            CargaGraficoChartSpace
        End If
    End If
End Sub

Private Sub txtDatoControl_LostFocus(Index As Integer)
    Dim lc_MensajeValidacion As String
    oRsProControlConfigDatos.MoveFirst
    Do While Not oRsProControlConfigDatos.EOF
        If oRsProControlConfigDatos.Fields!idcontroldato = Index Then
            If oRsProControlConfigDatos.Fields!Control_Tipo = "ValorFecha" Then
                If txtDatoControl(Index).Text <> sighentidades.FECHA_VACIA_DMY Then
                    If sighentidades.EsFecha(txtDatoControl(Index).Text, "DD/MM/AAAA", False) = False Then
                        MsgBox "La fecha '" & lblDatoControl(Index).Caption & "' no tiene el formato correcto", vbInformation, "Validación"
                        txtDatoControl(Index).SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        oRsProControlConfigDatos.MoveNext
    Loop
    'Genera datos calculado de control a partir de un dato de cabecera
    oRsProControlConfigDatos.MoveFirst
    Do While Not oRsProControlConfigDatos.EOF
        If oRsProControlConfigDatos.Fields!Control_EsDatoCalculado Then
            lc_MensajeValidacion = ActualizaDatosControlCalculados(oRsProControlConfigDatos.Fields!idcontroldato, oRsProControlConfigDatos.Fields!Control_FormulaCalculaValor)
            If lc_MensajeValidacion <> "" Then
                MsgBox lc_MensajeValidacion, vbInformation, "Validación"
                Exit Sub
            End If
        End If
        oRsProControlConfigDatos.MoveNext
    Loop
End Sub

Private Sub txtDatoControl_KeyPress(Index As Integer, KeyAscii As Integer)
    oRsProControlConfigDatos.MoveFirst
    Do While Not oRsProControlConfigDatos.EOF
        If oRsProControlConfigDatos.Fields!idcontroldato = Index Then
            If oRsProControlConfigDatos.Fields!Control_Tipo = "ValorEntero" Then
                If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
                    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                        KeyAscii = 0
                    End If
                End If
                Exit Do
            End If
        End If
        oRsProControlConfigDatos.MoveNext
    Loop
End Sub

Function ValidarReglas() As Boolean
    mc_MensajeValidacion = ""
    ValidarReglas = False
    ValidarReglas = True
End Function

Function ValidarDatosObligatorios() As Boolean
    mc_MensajeValidacion = ""
    ValidarDatosObligatorios = False
    
    'Casos Especiales Modulo Materno
    If ml_IdPrograma = 1 Then
        If chkDatoCabecera(7).Value = 1 Then
            If txtDatoCabecera(8).Text = "" Then
                mc_MensajeValidacion = "El embarazo es gemelar, por favor ingrese el número de gemelos"
                TabPrograma.Tab = 0
                Exit Function
            Else
                If Val(txtDatoCabecera(8).Text) <= 1 Then
                    mc_MensajeValidacion = "El número de gemelos debe ser mayor a 1"
                    TabPrograma.Tab = 0
                    Exit Function
                End If
            End If
        End If
        If chkDatoCabecera(9).Value = 1 Then
            If txtDatoCabecera(10).Text = sighentidades.FECHA_VACIA_DMY Then
                mc_MensajeValidacion = "La fecha de ecografía es obligatorio, por favor ingreselo"
                TabPrograma.Tab = 0
                Exit Function
            Else
                If sighentidades.EsFecha(txtDatoCabecera(10).Text, "DD/MM/AAAA", False) = False Then
                    mc_MensajeValidacion = "La fecha de ecografía' no tiene el formato correcto, por favor corrija"
                    TabPrograma.Tab = 0
                    Exit Function
                End If
            End If
            If txtDatoCabecera(11).Text = "" Then
                mc_MensajeValidacion = "Las semanas del resultado de la ecografía es obligatorio, por favor ingreselo"
                TabPrograma.Tab = 0
                Exit Function
            End If
            If txtDatoCabecera(12).Text = "" Then
                mc_MensajeValidacion = "Los dias del resultado de la ecografía es obligatorio, por favor ingreselo"
                TabPrograma.Tab = 0
                Exit Function
            End If
        End If

    End If
            
    If oRsProCabeceraConfigDatos.RecordCount > 0 Then
       oRsProCabeceraConfigDatos.MoveFirst
       Do While Not oRsProCabeceraConfigDatos.EOF
            If oRsProCabeceraConfigDatos.Fields!Cab_EsDatoObligatorio Then
                If txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Text = "" Then
                    mc_MensajeValidacion = "El dato de cabecera '" & oRsProCabeceraConfigDatos.Fields!Cab_Texto & "' es obligatorio, por favor ingreselo"
                    TabPrograma.Tab = 0
                    Exit Function
                End If
                If oRsProCabeceraConfigDatos.Fields!cab_tipo = "ValorFecha" Then
                    If txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Text = sighentidades.FECHA_VACIA_DMY Then
                        mc_MensajeValidacion = "El dato de cabecera '" & oRsProCabeceraConfigDatos.Fields!Cab_Texto & "' es obligatorio, por favor ingreselo"
                        TabPrograma.Tab = 0
                        Exit Function
                    Else
                        If sighentidades.EsFecha(txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Text, "DD/MM/AAAA", False) = False Then
                            mc_MensajeValidacion = "El dato de cabecera '" & oRsProCabeceraConfigDatos.Fields!Cab_Texto & "' no tiene el formato correcto, por favor corrija"
                            TabPrograma.Tab = 0
                            Exit Function
                        End If
                    End If
                End If
            End If
            If oRsProCabeceraConfigDatos.Fields!cab_tipo = "ValorFecha" Then
                If txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Text <> sighentidades.FECHA_VACIA_DMY Then
                    If sighentidades.EsFecha(txtDatoCabecera(oRsProCabeceraConfigDatos.Fields!IdCabDato).Text, "DD/MM/AAAA", False) = False Then
                        mc_MensajeValidacion = "El dato de cabecera '" & oRsProCabeceraConfigDatos.Fields!Cab_Texto & "' no tiene el formato correcto, por favor corrija"
                        TabPrograma.Tab = 0
                        Exit Function
                    End If
                End If
            End If
            oRsProCabeceraConfigDatos.MoveNext
        Loop
    End If
    If EsControlActual Then
        If txtFechaControl.Text = sighentidades.FECHA_VACIA_DMY Then
            mc_MensajeValidacion = "El dato de control 'Fecha de Control' es obligatorio, por favor ingreselo"
            TabPrograma.Tab = 1
            Exit Function
        Else
            If sighentidades.EsFecha(txtFechaControl.Text, "DD/MM/AAAA", False) = False Then
                mc_MensajeValidacion = "El dato de control 'Fecha de Control' no tiene el formato correcto, por favor corrija"
                TabPrograma.Tab = 1
                Exit Function
            End If
        End If
        
        If oRsProControlConfigDatos.RecordCount > 0 Then
           oRsProControlConfigDatos.MoveFirst
           Do While Not oRsProControlConfigDatos.EOF
                If oRsProControlConfigDatos.Fields!Control_EsDatoObligatorio Then
                    If txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Text = "" Then
                        mc_MensajeValidacion = "El dato de control '" & oRsProControlConfigDatos.Fields!control_texto & "' es obligatorio, por favor ingreselo"
                        TabPrograma.Tab = 1
                        Exit Function
                    End If
                    If oRsProControlConfigDatos.Fields!Control_Tipo = "ValorFecha" Then
                        If txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Text = sighentidades.FECHA_VACIA_DMY Then
                            mc_MensajeValidacion = "El dato de control '" & oRsProControlConfigDatos.Fields!control_texto & "' es obligatorio, por favor ingreselo"
                            TabPrograma.Tab = 1
                            Exit Function
                        Else
                            If sighentidades.EsFecha(txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Text, "DD/MM/AAAA", False) = False Then
                                mc_MensajeValidacion = "El dato de control '" & oRsProControlConfigDatos.Fields!control_texto & "' no tiene el formato correcto, por favor corrija"
                                TabPrograma.Tab = 1
                                Exit Function
                            End If
                        End If
                    End If
                End If
                If oRsProControlConfigDatos.Fields!Control_Tipo = "ValorFecha" Then
                    If txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Text <> sighentidades.FECHA_VACIA_DMY Then
                        If sighentidades.EsFecha(txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Text, "DD/MM/AAAA", False) = False Then
                            mc_MensajeValidacion = "El dato de control '" & oRsProControlConfigDatos.Fields!control_texto & "' no tiene el formato correcto, por favor corrija"
                            TabPrograma.Tab = 1
                            Exit Function
                        End If
                    End If
                End If
                oRsProControlConfigDatos.MoveNext
            Loop
        End If
        'Valida diagnosticos
        If oRsProCatalogoDiagnosticos.RecordCount = 0 Then
             mc_MensajeValidacion = "Debe ingresar al menos un Diagnóstico del programa"
             Exit Function
        End If
                   
        Dim lbExistePrincipal As Boolean
        lbExistePrincipal = False
        If oRsProCatalogoDiagnosticos.RecordCount > 0 Then
            oRsProCatalogoDiagnosticos.MoveFirst
            Do While Not oRsProCatalogoDiagnosticos.EOF
               If oRsProCatalogoDiagnosticos.Fields!Principal Then
                    lbExistePrincipal = True
               End If
               oRsProCatalogoDiagnosticos.MoveNext
            Loop
            oRsProCatalogoDiagnosticos.MoveFirst
            
            If lbExistePrincipal = False Then
                mc_MensajeValidacion = "Debe tener por lo menos un Diagnóstico principal"
                TabPrograma.Tab = 1
                Exit Function
            End If
        End If
    End If
    
    ValidarDatosObligatorios = True
End Function


Sub Actualiza_Peso_Presion(ml_peso_triaje As Double, ml_presion_triaje As String, ml_Talla_triaje As Long)
    Dim lc_MensajeValidacion As String
    ml_peso = ml_peso_triaje
    ml_presion = ml_presion_triaje
    ml_talla = ml_Talla_triaje
    If ml_IdControl = ml_IdControlSeleccionado Then
        oRsProControlConfigDatos.MoveFirst
        Do While Not oRsProControlConfigDatos.EOF
            If oRsProControlConfigDatos.Fields!Control_EsPeso Then txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Text = ml_peso
            If oRsProControlConfigDatos.Fields!Control_EsPresion Then txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Text = ml_presion
            If oRsProControlConfigDatos.Fields!Control_EsTalla Then txtDatoControl(oRsProControlConfigDatos.Fields!idcontroldato).Text = ml_talla
            If oRsProControlConfigDatos.Fields!Control_EsDatoCalculado Then
                    lc_MensajeValidacion = ActualizaDatosControlCalculados(oRsProControlConfigDatos.Fields!idcontroldato, oRsProControlConfigDatos.Fields!Control_FormulaCalculaValor)
                    If lc_MensajeValidacion <> "" Then
                        MsgBox lc_MensajeValidacion, vbInformation, "Validación"
                        Exit Sub
                    End If
            End If
            oRsProControlConfigDatos.MoveNext
        Loop
        CargaGraficoChartSpace
    End If
End Sub

Sub CargaGraficoChartSpace()
    Dim lnFor As Integer
    Dim lcTituloGrafico As String
    Dim lnNumCtrlsAnt As Integer
    Dim lnIdCtrlGraficoX As Long
    Dim lnColorLinea As Integer
    Dim oRsTmpGraficoX As ADODB.Recordset
    Dim oRsTmpGraficoY As ADODB.Recordset
    Dim oRsTmp1 As ADODB.Recordset
    Dim oRsTmp2 As ADODB.Recordset
    
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oRsTmpGraficoX = mo_reglasComunes.ProControlConfigGrafico(ml_IdPrograma, True, oConexion)
    Set oRsTmpGraficoY = mo_reglasComunes.ProControlConfigGrafico(ml_IdPrograma, False, oConexion)
    'Validacion obligatoria para cargar el grafico
    If oRsTmpGraficoX.RecordCount = 0 Then Exit Sub
    If oRsTmpGraficoX.RecordCount > 1 Then
        MsgBox "Se configuro mas de una coordenada X, por favor revise la tabla ProControlConfigDatos", vbInformation, "Validación"
    End If
    If oRsTmpGraficoY.RecordCount = 0 Then Exit Sub
              
    lcTituloGrafico = ""
    oRsTmpGraficoX.MoveFirst
    Do While Not oRsTmpGraficoX.EOF
       lnIdCtrlGraficoX = oRsTmpGraficoX.Fields!idcontroldato
       lcTituloGrafico = lcTituloGrafico & oRsTmpGraficoX.Fields!control_texto & " (X), "
       oRsTmpGraficoX.MoveNext
    Loop
    lnColorLinea = 1
    oRsTmpGraficoY.MoveFirst
    Do While Not oRsTmpGraficoY.EOF
       lcTituloGrafico = lcTituloGrafico & oRsTmpGraficoY.Fields!control_texto & Devuelve_Texto_ColorLinea(lnColorLinea) & ",  "
       lnColorLinea = lnColorLinea + 1
       oRsTmpGraficoY.MoveNext
    Loop
               
    ChartSpace1.Clear
    ChartSpace1.DisplayToolbar = False
    Set owcChart = ChartSpace1.Charts.Add
    owcChart.HasTitle = True
    owcChart.Title.Caption = lcTituloGrafico ' "Edad en Semanas (X), PT(rojo),  TE(verde),  PE(amarillo)"
    owcChart.Title.Font.Name = "Arial Narrow"
    owcChart.Title.Font.Size = 8
    owcChart.Title.Font.Color = vbBlue
    owcChart.Axes(chAxisPositionBottom).Font.Name = "Arial narrow"
    owcChart.Axes(chAxisPositionBottom).Font.Size = 8
    owcChart.Axes(chAxisPositionBottom).Font.Color = vbBlue
    owcChart.Axes(chAxisPositionBottom).Scaling.Minimum = 0
    owcChart.Axes(chAxisPositionLeft).Font.Name = "Arial narrow"
    owcChart.Axes(chAxisPositionLeft).Font.Size = "8"
    owcChart.Axes(chAxisPositionLeft).Font.Color = vbBlue
    owcChart.Axes(chAxisPositionLeft).Scaling.Minimum = 0
    owcChart.Axes(chAxisPositionLeft).Scaling.Maximum = 110
    owcChart.Axes(1).HasTitle = 1
    owcChart.Axes(1).Font.Name = "Arial Narrow"
    owcChart.Axes(1).Font.Size = 8
    owcChart.Axes(1).Font.Color = vbBlue
    owcChart.Axes(1).Title.Caption = "Percentil"
    owcChart.Axes(1).Title.Font.Name = "Arial Narrow"
    owcChart.Axes(1).Title.Font.Size = 8
    owcChart.Axes(1).Title.Font.Color = vbBlue
    
    If ml_IdProCabecera = 0 And txtDatoControl(lnIdCtrlGraficoX).Text = "" Then Exit Sub
      
    If ml_IdProCabecera = 0 Then
       lnNumCtrlsAnt = 0
    Else 'ProControlConfigGrafico
        'Consulta total controles anteriores
        Set oRsTmp1 = mo_reglasComunes.ProSeleccionarControlesAnteriores(ml_IdPrograma, ml_IdProCabecera, ml_IdControlSeleccionado, oConexion)
        If ml_IdControl = ml_IdControlSeleccionado Then
            If mb_ConsultaControl Then
                lnNumCtrlsAnt = oRsTmp1.RecordCount - 1
            Else
                lnNumCtrlsAnt = oRsTmp1.RecordCount
            End If
        Else
            lnNumCtrlsAnt = oRsTmp1.RecordCount - 1
        End If
    End If
    lnNroPuntosGraficos = lnNumCtrlsAnt + 1
    
    lnColorLinea = 1
    oRsTmpGraficoY.MoveFirst
    Do While Not oRsTmpGraficoY.EOF
'       lcTituloGrafico = Devuelve_Texto_ColorLinea(lnColorLinea) & ",  "
        xValues = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250, 280)
        yValues = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250, 280)
        ReDim xValues(lnNumCtrlsAnt)
        
        If ml_IdProCabecera > 0 Then
            If oRsTmp1.RecordCount > 0 Then
                oRsTmp1.MoveLast 'ProHistoricoDatoControlGrafico
                For lnFor = (oRsTmp1.RecordCount - 1) To 0 Step -1
                    Set oRsTmp2 = mo_reglasComunes.ProHistoricoDatoControlGrafico(ml_IdPrograma, ml_IdProCabecera, oRsTmp1.Fields!IdControl, lnIdCtrlGraficoX, oRsTmpGraficoY.Fields!idcontroldato, oConexion)
                    If oRsTmp2.RecordCount > 0 Then
                        oRsTmp2.MoveFirst
                        Do While Not oRsTmp2.EOF
                            If oRsTmp2.Fields!control_esgraficoejex = True Then xValues(lnFor) = oRsTmp2.Fields!controldato
                            If oRsTmp2.Fields!control_esgraficoejex = False Then yValues(lnFor) = oRsTmp2.Fields!controldato
                            oRsTmp2.MoveNext
                        Loop
                    End If
                    oRsTmp1.MovePrevious
                Next
            End If
        End If
        txtDatoControl(lnIdCtrlGraficoX).Refresh
        xValues(lnNumCtrlsAnt) = txtDatoControl(lnIdCtrlGraficoX).Text
        yValues(lnNumCtrlsAnt) = txtDatoControl(oRsTmpGraficoY.Fields!idcontroldato).Text
        
        Set owcSeries = owcChart.SeriesCollection.Add
        With owcSeries
            .Caption = ""
            .SetData chDimCategories, chDataLiteral, xValues
            .SetData chDimValues, chDataLiteral, yValues
            .Type = chChartTypeLineMarkers
            .Line.Color = vbRed
            If lnColorLinea = 1 Then .Line.Color = vbRed
            If lnColorLinea = 2 Then .Line.Color = vbGreen
            If lnColorLinea = 3 Then .Line.Color = vbYellow
            If lnColorLinea = 4 Then .Line.Color = vbBlue
            .Line.Color = vbRed
            .Line.Weight = 3
            .Marker.Style = chMarkerStyleCircle
            .Line.DashStyle = chLineSolid
            .DataLabelsCollection.Add
        End With
        
        lnColorLinea = lnColorLinea + 1
        oRsTmpGraficoY.MoveNext
    Loop
    
    oConexion.Close
    Set oConexion = Nothing
End Sub

Public Function Devuelve_Texto_ColorLinea(lnIdColor As Integer) As String
    Devuelve_Texto_ColorLinea = "(Rojo)"
    Select Case lnIdColor
        Case 1
            Devuelve_Texto_ColorLinea = "(rojo)"
        Case 2
            Devuelve_Texto_ColorLinea = "(verde)"
        Case 3
            Devuelve_Texto_ColorLinea = "(amarillo)"
        Case 4
            Devuelve_Texto_ColorLinea = "(azul)"
    End Select
End Function

'-------------------------------------------------------------------------------------
'               Funciones Personalidades del modulo materno
'-------------------------------------------------------------------------------------

Function Devuelve_Fecha_Posible_Parto(ldFechaFUM As Date) As Date
    Dim lnDiaFUM As Integer: Dim lnMesFUM As Integer: Dim lnAnioFUM As Integer
   
    lnDiaFUM = Day(ldFechaFUM)
    lnMesFUM = Month(ldFechaFUM)
    lnAnioFUM = Year(ldFechaFUM)
    
    If lnMesFUM <= 3 Then
        lnMesFUM = lnMesFUM + 9
    Else
        lnMesFUM = lnMesFUM - 3
        lnAnioFUM = lnAnioFUM + 1
    End If
    If lnDiaFUM > 28 Then
        Devuelve_Fecha_Posible_Parto = 10 + CDate(lnDiaFUM - 3 & "/" & lnMesFUM & "/" & lnAnioFUM)
    Else
        Devuelve_Fecha_Posible_Parto = 7 + CDate(lnDiaFUM & "/" & lnMesFUM & "/" & lnAnioFUM)
    End If
End Function

Function Devuelve_Fecha_FUM_Ecografia(ldFechaEcografia As Date, lnSemanasEcografia As Integer, lnDiasEcoGrafia As Integer) As Date
    Devuelve_Fecha_FUM_Ecografia = ldFechaEcografia - lnSemanasEcografia * 7 - lnDiasEcoGrafia
End Function

'Actualiza valores y Devuelve percentil IMC de la ATENCION ACTUAL DEL PACIENTE
Function CalculaPercentilIMC(lnPesoKg As Double, lnTallaCM As Long, lnEdadGestacional As Long) As Long
    CalculaPercentilIMC = 0
    If lnPesoKg > 0 And lnTallaCM > 0 And lnEdadGestacional > 0 Then
       On Error Resume Next
       Dim EXL As Excel.Application
       Set EXL = New Excel.Application
       Dim W As Excel.Workbook
       Set W = EXL.Workbooks.Open(App.Path & "\Plantillas\materno Percentiles.xls")

       Dim TallaM As Double 'Pasar la talla a metros
       TallaM = lnTallaCM / 100

       Dim s As Excel.Worksheet
       Set s = W.Sheets("IMC")
       lnPercentilPE = lnPercentilNull
       s.Cells(203, 6).Value = lnPesoKg
       s.Cells(205, 6).Value = TallaM
       s.Cells(209, 6).Value = lnEdadGestacional
       CalculaPercentilIMC = s.Cells(211, 6).Value

       W.Close False
       Set s = Nothing
       Set W = Nothing
       Set EXL = Nothing
    End If
End Function

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

Private Sub txtFechaControl_LostFocus()
    If Not EsFecha(txtFechaControl.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaControl.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Public Function DevuelveFechaControl() As String
    DevuelveFechaControl = txtFechaControl.Text
End Function

Public Function DevuelveNumeroControl() As String
    DevuelveNumeroControl = FrameControl.Caption
End Function



Private Sub txtFiltraCPT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        Dim oRsTmp1 As New Recordset
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        
        Set oRsTmp1 = mo_reglasComunes.ProCatalogoCptSeleccionarPorIdPrograma(ml_IdPrograma, oConexion)
        If Trim(txtFiltraCPT.Text) <> "" Then oRsTmp1.Filter = "PROCEDIMIENTO like '%" & txtFiltraCPT.Text & "%'"
        If Not oRsTmp1.EOF Then
            If oRsTmp1.RecordCount > 0 Then oRsTmp1.MoveFirst
        End If
        
        Set mo_cmbCptPrograma.MiComboBox = cmbProcPrograma
        mo_cmbCptPrograma.BoundColumn = "IdProducto"
        mo_cmbCptPrograma.ListField = "PROCEDIMIENTO"
        Set mo_cmbCptPrograma.RowSource = oRsTmp1
        cmbProcPrograma.SetFocus
        
        Set oRsTmp1 = Nothing
        oConexion.Close
        Set oConexion = Nothing
    Case vbKeyF11
'        BusquedaDx Trim(txtFiltroDx.Text)
    End Select
End Sub



Private Sub txtFiltroDx_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        Dim oRsTmp1 As New Recordset
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        
        Set oRsTmp1 = mo_reglasComunes.ProCatalogoDiagnosticosSeleccionarPorIdPrograma(ml_IdPrograma, oConexion)
        If Trim(txtFiltroDx.Text) <> "" Then oRsTmp1.Filter = "DIAGNOSTICO like '%" & txtFiltroDx.Text & "%'"
        If Not oRsTmp1.EOF Then
            If oRsTmp1.RecordCount > 0 Then oRsTmp1.MoveFirst
        End If
        
        Set mo_cmbDxPrograma.MiComboBox = cmbDxPrograma
        mo_cmbDxPrograma.BoundColumn = "IdDiagnostico"
        mo_cmbDxPrograma.ListField = "DIAGNOSTICO"
        Set mo_cmbDxPrograma.RowSource = oRsTmp1
        cmbDxPrograma.SetFocus
        
        Set oRsTmp1 = Nothing
        oConexion.Close
        Set oConexion = Nothing
    Case vbKeyF11
        BusquedaDx Trim(txtFiltroDx.Text)
    End Select
End Sub

Private Sub cmbDxPrograma_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmbLabHisDx.SetFocus
End Sub


Public Function DevuelveProHistorialControles() As ADODB.Recordset
    Dim oRsProHistorialControles As New Recordset
    If oRsProHistorialControles.State = 1 Then
       Set oRsProHistorialControles = Nothing
    End If
    With oRsProHistorialControles
          .Fields.Append "idcontrol", adInteger, 4, adFldIsNullable
          .Fields.Append "descripcion", adVarChar, 100, adFldIsNullable
          .Fields.Append "FechaControl", adVarChar, 100, adFldIsNullable
          .Fields.Append "ControlOtroEESS", adBoolean
          .Fields.Append "IdEstablecimiento", adInteger, 0, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    UserControl.grdHistoricoControles.Update
    If oRsProCatalogoControles.RecordCount > 0 Then
        oRsProCatalogoControles.MoveFirst
        Do While Not oRsProCatalogoControles.EOF
           oRsProHistorialControles.AddNew
           oRsProHistorialControles.Fields!IdControl = oRsProCatalogoControles.Fields!IdControl
           oRsProHistorialControles.Fields!FechaControl = oRsProCatalogoControles.Fields!FechaControl
           oRsProHistorialControles.Fields!ControlOtroEESS = oRsProCatalogoControles.Fields!ControlOtroEESS
           oRsProHistorialControles.Fields!IdEstablecimiento = oRsProCatalogoControles.Fields!IdEstablecimiento
           oRsProHistorialControles.Update
           oRsProCatalogoControles.MoveNext
        Loop
        oRsProCatalogoControles.MoveFirst
        If oRsProHistorialControles.RecordCount > 0 Then
            oRsProHistorialControles.MoveFirst
        End If
      End If
Set DevuelveProHistorialControles = oRsProHistorialControles
End Function

Private Function AsignarListaDeLabsEnGridaDiagnosticos(oGrilla As SSUltraGrid, cNombreColumna As String) As Boolean
On Error GoTo miError
               
    With oGrilla.ValueLists.Add("ListaLab").ValueListItems
           If mo_RsLabHis.RecordCount > 0 Then
              mo_RsLabHis.MoveFirst
              Do While Not mo_RsLabHis.EOF
                 .Add Right(Trim((mo_RsLabHis.Fields!valores)), 3), Trim(mo_RsLabHis.Fields!valores) '& "(" & Trim(mo_RsLabHis.Fields!descripcio) & ")"
                 mo_RsLabHis.MoveNext
              Loop
           End If
    End With
    oGrilla.Bands(0).Columns(cNombreColumna).ValueList = "ListaLab"
    AsignarListaDeLabsEnGridaDiagnosticos = True
miError:
    If Err > 0 Then
        'On Error Resume Next
        If Err <> 31101 Then MsgBox Err.Description & " : " & Err.Description, vbInformation, "Módulo Materno"
    End If
End Function

Private Function getIdPuntoCargaHospitalizacion(lIdServicioPaciente As Long) As Long
    Dim oRsTmp As New ADODB.Recordset
    Set oRsTmp = mo_reglasComunes.FactPuntosCargaSeleccionarPorFiltro("idServicio=" & Trim(Str(lIdServicioPaciente)))
    If oRsTmp.RecordCount > 0 Then
       getIdPuntoCargaHospitalizacion = oRsTmp.Fields!idPuntoCarga
    Else
       getIdPuntoCargaHospitalizacion = 9999
    End If
End Function

Private Sub txtFiltroInsumos_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        Dim oRsTmp1 As New Recordset
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        
        Set oRsTmp1 = mo_reglasComunes.ProCatalogoTratamientosSeleccionarPorIdPrograma(ml_IdPrograma, oConexion)
        If Trim(txtFiltroInsumos.Text) <> "" Then oRsTmp1.Filter = "Tratamientos like '%" & txtFiltroInsumos.Text & "%'"
        If Not oRsTmp1.EOF Then
            If oRsTmp1.RecordCount > 0 Then oRsTmp1.MoveFirst
        End If
        
        Set mo_cmbTtoPrograma.MiComboBox = cmbTtoProgama
        mo_cmbTtoPrograma.BoundColumn = "IdProducto"
        mo_cmbTtoPrograma.ListField = "Tratamientos"
        Set mo_cmbTtoPrograma.RowSource = oRsTmp1
        cmbTtoProgama.SetFocus
        
        Set oRsTmp1 = Nothing
        oConexion.Close
        Set oConexion = Nothing
    Case vbKeyF11
'        BusquedaDx Trim(txtFiltroDx.Text)
    End Select

End Sub

Private Function ValidarDiagnosticosMaterno(lIdDiagnostico As Long, cLabHIS As String, _
            Optional EditLab As Boolean = False, Optional ByVal oRegistroActual As Variant = 0) As Boolean
    On Error Resume Next
    ValidarDiagnosticosMaterno = False
    Dim oRegistroExplorar As Variant
    Dim lo_RsDxMaterno As ADODB.Recordset
    
    Set lo_RsDxMaterno = oRsProCatalogoDiagnosticos.Clone
    
    With lo_RsDxMaterno
        If Not (.BOF = True And .EOF = True) Then
            If .RecordCount > 0 Then
                .MoveFirst
                While .EOF = False
                    If Not (.CompareBookmarks(oRegistroActual, .Bookmark) = adCompareEqual) Or EditLab = False Then
                        If .Fields!idDiagnostico = lIdDiagnostico And Trim(cLabHIS) = Trim(IIf(IsNull(.Fields!labConfHIS), Space(3), .Fields!labConfHIS)) Then
                            If EditLab = True Then
                                .Bookmark = oRegistroActual
                            End If
                            MsgBox "DX ya ha sido agregados", vbInformation, "Modulo Materno"
                            Exit Function
                        End If
                    End If
                    .MoveNext
                Wend
            End If
        End If
    End With
    
    If Trim(cLabHIS) <> "" Then
        'mgaray201411a
        If mo_reglasComunes.existeCodigoLabHis(mo_RsLabHis, cLabHIS) = False Then
            MsgBox "Código LAB No Valido", vbInformation, "Módulo Niño Sano"
            Exit Function
        End If
    End If
    ValidarDiagnosticosMaterno = True
End Function


Private Sub grdDiagnosticos_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Cell.Column.Key = "labConfHIS" Then
        If ValidarDiagnosticosMaterno(Cell.Row.Cells("IdDiagnostico").Value, CStr(IIf(IsNull(NewValue), Space(3), NewValue)), True, oRsProCatalogoDiagnosticos.Bookmark) = False Then
            Cancel = True
        End If
    End If
End Sub
