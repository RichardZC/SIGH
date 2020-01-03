VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{22ACD161-99EB-11D2-9BB3-00400561D975}#1.0#0"; "PVCALE~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CFFE0A60-8E3A-11D3-BCC0-00104B9E0792}#1.0#0"; "ssInput1.ocx"
Object = "{0002E558-0000-0000-C000-000000000046}#1.1#0"; "OWC11.DLL"
Begin VB.UserControl ucPerinatal 
   BackColor       =   &H00FF8080&
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11685
   ScaleHeight     =   7125
   ScaleWidth      =   11685
   Begin TabDlg.SSTab STabPerinatal 
      Height          =   7125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12568
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Diag/Cpt/TTO."
      TabPicture(0)   =   "ucPerinatal.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraDxMorbilidad"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ChartSpace1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraCred"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FraCred1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FraMedicamentos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FraDxDesarrollo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FraInmunizaciones"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FraOtrosCpt"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Graficos Peso/Talla"
      TabPicture(1)   =   "ucPerinatal.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "stabGraficos"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Desarrollo Psicomotor"
      TabPicture(2)   =   "ucPerinatal.ctx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frAtenInteDesarrollo"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Inmunizaciones"
      TabPicture(3)   =   "ucPerinatal.ctx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frAtenInteInmunizaciones"
      Tab(3).Control(1)=   "btnImprime"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Suplemento"
      TabPicture(4)   =   "ucPerinatal.ctx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "frAtenInteSuplemento"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Tamizaje"
      TabPicture(5)   =   "ucPerinatal.ctx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "frAtenInteTamizajes"
      Tab(5).ControlCount=   1
      Begin TabDlg.SSTab stabGraficos 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   89
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   11668
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Peso para la Edad"
         TabPicture(0)   =   "ucPerinatal.ctx":00A8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frCrecimiento"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Longitud/Estatura para la Edad"
         TabPicture(1)   =   "ucPerinatal.ctx":00C4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frGraficoTalla"
         Tab(1).ControlCount=   1
         Begin VB.Frame frGraficoTalla 
            BorderStyle     =   0  'None
            Height          =   6120
            Left            =   -74880
            TabIndex        =   91
            Top             =   360
            Width           =   11175
            Begin OWC11.ChartSpace shaTallaEdad 
               Height          =   5985
               Left            =   0
               OleObjectBlob   =   "ucPerinatal.ctx":00E0
               TabIndex        =   92
               Top             =   0
               Width           =   11175
            End
         End
         Begin VB.Frame frCrecimiento 
            BorderStyle     =   0  'None
            Height          =   6135
            Left            =   120
            TabIndex        =   90
            Top             =   360
            Width           =   11175
            Begin OWC11.ChartSpace shaPesoEdad 
               Height          =   5985
               Left            =   0
               OleObjectBlob   =   "ucPerinatal.ctx":0CD8
               TabIndex        =   93
               Top             =   0
               Width           =   11175
            End
         End
      End
      Begin VB.CommandButton btnImprime 
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
         Left            =   -69600
         Picture         =   "ucPerinatal.ctx":18D0
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Imprimir Calendario de Inmunizaciones"
         Top             =   6480
         Width           =   1005
      End
      Begin VB.Frame frAtenInteTamizajes 
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   -74760
         TabIndex        =   81
         Top             =   360
         Width           =   11295
         Begin UltraGrid.SSUltraGrid grdPlanTamizajes 
            Height          =   3870
            Left            =   0
            TabIndex        =   82
            Top             =   2640
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   6826
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
            Caption         =   "grdPlanTamizajes"
         End
         Begin UltraGrid.SSUltraGrid grdPlanTamizajesPendientes 
            Height          =   2430
            Left            =   0
            TabIndex        =   83
            Top             =   0
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   4286
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
            Caption         =   "grdPlanTamizajesPendientes"
         End
      End
      Begin VB.Frame frAtenInteSuplemento 
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   -74760
         TabIndex        =   78
         Top             =   360
         Width           =   11295
         Begin UltraGrid.SSUltraGrid grdPlanSuplemento 
            Height          =   3870
            Left            =   0
            TabIndex        =   79
            Top             =   2640
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   6826
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
            Caption         =   "Plan  de Suplemento Nutricional"
         End
         Begin UltraGrid.SSUltraGrid grdPlanSuplementoPendientes 
            Height          =   2430
            Left            =   0
            TabIndex        =   80
            Top             =   0
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   4286
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
            Caption         =   "Suplemento Nutricional "
         End
      End
      Begin VB.Frame FraOtrosCpt 
         Caption         =   "Otros Procedimientos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   30
         TabIndex        =   5
         Top             =   360
         Width           =   6945
         Begin ActiveInput.SSComboBoxEx cmbProcedimientosFrecuentes 
            Height          =   345
            Left            =   60
            TabIndex        =   8
            Top             =   240
            Width           =   5550
            _ExtentX        =   9790
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
            Separator       =   " | "
         End
         Begin VB.CommandButton btnBusquedaProcedimientos 
            Caption         =   "..."
            Height          =   315
            Left            =   6540
            TabIndex        =   88
            TabStop         =   0   'False
            ToolTipText     =   "Busca Dx"
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton btnAgregarProcedminiento 
            DisabledPicture =   "ucPerinatal.ctx":1DA9
            DownPicture     =   "ucPerinatal.ctx":2192
            Height          =   315
            Left            =   5640
            Picture         =   "ucPerinatal.ctx":259E
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   390
         End
         Begin VB.CommandButton btnQuitaOtrosProcedimientos 
            DisabledPicture =   "ucPerinatal.ctx":29AA
            DownPicture     =   "ucPerinatal.ctx":2D35
            Height          =   315
            Left            =   6090
            Picture         =   "ucPerinatal.ctx":30C8
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Elimina todos los CPT"
            Top             =   240
            Width           =   405
         End
         Begin UltraGrid.SSUltraGrid grdCptFrecuentes 
            Height          =   1275
            Left            =   30
            TabIndex        =   7
            Top             =   540
            Width           =   6870
            _ExtentX        =   12118
            _ExtentY        =   2249
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   68157460
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Override        =   "ucPerinatal.ctx":3459
            Caption         =   "SSUltraGrid1"
         End
         Begin PVCOMBOLibCtl.PVComboBox cmbLabHisProcFrecuentes 
            Height          =   330
            Left            =   5640
            TabIndex        =   9
            Top             =   240
            Visible         =   0   'False
            Width           =   1065
            _Version        =   524288
            _cx             =   1879
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
      Begin VB.Frame frAtenInteInmunizaciones 
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   -74760
         TabIndex        =   58
         Top             =   360
         Width           =   11295
         Begin UltraGrid.SSUltraGrid grdPlanInmunizaciones 
            Height          =   3510
            Left            =   0
            TabIndex        =   60
            Top             =   2520
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   6191
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
            Caption         =   "Inmunizaciones Programadas"
         End
         Begin UltraGrid.SSUltraGrid grdPlanInmunizacionesPendientes 
            Height          =   2430
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   4286
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
            Caption         =   "Inmunizaciones Por Ejecutar"
         End
      End
      Begin VB.Frame frAtenInteDesarrollo 
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   -74760
         TabIndex        =   50
         Top             =   360
         Width           =   11175
         Begin VB.CommandButton btnActualizarSesionesPendientes 
            Caption         =   "Actualizar Sesiones"
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
            Left            =   5400
            TabIndex        =   85
            Top             =   6120
            Width           =   1335
         End
         Begin VB.CommandButton btnBuscaHistoricos 
            Caption         =   "Ver Plan"
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
            Left            =   3960
            Picture         =   "ucPerinatal.ctx":34AF
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   6120
            Width           =   1245
         End
         Begin VB.Frame frAtencionDesarrollo 
            Caption         =   "Sesión 1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   615
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   11010
            Begin VB.TextBox txtIdAtencionDesarrollo 
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
               Left            =   4200
               TabIndex        =   77
               ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
               Top             =   240
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.TextBox txtFechaProgramadaDesarrollo 
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
               Left            =   3000
               TabIndex        =   52
               ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
               Top             =   240
               Width           =   1125
            End
            Begin VB.TextBox txtEvalucionDesarrollo 
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
               Left            =   9600
               TabIndex        =   55
               ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
               Top             =   240
               Width           =   1125
            End
            Begin MSMask.MaskEdBox mskFechaEjecucionDes 
               Height          =   315
               Left            =   6840
               TabIndex        =   53
               Top             =   240
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Programada"
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
               Left            =   1440
               TabIndex        =   76
               Top             =   240
               Width           =   1500
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Evaluacion"
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
               Left            =   8640
               TabIndex        =   75
               Top             =   240
               Width           =   840
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Ejecución"
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
               Left            =   5400
               TabIndex        =   74
               Top             =   240
               Width           =   1320
            End
         End
         Begin UltraGrid.SSUltraGrid grdPlanDesarrollo 
            Height          =   3150
            Left            =   0
            TabIndex        =   57
            Top             =   2880
            Width           =   11010
            _ExtentX        =   19420
            _ExtentY        =   5556
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
            Caption         =   "Plan de Desarrollo Psicomotor"
         End
         Begin UltraGrid.SSUltraGrid grdPlanDesarrolloPendientes 
            Height          =   2070
            Left            =   0
            TabIndex        =   56
            Top             =   720
            Width           =   11010
            _ExtentX        =   19420
            _ExtentY        =   3651
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
            Caption         =   "Items a Evaluar"
         End
         Begin PVATLCALENDARLib.PVCalendar PVCFechaEjecDes 
            Height          =   2535
            Left            =   0
            TabIndex        =   54
            Top             =   4080
            Visible         =   0   'False
            Width           =   3015
            _Version        =   524288
            BorderStyle     =   0
            Appearance      =   1
            FirstDay        =   0
            Frame           =   3
            SelectMode      =   1
            DisplayFormat   =   0
            DateOrientation =   0
            CustomTextOrientation=   2
            ImageOrientation=   8
            DOWText0        =   "Sun"
            DOWText1        =   ""
            DOWText2        =   ""
            DOWText3        =   ""
            DOWText4        =   ""
            DOWText5        =   ""
            DOWText6        =   ""
            MonthText0      =   "January"
            MonthText1      =   "February"
            MonthText2      =   "March"
            MonthText3      =   "April"
            MonthText4      =   "May"
            MonthText5      =   "June"
            MonthText6      =   "July"
            MonthText7      =   "August"
            MonthText8      =   "September"
            MonthText9      =   "October"
            MonthText10     =   "November"
            MonthText11     =   "December"
            HeaderBackColor =   14215660
            HeaderForeColor =   0
            DisplayBackColor=   14215660
            DisplayForeColor=   0
            DayBackColor    =   14215660
            DayForeColor    =   0
            SelectedDayForeColor=   16777215
            SelectedDayBackColor=   12937777
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DOWFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DaysFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MultiLineText   =   -1  'True
            EditMode        =   0
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FraInmunizaciones 
         Caption         =   "Inmunizaciones"
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
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   6945
         Begin VB.CommandButton btnQuitarInmunizacion 
            DisabledPicture =   "ucPerinatal.ctx":3A39
            DownPicture     =   "ucPerinatal.ctx":3DC4
            Height          =   315
            Left            =   6420
            Picture         =   "ucPerinatal.ctx":4157
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Elimina todos los CPT"
            Top             =   120
            Width           =   435
         End
         Begin UltraGrid.SSUltraGrid grdInmunizaciones 
            Height          =   1005
            Left            =   30
            TabIndex        =   3
            Top             =   420
            Width           =   6870
            _ExtentX        =   12118
            _ExtentY        =   1773
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   68157460
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Override        =   "ucPerinatal.ctx":44E8
            Caption         =   "grdInmunizaciones"
         End
         Begin ActiveInput.SSComboBoxEx cmbEligeInmunizacion 
            Height          =   345
            Left            =   1470
            TabIndex        =   4
            Top             =   120
            Width           =   4950
            _ExtentX        =   8731
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
      End
      Begin VB.Frame FraDxDesarrollo 
         Caption         =   "Dx-Crecimiento Y Desarrollo"
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
         Left            =   30
         TabIndex        =   11
         Top             =   2190
         Width           =   6945
         Begin VB.CommandButton btnBusquedaDxDesarrollo 
            Caption         =   "..."
            Height          =   315
            Left            =   6540
            TabIndex        =   87
            TabStop         =   0   'False
            ToolTipText     =   "Busca Dx"
            Top             =   150
            Width           =   315
         End
         Begin ActiveInput.SSComboBoxEx cmbDxDesarrollo 
            Height          =   345
            Left            =   120
            TabIndex        =   14
            Top             =   180
            Width           =   5490
            _ExtentX        =   9684
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
            Separator       =   " | "
         End
         Begin VB.CommandButton btnAgregarDxCRED 
            DisabledPicture =   "ucPerinatal.ctx":453E
            DownPicture     =   "ucPerinatal.ctx":4927
            Height          =   315
            Left            =   5640
            Picture         =   "ucPerinatal.ctx":4D33
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   150
            Width           =   390
         End
         Begin VB.CommandButton btnQuitaDxDesarrollo 
            DisabledPicture =   "ucPerinatal.ctx":513F
            DownPicture     =   "ucPerinatal.ctx":54CA
            Height          =   315
            Left            =   6090
            Picture         =   "ucPerinatal.ctx":585D
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Elimina todos los Dx"
            Top             =   150
            Width           =   375
         End
         Begin UltraGrid.SSUltraGrid grdMorbilidadDesarollo 
            Height          =   1065
            Left            =   0
            TabIndex        =   13
            Top             =   450
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   1879
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   68157460
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Override        =   "ucPerinatal.ctx":5BEE
            Caption         =   "grdMorbilidadDesarollo"
         End
         Begin PVCOMBOLibCtl.PVComboBox cmbLabHisDxDesarrollo 
            Height          =   330
            Left            =   5640
            TabIndex        =   15
            Top             =   180
            Visible         =   0   'False
            Width           =   1065
            _Version        =   524288
            _cx             =   1879
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
      Begin VB.Frame FraMedicamentos 
         Caption         =   "Tratamiento recibido"
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
         Left            =   60
         TabIndex        =   48
         Top             =   5520
         Width           =   6870
         Begin UltraGrid.SSUltraGrid grdMedicamentos 
            Height          =   1245
            Left            =   30
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   240
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   2196
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            RowConnectorColor=   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "grdMedicamentos"
         End
      End
      Begin VB.Frame FraCred1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   6960
         TabIndex        =   38
         Top             =   2000
         Width           =   4635
         Begin VB.Frame FraAdulto 
            Height          =   1155
            Left            =   1440
            TabIndex        =   42
            Top             =   0
            Width           =   4605
            Begin VB.CheckBox chkPersonalSalud 
               Caption         =   "Personal de salud"
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
               Left            =   90
               TabIndex        =   43
               Top             =   150
               Width           =   2175
            End
            Begin VB.CheckBox chkDemandaIndividual 
               Caption         =   "Por demanda individual"
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
               Left            =   90
               TabIndex        =   44
               Top             =   390
               Width           =   3525
            End
            Begin VB.CheckBox chkMujerReproductiva 
               Caption         =   "Mujer en edad reproductiva"
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
               Left            =   90
               TabIndex        =   45
               Top             =   630
               Width           =   3345
            End
            Begin VB.CheckBox chkGestante 
               Caption         =   "Gestante"
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
               Left            =   90
               TabIndex        =   46
               Top             =   870
               Width           =   3165
            End
         End
         Begin VB.CheckBox chkAlimentacionComplementaria 
            Caption         =   "Alimentación complementaria"
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
            Left            =   30
            TabIndex        =   40
            Top             =   390
            Width           =   3165
         End
         Begin VB.CheckBox chkEstimulacionTemprana 
            Caption         =   "Estimulación temprana"
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
            Left            =   30
            TabIndex        =   39
            Top             =   150
            Width           =   2175
         End
         Begin VB.CheckBox chkLactanciaMaterna 
            Caption         =   "Lactancia materna complementaria"
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
            Left            =   30
            TabIndex        =   41
            Top             =   630
            Width           =   3165
         End
      End
      Begin VB.Frame FraCred 
         Caption         =   "Control del Crecimiento y desarrollo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   6960
         TabIndex        =   24
         Top             =   360
         Width           =   4635
         Begin VB.TextBox Cred7 
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
            Left            =   2670
            TabIndex        =   31
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   450
            Width           =   285
         End
         Begin VB.TextBox Cred8 
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
            Left            =   2970
            TabIndex        =   32
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   450
            Width           =   285
         End
         Begin VB.TextBox Cred9 
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
            Left            =   3270
            TabIndex        =   33
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   450
            Width           =   285
         End
         Begin VB.TextBox Cred10 
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
            Left            =   3570
            TabIndex        =   34
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   450
            Width           =   285
         End
         Begin VB.TextBox Cred11 
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
            Left            =   3870
            TabIndex        =   35
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   450
            Width           =   285
         End
         Begin VB.TextBox Cred12 
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
            Left            =   4170
            TabIndex        =   36
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   450
            Width           =   285
         End
         Begin VB.TextBox Cred6 
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
            TabIndex        =   30
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   450
            Width           =   285
         End
         Begin VB.TextBox Cred5 
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
            Left            =   2040
            TabIndex        =   29
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   450
            Width           =   285
         End
         Begin VB.TextBox Cred4 
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
            Left            =   1740
            TabIndex        =   28
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   450
            Width           =   285
         End
         Begin VB.TextBox Cred3 
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
            Left            =   1440
            TabIndex        =   27
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   450
            Width           =   285
         End
         Begin VB.TextBox Cred2 
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
            Left            =   1140
            TabIndex        =   26
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   450
            Width           =   285
         End
         Begin VB.TextBox Cred1 
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
            Left            =   840
            TabIndex        =   25
            ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
            Top             =   450
            Width           =   285
         End
         Begin UltraGrid.SSUltraGrid grdCred 
            Height          =   795
            Left            =   60
            TabIndex        =   37
            Top             =   780
            Width           =   4470
            _ExtentX        =   7885
            _ExtentY        =   1402
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   68157460
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Override        =   "ucPerinatal.ctx":5C44
            Caption         =   "SSUltraGrid1"
         End
         Begin VB.Label lblCred 
            AutoSize        =   -1  'True
            Caption         =   "1er año"
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
            TabIndex        =   73
            Top             =   480
            Width           =   690
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            Caption         =   "1"
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
            Left            =   930
            TabIndex        =   72
            Top             =   240
            Width           =   105
         End
         Begin VB.Label lbl2 
            AutoSize        =   -1  'True
            Caption         =   "2"
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
            Left            =   1230
            TabIndex        =   71
            Top             =   240
            Width           =   105
         End
         Begin VB.Label lbl3 
            AutoSize        =   -1  'True
            Caption         =   "3"
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
            Left            =   1500
            TabIndex        =   70
            Top             =   240
            Width           =   105
         End
         Begin VB.Label lbl4 
            AutoSize        =   -1  'True
            Caption         =   "4"
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
            Left            =   1800
            TabIndex        =   69
            Top             =   240
            Width           =   105
         End
         Begin VB.Label lbl5 
            AutoSize        =   -1  'True
            Caption         =   "5"
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
            Left            =   2130
            TabIndex        =   68
            Top             =   240
            Width           =   105
         End
         Begin VB.Label lbl6 
            AutoSize        =   -1  'True
            Caption         =   "6"
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
            Left            =   2430
            TabIndex        =   67
            Top             =   240
            Width           =   105
         End
         Begin VB.Label lbl7 
            AutoSize        =   -1  'True
            Caption         =   "7"
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
            Left            =   2730
            TabIndex        =   66
            Top             =   240
            Width           =   105
         End
         Begin VB.Label lbl8 
            AutoSize        =   -1  'True
            Caption         =   "8"
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
            Left            =   3000
            TabIndex        =   65
            Top             =   240
            Width           =   105
         End
         Begin VB.Label lbl9 
            AutoSize        =   -1  'True
            Caption         =   "9"
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
            Left            =   3300
            TabIndex        =   64
            Top             =   240
            Width           =   105
         End
         Begin VB.Label lbl10 
            AutoSize        =   -1  'True
            Caption         =   "10"
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
            Left            =   3570
            TabIndex        =   63
            Top             =   240
            Width           =   210
         End
         Begin VB.Label lbl11 
            AutoSize        =   -1  'True
            Caption         =   "11"
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
            Left            =   3870
            TabIndex        =   62
            Top             =   240
            Width           =   210
         End
         Begin VB.Label lbl12 
            AutoSize        =   -1  'True
            Caption         =   "12"
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
            Left            =   4170
            TabIndex        =   61
            Top             =   240
            Width           =   210
         End
      End
      Begin OWC11.ChartSpace ChartSpace1 
         Height          =   2595
         Left            =   6990
         OleObjectBlob   =   "ucPerinatal.ctx":5C9A
         TabIndex        =   47
         Top             =   3270
         Width           =   4605
      End
      Begin VB.Frame FraDxMorbilidad 
         Caption         =   "Dx-Morbilidad"
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
         Height          =   1695
         Left            =   30
         TabIndex        =   17
         Top             =   3780
         Width           =   6945
         Begin ActiveInput.SSComboBoxEx cmbMorbilidadFrec 
            Height          =   345
            Left            =   60
            TabIndex        =   21
            Top             =   240
            Width           =   5550
            _ExtentX        =   9790
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
            Separator       =   " | "
         End
         Begin VB.CommandButton btnAgregarDxMorbilidad 
            DisabledPicture =   "ucPerinatal.ctx":688E
            DownPicture     =   "ucPerinatal.ctx":6C77
            Height          =   315
            Left            =   5640
            Picture         =   "ucPerinatal.ctx":7083
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   240
            Width           =   390
         End
         Begin VB.CommandButton btnBusquedaDiagnostico 
            Caption         =   "..."
            Height          =   315
            Left            =   6540
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Busca Dx"
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton btnQuitaDxMorbilidad 
            DisabledPicture =   "ucPerinatal.ctx":748F
            DownPicture     =   "ucPerinatal.ctx":781A
            Height          =   315
            Left            =   6030
            Picture         =   "ucPerinatal.ctx":7BAD
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Elimina todos los Dx"
            Top             =   240
            Width           =   465
         End
         Begin UltraGrid.SSUltraGrid grdMorbilidadFrec 
            Height          =   1125
            Left            =   30
            TabIndex        =   20
            Top             =   540
            Width           =   6870
            _ExtentX        =   12118
            _ExtentY        =   1984
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   68157460
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Override        =   "ucPerinatal.ctx":7F3E
            Caption         =   "SSUltraGrid1"
         End
         Begin PVCOMBOLibCtl.PVComboBox cmbLabHisDxMorbilidad 
            Height          =   330
            Left            =   5640
            TabIndex        =   22
            Top             =   240
            Visible         =   0   'False
            Width           =   1065
            _Version        =   524288
            _cx             =   1879
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
   End
End
Attribute VB_Name = "ucPerinatal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim oRsInmunizaciones As New Recordset
Dim oRsCptFrecuentes As New Recordset
Dim oRsDxDesarrollo As New Recordset
Dim oRsMorbilidadFrec As New Recordset
Dim oRsCredHistorico As New Recordset
Dim oRsFarmaciaMI As New Recordset
Dim oRsPercentil As New Recordset
Dim oRsDxDesarrolloAutomaticos As New Recordset
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_Formulario As New sighentidades.Formulario
Dim lcSql As String
Dim lnIdModulo As sighentidades.sghPerinatalModulos
Dim mo_idPerinatalAtencion As Long
Dim lcEdadCredEnAtencion As String
Dim ml_IdPaciente As Long
Const lcCombo As String = "o"
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idTipoSexo As Long
Dim ml_EdadEnMeses As Long
Dim ml_idAtencion As Long
Dim lnPercentilPE As Long, lnPercentilTE As Long, lnPercentilPT As Long, lnPercentilIMC As Double
Dim lnPercentilPE_Z As Double, lnPercentilTE_Z As Double, lnPercentilPT_Z As Double, lnPercentilIMC_Z As Double
Const lcHasta28Dias As String = "<=28 días"
Const lcDe29diasHasta1anio As String = ">28d-<1año"
Const lnPercentilNull As Long = 0
Dim lnEdadEnAniosEnAtencion As Integer
Dim ml_idCuentaAtencion As Long
Dim ml_IdServicioIngreso As Long
Dim ml_IdPuntoCargaHosp As Long
'

Dim xValues As Variant, yValuesPT As Variant, yValuesTE As Variant, yValuesPE As Variant
Dim owcChart As OWC11.ChChart
Dim owcSeries As OWC11.ChSeries
Dim lnNroPuntosGraficos As Integer
Dim lnIdAtencionCred As Long    'usado para CONTROL
Dim lnIdAtencionCred1 As Long   'usado para los CHECK
Dim ml_FechaAtencion As Date
Dim ml_idUsuario As Long
Dim ml_YaCargoUnaSolaVez As Boolean
'mgaray
Dim oEdad As Edad
Dim md_fechaNacimiento As Date
Dim md_fechaActual As Date
Dim noEjecutarAccion As Boolean
'para alamcenar la celda activa
Dim ssRowActivate As SSRow
Dim ssCellActivate As SSCell
Dim mo_RsDesarrolloPendiente As ADODB.Recordset
Dim mo_RsCrecimientoPendiente As ADODB.Recordset
Dim ml_IdEstablecimiento As Long
Dim ms_MensajeError As String
Dim mo_RsLabHis As ADODB.Recordset
Dim ml_IdFormaPago As Long
'usado para la verificacion de duplicados entre procedimientos e inmunizaciones
Dim mo_rsImunizacionesPendientes As ADODB.Recordset
'mgaray201411b
Dim mb_EstaMarcadoEjecucionPsicomotor As Boolean
'mgaray201411e
Dim xValuesEdad As Variant, yValuesTallaD0 As Variant, yValuesTallaD1 As Variant, yValuesTallaD2 As Variant, yValuesTallaD_1 As Variant, yValuesTallaD_2 As Variant
Dim xValuesEdadPeso As Variant, yValuesPesoD0 As Variant, yValuesPesoD1 As Variant, yValuesPesoD2 As Variant, yValuesPesoD_1 As Variant, yValuesPesoD_2 As Variant
Dim xValuesEdadAtencion As Variant, xValuesEdadPesoAtencion As Variant, yValuesTalla As Variant, yValuesPeso As Variant
Dim mo_NroHistoriaClinica As Long
Dim mo_DOAtencionesCE As SIGHComun.DOAtencionesCE

Property Let idAtencion(lValue As Long)
   ml_idAtencion = lValue
End Property

Property Let NroHistoriaClinica(lValue As Long)
   mo_NroHistoriaClinica = lValue
End Property

Property Set DOAtencionesCE(lValue As SIGHComun.DOAtencionesCE)
   Set mo_DOAtencionesCE = lValue
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

Property Let EdadEnMeses(lValue As Long)
   ml_EdadEnMeses = lValue
End Property
Property Let idTipoSexo(lValue As Long)
   ml_idTipoSexo = lValue
End Property

Property Let idPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property

Property Let FechaAtencion(lValue As Date)
   ml_FechaAtencion = lValue
   calcularEdadPaciente
End Property

Property Let FechaNacimiento(lValue As Date)
   md_fechaNacimiento = lValue
   calcularEdadPaciente
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Property Get getRsDesarrolloPendiente() As ADODB.Recordset
    'mgaray201412a
    If mo_RsDesarrolloPendiente Is Nothing Then
        Set getRsDesarrolloPendiente = Nothing
    Else
        Set getRsDesarrolloPendiente = mo_RsDesarrolloPendiente.Clone()
    End If
End Property
Property Get getRsCrecimientoPendiente() As ADODB.Recordset
    'mgaray201412a
    If mo_RsCrecimientoPendiente Is Nothing Then
        Set getRsCrecimientoPendiente = Nothing
    Else
        Set getRsCrecimientoPendiente = mo_RsCrecimientoPendiente.Clone
    End If
End Property

'mgaray201410c
Property Get NumeroSesionDesarrollo() As String
   NumeroSesionDesarrollo = frAtencionDesarrollo.Tag
End Property

'@Implementar
Public Sub Inicializar()
    Set mo_RsLabHis = mo_reglasComunes.DevuelveHIS_SITUACIOporDescripcion()
    Set cmbLabHisDxDesarrollo.ListSource = mo_RsLabHis
    Set cmbLabHisDxMorbilidad.ListSource = mo_RsLabHis
    'mgaray201411a
    Set cmbLabHisProcFrecuentes.ListSource = mo_RsLabHis
    
    'mgaray201412a
'    cmbLabHisDxDesarrollo.Visible = True
'    cmbLabHisDxMorbilidad.Visible = True
'    cmbLabHisProcFrecuentes.Visible = True
    
    If ml_YaCargoUnaSolaVez = False Then
        ml_YaCargoUnaSolaVez = True
        CreaTemporales
        
        InicializarLaGrilla grdInmunizaciones
        InicializarLaGrilla grdCptFrecuentes
        InicializarLaGrilla grdMorbilidadDesarollo
        InicializarLaGrilla grdMorbilidadFrec
        'permisos
        Dim ms_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
        Dim oRsPermisos As New Recordset
        HabilitaDeshabilita "Inmunizaciones", False
        HabilitaDeshabilita "OtrosCpt", False
        HabilitaDeshabilita "DxDesarrollo", False
        HabilitaDeshabilita "DxMorbilidad", False
        HabilitaDeshabilita "Cred", False
        HabilitaDeshabilita "Cred1", False
        HabilitaDeshabilita "Medicamentos", False
        'mgaray201411e
        STabPerinatal.TabVisible(1) = True
        
        STabPerinatal.TabVisible(4) = False
        STabPerinatal.TabVisible(5) = False
        
        mo_Formulario.HabilitarDeshabilitar txtFechaProgramadaDesarrollo, False
        mo_Formulario.HabilitarDeshabilitar txtEvalucionDesarrollo, False
        
        
        Set oRsPermisos = ms_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_idUsuario)
        If oRsPermisos.RecordCount > 0 Then
           Do While Not oRsPermisos.EOF
              Select Case oRsPermisos.Fields!IdPermiso
              Case 400    'Perinatal - sólo Inmunizaciones
                   HabilitaDeshabilita "Inmunizaciones", True
              Case 401    'Perinatal - sólo Otros Procedimientos
                   HabilitaDeshabilita "OtrosCpt", True
              Case 402    'Perinatal - sólo Dx Desarrollo
                   HabilitaDeshabilita "DxDesarrollo", True
              Case 403    'Perinatal - sólo Dx Morbilidad
                   HabilitaDeshabilita "DxMorbilidad", True
              Case 404    'Perinatal - sólo Cred (control)
                   HabilitaDeshabilita "Cred", True
              Case 405    'Perinatal - sólo Tratamiento recibido
                   HabilitaDeshabilita "Medicamentos", True
              Case 406    'Perinatal -sólo Cred1 (lactante,estimulación temprana,etc)
                   HabilitaDeshabilita "Cred1", True
              End Select
              oRsPermisos.MoveNext
           Loop
        End If
        Set oRsPermisos = Nothing
        Set ms_ReglasSeguridad = Nothing
    End If
    'mgaray
    'ml_IdEstablecimiento = lcBuscaParametro.SeleccionaFilaParametro(280)
    Dim oDoEstablecimiento As New DOEstablecimiento
    'mgaray201411e
    Dim sCodigoRenaes As String
    sCodigoRenaes = lcBuscaParametro.SeleccionaFilaParametro(280)
    If mo_reglasComunes.EstablecimientosSeleccionarPorCodigo(sCodigoRenaes, _
                    oDoEstablecimiento) = True Then
        ml_IdEstablecimiento = oDoEstablecimiento.IdEstablecimiento
    Else
        MsgBox "Codigo RENAES " & sCodigoRenaes & " No Encontrado en la Lista de Establecimientos, revise la tabla parametros(280) ó actualice su listado de Establecimientos", vbInformation, "Modulo Niño Sano"
    End If
End Sub
Sub CargaMedicamentos(oConexion As Connection)
     Dim oRsTmp1 As New Recordset
     Set oRsTmp1 = mo_reglasComunes.PerinatalCatalogoMedicamentosSeleccionarPorIdModulo(lnIdModulo, oConexion)
     If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
           oRsFarmaciaMI.AddNew
           oRsFarmaciaMI.Fields!seleccionar = False
           oRsFarmaciaMI.Fields!Id = oRsTmp1.Fields!idProducto
           oRsFarmaciaMI.Fields!Medicamento = oRsTmp1.Fields!nombre
           oRsFarmaciaMI.Update
           oRsTmp1.MoveNext
        Loop
        oRsFarmaciaMI.MoveFirst
     End If
     oRsTmp1.Close
     Set oRsTmp1 = Nothing
End Sub

Sub CargaDatosAcombos(lnEdad As Integer, lnIdTipoEdad As Integer, oConexion As Connection)
     Dim oRsTmp As New Recordset
     Dim lnIdListItem As Integer, lnIdListItem1 As Integer
     On Error GoTo ErrCombo
     'Visualiza CRED segun Modulo
     Cred1.Visible = True: Cred2.Visible = True: Cred3.Visible = True: Cred4.Visible = True: Cred5.Visible = True: Cred6.Visible = True
     Cred7.Visible = True: Cred8.Visible = True: Cred9.Visible = True: Cred10.Visible = True: Cred11.Visible = True: Cred12.Visible = True
     lbl1.Visible = True: lbl2.Visible = True: lbl3.Visible = True: lbl4.Visible = True: lbl5.Visible = True: lbl6.Visible = True
     lbl7.Visible = True: lbl8.Visible = True: lbl9.Visible = True: lbl10.Visible = True: lbl11.Visible = True: lbl12.Visible = True
     lblCred = "Control"
     Select Case lnIdTipoEdad
     Case 1  'años
          lnEdadEnAniosEnAtencion = lnEdad
          If lnEdad >= 1 And lnEdad <= 4 Then
            lnIdModulo = sighDesde1Hasta4anios
            If lnEdad = 1 Then
               lcEdadCredEnAtencion = "1"
               lblCred = "1er año"
               Cred7.Visible = False: Cred8.Visible = False: Cred9.Visible = False: Cred10.Visible = False: Cred11.Visible = False: Cred12.Visible = False
               lbl7.Visible = False: lbl8.Visible = False: lbl9.Visible = False: lbl10.Visible = False: lbl11.Visible = False: lbl12.Visible = False
            ElseIf lnEdad = 2 Then
               lblCred = "2° año"
               lcEdadCredEnAtencion = "2"
               Cred5.Visible = False: Cred6.Visible = False: Cred7.Visible = False: Cred8.Visible = False: Cred9.Visible = False: Cred10.Visible = False: Cred11.Visible = False: Cred12.Visible = False
               lbl5.Visible = False: lbl6.Visible = False: lbl7.Visible = False: lbl8.Visible = False: lbl9.Visible = False: lbl10.Visible = False: lbl11.Visible = False: lbl12.Visible = False
            ElseIf lnEdad = 3 Then
               lblCred = "3er año"
               lcEdadCredEnAtencion = "3"
               Cred5.Visible = False: Cred6.Visible = False: Cred7.Visible = False: Cred8.Visible = False: Cred9.Visible = False: Cred10.Visible = False: Cred11.Visible = False: Cred12.Visible = False
               lbl5.Visible = False: lbl6.Visible = False: lbl7.Visible = False: lbl8.Visible = False: lbl9.Visible = False: lbl10.Visible = False: lbl11.Visible = False: lbl12.Visible = False
            Else
               lblCred = "4° año"
               lcEdadCredEnAtencion = "4"
               Cred5.Visible = False: Cred6.Visible = False: Cred7.Visible = False: Cred8.Visible = False: Cred9.Visible = False: Cred10.Visible = False: Cred11.Visible = False: Cred12.Visible = False
               lbl5.Visible = False: lbl6.Visible = False: lbl7.Visible = False: lbl8.Visible = False: lbl9.Visible = False: lbl10.Visible = False: lbl11.Visible = False: lbl12.Visible = False
            End If
          ElseIf lnEdad >= 5 And lnEdad <= 9 Then
            lnIdModulo = sighDesde5Hasta9anios
            lcEdadCredEnAtencion = "5-9"
            Cred6.Visible = False: Cred7.Visible = False: Cred8.Visible = False: Cred9.Visible = False: Cred10.Visible = False: Cred11.Visible = False: Cred12.Visible = False
            lbl6.Visible = False: lbl7.Visible = False: lbl8.Visible = False: lbl9.Visible = False: lbl10.Visible = False: lbl11.Visible = False: lbl12.Visible = False
          ElseIf lnEdad >= 10 And lnEdad <= 11 Then
            lnIdModulo = sighDesde10Hasta11anios
            lcEdadCredEnAtencion = "10-11"
            Cred3.Visible = False: Cred4.Visible = False: Cred5.Visible = False: Cred6.Visible = False: Cred7.Visible = False: Cred8.Visible = False: Cred9.Visible = False: Cred10.Visible = False: Cred11.Visible = False: Cred12.Visible = False
            lbl3.Visible = False: lbl4.Visible = False:  lbl5.Visible = False: lbl6.Visible = False: lbl7.Visible = False: lbl8.Visible = False: lbl9.Visible = False: lbl10.Visible = False: lbl11.Visible = False: lbl12.Visible = False
          ElseIf lnEdad >= 12 And lnEdad <= 17 Then
            lnIdModulo = sighDesde12Hasta17anios
            lcEdadCredEnAtencion = "12-17"
            Cred3.Visible = False: Cred4.Visible = False: Cred5.Visible = False: Cred6.Visible = False: Cred7.Visible = False: Cred8.Visible = False: Cred9.Visible = False: Cred10.Visible = False: Cred11.Visible = False: Cred12.Visible = False
            lbl3.Visible = False: lbl4.Visible = False:  lbl5.Visible = False: lbl6.Visible = False: lbl7.Visible = False: lbl8.Visible = False: lbl9.Visible = False: lbl10.Visible = False: lbl11.Visible = False: lbl12.Visible = False
          Else
            lnIdModulo = sighDesde18anios
            lcEdadCredEnAtencion = "desde 18"
            Cred1.Visible = False: Cred2.Visible = False: Cred3.Visible = False: Cred4.Visible = False: Cred5.Visible = False: Cred6.Visible = False: Cred7.Visible = False: Cred8.Visible = False: Cred9.Visible = False: Cred10.Visible = False: Cred11.Visible = False: Cred12.Visible = False
            lbl1.Visible = False: lbl2.Visible = False: lbl3.Visible = False: lbl4.Visible = False: lbl5.Visible = False: lbl6.Visible = False: lbl7.Visible = False: lbl8.Visible = False: lbl9.Visible = False: lbl10.Visible = False: lbl11.Visible = False: lbl12.Visible = False
          End If
     Case 2  'mes
          lnEdadEnAniosEnAtencion = 0
          If lnEdad <= 1 Then
             lnIdModulo = sighHasta28Dias
             lcEdadCredEnAtencion = lcHasta28Dias
             Cred9.Visible = False: Cred10.Visible = False: Cred11.Visible = False: Cred12.Visible = False
             lbl9.Visible = False: lbl10.Visible = False: lbl11.Visible = False: lbl12.Visible = False
          Else
             lnIdModulo = sighDesde29diasHasta1anio
             lcEdadCredEnAtencion = lcDe29diasHasta1anio
             Cred1.Visible = False
             lbl1.Visible = False
          End If
     Case 3  'dias
          lnEdadEnAniosEnAtencion = 0
          If lnEdad <= 28 Then
             lnIdModulo = sighHasta28Dias
             lcEdadCredEnAtencion = lcHasta28Dias
             Cred9.Visible = False: Cred10.Visible = False: Cred11.Visible = False: Cred12.Visible = False
             lbl9.Visible = False: lbl10.Visible = False: lbl11.Visible = False: lbl12.Visible = False
          Else
             lnIdModulo = sighDesde29diasHasta1anio
             lcEdadCredEnAtencion = lcDe29diasHasta1anio
             Cred1.Visible = False
             lbl1.Visible = False
          End If
     Case Else    'horas
          lnEdadEnAniosEnAtencion = 0
          lnIdModulo = sighHasta28Dias
          lcEdadCredEnAtencion = lcHasta28Dias
          Cred9.Visible = False: Cred10.Visible = False: Cred11.Visible = False: Cred12.Visible = False
          lbl9.Visible = False: lbl10.Visible = False: lbl11.Visible = False: lbl12.Visible = False
     End Select
     'Inicializa controles Checks segun modulos
     chkEstimulacionTemprana.Visible = False
     chkAlimentacionComplementaria.Visible = False
     chkLactanciaMaterna.Visible = False
     Select Case lnIdModulo
     Case sighHasta28Dias, sighDesde1Hasta4anios
            chkEstimulacionTemprana.Visible = True
            chkLactanciaMaterna.Visible = True
     Case sighDesde29diasHasta1anio
            chkEstimulacionTemprana.Visible = True
            chkAlimentacionComplementaria.Visible = True
            chkLactanciaMaterna.Visible = True
     End Select
     'Carga Combos: Dx Morbilidad y Dx Desarrollo*****Ademas carga Rangos Percentil, para Dx automaticos
     Set oRsTmp = mo_reglasComunes.PerinatalCatalogoCie10SeleccionarPorIdModulo(lnIdModulo, oConexion)
     cmbMorbilidadFrec.Clear
     cmbDxDesarrollo.Clear
     LimpiaDxDesarrolloAutomaticos
     If oRsTmp.RecordCount > 0 Then
        lnIdListItem = 0
        lnIdListItem1 = 0
        oRsTmp.MoveFirst
        Do While Not oRsTmp.EOF
           If oRsTmp.Fields!idLista = sghPerinatalListas.sighMorbilidadFrecuente Then
                cmbMorbilidadFrec.ListItems.Add lnIdListItem, lcCombo + Trim(Str(oRsTmp.Fields!idDiagnostico)), _
                                                getCodigoDiagnosticoIgualLongitud(IIf(IsNull(oRsTmp.Fields!CodigoCIE2004), "", oRsTmp.Fields!CodigoCIE2004)) & " = " & oRsTmp.Fields!Descripcion, , , , , , _
                                                IIf(IsNull(oRsTmp.Fields!CodigoCIE2004), "", Trim(oRsTmp.Fields!CodigoCIE2004))
                lnIdListItem = lnIdListItem + 1
           Else
                cmbDxDesarrollo.ListItems.Add lnIdListItem1, lcCombo + Trim(Str(oRsTmp.Fields!idDiagnostico)), _
                                                getCodigoDiagnosticoIgualLongitud(IIf(IsNull(oRsTmp.Fields!CodigoCIE2004), "", oRsTmp.Fields!CodigoCIE2004)) & " = " & oRsTmp.Fields!Descripcion, , , , , , _
                                                IIf(IsNull(oRsTmp.Fields!CodigoCIE2004), "", Trim(oRsTmp.Fields!CodigoCIE2004))
                lnIdListItem1 = lnIdListItem1 + 1
           End If
           '
           If (Not IsNull(oRsTmp.Fields!rangoInicio)) And (Not IsNull(oRsTmp.Fields!rangoFinal)) Then
                oRsDxDesarrolloAutomaticos.AddNew
                oRsDxDesarrolloAutomaticos.Fields!idDiagnostico = oRsTmp.Fields!idDiagnostico
                oRsDxDesarrolloAutomaticos.Fields!rangoInicio = oRsTmp.Fields!rangoInicio
                oRsDxDesarrolloAutomaticos.Fields!rangoFinal = oRsTmp.Fields!rangoFinal
                oRsDxDesarrolloAutomaticos.Fields!CodigoCIE2004 = Trim(oRsTmp.Fields!CodigoCIE2004)
                oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO = Trim(oRsTmp.Fields!Descripcion)
                oRsDxDesarrolloAutomaticos.Fields!cie10his = Trim(oRsTmp.Fields!CodigoHIS)
                oRsDxDesarrolloAutomaticos.Update
           End If
           '
           oRsTmp.MoveNext
        Loop
     End If
     oRsTmp.Close
     'Carga combos: CPT de Inmunizaciones, otros CPT
     Set oRsTmp = mo_reglasComunes.PerinatalCatalogoCptSeleccionarPorIdModulo(lnIdModulo, oConexion)
     cmbEligeInmunizacion.Clear
     cmbProcedimientosFrecuentes.Clear
     If oRsTmp.RecordCount > 0 Then
        lnIdListItem = 0
        lnIdListItem1 = 0
        oRsTmp.MoveFirst
        Do While Not oRsTmp.EOF
           If oRsTmp.Fields!idLista = sghPerinatalListas.sighInmunizaciones Then
              cmbEligeInmunizacion.ListItems.Add lnIdListItem, lcCombo + Trim(Str(oRsTmp.Fields!idProducto)), oRsTmp.Fields!nombre
              lnIdListItem = lnIdListItem + 1
           Else
              cmbProcedimientosFrecuentes.ListItems.Add lnIdListItem1, lcCombo + Trim(Str(oRsTmp.Fields!idProducto)), oRsTmp.Fields!nombre
              lnIdListItem1 = lnIdListItem1 + 1
           End If
           oRsTmp.MoveNext
        Loop
     End If
     cmbDxDesarrollo.Text = ""
     cmbProcedimientosFrecuentes.Text = ""
     oRsTmp.Close
     'Carga Combo de Medicamentos
     LimpiaMedicamentos
     CargaMedicamentos oConexion
     '
     Set oRsTmp = Nothing
     Exit Sub
ErrCombo:
   MsgBox Err.Description
   Resume
End Sub

Sub LimpiaMedicamentos()
    CreaTemporalFarmacia
End Sub

Private Sub btnActualizarSesionesPendientes_Click()
On Error GoTo miError
    Dim oFormulario As New FrmPerinatalDesarrolloPendiente
    
    oFormulario.idAtencion = ml_idAtencion
    oFormulario.FechaAtencion = ml_FechaAtencion
    oFormulario.FechaNacimiento = md_fechaNacimiento
    oFormulario.idPaciente = ml_IdPaciente
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.Inicializar
    oFormulario.Show 1
    Call cargarListaDesarrollo
miError:
    If Err Then
        MsgBox Err.Description, vbInformation, "Mensaje"
    End If
End Sub

'mgaray201401003
Private Sub btnAgregarDxCRED_Click()
    If cmbDxDesarrollo.SelectedItems.Count = 0 Then
        MsgBox "Seleccione DX de Crecimiento y Desarrollo que desea agregar", vbInformation, "Módulo Niño Sano"
        Exit Sub
    End If
    If AgregarDxCrecimientoDesarrolloSeleccionadoDesdeListado(cmbDxDesarrollo.SelectedItems(0), Right(Trim(cmbLabHisDxDesarrollo.Text), 3)) = False Then
'        MsgBox "DX y Lab de Crecimiento y Desarrollo ya ha sido agregado", vbInformation, "Módulo Perinatal"
    Else
        cmbDxDesarrollo.Text = ""
        cmbLabHisDxDesarrollo.Text = ""
    End If
End Sub
'mgaray201401003
Private Sub btnAgregarDxMorbilidad_Click()
    If cmbMorbilidadFrec.SelectedItems.Count = 0 Then
        MsgBox "Seleccione DX Morbilidad que desea agregar", vbInformation, "Módulo Niño Sano"
        Exit Sub
    End If
    
    
    If AgregarDxMorbilidadSeleccionadoDesdeListado(cmbMorbilidadFrec.SelectedItems(0), Right(Trim(cmbLabHisDxMorbilidad.Text), 3)) = False Then
'        MsgBox "DX y Lab de Morbilidad ya ha sido agregado", vbInformation, "Módulo Perinatal"
    Else
        cmbMorbilidadFrec.Text = ""
        cmbLabHisDxMorbilidad.Text = ""
    End If
End Sub
'mgaray201401003
Private Sub btnAgregarProcedminiento_Click()
    If cmbProcedimientosFrecuentes.SelectedItems.Count = 0 Then
        MsgBox "Seleccione Procedimiento que desea agregar", vbInformation, "Módulo Niño Sano"
        Exit Sub
    End If
    If AgregarProcedimientosSeleccionadoDesdeListado(cmbProcedimientosFrecuentes.SelectedItems(0), Right(Trim(cmbLabHisProcFrecuentes.Text), 3)) = False Then
'        MsgBox "Procedimiento ya ha sido agregado", vbInformation, "Módulo Perinatal"
    Else
        cmbProcedimientosFrecuentes.Text = ""
        cmbLabHisProcFrecuentes.Text = ""
    End If
End Sub

Private Sub btnBuscaHistoricos_Click()
On Error GoTo miError
    Dim oFormulario As New frmPlanDesarrollo
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente

    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    oDOAtenIntePlanIntePaciente.idPaciente = ml_IdPaciente
    Set oFormulario.rsDesarrollo = oReglasAtencionIntegral.ListarPlanDesarrolloPacientePendientesParaImpresion(oDOAtenIntePlanIntePaciente)
    oFormulario.Show 1
miError:
    If Err Then
        MsgBox Err.Description, vbInformation, "Mensaje"
    End If
End Sub

Private Sub btnBusquedaDiagnostico_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
    Dim oDODiagnostico As DODiagnostico
    oBusqueda.CodigoDx = ""
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_reglasComunes.DiagnosticosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            'mgaray201411a
            Call AgregarDxMorbilidadSeleccionado(oDODiagnostico, "", False, False)
'            If oRsMorbilidadFrec.RecordCount > 0 Then
'               oRsMorbilidadFrec.MoveFirst
'               oRsMorbilidadFrec.Find "id=" & oDODiagnostico.IdDiagnostico
'               If Not oRsMorbilidadFrec.EOF Then
'                  MsgBox "Ese Dx ya está registrado", vbInformation, "Mensaje"
'                  Exit Sub
'               End If
'            End If
'            oRsMorbilidadFrec.AddNew
'            oRsMorbilidadFrec.Fields!id = oDODiagnostico.IdDiagnostico
'            oRsMorbilidadFrec.Fields!DIAGNOSTICO = oDODiagnostico.Descripcion
'            oRsMorbilidadFrec.Fields!idAtencion = ml_idAtencion
'            oRsMorbilidadFrec.Fields!SeEligioConChek = False
'            oRsMorbilidadFrec.Fields!EsDxPerinatal = False
'            oRsMorbilidadFrec.Fields!CodigoCIE2004 = oDODiagnostico.CodigoCIE2004
'            oRsMorbilidadFrec.Update
        End If
    End If
    Set oBusqueda = Nothing
    
End Sub

Private Sub btnBusquedaDxDesarrollo_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
    Dim oDODiagnostico As DODiagnostico
    oBusqueda.CodigoDx = ""
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_reglasComunes.DiagnosticosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            If AgregarDxCrecimientoDesarrolloSeleccionado(oDODiagnostico, "") = False Then

            End If
        End If
    End If
    Set oBusqueda = Nothing
End Sub

Private Sub btnBusquedaProcedimientos_Click()
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
            
            If AgregarProcedimientosSeleccionado(oDoFactCatalogoServicio, "") = False Then
'                MsgBox "Procedimiento ya ha sido agregado", vbInformation, "Módulo Perinatal"
            Else
                cmbProcedimientosFrecuentes.Text = ""
            End If
        End If
    End If
    Set oBusqueda = Nothing

End Sub

Private Sub btnImprime_Click()
On Error GoTo miError
    Dim oDOPaciente As doPaciente
    Dim oAdminAdmision As New ReglasAdmision
    Dim sTipoSexo As String
    Dim sEdad As String
    Dim sDireccion As String

'    Set fgf = hshs(ml_idPaciente)
    Set oDOPaciente = oAdminAdmision.RetornaPacientesSeleccionarPorId(ml_IdPaciente)
    
    If oDOPaciente.idTipoSexo = 1 Then
        sTipoSexo = "Masculino"
    ElseIf oDOPaciente.idTipoSexo = 2 Then
        sTipoSexo = "Femenino"
    End If
    
    If md_fechaNacimiento <> 0 Then
        Dim ro_Edad As Edad
        ro_Edad = calcularEdadDisgregada(md_fechaNacimiento, Date)
        If ro_Edad.EdadAnio > 0 Then
            sEdad = ro_Edad.EdadAnio & " A"
        End If
        If ro_Edad.EdadMes > 0 Then
            sEdad = IIf(sEdad <> "", ",", "") & ro_Edad.EdadMes & " M"
        End If

    End If
    
    sDireccion = oDOPaciente.DireccionDomicilio
    Dim oDODistrito As New DODistrito
    If oDOPaciente.IdDistritoDomicilio > 0 Then
        Set oDODistrito = mo_reglasComunes.DistritoSeleccionarPorId(oDOPaciente.IdDistritoDomicilio)
        If Not (oDODistrito Is Nothing) Then
            sDireccion = sDireccion & IIf(sDireccion <> "", " - ", "") & oDODistrito.nombre
        End If
    End If
    
    Dim mrs_Shape As ADODB.Recordset
    Set mrs_Shape = grdPlanInmunizaciones.DataSource 'mo_ReglasLaboratorio.ReporteMuestraResultado
    With DsPlanInmunicaziones
      .Orientation = rptOrientPortrait
      .Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
      .Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
      .Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
      .Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
      .Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaServidorSQL
      .Sections("Cabecera").Controls("lblEstablecimiento").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
      .Sections("Cabecera").Controls("lblNombrePaciente").Caption = oDOPaciente.ApellidoPaterno & " " & oDOPaciente.ApellidoMaterno _
                                                                    & " " & oDOPaciente.PrimerNombre _
                                                                    & " " & oDOPaciente.SegundoNombre _
                                                                    & " " & oDOPaciente.TercerNombre
      .Sections("Cabecera").Controls("lblFechaNacimiento").Caption = Format(oDOPaciente.FechaNacimiento, sighentidades.DevuelveFechaSoloFormato_DMY)
      .Sections("Cabecera").Controls("lblSexo").Caption = sTipoSexo
      .Sections("Cabecera").Controls("lblEdad").Caption = sEdad
      .Sections("Cabecera").Controls("lblDireccion").Caption = sDireccion
      
      
      Set .Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
      Set .DataSource = mrs_Shape
      .Show 1
    End With
    Err = 0
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbInformation, "Error"
    End If
End Sub

'mgaray201401003
Private Sub btnQuitaDxDesarrollo_Click()
'    LimpiaDxDesarrollo
'
'    Dim lnFor As Integer, lnFor1 As Integer
'    On Error Resume Next
'    Do While True
'        For lnFor1 = 1 To 3
'            For lnFor = 0 To cmbDxDesarrollo.ListCount - 1
'                cmbDxDesarrollo.SelectedItems(lnFor).Selected = False
'            Next
'        Next
'        If cmbDxDesarrollo.SelectedItems.Count = 0 Then
'           Exit Do
'        End If
'    Loop
    Call EliminarDxCrecimientoDesarrolloSeleccionado
End Sub
'mgaray201401003
Private Sub btnQuitaDxMorbilidad_Click()
'    LimpiaMorbilidadFrecuente True
'
'    Dim lnFor As Integer, lnFor1 As Integer
'    On Error Resume Next
'    Do While True
'        For lnFor1 = 1 To 3
'            For lnFor = 0 To cmbMorbilidadFrec.ListCount - 1
'                cmbMorbilidadFrec.SelectedItems(lnFor).Selected = False
'            Next
'        Next
'        If cmbMorbilidadFrec.SelectedItems.Count = 0 Then
'           Exit Do
'        End If
'    Loop
    Call EliminarDxMorbilidadSeleccionado
End Sub
'mgaray201401003
Private Sub btnQuitaOtrosProcedimientos_Click()
'    LimpiaCPTFrecuentes
'
'    Dim lnFor As Integer, lnFor1 As Integer
'    On Error Resume Next
'    Do While True
'        For lnFor1 = 1 To 3
'            For lnFor = 0 To cmbProcedimientosFrecuentes.ListCount - 1
'                cmbProcedimientosFrecuentes.SelectedItems(lnFor).Selected = False
'            Next
'        Next
'        If cmbProcedimientosFrecuentes.SelectedItems.Count = 0 Then
'           Exit Do
'        End If
'    Loop
    Call EliminarProcedimientoSeleccionado
End Sub

Private Sub btnQuitarInmunizacion_Click()
    LimpiaInmunizaciones
    '
    Dim lnFor As Integer, lnFor1 As Integer
    On Error Resume Next
    Do While True
        For lnFor = 0 To cmbEligeInmunizacion.ListCount - 1
            cmbEligeInmunizacion.SelectedItems(lnFor).Selected = False
        Next
        If cmbEligeInmunizacion.SelectedItems.Count = 0 Then
           Exit Do
        End If
    Loop
End Sub







Private Sub cmbDxDesarrollo_LostFocus()
'Dim sItems As String
'Dim CBLI As SSCBListItem
'Dim iTrimLen As Integer
'     LimpiaDxDesarrollo
'     If cmbDxDesarrollo.SelectedItems.Count > 0 Then
'          For Each CBLI In cmbDxDesarrollo.SelectedItems
'              oRsDxDesarrollo.AddNew
'              oRsDxDesarrollo.Fields!ID = Val(Mid(CBLI.Key, 2, 100))
'              oRsDxDesarrollo.Fields!DIAGNOSTICO = CBLI.Text
'              oRsDxDesarrollo.Fields!idAtencion = ml_idAtencion
'              oRsDxDesarrollo.Update
'          Next CBLI
'          On Error Resume Next
'          oRsDxDesarrollo.MoveFirst
'     End If
End Sub

Private Sub cmbEligeInmunizacion_LostFocus()
Dim sItems As String
Dim CBLI As SSCBListItem
Dim iTrimLen As Integer
     LimpiaInmunizaciones
     If cmbEligeInmunizacion.SelectedItems.Count > 0 Then
          For Each CBLI In cmbEligeInmunizacion.SelectedItems
              oRsInmunizaciones.AddNew
              oRsInmunizaciones.Fields!Id = Val(Mid(CBLI.Key, 2, 100))
              oRsInmunizaciones.Fields!procedimiento = CBLI.Text
              oRsInmunizaciones.Fields!idAtencion = ml_idAtencion
              oRsInmunizaciones.Update
          Next CBLI
          On Error Resume Next
          oRsInmunizaciones.MoveFirst
     End If
End Sub


Sub CreaTemporales()
    With oRsInmunizaciones
          .Fields.Append "Id", adInteger
          .Fields.Append "Procedimiento", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdAtencion", adInteger
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdInmunizaciones.DataSource = oRsInmunizaciones
    mo_Apariencia.ConfigurarFilasBiColores grdInmunizaciones, sighentidades.GrillaConFilasBicolor
    grdInmunizaciones.Caption = ""
    '
    With oRsCptFrecuentes
          .Fields.Append "Id", adInteger
          .Fields.Append "Procedimiento", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdAtencion", adInteger
          'mgaray201411a
          .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable + adFldUpdatable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    
    End With
    Set grdCptFrecuentes.DataSource = oRsCptFrecuentes
    mo_Apariencia.ConfigurarFilasBiColores grdCptFrecuentes, sighentidades.GrillaConFilasBicolor
    grdCptFrecuentes.Caption = ""
    '
    With oRsDxDesarrollo
          .Fields.Append "Id", adInteger
          'mgaray201411a
          .Fields.Append "IdClasificacionDx", adInteger
          .Fields.Append "IdSubclasificacionDx", adInteger
          .Fields.Append "CodigoCIE2004", adVarChar, 7, adFldIsNullable
          .Fields.Append "Diagnostico", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdAtencion", adInteger
          .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable + adFldUpdatable
          .CursorType = adOpenDynamic
'          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdMorbilidadDesarollo.DataSource = oRsDxDesarrollo
    mo_Apariencia.ConfigurarFilasBiColores grdMorbilidadDesarollo, sighentidades.GrillaConFilasBicolor
    '
    With oRsMorbilidadFrec
          .Fields.Append "Id", adInteger
          'mgaray201411a
          .Fields.Append "IdClasificacionDx", adInteger
          .Fields.Append "IdSubclasificacionDx", adInteger
          .Fields.Append "CodigoCIE2004", adVarChar, 7, adFldIsNullable
          .Fields.Append "Diagnostico", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdAtencion", adInteger
          .Fields.Append "SeEligioConChek", adBoolean
          .Fields.Append "EsDxPerinatal", adBoolean
          .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable + adFldUpdatable
          .CursorType = adOpenDynamic
'         .CursorType = adOpenStatic
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdMorbilidadFrec.DataSource = oRsMorbilidadFrec
    mo_Apariencia.ConfigurarFilasBiColores grdMorbilidadFrec, sighentidades.GrillaConFilasBicolor
    grdMorbilidadFrec.Caption = ""
    '
    With oRsCredHistorico
          .Fields.Append "Anio", adVarChar, 20, adFldIsNullable
          .Fields.Append "Control", adVarChar, 255, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    
    End With
    Set grdCred.DataSource = oRsCredHistorico
    mo_Apariencia.ConfigurarFilasBiColores grdCred, sighentidades.GrillaConFilasBicolor
    grdCred.Caption = ""
    '
    CreaTemporalFarmacia
    '
    With oRsPercentil
          .Fields.Append "IdAtencion", adInteger
          .Fields.Append "EdadEnMeses", adInteger
          .Fields.Append "PercentilPE", adInteger
          .Fields.Append "PercentilTE", adInteger
          .Fields.Append "PercentilPT", adInteger
          .Fields.Append "PercentilIMC", adDouble
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    '
     CreaTemporalDxDesarrolloAutomaticos
     
End Sub
Sub CreaTemporalDxDesarrolloAutomaticos()
    If oRsDxDesarrolloAutomaticos.State = 1 Then Set oRsDxDesarrolloAutomaticos = Nothing
    With oRsDxDesarrolloAutomaticos
          .Fields.Append "idDiagnostico", adInteger
          .Fields.Append "RangoInicio", adDouble
          .Fields.Append "RangoFinal", adDouble
          .Fields.Append "CodigoCIE2004", adVarChar, 7, adFldIsNullable
          .Fields.Append "Diagnostico", adVarChar, 255, adFldIsNullable
          .Fields.Append "cie10his", adVarChar, 20, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With

End Sub

Sub CreaTemporalFarmacia()
    If oRsFarmaciaMI.State = 1 Then
       Set oRsFarmaciaMI = Nothing
    End If
    With oRsFarmaciaMI
          .Fields.Append "Seleccionar", adBoolean
          .Fields.Append "Id", adInteger
          .Fields.Append "Medicamento", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdAtencion", adInteger
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdMedicamentos.DataSource = oRsFarmaciaMI
    mo_Apariencia.ConfigurarFilasBiColores grdMedicamentos, sighentidades.GrillaConFilasBicolor
End Sub

Sub LimpiaPercentil()
    On Error GoTo errLimp
    With oRsPercentil
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

Sub LimpiaInmunizaciones()
    On Error GoTo errLimp
    With oRsInmunizaciones
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
Sub LimpiaCheckDeMedicamentos()
    On Error GoTo errLimp
    With oRsFarmaciaMI
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              .Fields!seleccionar = False
              .Fields!idAtencion = 0
              .Update
              .MoveNext
           Loop
        End If
    End With
errLimp:
End Sub

Sub LimpiaCPTFrecuentes()
    On Error GoTo errLimp
    With oRsCptFrecuentes
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


Sub LimpiaDxDesarrollo()
    On Error GoTo errLimp
    With oRsDxDesarrollo
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

Sub LimpiaMorbilidadFrecuente(lbLimpiarTodosDx As Boolean)
    On Error GoTo errLimp
    Dim lbContinuar As Boolean
    With oRsMorbilidadFrec
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              lbContinuar = True
              If lbLimpiarTodosDx = False Then
                 If oRsMorbilidadFrec.Fields!SeEligioConChek = False Then
                    lbContinuar = False
                 ElseIf oRsMorbilidadFrec.Fields!EsDxPerinatal = False Then
                    lbContinuar = False
                 End If
              End If
              
              If lbContinuar = True Then
                    .Delete
                    .Update
              End If
              .MoveNext
           Loop
        End If
    End With
errLimp:
End Sub

Sub LimpiaDxDesarrolloAutomaticos()
    On Error GoTo errLimp
    Set oRsDxDesarrolloAutomaticos = Nothing
    CreaTemporalDxDesarrolloAutomaticos
'    With oRsDxDesarrolloAutomaticos
'        If .RecordCount > 0 Then
'           .MoveFirst
'           Do While Not .EOF
'              .Delete
'              .Update
'              .MoveNext
'           Loop
'        End If
'    End With
errLimp:
    
End Sub


Private Sub cmbMorbilidadFrec_LostFocus()
'Dim sItems As String
'Dim CBLI As SSCBListItem
'Dim iTrimLen As Integer
'     LimpiaMorbilidadFrecuente False
'     If cmbMorbilidadFrec.SelectedItems.Count > 0 Then
'          For Each CBLI In cmbMorbilidadFrec.SelectedItems
'              oRsMorbilidadFrec.AddNew
'              oRsMorbilidadFrec.Fields!ID = Val(Mid(CBLI.Key, 2, 100))
'              oRsMorbilidadFrec.Fields!DIAGNOSTICO = CBLI.Text
'              oRsMorbilidadFrec.Fields!idAtencion = ml_idAtencion
'              oRsMorbilidadFrec.Fields!SeEligioConChek = True
'              oRsMorbilidadFrec.Fields!EsDxPerinatal = True
'              oRsMorbilidadFrec.Update
'          Next CBLI
'          On Error Resume Next
'          oRsMorbilidadFrec.MoveFirst
'     End If

End Sub

Private Sub cmbProcedimientosFrecuentes_LostFocus()
'Dim sItems As String
'Dim CBLI As SSCBListItem
'Dim iTrimLen As Integer
'     LimpiaCPTFrecuentes
'     If cmbProcedimientosFrecuentes.SelectedItems.Count > 0 Then
'          For Each CBLI In cmbProcedimientosFrecuentes.SelectedItems
'              oRsCptFrecuentes.AddNew
'              oRsCptFrecuentes.Fields!ID = Val(Mid(CBLI.Key, 2, 100))
'              oRsCptFrecuentes.Fields!procedimiento = CBLI.Text
'              oRsCptFrecuentes.Fields!idAtencion = ml_idAtencion
'              oRsCptFrecuentes.Update
'          Next CBLI
'          On Error Resume Next
'          oRsCptFrecuentes.MoveFirst
'     End If
End Sub


Private Sub Cred1_LostFocus()
    If Trim(Cred1.Text) <> "" Then
        Cred1.Text = UCase(Cred1.Text)
        If Cred1.Text <> "X" And Cred1.Text <> "E" Then
           MsgBox "Los valores pueden ser:  X (control en el Establecimiento),  E (control externo)", vbInformation, "Mensaje"
           Cred1.Text = ""
        End If
    End If
End Sub






Private Sub Cred10_LostFocus()
  If Trim(Cred10.Text) <> "" Then
    Cred10.Text = UCase(Cred1.Text)
    If Cred10.Text <> "X" And Cred10.Text <> "E" Then
       MsgBox "Los valores pueden ser:  X (control en el Establecimiento),  E (control externo)", vbInformation, "Mensaje"
       Cred10.Text = ""
    End If
  End If
End Sub


Private Sub Cred11_LostFocus()
  If Trim(Cred11.Text) <> "" Then
    Cred11.Text = UCase(Cred11.Text)
    If Cred11.Text <> "X" And Cred11.Text <> "E" Then
       MsgBox "Los valores pueden ser:  X (control en el Establecimiento),  E (control externo)", vbInformation, "Mensaje"
       Cred11.Text = ""
    End If
  End If
End Sub



Private Sub Cred12_LostFocus()
  If Trim(Cred12.Text) <> "" Then
    Cred12.Text = UCase(Cred12.Text)
    If Cred12.Text <> "X" And Cred12.Text <> "E" Then
       MsgBox "Los valores pueden ser:  X (control en el Establecimiento),  E (control externo)", vbInformation, "Mensaje"
       Cred12.Text = ""
    End If
  End If
End Sub

Private Sub Cred2_LostFocus()
  If Trim(Cred2.Text) <> "" Then
    Cred2.Text = UCase(Cred2.Text)
    If Cred2.Text <> "X" And Cred2.Text <> "E" Then
       MsgBox "Los valores pueden ser:  X (control en el Establecimiento),  E (control externo)", vbInformation, "Mensaje"
       Cred2.Text = ""
    End If
  End If
End Sub



Private Sub Cred3_LostFocus()
  If Trim(Cred3.Text) <> "" Then
    Cred3.Text = UCase(Cred3.Text)
    If Cred3.Text <> "X" And Cred3.Text <> "E" Then
       MsgBox "Los valores pueden ser:  X (control en el Establecimiento),  E (control externo)", vbInformation, "Mensaje"
       Cred3.Text = ""
    End If
  End If
End Sub



Private Sub Cred4_LostFocus()
  If Trim(Cred4.Text) <> "" Then
    Cred4.Text = UCase(Cred4.Text)
    If Cred4.Text <> "X" And Cred4.Text <> "E" Then
       MsgBox "Los valores pueden ser:  X (control en el Establecimiento),  E (control externo)", vbInformation, "Mensaje"
       Cred4.Text = ""
    End If
  End If
End Sub



Private Sub Cred5_LostFocus()
   If Trim(Cred5.Text) <> "" Then
    Cred5.Text = UCase(Cred5.Text)
    If Cred5.Text <> "X" And Cred5.Text <> "E" Then
       MsgBox "Los valores pueden ser:  X (control en el Establecimiento),  E (control externo)", vbInformation, "Mensaje"
       Cred5.Text = ""
    End If
  End If
End Sub



Private Sub Cred6_LostFocus()
   If Trim(Cred6.Text) <> "" Then
    Cred6.Text = UCase(Cred6.Text)
    If Cred6.Text <> "X" And Cred6.Text <> "E" Then
       MsgBox "Los valores pueden ser:  X (control en el Establecimiento),  E (control externo)", vbInformation, "Mensaje"
       Cred6.Text = ""
    End If
   End If
End Sub



Private Sub Cred7_LostFocus()
  If Trim(Cred7.Text) <> "" Then
    Cred7.Text = UCase(Cred7.Text)
    If Cred7.Text <> "X" And Cred7.Text <> "E" Then
       MsgBox "Los valores pueden ser:  X (control en el Establecimiento),  E (control externo)", vbInformation, "Mensaje"
       Cred7.Text = ""
    End If
  End If
End Sub



Private Sub Cred8_LostFocus()
  If Trim(Cred8.Text) <> "" Then
    Cred8.Text = UCase(Cred8.Text)
    If Cred8.Text <> "X" And Cred8.Text <> "E" Then
       MsgBox "Los valores pueden ser:  X (control en el Establecimiento),  E (control externo)", vbInformation, "Mensaje"
       Cred8.Text = ""
    End If
  End If
End Sub



Private Sub Cred9_LostFocus()
  If Trim(Cred9.Text) <> "" Then
    Cred9.Text = UCase(Cred9.Text)
    If Cred9.Text <> "X" And Cred9.Text <> "E" Then
       MsgBox "Los valores pueden ser:  X (control en el Establecimiento),  E (control externo)", vbInformation, "Mensaje"
       Cred9.Text = ""
    End If
  End If
End Sub


Private Sub grdCptFrecuentes_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Cell.Column.Key = "labConfHIS" Then
        If ValidarCptFrecuente(Cell.Row.Cells("ID").Value, CStr(IIf(IsNull(NewValue), "", NewValue)), True, oRsCptFrecuentes.Bookmark) = False Then
            Cancel = True
        End If
    End If
End Sub

Private Sub grdCptFrecuentes_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
    Call EliminarProcedimientoSeleccionado
End Sub

Private Sub grdCptFrecuentes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
'    InicializarLaGrilla grdCptFrecuentes
'    grdCptFrecuentes.Bands(0).Columns("Id").Hidden = True
'    grdCptFrecuentes.Bands(0).Columns("Procedimiento").Header.Caption = "Procedimiento"
'    grdCptFrecuentes.Bands(0).Columns("Procedimiento").Width = 6300
End Sub

Private Sub grdCred_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdCred.Bands(0).Columns("Anio").Header.Caption = "Año"
    grdCred.Bands(0).Columns("Control").Header.Caption = "01  02  03  04  05  06  07  08  09  10  11  12"
                                                        '"01xx02xx03xx04xx05xx06xx07xx08xx09xx10xx11xx12xx13"
                                                        ' 12345678901234567890123456789012345678901234567890
    grdCred.Bands(0).Columns("anio").Width = 900
    grdCred.Bands(0).Columns("control").Width = 3000

End Sub

Private Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
    Dim oRsLabHis As ADODB.Recordset
    
    
    
    Select Case oGrilla.Name
    Case "grdInmunizaciones"
         oGrilla.Bands(0).Columns("Id").Hidden = True
         oGrilla.Bands(0).Columns("IdAtencion").Hidden = True
         oGrilla.Bands(0).Columns("Procedimiento").Header.Caption = "Procedimiento"
         oGrilla.Bands(0).Columns("Procedimiento").Width = 6300
    Case "grdMedicamentos"
         oGrilla.Bands(0).Columns("Id").Hidden = True
         oGrilla.Bands(0).Columns("IdAtencion").Hidden = True
         oGrilla.Bands(0).Columns("seleccionar").Width = 500
         oGrilla.Bands(0).Columns("medicamento").Width = 5650
    Case "grdMorbilidadDesarollo"
        
        'mgaray20141003
        oGrilla.Bands(0).Override.AllowUpdate = ssAllowUpdateDefault

         oGrilla.Bands(0).Columns("Id").Hidden = True
         oGrilla.Bands(0).Columns("IdAtencion").Hidden = True
         oGrilla.Bands(0).Columns("IdClasificacionDx").Hidden = True
         
         
         oGrilla.Bands(0).Columns("CodigoCIE2004").Header.Caption = "Código"
         oGrilla.Bands(0).Columns("CodigoCIE2004").Width = 700
         oGrilla.Bands(0).Columns("CodigoCIE2004").Activation = ssActivationActivateNoEdit
         
         oGrilla.Bands(0).Columns("IdSubclasificacionDx").Header.Caption = "Tipo Dx"
         oGrilla.Bands(0).Columns("IdSubclasificacionDx").Width = 900
         oGrilla.Bands(0).Columns("IdSubclasificacionDx").Activation = ssActivationAllowEdit
         oGrilla.Bands(0).Columns("IdSubclasificacionDx").Style = ssStyleDropDownValidate
         
         oGrilla.Bands(0).Columns("labConfHIS").Header.Caption = "Lab"
         oGrilla.Bands(0).Columns("labConfHIS").Width = 800
         oGrilla.Bands(0).Columns("labConfHIS").Activation = ssActivationAllowEdit
         
         oGrilla.Bands(0).Columns("diagnostico").Header.Caption = "Diagnóstico"
         oGrilla.Bands(0).Columns("diagnostico").Width = oGrilla.Width - oGrilla.Bands(0).Columns("CodigoCIE2004").Width - _
                            oGrilla.Bands(0).Columns("IdSubclasificacionDx").Width - 600 ' - oGrilla.Bands(0).Columns("labConfHIS").Width - 600  '6300
        oGrilla.Bands(0).Columns("diagnostico").Activation = ssActivationActivateNoEdit
        
        'mgaray201412a
         oGrilla.Bands(0).Columns("labConfHIS").Hidden = True
        Call AsignarListaDeLabsEnGridaDiagnosticos(oGrilla, "labConfHIS")
        Call AsignarListaDeTipoDxEnGrida(oGrilla, "IdSubclasificacionDx")
        
    Case "grdMorbilidadFrec"
        'mgaray20141003
        oGrilla.Bands(0).Override.AllowUpdate = ssAllowUpdateDefault
        
         oGrilla.Bands(0).Columns("Id").Hidden = True
         oGrilla.Bands(0).Columns("IdAtencion").Hidden = True
         oGrilla.Bands(0).Columns("SeEligioConChek").Hidden = True
         oGrilla.Bands(0).Columns("EsDxPerinatal").Hidden = True
         oGrilla.Bands(0).Columns("IdClasificacionDx").Hidden = True
         oGrilla.Bands(0).Columns("CodigoCIE2004").Header.Caption = "Código"
         oGrilla.Bands(0).Columns("CodigoCIE2004").Width = 800
         oGrilla.Bands(0).Columns("CodigoCIE2004").Activation = ssActivationActivateNoEdit
         
         oGrilla.Bands(0).Columns("IdSubclasificacionDx").Header.Caption = "Tipo Dx"
         oGrilla.Bands(0).Columns("IdSubclasificacionDx").Width = 900
         oGrilla.Bands(0).Columns("IdSubclasificacionDx").Activation = ssActivationAllowEdit
         oGrilla.Bands(0).Columns("IdSubclasificacionDx").Style = ssStyleDropDownValidate
         
         oGrilla.Bands(0).Columns("labConfHIS").Header.Caption = "Lab"
         oGrilla.Bands(0).Columns("labConfHIS").Width = 700
         oGrilla.Bands(0).Columns("labConfHIS").Activation = ssActivationAllowEdit
         
         oGrilla.Bands(0).Columns("diagnostico").Header.Caption = "Diagnóstico"
         oGrilla.Bands(0).Columns("diagnostico").Width = oGrilla.Width - oGrilla.Bands(0).Columns("CodigoCIE2004").Width - _
                            oGrilla.Bands(0).Columns("IdSubclasificacionDx").Width - 600 ' - oGrilla.Bands(0).Columns("labConfHIS").Width - 600  '6300
        oGrilla.Bands(0).Columns("diagnostico").Activation = ssActivationActivateNoEdit
        'mgaray201412a
         oGrilla.Bands(0).Columns("labConfHIS").Hidden = True
        Call AsignarListaDeLabsEnGridaDiagnosticos(oGrilla, "labConfHIS")
        Call AsignarListaDeTipoDxEnGrida(oGrilla, "IdSubclasificacionDx")
        
    Case "grdCptFrecuentes"
         oGrilla.Bands(0).Columns("Id").Hidden = True
         oGrilla.Bands(0).Columns("IdAtencion").Hidden = True
         
         oGrilla.Bands(0).Columns("labConfHIS").Header.Caption = "Lab"
         oGrilla.Bands(0).Columns("labConfHIS").Width = 700
         oGrilla.Bands(0).Columns("labConfHIS").Activation = ssActivationAllowEdit
         
         oGrilla.Bands(0).Columns("Procedimiento").Header.Caption = "Procedimiento"
         oGrilla.Bands(0).Columns("Procedimiento").Width = oGrilla.Width - _
                            600 'oGrilla.Bands(0).Columns("labConfHIS").Width - 600 '6300
        oGrilla.Bands(0).Columns("Procedimiento").Activation = ssActivationActivateNoEdit
         
         'mgaray201412a
         oGrilla.Bands(0).Columns("labConfHIS").Hidden = True
         Call AsignarListaDeLabsEnGridaDiagnosticos(oGrilla, "labConfHIS")
    End Select
End Sub
Private Sub grdInmunizaciones_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdInmunizaciones
End Sub


Private Sub grdMedicamentos_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdMedicamentos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdMedicamentos
End Sub

'mgaray20141003
Private Sub grdMorbilidadDesarollo_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Cell.Column.Key = "labConfHIS" Then
        If ValidarDiagnosticosCRED(Cell.Row.Cells("ID").Value, CStr(IIf(IsNull(NewValue), "", NewValue)), True, oRsDxDesarrollo.Bookmark) = False Then
            Cancel = True
        End If
    End If
End Sub

Private Sub grdMorbilidadDesarollo_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
    Call EliminarDxCrecimientoDesarrolloSeleccionado
End Sub

Private Sub grdMorbilidadDesarollo_Error(ByVal ErrorInfo As UltraGrid.SSError)
'On Error GoTo miError
    If ErrorInfo.Code = 16389 And ErrorInfo.DataError.Cell.Column.Key = "IdSubclasificacionDx" Then
        ErrorInfo.DisplayErrorDialog = False
        MsgBox "El valor del tipo de diagnóstico no es valido", vbInformation, "Validación"
    End If
'miError:
'    Err = 0
End Sub

'mgaray20141003
Private Sub grdMorbilidadDesarollo_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
'    InicializarLaGrilla grdMorbilidadDesarollo
End Sub

Private Sub grdMorbilidadDesarollo_LostFocus()
    Dim a As String
    a = "test"
End Sub

Private Sub grdMorbilidadFrec_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Cell.Column.Key = "labConfHIS" Then
        If ValidarDiagnosticosFrecuentes(Cell.Row.Cells("ID").Value, CStr(IIf(IsNull(NewValue), "", NewValue)), True, oRsMorbilidadFrec.Bookmark) = False Then
            Cancel = True
    '        MsgBox "DX y Lab de Morbilidad ya ha sido agregado", vbInformation, "Módulo Perinatal"
        End If
    End If
End Sub

Private Sub grdMorbilidadFrec_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
    Call EliminarDxMorbilidadSeleccionado
End Sub

Private Sub grdMorbilidadFrec_Error(ByVal ErrorInfo As UltraGrid.SSError)
'On Error GoTo miError
    If ErrorInfo.Code = 16389 And ErrorInfo.DataError.Cell.Column.Key = "IdSubclasificacionDx" Then
        ErrorInfo.DisplayErrorDialog = False
        MsgBox "El valor del tipo de diagnóstico no es valido", vbInformation, "Validación"
    End If
'miError:
'    Err = 0
End Sub

'mgaray20141003
Private Sub grdMorbilidadFrec_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
'    InicializarLaGrilla grdMorbilidadFrec
End Sub


Public Function DevuelveDatosGenerales() As DoPerinatalAtencion
    Dim oDoPerinatalAtencion As New DoPerinatalAtencion
    With oDoPerinatalAtencion
        .idModulo = lnIdModulo
        .idPerinatalAtencion = mo_idPerinatalAtencion
        .GrafXedadEnMeses = ml_EdadEnMeses
        .GrafYpercentilPE = lnPercentilPE
        .GrafYpercentilPT = lnPercentilPT
        .GrafYpercentilTE = lnPercentilTE
        .GrafYimc = lnPercentilIMC
        .FechaAtencion = Format(ml_FechaAtencion, sighentidades.DevuelveFechaSoloFormato_DMY)
    End With
    Set DevuelveDatosGenerales = oDoPerinatalAtencion
End Function

Public Function DevuelvePerinatalAtencionCred1() As DoPerinatalAtencionCred1
    Dim oDoPerinatalAtencionCred1 As New DoPerinatalAtencionCred1
    With oDoPerinatalAtencionCred1
        .AlimentacionComplementaria = IIf(chkAlimentacionComplementaria.Value = 1, True, False)
        .DemandaIndividual = IIf(chkDemandaIndividual.Value = 1, True, False)
        .EstimulacionTemprana = IIf(chkEstimulacionTemprana.Value = 1, True, False)
        .idAtencion = IIf(FraCred1.Enabled = True, ml_idAtencion, 0)
        .idModulo = lnIdModulo
        .idPerinatalAtencion = mo_idPerinatalAtencion
        .LactanciaMaterna = IIf(chkLactanciaMaterna.Value = 1, True, False)
        .MujerEdadReproductiva = IIf(chkMujerReproductiva.Value = 1, True, False)
        .MujerGestante = IIf(chkGestante.Value = 1, True, False)
        .PersonalSalud = IIf(chkPersonalSalud.Value = 1, True, False)
    End With
    Set DevuelvePerinatalAtencionCred1 = oDoPerinatalAtencionCred1
End Function

Public Function DevuelveDatosCred() As Recordset
    Dim oRsCred As New Recordset
    With oRsCred
          .Fields.Append "EdadEnAnios", adVarChar, 10, adFldIsNullable
          .Fields.Append "CredNumero", adInteger
          .Fields.Append "CredCheck", adVarChar, 1, adFldIsNullable
          .Fields.Append "idAtencion", adInteger
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    '
    If Cred1.Visible = True And Cred1.Enabled = True And Cred1.Text <> "" Then
       oRsCred.AddNew
       oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
       oRsCred.Fields!credNumero = 1
       oRsCred.Fields!credCheck = Cred1.Text
       oRsCred.Fields!idAtencion = ml_idAtencion
       oRsCred.Update
    End If
    If Cred2.Visible = True And Cred2.Enabled = True And Cred2.Text <> "" Then
       oRsCred.AddNew
       oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
       oRsCred.Fields!credNumero = 2
       oRsCred.Fields!credCheck = Cred2.Text
       oRsCred.Fields!idAtencion = ml_idAtencion
       oRsCred.Update
    End If
    If Cred3.Visible = True And Cred3.Enabled = True And Cred3.Text <> "" Then
       oRsCred.AddNew
       oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
       oRsCred.Fields!credNumero = 3
       oRsCred.Fields!credCheck = Cred3.Text
       oRsCred.Fields!idAtencion = ml_idAtencion
       oRsCred.Update
    End If
    If Cred4.Visible = True And Cred4.Enabled = True And Cred4.Text <> "" Then
       oRsCred.AddNew
       oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
       oRsCred.Fields!credNumero = 4
       oRsCred.Fields!credCheck = Cred4.Text
       oRsCred.Fields!idAtencion = ml_idAtencion
       oRsCred.Update
    End If
    If Cred5.Visible = True And Cred5.Enabled = True And Cred5.Text <> "" Then
       oRsCred.AddNew
       oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
       oRsCred.Fields!credNumero = 5
       oRsCred.Fields!credCheck = Cred5.Text
       oRsCred.Fields!idAtencion = ml_idAtencion
       oRsCred.Update
    End If
    If Cred6.Visible = True And Cred6.Enabled = True And Cred6.Text <> "" Then
       oRsCred.AddNew
       oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
       oRsCred.Fields!credNumero = 6
       oRsCred.Fields!credCheck = Cred6.Text
       oRsCred.Fields!idAtencion = ml_idAtencion
       oRsCred.Update
    End If
    If Cred7.Visible = True And Cred7.Enabled = True And Cred7.Text <> "" Then
       oRsCred.AddNew
       oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
       oRsCred.Fields!credNumero = 7
       oRsCred.Fields!credCheck = Cred7.Text
       oRsCred.Fields!idAtencion = ml_idAtencion
       oRsCred.Update
    End If
    If Cred8.Visible = True And Cred8.Enabled = True And Cred8.Text <> "" Then
       oRsCred.AddNew
       oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
       oRsCred.Fields!credNumero = 8
       oRsCred.Fields!credCheck = Cred8.Text
       oRsCred.Fields!idAtencion = ml_idAtencion
       oRsCred.Update
    End If
    If Cred9.Visible = True And Cred9.Enabled = True And Cred9.Text <> "" Then
       oRsCred.AddNew
       oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
       oRsCred.Fields!credNumero = 9
       oRsCred.Fields!credCheck = Cred9.Text
       oRsCred.Fields!idAtencion = ml_idAtencion
       oRsCred.Update
    End If
    If Cred10.Visible = True And Cred10.Enabled = True And Cred10.Text <> "" Then
       oRsCred.AddNew
       oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
       oRsCred.Fields!credNumero = 10
       oRsCred.Fields!credCheck = Cred10.Text
       oRsCred.Fields!idAtencion = ml_idAtencion
       oRsCred.Update
    End If
    If Cred11.Visible = True And Cred11.Enabled = True And Cred11.Text <> "" Then
       oRsCred.AddNew
       oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
       oRsCred.Fields!credNumero = 11
       oRsCred.Fields!credCheck = Cred11.Text
       oRsCred.Fields!idAtencion = ml_idAtencion
       oRsCred.Update
    End If
    If Cred12.Visible = True And Cred12.Enabled = True And Cred12.Text <> "" Then
       oRsCred.AddNew
       oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
       oRsCred.Fields!credNumero = 12
       oRsCred.Fields!credCheck = Cred12.Text
       oRsCred.Fields!idAtencion = ml_idAtencion
       oRsCred.Update
    End If
    '
    Set DevuelveDatosCred = oRsCred
End Function

Public Function DevuelveCptInmunizaciones() As Recordset
    'Set DevuelveCptInmunizaciones = oRsInmunizaciones
    Set DevuelveCptInmunizaciones = DevuelveCptInmunizacionesAtencionIntegral()
End Function
Public Function DevuelveCptFrecuentes() As Recordset
    'Set DevuelveCptFrecuentes = oRsCptFrecuentes
    Set DevuelveCptFrecuentes = DevuelveCptFrecuentesAtencionIntegral()
End Function
Public Function DevuelveDxDesarrollo() As Recordset
    'mgaray201412a
    If oRsDxDesarrollo Is Nothing Then
        Set DevuelveDxDesarrollo = Nothing
    Else
        Set DevuelveDxDesarrollo = oRsDxDesarrollo.Clone()
    End If
End Function
Public Function DevuelveDxMorbilidad() As Recordset
    'mgaray201412a
    If oRsMorbilidadFrec Is Nothing Then
        Set DevuelveDxMorbilidad = Nothing
    Else
        Set DevuelveDxMorbilidad = oRsMorbilidadFrec.Clone
    End If
End Function
'@implementar
Public Function DevuelveMedicamentos() As Recordset
    On Error GoTo errDM
'    If oRsFarmaciaMI.RecordCount > 0 Then
'       oRsFarmaciaMI.MoveFirst
'       Do While Not oRsFarmaciaMI.EOF
'          If oRsFarmaciaMI.Fields!seleccionar = True Then
'             oRsFarmaciaMI.Fields!idAtencion = ml_idAtencion
'             oRsFarmaciaMI.Update
'          End If
'          oRsFarmaciaMI.MoveNext
'       Loop
'    End If
'    Set DevuelveMedicamentos = oRsFarmaciaMI
    
    Set DevuelveMedicamentos = DevuelveMedicamentosAtencionIntegral()
    Exit Function
errDM:
    Set DevuelveMedicamentos = Nothing
End Function



Public Sub CargaDatosAcontroles(lnEdad As Integer, lnIdTipoEdad As Integer, lnPesoKg As Double, lnTallaCM As Long, oConexion As Connection)
    CargaDatosAcombos lnEdad, lnIdTipoEdad, oConexion
    '
    Dim oDoPerinatalAtencion As New DoPerinatalAtencion, oPerinatalAtencion As New PerinatalAtencion
    Dim oDoPerinatalAtencionCred1 As New DoPerinatalAtencionCred1, oPerinatalAtencionCred1 As New PerinatalAtencionCred1
    Dim oRsTmp As New Recordset
    Dim lnFor As Integer, lbEsDxPerinatal As Boolean
    '
    FraAdulto.Left = 1060
    FraAdulto.Visible = False
    If lnIdModulo = sighDesde18anios Then
        FraAdulto.Left = 60
        FraAdulto.Visible = True
    End If
    '
    Set oPerinatalAtencion.Conexion = oConexion
    Set oPerinatalAtencionCred1.Conexion = oConexion
    LimpiaInmunizaciones
    LimpiaCPTFrecuentes
    LimpiaDxDesarrollo
    LimpiaMorbilidadFrecuente True
    LimpiaCheckDeMedicamentos
    chkEstimulacionTemprana.Value = 0
    chkAlimentacionComplementaria.Value = 0
    chkLactanciaMaterna.Value = 0
    chkPersonalSalud.Value = 0
    chkDemandaIndividual.Value = 0
    chkMujerReproductiva.Value = 0
    chkGestante.Value = 0
    lnIdAtencionCred1 = 0
    mo_idPerinatalAtencion = 0
    If oPerinatalAtencion.SeleccionarPorIdAtencion(oDoPerinatalAtencion, ml_idAtencion) = True Then
    'If oPerinatalAtencion.SeleccionarPorIdPaciente(oDoPerinatalAtencion, ml_idPaciente, ml_FechaAtencion) = True Then
       mo_idPerinatalAtencion = oDoPerinatalAtencion.idPerinatalAtencion
       lnPercentilPE = oDoPerinatalAtencion.GrafYpercentilPE
       lnPercentilTE = oDoPerinatalAtencion.GrafYpercentilTE
       lnPercentilPT = oDoPerinatalAtencion.GrafYpercentilPT
       lnPercentilIMC = oDoPerinatalAtencion.GrafYimc
       'Carga Cred1
       oDoPerinatalAtencionCred1.idPerinatalAtencion = mo_idPerinatalAtencion
       If oPerinatalAtencionCred1.SeleccionarPorId(oDoPerinatalAtencionCred1) Then
            chkEstimulacionTemprana.Value = IIf(oDoPerinatalAtencionCred1.EstimulacionTemprana = True, 1, 0)
            chkAlimentacionComplementaria.Value = IIf(oDoPerinatalAtencionCred1.AlimentacionComplementaria = True, 1, 0)
            chkLactanciaMaterna.Value = IIf(oDoPerinatalAtencionCred1.LactanciaMaterna = True, 1, 0)
            chkPersonalSalud.Value = IIf(oDoPerinatalAtencionCred1.PersonalSalud = True, 1, 0)
            chkDemandaIndividual.Value = IIf(oDoPerinatalAtencionCred1.DemandaIndividual = True, 1, 0)
            chkMujerReproductiva.Value = IIf(oDoPerinatalAtencionCred1.MujerEdadReproductiva = True, 1, 0)
            chkGestante.Value = IIf(oDoPerinatalAtencionCred1.MujerGestante = True, 1, 0)
            lnIdAtencionCred1 = oDoPerinatalAtencionCred1.idAtencion
       End If
       '
       GraficoRegistraDatosParaFilaColumnas True
       '
       Set oRsTmp = mo_reglasComunes.PerinatalAtencionCptSeleccionarPorIdPerinatalAtencion(mo_idPerinatalAtencion, oConexion)
       If oRsTmp.RecordCount > 0 Then
          
          oRsTmp.MoveFirst
          Do While Not oRsTmp.EOF
             If oRsTmp.Fields!idLista = sghPerinatalListas.sighInmunizaciones Then
                If oRsTmp.Fields!CptEsAutomatico = False Then
                    For lnFor = 0 To cmbEligeInmunizacion.ListCount - 1
                        If cmbEligeInmunizacion.ListItems.Item(lnFor).Key = lcCombo & Trim(Str(oRsTmp.Fields!idProducto)) Then
                           cmbEligeInmunizacion.ListItems.Item(lnFor).Selected = True
                           Exit For
                        End If
                    Next
                    oRsInmunizaciones.AddNew
                    oRsInmunizaciones.Fields!Id = oRsTmp.Fields!idProducto
                    oRsInmunizaciones.Fields!procedimiento = oRsTmp.Fields!nombre
                    oRsInmunizaciones.Fields!idAtencion = oRsTmp.Fields!idAtencion
                    oRsInmunizaciones.Update
                End If
             Else
                For lnFor = 0 To cmbProcedimientosFrecuentes.ListCount - 1
                    If cmbProcedimientosFrecuentes.ListItems.Item(lnFor).Key = lcCombo & Trim(Str(oRsTmp.Fields!idProducto)) Then
                       cmbProcedimientosFrecuentes.ListItems.Item(lnFor).Selected = True
                       Exit For
                    End If
                Next
                oRsCptFrecuentes.AddNew
                oRsCptFrecuentes.Fields!Id = oRsTmp.Fields!idProducto
                oRsCptFrecuentes.Fields!procedimiento = oRsTmp.Fields!nombre
                oRsCptFrecuentes.Fields!idAtencion = oRsTmp.Fields!idAtencion
                oRsCptFrecuentes.Fields!labConfHIS = oRsTmp.Fields!labConfHIS
                oRsCptFrecuentes.Update
             End If
             oRsTmp.MoveNext
          Loop
          If oRsInmunizaciones.RecordCount > 0 Then oRsInmunizaciones.MoveFirst
          If oRsCptFrecuentes.RecordCount > 0 Then oRsCptFrecuentes.MoveFirst
       End If
       oRsTmp.Close
       '
       Set oRsTmp = mo_reglasComunes.PerinatalAtencionDxSeleccionarPorIdPerinatalAtencion(mo_idPerinatalAtencion, oConexion)
       If oRsTmp.RecordCount > 0 Then
          oRsTmp.MoveFirst
          Do While Not oRsTmp.EOF
             If oRsTmp.Fields!idLista = sghPerinatalListas.sighMorbilidadDesarrollo Then
                For lnFor = 0 To cmbDxDesarrollo.ListCount - 1
                    If cmbDxDesarrollo.ListItems.Item(lnFor).Key = lcCombo & Trim(Str(oRsTmp.Fields!idDiagnostico)) Then
                       cmbDxDesarrollo.ListItems.Item(lnFor).Selected = True
                       Exit For
                    End If
                Next
                oRsDxDesarrollo.AddNew
                oRsDxDesarrollo.Fields!Id = oRsTmp.Fields!idDiagnostico
                oRsDxDesarrollo.Fields!CodigoCIE2004 = oRsTmp.Fields!CodigoCIE2004
                oRsDxDesarrollo.Fields!DIAGNOSTICO = oRsTmp.Fields!Descripcion
                oRsDxDesarrollo.Fields!idAtencion = oRsTmp.Fields!idAtencion
                oRsDxDesarrollo.Fields!labConfHIS = oRsTmp.Fields!labConfHIS
                oRsDxDesarrollo.Fields!IdClasificacionDx = oRsTmp.Fields!IdClasificacionDx
                oRsDxDesarrollo.Fields!IdSubclasificacionDx = oRsTmp.Fields!IdSubclasificacionDx
                oRsDxDesarrollo.Update
             Else
                lbEsDxPerinatal = False
                For lnFor = 0 To cmbMorbilidadFrec.ListCount - 1
                    If cmbMorbilidadFrec.ListItems.Item(lnFor).Key = lcCombo & Trim(Str(oRsTmp.Fields!idDiagnostico)) Then
                       cmbMorbilidadFrec.ListItems.Item(lnFor).Selected = True
                       lbEsDxPerinatal = True
                       Exit For
                    End If
                Next
                oRsMorbilidadFrec.AddNew
                oRsMorbilidadFrec.Fields!Id = oRsTmp.Fields!idDiagnostico
                oRsMorbilidadFrec.Fields!CodigoCIE2004 = oRsTmp.Fields!CodigoCIE2004
                oRsMorbilidadFrec.Fields!DIAGNOSTICO = oRsTmp.Fields!Descripcion
                oRsMorbilidadFrec.Fields!idAtencion = oRsTmp.Fields!idAtencion
                oRsMorbilidadFrec.Fields!SeEligioConChek = True
                oRsMorbilidadFrec.Fields!EsDxPerinatal = lbEsDxPerinatal
                oRsMorbilidadFrec.Fields!labConfHIS = oRsTmp.Fields!labConfHIS
                oRsMorbilidadFrec.Fields!IdClasificacionDx = oRsTmp.Fields!IdClasificacionDx
                oRsMorbilidadFrec.Fields!IdSubclasificacionDx = oRsTmp.Fields!IdSubclasificacionDx
                oRsMorbilidadFrec.Update
             End If
             oRsTmp.MoveNext
          Loop
          If oRsDxDesarrollo.RecordCount > 0 Then oRsDxDesarrollo.MoveFirst
          If oRsMorbilidadFrec.RecordCount > 0 Then oRsMorbilidadFrec.MoveFirst
       End If
       oRsTmp.Close
       '
       If oRsFarmaciaMI.RecordCount > 0 Then
            Set oRsTmp = mo_reglasComunes.PerinatalAtencionMedicamentoSeleccionarPorIdPerinatalAtencion(mo_idPerinatalAtencion, oConexion)
            If oRsTmp.RecordCount > 0 Then
               oRsTmp.MoveFirst
               Do While Not oRsTmp.EOF
                  oRsFarmaciaMI.MoveFirst
                  oRsFarmaciaMI.Find "ID=" & oRsTmp.Fields!idProducto
                  If Not oRsFarmaciaMI.EOF Then
                     oRsFarmaciaMI.Fields!seleccionar = True
                     oRsFarmaciaMI.Fields!idAtencion = oRsTmp.Fields!idAtencion
                     oRsFarmaciaMI.Update
                  End If
                  oRsTmp.MoveNext
               Loop
            End If
            oRsTmp.Close
       End If
    Else
        'debb-2/3/2015*****inicio
        Dim oProcesos As New Procesos
        oProcesos.CalculaPercentiles lnPesoKg, lnTallaCM, ml_EdadEnMeses, _
                                     ml_idTipoSexo, lnEdadEnAniosEnAtencion, _
                                     lnPercentilPE, lnPercentilTE, lnPercentilPT, lnPercentilIMC, _
                                     lnPercentilPE_Z, lnPercentilTE_Z, lnPercentilPT_Z, lnPercentilIMC_Z
        Set oProcesos = Nothing
        GraficoRegistraDatosParaFilaColumnas True
        CargaDxAutomaticosParaMorbilidadEnDesarrollo lnPesoKg, lnTallaCM
        'debb-2/3/2015***** fin
    End If
    
    '
    If oRsFarmaciaMI.RecordCount > 0 Then oRsFarmaciaMI.MoveFirst
    '
    If oRsCredHistorico.RecordCount > 0 Then
       oRsCredHistorico.MoveFirst
       Do While Not oRsCredHistorico.EOF
          oRsCredHistorico.Delete
          oRsCredHistorico.Update
          oRsCredHistorico.MoveNext
       Loop
    End If
    Cred1.Text = "": Cred2.Text = "": Cred3.Text = "": Cred4.Text = "": Cred5.Text = "": Cred6.Text = ""
    Cred7.Text = "": Cred8.Text = "": Cred9.Text = "": Cred10.Text = "": Cred11.Text = "": Cred12.Text = ""
    lnIdAtencionCred = 0
    mo_Formulario.HabilitarDeshabilitar Cred1, True: mo_Formulario.HabilitarDeshabilitar Cred2, True
    mo_Formulario.HabilitarDeshabilitar Cred3, True: mo_Formulario.HabilitarDeshabilitar Cred4, True
    mo_Formulario.HabilitarDeshabilitar Cred5, True: mo_Formulario.HabilitarDeshabilitar Cred6, True
    mo_Formulario.HabilitarDeshabilitar Cred7, True: mo_Formulario.HabilitarDeshabilitar Cred8, True
    mo_Formulario.HabilitarDeshabilitar Cred9, True: mo_Formulario.HabilitarDeshabilitar Cred10, True
    mo_Formulario.HabilitarDeshabilitar Cred11, True: mo_Formulario.HabilitarDeshabilitar Cred12, True
    'mgaray
    BloqueoControlesAtencionCRED
    Dim lcEdad As String, lcCred As String, lnPos As Integer, lcIzquierda As String, lcCentro As String, lcDerecha As String
    Set oRsTmp = mo_reglasComunes.PerinatalAtencionCredSeleccionarPorIdPaciente(ml_IdPaciente, oConexion)
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          lcEdad = oRsTmp.Fields!edadEnAnios
          lcCred = Space(50)
          Do While Not oRsTmp.EOF And lcEdad = oRsTmp.Fields!edadEnAnios
             If lcEdadCredEnAtencion = oRsTmp.Fields!edadEnAnios Then
                'If ml_idAtencion >= oRsTmp.Fields!idAtencion Then
                
                    Select Case oRsTmp.Fields!credNumero
                    Case 1
                        If mo_idPerinatalAtencion <> oRsTmp.Fields!idPerinatalAtencion Then
                           mo_Formulario.HabilitarDeshabilitar Cred1, False
                        End If
                        Cred1.Text = oRsTmp.Fields!credCheck
                    Case 2
                        If mo_idPerinatalAtencion <> oRsTmp.Fields!idPerinatalAtencion Then
                           mo_Formulario.HabilitarDeshabilitar Cred2, False
                        End If
                        Cred2.Text = oRsTmp.Fields!credCheck
                    Case 3
                        If mo_idPerinatalAtencion <> oRsTmp.Fields!idPerinatalAtencion Then
                           mo_Formulario.HabilitarDeshabilitar Cred3, False
                        End If
                        Cred3.Text = oRsTmp.Fields!credCheck
                    Case 4
                        If mo_idPerinatalAtencion <> oRsTmp.Fields!idPerinatalAtencion Then
                           mo_Formulario.HabilitarDeshabilitar Cred4, False
                        End If
                        Cred4.Text = oRsTmp.Fields!credCheck
                    Case 5
                        If mo_idPerinatalAtencion <> oRsTmp.Fields!idPerinatalAtencion Then
                           mo_Formulario.HabilitarDeshabilitar Cred5, False
                        End If
                        Cred5.Text = oRsTmp.Fields!credCheck
                    Case 6
                        If mo_idPerinatalAtencion <> oRsTmp.Fields!idPerinatalAtencion Then
                           mo_Formulario.HabilitarDeshabilitar Cred6, False
                        End If
                        Cred6.Text = oRsTmp.Fields!credCheck
                    Case 7
                        If mo_idPerinatalAtencion <> oRsTmp.Fields!idPerinatalAtencion Then
                           mo_Formulario.HabilitarDeshabilitar Cred7, False
                        End If
                        Cred7.Text = oRsTmp.Fields!credCheck
                    Case 8
                        If mo_idPerinatalAtencion <> oRsTmp.Fields!idPerinatalAtencion Then
                           mo_Formulario.HabilitarDeshabilitar Cred8, False
                        End If
                        Cred8.Text = oRsTmp.Fields!credCheck
                    Case 9
                        If mo_idPerinatalAtencion <> oRsTmp.Fields!idPerinatalAtencion Then
                           mo_Formulario.HabilitarDeshabilitar Cred9, False
                        End If
                        Cred9.Text = oRsTmp.Fields!credCheck
                    Case 10
                        If mo_idPerinatalAtencion <> oRsTmp.Fields!idPerinatalAtencion Then
                           mo_Formulario.HabilitarDeshabilitar Cred10, False
                        End If
                        Cred10.Text = oRsTmp.Fields!credCheck
                    Case 11
                        If mo_idPerinatalAtencion <> oRsTmp.Fields!idPerinatalAtencion Then
                           mo_Formulario.HabilitarDeshabilitar Cred11, False
                        End If
                        Cred11.Text = oRsTmp.Fields!credCheck
                    Case 12
                        If mo_idPerinatalAtencion <> oRsTmp.Fields!idPerinatalAtencion Then
                           mo_Formulario.HabilitarDeshabilitar Cred12, False
                        End If
                        Cred12.Text = oRsTmp.Fields!credCheck
                    End Select
                    If mo_idPerinatalAtencion = oRsTmp.Fields!idPerinatalAtencion Then
                       lnIdAtencionCred = oRsTmp.Fields!idAtencion
                    End If
                'End If
             Else
                '"01xx02xx03xx04xx05xx06xx07xx08xx09xx10xx11xx12xx13"
                ' 12345678901234567890123456789012345678901234567890
                lnPos = (4 * (oRsTmp.Fields!credNumero - 1)) + 1
                lcIzquierda = Left(lcCred, lnPos - 1)
                lcCentro = oRsTmp.Fields!credCheck & String(1, " ")
                lcDerecha = Mid(lcCred, lnPos + 2, 50)
                lcCred = lcIzquierda & lcCentro & lcDerecha
             End If
             oRsTmp.MoveNext
             If oRsTmp.EOF Then
                Exit Do
             End If
          Loop
          If Trim(lcCred) <> "" Then
                lcCred = Space(2) & lcCred
                oRsCredHistorico.AddNew
                oRsCredHistorico.Fields!Anio = lcEdad
                oRsCredHistorico.Fields!Control = lcCred
                oRsCredHistorico.Update
          End If
       Loop
    End If
    oRsTmp.Close
    
    '
    Set oDoPerinatalAtencion = Nothing
    Set oPerinatalAtencion = Nothing
    Set oRsTmp = Nothing
    Set oDoPerinatalAtencionCred1 = Nothing
    Set oPerinatalAtencionCred1 = Nothing
    '
     CargaGraficoChartSpace True
     'mgaray201411e
     CargaGraficoTallaEdad True
     CargaGraficoPesoEdad True
End Sub
'mgaray201411e
Public Sub ActualizaGraficoYDiagnosticosAutomaticamente(oDOAtencionesCE As SIGHComun.DOAtencionesCE)
'Public Sub ActualizaGraficoYDiagnosticosAutomaticamente(lnPesoKg As Double, lnTallaCM As Long)
    'As SIGHComun.DOAtencionesCE
    
    'mgaray201411e
    Dim lnPesoKg As Double, lnTallaCM As Long
    
    Set mo_DOAtencionesCE = oDOAtencionesCE
    lnPesoKg = oDOAtencionesCE.triajePeso
    lnTallaCM = oDOAtencionesCE.triajeTalla
    
    If lnPesoKg > 0 And lnTallaCM > 0 Then
       'debb-2/3/2015****inicio
       Dim oProcesos As New Procesos
       oProcesos.CalculaPercentiles lnPesoKg, lnTallaCM, ml_EdadEnMeses, _
                                     ml_idTipoSexo, lnEdadEnAniosEnAtencion, _
                                     lnPercentilPE, lnPercentilTE, lnPercentilPT, lnPercentilIMC, _
                                     lnPercentilPE_Z, lnPercentilTE_Z, lnPercentilPT_Z, lnPercentilIMC_Z
       Set oProcesos = Nothing
       'debb-2/3/2015****fin
       GraficoRegistraDatosParaFilaColumnas False
       CargaDxAutomaticosParaMorbilidadEnDesarrollo lnPesoKg, lnTallaCM
       '
       Select Case lnIdModulo
       Case sighHasta28Dias
       Case sighDesde29diasHasta1anio
       Case sighDesde1Hasta4anios
       Case sighDesde5Hasta9anios
       Case sighDesde10Hasta11anios
       Case sighDesde12Hasta17anios
       Case sighDesde18anios
       End Select
       CargaGraficoChartSpace False
       CargaGraficoTallaEdad False
       CargaGraficoPesoEdad False
    End If
End Sub

''Actualiza valores y Devuelve percentil de la ATENCION ACTUAL DEL PACIENTE
'Sub CalculaPercentiles(lnPesoKg As Double, lnTallaCM As Long)
'    lnPercentilPE = lnPercentilNull: lnPercentilTE = lnPercentilNull: lnPercentilPT = lnPercentilNull: lnPercentilIMC = lnPercentilNull
'    lnPercentilPE_Z = lnPercentilNull: lnPercentilTE_Z = lnPercentilNull: lnPercentilPT_Z = lnPercentilNull: lnPercentilIMC_Z = lnPercentilNull
'    If lnPesoKg > 0 And lnTallaCM > 0 Then
'       On Error Resume Next
'       Dim EXL As Excel.Application
'       Set EXL = New Excel.Application
'       Dim W As Excel.Workbook
'       Dim lnEdadUSAmaxima
'       lnEdadUSAmaxima = 5
'       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
'           Set W = EXL.Workbooks.Open(App.Path & "\Plantillas\cred.xls")       'usa
'       Else
'           Set W = EXL.Workbooks.Open(App.Path & "\Plantillas\cred who.xls")    'oms
'       End If
'       Dim s As Excel.Worksheet
'       Dim lnEdadEnMesesMasPuntoCinco As Double, lnMinimo As Double, lnMaximo As Double, lnIMC As Double
'       Dim lnTallaEnCmMasPuntoCinco As Double
'       lnEdadEnMesesMasPuntoCinco = ml_EdadEnMeses + 0.5
'       lnTallaEnCmMasPuntoCinco = lnTallaCM + 0.5
'       'Peso Edad
'       Set s = W.Sheets("P-E")
'
'       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
'          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
'          lnMaximo = s.Cells(243, IIf(ml_idTipoSexo = 1, 2, 18)).Value
'          lnPercentilPE = lnPercentilNull
'          s.Cells(246, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnEdadEnMesesMasPuntoCinco
'          s.Cells(247, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnPesoKg
'          lnPercentilPE = s.Cells(255, IIf(ml_idTipoSexo = 1, 3, 19)).Value
'          lnPercentilPE_Z = s.Cells(254, IIf(ml_idTipoSexo = 1, 3, 19)).Value
'       Else
'          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
'          lnMaximo = s.Cells(1858, IIf(ml_idTipoSexo = 1, 2, 18)).Value
'          lnPercentilPE = lnPercentilNull
'          s.Cells(1861, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnEdadEnMesesMasPuntoCinco
'          s.Cells(1862, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnPesoKg
'          lnPercentilPE = s.Cells(1870, IIf(ml_idTipoSexo = 1, 4, 20)).Value
'          lnPercentilPE_Z = s.Cells(1869, IIf(ml_idTipoSexo = 1, 4, 20)).Value
'       End If
'       'Talla Edad
'       Set s = W.Sheets("T-E")
'       'If lnEdadEnAniosEnAtencion <= lnEdadUSAmaxima Then
'       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
'          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
'          lnMaximo = s.Cells(243, IIf(ml_idTipoSexo = 1, 2, 18)).Value
'          lnPercentilTE = lnPercentilNull
'          s.Cells(246, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnEdadEnMesesMasPuntoCinco
'          s.Cells(247, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnTallaCM
'          lnPercentilTE = s.Cells(255, IIf(ml_idTipoSexo = 1, 3, 19)).Value
'          lnPercentilTE_Z = s.Cells(254, IIf(ml_idTipoSexo = 1, 3, 19)).Value
'       Else
'          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
'          lnMaximo = s.Cells(1858, IIf(ml_idTipoSexo = 1, 2, 18)).Value
'          lnPercentilTE = lnPercentilNull
'          s.Cells(1861, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnEdadEnMesesMasPuntoCinco
'          s.Cells(1862, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnTallaCM
'          lnPercentilTE = s.Cells(1870, IIf(ml_idTipoSexo = 1, 4, 20)).Value
'          lnPercentilTE_Z = s.Cells(1869, IIf(ml_idTipoSexo = 1, 4, 20)).Value
'       End If
'       'Peso Talla
'       Set s = W.Sheets("P-T")
'       'If lnEdadEnAniosEnAtencion <= lnEdadUSAmaxima Then
'       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
'          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
'          lnMaximo = s.Cells(61, IIf(ml_idTipoSexo = 1, 2, 18)).Value
'          lnPercentilPT = lnPercentilNull
'          s.Cells(64, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnTallaEnCmMasPuntoCinco
'          s.Cells(65, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnPesoKg
'          lnPercentilPT = s.Cells(73, IIf(ml_idTipoSexo = 1, 3, 19)).Value
'          lnPercentilPT_Z = s.Cells(72, IIf(ml_idTipoSexo = 1, 3, 19)).Value
'       Else
'          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
'          lnMaximo = s.Cells(652, IIf(ml_idTipoSexo = 1, 2, 18)).Value
'          lnPercentilPT = lnPercentilNull
'          s.Cells(655, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnTallaEnCmMasPuntoCinco
'          s.Cells(656, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnPesoKg
'          lnPercentilPT = s.Cells(664, IIf(ml_idTipoSexo = 1, 4, 20)).Value
'          lnPercentilPT_Z = s.Cells(663, IIf(ml_idTipoSexo = 1, 4, 20)).Value
'       End If
'       'Edad IMC
'       lnIMC = Round((lnPesoKg / (lnTallaCM * lnTallaCM)), 0)
'       Set s = W.Sheets("E-IMC")
'       'If lnEdadEnAniosEnAtencion <= lnEdadUSAmaxima Then
'       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
'          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 19)).Value
'          lnMaximo = s.Cells(220, IIf(ml_idTipoSexo = 1, 2, 19)).Value
'          lnPercentilIMC = lnPercentilNull
'          s.Cells(223, IIf(ml_idTipoSexo = 1, 4, 21)).Value = lnEdadEnMesesMasPuntoCinco
'          s.Cells(224, IIf(ml_idTipoSexo = 1, 4, 21)).Value = lnIMC
'          lnPercentilIMC = s.Cells(232, IIf(ml_idTipoSexo = 1, 3, 20)).Value
'          lnPercentilIMC_Z = s.Cells(231, IIf(ml_idTipoSexo = 1, 3, 20)).Value
'       Else
'          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
'          lnMaximo = s.Cells(1858, IIf(ml_idTipoSexo = 1, 2, 18)).Value
'          lnPercentilIMC = lnPercentilNull
'          s.Cells(1861, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnEdadEnMesesMasPuntoCinco
'          s.Cells(1862, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnIMC
'          lnPercentilIMC = s.Cells(1870, IIf(ml_idTipoSexo = 1, 4, 20)).Value
'          lnPercentilIMC_Z = s.Cells(1869, IIf(ml_idTipoSexo = 1, 4, 20)).Value
'       End If
'       W.Close False
'       Set s = Nothing
'       Set W = Nothing
'       Set EXL = Nothing
'    End If
'End Sub

Sub GraficoRegistraDatosParaFilaColumnas(lbSeCargaPercentilHistoricos As Boolean)
    If lbSeCargaPercentilHistoricos = True Then
       LimpiaPercentil
       Dim oRsTmp As New Recordset
       Set oRsTmp = mo_reglasComunes.PerinatalAtencionSeleccionarPorIdPaciente(ml_IdPaciente)
       If oRsTmp.RecordCount > 0 Then
          oRsTmp.MoveFirst
          Do While Not oRsTmp.EOF
             If ml_idAtencion >= oRsTmp.Fields!idAtencion Then
                oRsPercentil.AddNew
                oRsPercentil.Fields!idAtencion = oRsTmp.Fields!idAtencion
                oRsPercentil.Fields!EdadEnMeses = IIf(IsNull(oRsTmp.Fields!GrafXedadEnMeses), 0, oRsTmp.Fields!GrafXedadEnMeses)
                oRsPercentil.Fields!PercentilPE = IIf(IsNull(oRsTmp.Fields!GrafYpercentilPE), 0, oRsTmp.Fields!GrafYpercentilPE)
                oRsPercentil.Fields!PercentilTE = IIf(IsNull(oRsTmp.Fields!GrafYpercentilTE), 0, oRsTmp.Fields!GrafYpercentilTE)
                oRsPercentil.Fields!PercentilPT = IIf(IsNull(oRsTmp.Fields!GrafYpercentilPT), 0, oRsTmp.Fields!GrafYpercentilPT)
                oRsPercentil.Fields!PercentilIMC = IIf(IsNull(oRsTmp.Fields!GrafYimc), 0, oRsTmp.Fields!GrafYimc)
                oRsPercentil.Update
             End If
             oRsTmp.MoveNext
          Loop
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
    End If
    If oRsPercentil.RecordCount > 0 Then
       oRsPercentil.MoveFirst
       oRsPercentil.Find "idAtencion=" & ml_idAtencion
    End If
    If oRsPercentil.EOF Then
       oRsPercentil.AddNew
       oRsPercentil.Fields!idAtencion = ml_idAtencion
       oRsPercentil.Fields!EdadEnMeses = ml_EdadEnMeses
    End If
    oRsPercentil.Fields!PercentilPE = lnPercentilPE
    oRsPercentil.Fields!PercentilTE = lnPercentilTE
    oRsPercentil.Fields!PercentilPT = lnPercentilPT
    oRsPercentil.Fields!PercentilIMC = lnPercentilIMC
    oRsPercentil.Update
    'GraficoPresentaLineasEnGrafico
    'btnRefrescaGrafico_Click
   
End Sub







Sub CargaGraficoChartSpace(lbActualizaDesdeInicioRs As Boolean)
    Dim lnFor As Integer
    If lbActualizaDesdeInicioRs = True Then
            xValues = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250, 280)
            yValuesPT = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250, 280)
            yValuesTE = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250, 280)
            yValuesPE = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250, 280)
            ReDim xValues(oRsPercentil.RecordCount - 1)
            lnNroPuntosGraficos = oRsPercentil.RecordCount - 1
            oRsPercentil.MoveLast
            For lnFor = (oRsPercentil.RecordCount - 1) To 0 Step -1
               xValues(lnFor) = oRsPercentil.Fields!EdadEnMeses
               yValuesPT(lnFor) = oRsPercentil.Fields!PercentilPT
               yValuesTE(lnFor) = oRsPercentil.Fields!PercentilTE
               yValuesPE(lnFor) = oRsPercentil.Fields!PercentilPE
               oRsPercentil.MovePrevious
            Next
    Else
            lnFor = lnNroPuntosGraficos
            If lnFor < 0 Then
               lnFor = 0
            End If
            yValuesPT(lnFor) = lnPercentilPT
            yValuesTE(lnFor) = lnPercentilTE
            yValuesPE(lnFor) = lnPercentilPE
    End If
    '
    ChartSpace1.Clear
    ChartSpace1.DisplayToolbar = False
    Set owcChart = ChartSpace1.Charts.Add
    owcChart.HasTitle = True
    owcChart.Title.Caption = "Edad en Semanas (X), PT(rojo),  TE(verde),  PE(amarillo)"
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
    '
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "idTipoFinanciamiento"
        .SetData chDimCategories, chDataLiteral, xValues
        .SetData chDimValues, chDataLiteral, yValuesPT
        .Type = chChartTypeLineMarkers
        .Line.Color = vbRed
        .Line.Weight = 3
        .Marker.Style = chMarkerStyleCircle
        .Line.DashStyle = chLineSolid
        .DataLabelsCollection.Add
    End With
    '
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "idTipoFinanciamiento"
        .SetData chDimCategories, chDataLiteral, xValues
        .SetData chDimValues, chDataLiteral, yValuesTE
        .Type = chChartTypeLineMarkers
        .Line.Color = vbGreen
        .Line.Weight = 3
        .Marker.Style = chMarkerStyleCircle
        .Line.DashStyle = chLineSolid
        .DataLabelsCollection.Add
    End With
    '
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "idTipoFinanciamiento"
        .SetData chDimCategories, chDataLiteral, xValues
        .SetData chDimValues, chDataLiteral, yValuesPE
        .Type = chChartTypeLineMarkers
        .Line.Color = vbYellow
        .Line.Weight = 3
        .Marker.Style = chMarkerStyleCircle
        .Line.DashStyle = chLineSolid
        .DataLabelsCollection.Add
    End With
End Sub


'@Implementar
Sub HabilitaDeshabilita(lcFrame As String, lbEstado As Boolean)
    Select Case lcFrame
    Case "Inmunizaciones"
         cmbEligeInmunizacion.Enabled = lbEstado
         btnQuitarInmunizacion.Enabled = lbEstado
         'FraInmunizaciones.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         mo_Formulario.HabilitarDeshabilitar FraInmunizaciones, lbEstado
         grdInmunizaciones.Appearance.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         frAtenInteInmunizaciones.Tag = IIf(lbEstado = True, 1, 0)
    Case "OtrosCpt"
         cmbProcedimientosFrecuentes.Enabled = lbEstado
         btnQuitaOtrosProcedimientos.Enabled = lbEstado
         'FraOtrosCpt.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         mo_Formulario.HabilitarDeshabilitar FraOtrosCpt, lbEstado
         grdCptFrecuentes.Appearance.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         frAtenInteTamizajes.Tag = IIf(lbEstado = True, 1, 0)
    Case "DxDesarrollo"
         cmbDxDesarrollo.Enabled = lbEstado
         btnQuitaDxDesarrollo.Enabled = lbEstado
         'FraDxDesarrollo.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         mo_Formulario.HabilitarDeshabilitar FraDxDesarrollo, lbEstado
         'mgaray20141012
         mo_Formulario.HabilitarDeshabilitar btnActualizarSesionesPendientes, lbEstado
         grdMorbilidadDesarollo.Appearance.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         
         frAtenInteDesarrollo.Tag = IIf(lbEstado = True, 1, 0)
    Case "DxMorbilidad"
         cmbMorbilidadFrec.Enabled = lbEstado
         btnQuitaDxMorbilidad.Enabled = lbEstado
         btnBusquedaDiagnostico.Enabled = lbEstado
         'FraDxMorbilidad.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         mo_Formulario.HabilitarDeshabilitar FraDxMorbilidad, lbEstado
         grdMorbilidadFrec.Appearance.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
    Case "Cred"
         mo_Formulario.HabilitarDeshabilitar FraCred, lbEstado
         'FraCred.Enabled = lbEstado
         'FraCred.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         lblCred.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         lbl1.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         lbl2.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         lbl3.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         lbl4.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         lbl5.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         lbl6.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         lbl7.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         lbl8.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         lbl9.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         lbl10.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         lbl11.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         lbl12.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
    Case "Cred1"
         mo_Formulario.HabilitarDeshabilitar FraCred1, lbEstado
         'FraCred1.Enabled = lbEstado
         chkEstimulacionTemprana.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         chkAlimentacionComplementaria.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         chkLactanciaMaterna.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         chkPersonalSalud.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         chkDemandaIndividual.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         chkMujerReproductiva.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         chkGestante.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
    Case "Medicamentos"
         mo_Formulario.HabilitarDeshabilitar FraMedicamentos, lbEstado
         'FraMedicamentos.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         grdMedicamentos.Bands(0).Columns("Seleccionar").Activation = IIf(lbEstado = True, ssActivationAllowEdit, ssActivationActivateNoEdit)
         grdMedicamentos.Bands(0).Columns("Medicamento").Activation = ssActivationActivateNoEdit
         grdMedicamentos.Appearance.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         frAtenInteSuplemento.Tag = IIf(lbEstado = True, 1, 0)
    End Select
End Sub

'Solamente carga Dx automaticos si esta vacio (es decir no se eligió nada aun)@revisar
Sub CargaDxAutomaticosParaMorbilidadEnDesarrollo(lnPesoKg As Double, lnTallaCM As Long)
    Exit Sub
    If oRsDxDesarrolloAutomaticos.RecordCount > 0 And oRsDxDesarrollo.RecordCount = 0 And lnPesoKg > 0 And lnTallaCM > 0 And FraDxDesarrollo.Enabled = True Then
       Dim lbContinuar As Boolean
       'mgaray201411a
       Dim oDODiagnostico As New DODiagnostico
       
       oRsDxDesarrolloAutomaticos.MoveFirst
       Do While Not oRsDxDesarrolloAutomaticos.EOF
            Select Case lnIdModulo
            Case sighHasta28Dias
                 'Usa el PESO
                 If (lnPesoKg * 1000) >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And (lnPesoKg * 1000) <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                    'mgaray201411a
                    oDODiagnostico.idDiagnostico = oRsDxDesarrolloAutomaticos.Fields!idDiagnostico
                    oDODiagnostico.CodigoCIE2004 = oRsDxDesarrolloAutomaticos.Fields!CodigoCIE2004
                    oDODiagnostico.Descripcion = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
                    Call AgregarDxCrecimientoDesarrolloSeleccionado(oDODiagnostico, "")
                    
'                    oRsDxDesarrollo.AddNew
'                    oRsDxDesarrollo.Fields!ID = oRsDxDesarrolloAutomaticos.Fields!IdDiagnostico
'                    oRsDxDesarrollo.Fields!DIAGNOSTICO = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
'                    'mgaray20141022
'                    oRsDxDesarrollo.Fields!CodigoCIE2004 = oRsDxDesarrolloAutomaticos.Fields!CodigoCIE2004
'                    oRsDxDesarrollo.Fields!idAtencion = ml_idAtencion
'                    oRsDxDesarrollo.Update
                 End If
            Case sighDesde29diasHasta1anio
                 'Usa Z de Peso-Edad,Peso-Talla,Talla-Edad, Cie10 de mas de 4 digitos
                 lbContinuar = False
                 Select Case UCase(oRsDxDesarrolloAutomaticos.Fields!cie10his)
                 Case "E343", "E3431", "E3441", "E344"
                     If lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                     End If
                 Case "E660", "E669"
                     If (lnPercentilPE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilPE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) _
                        Or (lnPercentilPT_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilPT_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                     End If
                 Case Else
                     If (lnPercentilPE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilPE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) _
                        Or (lnPercentilPT_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilPT_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) _
                        Or (lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                     End If
                 End Select
                 If lbContinuar = True Then
                    'mgaray201411a
                    oDODiagnostico.idDiagnostico = oRsDxDesarrolloAutomaticos.Fields!idDiagnostico
                    oDODiagnostico.CodigoCIE2004 = oRsDxDesarrolloAutomaticos.Fields!CodigoCIE2004
                    oDODiagnostico.Descripcion = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
                    Call AgregarDxCrecimientoDesarrolloSeleccionado(oDODiagnostico, "")
                    
'                    oRsDxDesarrollo.AddNew
'                    oRsDxDesarrollo.Fields!ID = oRsDxDesarrolloAutomaticos.Fields!IdDiagnostico
'                    oRsDxDesarrollo.Fields!DIAGNOSTICO = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
'                    'mgaray20141022
'                    oRsDxDesarrollo.Fields!CodigoCIE2004 = oRsDxDesarrolloAutomaticos.Fields!CodigoCIE2004
'                    oRsDxDesarrollo.Fields!idAtencion = ml_idAtencion
'                    oRsDxDesarrollo.Update
                 End If
            Case sighDesde1Hasta4anios
                 'Usa Z de Peso-Edad,Peso-Talla,Talla-Edad, Cie10 de mas de 4 digitos
                 lbContinuar = False
                 Select Case UCase(oRsDxDesarrolloAutomaticos.Fields!cie10his)
                 Case "E343", "E3431"
                     If lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                     End If
                 Case Else
                     If (lnPercentilPE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilPE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) _
                        Or (lnPercentilPT_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilPT_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) _
                        Or (lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                     End If
                 End Select
                 If lbContinuar = True Then
                    'mgaray201411a
                    oDODiagnostico.idDiagnostico = oRsDxDesarrolloAutomaticos.Fields!idDiagnostico
                    oDODiagnostico.CodigoCIE2004 = oRsDxDesarrolloAutomaticos.Fields!CodigoCIE2004
                    oDODiagnostico.Descripcion = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
                    Call AgregarDxCrecimientoDesarrolloSeleccionado(oDODiagnostico, "")
                    
'                    oRsDxDesarrollo.AddNew
'                    oRsDxDesarrollo.Fields!ID = oRsDxDesarrolloAutomaticos.Fields!IdDiagnostico
'                    oRsDxDesarrollo.Fields!DIAGNOSTICO = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
'                    'mgaray20141022
'                    oRsDxDesarrollo.Fields!CodigoCIE2004 = oRsDxDesarrolloAutomaticos.Fields!CodigoCIE2004
'                    oRsDxDesarrollo.Fields!idAtencion = ml_idAtencion
'                    oRsDxDesarrollo.Update
                 End If
            Case sighDesde5Hasta9anios
                 'Usa Z de IMC,Talla-Edad, Cie10 de mas de 4 digitos
                 lbContinuar = False
                 Select Case UCase(oRsDxDesarrolloAutomaticos.Fields!cie10his)
                 Case "Z006"
                     If (lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) _
                        And (lnPercentilTE_Z >= 10 And lnPercentilTE_Z <= 90) Then
                        lbContinuar = True
                     End If
                 Case "E46X", "E660", "E669"
                     If lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                     End If
                 Case "E3431", "E3441", "E344"
                     If lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                     End If
                 Case Else
                     If (lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) _
                        Or (lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                     End If
                 End Select
                 If lbContinuar = True Then
                    'mgaray201411a
                    oDODiagnostico.idDiagnostico = oRsDxDesarrolloAutomaticos.Fields!idDiagnostico
                    oDODiagnostico.CodigoCIE2004 = oRsDxDesarrolloAutomaticos.Fields!CodigoCIE2004
                    oDODiagnostico.Descripcion = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
                    Call AgregarDxCrecimientoDesarrolloSeleccionado(oDODiagnostico, "")
                    
'                    oRsDxDesarrollo.AddNew
'                    oRsDxDesarrollo.Fields!ID = oRsDxDesarrolloAutomaticos.Fields!IdDiagnostico
'                    oRsDxDesarrollo.Fields!DIAGNOSTICO = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
'                    'mgaray20141022
'                    oRsDxDesarrollo.Fields!CodigoCIE2004 = oRsDxDesarrolloAutomaticos.Fields!CodigoCIE2004
'                    oRsDxDesarrollo.Fields!idAtencion = ml_idAtencion
'                    oRsDxDesarrollo.Update
                 End If
            Case sighDesde10Hasta11anios
                 'Usa Z de IMC,Talla-Edad, Cie10 de mas de 4 digitos
                 lbContinuar = False
                 Select Case UCase(oRsDxDesarrolloAutomaticos.Fields!cie10his)
                 Case "Z006"
                     If (lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) _
                        And (lnPercentilTE_Z >= 10 And lnPercentilTE_Z <= 90) Then
                        lbContinuar = True
                     End If
                 Case "E46X", "E660", "E669"
                     If lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                     End If
                 Case "E3431", "E3441", "E344"
                     If lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                     End If
                 Case Else
                     If (lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) _
                        Or (lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                     End If
                 End Select
                 If lbContinuar = True Then
                    'mgaray201411a
                    oDODiagnostico.idDiagnostico = oRsDxDesarrolloAutomaticos.Fields!idDiagnostico
                    oDODiagnostico.CodigoCIE2004 = oRsDxDesarrolloAutomaticos.Fields!CodigoCIE2004
                    oDODiagnostico.Descripcion = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
                    Call AgregarDxCrecimientoDesarrolloSeleccionado(oDODiagnostico, "")
                    
                    
'                    oRsDxDesarrollo.AddNew
'                    oRsDxDesarrollo.Fields!ID = oRsDxDesarrolloAutomaticos.Fields!IdDiagnostico
'                    oRsDxDesarrollo.Fields!DIAGNOSTICO = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
'                    'mgaray20141022
'                    oRsDxDesarrollo.Fields!CodigoCIE2004 = oRsDxDesarrolloAutomaticos.Fields!CodigoCIE2004
'                    oRsDxDesarrollo.Fields!idAtencion = ml_idAtencion
'                    oRsDxDesarrollo.Update
                 End If
            Case sighDesde12Hasta17anios
                 'Usa Z de IMC,Talla-Edad, Cie10 de mas de 4 digitos
                 lbContinuar = False
                 Select Case UCase(oRsDxDesarrolloAutomaticos.Fields!cie10his)
                 Case "Z006"
                     If (lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) _
                        And (lnPercentilTE_Z >= 10 And lnPercentilTE_Z <= 90) Then
                        lbContinuar = True
                     End If
                 Case "E46X", "E660", "E669"
                     If lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                     End If
                 Case "E3431", "E3441", "E344"
                     If lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                     End If
                 Case Else
                     If (lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) _
                        Or (lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                     End If
                 End Select
                 If lbContinuar = True Then
                    'mgaray201411a
                    oDODiagnostico.idDiagnostico = oRsDxDesarrolloAutomaticos.Fields!idDiagnostico
                    oDODiagnostico.CodigoCIE2004 = oRsDxDesarrolloAutomaticos.Fields!CodigoCIE2004
                    oDODiagnostico.Descripcion = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
                    Call AgregarDxCrecimientoDesarrolloSeleccionado(oDODiagnostico, "")
                    
'                    oRsDxDesarrollo.AddNew
'                    oRsDxDesarrollo.Fields!ID = oRsDxDesarrolloAutomaticos.Fields!IdDiagnostico
'                    oRsDxDesarrollo.Fields!DIAGNOSTICO = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
'                    'mgaray20141022
'                    oRsDxDesarrollo.Fields!CodigoCIE2004 = oRsDxDesarrolloAutomaticos.Fields!CodigoCIE2004
'                    oRsDxDesarrollo.Fields!idAtencion = ml_idAtencion
'                    oRsDxDesarrollo.Update
                 End If
            Case sighDesde18anios
            End Select
            oRsDxDesarrolloAutomaticos.MoveNext
       Loop
    End If
End Sub


'preguntas:
'- si son 2 servicios donde se consulta el paciente. ejm Cred y Medicina, el Dx de medicina debe copiarse como dx en Cred ?
'- terminar de registrar CPT y DX que no se ha encontrado en tablas (od=otra descripcion, ya está registrado aquí como 'consulta en consultorios externos')
Private Sub grdPlanDesarrollo_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not IsNull(NewValue) And Cell.Column.Key = "FechaProgramada" Then
        If NewValue < md_fechaActual Then
            MsgBox "Fecha no puede ser menor que la fecha actual", vbInformation, "Advertencia"
            NewValue = Cell.Value
        End If
    End If
End Sub

Private Sub grdPlanDesarrollo_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdPlanDesarrollo_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPlanDesarrollo.ViewStyleBand = ssViewStyleBandVertical
    'evitar que los cambios en las celdas editables se hagan directamente en la base de datos
    grdPlanDesarrollo.UpdateMode = ssUpdateOnUpdate
    grdPlanDesarrollo.CollapseAll
    'Cabecera de grupo
    With Layout.Bands(0)
        '.ColHeadersVisible = False
        'establecer etiqueta de columnas y formato
        .Columns("IdPlanAtencion").Header.Caption = "Id Plan"
        
        .Columns("FechaProgramada").Header.Caption = "F. Programada"
        .Columns("FechaProgramada").Width = 1400
                
        .Columns("FechaEjecucion").Header.Caption = "F. Ejecución"
        .Columns("FechaEjecucion").Width = 1400
        
        .Columns("NumeroSesion").Header.Caption = "N° Sesión"
        .Columns("NumeroSesion").Width = 1000
        
        .Columns("Evaluacion").Header.Caption = "Evaluación"
        .Columns("Evaluacion").Width = 1200
        
        .Columns("Descripcion").Header.Caption = "Edad"
        .Columns("Descripcion").Width = grdPlanDesarrollo.Width - 1000 - _
                                    .Columns("FechaProgramada").Width - _
                                    .Columns("FechaEjecucion").Width - _
                                    .Columns("NumeroSesion").Width - _
                                    .Columns("Evaluacion").Width
        
        'ocultar columnas
        Call mo_Apariencia.ocultarColumnas(Layout, 0, "IdPlanDesarrolloPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "EdadAnio", "EdadMes", "EdadDia", _
                                        "IdEstablecimiento", "Establecimiento")
        
        
        'desactivar edicion de columnas
        Call mo_Apariencia.modificarActivationColumnas(Layout, 0, ssActivationActivateNoEdit, "IdPlanDesarrolloPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "Descripcion", "EdadAnio", "EdadMes", "EdadDia", _
                                        "FechaEjecucion", "NumeroSesion", "Evaluacion", _
                                        "IdEstablecimiento", "Establecimiento")
                                        
        Call mo_Apariencia.modificarAlineacionHColumnas(Layout, 0, ssAlignCenter, _
                                "FechaEjecucion", "FechaProgramada", "NumeroSesion")
    
    End With
    'detalle del grupo
    With Layout.Bands(1)
        .Columns.Add "SiEjecutaAccion", "SI"
        .Columns.Add "NoEjecutaAccion", "NO"
        
        .Columns("SiEjecutaAccion").DataType = ssDataTypeBoolean
        .Columns("SiEjecutaAccion").Style = ssStyleCheckBox
        
        .Columns("NoEjecutaAccion").DataType = ssDataTypeBoolean
        .Columns("NoEjecutaAccion").Style = ssStyleCheckBox
        
        'establecer etiqueta de columnas Y formato
        
        .Columns("SiEjecutaAccion").Width = 1000
        .Columns("NoEjecutaAccion").Width = 1000
                
        .Columns("ItemDesarrollo").Header.Caption = "Descripción de Item a Evaluar"
        '.Columns("ItemDesarrollo").Width = grdPlanDesarrollo.Width - 500 - .Columns("SiEjecutaAccion").Width - _
         '                                           .Columns("NoEjecutaAccion").Width
        .Columns("ItemDesarrollo").ColSpan = 3 'Layout.Bands(0).Columns.Count - 1 - 9  '(filas ocultas del detalle del grupo)
        
        Call mo_Apariencia.modificarAlineacionHColumnas(Layout, 1, ssAlignCenter, "SiEjecutaAccion", "NoEjecutaAccion")
        'ocultar columnas
        Call mo_Apariencia.ocultarColumnas(Layout, 1, "IdPlanDesarrolloPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "IdItemDesarrollo", "OrdenItem", "EjecutaAccion", _
                                        "EdadAnio", "EdadMes", "EdadDia", "FechaEjecucion")
        
        '.Columns("EsEjecutada").Activation = ssActivationAllowEdit
        'desactivar edicion de columnas
        Call mo_Apariencia.modificarActivationColumnas(Layout, 1, ssActivationActivateNoEdit, "IdPlanDesarrolloPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "IdItemDesarrollo", "ItemDesarrollo", "OrdenItem", _
                                        "EjecutaAccion", "EdadAnio", "EdadMes", "EdadDia", _
                                        "FechaEjecucion", "SiEjecutaAccion", "NoEjecutaAccion")
        
    End With
End Sub

Private Sub grdPlanDesarrollo_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
    If Row.HasParent = False Then
        If frAtenInteDesarrollo.Tag = "1" Then
            Call activarEdicionPlanIntegral(Row)
        Else
            Row.Cells("FechaProgramada").Activation = ssActivationActivateNoEdit
        End If
    Else
        noEjecutarAccion = True
        Call SeleccionarRespuestaAccionDesarrollo(Row)
        noEjecutarAccion = False
    End If
    Call formatoFilaPlanIntegral(Row)
End Sub

Private Sub grdPlanDesarrolloPendientes_BeforeCellActivate(ByVal Cell As UltraGrid.SSCell, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Set ssCellActivate = Cell
    'mgaray201411b
    If Cell.Column.Key = "SiEjecutaAccion" Or Cell.Column.Key = "NoEjecutaAccion" Then
        mb_EstaMarcadoEjecucionPsicomotor = Cell.Value
    End If
End Sub

Private Sub grdPlanDesarrolloPendientes_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not (ssCellActivate Is Nothing) Then
        If ssCellActivate.Column.Key = "SiEjecutaAccion" Or ssCellActivate.Column.Key = "NoEjecutaAccion" Then
            EventsSeleccinarEjecutaAccion ssCellActivate, ssCellActivate.Row.Cells(ssCellActivate.Column.Key).Value
        End If
    End If
End Sub

Private Sub grdPlanDesarrolloPendientes_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
'    If Not IsNull(NewValue) And Cell.Column.Key = "FechaEjecucion" Then
'        If NewValue < ml_FechaAtencion Or NewValue > md_fechaActual Then
'            MsgBox "Fecha no puede ser menor que la fecha de atención ni mayor que la fecha actual", vbInformation, "Advertencia"
'            NewValue = Cell.Value
'        End If
'    End If
End Sub

Private Sub grdPlanDesarrolloPendientes_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdPlanDesarrolloPendientes_CellChange(ByVal Cell As UltraGrid.SSCell)
    If noEjecutarAccion = True Then: Exit Sub
    If Cell.Column.Key = "SiEjecutaAccion" Or Cell.Column.Key = "NoEjecutaAccion" Then
        'mgaray201411b
        mb_EstaMarcadoEjecucionPsicomotor = Not mb_EstaMarcadoEjecucionPsicomotor
        If mb_EstaMarcadoEjecucionPsicomotor = True Then
'        If Cell.Value = True Then
            noEjecutarAccion = True
            If Cell.Column.Key = "SiEjecutaAccion" Then
                Cell.Row.Cells("NoEjecutaAccion").Value = False
            Else
                Cell.Row.Cells("SiEjecutaAccion").Value = False
            End If
            noEjecutarAccion = False
        End If
    End If
End Sub

Private Sub grdPlanDesarrolloPendientes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPlanDesarrolloPendientes.ViewStyleBand = ssViewStyleBandVertical
    'evitar que los cambios en las celdas editables se hagan directamente en la base de datos
    grdPlanDesarrolloPendientes.UpdateMode = ssUpdateOnUpdate
    
    'detalle del grupo
    With Layout.Bands(0)
        .Columns.Add "SiEjecutaAccion", "SI"
        .Columns.Add "NoEjecutaAccion", "NO"
        
        .Columns("SiEjecutaAccion").DataType = ssDataTypeBoolean
        .Columns("SiEjecutaAccion").Style = ssStyleCheckBox
        
        .Columns("NoEjecutaAccion").DataType = ssDataTypeBoolean
        .Columns("NoEjecutaAccion").Style = ssStyleCheckBox
        
        'establecer etiqueta de columnas Y formato
        
        .Columns("SiEjecutaAccion").Width = 1000
        .Columns("NoEjecutaAccion").Width = 1000
                
        .Columns("ItemDesarrollo").Header.Caption = "Descripción de Item a Evaluar"
        .Columns("ItemDesarrollo").Width = grdPlanDesarrolloPendientes.Width - 500 - .Columns("SiEjecutaAccion").Width - _
                                                    .Columns("NoEjecutaAccion").Width
        
        Call mo_Apariencia.modificarAlineacionHColumnas(Layout, 0, ssAlignCenter, "SiEjecutaAccion", "NoEjecutaAccion")
        'ocultar columnas
        Call mo_Apariencia.ocultarColumnas(Layout, 0, "IdPlanDesarrolloPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "IdItemDesarrollo", "OrdenItem", "EjecutaAccion", _
                                        "EdadAnio", "EdadMes", "EdadDia", "FechaEjecucion")
        
        '.Columns("EsEjecutada").Activation = ssActivationAllowEdit
        'desactivar edicion de columnas
        Call mo_Apariencia.modificarActivationColumnas(Layout, 0, ssActivationActivateNoEdit, "IdPlanDesarrolloPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "IdItemDesarrollo", "ItemDesarrollo", "OrdenItem", "EjecutaAccion", "EdadAnio", "EdadMes", "EdadDia", "FechaEjecucion")
    End With
End Sub

Private Sub grdPlanDesarrolloPendientes_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
    If noEjecutarAccion = True Then Exit Sub
    noEjecutarAccion = True
    Call SeleccionarRespuestaAccionDesarrollo(Row)
    noEjecutarAccion = False
    Call formatoFilaPlanIntegral(Row)
    If frAtenInteDesarrollo.Tag = "1" Then
        Row.Cells("SiEjecutaAccion").Activation = ssActivationAllowEdit
        Row.Cells("NoEjecutaAccion").Activation = ssActivationAllowEdit
    Else
        Row.Cells("SiEjecutaAccion").Activation = ssActivationActivateNoEdit
        Row.Cells("NoEjecutaAccion").Activation = ssActivationActivateNoEdit
    End If
End Sub

Private Sub grdPlanDesarrolloPendientes_LostFocus()
    Dim ssReturnValue As SSReturnBoolean
    
    grdPlanDesarrolloPendientes_BeforeCellDeactivate ssReturnValue
    UserControl.txtEvalucionDesarrollo.Tag = ObtenerEvaluacion()
    UserControl.txtEvalucionDesarrollo.Text = ObtenerEvaluacionDescripcion(UserControl.txtEvalucionDesarrollo.Tag)
End Sub

Private Sub grdPlanInmunizaciones_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    Cell.Row.Cells("UserChange").Value = True
End Sub

Private Sub grdPlanInmunizaciones_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not IsNull(NewValue) And Cell.Column.Key = "FechaProgramada" Then
        If NewValue < md_fechaActual Then
            MsgBox "Fecha no puede ser menor que la fecha actual", vbInformation, "Advertencia"
            'Cell.CancelUpdate
            NewValue = Cell.Value
        End If
    End If
End Sub

Private Sub grdPlanInmunizaciones_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

'preguntas:
'- si son 2 servicios donde se consulta el paciente. ejm Cred y Medicina, el Dx de medicina debe copiarse como dx en Cred ?
'- terminar de registrar CPT y DX que no se ha encontrado en tablas (od=otra descripcion, ya está registrado aquí como 'consulta en consultorios externos')

Private Sub grdPlanInmunizaciones_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPlanInmunizaciones.ViewStyleBand = ssViewStyleBandVertical
    'evitar que los cambios en las celdas editables se hagan directamente en la base de datos
    grdPlanInmunizaciones.UpdateMode = ssUpdateOnUpdate
    grdPlanInmunizaciones.CollapseAll
    'Cabecera de grupo
    With Layout.Bands(0)
        .ColHeadersVisible = False
        'establecer etiqueta de columnas y formato
        .Columns("IdPlanAtencion").Header.Caption = "Id Plan"
        
        .Columns("Descripcion").Header.Caption = "Edad"
        .Columns("Descripcion").Width = grdPlanInmunizaciones.Width - 1200
        .Columns("Descripcion").ColSpan = Layout.Bands(1).Columns.Count - 1 - 9  '(filas ocultas del detalle del grupo)
        
        'ocultar columnas
        .Columns("IdPlanAtencion").Hidden = True
        
        'desactivar edicion de columnas
        .Columns("IdPlanAtencion").Activation = ssActivationActivateNoEdit
        .Columns("Descripcion").Activation = ssActivationActivateNoEdit
    
    End With
    'detalle del grupo
    With Layout.Bands(1)
        .Columns.Add "UserChange", "UserChange"
        .Columns("UserChange").DataType = ssDataTypeBoolean
        
        'establecer etiqueta de columnas Y formato
        .Columns("NumeroDosis").Header.Caption = "Dosis"
        .Columns("NumeroDosis").Width = 1200
                
        .Columns("FechaProgramada").Header.Caption = "F. Programada"
        .Columns("FechaProgramada").Width = 1400
                
        .Columns("FechaEjecucion").Header.Caption = "F. Ejecución"
        .Columns("FechaEjecucion").Width = 1400
                
        'mgaray20141009
        .Columns("Codigo").Header.Caption = "Código"
        .Columns("Codigo").Width = 1000
        
        .Columns("Nombre").Header.Caption = "Tipo Inmunización"
        .Columns("Nombre").Width = Layout.Bands(0).Columns("Descripcion").Width - _
                                                    .Columns("Codigo").Width - _
                                                    .Columns("NumeroDosis").Width - _
                                                    .Columns("FechaProgramada").Width - _
                                                    .Columns("FechaEjecucion").Width
        
        Call mo_Apariencia.modificarAlineacionHColumnas(Layout, 1, ssAlignCenter, _
                        "NumeroDosis", "FechaProgramada", "FechaEjecucion")
        'ocultar columnas
        Call mo_Apariencia.ocultarColumnas(Layout, 1, "IdPlanProcedimientoPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "Descripcion", "IdProducto", "EdadAnio", _
                                        "EdadMes", "EdadDia", "IdEstablecimiento", _
                                        "Establecimiento", "Userchange")
        
        'desactivar edicion de columnas
        'mgaray20141009
        Call mo_Apariencia.modificarActivationColumnas(Layout, 1, ssActivationActivateNoEdit, "IdPlanProcedimientoPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "Descripcion", "IdProducto", "Nombre", "NumeroDosis", _
                                        "FechaEjecucion", "Codigo")
    End With
    
End Sub

Public Sub cargarListaInmunizaciones()
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente
    
    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    oDOAtenIntePlanIntePaciente.idPaciente = ml_IdPaciente
        
    Set grdPlanInmunizaciones.DataSource = oReglasAtencionIntegral.ListarPlanInmunizacionPaciente(oDOAtenIntePlanIntePaciente)
    If oReglasAtencionIntegral.MensajeError <> "" Then
        MsgBox oReglasAtencionIntegral.MensajeError, vbInformation, "Error"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdPlanInmunizaciones, sighentidades.GrillaConFilasBicolor
End Sub

Public Sub cargarListaInmunizacionesPendientes()
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente
    
    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    oDOAtenIntePlanIntePaciente.idPaciente = ml_IdPaciente
    oDOAtenIntePlanIntePaciente.idAtencion = ml_idAtencion
    
    'duplica consultas para hacer validaciones sin afecta la grida que contiene la ejecucion de las inmunizaciones
    Set mo_rsImunizacionesPendientes = oReglasAtencionIntegral.ListarPlanInmunizacionPacientePendientes(oDOAtenIntePlanIntePaciente)
    
    Set grdPlanInmunizacionesPendientes.DataSource = oReglasAtencionIntegral.ListarPlanInmunizacionPacientePendientes(oDOAtenIntePlanIntePaciente)
    If oReglasAtencionIntegral.MensajeError <> "" Then
        MsgBox oReglasAtencionIntegral.MensajeError, vbInformation, "Error"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdPlanInmunizacionesPendientes, sighentidades.GrillaConFilasBicolor
End Sub


Private Sub grdPlanInmunizaciones_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
    If Row.HasParent = True Then
        'si tiene permiso para ejecutar inmunizaciones
        If frAtenInteInmunizaciones.Tag = "1" Then
            Call activarEdicionPlanIntegral(Row)
        Else
            Row.Cells("FechaProgramada").Activation = ssActivationActivateNoEdit
        End If
        Call formatoFilaPlanIntegral(Row)
    End If
End Sub

Private Sub grdPlanInmunizacionesPendientes_BeforeCellActivate(ByVal Cell As UltraGrid.SSCell, ByVal Cancel As UltraGrid.SSReturnBoolean)
    'programada para controlar la ejecucion por que el evento CellChange de la celda no
    'es confiable(al menos para un checkButton) el valor se queda pegado en algunas
    'ocaciones(Reproducir problema ponga un Debug.Print cell.value en el evento y haga click varias veces en el control)
    Set ssCellActivate = Cell
End Sub

Private Sub grdPlanInmunizacionesPendientes_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not (ssCellActivate Is Nothing) Then
        If ssCellActivate.Column.Key = "EsEjecutada" Then
            CambiarIndicadorDeEjecutarInmunizacion ssCellActivate, ssCellActivate.Row.Cells("EsEjecutada").Value
            'Debug.Print ssCellActivate.Row.Cells("EsEjecutada").Value & "antes de desactivar a celda"
        End If
    End If
End Sub

Private Sub grdPlanInmunizacionesPendientes_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not IsNull(NewValue) And Cell.Column.Key = "FechaEjecucion" Then
        'mgaray201411b
        Dim dFechaProgramada As Date
        dFechaProgramada = Cell.Row.Cells("FechaProgramada").Value
        If NewValue < dFechaProgramada Or NewValue > md_fechaActual Then
            MsgBox "Fecha no puede ser menor que la fecha para la que fue programada la inmunización ni mayor que la fecha actual", vbInformation, "Advertencia"
            NewValue = Cell.Value
        End If
        
'        If NewValue < ml_FechaAtencion Or NewValue > md_fechaActual Then
'            MsgBox "Fecha no puede ser menor que la fecha de atención ni mayor que la fecha actual", vbInformation, "Advertencia"
'            NewValue = Cell.Value
'        End If
    End If
End Sub

Private Sub grdPlanInmunizacionesPendientes_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdPlanInmunizacionesPendientes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPlanInmunizacionesPendientes.ViewStyleBand = ssViewStyleBandVertical
    'evitar que los cambios en las celdas editables se hagan directamente en la base de datos
    grdPlanInmunizacionesPendientes.UpdateMode = ssUpdateOnUpdate
    
    'detalle del grupo
    With Layout.Bands(0)
        .Columns.Add "EsEjecutada", "Aplicar"
        .Columns("EsEjecutada").DataType = ssDataTypeBoolean
        .Columns("EsEjecutada").Style = ssStyleCheckBox
        '.Columns("EsEjecutada").ButtonDisplayStyle = ssButtonDisplayStyleOnRowActivate
        
        'establecer etiqueta de columnas Y formato
        .Columns("NumeroDosis").Header.Caption = "Dosis"
        .Columns("NumeroDosis").Width = 1200
                
        .Columns("FechaProgramada").Header.Caption = "F. Programada"
        .Columns("FechaProgramada").Width = 1400
        
        .Columns("EsEjecutada").Header.Caption = "Aplicar"
        .Columns("EsEjecutada").Width = 1000
                
        .Columns("FechaEjecucion").Header.Caption = "F. Ejecución"
        .Columns("FechaEjecucion").Width = 1400
        
        .Columns("Descripcion").Width = 0
        'mgaray20141009
        .Columns("Codigo").Header.Caption = "Código"
        .Columns("Codigo").Width = 1000
                
        .Columns("Nombre").Header.Caption = "Tipo Inmunización"
        .Columns("Nombre").Width = grdPlanInmunizacionesPendientes.Width - 500 - .Columns("Descripcion").Width - _
                                                    .Columns("Codigo").Width - _
                                                    .Columns("NumeroDosis").Width - _
                                                    .Columns("FechaProgramada").Width - _
                                                    .Columns("FechaEjecucion").Width - .Columns("EsEjecutada").Width
        
        Call mo_Apariencia.modificarAlineacionHColumnas(Layout, 0, ssAlignCenter, "NumeroDosis", "FechaProgramada", "FechaEjecucion", "EsEjecutada")
        'ocultar columnas
        Call mo_Apariencia.ocultarColumnas(Layout, 0, "IdPlanProcedimientoPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "Descripcion", "IdProducto", "EdadAnio", _
                                        "EdadMes", "EdadDia", "IdEstablecimiento", _
                                        "Establecimiento", "IdAtencion", "CodigoHIS") ', "Userchange")
        
        '.Columns("EsEjecutada").Activation = ssActivationAllowEdit
        'desactivar edicion de columnas
        'mgaray20141009
        Call mo_Apariencia.modificarActivationColumnas(Layout, 0, ssActivationActivateNoEdit, "IdPlanProcedimientoPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "Descripcion", "IdProducto", "Nombre", "NumeroDosis", _
                                        "FechaProgramada", "Codigo")
    End With
End Sub

Public Sub cargarListaSuplemento()
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente

    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    oDOAtenIntePlanIntePaciente.idPaciente = ml_IdPaciente

    Set grdPlanSuplemento.DataSource = oReglasAtencionIntegral.ListarPlanSuplementoPaciente(oDOAtenIntePlanIntePaciente)
    If oReglasAtencionIntegral.MensajeError <> "" Then
        MsgBox oReglasAtencionIntegral.MensajeError, vbInformation, "Error"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdPlanSuplemento, sighentidades.GrillaConFilasBicolor
End Sub

Public Sub cargarListaSuplementoPendientes()
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente

    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    oDOAtenIntePlanIntePaciente.idPaciente = ml_IdPaciente
    oDOAtenIntePlanIntePaciente.idAtencion = ml_idAtencion

    Set grdPlanSuplementoPendientes.DataSource = oReglasAtencionIntegral.ListarPlanSuplementoPacientePendientes(oDOAtenIntePlanIntePaciente)
    If oReglasAtencionIntegral.MensajeError <> "" Then
        MsgBox oReglasAtencionIntegral.MensajeError, vbInformation, "Error"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdPlanSuplementoPendientes, sighentidades.GrillaConFilasBicolor
    depurarDatosGridaMedicamentos
End Sub


Public Sub cargarListaCrecimiento()
'    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
'    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente
'
'    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
'    oDOAtenIntePlanIntePaciente.idPaciente = ml_idPaciente
'
'    Set grdPlanDesarrollo.DataSource = oReglasAtencionIntegral.ListarPlanDesarrolloPaciente(oDOAtenIntePlanIntePaciente)
'    If oReglasAtencionIntegral.MensajeError <> "" Then
'        MsgBox oReglasAtencionIntegral.MensajeError, vbInformation, "Error"
'    End If
'    mo_Apariencia.ConfigurarFilasBiColores grdPlanDesarrollo, sighEntidades.GrillaConFilasBicolor
End Sub

Public Sub cargarListaCrecimientoPendientes()
'    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
'    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente
'
'    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
'    oDOAtenIntePlanIntePaciente.idPaciente = ml_idPaciente
'    oDOAtenIntePlanIntePaciente.idAtencion = ml_idAtencion
'
'
'    Set grdPlanDesarrolloPendientes.DataSource = oReglasAtencionIntegral.ListarPlanDesarrolloPacientePendientes(oDOAtenIntePlanIntePaciente)
'
'    Call LimpiarDatosAControlesDesarrollo
'
'    If oReglasAtencionIntegral.MensajeError <> "" Then
'        MsgBox oReglasAtencionIntegral.MensajeError, vbInformation, "Error"
'    Else
'        Call AsignarDatosAControlesDesarrollo(oDOAtenIntePlanIntePaciente)
'    End If
'    mo_Apariencia.ConfigurarFilasBiColores grdPlanDesarrolloPendientes, sighEntidades.GrillaConFilasBicolor
End Sub

Public Sub cargarListaTamizajes()
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente

    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    oDOAtenIntePlanIntePaciente.idPaciente = ml_IdPaciente

    Set grdPlanTamizajes.DataSource = oReglasAtencionIntegral.ListarPlanTamizajePaciente(oDOAtenIntePlanIntePaciente)
    If oReglasAtencionIntegral.MensajeError <> "" Then
        MsgBox oReglasAtencionIntegral.MensajeError, vbInformation, "Error"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdPlanTamizajes, sighentidades.GrillaConFilasBicolor
End Sub

Public Sub cargarListaTamizajesPendientes()
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente

    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    oDOAtenIntePlanIntePaciente.idPaciente = ml_IdPaciente
    oDOAtenIntePlanIntePaciente.idAtencion = ml_idAtencion

    Set grdPlanTamizajesPendientes.DataSource = oReglasAtencionIntegral.ListarPlanTamizajePacientePendientes(oDOAtenIntePlanIntePaciente)
    If oReglasAtencionIntegral.MensajeError <> "" Then
        MsgBox oReglasAtencionIntegral.MensajeError, vbInformation, "Error"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdPlanTamizajesPendientes, sighentidades.GrillaConFilasBicolor
    cargarListaAComboProcedimientos
    depurarDatosGridaProcedimientos
    
End Sub

'mgaray
Public Function SeleccionarRespuestaAccionDesarrollo(ByVal Row As UltraGrid.SSRow)
    If IsNull(Row.Cells("EjecutaAccion").Value) Then
        Row.Cells("SiEjecutaAccion").Value = False
        Row.Cells("NoEjecutaAccion").Value = False
        
    Else
        If Row.Cells("EjecutaAccion").Value = True Then
            Row.Cells("SiEjecutaAccion").Value = True
            Row.Cells("NoEjecutaAccion").Value = False
        Else
            Row.Cells("SiEjecutaAccion").Value = False
        Row.Cells("NoEjecutaAccion").Value = True
        End If
    End If
End Function

Private Function EventsSeleccinarEjecutaAccion(Cell As SSCell, EjecutaAccion As Boolean)
    If EjecutaAccion = True Then
        noEjecutarAccion = True
        If Cell.Column.Key = "SiEjecutaAccion" Then
            Cell.Row.Cells("NoEjecutaAccion").Value = False
        Else
            Cell.Row.Cells("SiEjecutaAccion").Value = False
        End If
        noEjecutarAccion = False
    End If
    
End Function


Public Function activarEdicionPlanIntegral(ByVal Row As UltraGrid.SSRow)
    If Row.Cells("FechaProgramada").Value < md_fechaActual Or Not IsNull(Row.Cells("FechaEjecucion").Value) Then
        Row.Cells("FechaProgramada").Activation = ssActivationActivateNoEdit
    Else
        Row.Cells("FechaProgramada").Activation = ssActivationAllowEdit
    End If
End Function

Public Function activarEdicionInmunizacionesPendiente(ByVal Row As UltraGrid.SSRow)
    If Row.Cells("EsEjecutada").Value = True Then
        Row.Cells("FechaEjecucion").Activation = ssActivationAllowEdit
    Else
        Row.Cells("FechaEjecucion").Activation = ssActivationActivateNoEdit
    End If
End Function


Public Function formatoFilaPlanIntegral(ByVal Row As UltraGrid.SSRow)
    If IsNull(Row.Cells("FechaEjecucion").Value) Then
        
        If (Row.Cells("EdadAnio").Value = oEdad.EdadAnio _
                        And Row.Cells("EdadMes").Value < oEdad.EdadMes) Or Row.Cells("EdadAnio").Value < oEdad.EdadAnio Then
            Row.CellAppearance.ForeColor = vbRed
        End If
    Else
        Row.CellAppearance.ForeColor = vbBlue
    End If
End Function

Private Function calcularEdadPaciente() As Edad
    If md_fechaNacimiento <> 0 And ml_FechaAtencion <> 0 Then
        oEdad = calcularEdadDisgregada(md_fechaNacimiento, ml_FechaAtencion)
        calcularEdadPaciente = oEdad
    End If
End Function

Private Function getFechaActual() As Date
'    If md_fechaActual = 0 Then
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        md_fechaActual = lcBuscaParametro.RetornaFechaServidorSQL
'    End If
    getFechaActual = md_fechaActual
End Function


Private Sub grdPlanInmunizacionesPendientes_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
    If noEjecutarAccion = True Then Exit Sub
    Call activarEdicionInmunizacionesPendiente(Row)
    Call formatoFilaPlanIntegral(Row)
    noEjecutarAccion = True
    Row.Cells("FechaEjecucion").Activation = ssActivationAllowEdit
    'Row.Cells("EsEjecutada").Activation = ssActivationAllowEdit
    'si tiene permiso para ejecutar inmunizaciones
    If frAtenInteInmunizaciones.Tag = "1" Then
        Row.Cells("EsEjecutada").Activation = ssActivationAllowEdit
        If IsNull(Row.Cells("FechaEjecucion").Value) Then
            Row.Cells("EsEjecutada").Value = False
            Row.Cells("FechaEjecucion").Activation = ssActivationActivateNoEdit
        Else
            Row.Cells("EsEjecutada").Value = True
        End If
    Else
        Row.Cells("EsEjecutada").Activation = ssActivationActivateNoEdit
    End If
    noEjecutarAccion = False
End Sub


Public Function cargarDatosAtencionIntegral() As Boolean
    Call getFechaActual
    Call cargarListaInmunizaciones
    Call cargarListaInmunizacionesPendientes
    Call cargarListaTamizajes
    Call cargarListaTamizajesPendientes
    Call cargarListaDesarrollo
    Call cargarListaDesarrolloPendientes
    Call cargarListaCrecimiento
    Call cargarListaCrecimientoPendientes
    Call cargarListaSuplemento
    Call cargarListaSuplementoPendientes
    
    cargarDatosAtencionIntegral = True
End Function


Public Sub CambiarIndicadorDeEjecutarInmunizacion(Cell As SSCell, EsEjecutar As Boolean)
    If noEjecutarAccion = True Then: Exit Sub
    If Cell.Column.Key = "EsEjecutada" Then
        If EsEjecutar = True Then
            Cell.Row.Cells("FechaEjecucion").Activation = ssActivationAllowEdit
            If IsNull(Cell.Row.Cells("FechaEjecucion").Value) Then
                noEjecutarAccion = True
                Cell.Row.Cells("FechaEjecucion").Value = ml_FechaAtencion
                Cell.Row.Cells("FechaEjecucion").Activation = ssActivationAllowEdit
                noEjecutarAccion = False
            End If

        Else
            Cell.Row.Cells("FechaEjecucion").Activation = ssActivationActivateNoEdit
        End If
    End If
End Sub


Public Sub cargarListaDesarrollo()
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente
    
    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    oDOAtenIntePlanIntePaciente.idPaciente = ml_IdPaciente
        
    Set grdPlanDesarrollo.DataSource = oReglasAtencionIntegral.ListarPlanDesarrolloPaciente(oDOAtenIntePlanIntePaciente)
    If oReglasAtencionIntegral.MensajeError <> "" Then
        MsgBox oReglasAtencionIntegral.MensajeError, vbInformation, "Error"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdPlanDesarrollo, sighentidades.GrillaConFilasBicolor
    Err = 0
End Sub

Public Sub cargarListaDesarrolloPendientes()
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente
    
    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    oDOAtenIntePlanIntePaciente.idPaciente = ml_IdPaciente
    oDOAtenIntePlanIntePaciente.idAtencion = ml_idAtencion
    
        
    Set grdPlanDesarrolloPendientes.DataSource = oReglasAtencionIntegral.ListarPlanDesarrolloPacientePendientes(oDOAtenIntePlanIntePaciente)
    
    Call LimpiarDatosAControlesDesarrollo
    
    If oReglasAtencionIntegral.MensajeError <> "" Then
        MsgBox oReglasAtencionIntegral.MensajeError, vbInformation, "Error"
    Else
        Call AsignarDatosAControlesDesarrollo(oDOAtenIntePlanIntePaciente)
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdPlanDesarrolloPendientes, sighentidades.GrillaConFilasBicolor
End Sub

Private Function LimpiarDatosAControlesDesarrollo() As Boolean
    Dim oFechaHOra As New FechaHora
    frAtencionDesarrollo.Caption = ""
    frAtencionDesarrollo.Tag = ""
    txtFechaProgramadaDesarrollo.Text = ""
    txtIdAtencionDesarrollo.Text = ""
    mskFechaEjecucionDes.Text = oFechaHOra.FECHA_VACIA_DMY
    txtEvalucionDesarrollo.Text = ""
    txtEvalucionDesarrollo.Tag = ""
    
    mo_Formulario.HabilitarDeshabilitar mskFechaEjecucionDes, False
End Function

Private Function AsignarDatosAControlesDesarrollo(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As Boolean
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oFechaHOra As New FechaHora
'    Dim oRsDesarrolloPendiente As ADODB.Recordset
    
    Set mo_RsDesarrolloPendiente = oReglasAtencionIntegral.GetItemDesarrolloPacientePendiente(oDOAtenIntePlanIntePaciente)
    If mo_RsDesarrolloPendiente.RecordCount > 0 Then
        mo_Formulario.HabilitarDeshabilitar mskFechaEjecucionDes, True
        mo_RsDesarrolloPendiente.MoveFirst
        frAtencionDesarrollo.Caption = IIf(IsNull(mo_RsDesarrolloPendiente!NumeroSesion), "", "Sesión " & mo_RsDesarrolloPendiente!NumeroSesion)
        frAtencionDesarrollo.Tag = IIf(IsNull(mo_RsDesarrolloPendiente!NumeroSesion), "", mo_RsDesarrolloPendiente!NumeroSesion)
        txtFechaProgramadaDesarrollo.Text = IIf(IsNull(mo_RsDesarrolloPendiente!FechaProgramada), oFechaHOra.FECHA_VACIA_DMY, mo_RsDesarrolloPendiente!FechaProgramada)
        txtIdAtencionDesarrollo.Text = IIf(IsNull(mo_RsDesarrolloPendiente!idAtencion), "", mo_RsDesarrolloPendiente!idAtencion)
        mskFechaEjecucionDes.Text = IIf(IsNull(mo_RsDesarrolloPendiente!FechaEjecucion), md_fechaActual, mo_RsDesarrolloPendiente!FechaEjecucion)
        txtEvalucionDesarrollo.Text = IIf(IsNull(mo_RsDesarrolloPendiente!EvaluacionDesc), "", mo_RsDesarrolloPendiente!EvaluacionDesc)
        txtEvalucionDesarrollo.Tag = IIf(IsNull(mo_RsDesarrolloPendiente!evaluacion), "", mo_RsDesarrolloPendiente!evaluacion)
    End If
End Function

Private Function ObtenerEvaluacionDescripcion(idEvaluacion As Integer) As String
    Select Case idEvaluacion
        Case 1:
            ObtenerEvaluacionDescripcion = "NORMAL"
        Case 2:
            ObtenerEvaluacionDescripcion = "DEFICIT"
        Case Else:
            ObtenerEvaluacionDescripcion = ""
    End Select
End Function

Public Function ObtenerEvaluacion() As Integer
    Dim oRs As ADODB.Recordset
    Dim totalItems As Integer
    Dim itemNoEjecutados As Integer, itemEjecutados As Integer
    Dim evaluacion As Integer
    
    evaluacion = 0
    itemEjecutados = 0
    itemNoEjecutados = 0
    
    If validarEvaluacionDesarrollo() = True Then
        Set oRs = grdPlanDesarrolloPendientes.DataSource
        If Not (oRs Is Nothing) Then
            totalItems = oRs.RecordCount
            If totalItems > 0 Then
                Dim siguienteFila As Boolean
                Dim oSRow As SSRow
                
                'leer las filas debido a que se necesita acceder a dos columnas agregadas que no estan en el recorset
                Set oSRow = grdPlanDesarrolloPendientes.GetRow(ssChildRowFirst)
                siguienteFila = True
                If Not (oSRow Is Nothing) Then
                    While siguienteFila = True
                        If oSRow.Cells("SiEjecutaAccion").Value = True Or oSRow.Cells("NoEjecutaAccion").Value = True Then
                            If oSRow.Cells("SiEjecutaAccion").Value = True Then
                                itemEjecutados = itemEjecutados + 1
                            Else
                                itemNoEjecutados = itemNoEjecutados + 1
                            End If
                        End If
                        siguienteFila = oSRow.HasNextSibling
                        If siguienteFila = True Then
                            Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
                        End If
                    Wend
                End If
            
'                oRs.MoveFirst
'                While oRs.EOF = True
'                    If oRs!SiEjecutaAccion = True Or oRs!NoEjecutaAccion = True Then
'                        If oRs!SiEjecutaAccion = True Then
'                            itemEjecutados = itemEjecutados + 1
'                        Else
'                            itemNoEjecutados = itemNoEjecutados + 1
'                        End If
'                    End If
'                    oRs.MoveNext
'                Wend
            End If
        End If
        If itemNoEjecutados = 0 Then
            evaluacion = 1
        Else
            evaluacion = 2
        End If
    End If
    ObtenerEvaluacion = evaluacion
End Function

Public Function validarEvaluacionDesarrollo() As Boolean
    validarEvaluacionDesarrollo = False
    Dim oRs As ADODB.Recordset
    Dim totalItems As Integer
    Dim itemNoEjecutados As Integer, itemEjecutados As Integer
    Err = 0
    itemEjecutados = 0
    itemNoEjecutados = 0
    totalItems = 0
    
    Set oRs = grdPlanDesarrolloPendientes.DataSource
    
    If Not (oRs Is Nothing) Then
        totalItems = oRs.RecordCount
        
        If totalItems > 0 Then
            Dim siguienteFila As Boolean
            Dim oSRow As SSRow
            
            'leer las filas debido a que se necesita acceder a dos columnas agregadas que no estan en el recorset
            Set oSRow = grdPlanDesarrolloPendientes.GetRow(ssChildRowFirst)
            siguienteFila = True
            If Not (oSRow Is Nothing) Then
                While siguienteFila = True
                    If oSRow.Cells("SiEjecutaAccion").Value = True Or oSRow.Cells("NoEjecutaAccion").Value = True Then
                        If oSRow.Cells("SiEjecutaAccion").Value = True Then
                            itemEjecutados = itemEjecutados + 1
                        Else
                            itemNoEjecutados = itemNoEjecutados + 1
                        End If
                    End If
                    siguienteFila = oSRow.HasNextSibling
                    If siguienteFila = True Then
                        Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
                    End If
                Wend
            End If
        End If
    End If
    If totalItems = itemEjecutados + itemNoEjecutados Then
        validarEvaluacionDesarrollo = True
    End If
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation, "Advertencia"
    End If
End Function

Private Sub grdPlanSuplemento_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not IsNull(NewValue) And Cell.Column.Key = "FechaProgramada" Then
        If NewValue < md_fechaActual Then
            MsgBox "Fecha no puede ser menor que la fecha actual", vbInformation, "Advertencia"
            'Cell.CancelUpdate
            NewValue = Cell.Value
        End If
    End If
End Sub

Private Sub grdPlanSuplemento_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdPlanSuplemento_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPlanSuplemento.ViewStyleBand = ssViewStyleBandVertical
    'evitar que los cambios en las celdas editables se hagan directamente en la base de datos
    grdPlanSuplemento.UpdateMode = ssUpdateOnUpdate
    grdPlanSuplemento.CollapseAll
    'Cabecera de grupo
    With Layout.Bands(0)
        .ColHeadersVisible = False
        'establecer etiqueta de columnas y formato
        .Columns("IdPlanAtencion").Header.Caption = "Id Plan"
        
        .Columns("Descripcion").Header.Caption = "Edad"
        .Columns("Descripcion").Width = grdPlanSuplemento.Width - 1200
        .Columns("Descripcion").ColSpan = Layout.Bands(1).Columns.Count - 1 - 9  '(filas ocultas del detalle del grupo)
        
        'ocultar columnas
        .Columns("IdPlanAtencion").Hidden = True
        
        'desactivar edicion de columnas
        .Columns("IdPlanAtencion").Activation = ssActivationActivateNoEdit
        .Columns("Descripcion").Activation = ssActivationActivateNoEdit
    
    End With
    'detalle del grupo
    With Layout.Bands(1)
               
        'establecer etiqueta de columnas Y formato
        .Columns("NumeroDosis").Header.Caption = "Nro. Dosis"
        .Columns("NumeroDosis").Width = 1200

        .Columns("FechaProgramada").Header.Caption = "F. Programada"
        .Columns("FechaProgramada").Width = 1400

        .Columns("FechaEjecucion").Header.Caption = "F. Ejecución"
        .Columns("FechaEjecucion").Width = 1400

        .Columns("Nombre").Header.Caption = "Suplemento"
        .Columns("Nombre").Width = Layout.Bands(0).Columns("Descripcion").Width - _
                                                    .Columns("NumeroDosis").Width - _
                                                    .Columns("FechaProgramada").Width - _
                                                    .Columns("FechaEjecucion").Width

        Call mo_Apariencia.modificarAlineacionHColumnas(Layout, 1, ssAlignCenter, _
                        "NumeroDosis", "FechaProgramada", "FechaEjecucion")
        'ocultar columnas
        Call mo_Apariencia.ocultarColumnas(Layout, 1, "IdPlanSuplementoPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "Descripcion", "IdProducto", "EdadAnio", _
                                        "EdadMes", "EdadDia", "IdEstablecimiento", _
                                        "Establecimiento")

        'desactivar edicion de columnas
        Call mo_Apariencia.modificarActivationColumnas(Layout, 1, ssActivationActivateNoEdit, "IdPlanSuplementoPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "Descripcion", "IdProducto", "Nombre", "NumeroDosis", _
                                        "FechaEjecucion")
    End With
End Sub

Private Sub grdPlanSuplemento_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
    If Row.HasParent = True Then
        'si tiene permiso para ejecutar inmunizaciones
        If frAtenInteSuplemento.Tag = "1" Then
            Call activarEdicionPlanIntegral(Row)
        Else
            Row.Cells("FechaProgramada").Activation = ssActivationActivateNoEdit
        End If
        Call formatoFilaPlanIntegral(Row)
    End If
End Sub

Private Sub grdPlanSuplementoPendientes_BeforeCellActivate(ByVal Cell As UltraGrid.SSCell, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Set ssCellActivate = Cell
End Sub

Private Sub grdPlanSuplementoPendientes_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not (ssCellActivate Is Nothing) Then
        If ssCellActivate.Column.Key = "EsEjecutada" Then
            CambiarIndicadorDeEjecutarInmunizacion ssCellActivate, ssCellActivate.Row.Cells("EsEjecutada").Value
        End If
    End If
End Sub

Private Sub grdPlanSuplementoPendientes_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not IsNull(NewValue) And Cell.Column.Key = "FechaEjecucion" Then
        If NewValue < ml_FechaAtencion Or NewValue > md_fechaActual Then
            MsgBox "Fecha no puede ser menor que la fecha de atención ni mayor que la fecha actual", vbInformation, "Advertencia"
            NewValue = Cell.Value
        End If
    End If
End Sub

Private Sub grdPlanSuplementoPendientes_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdPlanSuplementoPendientes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPlanSuplementoPendientes.ViewStyleBand = ssViewStyleBandVertical
    'evitar que los cambios en las celdas editables se hagan directamente en la base de datos
    grdPlanSuplementoPendientes.UpdateMode = ssUpdateOnUpdate
    
    'detalle del grupo
    With Layout.Bands(0)
        .Columns.Add "EsEjecutada", "Aplicar"
        .Columns("EsEjecutada").DataType = ssDataTypeBoolean
        .Columns("EsEjecutada").Style = ssStyleCheckBox
        '.Columns("EsEjecutada").ButtonDisplayStyle = ssButtonDisplayStyleOnRowActivate
        
        'establecer etiqueta de columnas Y formato
        .Columns("NumeroDosis").Header.Caption = "Nro. Dosis"
        .Columns("NumeroDosis").Width = 1200

        .Columns("FechaProgramada").Header.Caption = "F. Programada"
        .Columns("FechaProgramada").Width = 1400

        .Columns("EsEjecutada").Header.Caption = "Aplicar"
        .Columns("EsEjecutada").Width = 1000

        .Columns("FechaEjecucion").Header.Caption = "F. Ejecución"
        .Columns("FechaEjecucion").Width = 1400

        .Columns("Descripcion").Width = 0

        .Columns("Nombre").Header.Caption = "Tipo Inmunización"
        .Columns("Nombre").Width = grdPlanInmunizacionesPendientes.Width - 500 - .Columns("Descripcion").Width - _
                                                    .Columns("NumeroDosis").Width - _
                                                    .Columns("FechaProgramada").Width - _
                                                    .Columns("FechaEjecucion").Width - .Columns("EsEjecutada").Width

        Call mo_Apariencia.modificarAlineacionHColumnas(Layout, 0, ssAlignCenter, "NumeroDosis", "FechaProgramada", "FechaEjecucion", "EsEjecutada")
        'ocultar columnas
        Call mo_Apariencia.ocultarColumnas(Layout, 0, "IdPlanSuplementoPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "Descripcion", "IdProducto", "EdadAnio", _
                                        "EdadMes", "EdadDia", "IdEstablecimiento", _
                                        "Establecimiento", "IdAtencion")

        '.Columns("EsEjecutada").Activation = ssActivationAllowEdit
        'desactivar edicion de columnas
        Call mo_Apariencia.modificarActivationColumnas(Layout, 0, ssActivationActivateNoEdit, "IdPlanSuplementoPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "Descripcion", "IdProducto", "Nombre", "NumeroDosis", _
                                        "FechaProgramada")
    End With
End Sub

Private Sub grdPlanSuplementoPendientes_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
    If noEjecutarAccion = True Then Exit Sub
    Call activarEdicionInmunizacionesPendiente(Row)
    Call formatoFilaPlanIntegral(Row)
    noEjecutarAccion = True
    Row.Cells("FechaEjecucion").Activation = ssActivationAllowEdit
    'Row.Cells("EsEjecutada").Activation = ssActivationAllowEdit
    'si tiene permiso para ejecutar inmunizaciones
    If frAtenInteSuplemento.Tag = "1" Then
        Row.Cells("EsEjecutada").Activation = ssActivationAllowEdit
        If IsNull(Row.Cells("FechaEjecucion").Value) Then
            Row.Cells("EsEjecutada").Value = False
            Row.Cells("FechaEjecucion").Activation = ssActivationActivateNoEdit
        Else
            Row.Cells("EsEjecutada").Value = True
        End If
    Else
        Row.Cells("EsEjecutada").Activation = ssActivationActivateNoEdit
    End If
    noEjecutarAccion = False
End Sub


Private Sub grdPlanTamizajes_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not IsNull(NewValue) And Cell.Column.Key = "FechaProgramada" Then
        If NewValue < md_fechaActual Then
            MsgBox "Fecha no puede ser menor que la fecha actual", vbInformation, "Advertencia"
            'Cell.CancelUpdate
            NewValue = Cell.Value
        End If
    End If
End Sub

Private Sub grdPlanTamizajes_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

'preguntas:
'- si son 2 servicios donde se consulta el paciente. ejm Cred y Medicina, el Dx de medicina debe copiarse como dx en Cred ?
'- terminar de registrar CPT y DX que no se ha encontrado en tablas (od=otra descripcion, ya está registrado aquí como 'consulta en consultorios externos')

Private Sub grdPlanTamizajes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPlanTamizajes.ViewStyleBand = ssViewStyleBandVertical
    'evitar que los cambios en las celdas editables se hagan directamente en la base de datos
    grdPlanTamizajes.UpdateMode = ssUpdateOnUpdate
    grdPlanTamizajes.CollapseAll
    'Cabecera de grupo
    With Layout.Bands(0)
        .ColHeadersVisible = False
        'establecer etiqueta de columnas y formato
        .Columns("IdPlanAtencion").Header.Caption = "Id Plan"
        
        .Columns("Descripcion").Header.Caption = "Edad"
        .Columns("Descripcion").Width = grdPlanTamizajes.Width - 1200
        .Columns("Descripcion").ColSpan = Layout.Bands(1).Columns.Count - 1 - 9  '(filas ocultas del detalle del grupo)
        
        'ocultar columnas
        .Columns("IdPlanAtencion").Hidden = True
        
        'desactivar edicion de columnas
        .Columns("IdPlanAtencion").Activation = ssActivationActivateNoEdit
        .Columns("Descripcion").Activation = ssActivationActivateNoEdit
    
    End With
    'detalle del grupo
    With Layout.Bands(1)
        
        'establecer etiqueta de columnas Y formato
        .Columns("NumeroDosis").Header.Caption = "Dosis"
        .Columns("NumeroDosis").Width = 1200
                
        .Columns("FechaProgramada").Header.Caption = "F. Programada"
        .Columns("FechaProgramada").Width = 1400
                
        .Columns("FechaEjecucion").Header.Caption = "F. Ejecución"
        .Columns("FechaEjecucion").Width = 1400
                
        .Columns("Nombre").Header.Caption = "Tipo Inmunización"
        .Columns("Nombre").Width = Layout.Bands(0).Columns("Descripcion").Width - _
                                                    .Columns("NumeroDosis").Width - _
                                                    .Columns("FechaProgramada").Width - _
                                                    .Columns("FechaEjecucion").Width
        
        Call mo_Apariencia.modificarAlineacionHColumnas(Layout, 1, ssAlignCenter, _
                        "NumeroDosis", "FechaProgramada", "FechaEjecucion")
        'ocultar columnas
        Call mo_Apariencia.ocultarColumnas(Layout, 1, "IdPlanProcedimientoPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "Descripcion", "IdProducto", "EdadAnio", _
                                        "EdadMes", "EdadDia", "IdEstablecimiento", _
                                        "Establecimiento")
        
        'desactivar edicion de columnas
        Call mo_Apariencia.modificarActivationColumnas(Layout, 1, ssActivationActivateNoEdit, "IdPlanProcedimientoPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "Descripcion", "IdProducto", "Nombre", "NumeroDosis", _
                                        "FechaEjecucion")
    End With
    
End Sub

Private Sub grdPlanTamizajes_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
    If Row.HasParent = True Then
        'si tiene permiso para ejecutar Tamizajes
        If frAtenInteTamizajes.Tag = "1" Then
            Call activarEdicionPlanIntegral(Row)
        Else
            Row.Cells("FechaProgramada").Activation = ssActivationActivateNoEdit
        End If
        Call formatoFilaPlanIntegral(Row)
    End If
End Sub

Private Sub grdPlanTamizajesPendientes_BeforeCellActivate(ByVal Cell As UltraGrid.SSCell, ByVal Cancel As UltraGrid.SSReturnBoolean)
    'programada para controlar la ejecucion por que el evento CellChange de la celda no
    'es confiable(al menos para un checkButton) el valor se queda pegado en algunas
    'ocaciones(Reproducir problema ponga un Debug.Print cell.value en el evento y haga click varias veces en el control)
    Set ssCellActivate = Cell
End Sub

Private Sub grdPlanTamizajesPendientes_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not (ssCellActivate Is Nothing) Then
        If ssCellActivate.Column.Key = "EsEjecutada" Then
            CambiarIndicadorDeEjecutarInmunizacion ssCellActivate, ssCellActivate.Row.Cells("EsEjecutada").Value
            'Debug.Print ssCellActivate.Row.Cells("EsEjecutada").Value & "antes de desactivar a celda"
        End If
    End If
End Sub

Private Sub grdPlanTamizajesPendientes_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not IsNull(NewValue) And Cell.Column.Key = "FechaEjecucion" Then
        If NewValue < ml_FechaAtencion Or NewValue > md_fechaActual Then
            MsgBox "Fecha no puede ser menor que la fecha de atención ni mayor que la fecha actual", vbInformation, "Advertencia"
            NewValue = Cell.Value
        End If
    End If
End Sub

Private Sub grdPlanTamizajesPendientes_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdPlanTamizajesPendientes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdPlanTamizajesPendientes.ViewStyleBand = ssViewStyleBandVertical
    'evitar que los cambios en las celdas editables se hagan directamente en la base de datos
    grdPlanTamizajesPendientes.UpdateMode = ssUpdateOnUpdate
    
    'detalle del grupo
    With Layout.Bands(0)
        .Columns.Add "EsEjecutada", "Aplicar"
        .Columns("EsEjecutada").DataType = ssDataTypeBoolean
        .Columns("EsEjecutada").Style = ssStyleCheckBox
        '.Columns("EsEjecutada").ButtonDisplayStyle = ssButtonDisplayStyleOnRowActivate
        
        'establecer etiqueta de columnas Y formato
        .Columns("NumeroDosis").Header.Caption = "Dosis"
        .Columns("NumeroDosis").Width = 1200
                
        .Columns("FechaProgramada").Header.Caption = "F. Programada"
        .Columns("FechaProgramada").Width = 1400
        
        .Columns("EsEjecutada").Header.Caption = "Aplicar"
        .Columns("EsEjecutada").Width = 1000
                
        .Columns("FechaEjecucion").Header.Caption = "F. Ejecución"
        .Columns("FechaEjecucion").Width = 1400
        
        .Columns("Descripcion").Width = 0
                
        .Columns("Nombre").Header.Caption = "Tipo Inmunización"
        .Columns("Nombre").Width = grdPlanTamizajesPendientes.Width - 500 - .Columns("Descripcion").Width - _
                                                    .Columns("NumeroDosis").Width - _
                                                    .Columns("FechaProgramada").Width - _
                                                    .Columns("FechaEjecucion").Width - .Columns("EsEjecutada").Width
        
        Call mo_Apariencia.modificarAlineacionHColumnas(Layout, 0, ssAlignCenter, "NumeroDosis", "FechaProgramada", "FechaEjecucion", "EsEjecutada")
        'ocultar columnas
        Call mo_Apariencia.ocultarColumnas(Layout, 0, "IdPlanProcedimientoPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "Descripcion", "IdProducto", "EdadAnio", _
                                        "EdadMes", "EdadDia", "IdEstablecimiento", _
                                        "Establecimiento", "IdAtencion", "CodigoHIS") ', "Userchange")
        
        '.Columns("EsEjecutada").Activation = ssActivationAllowEdit
        'desactivar edicion de columnas
        Call mo_Apariencia.modificarActivationColumnas(Layout, 0, ssActivationActivateNoEdit, "IdPlanProcedimientoPaciente", _
                                        "IdPlanIntegralPaciente", "IdPlanAtencion", _
                                        "Descripcion", "IdProducto", "Nombre", "NumeroDosis", _
                                        "FechaProgramada")
    End With
End Sub

Private Sub grdPlanTamizajesPendientes_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
    If noEjecutarAccion = True Then Exit Sub
    Call activarEdicionInmunizacionesPendiente(Row)
    Call formatoFilaPlanIntegral(Row)
    noEjecutarAccion = True
    Row.Cells("FechaEjecucion").Activation = ssActivationAllowEdit
    'Row.Cells("EsEjecutada").Activation = ssActivationAllowEdit
    'si tiene permiso para ejecutar Tamizajes
    If frAtenInteTamizajes.Tag = "1" Then
        Row.Cells("EsEjecutada").Activation = ssActivationAllowEdit
        If IsNull(Row.Cells("FechaEjecucion").Value) Then
            Row.Cells("EsEjecutada").Value = False
            Row.Cells("FechaEjecucion").Activation = ssActivationActivateNoEdit
        Else
            Row.Cells("EsEjecutada").Value = True
        End If
    Else
        Row.Cells("EsEjecutada").Activation = ssActivationActivateNoEdit
    End If
    noEjecutarAccion = False
End Sub

Private Sub mskFechaEjecucionDes_GotFocus()
    mskFechaEjecucionDes.Tag = mskFechaEjecucionDes.Text
End Sub

Private Sub mskFechaEjecucionDes_LostFocus()
    If mskFechaEjecucionDes.Text <> sighentidades.FECHA_VACIA_DMY Then
            On Error Resume Next
            If Not EsFecha(mskFechaEjecucionDes.Text, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, "Datos de Desarrollo"
                mskFechaEjecucionDes.Text = mskFechaEjecucionDes.Tag
                mskFechaEjecucionDes.SetFocus
            ElseIf CDate(mskFechaEjecucionDes.Text) < ml_FechaAtencion _
                            Or CDate(mskFechaEjecucionDes.Text) > md_fechaActual Then
                MsgBox "Fecha no puede ser menor que la fecha de atención ni mayor que la fecha actual", vbInformation, "Datos de Desarrollo"
                mskFechaEjecucionDes.Text = mskFechaEjecucionDes.Tag
                mskFechaEjecucionDes.SetFocus
            End If
            
        End If
   'mo_Formulario.MarcarComoVacio txtFechaNacimiento
End Sub


Private Function DevuelveCptInmunizacionesAtencionIntegral() As Recordset
    Call LimpiaInmunizaciones
    
    Dim siguienteFila As Boolean
    Dim oSRow As SSRow
    
    'leer las filas debido a que se necesita acceder a dos columnas agregadas que no estan en el recorset
    Set oSRow = grdPlanInmunizacionesPendientes.GetRow(ssChildRowFirst)
    siguienteFila = True
    
    If Not (oSRow Is Nothing) Then
        While siguienteFila = True
            If oSRow.Cells("EsEjecutada").Value = True Then
                oRsInmunizaciones.AddNew
                oRsInmunizaciones.Fields!Id = oSRow.Cells("IdProducto").Value
                oRsInmunizaciones.Fields!procedimiento = oSRow.Cells("Nombre").Value
'                oRsInmunizaciones.Fields!CodigoHIS = oSRow.Cells("CodigoHIS").Value
                oRsInmunizaciones.Fields!idAtencion = ml_idAtencion
                oRsInmunizaciones.Update
            
            End If
            siguienteFila = oSRow.HasNextSibling
            If siguienteFila = True Then
                Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
            End If
        Wend
    End If
            
    On Error Resume Next
    
    Set DevuelveCptInmunizacionesAtencionIntegral = oRsInmunizaciones
End Function

Private Function DevuelveCptFrecuentesAtencionIntegral() As Recordset
    Dim oRsCptFrecuentesLocal As New ADODB.Recordset
    Dim i As Integer
    
    Set oRsCptFrecuentesLocal = clonarRecorset(oRsCptFrecuentes)
    
    If oRsCptFrecuentesLocal.Fields.Count = 0 Then 'Actualizado 07102014
        oRsCptFrecuentesLocal.Fields.Append "Id", adChar, 1, adFldIsNullable
    End If
    
    oRsCptFrecuentesLocal.Open
    
    If oRsCptFrecuentes.State <> 0 Then                  'debb-13/05/2016
        If oRsCptFrecuentes.RecordCount > 0 Then
            oRsCptFrecuentes.MoveFirst
            While oRsCptFrecuentes.EOF = False
                oRsCptFrecuentesLocal.AddNew
                For i = 0 To oRsCptFrecuentesLocal.Fields.Count - 1
                    oRsCptFrecuentesLocal.Fields(i).Value = oRsCptFrecuentes.Fields(i).Value
                Next i
                oRsCptFrecuentesLocal.Update
                oRsCptFrecuentes.MoveNext
            Wend
        End If
    End If                                                  'debb-13/05/2016
    
    Dim siguienteFila As Boolean
    Dim oSRow As SSRow
    
    'leer las filas debido a que se necesita acceder a dos columnas agregadas que no estan en el recorset
    Set oSRow = grdPlanTamizajesPendientes.GetRow(ssChildRowFirst)
    siguienteFila = True
    
    If Not (oSRow Is Nothing) Then
        While siguienteFila = True
            If oSRow.Cells("EsEjecutada").Value = True Then
                oRsCptFrecuentesLocal.AddNew
                oRsCptFrecuentesLocal.Fields!Id = oSRow.Cells("IdProducto").Value
                oRsCptFrecuentesLocal.Fields!procedimiento = oSRow.Cells("Nombre").Value
                oRsCptFrecuentesLocal.Fields!idAtencion = ml_idAtencion
                oRsCptFrecuentesLocal.Update
            End If
            siguienteFila = oSRow.HasNextSibling
            If siguienteFila = True Then
                Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
            End If
        Wend
    End If
            
    On Error Resume Next
    
    Set DevuelveCptFrecuentesAtencionIntegral = oRsCptFrecuentesLocal
End Function

Private Function DevuelveMedicamentosAtencionIntegral() As Recordset
    Dim oRsMedicamentosLocal As New ADODB.Recordset
    Dim i As Integer
    
    Set oRsMedicamentosLocal = clonarRecorset(oRsFarmaciaMI)
    oRsMedicamentosLocal.Open
    
    If oRsFarmaciaMI.RecordCount > 0 Then
        oRsFarmaciaMI.MoveFirst
        While oRsFarmaciaMI.EOF = False
            If oRsFarmaciaMI.Fields!seleccionar = True Then
                oRsMedicamentosLocal.AddNew
                For i = 0 To oRsMedicamentosLocal.Fields.Count - 1
                    oRsMedicamentosLocal.Fields(i).Value = oRsFarmaciaMI.Fields(i).Value
                Next i
                oRsMedicamentosLocal.Fields!idAtencion = ml_idAtencion
                oRsMedicamentosLocal.Update
            End If
            oRsFarmaciaMI.MoveNext
        Wend
    End If
    
    Dim siguienteFila As Boolean
    Dim oSRow As SSRow
    
    'leer las filas debido a que se necesita acceder a dos columnas agregadas que no estan en el recorset
    Set oSRow = grdPlanSuplementoPendientes.GetRow(ssChildRowFirst)
    siguienteFila = True
    
    If Not (oSRow Is Nothing) Then
        While siguienteFila = True
            If oSRow.Cells("EsEjecutada").Value = True Then
                oRsMedicamentosLocal.AddNew
                oRsMedicamentosLocal.Fields!seleccionar = True
                oRsMedicamentosLocal.Fields!Id = oSRow.Cells("IdProducto").Value
                oRsMedicamentosLocal.Fields!Medicamento = oSRow.Cells("Nombre").Value
                oRsMedicamentosLocal.Fields!idAtencion = ml_idAtencion
                oRsMedicamentosLocal.Update
            End If
            siguienteFila = oSRow.HasNextSibling
            If siguienteFila = True Then
                Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
            End If
        Wend
    End If
            
    On Error Resume Next
    
    Set DevuelveMedicamentosAtencionIntegral = oRsMedicamentosLocal
End Function
'DevuelveMedicamentos

Public Function getPlanInmunizaciones() As ADODB.Recordset
    Set getPlanInmunizaciones = getRecorsetPlanInmunizaciones()
End Function


Public Function getPlanDesarrollo() As ADODB.Recordset
    Set getPlanDesarrollo = getRecorsetPlanDesarrollo()
End Function

Public Function getPlanTamizaje() As ADODB.Recordset
    Set getPlanTamizaje = getRecorsetPlanTamizajes()
End Function

Public Function getPlanSuplemento() As ADODB.Recordset
    Set getPlanSuplemento = getRecorsetPlanSuplemento()
End Function

Public Function getPlanCrecimiento() As ADODB.Recordset
    Set getPlanCrecimiento = Nothing
End Function


Public Function getAtencionIntegralInmunizaciones() As ADODB.Recordset
    Set getAtencionIntegralInmunizaciones = getRecorsetInmunizaciones()
End Function


Public Function getAtencionIntegralDesarrollo() As ADODB.Recordset
    Set getAtencionIntegralDesarrollo = getRecorsetDesarrollo()
End Function

Public Function getAtencionIntegralTamizaje() As ADODB.Recordset
    Set getAtencionIntegralTamizaje = getRecorsetTamizajes()
End Function

Public Function getAtencionIntegralSuplemento() As ADODB.Recordset
    Set getAtencionIntegralSuplemento = getRecorsetSuplemento()
End Function

Public Function getAtencionIntegralCrecimiento() As ADODB.Recordset
    Set getAtencionIntegralCrecimiento = Nothing
End Function

Private Function getRecorsetInmunizaciones()
    Dim oRsInmunizaciones As New Recordset
    Dim oRsInmunizacionesGrida As ADODB.Recordset
    
    Set oRsInmunizacionesGrida = grdPlanInmunizacionesPendientes.DataSource
    
    
    If Not (oRsInmunizacionesGrida Is Nothing) Then
    
        Set oRsInmunizaciones = clonarRecorset(oRsInmunizacionesGrida)
        oRsInmunizaciones.Fields.Append "EsEjecutada", adBoolean, 1, adFldIsNullable
        oRsInmunizaciones.Open
        
        If oRsInmunizacionesGrida.RecordCount > 0 Then
            Dim i As Integer
            'leer datos cambiados de la grida
            Dim siguienteFila As Boolean
            Dim oSRow As SSRow
            
            Set oSRow = grdPlanInmunizacionesPendientes.GetRow(ssChildRowFirst)
            siguienteFila = True
            
            If Not (oSRow Is Nothing) Then
                While siguienteFila = True
                    If oSRow.DataChanged = True Then
                        agregarFilaAtencionIntegral oRsInmunizaciones, oSRow
                        oRsInmunizaciones!EsEjecutada = oSRow.Cells("EsEjecutada").Value
                        Call actualizarDatosEjecucionPlanIntegral(oRsInmunizaciones, Not oSRow.Cells("EsEjecutada").Value)
                            
                        If oSRow.Cells("EsEjecutada").Value = True Then
                            oRsInmunizaciones!CodigoHIS = getCodigoHisProcedimiento(oRsInmunizaciones!idProducto)
                        Else
                            oRsInmunizaciones!CodigoHIS = Null
                        End If
                        oRsInmunizaciones.Update
                    End If
                    siguienteFila = oSRow.HasNextSibling
                    If siguienteFila = True Then
                        Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
                    End If
                Wend
            End If
        End If
    End If
    Set getRecorsetInmunizaciones = oRsInmunizaciones
End Function

Private Function getRecorsetPlanInmunizaciones() As ADODB.Recordset
'    Set getRecorsetPlanInmunizaciones = grdPlanInmunizaciones.DataSource
'    Exit Function
    Dim oRsInmunizaciones As New Recordset
    Dim oRsInmunizacionesGrida As ADODB.Recordset
    Dim oSegundoNivel As Variant
    
    Set oRsInmunizacionesGrida = grdPlanInmunizaciones.DataSource
    
    If Not (oRsInmunizacionesGrida Is Nothing) Then
'        Set oRsDetalle = oSegundoNivel
        If oRsInmunizacionesGrida.RecordCount > 0 Then
            oSegundoNivel = oRsInmunizacionesGrida("detalleProcedimiento")
            Set oRsInmunizaciones = clonarVariantARecorset(oSegundoNivel)
        Else
            oRsInmunizaciones.Fields.Append "Id", adBigInt, 0, adFldIsNullable
        End If
        oRsInmunizaciones.Open
        
        If oRsInmunizacionesGrida.RecordCount > 0 Then
            Dim i As Integer
            'leer datos cambiados de la grida
            Dim siguienteFila As Boolean
            Dim siguienteFilaHija As Boolean
            Dim oSRow As SSRow
            Dim oSRowHija As SSRow
            
            Set oSRow = grdPlanInmunizaciones.GetRow(ssChildRowFirst)
            siguienteFila = True
            
            If Not (oSRow Is Nothing) Then
                While siguienteFila = True
                    If oSRow.HasChild = True Then
                        siguienteFilaHija = True
                        
                        Set oSRowHija = oSRow.GetChild(ssChildRowFirst)
                        While siguienteFilaHija = True
                            If oSRowHija.DataChanged = True Then
                                agregarFilaAtencionIntegral oRsInmunizaciones, oSRowHija
                                oRsInmunizaciones.Update
                            End If
                            siguienteFilaHija = oSRowHija.HasNextSibling
                            If siguienteFilaHija = True Then
                                Set oSRowHija = oSRowHija.GetSibling(ssSiblingRowNext)
                            End If
                        Wend
                    End If
                    siguienteFila = oSRow.HasNextSibling
                    If siguienteFila = True Then
                        Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
                    End If
                Wend
            End If
        End If
    End If
    Set getRecorsetPlanInmunizaciones = oRsInmunizaciones
End Function

Private Function getRecorsetTamizajes()
    Dim oRsTamizajes As New Recordset
    Dim oRsTamizajesGrida As ADODB.Recordset
    
    Set oRsTamizajesGrida = grdPlanTamizajesPendientes.DataSource
    
    
    If Not (oRsTamizajesGrida Is Nothing) Then
    
        Set oRsTamizajes = clonarRecorset(oRsTamizajesGrida)
        oRsTamizajes.Fields.Append "EsEjecutada", adBoolean, 1, adFldIsNullable
        oRsTamizajes.Open
        
        If oRsTamizajesGrida.RecordCount > 0 Then
            Dim i As Integer
            'leer datos cambiados de la grida
            Dim siguienteFila As Boolean
            Dim oSRow As SSRow
            
            Set oSRow = grdPlanTamizajesPendientes.GetRow(ssChildRowFirst)
            siguienteFila = True
            
            If Not (oSRow Is Nothing) Then
                While siguienteFila = True
                    If oSRow.DataChanged = True Then
                        agregarFilaAtencionIntegral oRsTamizajes, oSRow
                        oRsTamizajes!EsEjecutada = oSRow.Cells("EsEjecutada").Value
                        Call actualizarDatosEjecucionPlanIntegral(oRsTamizajes, Not oSRow.Cells("EsEjecutada").Value)
                            
                        If oSRow.Cells("EsEjecutada").Value = True Then
                            oRsTamizajes!CodigoHIS = getCodigoHisProcedimiento(oRsTamizajes!idProducto)
                        Else
                            oRsTamizajes!CodigoHIS = Null
                        End If
                        oRsTamizajes.Update
                    End If
                    siguienteFila = oSRow.HasNextSibling
                    If siguienteFila = True Then
                        Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
                    End If
                Wend
            End If
        End If
    End If
    Set getRecorsetTamizajes = oRsTamizajes
End Function

Private Function getRecorsetPlanTamizajes() As ADODB.Recordset
'    Set getRecorsetPlanTamizajes = grdPlanTamizajes.DataSource
'    Exit Function
    Dim oRsTamizajes As New Recordset
    Dim oRsTamizajesGrida As ADODB.Recordset
    Dim oSegundoNivel As Variant
    
    Set oRsTamizajesGrida = grdPlanTamizajes.DataSource
    
    
    If Not (oRsTamizajesGrida Is Nothing) Then
'        Set oRsDetalle = oSegundoNivel
        If oRsTamizajesGrida.RecordCount > 0 Then
            oSegundoNivel = oRsTamizajesGrida("detalleProcedimiento")
            Set oRsTamizajes = clonarVariantARecorset(oSegundoNivel)
        Else
            oRsTamizajes.Fields.Append "Id", adBigInt, 0, adFldIsNullable
        End If
        
        oRsTamizajes.Open
        
        If oRsTamizajesGrida.RecordCount > 0 Then
            Dim i As Integer
            'leer datos cambiados de la grida
            Dim siguienteFila As Boolean
            Dim siguienteFilaHija As Boolean
            Dim oSRow As SSRow
            Dim oSRowHija As SSRow
            
            Set oSRow = grdPlanTamizajes.GetRow(ssChildRowFirst)
            siguienteFila = True
            
            If Not (oSRow Is Nothing) Then
                While siguienteFila = True
                    If oSRow.HasChild = True Then
                        siguienteFilaHija = True
                        
                        Set oSRowHija = oSRow.GetChild(ssChildRowFirst)
                        While siguienteFilaHija = True
                            If oSRowHija.DataChanged = True Then
                                agregarFilaAtencionIntegral oRsTamizajes, oSRowHija
                                oRsTamizajes.Update
                            End If
                            siguienteFilaHija = oSRowHija.HasNextSibling
                            If siguienteFilaHija = True Then
                                Set oSRowHija = oSRowHija.GetSibling(ssSiblingRowNext)
                            End If
                        Wend
                    End If
                    siguienteFila = oSRow.HasNextSibling
                    If siguienteFila = True Then
                        Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
                    End If
                Wend
            End If
        End If
    End If
    Set getRecorsetPlanTamizajes = oRsTamizajes
End Function

Private Function getRecorsetSuplemento()
    Dim oRsSuplemento As New Recordset
    Dim oRsSuplementoGrida As ADODB.Recordset
    
    Set oRsSuplementoGrida = grdPlanSuplementoPendientes.DataSource
    
    
    If Not (oRsSuplementoGrida Is Nothing) Then
    
        Set oRsSuplemento = clonarRecorset(oRsSuplementoGrida)
        oRsSuplemento.Fields.Append "EsEjecutada", adBoolean, 1, adFldIsNullable
        oRsSuplemento.Open
        
        If oRsSuplementoGrida.RecordCount > 0 Then
            Dim i As Integer
            'leer datos cambiados de la grida
            Dim siguienteFila As Boolean
            Dim oSRow As SSRow
            
            Set oSRow = grdPlanSuplementoPendientes.GetRow(ssChildRowFirst)
            siguienteFila = True
            
            If Not (oSRow Is Nothing) Then
                While siguienteFila = True
                    If oSRow.DataChanged = True Then
                        agregarFilaAtencionIntegral oRsSuplemento, oSRow
                        oRsSuplemento!EsEjecutada = oSRow.Cells("EsEjecutada").Value
                        Call actualizarDatosEjecucionPlanIntegral(oRsSuplemento, Not oSRow.Cells("EsEjecutada").Value)
                        oRsSuplemento.Update
                    End If
                    siguienteFila = oSRow.HasNextSibling
                    If siguienteFila = True Then
                        Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
                    End If
                Wend
            End If
        End If
    End If
    Set getRecorsetSuplemento = oRsSuplemento
End Function

Private Function getRecorsetPlanSuplemento() As ADODB.Recordset
'    Set getRecorsetPlanSuplemento = grdPlanSuplemento.DataSource
'    Exit Function
    Dim oRsSuplemento As New Recordset
    Dim oRsSuplementoGrida As ADODB.Recordset
    Dim oSegundoNivel As Variant
    
    Set oRsSuplementoGrida = grdPlanSuplemento.DataSource
    
    If Not (oRsSuplementoGrida Is Nothing) Then
    
        If oRsSuplementoGrida.RecordCount > 0 Then
            oSegundoNivel = oRsSuplementoGrida("detalleSuplemento")
    '        Set oRsDetalle = oSegundoNivel
            Set oRsSuplemento = clonarVariantARecorset(oSegundoNivel)
        Else
            oRsSuplemento.Fields.Append "Id", adBigInt, 0, adFldIsNullable
        End If
        oRsSuplemento.Open
        
        
        If oRsSuplementoGrida.RecordCount > 0 Then
            Dim i As Integer
            'leer datos cambiados de la grida
            Dim siguienteFila As Boolean
            Dim siguienteFilaHija As Boolean
            Dim oSRow As SSRow
            Dim oSRowHija As SSRow
            
            Set oSRow = grdPlanSuplemento.GetRow(ssChildRowFirst)
            siguienteFila = True
            
            If Not (oSRow Is Nothing) Then
                While siguienteFila = True
                    If oSRow.HasChild = True Then
                        siguienteFilaHija = True
                        
                        Set oSRowHija = oSRow.GetChild(ssChildRowFirst)
                        While siguienteFilaHija = True
                            If oSRowHija.DataChanged = True Then
                                agregarFilaAtencionIntegral oRsSuplemento, oSRowHija
                                oRsSuplemento.Update
                            End If
                            siguienteFilaHija = oSRowHija.HasNextSibling
                            If siguienteFilaHija = True Then
                                Set oSRowHija = oSRowHija.GetSibling(ssSiblingRowNext)
                            End If
                        Wend
                    End If
                    siguienteFila = oSRow.HasNextSibling
                    If siguienteFila = True Then
                        Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
                    End If
                Wend
            End If
        End If
    
    End If
    Set getRecorsetPlanSuplemento = oRsSuplemento
End Function

Private Function verificaEjecucionItemDesarrollo(ByVal oRsDesarrollo As ADODB.Recordset) As Boolean
    verificaEjecucionItemDesarrollo = False
    If Not (oRsDesarrollo Is Nothing) Then
        oRsDesarrollo.MoveFirst
        While oRsDesarrollo.EOF = False
            If IsNull(oRsDesarrollo!EjecutaAccion) Then
                Exit Function
            End If
            oRsDesarrollo.MoveNext
        Wend
    End If
    verificaEjecucionItemDesarrollo = True
End Function

Private Function SetDatosEjecucionDesarrollo(cambioEjecucion As Boolean, _
        oRsDesarrollo As ADODB.Recordset) As ADODB.Recordset
    If cambioEjecucion = True Then
        If Not (mo_RsDesarrolloPendiente Is Nothing) Then
            If verificaEjecucionItemDesarrollo(oRsDesarrollo) = True Then
                Call actualizarDatosEjecucionPlanIntegral(mo_RsDesarrolloPendiente, False)
                mo_RsDesarrolloPendiente!evaluacion = ObtenerEvaluacion()
                mo_RsDesarrolloPendiente!FechaEjecucion = mskFechaEjecucionDes.Text
            Else
                Call actualizarDatosEjecucionPlanIntegral(mo_RsDesarrolloPendiente, True)
                mo_RsDesarrolloPendiente!evaluacion = Null
            End If
        End If
    End If
    Set SetDatosEjecucionDesarrollo = mo_RsDesarrolloPendiente
End Function

Private Function getRecorsetDesarrollo()
    Dim oRsDesarrollo As New Recordset
    Dim oRsDesarrolloGrida As ADODB.Recordset
    Dim cambioEjecucion As Boolean
    
    Set oRsDesarrolloGrida = grdPlanDesarrolloPendientes.DataSource
    'Datos de la session de desarrollo
    cambioEjecucion = False
    
    
    If Not (oRsDesarrolloGrida Is Nothing) Then
        'crear un recorset temporal basado en el recorset de la grida
        Set oRsDesarrollo = clonarRecorset(oRsDesarrolloGrida)
'        oRsDesarrollo.Fields.Append "SiEjecutaAccion", adBoolean, 1, adFldIsNullable
'        oRsDesarrollo.Fields.Append "NoEjecutaAccion", adBoolean, 1, adFldIsNullable
        oRsDesarrollo.Open
        
        If oRsDesarrolloGrida.RecordCount > 0 Then
            Dim i As Integer
            'leer datos cambiados de la grida
            Dim siguienteFila As Boolean
            Dim oSRow As SSRow
            
            Set oSRow = grdPlanDesarrolloPendientes.GetRow(ssChildRowFirst)
            siguienteFila = True
            
            If Not (oSRow Is Nothing) Then
                While siguienteFila = True
                    If oSRow.DataChanged = True Then
                        'trasladar los valores de la grida al recorset temporral
                        agregarFilaAtencionIntegral oRsDesarrollo, oSRow
                        
                        If oSRow.Cells("SiEjecutaAccion").Value = True Or _
                                        oSRow.Cells("NoEjecutaAccion").Value = True Then
                            If oSRow.Cells("SiEjecutaAccion").Value = True Then
                                oRsDesarrollo!EjecutaAccion = 1
                            Else
                                oRsDesarrollo!EjecutaAccion = 0
                            End If
                        Else
                            oRsDesarrollo!EjecutaAccion = Null
                        End If
                        oRsDesarrollo.Update
                        cambioEjecucion = True
                    End If
                    siguienteFila = oSRow.HasNextSibling
                    If siguienteFila = True Then
                        Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
                    End If
                Wend
            End If
        End If
    End If
    
    Call SetDatosEjecucionDesarrollo(cambioEjecucion, oRsDesarrollo)
    
    Set getRecorsetDesarrollo = oRsDesarrollo
End Function

Private Function getRecorsetPlanDesarrollo()
    Dim oRsDesarrollo As New Recordset
    Dim oRsDesarrolloGrida As ADODB.Recordset
    Dim cambioEjecucion As Boolean
    
    Set oRsDesarrolloGrida = grdPlanDesarrollo.DataSource
    'Datos de la session de desarrollo
    
    If Not (oRsDesarrolloGrida Is Nothing) Then
        'crear un recorset temporal basado en el recorset de la grida
        Set oRsDesarrollo = clonarRecorset(oRsDesarrolloGrida, "detalleDesarrollo")
        oRsDesarrollo.Open
        
        If oRsDesarrolloGrida.RecordCount > 0 Then
            Dim i As Integer
            'leer datos cambiados de la grida
            Dim siguienteFila As Boolean
            Dim oSRow As SSRow
            
            Set oSRow = grdPlanDesarrollo.GetRow(ssChildRowFirst)
            siguienteFila = True
            
            If Not (oSRow Is Nothing) Then
                While siguienteFila = True
                    If oSRow.DataChanged = True Then
                        'trasladar los valores de la grida al recorset temporral
                        agregarFilaAtencionIntegral oRsDesarrollo, oSRow
                        oRsDesarrollo.Update
'                        cambioEjecucion = True
                    End If
                    siguienteFila = oSRow.HasNextSibling
                    If siguienteFila = True Then
                        Set oSRow = oSRow.GetSibling(ssSiblingRowNext)
                    End If
                Wend
            End If
        End If
    End If
    Set getRecorsetPlanDesarrollo = oRsDesarrollo
End Function


Private Function clonarRecorset(oRsOriginal As ADODB.Recordset, Optional fieldExclude As String = "") As ADODB.Recordset
    Dim oRsLocal As New Recordset
    If Not (oRsOriginal Is Nothing) Then
        'If oRsOriginal.RecordCount > 0 Then
            Dim i As Integer
            With oRsLocal
                For i = 0 To oRsOriginal.Fields.Count - 1
                    If fieldExclude = "" Or oRsOriginal.Fields(i).Name <> fieldExclude Then
                        .Fields.Append oRsOriginal.Fields(i).Name, _
                                        oRsOriginal.Fields(i).Type, _
                                        oRsOriginal.Fields(i).DefinedSize, _
                                        adFldIsNullable
                    End If
                Next i
                '.Fields.Append "EsEjecutada", adBoolean, 1, adFldIsNullable
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
            End With
        'End If
    End If
    Set clonarRecorset = oRsLocal
End Function

Private Function agregarFilaAtencionIntegral(ByRef oRs As ADODB.Recordset, oSRow As SSRow) As Boolean
On Error GoTo miError
    Dim i As Integer
    oRs.AddNew
    For i = 0 To oRs.Fields.Count - 1
        oRs.Fields(oSRow.Cells(i).Column.Key).Value = oSRow.Cells(i).Value
    Next i
    agregarFilaAtencionIntegral = True
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation, "Niño Sano - Inmunizaciones"
    End If
End Function

Private Function actualizarDatosEjecucionPlanIntegral(ByRef oRs As ADODB.Recordset, _
        limpiarDatosEjecucion As Boolean)
        
    If limpiarDatosEjecucion = False Then
        If IsNull(oRs!FechaEjecucion) Then
            oRs!FechaEjecucion = md_fechaActual
        End If
        If IsNull(oRs!idAtencion) Then
            oRs!idAtencion = ml_idAtencion
        End If
        If IsNull(oRs!IdEstablecimiento) Then
            oRs!IdEstablecimiento = ml_IdEstablecimiento
        End If
    Else
        oRs!FechaEjecucion = Null
        oRs!idAtencion = Null
        oRs!IdEstablecimiento = Null
    End If
End Function

Private Function getCodigoHisProcedimiento(idProducto As Long) As String
    getCodigoHisProcedimiento = ""
End Function

Public Function validarReglasAtencionIntegral() As Boolean
    validarReglasAtencionIntegral = False
    If frAtenInteDesarrollo.Tag = "1" Then
        If validarEvaluacionDesarrollo() = False Then
            ms_MensajeError = "Debe de Evaluar Todos los Item de Desarrollo"
            STabPerinatal.Tab = 2
            Exit Function
        End If
    End If
    validarReglasAtencionIntegral = True
End Function

Private Function BloqueoControlesAtencionCRED()
    BloquearControlesCred False, Cred1, Cred2, Cred3, Cred4, _
                                Cred5, Cred6, Cred7, Cred8, Cred9, Cred10, Cred11, Cred12
    If oEdad.EdadAnio < 1 Or (oEdad.EdadAnio = 1 And oEdad.EdadMes = 0) Then
        Select Case oEdad.EdadMes
            Case 1:
                BloquearControlesCred True, Cred1
            Case 2:
                BloquearControlesCred True, Cred1, Cred2
            Case 3:
                BloquearControlesCred True, Cred1, Cred2, Cred3
            Case 4:
                BloquearControlesCred True, Cred1, Cred2, Cred3, Cred4
            Case 5
                BloquearControlesCred True, Cred1, Cred2, Cred3, Cred4, Cred5
            Case 6:
                BloquearControlesCred True, Cred1, Cred2, Cred3, Cred4, Cred5, Cred6
            Case 7:
                BloquearControlesCred True, Cred1, Cred2, Cred3, Cred4, Cred5, Cred6, Cred7
            Case 8:
                BloquearControlesCred True, Cred1, Cred2, Cred3, Cred4, Cred5, Cred6, Cred7, Cred8
            Case 9:
                BloquearControlesCred True, Cred1, Cred2, Cred3, Cred4, Cred5, Cred6, Cred7, Cred8, Cred9
            Case 10:
                BloquearControlesCred True, Cred1, Cred2, Cred3, Cred4, Cred5, Cred6, Cred7, Cred8, Cred9, Cred10
            Case 11:
                BloquearControlesCred True, Cred1, Cred2, Cred3, Cred4, Cred5, Cred6, Cred7, Cred8, Cred9, Cred10, Cred11
            Case Else:
                BloquearControlesCred True, Cred1, Cred2, Cred3, Cred4, Cred5, Cred6, Cred7, Cred8, Cred9, Cred10, Cred11, Cred12
        End Select
    ElseIf (oEdad.EdadAnio = 1 And oEdad.EdadMes <= 11) Or (oEdad.EdadAnio = 2 And oEdad.EdadMes <= 0) Then
        Select Case oEdad.EdadMes
            Case 2:
                BloquearControlesCred True, Cred1
            Case 4:
                BloquearControlesCred True, Cred1, Cred2
            Case 6:
                BloquearControlesCred True, Cred1, Cred2, Cred3
            Case 8:
                BloquearControlesCred True, Cred1, Cred2, Cred3, Cred4
            Case 10:
                BloquearControlesCred True, Cred1, Cred2, Cred3, Cred4, Cred5
            Case Else:
                BloquearControlesCred True, Cred1, Cred2, Cred3, Cred4, Cred5, Cred6
        End Select
    Else
        BloquearControlesCred True, Cred1, Cred2, Cred3, Cred4, Cred5, Cred6, Cred7, Cred8, Cred9, Cred10, Cred11, Cred12
    End If
End Function

Private Function BloquearControlesCred(Bloqueo As Boolean, ParamArray oControls() As Variant)
    Dim oControl As Variant
    Dim oObject  As Object
    For Each oControl In oControls
        Set oObject = oControl
        mo_Formulario.HabilitarDeshabilitar oObject, Bloqueo
    Next
End Function

Private Function clonarVariantARecorset(oRsOriginal As Variant, _
                    Optional fieldExclude As String = "") As ADODB.Recordset
    Dim oRsLocal As New Recordset
    If Not (oRsOriginal Is Nothing) Then
        Dim i As Integer
        With oRsLocal
            For i = 0 To oRsOriginal.Fields.Count - 1
                If fieldExclude = "" Or oRsOriginal.Fields(i).Name <> fieldExclude Then
                    .Fields.Append oRsOriginal.Fields(i).Name, _
                                    oRsOriginal.Fields(i).Type, _
                                    oRsOriginal.Fields(i).DefinedSize, _
                                    adFldIsNullable
                End If
            Next i
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
        End With
    End If
    Set clonarVariantARecorset = oRsLocal
End Function

Private Function cargarListaAComboProcedimientos(Optional oRsTmp As Recordset = Nothing) As Boolean
    Dim oConexion As New Connection
    Dim SeEstablecioConexionLocal As Boolean
    Dim lnIdListItem1 As Long
    'Dim oRsTmp As New Recordset
    SeEstablecioConexionLocal = False
    If oRsTmp Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        Set oRsTmp = mo_reglasComunes.PerinatalCatalogoCptSeleccionarPorIdModulo(lnIdModulo, oConexion)
        SeEstablecioConexionLocal = True
    End If
    
     cmbProcedimientosFrecuentes.Clear
     If oRsTmp.RecordCount > 0 Then
        Dim oRsTamizajes As ADODB.Recordset
        
        'Set oRsTamizajes = grdPlanTamizajesPendientes.DataSource
        Set oRsTamizajes = getPlanTamizajesPendientesParaValidar()
        lnIdListItem1 = 0
        oRsTmp.MoveFirst
        Do While Not oRsTmp.EOF
            If oRsTmp.Fields!idLista <> sghPerinatalListas.sighInmunizaciones Then
                If existeProcedimientoEnPlanDeTamizajesPendientes(oRsTmp.Fields!idProducto, oRsTamizajes) = False Then
                    cmbProcedimientosFrecuentes.ListItems.Add lnIdListItem1, lcCombo + Trim(Str(oRsTmp.Fields!idProducto)), oRsTmp.Fields!nombre
                    lnIdListItem1 = lnIdListItem1 + 1
'                Else
'                    Debug.Print "encontrol"
                End If
            End If
            oRsTmp.MoveNext
        Loop
     End If
     oRsTmp.Close
     
    If SeEstablecioConexionLocal = True Then
        oConexion.Close
    End If
    Set oRsTmp = Nothing
    Set oConexion = Nothing
End Function

Private Function depurarDatosGridaProcedimientos() As Boolean
    If Not (oRsCptFrecuentes Is Nothing) Then
        If oRsCptFrecuentes.RecordCount > 0 Then
            Dim oRsTamizajes As ADODB.Recordset
            'Set oRsTamizajes = grdPlanTamizajesPendientes.DataSource
            Set oRsTamizajes = getPlanTamizajesPendientesParaValidar()
        
            oRsCptFrecuentes.MoveFirst
            While oRsCptFrecuentes.EOF = False
                If existeProcedimientoEnPlanDeTamizajesPendientes(oRsCptFrecuentes.Fields!Id, oRsTamizajes) = True Then
                    oRsCptFrecuentes.Delete
                    oRsCptFrecuentes.Update
                End If
                oRsCptFrecuentes.MoveNext
            Wend
        End If
    End If
End Function

Private Function existeProcedimientoEnPlanDeTamizajesPendientes(lIdProducto As Long, ByVal oRsTamizajesParam As ADODB.Recordset) As Boolean
    Dim oRsTamizajes As ADODB.Recordset
    
    Set oRsTamizajes = oRsTamizajesParam
    
    existeProcedimientoEnPlanDeTamizajesPendientes = False
    If Not (oRsTamizajes Is Nothing) Then
        If oRsTamizajes.RecordCount > 0 Then
'            noEjecutarAccion = True
            oRsTamizajes.MoveFirst
            oRsTamizajes.Find "IdProducto=" & lIdProducto
'            noEjecutarAccion = False
            If Not oRsTamizajes.EOF Then
                existeProcedimientoEnPlanDeTamizajesPendientes = True
            End If
        End If
    End If
End Function


Private Function depurarDatosGridaMedicamentos() As Boolean
    If Not (oRsFarmaciaMI Is Nothing) Then
        If oRsFarmaciaMI.RecordCount > 0 Then
            Dim oRsMedicamentos As ADODB.Recordset
    
            'Set oRsMedicamentos = grdPlanSuplementoPendientes.DataSource
            Set oRsMedicamentos = getPlanMedicamentosPendientesParaValidar()
        
            oRsFarmaciaMI.MoveFirst
            While oRsFarmaciaMI.EOF = False
                If existeMedicamentoEnPlanDeSuplementoPendientes(oRsFarmaciaMI.Fields!Id, oRsMedicamentos) = True Then
                    oRsFarmaciaMI.Delete
                    oRsFarmaciaMI.Update
                End If
                oRsFarmaciaMI.MoveNext
            Wend
        End If
    End If
End Function

Private Function existeMedicamentoEnPlanDeSuplementoPendientes(lIdProducto As Long, oRsMedicamentos As ADODB.Recordset) As Boolean
    existeMedicamentoEnPlanDeSuplementoPendientes = False
    If Not (oRsMedicamentos Is Nothing) Then
        If oRsMedicamentos.RecordCount > 0 Then
            oRsMedicamentos.MoveFirst
            oRsMedicamentos.Find "IdProducto=" & lIdProducto
            If Not oRsMedicamentos.EOF Then
                existeMedicamentoEnPlanDeSuplementoPendientes = True
            End If
        End If
    End If
End Function

Private Function getPlanTamizajesPendientesParaValidar() As ADODB.Recordset
    Dim oRsTamizajes As ADODB.Recordset
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente
    
    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    oDOAtenIntePlanIntePaciente.idPaciente = ml_IdPaciente
    oDOAtenIntePlanIntePaciente.idAtencion = ml_idAtencion
    
    Set getPlanTamizajesPendientesParaValidar = oReglasAtencionIntegral.ListarPlanTamizajePacientePendientes(oDOAtenIntePlanIntePaciente)
End Function

Private Function getPlanMedicamentosPendientesParaValidar() As ADODB.Recordset
    Dim oRsTamizajes As ADODB.Recordset
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente
    
    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    oDOAtenIntePlanIntePaciente.idPaciente = ml_IdPaciente
    oDOAtenIntePlanIntePaciente.idAtencion = ml_idAtencion
    
    Set getPlanMedicamentosPendientesParaValidar = oReglasAtencionIntegral.ListarPlanSuplementoPacientePendientes(oDOAtenIntePlanIntePaciente)
End Function

Public Function EnfoqueTabProcedimientos()
    STabPerinatal.Tab = 0
End Function

'mgaray20141003
Private Function EliminarProcedimientoSeleccionado() As Boolean
On Error GoTo miError
    EliminarProcedimientoSeleccionado = False
    
    If MsgBox("¿Desea eliminar el Procedimiento seleccionado?", vbYesNo, "Eliminar Procedimiento") = vbYes Then
        If oRsCptFrecuentes.RecordCount > 0 Then
            With oRsCptFrecuentes
                If Not .EOF And Not .BOF Then
                   .Delete
                   .Update
                   If Not (.BOF = True And .EOF = True) Then
                        .MovePrevious
                        If .BOF = True Then
                            .MoveNext
                        End If
                   End If
                   EliminarProcedimientoSeleccionado = True
                Else
                    MsgBox "Seleccione Procedimiento a Eliminar", vbInformation, "Módulo  Niño Sano"
                End If
            End With
        End If
    End If
miError:
    If Err Then
        MsgBox Err.Number & " " & Err.Description, vbExclamation, "Módulo Niño Sano"
    End If
End Function

Private Function AgregarProcedimientosSeleccionadoDesdeListado(CBLI As SSCBListItem, _
                Optional cLabHIS As String = "") As Boolean
On Error GoTo miError
    Dim lIdProducto As Long
    Dim returnValue As Boolean
    Dim oDoFactCatalogoServicios As New DOCatalogoServicio
    
    returnValue = False
    
    oDoFactCatalogoServicios.idProducto = Val(Mid(CBLI.Key, 2, 100))
    oDoFactCatalogoServicios.nombre = CBLI.Text
    AgregarProcedimientosSeleccionadoDesdeListado = AgregarProcedimientosSeleccionado(oDoFactCatalogoServicios, cLabHIS)
    
miError:
    If Err Then
        MsgBox Err.Number & " " & Err.Description, vbExclamation, "Módulo Niño Sano"
    End If
End Function

Private Function AgregarProcedimientosSeleccionado(oDoFactCatalogoServicios As DOCatalogoServicio, _
                Optional cLabHIS As String = "") As Boolean
On Error GoTo miError
    Dim lIdProducto As Long
    Dim returnValue As Boolean
    returnValue = False
        
    lIdProducto = oDoFactCatalogoServicios.idProducto  'Val(Mid(CBLI.Key, 2, 100))
    
    If ValidarCptFrecuente(lIdProducto, cLabHIS) = False Then
        Exit Function
    End If
    
    With oRsCptFrecuentes
        .AddNew
        .Fields!Id = lIdProducto
        .Fields!procedimiento = oDoFactCatalogoServicios.nombre  'CBLI.Text
        .Fields!idAtencion = ml_idAtencion
        .Fields!labConfHIS = cLabHIS
        .Update
    End With
          
    returnValue = True
    AgregarProcedimientosSeleccionado = returnValue
miError:
    If Err Then
        MsgBox Err.Number & " " & Err.Description, vbExclamation, "Módulo Niño Sano"
    End If
End Function


Private Function EliminarDxCrecimientoDesarrolloSeleccionado() As Boolean
On Error GoTo miError
    EliminarDxCrecimientoDesarrolloSeleccionado = False
    
    If MsgBox("¿Desea eliminar Dx Crecimiento y Desarrollo seleccionado?", vbYesNo, "Eliminar DX") = vbYes Then
        If oRsDxDesarrollo.RecordCount > 0 Then
            With oRsDxDesarrollo
                If Not .EOF And Not .BOF Then
                   .Delete
                   .Update
                   If Not (.BOF = True And .EOF = True) Then
                        .MovePrevious
                        If .BOF = True Then
                            .MoveNext
                        End If
                   End If
                   EliminarDxCrecimientoDesarrolloSeleccionado = True
                Else
                    MsgBox "Seleccione Dx Crecimiento y Desarrollo a Eliminar", vbInformation, "Módulo Niño Sano"
                End If
            End With
        End If
    End If
miError:
    If Err Then
        MsgBox Err.Number & " " & Err.Description, vbExclamation, "Módulo Niño Sano"
    End If
End Function

Private Function AgregarDxCrecimientoDesarrolloSeleccionadoDesdeListado(CBLI As SSCBListItem, Optional cLabHIS As String = "") As Boolean
On Error GoTo miError
    Dim lIdProducto As Long
    Dim returnValue As Boolean
    returnValue = False
    
    Dim oDODiagnostico As New DODiagnostico
    oDODiagnostico.idDiagnostico = Val(Mid(CBLI.Key, 2, 100))
    oDODiagnostico.CodigoCIE2004 = CBLI.TagVariant
    oDODiagnostico.Descripcion = Trim(Mid(CBLI.Text, InStr(CBLI.Text, "=") + 1)) 'Actualizado 07102014
    AgregarDxCrecimientoDesarrolloSeleccionadoDesdeListado = AgregarDxCrecimientoDesarrolloSeleccionado(oDODiagnostico, cLabHIS)

miError:
    If Err Then
        MsgBox Err.Number & " " & Err.Description, vbExclamation, "Módulo Niño Sano"
    End If
End Function


Private Function AgregarDxCrecimientoDesarrolloSeleccionado(oDODiagnostico As DODiagnostico, Optional cLabHIS As String = "", _
                Optional lIdClasificacionDx As Long = 0, Optional lIdSubclasificacionDx As Long = 0) As Boolean
On Error GoTo miError
    Dim lIdProducto As Long
    Dim returnValue As Boolean
    returnValue = False
    
    If lIdClasificacionDx = 0 Then
        lIdClasificacionDx = sghTiposDiagnostico.sghAtencionConsultaExterna
    End If
    If IdSubclasificacionDx = 0 Then
        lIdSubclasificacionDx = sghDxDefinitivos.sighDxCeDefinitivo
    End If
    
    lIdProducto = oDODiagnostico.idDiagnostico 'Val(Mid(CBLI.Key, 2, 100))
    
    'mgaray201412a
    If cLabHIS = "" Then
        cLabHIS = obtenerLabAutomatico(oDODiagnostico)
    End If
    
    If ValidarDiagnosticosCRED(lIdProducto, cLabHIS) = False Then
        Exit Function
    End If
    
    With oRsDxDesarrollo
        .AddNew
        .Fields!Id = lIdProducto
        .Fields!CodigoCIE2004 = oDODiagnostico.CodigoCIE2004 'CBLI.TagVariant
        .Fields!DIAGNOSTICO = Trim(oDODiagnostico.Descripcion) 'Trim(Mid(CBLI.Text, InStr(CBLI.Text, "=") + 1)) 'Actualizado 07102014
        .Fields!idAtencion = ml_idAtencion
        .Fields!labConfHIS = cLabHIS
        .Fields!IdClasificacionDx = lIdClasificacionDx
        .Fields!IdSubclasificacionDx = lIdSubclasificacionDx
        .Update
    End With
          
    returnValue = True
    AgregarDxCrecimientoDesarrolloSeleccionado = returnValue
miError:
    If Err Then
        MsgBox Err.Number & " " & Err.Description, vbExclamation, "Módulo Niño Sano"
    End If
End Function

Private Function EliminarDxMorbilidadSeleccionado() As Boolean
On Error GoTo miError
    EliminarDxMorbilidadSeleccionado = False
    
    If MsgBox("¿Desea eliminar Dx Morbilidad seleccionado?", vbYesNo, "Eliminar DX") = vbYes Then
        If oRsMorbilidadFrec.RecordCount > 0 Then
            With oRsMorbilidadFrec
                If Not .EOF And Not .BOF Then
                   .Delete
                   .Update
                   If Not (.BOF = True And .EOF = True) Then
                        .MovePrevious
                        If .BOF = True Then
                            .MoveNext
                        End If
                   End If
                   EliminarDxMorbilidadSeleccionado = True
                Else
                    MsgBox "Seleccione Dx Morbilidad a Eliminar", vbInformation, "Módulo Niño Sano"
                End If
            End With
        End If
    End If
miError:
    If Err Then
        MsgBox Err.Number & " " & Err.Description, vbExclamation, "Módulo Niño Sano"
    End If
End Function
'mgaray201411a
Private Function AgregarDxMorbilidadSeleccionadoDesdeListado(CBLI As SSCBListItem, Optional cLabHIS As String = "") As Boolean
On Error GoTo miError
    Dim lIdProducto As Long
    Dim returnValue As Boolean
    returnValue = False
    
    Dim oDODiagnostico As New DODiagnostico
    oDODiagnostico.idDiagnostico = Val(Mid(CBLI.Key, 2, 100))
    oDODiagnostico.CodigoCIE2004 = CBLI.TagVariant
    oDODiagnostico.Descripcion = Trim(Mid(CBLI.Text, InStr(CBLI.Text, "=") + 1))  'Actualizado 07102014
    AgregarDxMorbilidadSeleccionadoDesdeListado = AgregarDxMorbilidadSeleccionado(oDODiagnostico, cLabHIS)

miError:
    If Err Then
        MsgBox Err.Number & " " & Err.Description, vbExclamation, "Módulo Niño Sano"
    End If
End Function

'Private Function AgregarDxMorbilidadSeleccionado(CBLI As SSCBListItem, Optional cLabHIS As String = "",
Private Function AgregarDxMorbilidadSeleccionado(oDODiagnostico As DODiagnostico, Optional cLabHIS As String = "", _
                Optional bSeEligioConChek As Boolean = True, Optional bEsDxPerinatal As Boolean = True, _
                Optional lIdClasificacionDx As Long = 0, Optional lIdSubclasificacionDx As Long = 0) As Boolean
On Error GoTo miError
    Dim lIdProducto As Long
    Dim returnValue As Boolean
    returnValue = False
    
'    lIdProducto = Val(Mid(CBLI.Key, 2, 100))
    lIdProducto = oDODiagnostico.idDiagnostico
    
    If lIdClasificacionDx = 0 Then
        lIdClasificacionDx = sghTiposDiagnostico.sghAtencionConsultaExterna
    End If
    If IdSubclasificacionDx = 0 Then
        lIdSubclasificacionDx = sghDxDefinitivos.sighDxCeDefinitivo
    End If
    'mgaray201412a
    If cLabHIS = "" Then
        cLabHIS = obtenerLabAutomatico(oDODiagnostico)
    End If
    If ValidarDiagnosticosFrecuentes(lIdProducto, cLabHIS) = False Then
        Exit Function
    End If
    
    With oRsMorbilidadFrec
        .AddNew
        .Fields!Id = lIdProducto
        .Fields!CodigoCIE2004 = oDODiagnostico.CodigoCIE2004 'CBLI.TagVariant
'        .Fields!DIAGNOSTICO = CBLI.Text
        .Fields!DIAGNOSTICO = Trim(oDODiagnostico.Descripcion) 'Trim(Mid(CBLI.Text, InStr(CBLI.Text, "=") + 1)) 'Actualizado 07102014
        .Fields!idAtencion = ml_idAtencion
        .Fields!SeEligioConChek = bSeEligioConChek 'True
        .Fields!EsDxPerinatal = bEsDxPerinatal 'True
        .Fields!labConfHIS = cLabHIS
        .Fields!IdClasificacionDx = lIdClasificacionDx
        .Fields!IdSubclasificacionDx = lIdSubclasificacionDx
        .Update
    End With
          
    returnValue = True
    AgregarDxMorbilidadSeleccionado = returnValue
miError:
    If Err Then
        MsgBox Err.Number & " " & Err.Description, vbExclamation, "Módulo Niño Sano"
    End If
End Function

Private Function getCodigoCieIgualLongitud() As String
    Dim cCodigo As String
    
End Function


Private Function getCodigoDiagnosticoIgualLongitud(cCodigo As String) As String
    cCodigo = Trim(cCodigo) & Space(7)
    getCodigoDiagnosticoIgualLongitud = Left(cCodigo, 7)
End Function


'mgaray20141003
Private Function AsignarListaDeLabsEnGridaDiagnosticos(oGrilla As SSUltraGrid, cNombreColumna As String) As Boolean
On Error GoTo miError
'    Dim oRsLabHis As ADODB.Recordset
'
'    Set oRsLabHis = mo_reglasComunes.DevuelveHIS_SITUACIOporDescripcion()
                
    With oGrilla.ValueLists.Add("ListaLab").ValueListItems
           If mo_RsLabHis.RecordCount > 0 Then
              mo_RsLabHis.MoveFirst
              Do While Not mo_RsLabHis.EOF
                 .Add Right(Trim((mo_RsLabHis.Fields!valores)), 3), Trim(mo_RsLabHis.Fields!valores) '& "(" & Trim(mo_RsLabHis.Fields!descripcio) & ")"
                 mo_RsLabHis.MoveNext
              Loop
           End If
    End With
'    oRsLabHis.Close
    oGrilla.Bands(0).Columns(cNombreColumna).ValueList = "ListaLab"
    
    AsignarListaDeLabsEnGridaDiagnosticos = True
miError:
    If Err Then
        MsgBox Err.Description & " : " & Err.Description, vbInformation, "Módulo Niño Sano"
    End If
End Function

Private Function ValidarDiagnosticosCRED(lIdDiagnostico As Long, cLabHIS As String, _
            Optional EditLab As Boolean = False, Optional ByVal oRegistroActual As Variant = 0) As Boolean
    On Error Resume Next
    ValidarDiagnosticosCRED = False
    Dim oRegistroExplorar As Variant
    Dim lo_RsDxDesarrollo As ADODB.Recordset
    
    Set lo_RsDxDesarrollo = oRsDxDesarrollo.Clone
    
    With lo_RsDxDesarrollo
        If Not (.BOF = True And .EOF = True) Then
            'mgaray20141009
            If .RecordCount > 0 Then
                .MoveFirst
                While .EOF = False
                    If Not (.CompareBookmarks(oRegistroActual, .Bookmark) = adCompareEqual) Or EditLab = False Then
                        'mgaray201412a
'                        If .Fields!ID = lIdDiagnostico And Trim(cLabHIS) = Trim(IIf(IsNull(.Fields!labConfHIS), "", .Fields!labConfHIS)) Then
                        If .Fields!Id = lIdDiagnostico Then
                            If EditLab = True Then
                                .Bookmark = oRegistroActual
                            End If
                            MsgBox "DX de Crecimiento y Desarrollo ya ha sido agregado", vbInformation, "Módulo Niño Sano"
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
    ValidarDiagnosticosCRED = True
End Function

'Private Function ValidarDiagnosticosCRED(lIdDiagnostico As Long, cLabHIS As String, _
'            Optional EditLab As Boolean = False) As Boolean
'
'    ValidarDiagnosticosCRED = False
'
'
'    If EditLab = False Then
'        With oRsDxDesarrollo
'            If Not (.BOF = True And .EOF = True) Then
'                'mgaray20141009
'                If .RecordCount > 0 Then
'                    .MoveFirst
'                    .Find "id=" & lIdDiagnostico
'                    If Not .EOF Then
'                       Exit Function
'                    End If
'                End If
'            End If
'        End With
'    End If
'    If Trim(cLabHIS) <> "" Then
'        If existeCodigoLabHis(mo_RsLabHis, cLabHIS) = False Then
'            MsgBox "Código LAB No Valido", vbInformation, "Módulo Perinatal"
'            Exit Function
'        End If
'    End If
'    ValidarDiagnosticosCRED = True
'End Function

Private Function ValidarDiagnosticosFrecuentes(lIdDiagnostico As Long, cLabHIS As String, _
            Optional EditLab As Boolean = False, Optional ByVal oRegistroActual As Variant = 0) As Boolean
   Dim lo_RsMorbilidadFrec As ADODB.Recordset
   ValidarDiagnosticosFrecuentes = False
   
    Set lo_RsMorbilidadFrec = oRsMorbilidadFrec.Clone
    
    With lo_RsMorbilidadFrec
        If Not (.BOF = True And .EOF = True) Then
            'mgaray20141009
            If .RecordCount > 0 Then
                .MoveFirst
                While .EOF = False
                    If Not (.CompareBookmarks(oRegistroActual, .Bookmark) = adCompareEqual) Or EditLab = False Then
                        'mgaray201412a
'                        If .Fields!ID = lIdDiagnostico And Trim(cLabHIS) = Trim(IIf(IsNull(.Fields!labConfHIS), "", .Fields!labConfHIS)) Then
                        If .Fields!Id = lIdDiagnostico Then
                            If EditLab = True Then
                                .Bookmark = oRegistroActual
                            End If
                            MsgBox "DX de Morbilidad ya ha sido agregado", vbInformation, "Módulo Niño Sano"
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
    ValidarDiagnosticosFrecuentes = True
End Function

'Private Function ValidarDiagnosticosFrecuentes(lIdDiagnostico As Long, cLabHIS As String, _
'            Optional EditLab As Boolean = False) As Boolean
'
'   ValidarDiagnosticosFrecuentes = False
'
'   If EditLab = False Then
'        With oRsMorbilidadFrec
'            If Not (.BOF = True And .EOF = True) Then
'                'mgaray20141009
'                If .RecordCount > 0 Then
'                    .MoveFirst
'                    .Find "id=" & lIdDiagnostico
'                    If Not .EOF Then
'                       Exit Function
'                    End If
'                End If
'            End If
'        End With
'    End If
'    If Trim(cLabHIS) <> "" Then
'        If existeCodigoLabHis(mo_RsLabHis, cLabHIS) = False Then
'            MsgBox "Código LAB No Valido", vbInformation, "Módulo Perinatal"
'            Exit Function
'        End If
'    End If
'    ValidarDiagnosticosFrecuentes = True
'End Function

'mgaray201411a
Private Function ValidarCptFrecuente(lIdProducto As Long, cLabHIS As String, _
            Optional EditLab As Boolean = False, Optional ByVal oRegistroActual As Variant = 0) As Boolean
    On Error Resume Next
    ValidarCptFrecuente = False
    Dim oRegistroExplorar As Variant
    Dim lo_RsCptFrecuentes As ADODB.Recordset
    
    Set lo_RsCptFrecuentes = oRsCptFrecuentes.Clone
    
    With lo_RsCptFrecuentes
        If Not (.BOF = True And .EOF = True) Then
            'mgaray20141009
            If .RecordCount > 0 Then
                .MoveFirst
                While .EOF = False
                    If Not (.CompareBookmarks(oRegistroActual, .Bookmark) = adCompareEqual) Or EditLab = False Then
                        'If .Fields!Id = lIdProducto And Trim(cLabHIS) = Trim(IIf(IsNull(.Fields!labConfHIS), "", .Fields!labConfHIS)) Then
                        If .Fields!Id = lIdProducto Then
                            If EditLab = True Then
                                .Bookmark = oRegistroActual
                            End If
                            'MsgBox "Procedimiento y Lab ya ha sido agregado", vbInformation, "Módulo Perinatal"
                            MsgBox "Procedimiento ya ha sido agregado", vbInformation, "Módulo Niño Sano"
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
    
    Dim oClonRsInmunizaciones As New ADODB.Recordset
    
    Set oClonRsInmunizaciones = mo_rsImunizacionesPendientes.Clone()
    
    If Not (oClonRsInmunizaciones Is Nothing) Then
        With oClonRsInmunizaciones
            If Not (.BOF = True And .EOF = True) Then
                If .RecordCount > 0 Then
                    .MoveFirst
                    .Find "IdProducto=" & lIdProducto
                    If Not .EOF Then
                        MsgBox "Procedimiento es una inmunización contenida en el plan de atención integral, proceda a ejecutar desde la ficha correspondiente a inmunizaciones", vbInformation, "Módulo Niño Sano"
                       Exit Function
                    End If
                End If
    
            End If
        End With
    End If
    ValidarCptFrecuente = True

End Function

'mgaray201411a
'Private Function existeCodigoLabHis(oRsLabHis As ADODB.Recordset, cLabHIS As String) As Boolean
'    Dim returnValue As Boolean
'
'    returnValue = False
'
''    If Trim(cLabHIS) <> "" Then
'        If oRsLabHis.RecordCount > 0 Then
'            oRsLabHis.MoveFirst
'            oRsLabHis.Find "valores='" & cLabHIS & "'"
'            If oRsLabHis.EOF = False Then
'                returnValue = True
'            End If
'        End If
''    End If
'    existeCodigoLabHis = returnValue
'End Function

Private Function getIdPuntoCargaHospitalizacion(lIdServicioPaciente As Long) As Long
    Dim oRsTmp As New ADODB.Recordset
    Set oRsTmp = mo_reglasComunes.FactPuntosCargaSeleccionarPorFiltro("idServicio=" & Trim(Str(lIdServicioPaciente)))
    If oRsTmp.RecordCount > 0 Then
       getIdPuntoCargaHospitalizacion = oRsTmp.Fields!idPuntoCarga
    Else
       getIdPuntoCargaHospitalizacion = 9999
    End If
End Function

Private Function AsignarListaDeTipoDxEnGrida(oGrilla As SSUltraGrid, cNombreColumna As String) As Boolean
On Error GoTo miError
    Dim oRsTipoDx As ADODB.Recordset

    Set oRsTipoDx = mo_reglasComunes.SubclasificacionDiagnosticosSeleccionarDxConsultaExterna
                
    With oGrilla.ValueLists.Add("ListaTipoDx").ValueListItems
           If oRsTipoDx.RecordCount > 0 Then
              oRsTipoDx.MoveFirst
              Do While Not oRsTipoDx.EOF
                 .Add Right(Trim((oRsTipoDx.Fields!IdSubclasificacionDx)), 3), Trim(oRsTipoDx.Fields!DescripcionLarga)
                 oRsTipoDx.MoveNext
              Loop
           End If
    End With
'    oRsLabHis.Close
    oGrilla.Bands(0).Columns(cNombreColumna).ValueList = "ListaTipoDx"
    
    AsignarListaDeTipoDxEnGrida = True
miError:
    If Err Then
        MsgBox Err.Description & " : " & Err.Description, vbInformation, "Módulo Niño Sano"
    End If
End Function


Public Sub CargaGraficoTallaEdad(lbActualizaDesdeInicioRs As Boolean)
On Error GoTo miError
    Dim lnFor As Integer
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oRsTalla As ADODB.Recordset
    Dim lIdTipoSexo As Long, lGrosorLinea As Integer, lGrosorLineaValor As Integer
'    Dim xValuesEdadAtencion As Variant, yValuesTalla As Variant
    Dim oRsListaTriaje As ADODB.Recordset
    
'    ml_PuntoTallaAtencion = 0
    
    lIdTipoSexo = ml_idTipoSexo
    If lIdTipoSexo = 0 Then
        lIdTipoSexo = sghSexo.Femenino
    End If
    lGrosorLinea = 1
    lGrosorLineaValor = 2
    
    
    If lbActualizaDesdeInicioRs = True Then
        Set oRsListaTriaje = oReglasAtencionIntegral.AtencionesCeListaTriaje(ml_idAtencion, mo_NroHistoriaClinica)
        
        Set oRsTalla = oReglasAtencionIntegral.DevuelveValorTallaPorSexoDesviacion(lIdTipoSexo, sghDesviacion.sghNormal)
        If oRsTalla.RecordCount > 0 Then
            ReDim xValuesEdad(oRsTalla.RecordCount - 1)
            ReDim yValuesTallaD0(oRsTalla.RecordCount - 1)
            oRsTalla.MoveFirst
            For lnFor = 0 To oRsTalla.RecordCount - 1
               xValuesEdad(lnFor) = oRsTalla.Fields!EdadMeses
               yValuesTallaD0(lnFor) = CDbl(oRsTalla.Fields!ValorTalla)
               oRsTalla.MoveNext
            Next
        Else
            ReDim xValuesEdad(0)
            ReDim yValuesTallaD0(0)
        End If
        
        Set oRsTalla = oReglasAtencionIntegral.DevuelveValorTallaPorSexoDesviacion(lIdTipoSexo, sghDesviacion.sghDesviacion1)
        If oRsTalla.RecordCount > 0 Then
            ReDim xValuesEdad(oRsTalla.RecordCount - 1)
            ReDim yValuesTallaD1(oRsTalla.RecordCount - 1)
            oRsTalla.MoveFirst
            For lnFor = 0 To oRsTalla.RecordCount - 1
               xValuesEdad(lnFor) = oRsTalla.Fields!EdadMeses
               yValuesTallaD1(lnFor) = CDbl(oRsTalla.Fields!ValorTalla)
               oRsTalla.MoveNext
            Next
        Else
            ReDim xValuesEdad(0)
            ReDim yValuesTallaD1(0)
        End If
        
        
        Set oRsTalla = oReglasAtencionIntegral.DevuelveValorTallaPorSexoDesviacion(lIdTipoSexo, sghDesviacion.sghDesviacion2)
        If oRsTalla.RecordCount > 0 Then
            ReDim xValuesEdad(oRsTalla.RecordCount - 1)
            ReDim yValuesTallaD2(oRsTalla.RecordCount - 1)
            oRsTalla.MoveFirst
            For lnFor = 0 To oRsTalla.RecordCount - 1
               xValuesEdad(lnFor) = oRsTalla.Fields!EdadMeses
               yValuesTallaD2(lnFor) = CDbl(oRsTalla.Fields!ValorTalla)
               oRsTalla.MoveNext
            Next
        Else
            ReDim xValuesEdad(0)
            ReDim yValuesTallaD2(0)
        End If
        
        Set oRsTalla = oReglasAtencionIntegral.DevuelveValorTallaPorSexoDesviacion(lIdTipoSexo, sghDesviacion.sghDesviacionMenos1)
        
        If oRsTalla.RecordCount > 0 Then
            ReDim xValuesEdad(oRsTalla.RecordCount - 1)
            ReDim yValuesTallaD_1(oRsTalla.RecordCount - 1)
            oRsTalla.MoveFirst
            For lnFor = 0 To oRsTalla.RecordCount - 1
               xValuesEdad(lnFor) = oRsTalla.Fields!EdadMeses
               yValuesTallaD_1(lnFor) = CDbl(oRsTalla.Fields!ValorTalla)
               oRsTalla.MoveNext
            Next
        Else
            ReDim xValuesEdad(0)
            ReDim yValuesTallaD_1(0)
        End If
        
        Set oRsTalla = oReglasAtencionIntegral.DevuelveValorTallaPorSexoDesviacion(lIdTipoSexo, sghDesviacion.sghDesviacionMenos2)
        
        If oRsTalla.RecordCount > 0 Then
            ReDim xValuesEdad(oRsTalla.RecordCount - 1)
            ReDim yValuesTallaD_2(oRsTalla.RecordCount - 1)
            oRsTalla.MoveFirst
            For lnFor = 0 To oRsTalla.RecordCount - 1
               xValuesEdad(lnFor) = oRsTalla.Fields!EdadMeses
               yValuesTallaD_2(lnFor) = CDbl(oRsTalla.Fields!ValorTalla)
               oRsTalla.MoveNext
            Next
        Else
            ReDim xValuesEdad(0)
            ReDim yValuesTallaD_2(0)
        End If
        
    
        If oRsListaTriaje.RecordCount > 0 Then
            ReDim xValuesEdadAtencion(oRsListaTriaje.RecordCount - 1)
            ReDim yValuesTalla(oRsListaTriaje.RecordCount - 1)
            oRsListaTriaje.MoveFirst
            For lnFor = 0 To oRsListaTriaje.RecordCount - 1
                If Not IsNull(oRsListaTriaje.Fields!triajeTalla) And Not IsNull(oRsListaTriaje.Fields!TriajeFecha) Then
                    xValuesEdadAtencion(lnFor) = devuelveEdadTriajeEnMeses(md_fechaNacimiento, oRsListaTriaje.Fields!TriajeFecha)
                    yValuesTalla(lnFor) = CDbl(oRsListaTriaje.Fields!triajeTalla)
               End If
               oRsListaTriaje.MoveNext
            Next
        Else
            ReDim xValuesEdadAtencion(0)
            ReDim yValuesTalla(0)
        End If
        
        ReDim Preserve xValuesEdadAtencion(UBound(xValuesEdadAtencion) + 1)
        ReDim Preserve yValuesTalla(UBound(yValuesTalla) + 1)
    End If
    If Not (mo_DOAtencionesCE Is Nothing) Then
        If Val(mo_DOAtencionesCE.triajeTalla) > 0 Then
            xValuesEdadAtencion(UBound(xValuesEdadAtencion)) = devuelveEdadTriajeEnMeses(md_fechaNacimiento, mo_DOAtencionesCE.TriajeFecha)
            yValuesTalla(UBound(yValuesTalla)) = CDbl(mo_DOAtencionesCE.triajeTalla)
        End If
    End If
    
    lnFor = 0
    
    shaTallaEdad.Clear
    shaTallaEdad.DisplayToolbar = False
    Set owcChart = shaTallaEdad.Charts.Add
    owcChart.HasTitle = True
    owcChart.Title.Caption = "Edad en Meses (X), Talla (Y)"
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
    owcChart.Axes(chAxisPositionLeft).Scaling.Maximum = 130
    owcChart.Axes(1).HasTitle = 1
    owcChart.Axes(1).Font.Name = "Arial Narrow"
    owcChart.Axes(1).Font.Size = 8
    owcChart.Axes(1).Font.Color = vbBlue
    owcChart.Axes(1).Title.Caption = "Talla"
    owcChart.Axes(1).Title.Font.Name = "Arial Narrow"
    owcChart.Axes(1).Title.Font.Size = 8
    owcChart.Axes(1).Title.Font.Color = vbBlue
    
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "Normal"
        .SetData chDimCategories, chDataLiteral, xValuesEdad
        .SetData chDimValues, chDataLiteral, yValuesTallaD0
        .Type = chChartTypeLine
        .Line.Color = sghDesviacionColor.sghNormal
        .Line.Weight = lGrosorLinea
        .Marker.Style = chMarkerStyleNone
        .Line.DashStyle = chLineSolid
    End With
    
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "Desviacion 1"
        .SetData chDimCategories, chDataLiteral, xValuesEdad
        .SetData chDimValues, chDataLiteral, yValuesTallaD1
        .Type = chChartTypeLine
        .Line.Color = sghDesviacionColor.sghDesviacion1
        .Line.Weight = lGrosorLinea
        .Marker.Style = chMarkerStyleNone
        .Line.DashStyle = chLineSolid
    End With
    
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "Desviacion 2"
        .SetData chDimCategories, chDataLiteral, xValuesEdad
        .SetData chDimValues, chDataLiteral, yValuesTallaD2
        .Type = chChartTypeLine
        .Line.Color = sghDesviacionColor.sghDesviacion2
        .Line.Weight = lGrosorLinea
        .Marker.Style = chMarkerStyleNone
        .Line.DashStyle = chLineSolid
    End With
    
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "Desviacion -1"
        .SetData chDimCategories, chDataLiteral, xValuesEdad
        .SetData chDimValues, chDataLiteral, yValuesTallaD_1
        .Type = chChartTypeLine
        .Line.Color = sghDesviacionColor.sghDesviacionMenos1
        .Line.Weight = lGrosorLinea
        .Marker.Style = chMarkerStyleNone
        .Line.DashStyle = chLineSolid
    End With
    
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "Desviacion -2"
        .SetData chDimCategories, chDataLiteral, xValuesEdad
        .SetData chDimValues, chDataLiteral, yValuesTallaD_2
        .Type = chChartTypeLine
        .Line.Color = sghDesviacionColor.sghDesviacionMenos2
        .Line.Weight = lGrosorLinea
        .Marker.Style = chMarkerStyleNone
        .Line.DashStyle = chLineSolid
    End With
    
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "Talla"
        .SetData chDimCategories, chDataLiteral, xValuesEdadAtencion
        .SetData chDimValues, chDataLiteral, yValuesTalla
        .Type = chChartTypeLine
        .Line.Color = sghDesviacionColor.sghValorTriaje
        .Line.Weight = lGrosorLineaValor
        .Marker.Style = chMarkerStyleNone
        .Line.DashStyle = chLineSolid
    End With
    
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical, "Niño Sano"
    End If
End Sub


Public Sub CargaGraficoPesoEdad(lbActualizaDesdeInicioRs As Boolean)
On Error GoTo miError
    Dim lnFor As Integer
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oRsPeso As ADODB.Recordset
    Dim lIdTipoSexo As Long, lGrosorLinea As Integer, lGrosorLineaValor As Integer
'    Dim xValuesEdadPesoAtencion As Variant, yValuesPeso As Variant
    Dim oRsListaTriaje As ADODB.Recordset
    
'    ml_PuntoPesoAtencion = 0
    
    lIdTipoSexo = ml_idTipoSexo
    If lIdTipoSexo = 0 Then
        lIdTipoSexo = sghSexo.Femenino
    End If
    lGrosorLinea = 1
    lGrosorLineaValor = 2
    
    
    If lbActualizaDesdeInicioRs = True Then
        Set oRsListaTriaje = oReglasAtencionIntegral.AtencionesCeListaTriaje(ml_idAtencion, mo_NroHistoriaClinica)
        
        Set oRsPeso = oReglasAtencionIntegral.DevuelveValorPesoPorSexoDesviacion(lIdTipoSexo, sghDesviacion.sghNormal)
        If oRsPeso.RecordCount > 0 Then
            ReDim xValuesEdadPeso(oRsPeso.RecordCount - 1)
            ReDim yValuesPesoD0(oRsPeso.RecordCount - 1)
            oRsPeso.MoveFirst
            For lnFor = 0 To oRsPeso.RecordCount - 1
               xValuesEdadPeso(lnFor) = oRsPeso.Fields!EdadMeses
               yValuesPesoD0(lnFor) = CDbl(oRsPeso.Fields!ValorPeso)
               oRsPeso.MoveNext
            Next
        Else
            ReDim xValuesEdadPeso(0)
            ReDim yValuesPesoD0(0)
        End If
        
        Set oRsPeso = oReglasAtencionIntegral.DevuelveValorPesoPorSexoDesviacion(lIdTipoSexo, sghDesviacion.sghDesviacion1)
        If oRsPeso.RecordCount > 0 Then
            ReDim xValuesEdadPeso(oRsPeso.RecordCount - 1)
            ReDim yValuesPesoD1(oRsPeso.RecordCount - 1)
            oRsPeso.MoveFirst
            For lnFor = 0 To oRsPeso.RecordCount - 1
               xValuesEdadPeso(lnFor) = oRsPeso.Fields!EdadMeses
               yValuesPesoD1(lnFor) = CDbl(oRsPeso.Fields!ValorPeso)
               oRsPeso.MoveNext
            Next
        Else
            ReDim xValuesEdadPeso(0)
            ReDim yValuesPesoD1(0)
        End If
        
        
        Set oRsPeso = oReglasAtencionIntegral.DevuelveValorPesoPorSexoDesviacion(lIdTipoSexo, sghDesviacion.sghDesviacion2)
        If oRsPeso.RecordCount > 0 Then
            ReDim xValuesEdadPeso(oRsPeso.RecordCount - 1)
            ReDim yValuesPesoD2(oRsPeso.RecordCount - 1)
            oRsPeso.MoveFirst
            For lnFor = 0 To oRsPeso.RecordCount - 1
               xValuesEdadPeso(lnFor) = oRsPeso.Fields!EdadMeses
               yValuesPesoD2(lnFor) = CDbl(oRsPeso.Fields!ValorPeso)
               oRsPeso.MoveNext
            Next
        Else
            ReDim xValuesEdadPeso(0)
            ReDim yValuesPesoD2(0)
        End If
        
        Set oRsPeso = oReglasAtencionIntegral.DevuelveValorPesoPorSexoDesviacion(lIdTipoSexo, sghDesviacion.sghDesviacionMenos1)
        
        If oRsPeso.RecordCount > 0 Then
            ReDim xValuesEdadPeso(oRsPeso.RecordCount - 1)
            ReDim yValuesPesoD_1(oRsPeso.RecordCount - 1)
            oRsPeso.MoveFirst
            For lnFor = 0 To oRsPeso.RecordCount - 1
               xValuesEdadPeso(lnFor) = oRsPeso.Fields!EdadMeses
               yValuesPesoD_1(lnFor) = CDbl(oRsPeso.Fields!ValorPeso)
               oRsPeso.MoveNext
            Next
        Else
            ReDim xValuesEdadPeso(0)
            ReDim yValuesPesoD_1(0)
        End If
        
        Set oRsPeso = oReglasAtencionIntegral.DevuelveValorPesoPorSexoDesviacion(lIdTipoSexo, sghDesviacion.sghDesviacionMenos2)
        
        If oRsPeso.RecordCount > 0 Then
            ReDim xValuesEdadPeso(oRsPeso.RecordCount - 1)
            ReDim yValuesPesoD_2(oRsPeso.RecordCount - 1)
            oRsPeso.MoveFirst
            For lnFor = 0 To oRsPeso.RecordCount - 1
               xValuesEdadPeso(lnFor) = oRsPeso.Fields!EdadMeses
               yValuesPesoD_2(lnFor) = CDbl(oRsPeso.Fields!ValorPeso)
               oRsPeso.MoveNext
            Next
        Else
            ReDim xValuesEdadPeso(0)
            ReDim yValuesPesoD_2(0)
        End If
        
    
        If oRsListaTriaje.RecordCount > 0 Then
            ReDim xValuesEdadPesoAtencion(oRsListaTriaje.RecordCount - 1)
            ReDim yValuesPeso(oRsListaTriaje.RecordCount - 1)
            oRsListaTriaje.MoveFirst
            For lnFor = 0 To oRsListaTriaje.RecordCount - 1
                If Not IsNull(oRsListaTriaje.Fields!triajePeso) And Not IsNull(oRsListaTriaje.Fields!TriajeFecha) Then
                    xValuesEdadPesoAtencion(lnFor) = devuelveEdadTriajeEnMeses(md_fechaNacimiento, oRsListaTriaje.Fields!TriajeFecha)
                    yValuesPeso(lnFor) = CDbl(oRsListaTriaje.Fields!triajePeso)
               End If
               oRsListaTriaje.MoveNext
            Next
        Else
            ReDim xValuesEdadPesoAtencion(0)
            ReDim yValuesPeso(0)
        End If
        
        ReDim Preserve xValuesEdadPesoAtencion(UBound(xValuesEdadPesoAtencion) + 1)
        ReDim Preserve yValuesPeso(UBound(yValuesPeso) + 1)
    End If
    If Not (mo_DOAtencionesCE Is Nothing) Then
        If Val(mo_DOAtencionesCE.triajePeso) > 0 Then
            xValuesEdadPesoAtencion(UBound(xValuesEdadPesoAtencion)) = devuelveEdadTriajeEnMeses(md_fechaNacimiento, mo_DOAtencionesCE.TriajeFecha)
            yValuesPeso(UBound(yValuesPeso)) = CDbl(mo_DOAtencionesCE.triajePeso)
        End If
    End If
    
    lnFor = 0
    
    shaPesoEdad.Clear
    shaPesoEdad.DisplayToolbar = False
    Set owcChart = shaPesoEdad.Charts.Add
    owcChart.HasTitle = True
    owcChart.Title.Caption = "Edad en Meses (X), Peso (Y)"
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
    owcChart.Axes(chAxisPositionLeft).Scaling.Maximum = 40
    owcChart.Axes(1).HasTitle = 1
    owcChart.Axes(1).Font.Name = "Arial Narrow"
    owcChart.Axes(1).Font.Size = 8
    owcChart.Axes(1).Font.Color = vbBlue
    owcChart.Axes(1).Title.Caption = "Peso"
    owcChart.Axes(1).Title.Font.Name = "Arial Narrow"
    owcChart.Axes(1).Title.Font.Size = 8
    owcChart.Axes(1).Title.Font.Color = vbBlue
    
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "Normal"
        .SetData chDimCategories, chDataLiteral, xValuesEdadPeso
        .SetData chDimValues, chDataLiteral, yValuesPesoD0
        .Type = chChartTypeLine
        .Line.Color = sghDesviacionColor.sghNormal
        .Line.Weight = lGrosorLinea
        .Marker.Style = chMarkerStyleNone
        .Line.DashStyle = chLineSolid
    End With
    
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "Desviacion 1"
        .SetData chDimCategories, chDataLiteral, xValuesEdadPeso
        .SetData chDimValues, chDataLiteral, yValuesPesoD1
        .Type = chChartTypeLine
        .Line.Color = sghDesviacionColor.sghDesviacion1
        .Line.Weight = lGrosorLinea
        .Marker.Style = chMarkerStyleNone
        .Line.DashStyle = chLineSolid
    End With
    
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "Desviacion 2"
        .SetData chDimCategories, chDataLiteral, xValuesEdadPeso
        .SetData chDimValues, chDataLiteral, yValuesPesoD2
        .Type = chChartTypeLine
        .Line.Color = sghDesviacionColor.sghDesviacion2
        .Line.Weight = lGrosorLinea
        .Marker.Style = chMarkerStyleNone
        .Line.DashStyle = chLineSolid
    End With
    
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "Desviacion -1"
        .SetData chDimCategories, chDataLiteral, xValuesEdadPeso
        .SetData chDimValues, chDataLiteral, yValuesPesoD_1
        .Type = chChartTypeLine
        .Line.Color = sghDesviacionColor.sghDesviacionMenos1
        .Line.Weight = lGrosorLinea
        .Marker.Style = chMarkerStyleNone
        .Line.DashStyle = chLineSolid
    End With
    
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "Desviacion -2"
        .SetData chDimCategories, chDataLiteral, xValuesEdadPeso
        .SetData chDimValues, chDataLiteral, yValuesPesoD_2
        .Type = chChartTypeLine
        .Line.Color = sghDesviacionColor.sghDesviacionMenos2
        .Line.Weight = lGrosorLinea
        .Marker.Style = chMarkerStyleNone
        .Line.DashStyle = chLineSolid
    End With
    
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = "Peso"
        .SetData chDimCategories, chDataLiteral, xValuesEdadPesoAtencion
        .SetData chDimValues, chDataLiteral, yValuesPeso
        .Type = chChartTypeLine
        .Line.Color = sghDesviacionColor.sghValorTriaje
        .Line.Weight = lGrosorLineaValor
        .Marker.Style = chMarkerStyleNone
        .Line.DashStyle = chLineSolid
    End With
    
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbCritical, "Niño Sano"
    End If
End Sub


Private Function devuelveEdadTriajeEnMeses(dFechaNacimiento As Date, dFechaTriaje As Date) As Long
        Dim oEdad As Edad
        Dim lEdad As Long
        
        oEdad = calcularEdadDisgregada(dFechaNacimiento, dFechaTriaje)
        lEdad = (oEdad.EdadAnio * 12) + oEdad.EdadMes
        devuelveEdadTriajeEnMeses = lEdad
End Function
'mgaray201412a
Private Function obtenerLabAutomatico(oDODiagnostico As DODiagnostico) As String
    Dim sCodigoCIECRED As String
    Dim sCodigoCie As String
    Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
    Dim sLab As String
    
    sCodigoCIECRED = UCase(Trim(mo_AdminAdmision.ObtenerCodigoCIEParaAtencionCRED()))
    
    sCodigoCie = UCase(Trim(oDODiagnostico.CodigoCIE2004))
    
    Select Case sCodigoCie
        Case sCodigoCIECRED:
            If frAtencionDesarrollo.Tag <> "" Then
                sLab = frAtencionDesarrollo.Tag
            End If
    End Select
    obtenerLabAutomatico = sLab
End Function
