VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{CFFE0A60-8E3A-11D3-BCC0-00104B9E0792}#1.0#0"; "ssInput1.ocx"
Object = "{0002E558-0000-0000-C000-000000000046}#1.1#0"; "OWC11.DLL"
Begin VB.UserControl ucPerinatalAS 
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11715
   LockControls    =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   11715
   Begin VB.CheckBox chkPT 
      BackColor       =   &H00FFFF00&
      Caption         =   "TP"
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
      Left            =   9450
      TabIndex        =   64
      Top             =   3165
      Width           =   630
   End
   Begin VB.CheckBox chkPe 
      BackColor       =   &H0000FFFF&
      Caption         =   "PE"
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
      Left            =   8700
      TabIndex        =   62
      Top             =   3165
      Width           =   630
   End
   Begin VB.CheckBox chkTe 
      BackColor       =   &H0000FF00&
      Caption         =   "TE"
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
      Left            =   7950
      TabIndex        =   61
      Top             =   3165
      Width           =   630
   End
   Begin VB.CheckBox chkImc 
      BackColor       =   &H000000FF&
      Caption         =   "IMC"
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
      Left            =   7200
      TabIndex        =   60
      Top             =   3165
      Value           =   1  'Checked
      Width           =   630
   End
   Begin VB.CommandButton cmdZoom 
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
      Left            =   11100
      Picture         =   "ucPerinatalAS.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Máximiza gráfico"
      Top             =   3165
      Width           =   405
   End
   Begin VB.ComboBox cmbGraficoSM 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "ucPerinatalAS.ctx":058A
      Left            =   10185
      List            =   "ucPerinatalAS.ctx":0597
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   3165
      Width           =   915
   End
   Begin OWC11.ChartSpace ChartSpace1 
      Height          =   2355
      Left            =   7020
      OleObjectBlob   =   "ucPerinatalAS.ctx":05B1
      TabIndex        =   44
      Top             =   3285
      Width           =   4620
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
      Height          =   1470
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   6945
      Begin VB.CheckBox chkTodaVacuna 
         Caption         =   "Todos"
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
         Left            =   6375
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Clic para MOSTRAR TODAS LAS VACUNAS"
         Top             =   135
         Width           =   510
      End
      Begin VB.CommandButton btnQuitarInmunizacion 
         DisabledPicture =   "ucPerinatalAS.ctx":11A5
         DownPicture     =   "ucPerinatalAS.ctx":1530
         Height          =   315
         Left            =   5925
         Picture         =   "ucPerinatalAS.ctx":18C3
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Elimina todos los CPT"
         Top             =   135
         Width           =   435
      End
      Begin UltraGrid.SSUltraGrid grdInmunizaciones 
         Height          =   990
         Left            =   30
         TabIndex        =   42
         Top             =   420
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   1746
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
         Override        =   "ucPerinatalAS.ctx":1C54
         Caption         =   "grdInmunizaciones"
      End
      Begin ActiveInput.SSComboBoxEx cmbEligeInmunizacion 
         Height          =   345
         Left            =   1470
         TabIndex        =   43
         Top             =   120
         Width           =   4425
         _ExtentX        =   7805
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
      Height          =   1965
      Left            =   0
      TabIndex        =   36
      Top             =   1500
      Width           =   6945
      Begin UltraGrid.SSUltraGrid grdCptFrecuentes 
         Height          =   1515
         Left            =   30
         TabIndex        =   38
         Top             =   420
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   2672
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
         Override        =   "ucPerinatalAS.ctx":1CAA
         Caption         =   "SSUltraGrid1"
      End
      Begin VB.CommandButton btnQuitaOtrosProcedimientos 
         DisabledPicture =   "ucPerinatalAS.ctx":1D00
         DownPicture     =   "ucPerinatalAS.ctx":208B
         Height          =   315
         Left            =   6450
         Picture         =   "ucPerinatalAS.ctx":241E
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Elimina todos los CPT"
         Top             =   120
         Width           =   405
      End
      Begin ActiveInput.SSComboBoxEx cmbProcedimientosFrecuentes 
         Height          =   345
         Left            =   1980
         TabIndex        =   39
         Top             =   120
         Width           =   4470
         _ExtentX        =   7885
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
      Caption         =   "Dx-Desarrollo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   30
      TabIndex        =   33
      Top             =   3510
      Width           =   6945
      Begin VB.CommandButton btnQuitaDxDesarrollo 
         DisabledPicture =   "ucPerinatalAS.ctx":27AF
         DownPicture     =   "ucPerinatalAS.ctx":2B3A
         Height          =   315
         Left            =   6450
         Picture         =   "ucPerinatalAS.ctx":2ECD
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Elimina todos los Dx"
         Top             =   150
         Width           =   375
      End
      Begin ActiveInput.SSComboBoxEx cmbDxDesarrollo 
         Height          =   345
         Left            =   1200
         TabIndex        =   35
         Top             =   180
         Width           =   5250
         _ExtentX        =   9260
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
      Begin UltraGrid.SSUltraGrid grdMorbilidadDesarollo 
         Height          =   1080
         Left            =   45
         TabIndex        =   55
         Top             =   525
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   1905
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
      Height          =   1815
      Left            =   30
      TabIndex        =   29
      Top             =   5220
      Width           =   6945
      Begin VB.CommandButton btnBusquedaDiagnostico 
         Caption         =   "..."
         Height          =   315
         Left            =   6540
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Busca Dx"
         Top             =   120
         Width           =   315
      End
      Begin VB.CommandButton btnQuitaDxMorbilidad 
         DisabledPicture =   "ucPerinatalAS.ctx":325E
         DownPicture     =   "ucPerinatalAS.ctx":35E9
         Height          =   315
         Left            =   6030
         Picture         =   "ucPerinatalAS.ctx":397C
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Elimina todos los Dx"
         Top             =   120
         Width           =   465
      End
      Begin ActiveInput.SSComboBoxEx cmbMorbilidadFrec 
         Height          =   345
         Left            =   1260
         TabIndex        =   32
         Top             =   120
         Width           =   4755
         _ExtentX        =   8387
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
      Begin UltraGrid.SSUltraGrid grdMorbilidadFrec 
         Height          =   1290
         Left            =   45
         TabIndex        =   56
         Top             =   480
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   2275
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
         Caption         =   "grdMorbilidadFrec"
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
      Height          =   1410
      Left            =   7020
      TabIndex        =   0
      Top             =   5625
      Width           =   4575
      Begin UltraGrid.SSUltraGrid grdMedicamentos 
         Height          =   1110
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   1958
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
      Height          =   1260
      Left            =   6960
      TabIndex        =   45
      Top             =   1860
      Width           =   4635
      Begin VB.CheckBox chkLactanciaMaternaComp 
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
         Left            =   45
         TabIndex        =   59
         Top             =   885
         Width           =   3165
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
         TabIndex        =   53
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
         TabIndex        =   52
         Top             =   150
         Width           =   2175
      End
      Begin VB.CheckBox chkLactanciaMaterna 
         Caption         =   "Lactancia materna exclusiva"
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
         TabIndex        =   51
         Top             =   630
         Width           =   3165
      End
      Begin VB.Frame FraAdulto 
         Height          =   1155
         Left            =   1080
         TabIndex        =   46
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
            Top             =   870
            Width           =   3165
         End
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
      Height          =   1875
      Left            =   6990
      TabIndex        =   2
      Top             =   0
      Width           =   4635
      Begin VB.TextBox txtNcontrol 
         Alignment       =   2  'Center
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
         Height          =   270
         Left            =   4110
         TabIndex        =   58
         Text            =   "..."
         Top             =   0
         Visible         =   0   'False
         Width           =   435
      End
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
         ToolTipText     =   "X (control en el Establecimiento),  E (control externo)"
         Top             =   450
         Width           =   285
      End
      Begin UltraGrid.SSUltraGrid grdCred 
         Height          =   1035
         Left            =   60
         TabIndex        =   15
         ToolTipText     =   "Pulse ENTER o DOBLE CLIC "
         Top             =   780
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   1826
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
         Override        =   "ucPerinatalAS.ctx":3D0D
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   240
         Width           =   210
      End
   End
End
Attribute VB_Name = "ucPerinatalAS"
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
Dim ml_idPaciente As Long
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
'

Dim xValues As Variant, yValuesPT As Variant, yValuesTE As Variant, yValuesPE As Variant, yValuesIMC As Variant
Dim owcChart As OWC11.ChChart
Dim owcSeries As OWC11.ChSeries
Dim lnNroPuntosGraficos As Integer
Dim lnIdAtencionCred As Long    'usado para CONTROL
Dim lnIdAtencionCred1 As Long   'usado para los CHECK
Dim ml_FechaAtencion As Date
Dim ml_idUsuario As Long
Dim ml_YaCargoUnaSolaVez As Boolean
Dim ml_EdadEnSemanas As Long
Dim ml_EdadEnAnios As Long
Dim ld_FechaNacimiento As Date
Dim lbEstaMaximizadoElGrafico As Boolean
Dim ln_EdadEnDias As Long
Const labPE As String = "PE ", labPT As String = "TP ", labTE As String = "TE ", labIMC As String = "IMC"
Dim lnPesoKgActual As Double, lnTallaCMActual As Double
Dim lbSeCargaDatosDesdeTablasPerinatal As Boolean
'
Property Let FechaNacimiento(lValue As Date)
    ld_FechaNacimiento = lValue
    ml_EdadEnSemanas = sighentidades.DevuelveEdadEnSemanas(lValue, ml_FechaAtencion)
    ml_EdadEnAnios = DateDiff("yyyy", lValue, ml_FechaAtencion)
    ln_EdadEnDias = DateDiff("d", lValue, ml_FechaAtencion)
End Property
Property Let idAtencion(lValue As Long)
   ml_idAtencion = lValue

End Property

Property Let EdadEnMeses(lValue As Long)
   ml_EdadEnMeses = lValue
End Property
Property Let idTipoSexo(lValue As Long)
   ml_idTipoSexo = lValue
End Property

Property Let idPaciente(lValue As Long)
   ml_idPaciente = lValue
End Property

Property Let FechaAtencion(lValue As Date)
   ml_FechaAtencion = lValue
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Public Sub inicializar()
    
    If ml_YaCargoUnaSolaVez = False Then
        ml_YaCargoUnaSolaVez = True
        mo_Formulario.HabilitarDeshabilitar txtNcontrol, False
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
        cmbGraficoSM.ListIndex = 1
        chkImc.Value = 1
    End If
    lbSeCargaDatosDesdeTablasPerinatal = False
    '
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
           oRsFarmaciaMI.Fields!Medicamento = oRsTmp1.Fields!Nombre
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
          If ln_EdadEnDias <= 28 Then
             lnIdModulo = sighHasta28Dias
             lcEdadCredEnAtencion = lcHasta28Dias
             Cred9.Visible = False: Cred10.Visible = False: Cred11.Visible = False: Cred12.Visible = False
             lbl9.Visible = False: lbl10.Visible = False: lbl11.Visible = False: lbl12.Visible = False
          Else
             lnIdModulo = sighDesde29diasHasta1anio
             lcEdadCredEnAtencion = lcDe29diasHasta1anio
             Cred12.Visible = False
             lbl12.Visible = False
          End If
     Case 3  'dias
          lnEdadEnAniosEnAtencion = 0
          If lnEdad <= 28 Then
             lnIdModulo = sighHasta28Dias
             lcEdadCredEnAtencion = lcHasta28Dias
             Cred3.Visible = False: Cred4.Visible = False: Cred5.Visible = False: Cred6.Visible = False
             lbl3.Visible = False: lbl4.Visible = False: lbl5.Visible = False: lbl6.Visible = False
             Cred7.Visible = False: Cred8.Visible = False
             lbl7.Visible = False: lbl8.Visible = False
             Cred9.Visible = False: Cred10.Visible = False: Cred11.Visible = False: Cred12.Visible = False
             lbl9.Visible = False: lbl10.Visible = False: lbl11.Visible = False: lbl12.Visible = False
          Else
             lnIdModulo = sighDesde29diasHasta1anio
             lcEdadCredEnAtencion = lcDe29diasHasta1anio
             Cred12.Visible = False
             lbl12.Visible = False
          End If
     Case Else    'horas
          lnEdadEnAniosEnAtencion = 0
          lnIdModulo = sighHasta28Dias
          lcEdadCredEnAtencion = lcHasta28Dias
          Cred3.Visible = False: Cred4.Visible = False: Cred5.Visible = False: Cred6.Visible = False
          lbl3.Visible = False: lbl4.Visible = False: lbl5.Visible = False: lbl6.Visible = False
          Cred7.Visible = False: Cred8.Visible = False
          lbl7.Visible = False: lbl8.Visible = False
          Cred9.Visible = False: Cred10.Visible = False: Cred11.Visible = False: Cred12.Visible = False
          lbl9.Visible = False: lbl10.Visible = False: lbl11.Visible = False: lbl12.Visible = False
     End Select
     'Inicializa controles Checks segun modulos
     chkEstimulacionTemprana.Visible = False
     chkAlimentacionComplementaria.Visible = False
     chkLactanciaMaterna.Visible = False
     chkLactanciaMaternaComp.Visible = False
     FraAdulto.Left = 1060
     FraAdulto.Visible = False
     Select Case lnIdModulo
     Case sighHasta28Dias
            chkLactanciaMaterna.Visible = True
            chkEstimulacionTemprana.Visible = True
     Case sighDesde29diasHasta1anio
            chkLactanciaMaterna.Visible = True
            chkAlimentacionComplementaria.Visible = True
            chkEstimulacionTemprana.Visible = True
     Case sighDesde1Hasta4anios
            chkEstimulacionTemprana.Visible = True
            chkLactanciaMaternaComp.Visible = True
     Case sighDesde18anios
            FraAdulto.Left = 60
            FraAdulto.Visible = True
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
                cmbMorbilidadFrec.ListItems.Add lnIdListItem, lcCombo + Trim(Str(oRsTmp.Fields!IdDiagnostico)), oRsTmp.Fields!Descripcion
                lnIdListItem = lnIdListItem + 1
           Else
                cmbDxDesarrollo.ListItems.Add lnIdListItem1, lcCombo + Trim(Str(oRsTmp.Fields!IdDiagnostico)), oRsTmp.Fields!Descripcion
                lnIdListItem1 = lnIdListItem1 + 1
           End If
           '
           If (Not IsNull(oRsTmp.Fields!rangoInicio)) And (Not IsNull(oRsTmp.Fields!rangoFinal)) Then
                oRsDxDesarrolloAutomaticos.AddNew
                oRsDxDesarrolloAutomaticos.Fields!IdDiagnostico = oRsTmp.Fields!IdDiagnostico
                oRsDxDesarrolloAutomaticos.Fields!rangoInicio = oRsTmp.Fields!rangoInicio
                oRsDxDesarrolloAutomaticos.Fields!rangoFinal = oRsTmp.Fields!rangoFinal
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
              cmbEligeInmunizacion.ListItems.Add lnIdListItem, lcCombo + Trim(Str(oRsTmp.Fields!idProducto)), oRsTmp.Fields!Nombre
              lnIdListItem = lnIdListItem + 1
           Else
              cmbProcedimientosFrecuentes.ListItems.Add lnIdListItem1, lcCombo + Trim(Str(oRsTmp.Fields!idProducto)), oRsTmp.Fields!Nombre
              lnIdListItem1 = lnIdListItem1 + 1
           End If
           oRsTmp.MoveNext
        Loop
     End If
     oRsTmp.Close
     'inicialmente debe cargar todas las INMUNIZACIONES
     chkTodaVacuna.Value = 1
     chkTodaVacuna_Click
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

Sub LlenaComboInmunizacionesTOTALoSEGUN_EDAD(lbLlenaTotal As Boolean)
     Dim oConexion As New Connection
     Dim oRsTmp As New Recordset
     Dim lnIdListItem As Integer
     Dim lcID As String
     oConexion.CommandTimeout = 900
     oConexion.CursorLocation = adUseClient
     oConexion.Open sighentidades.CadenaConexion
     cmbEligeInmunizacion.Clear
     If lbLlenaTotal = True Then
        Set oRsTmp = mo_reglasComunes.PerinatalCatalogoCptSeleccionarPorIdModulo(0, oConexion)
     Else
        Set oRsTmp = mo_reglasComunes.PerinatalCatalogoCptSeleccionarPorIdModulo(lnIdModulo, oConexion)
     End If
     oRsTmp.Filter = "idLista=" & sghPerinatalListas.sighInmunizaciones
     If oRsTmp.RecordCount > 0 Then
        lnIdListItem = 0
        lcID = "/"
        oRsTmp.MoveFirst
        Do While Not oRsTmp.EOF
           If InStr(lcID, oRsTmp!idProducto) = 0 Then
              cmbEligeInmunizacion.ListItems.Add lnIdListItem, lcCombo + Trim(Str(oRsTmp.Fields!idProducto)), oRsTmp.Fields!Nombre
              lnIdListItem = lnIdListItem + 1
              lcID = lcID & Trim(Str(oRsTmp!idProducto)) & "/"
           End If
           oRsTmp.MoveNext
        Loop
     End If
     oRsTmp.Close
     oConexion.Close
     Set oRsTmp = Nothing
     Set oConexion = Nothing
End Sub

Sub LimpiaMedicamentos()
    CreaTemporalFarmacia
End Sub

Private Sub btnBusquedaDiagnostico_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
    Dim oDODiagnostico As DODiagnostico
    oBusqueda.CodigoDx = ""
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_reglasComunes.DiagnosticosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            If oRsMorbilidadFrec.RecordCount > 0 Then
               oRsMorbilidadFrec.MoveFirst
               oRsMorbilidadFrec.Find "id=" & oDODiagnostico.IdDiagnostico
               If Not oRsMorbilidadFrec.EOF Then
                  MsgBox "Ese Dx ya está registrado", vbInformation, "Mensaje"
                  Exit Sub
               End If
            End If
            oRsMorbilidadFrec.AddNew
            oRsMorbilidadFrec.Fields!Id = oDODiagnostico.IdDiagnostico
            oRsMorbilidadFrec.Fields!DIAGNOSTICO = oDODiagnostico.Descripcion
            oRsMorbilidadFrec.Fields!idAtencion = ml_idAtencion
            oRsMorbilidadFrec.Fields!SeEligioConChek = False
            oRsMorbilidadFrec.Fields!EsDxPerinatal = False
            oRsMorbilidadFrec.Fields!IdClasificacionDx = sghTiposDiagnostico.sghAtencionConsultaExterna
            oRsMorbilidadFrec.Fields!IdSubclasificacionDx = 1   'sghDxDefinitivos.sighDxCeDefinitivo
            oRsMorbilidadFrec.Update
        End If
    End If
    Set oBusqueda = Nothing
End Sub

Private Sub btnQuitaDxDesarrollo_Click()
    LimpiaDxDesarrollo
    '
    Dim lnFor As Integer, lnFor1 As Integer
    On Error Resume Next
    Do While True
        For lnFor1 = 1 To 3
            For lnFor = 0 To cmbDxDesarrollo.ListCount - 1
                cmbDxDesarrollo.SelectedItems(lnFor).Selected = False
            Next
        Next
        If cmbDxDesarrollo.SelectedItems.Count = 0 Then
           Exit Do
        End If
    Loop
End Sub

Private Sub btnQuitaDxMorbilidad_Click()
    LimpiaMorbilidadFrecuente True
    '
    Dim lnFor As Integer, lnFor1 As Integer
    On Error Resume Next
    Do While True
        For lnFor1 = 1 To 3
            For lnFor = 0 To cmbMorbilidadFrec.ListCount - 1
                cmbMorbilidadFrec.SelectedItems(lnFor).Selected = False
            Next
        Next
        If cmbMorbilidadFrec.SelectedItems.Count = 0 Then
           Exit Do
        End If
    Loop
End Sub

Private Sub btnQuitaOtrosProcedimientos_Click()
    LimpiaCPTFrecuentes
    '
    Dim lnFor As Integer, lnFor1 As Integer
    On Error Resume Next
    Do While True
        For lnFor1 = 1 To 3
            For lnFor = 0 To cmbProcedimientosFrecuentes.ListCount - 1
                cmbProcedimientosFrecuentes.SelectedItems(lnFor).Selected = False
            Next
        Next
        If cmbProcedimientosFrecuentes.SelectedItems.Count = 0 Then
           Exit Do
        End If
    Loop
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







Private Sub chkImc_Click()
    cmbGraficoSM_Click
End Sub

Private Sub chkPe_Click()
     cmbGraficoSM_Click
End Sub

Private Sub chkPT_Click()
    cmbGraficoSM_Click
End Sub

Private Sub chkTe_Click()
    cmbGraficoSM_Click
End Sub

Private Sub chkTodaVacuna_Click()
    If chkTodaVacuna.Value = 1 Then
       chkTodaVacuna.Caption = "Edad"
       LlenaComboInmunizacionesTOTALoSEGUN_EDAD True
       chkTodaVacuna.ToolTipText = "Clic para mostrar solo VACUNAS para la EDAD"
    Else
       chkTodaVacuna.Caption = "Todos"
       LlenaComboInmunizacionesTOTALoSEGUN_EDAD False
       chkTodaVacuna.ToolTipText = "Clic para MOSTRAR TODAS LAS VACUNAS"
    End If
End Sub

Private Sub cmbDxDesarrollo_LostFocus()
Dim sItems As String
Dim CBLI As SSCBListItem
Dim iTrimLen As Integer
     LimpiaDxDesarrollo
     If cmbDxDesarrollo.SelectedItems.Count > 0 Then
          For Each CBLI In cmbDxDesarrollo.SelectedItems
              oRsDxDesarrollo.AddNew
              oRsDxDesarrollo.Fields!Id = Val(Mid(CBLI.Key, 2, 100))
              oRsDxDesarrollo.Fields!DIAGNOSTICO = CBLI.Text
              oRsDxDesarrollo.Fields!idAtencion = ml_idAtencion
              oRsDxDesarrollo.Fields!IdClasificacionDx = sghTiposDiagnostico.sghAtencionConsultaExterna
              oRsDxDesarrollo.Fields!IdSubclasificacionDx = 2  'sghDxDefinitivos.sighDxCeDefinitivo
              oRsDxDesarrollo.Update
          Next CBLI
          On Error Resume Next
          oRsDxDesarrollo.MoveFirst
     End If
     
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
          .Fields.Append "Diagnostico", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdAtencion", adInteger
          .Fields.Append "IdClasificacionDx", adInteger
          .Fields.Append "IdSubclasificacionDx", adInteger
          .Fields.Append "CodigoCIE2004", adVarChar, 7, adFldIsNullable
           .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable + adFldUpdatable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdMorbilidadDesarollo.DataSource = oRsDxDesarrollo
    mo_Apariencia.ConfigurarFilasBiColores grdMorbilidadDesarollo, sighentidades.GrillaConFilasBicolor
    '
    With oRsMorbilidadFrec
          .Fields.Append "Id", adInteger
          .Fields.Append "Diagnostico", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdAtencion", adInteger
          .Fields.Append "SeEligioConChek", adBoolean
          .Fields.Append "EsDxPerinatal", adBoolean
          .Fields.Append "IdClasificacionDx", adInteger
          .Fields.Append "IdSubclasificacionDx", adInteger
          .Fields.Append "CodigoCIE2004", adVarChar, 7, adFldIsNullable
          .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable + adFldUpdatable
          .CursorType = adOpenDynamic
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
Dim sItems As String
Dim CBLI As SSCBListItem
Dim iTrimLen As Integer
     LimpiaMorbilidadFrecuente False
     If cmbMorbilidadFrec.SelectedItems.Count > 0 Then
          For Each CBLI In cmbMorbilidadFrec.SelectedItems
              oRsMorbilidadFrec.AddNew
              oRsMorbilidadFrec.Fields!Id = Val(Mid(CBLI.Key, 2, 100))
              oRsMorbilidadFrec.Fields!DIAGNOSTICO = CBLI.Text
              oRsMorbilidadFrec.Fields!idAtencion = ml_idAtencion
              oRsMorbilidadFrec.Fields!SeEligioConChek = True
              oRsMorbilidadFrec.Fields!EsDxPerinatal = True
              oRsMorbilidadFrec.Fields!IdClasificacionDx = sghTiposDiagnostico.sghAtencionConsultaExterna
              oRsMorbilidadFrec.Fields!IdSubclasificacionDx = 1
              
              oRsMorbilidadFrec.Update
          Next CBLI
          On Error Resume Next
          oRsMorbilidadFrec.MoveFirst
     End If

End Sub

Private Sub cmbProcedimientosFrecuentes_LostFocus()
Dim sItems As String
Dim CBLI As SSCBListItem
Dim iTrimLen As Integer
     LimpiaCPTFrecuentes
     If cmbProcedimientosFrecuentes.SelectedItems.Count > 0 Then
          For Each CBLI In cmbProcedimientosFrecuentes.SelectedItems
              oRsCptFrecuentes.AddNew
              oRsCptFrecuentes.Fields!Id = Val(Mid(CBLI.Key, 2, 100))
              oRsCptFrecuentes.Fields!procedimiento = CBLI.Text
              oRsCptFrecuentes.Fields!idAtencion = ml_idAtencion
              oRsCptFrecuentes.Update
          Next CBLI
          On Error Resume Next
          oRsCptFrecuentes.MoveFirst
     End If
End Sub





Private Sub cmbGraficoSM_Click()
    GraficoRegistraDatosParaFilaColumnas True
    CargaGraficoChartSpace True
End Sub

Private Sub cmdZoom_Click()
    If lbEstaMaximizadoElGrafico = False Then
       lbEstaMaximizadoElGrafico = True
       cmbGraficoSM.Top = 0
       cmdZoom.Top = 0
       ChartSpace1.Top = 0
       ChartSpace1.Height = UserControl.Height - 2700
       ChartSpace1.Width = UserControl.Width - 50
       ChartSpace1.Left = 0
       chkImc.Top = 0
       chkPe.Top = 0
       chkTe.Top = 0
       chkPT.Top = 0
    Else
       lbEstaMaximizadoElGrafico = False
       cmbGraficoSM.Top = 3165     '2820
       cmdZoom.Top = 3165   '2820
       ChartSpace1.Top = 3285  '2985
       ChartSpace1.Height = 2355    '2520
       ChartSpace1.Width = 4620
       ChartSpace1.Left = 7020
       chkImc.Top = 3165   '2820
       chkPe.Top = 3165   '2820
       chkTe.Top = 3165   '2820
       chkPT.Top = 3165   '2820
    End If
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
    Cred10.Text = UCase(Cred10.Text)
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


Private Sub grdCptFrecuentes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdCptFrecuentes
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
    Dim oRsTmp198 As New Recordset
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
         oGrilla.Bands(0).Columns("medicamento").Width = 3400
    Case "grdMorbilidadDesarollo"
         oGrilla.Bands(0).Columns("Id").Hidden = True
         oGrilla.Bands(0).Columns("IdAtencion").Hidden = True
         oGrilla.Bands(0).Columns("diagnostico").Header.Caption = "Diagnóstico"
         oGrilla.Bands(0).Columns("diagnostico").Width = 5000
         oGrilla.Bands(0).Columns("labConfHIS").Width = 400
         oGrilla.Bands(0).Columns("labConfHIS").Header.Caption = "  "
         oGrilla.Bands(0).Columns("codigoCIE2004").Hidden = True
         oGrilla.Bands(0).Columns("IdClasificacionDx").Hidden = True
         '
         oGrilla.Bands(0).Columns("IdSubClasificacionDx").Style = ssStyleDropDownList
         oGrilla.Bands(0).Columns("IdSubClasificacionDx").Activation = ssActivationAllowEdit
         oGrilla.Bands(0).Columns("IdSubClasificacionDx").Width = 700
         oGrilla.Bands(0).Columns("IdSubClasificacionDx").Header.Caption = "Tipo"
         CargaListaComboSubClasificacionDx oGrilla, oGrilla.Bands(0).Columns("IdSubClasificacionDx")
    Case "grdMorbilidadFrec"
         oGrilla.Bands(0).Columns("Id").Hidden = True
         oGrilla.Bands(0).Columns("IdAtencion").Hidden = True
         oGrilla.Bands(0).Columns("SeEligioConChek").Hidden = True
         oGrilla.Bands(0).Columns("EsDxPerinatal").Hidden = True
         oGrilla.Bands(0).Columns("diagnostico").Header.Caption = "Diagnóstico"
         oGrilla.Bands(0).Columns("diagnostico").Width = 5300
         oGrilla.Bands(0).Columns("labConfHIS").Hidden = True
         oGrilla.Bands(0).Columns("codigoCIE2004").Hidden = True
         oGrilla.Bands(0).Columns("IdClasificacionDx").Hidden = True
         '
         oGrilla.Bands(0).Columns("IdSubClasificacionDx").Style = ssStyleDropDownList
         oGrilla.Bands(0).Columns("IdSubClasificacionDx").Activation = ssActivationAllowEdit
         oGrilla.Bands(0).Columns("IdSubClasificacionDx").Width = 700
         oGrilla.Bands(0).Columns("IdSubClasificacionDx").Header.Caption = "Tipo"
         CargaListaComboSubClasificacionDx oGrilla, oGrilla.Bands(0).Columns("IdSubClasificacionDx")
    Case "grdCptFrecuentes"
         oGrilla.Bands(0).Columns("Id").Hidden = True
         oGrilla.Bands(0).Columns("IdAtencion").Hidden = True
         oGrilla.Bands(0).Columns("Procedimiento").Header.Caption = "Procedimiento"
         oGrilla.Bands(0).Columns("Procedimiento").Width = 6300
         oGrilla.Bands(0).Columns("labConfHIS").Hidden = True
         '
    End Select
    Set oRsTmp198 = Nothing
End Sub

Sub CargaListaComboSubClasificacionDx(oGrilla As SSUltraGrid, oColumn As SSColumn)
    'On Error GoTo ErrCLCSD
'    Dim oRsTmp198 As New Recordset
'    Set oRsTmp198 = mo_reglasComunes.SubclasificacionDiagnosticosSeleccionarDxConsultaExterna
'    If oRsTmp198.RecordCount > 0 Then
'       oGrilla.Bands(0).Columns("IdSubClasificacionDx").Width = 1000
'       With oGrilla.ValueLists.Add("DxPrincipal9").ValueListItems
'            oRsTmp198.MoveFirst
'            Do While Not oRsTmp198.EOF
'               .Add oRsTmp198!IdSubclasificacionDx, oRsTmp198!DescripcionLarga
'               oRsTmp198.MoveNext
'            Loop
'       End With
'       oGrilla.Bands(0).Columns("IdSubClasificacionDx").ValueList = "DxPrincipal9"
'    End If
'ErrCLCSD:
'    Set oRsTmp198 = Nothing


Dim rs As ADODB.Recordset
Dim i As Integer
Dim oValueEstado As SSValueList
    If Not oGrilla.ValueLists.Exists("listaTipoDx") Then
        Set oValueEstado = oGrilla.ValueLists.Add("listaTipoDx")
        oValueEstado.ValueListItems.Add 1, "Presuntivo"
        oValueEstado.ValueListItems.Add 2, "Definitivo"
        oValueEstado.ValueListItems.Add 3, "Repetido"

'        Set rs = mo_reglasComunes.SubclasificacionDiagnosticosSeleccionarDxConsultaExterna
'        Do While Not rs.EOF
'            oValueEstado.ValueListItems.Add rs!IdSubclasificacionDx, Trim(rs!DescripcionLarga)
'            'oValueEstado.ValueListItems.Add Trim(rs!DescripcionLarga), rs!IdSubclasificacionDx
'            rs.MoveNext
'        Loop
'        rs.Close
    Else
        Set oValueEstado = oGrilla.ValueLists.Item("listaTipoDx")
    End If

    Set oColumn.ValueList = oValueEstado

End Sub

Private Sub grdInmunizaciones_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdInmunizaciones
End Sub


Private Sub grdMedicamentos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdMedicamentos
End Sub

Private Sub grdMorbilidadDesarollo_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     InicializarLaGrilla grdMorbilidadDesarollo

    
End Sub

Private Sub grdMorbilidadFrec_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
     InicializarLaGrilla grdMorbilidadFrec
     '

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
        .CredN = Val(txtNcontrol.Text)
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
        .LactanciaMaternaComp = IIf(chkLactanciaMaternaComp.Value = 1, True, False)
    End With
    Set DevuelvePerinatalAtencionCred1 = oDoPerinatalAtencionCred1
End Function

Public Function DevuelveDatosCred(lbSoloTextoActivo As Boolean) As Recordset
    Dim oRsCred As New Recordset
    Dim lbSiguee As Boolean
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
       lbSiguee = True
       If lbSoloTextoActivo = True And Cred1.Locked = True Then
          lbSiguee = False
       End If
       If lbSiguee = True Then
            oRsCred.AddNew
            oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
            oRsCred.Fields!credNumero = 1
            oRsCred.Fields!credCheck = Cred1.Text
            oRsCred.Fields!idAtencion = ml_idAtencion
            oRsCred.Update
       End If
    End If
    If Cred2.Visible = True And Cred2.Enabled = True And Cred2.Text <> "" Then
       lbSiguee = True
       If lbSoloTextoActivo = True And Cred2.Locked = True Then
          lbSiguee = False
       End If
       If lbSiguee = True Then
            oRsCred.AddNew
            oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
            oRsCred.Fields!credNumero = 2
            oRsCred.Fields!credCheck = Cred2.Text
            oRsCred.Fields!idAtencion = ml_idAtencion
            oRsCred.Update
       End If
    End If
    If Cred3.Visible = True And Cred3.Enabled = True And Cred3.Text <> "" Then
       lbSiguee = True
       If lbSoloTextoActivo = True And Cred3.Locked = True Then
          lbSiguee = False
       End If
       If lbSiguee = True Then
            oRsCred.AddNew
            oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
            oRsCred.Fields!credNumero = 3
            oRsCred.Fields!credCheck = Cred3.Text
            oRsCred.Fields!idAtencion = ml_idAtencion
            oRsCred.Update
       End If
    End If
    If Cred4.Visible = True And Cred4.Enabled = True And Cred4.Text <> "" Then
       lbSiguee = True
       If lbSoloTextoActivo = True And Cred4.Locked = True Then
          lbSiguee = False
       End If
       If lbSiguee = True Then
            oRsCred.AddNew
            oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
            oRsCred.Fields!credNumero = 4
            oRsCred.Fields!credCheck = Cred4.Text
            oRsCred.Fields!idAtencion = ml_idAtencion
            oRsCred.Update
       End If
    End If
    If Cred5.Visible = True And Cred5.Enabled = True And Cred5.Text <> "" Then
       lbSiguee = True
       If lbSoloTextoActivo = True And Cred5.Locked = True Then
          lbSiguee = False
       End If
       If lbSiguee = True Then
            oRsCred.AddNew
            oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
            oRsCred.Fields!credNumero = 5
            oRsCred.Fields!credCheck = Cred5.Text
            oRsCred.Fields!idAtencion = ml_idAtencion
            oRsCred.Update
       End If
    End If
    If Cred6.Visible = True And Cred6.Enabled = True And Cred6.Text <> "" Then
       lbSiguee = True
       If lbSoloTextoActivo = True And Cred6.Locked = True Then
          lbSiguee = False
       End If
       If lbSiguee = True Then
            oRsCred.AddNew
            oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
            oRsCred.Fields!credNumero = 6
            oRsCred.Fields!credCheck = Cred6.Text
            oRsCred.Fields!idAtencion = ml_idAtencion
            oRsCred.Update
       End If
    End If
    If Cred7.Visible = True And Cred7.Enabled = True And Cred7.Text <> "" Then
       lbSiguee = True
       If lbSoloTextoActivo = True And Cred7.Locked = True Then
          lbSiguee = False
       End If
       If lbSiguee = True Then
            oRsCred.AddNew
            oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
            oRsCred.Fields!credNumero = 7
            oRsCred.Fields!credCheck = Cred7.Text
            oRsCred.Fields!idAtencion = ml_idAtencion
            oRsCred.Update
       End If
    End If
    If Cred8.Visible = True And Cred8.Enabled = True And Cred8.Text <> "" Then
       lbSiguee = True
       If lbSoloTextoActivo = True And Cred8.Locked = True Then
          lbSiguee = False
       End If
       If lbSiguee = True Then
            oRsCred.AddNew
            oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
            oRsCred.Fields!credNumero = 8
            oRsCred.Fields!credCheck = Cred8.Text
            oRsCred.Fields!idAtencion = ml_idAtencion
            oRsCred.Update
       End If
    End If
    If Cred9.Visible = True And Cred9.Enabled = True And Cred9.Text <> "" Then
       lbSiguee = True
       If lbSoloTextoActivo = True And Cred9.Locked = True Then
          lbSiguee = False
       End If
       If lbSiguee = True Then
            oRsCred.AddNew
            oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
            oRsCred.Fields!credNumero = 9
            oRsCred.Fields!credCheck = Cred9.Text
            oRsCred.Fields!idAtencion = ml_idAtencion
            oRsCred.Update
       End If
    End If
    If Cred10.Visible = True And Cred10.Enabled = True And Cred10.Text <> "" Then
       lbSiguee = True
       If lbSoloTextoActivo = True And Cred10.Locked = True Then
          lbSiguee = False
       End If
       If lbSiguee = True Then
            oRsCred.AddNew
            oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
            oRsCred.Fields!credNumero = 10
            oRsCred.Fields!credCheck = Cred10.Text
            oRsCred.Fields!idAtencion = ml_idAtencion
            oRsCred.Update
       End If
    End If
    If Cred11.Visible = True And Cred11.Enabled = True And Cred11.Text <> "" Then
       lbSiguee = True
       If lbSoloTextoActivo = True And Cred11.Locked = True Then
          lbSiguee = False
       End If
       If lbSiguee = True Then
            oRsCred.AddNew
            oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
            oRsCred.Fields!credNumero = 11
            oRsCred.Fields!credCheck = Cred11.Text
            oRsCred.Fields!idAtencion = ml_idAtencion
            oRsCred.Update
       End If
    End If
    If Cred12.Visible = True And Cred12.Enabled = True And Cred12.Text <> "" Then
       lbSiguee = True
       If lbSoloTextoActivo = True And Cred12.Locked = True Then
          lbSiguee = False
       End If
       If lbSiguee = True Then
            oRsCred.AddNew
            oRsCred.Fields!edadEnAnios = lcEdadCredEnAtencion
            oRsCred.Fields!credNumero = 12
            oRsCred.Fields!credCheck = Cred12.Text
            oRsCred.Fields!idAtencion = ml_idAtencion
            oRsCred.Update
       End If
    End If
    '
    Set DevuelveDatosCred = oRsCred
End Function

Public Function DevuelveCptInmunizaciones() As Recordset
    Set DevuelveCptInmunizaciones = oRsInmunizaciones
End Function
Public Function DevuelveCptFrecuentes() As Recordset
    Set DevuelveCptFrecuentes = oRsCptFrecuentes
End Function
Public Function DevuelveDxDesarrollo() As Recordset
 

    If oRsDxDesarrollo.RecordCount > 0 Then
       oRsDxDesarrollo.MoveFirst
       Do While Not oRsDxDesarrollo.EOF
          If oRsDxDesarrollo!IdSubclasificacionDx < 4 Then
             oRsDxDesarrollo!IdSubclasificacionDx = oRsDxDesarrollo!IdSubclasificacionDx + 100
             oRsDxDesarrollo.Update
          End If
          oRsDxDesarrollo.MoveNext
       Loop
    End If
    Set DevuelveDxDesarrollo = oRsDxDesarrollo
End Function
Public Function DevuelveDxMorbilidad() As Recordset
    If oRsMorbilidadFrec.RecordCount > 0 Then
       oRsMorbilidadFrec.MoveFirst
       Do While Not oRsMorbilidadFrec.EOF
          If oRsMorbilidadFrec!IdSubclasificacionDx < 4 Then
             oRsMorbilidadFrec!IdSubclasificacionDx = oRsMorbilidadFrec!IdSubclasificacionDx + 100
             oRsMorbilidadFrec.Update
          End If
          oRsMorbilidadFrec.MoveNext
       Loop
    End If
    Set DevuelveDxMorbilidad = oRsMorbilidadFrec
End Function

Public Function DevuelveMedicamentos() As Recordset
    On Error GoTo errDM
    If oRsFarmaciaMI.RecordCount > 0 Then
       oRsFarmaciaMI.MoveFirst
       Do While Not oRsFarmaciaMI.EOF
          If oRsFarmaciaMI.Fields!seleccionar = True Then
             oRsFarmaciaMI.Fields!idAtencion = ml_idAtencion
             oRsFarmaciaMI.Update
          End If
          oRsFarmaciaMI.MoveNext
       Loop
    End If
    Set DevuelveMedicamentos = oRsFarmaciaMI
    Exit Function
errDM:
    Set DevuelveMedicamentos = Nothing
End Function



Public Sub CargaDatosAcontroles(lnEdad As Integer, lnIdTipoEdad As Integer, lnPesoKg As Double, lnTallaCM As Long, oConexion As Connection)
    lnPesoKgActual = lnPesoKg
    lnTallaCMActual = lnTallaCM
    
    '
    CargaDatosAcombos lnEdad, lnIdTipoEdad, oConexion
    '
    Dim oDoPerinatalAtencion As New DoPerinatalAtencion, oPerinatalAtencion As New PerinatalAtencion
    Dim oDoPerinatalAtencionCred1 As New DoPerinatalAtencionCred1, oPerinatalAtencionCred1 As New PerinatalAtencionCred1
    Dim oRsTmp As New Recordset
    Dim lnFor As Integer, lbEsDxPerinatal As Boolean, lbSiTieneVacuna As Boolean
'    '
'    FraAdulto.Left = 1060
'    FraAdulto.Visible = False
'    If lnIdModulo = sighDesde18anios Then
'        FraAdulto.Left = 60
'        FraAdulto.Visible = True
'    End If
'    '
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
    chkLactanciaMaternaComp.Value = 0
    lnIdAtencionCred1 = 0
    mo_idPerinatalAtencion = 0
    txtNcontrol.Text = 1
    If oPerinatalAtencion.SeleccionarPorIdAtencion(oDoPerinatalAtencion, ml_idAtencion) = True Then
       lbSeCargaDatosDesdeTablasPerinatal = True
       txtNcontrol.Text = oDoPerinatalAtencion.CredN
       'mo_Formulario.HabilitarDeshabilitar txtNcontrol, True
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
            chkLactanciaMaternaComp.Value = IIf(oDoPerinatalAtencionCred1.LactanciaMaternaComp = True, 1, 0)
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
                    oRsInmunizaciones.Fields!procedimiento = oRsTmp.Fields!Nombre
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
                oRsCptFrecuentes.Fields!procedimiento = oRsTmp.Fields!Nombre
                oRsCptFrecuentes.Fields!idAtencion = oRsTmp.Fields!idAtencion
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
                    If cmbDxDesarrollo.ListItems.Item(lnFor).Key = lcCombo & Trim(Str(oRsTmp.Fields!IdDiagnostico)) Then
                       cmbDxDesarrollo.ListItems.Item(lnFor).Selected = True
                       Exit For
                    End If
                Next
                oRsDxDesarrollo.AddNew
                oRsDxDesarrollo.Fields!Id = oRsTmp.Fields!IdDiagnostico
                oRsDxDesarrollo.Fields!DIAGNOSTICO = oRsTmp.Fields!Descripcion
                oRsDxDesarrollo.Fields!idAtencion = oRsTmp.Fields!idAtencion
                oRsDxDesarrollo.Fields!IdClasificacionDx = sghTiposDiagnostico.sghAtencionConsultaExterna
                oRsDxDesarrollo.Fields!IdSubclasificacionDx = oRsTmp.Fields!IdSubclasificacionDx - 100  'sghDxDefinitivos.sighDxCeDefinitivo
                oRsDxDesarrollo.Fields!labConfHIS = oRsTmp!labConfHIS
                oRsDxDesarrollo.Update
             Else
                lbEsDxPerinatal = False
                For lnFor = 0 To cmbMorbilidadFrec.ListCount - 1
                    If cmbMorbilidadFrec.ListItems.Item(lnFor).Key = lcCombo & Trim(Str(oRsTmp.Fields!IdDiagnostico)) Then
                       cmbMorbilidadFrec.ListItems.Item(lnFor).Selected = True
                       lbEsDxPerinatal = True
                       Exit For
                    End If
                Next
                oRsMorbilidadFrec.AddNew
                oRsMorbilidadFrec.Fields!Id = oRsTmp.Fields!IdDiagnostico
                oRsMorbilidadFrec.Fields!DIAGNOSTICO = oRsTmp.Fields!Descripcion
                oRsMorbilidadFrec.Fields!idAtencion = oRsTmp.Fields!idAtencion
                oRsMorbilidadFrec.Fields!SeEligioConChek = True
                oRsMorbilidadFrec.Fields!EsDxPerinatal = lbEsDxPerinatal
                oRsMorbilidadFrec.Fields!IdClasificacionDx = sghTiposDiagnostico.sghAtencionConsultaExterna
                oRsMorbilidadFrec.Fields!IdSubclasificacionDx = oRsTmp.Fields!IdSubclasificacionDx - 100  'sghDxDefinitivos.sighDxCeDefinitivo
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
        CalculaPercentiles lnPesoKg, lnTallaCM
        GraficoRegistraDatosParaFilaColumnas True
        CargaDxAutomaticosParaMorbilidadEnDesarrollo lnPesoKg, lnTallaCM
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
    Dim lcEdad As String, lcCred As String, lnPos As Integer, lcIzquierda As String, lcCentro As String, lcDerecha As String
    Set oRsTmp = mo_reglasComunes.PerinatalAtencionCredSeleccionarPorIdPaciente(ml_idPaciente, oConexion)
    oRsTmp.Filter = "fechaAtencion<=" & ml_FechaAtencion
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
    If oRsCredHistorico.RecordCount > 0 Then
       oRsCredHistorico.MoveFirst
    End If
    '
    Set oDoPerinatalAtencion = Nothing
    Set oPerinatalAtencion = Nothing
    Set oRsTmp = Nothing
    Set oDoPerinatalAtencionCred1 = Nothing
    Set oPerinatalAtencionCred1 = Nothing
    '
     CargaGraficoChartSpace True
   '
   If lbSeCargaDatosDesdeTablasPerinatal = False Then AsignaEQUISaCREDautomaticamente
   '
   lbEstaMaximizadoElGrafico = True
   cmdZoom_Click
End Sub

Public Sub ActualizaGraficoYDiagnosticosAutomaticamente(lnPesoKg As Double, lnTallaCM As Long)
    If lnPesoKg > 0 And lnTallaCM > 0 Then
       CalculaPercentiles lnPesoKg, lnTallaCM
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
       
    End If
End Sub

'Actualiza valores y Devuelve percentil de la ATENCION ACTUAL DEL PACIENTE
Sub CalculaPercentiles(lnPesoKg As Double, lnTallaCM As Long)
    lnPercentilPE = lnPercentilNull: lnPercentilTE = lnPercentilNull: lnPercentilPT = lnPercentilNull: lnPercentilIMC = lnPercentilNull
    lnPercentilPE_Z = lnPercentilNull: lnPercentilTE_Z = lnPercentilNull: lnPercentilPT_Z = lnPercentilNull: lnPercentilIMC_Z = lnPercentilNull
    If lnPesoKg > 0 And lnTallaCM > 0 Then
       On Error Resume Next
       Dim EXL As Excel.Application
       Set EXL = New Excel.Application
       Dim W As Excel.Workbook
       Dim lnEdadUSAmaxima
       lnEdadUSAmaxima = 5
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
           Set W = EXL.Workbooks.Open(App.Path & "\Plantillas\cred.xls")       'usa
       Else
           Set W = EXL.Workbooks.Open(App.Path & "\Plantillas\cred who.xls")    'oms
       End If
       Dim s As Excel.Worksheet
       Dim lnEdadEnMesesMasPuntoCinco As Double, lnMinimo As Double, lnMaximo As Double, lnIMC As Double
       Dim lnTallaEnCmMasPuntoCinco As Double, lnEdadEnSemanasMasPuntoCinco As Double, lnEdadEnDiasMasPuntoCinco As Double
       lnEdadEnSemanasMasPuntoCinco = ml_EdadEnSemanas + 0.5
       lnEdadEnMesesMasPuntoCinco = ml_EdadEnMeses + 0.5
       lnEdadEnDiasMasPuntoCinco = ln_EdadEnDias + 0.5
       lnTallaEnCmMasPuntoCinco = lnTallaCM + 0.5
       
       'Peso Edad
       Set s = W.Sheets("P-E")
       
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(243, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilPE = lnPercentilNull
          s.Cells(246, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(247, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnPesoKg
          lnPercentilPE = s.Cells(255, IIf(ml_idTipoSexo = 1, 3, 19)).Value
          lnPercentilPE_Z = s.Cells(254, IIf(ml_idTipoSexo = 1, 3, 19)).Value
       Else
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(1858, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilPE = lnPercentilNull
          s.Cells(1861, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnEdadEnDiasMasPuntoCinco
          s.Cells(1862, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnPesoKg
          lnPercentilPE = s.Cells(1870, IIf(ml_idTipoSexo = 1, 4, 20)).Value
          lnPercentilPE_Z = s.Cells(1869, IIf(ml_idTipoSexo = 1, 4, 20)).Value
       End If
       'Talla Edad
       Set s = W.Sheets("T-E")
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(243, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilTE = lnPercentilNull
          s.Cells(246, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(247, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnTallaCM
          lnPercentilTE = s.Cells(255, IIf(ml_idTipoSexo = 1, 3, 19)).Value
          lnPercentilTE_Z = s.Cells(254, IIf(ml_idTipoSexo = 1, 3, 19)).Value
       Else
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(1858, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilTE = lnPercentilNull
          s.Cells(1861, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnEdadEnDiasMasPuntoCinco
          s.Cells(1862, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnTallaCM
          lnPercentilTE = s.Cells(1870, IIf(ml_idTipoSexo = 1, 4, 20)).Value
          lnPercentilTE_Z = s.Cells(1869, IIf(ml_idTipoSexo = 1, 4, 20)).Value
       End If
       
       'Peso Talla
       Set s = W.Sheets("P-T")
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(61, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilPT = lnPercentilNull
          s.Cells(64, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnTallaEnCmMasPuntoCinco
          s.Cells(65, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnPesoKg
          lnPercentilPT = s.Cells(73, IIf(ml_idTipoSexo = 1, 3, 19)).Value
          lnPercentilPT_Z = s.Cells(72, IIf(ml_idTipoSexo = 1, 3, 19)).Value
       Else
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(652, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilPT = lnPercentilNull
          s.Cells(655, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnTallaEnCmMasPuntoCinco
          s.Cells(656, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnPesoKg
          lnPercentilPT = s.Cells(664, IIf(ml_idTipoSexo = 1, 4, 20)).Value
          lnPercentilPT_Z = s.Cells(663, IIf(ml_idTipoSexo = 1, 4, 20)).Value
       End If
       'Edad IMC
       lnIMC = Round(lnPesoKg / (lnTallaCM * lnTallaCM * 0.0001), 0)
       Set s = W.Sheets("E-IMC")
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 19)).Value
          lnMaximo = s.Cells(220, IIf(ml_idTipoSexo = 1, 2, 19)).Value
          lnPercentilIMC = lnPercentilNull
          s.Cells(223, IIf(ml_idTipoSexo = 1, 4, 21)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(224, IIf(ml_idTipoSexo = 1, 4, 21)).Value = lnIMC
          lnPercentilIMC = s.Cells(232, IIf(ml_idTipoSexo = 1, 3, 20)).Value
          lnPercentilIMC_Z = s.Cells(231, IIf(ml_idTipoSexo = 1, 3, 20)).Value
       Else
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(1858, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilIMC = lnPercentilNull
          s.Cells(1861, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnEdadEnDiasMasPuntoCinco
          s.Cells(1862, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnIMC
          lnPercentilIMC = s.Cells(1870, IIf(ml_idTipoSexo = 1, 4, 20)).Value
          lnPercentilIMC_Z = s.Cells(1869, IIf(ml_idTipoSexo = 1, 4, 20)).Value
       End If
       W.Close False
       Set s = Nothing
       Set W = Nothing
       Set EXL = Nothing
    End If
End Sub

Sub GraficoRegistraDatosParaFilaColumnas(lbSeCargaPercentilHistoricos As Boolean)
    If lbSeCargaPercentilHistoricos = True Then
       LimpiaPercentil
       Dim oRsTmp As New Recordset
       Dim oRsMesAnio As New Recordset
       Dim lnEdadEnSemanas111 As Long, lnEdadEnAnios111 As Long, lnNuevoMesAnio As Boolean, lbLaAtencionActualEstaEnTemporal As Boolean
       With oRsMesAnio
              .Fields.Append "EdadEnMeses", adInteger
              .Fields.Append "PercentilPE", adInteger
              .Fields.Append "PercentilTE", adInteger
              .Fields.Append "PercentilPT", adInteger
              .Fields.Append "PercentilIMC", adInteger
              .Fields.Append "Veces", adInteger
              .CursorType = adOpenDynamic
              .LockType = adLockOptimistic
              .Open
       End With
       Set oRsTmp = mo_reglasComunes.PerinatalAtencionSeleccionarPorIdPaciente(ml_idPaciente)
       If oRsTmp.RecordCount > 0 Then
          oRsTmp.MoveFirst
          Do While Not oRsTmp.EOF
             'If ml_idAtencion >= oRsTmp.Fields!idAtencion Then
             
             
             
              If ml_FechaAtencion >= oRsTmp!FechaAtencion Then
             
             
                If mo_idPerinatalAtencion = 0 Then
                   txtNcontrol.Text = oRsTmp!CredN + 1
                End If
                lnEdadEnSemanas111 = sighentidades.DevuelveEdadEnSemanas(ld_FechaNacimiento, oRsTmp!FechaAtencion) 'si  se cambia el GRAFICO a: edad en semanas
                lnEdadEnAnios111 = DateDiff("yyyy", ld_FechaNacimiento, oRsTmp!FechaAtencion)
                If oRsTmp.Fields!GrafYpercentilPE > 0 Or oRsTmp.Fields!GrafYpercentilTE > 0 Or oRsTmp.Fields!GrafYpercentilPT > 0 Or oRsTmp.Fields!GrafYimc > 0 Then
                    oRsPercentil.AddNew
                    oRsPercentil.Fields!idAtencion = oRsTmp.Fields!idAtencion
                    If cmbGraficoSM.ListIndex = 0 Then
                       oRsPercentil.Fields!EdadEnMeses = IIf(IsNull(oRsTmp.Fields!GrafXedadEnMeses), 0, oRsTmp.Fields!GrafXedadEnMeses)
                    ElseIf cmbGraficoSM.ListIndex = 1 Then
                       oRsPercentil.Fields!EdadEnMeses = lnEdadEnSemanas111
                    Else
                       oRsPercentil.Fields!EdadEnMeses = lnEdadEnAnios111
                    End If
                
                    oRsPercentil.Fields!PercentilPE = IIf(IsNull(oRsTmp.Fields!GrafYpercentilPE), 0, oRsTmp.Fields!GrafYpercentilPE)
                    oRsPercentil.Fields!PercentilTE = IIf(IsNull(oRsTmp.Fields!GrafYpercentilTE), 0, oRsTmp.Fields!GrafYpercentilTE)
                    oRsPercentil.Fields!PercentilPT = IIf(IsNull(oRsTmp.Fields!GrafYpercentilPT), 0, oRsTmp.Fields!GrafYpercentilPT)
                    oRsPercentil.Fields!PercentilIMC = IIf(IsNull(oRsTmp.Fields!GrafYimc), 0, oRsTmp.Fields!GrafYimc)
                    oRsPercentil.Update
                    'solo si se elige MES o AÑO
                    If cmbGraficoSM.ListIndex <> 1 Then
                       lnNuevoMesAnio = True
                       If oRsMesAnio.RecordCount > 0 Then
                          oRsMesAnio.MoveFirst
                          oRsMesAnio.Find "EdadEnMeses=" & oRsPercentil!EdadEnMeses
                          If Not oRsMesAnio.EOF Then
                             lnNuevoMesAnio = False
                          End If
                       End If
                       If lnNuevoMesAnio = True Then
                            oRsMesAnio.AddNew
                            oRsMesAnio!EdadEnMeses = oRsPercentil!EdadEnMeses
                            oRsMesAnio!PercentilPE = oRsPercentil.Fields!PercentilPE
                            oRsMesAnio!PercentilTE = oRsPercentil.Fields!PercentilTE
                            oRsMesAnio!PercentilPT = oRsPercentil.Fields!PercentilPT
                            oRsMesAnio!PercentilIMC = oRsPercentil.Fields!PercentilIMC
                            oRsMesAnio!veces = 1
                       Else
                            oRsMesAnio!PercentilPE = oRsPercentil.Fields!PercentilPE
                            oRsMesAnio!PercentilTE = oRsPercentil.Fields!PercentilTE
                            oRsMesAnio!PercentilPT = oRsPercentil.Fields!PercentilPT
                            oRsMesAnio!PercentilIMC = oRsPercentil.Fields!PercentilIMC
                            oRsMesAnio!veces = oRsMesAnio!veces + 1
                       End If
                       oRsMesAnio.Update
                   End If
                   
                End If
                '
             End If
             oRsTmp.MoveNext
          Loop
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
    End If
    lbLaAtencionActualEstaEnTemporal = True
    If lnPercentilPE > 0 Or lnPercentilTE > 0 Or lnPercentilPT > 0 And lnPercentilIMC > 0 Then
        If oRsPercentil.RecordCount > 0 Then
           oRsPercentil.MoveFirst
           oRsPercentil.Find "idAtencion=" & ml_idAtencion
           
        End If
        If oRsPercentil.EOF Then
           lbLaAtencionActualEstaEnTemporal = False
           lnEdadEnSemanas111 = ml_EdadEnSemanas   'si  se cambia el GRAFICO a: edad en semanas
           lnEdadEnAnios111 = ml_EdadEnAnios       'si  se cambia el GRAFICO a: edad en años
           oRsPercentil.AddNew
           oRsPercentil.Fields!idAtencion = ml_idAtencion
           If cmbGraficoSM.ListIndex = 0 Then
              oRsPercentil.Fields!EdadEnMeses = ml_EdadEnMeses
           ElseIf cmbGraficoSM.ListIndex = 1 Then
              oRsPercentil.Fields!EdadEnMeses = lnEdadEnSemanas111
           Else
              oRsPercentil.Fields!EdadEnMeses = lnEdadEnAnios111
           End If
        End If
        
        oRsPercentil.Fields!PercentilPE = lnPercentilPE
        oRsPercentil.Fields!PercentilTE = lnPercentilTE
        oRsPercentil.Fields!PercentilPT = lnPercentilPT
        oRsPercentil.Fields!PercentilIMC = lnPercentilIMC
        oRsPercentil.Update
    End If
    'solo si se elige MES o AÑO
    If cmbGraficoSM.ListIndex <> 1 And lbSeCargaPercentilHistoricos = True Then
       If lbLaAtencionActualEstaEnTemporal = False Then
            lnNuevoMesAnio = True
            If oRsMesAnio.RecordCount > 0 Then
               oRsMesAnio.MoveFirst
               oRsMesAnio.Find "EdadEnMeses=" & oRsPercentil!EdadEnMeses
               If Not oRsMesAnio.EOF Then
                  lnNuevoMesAnio = False
               End If
            End If
            If lnNuevoMesAnio = True Then
                 oRsMesAnio.AddNew
                 oRsMesAnio!EdadEnMeses = oRsPercentil!EdadEnMeses
                 oRsMesAnio!PercentilPE = oRsPercentil.Fields!PercentilPE
                 oRsMesAnio!PercentilTE = oRsPercentil.Fields!PercentilTE
                 oRsMesAnio!PercentilPT = oRsPercentil.Fields!PercentilPT
                 oRsMesAnio!PercentilIMC = oRsPercentil.Fields!PercentilIMC
                 oRsMesAnio!veces = 1
            Else
                 oRsMesAnio!PercentilPE = oRsPercentil.Fields!PercentilPE
                 oRsMesAnio!PercentilTE = oRsPercentil.Fields!PercentilTE
                 oRsMesAnio!PercentilPT = oRsPercentil.Fields!PercentilPT
                 oRsMesAnio!PercentilIMC = oRsPercentil.Fields!PercentilIMC
                 oRsMesAnio!veces = oRsMesAnio!veces + 1
            End If
            oRsMesAnio.Update
       End If
       If oRsMesAnio.State = 1 Then
            'queda el promedio por cada MES o AÑO
            oRsMesAnio.Filter = "veces>1"
            If oRsMesAnio.RecordCount > 0 Then
               oRsMesAnio.MoveFirst
               Do While Not oRsMesAnio.EOF
                  oRsPercentil.Filter = "edadEnMeses=" & oRsMesAnio!EdadEnMeses
                  oRsPercentil.MoveFirst
                  oRsPercentil.Fields!PercentilPE = oRsMesAnio!PercentilPE
                  oRsPercentil.Fields!PercentilTE = oRsMesAnio!PercentilTE
                  oRsPercentil.Fields!PercentilPT = oRsMesAnio!PercentilPT
                  oRsPercentil.Fields!PercentilIMC = oRsMesAnio!PercentilIMC
                  oRsPercentil.Update
                  oRsPercentil.MoveNext
                  Do While Not oRsPercentil.EOF
                     oRsPercentil.Delete
                     oRsPercentil.Update
                     oRsPercentil.MoveNext
                  Loop
                  oRsMesAnio.MoveNext
               Loop
               oRsPercentil.Filter = ""
            End If
       End If
    End If
    '
    Set oRsMesAnio = Nothing
End Sub







Sub CargaGraficoChartSpace(lbActualizaDesdeInicioRs As Boolean)
    Dim lnRegistrosPerc As Integer
    lnRegistrosPerc = oRsPercentil.RecordCount
    If lnRegistrosPerc = 0 Then
        ChartSpace1.Clear
        Exit Sub
    End If
    Dim lnFor As Integer
    If lbActualizaDesdeInicioRs = True Or lnNroPuntosGraficos >= 1 Then
            xValues = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250)
            yValuesPT = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250)
            yValuesTE = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250)
            yValuesPE = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250)
            yValuesIMC = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250)
            If lnRegistrosPerc > 15 Then
                ReDim yValuesPT(lnRegistrosPerc - 1)
                ReDim yValuesTE(lnRegistrosPerc - 1)
                ReDim yValuesPE(lnRegistrosPerc - 1)
                ReDim yValuesIMC(lnRegistrosPerc - 1)
            End If
            ReDim xValues(lnRegistrosPerc - 1)
            lnNroPuntosGraficos = lnRegistrosPerc - 1
            oRsPercentil.MoveLast
            For lnFor = (lnRegistrosPerc - 1) To 0 Step -1
               xValues(lnFor) = oRsPercentil.Fields!EdadEnMeses
               yValuesPT(lnFor) = oRsPercentil.Fields!PercentilPT
               yValuesTE(lnFor) = oRsPercentil.Fields!PercentilTE
               yValuesPE(lnFor) = oRsPercentil.Fields!PercentilPE
               yValuesIMC(lnFor) = oRsPercentil.Fields!PercentilIMC
               oRsPercentil.MovePrevious
            Next
    Else
            xValues = Array(10)
            yValuesPT = Array(10)
            yValuesTE = Array(10)
            yValuesPE = Array(10)
            yValuesIMC = Array(10)
            lnFor = lnNroPuntosGraficos
            If lnFor < 0 Then
               lnFor = 0
            End If
            yValuesPT(lnFor) = lnPercentilPT
            yValuesTE(lnFor) = lnPercentilTE
            yValuesPE(lnFor) = lnPercentilPE
            yValuesIMC(lnFor) = lnPercentilIMC
    End If
    '
    ChartSpace1.Clear
    ChartSpace1.DisplayToolbar = False
    Set owcChart = ChartSpace1.Charts.Add
    owcChart.HasTitle = True
    owcChart.Title.Caption = "                    Edad"
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
    If chkImc.Value = 1 Then
        Set owcSeries = owcChart.SeriesCollection.Add
        With owcSeries
            .Caption = "idTipoFinanciamiento"
            .SetData chDimCategories, chDataLiteral, xValues
            .SetData chDimValues, chDataLiteral, yValuesIMC
            .Type = chChartTypeLineMarkers
            .Line.Color = vbRed
            .Line.Weight = 3
            .Marker.Style = chMarkerStyleCircle
            .Line.DashStyle = chLineSolid
            .DataLabelsCollection.Add
        End With
    End If
    '
    If chkTe.Value = 1 Then
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
    End If
    '
    If chkPe.Value = 1 Then
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
    End If
    '
    If chkPT.Value = 1 Then
        Set owcSeries = owcChart.SeriesCollection.Add
        With owcSeries
            .Caption = "idTipoFinanciamiento"
            .SetData chDimCategories, chDataLiteral, xValues
            .SetData chDimValues, chDataLiteral, yValuesPT
            .Type = chChartTypeLineMarkers
            .Line.Color = vbCyan
            .Line.Weight = 3
            .Marker.Style = chMarkerStyleCircle
            .Line.DashStyle = chLineSolid
            .DataLabelsCollection.Add
        End With
    End If


End Sub



Sub HabilitaDeshabilita(lcFrame As String, lbEstado As Boolean)
    Select Case lcFrame
    Case "Inmunizaciones"
         cmbEligeInmunizacion.Enabled = lbEstado
         btnQuitarInmunizacion.Enabled = lbEstado
         'FraInmunizaciones.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         mo_Formulario.HabilitarDeshabilitar FraInmunizaciones, lbEstado
         grdInmunizaciones.Appearance.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
    Case "OtrosCpt"
         cmbProcedimientosFrecuentes.Enabled = lbEstado
         btnQuitaOtrosProcedimientos.Enabled = lbEstado
         'FraOtrosCpt.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         mo_Formulario.HabilitarDeshabilitar FraOtrosCpt, lbEstado
         grdCptFrecuentes.Appearance.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
    Case "DxDesarrollo"
         cmbDxDesarrollo.Enabled = lbEstado
         btnQuitaDxDesarrollo.Enabled = lbEstado
         'FraDxDesarrollo.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
         mo_Formulario.HabilitarDeshabilitar FraDxDesarrollo, lbEstado
         grdMorbilidadDesarollo.Appearance.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
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
         grdMedicamentos.Appearance.ForeColor = IIf(lbEstado = True, vbBlack, vbRed)
    End Select
End Sub

'Solamente carga Dx automaticos si esta vacio (es decir no se eligió nada aun)
Sub CargaDxAutomaticosParaMorbilidadEnDesarrollo(lnPesoKg As Double, lnTallaCM As Long)
    lnPesoKgActual = lnPesoKg
    lnTallaCMActual = lnTallaCM
    If lbSeCargaDatosDesdeTablasPerinatal = False Then AsignaEQUISaCREDautomaticamente
    '
    If oRsDxDesarrolloAutomaticos.RecordCount > 0 And oRsDxDesarrollo.RecordCount = 0 And lnPesoKg > 0 And lnTallaCM > 0 And FraDxDesarrollo.Enabled = True Then
       Dim lbContinuar As Boolean
       Dim lcLab As String
       oRsDxDesarrolloAutomaticos.MoveFirst
       Do While Not oRsDxDesarrolloAutomaticos.EOF
            lcLab = ""
            Select Case lnIdModulo
            Case sighHasta28Dias
                 'Usa el PESO
                 If (lnPesoKg * 1000) >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And (lnPesoKg * 1000) <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                    lcLab = labPE
                    oRsDxDesarrollo.AddNew
                    oRsDxDesarrollo.Fields!Id = oRsDxDesarrolloAutomaticos.Fields!IdDiagnostico
                    oRsDxDesarrollo.Fields!DIAGNOSTICO = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
                    oRsDxDesarrollo.Fields!idAtencion = ml_idAtencion
                    oRsDxDesarrollo.Fields!IdClasificacionDx = sghTiposDiagnostico.sghAtencionConsultaExterna
                    oRsDxDesarrollo.Fields!IdSubclasificacionDx = 2  'sghDxDefinitivos.sighDxCeDefinitivo
                    If lcLab <> "" Then
                       oRsDxDesarrollo.Fields!labConfHIS = lcLab
                    End If
                    oRsDxDesarrollo.Update
                 End If
            Case sighDesde29diasHasta1anio
                 'Usa Z de Peso-Edad,Peso-Talla,Talla-Edad, Cie10 de mas de 4 digitos
                 lbContinuar = False
                 Select Case UCase(oRsDxDesarrolloAutomaticos.Fields!cie10his)
                 Case "E343", "E3431", "E3441", "E344"
                     If lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                        lcLab = labTE
                     End If
                 Case "E660", "E669"
                     If (lnPercentilPE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilPE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                        lcLab = labPE
                     End If
'                     If (lnPercentilPT_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilPT_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
'                        lbContinuar = True
'                        lcLab = labPT
'                     End If
                 Case Else
'                     If (lnPercentilPT_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilPT_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
'                         lbContinuar = True
'                         lcLab = labPT
'                     End If
                     If (lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                          lbContinuar = True
                          lcLab = labTE
                     End If
                     If (lnPercentilPE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilPE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                        lcLab = labPE
                     End If
                 End Select
                 If lbContinuar = True Then
                    oRsDxDesarrollo.AddNew
                    oRsDxDesarrollo.Fields!Id = oRsDxDesarrolloAutomaticos.Fields!IdDiagnostico
                    oRsDxDesarrollo.Fields!DIAGNOSTICO = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
                    oRsDxDesarrollo.Fields!idAtencion = ml_idAtencion
                    oRsDxDesarrollo.Fields!IdClasificacionDx = sghTiposDiagnostico.sghAtencionConsultaExterna
                    oRsDxDesarrollo.Fields!IdSubclasificacionDx = 2   'sghDxDefinitivos.sighDxCeDefinitivo
                    If lcLab <> "" Then
                       oRsDxDesarrollo.Fields!labConfHIS = lcLab
                    End If
                    oRsDxDesarrollo.Update
                 End If
            Case sighDesde1Hasta4anios
                 'Usa Z de Peso-Edad,Peso-Talla,Talla-Edad, Cie10 de mas de 4 digitos
                 lbContinuar = False
                 Select Case UCase(oRsDxDesarrolloAutomaticos.Fields!cie10his)
                 Case "E343", "E3431"
                     If lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                        lcLab = labTE
                     End If
                 Case Else
'                     If (lnPercentilPT_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilPT_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
'                            lbContinuar = True
'                            lcLab = labPT
'                     End If
                     If (lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                            lbContinuar = True
                            lcLab = labTE
                     End If
                     If (lnPercentilPE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilPE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                            lbContinuar = True
                            lcLab = labPE
                     End If
                 End Select
                 If lbContinuar = True Then
                    oRsDxDesarrollo.AddNew
                    oRsDxDesarrollo.Fields!Id = oRsDxDesarrolloAutomaticos.Fields!IdDiagnostico
                    oRsDxDesarrollo.Fields!DIAGNOSTICO = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
                    oRsDxDesarrollo.Fields!idAtencion = ml_idAtencion
                    oRsDxDesarrollo.Fields!IdClasificacionDx = sghTiposDiagnostico.sghAtencionConsultaExterna
                    oRsDxDesarrollo.Fields!IdSubclasificacionDx = 2     'sghDxDefinitivos.sighDxCeDefinitivo
                    If lcLab <> "" Then
                       oRsDxDesarrollo.Fields!labConfHIS = lcLab
                    End If
                    oRsDxDesarrollo.Update
                 End If
            Case sighDesde5Hasta9anios
                 'Usa Z de IMC,Talla-Edad, Cie10 de mas de 4 digitos
                 lbContinuar = False
                 Select Case UCase(oRsDxDesarrolloAutomaticos.Fields!cie10his)
                 Case "Z006"
                     If (lnPercentilTE_Z >= 10 And lnPercentilTE_Z <= 89.999) Then
                        lbContinuar = True
                        lcLab = labTE
                     End If
                     If (lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                        lcLab = labIMC
                     End If
                 Case "E46X", "E660", "E669"
                     If lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                        lcLab = labIMC
                     End If
                 Case "E3431", "E3441", "E344"
                     If lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                        lcLab = labTE
                     End If
                 Case Else
                     If (lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                        lcLab = labTE
                     End If
                     If (lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                        lcLab = labIMC
                     End If
                 End Select
                 If lbContinuar = True Then
                    oRsDxDesarrollo.AddNew
                    oRsDxDesarrollo.Fields!Id = oRsDxDesarrolloAutomaticos.Fields!IdDiagnostico
                    oRsDxDesarrollo.Fields!DIAGNOSTICO = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
                    oRsDxDesarrollo.Fields!idAtencion = ml_idAtencion
                    oRsDxDesarrollo.Fields!IdClasificacionDx = sghTiposDiagnostico.sghAtencionConsultaExterna
                    oRsDxDesarrollo.Fields!IdSubclasificacionDx = 2   'sghDxDefinitivos.sighDxCeDefinitivo
                    If lcLab <> "" Then
                       oRsDxDesarrollo.Fields!labConfHIS = lcLab
                    End If
                    oRsDxDesarrollo.Update
                 End If
            Case sighDesde10Hasta11anios
                 'Usa Z de IMC,Talla-Edad, Cie10 de mas de 4 digitos
                 lbContinuar = False
                 Select Case UCase(oRsDxDesarrolloAutomaticos.Fields!cie10his)
                 Case "Z006"
                     If (lnPercentilTE_Z >= 10 And lnPercentilTE_Z <= 89.999) Then
                        lbContinuar = True
                        lcLab = labTE
                     End If
                     If (lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                        lcLab = labIMC
                     End If
                 Case "E46X", "E660", "E669"
                     If lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                        lcLab = labIMC
                     End If
                 Case "E3431", "E3441", "E344"
                     If lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                        lcLab = labTE
                     End If
                 Case Else
                     If (lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                        lcLab = labTE
                     End If
                     If (lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                        lcLab = labIMC
                     End If
                 End Select
                 If lbContinuar = True Then
                    oRsDxDesarrollo.AddNew
                    oRsDxDesarrollo.Fields!Id = oRsDxDesarrolloAutomaticos.Fields!IdDiagnostico
                    oRsDxDesarrollo.Fields!DIAGNOSTICO = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
                    oRsDxDesarrollo.Fields!idAtencion = ml_idAtencion
                    oRsDxDesarrollo.Fields!IdClasificacionDx = sghTiposDiagnostico.sghAtencionConsultaExterna
                    oRsDxDesarrollo.Fields!IdSubclasificacionDx = 2      'sghDxDefinitivos.sighDxCeDefinitivo
                    If lcLab <> "" Then
                       oRsDxDesarrollo.Fields!labConfHIS = lcLab
                    End If
                    oRsDxDesarrollo.Update
                 End If
            Case sighDesde12Hasta17anios
                 'Usa Z de IMC,Talla-Edad, Cie10 de mas de 4 digitos
                 lbContinuar = False
                 Select Case UCase(oRsDxDesarrolloAutomaticos.Fields!cie10his)
                 Case "Z006"
                     If (lnPercentilTE_Z >= 10 And lnPercentilTE_Z <= 89.999) Then
                        lbContinuar = True
                        lcLab = labTE
                     End If
                     If (lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                        lcLab = labIMC
                     End If
                 Case "E46X", "E660", "E669"
                     If lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                        lcLab = labIMC
                     End If
                 Case "E3431", "E3441", "E344"
                     If lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal Then
                        lbContinuar = True
                        lcLab = labTE
                     End If
                 Case Else
                     If (lnPercentilTE_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilTE_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                        lcLab = labTE
                     End If
                     If (lnPercentilIMC_Z >= oRsDxDesarrolloAutomaticos.Fields!rangoInicio And lnPercentilIMC_Z <= oRsDxDesarrolloAutomaticos.Fields!rangoFinal) Then
                        lbContinuar = True
                        lcLab = labIMC
                     End If
                 End Select
                 If lbContinuar = True Then
                    oRsDxDesarrollo.AddNew
                    oRsDxDesarrollo.Fields!Id = oRsDxDesarrolloAutomaticos.Fields!IdDiagnostico
                    oRsDxDesarrollo.Fields!DIAGNOSTICO = oRsDxDesarrolloAutomaticos.Fields!DIAGNOSTICO
                    oRsDxDesarrollo.Fields!idAtencion = ml_idAtencion
                    oRsDxDesarrollo.Fields!IdClasificacionDx = sghTiposDiagnostico.sghAtencionConsultaExterna
                    oRsDxDesarrollo.Fields!IdSubclasificacionDx = 2     'sghDxDefinitivos.sighDxCeDefinitivo
                    If lcLab <> "" Then
                       oRsDxDesarrollo.Fields!labConfHIS = lcLab
                    End If
                    oRsDxDesarrollo.Update
                 End If
            Case sighDesde18anios
            End Select
            oRsDxDesarrolloAutomaticos.MoveNext
       Loop
       If oRsDxDesarrollo.RecordCount > 0 Then
          oRsDxDesarrollo.MoveFirst
       End If
    End If
End Sub


'preguntas:
'- si son 2 servicios donde se consulta el paciente. ejm Cred y Medicina, el Dx de medicina debe copiarse como dx en Cred ?
'- terminar de registrar CPT y DX que no se ha encontrado en tablas (od=otra descripcion, ya está registrado aquí como 'consulta en consultorios externos')



Public Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   '
'   Dim oRsTmp1 As New Recordset
'   Set oRsTmp1 = DevuelveDatosCred(True)
'   oRsTmp1.Filter = "credCheck='X'"
'   If oRsTmp1.RecordCount = 0 And lnPesoKgActual > 0 And lnTallaCMActual > 0 Then
'        Set oRsTmp1 = Nothing
'        MsgBox "Debe marcar el " & FraCred.Caption & " correspondiente", vbInformation, "MODULO PERINATAL"
'        Exit Function
'   'ElseIf oRsTmp1.RecordCount > 1 Then
'       ' Set oRsTmp1 = Nothing
'       ' MsgBox "Solo debe marcar una sola vez    X    en " & FraCred.Caption, vbInformation, "MODULO PERINATAL"
'       ' Exit Function
'   End If
   If lnPesoKgActual > 0 And lnTallaCMActual > 0 Then
      AsignaEQUISaCREDautomaticamente
   End If
   '
   If oRsDxDesarrollo.RecordCount = 0 And oRsMorbilidadFrec.RecordCount = 0 Then
        'Set oRsTmp1 = Nothing
        MsgBox "Debe registrar un Diagnóstico_Desarrollo o Morbilidad_Frecuente", vbInformation, "MODULO PERINATAL"
        Exit Function
   End If
   '
   ValidarDatosObligatorios = True
   'Set oRsTmp1 = Nothing
End Function

Sub AsignaEQUISaCREDautomaticamente()
    If lnPesoKgActual > 0 And lnTallaCMActual > 0 Then
       If Cred1.Locked = False And Cred1.Visible = True Then
           If Cred1.Text = "" Then Cred1.Text = "X"
       ElseIf Cred2.Locked = False And Cred2.Visible = True Then
           If Cred2.Text = "" Then Cred2.Text = "X"
       ElseIf Cred3.Locked = False And Cred3.Visible = True Then
           If Cred3.Text = "" Then Cred3.Text = "X"
       ElseIf Cred4.Locked = False And Cred4.Visible = True Then
           If Cred4.Text = "" Then Cred4.Text = "X"
       ElseIf Cred5.Locked = False And Cred5.Visible = True Then
           If Cred5.Text = "" Then Cred5.Text = "X"
       ElseIf Cred6.Locked = False And Cred6.Visible = True Then
           If Cred6.Text = "" Then Cred6.Text = "X"
       ElseIf Cred7.Locked = False And Cred7.Visible = True Then
           If Cred7.Text = "" Then Cred7.Text = "X"
       ElseIf Cred8.Locked = False And Cred8.Visible = True Then
           If Cred8.Text = "" Then Cred8.Text = "X"
       ElseIf Cred9.Locked = False And Cred9.Visible = True Then
           If Cred9.Text = "" Then Cred9.Text = "X"
       ElseIf Cred10.Locked = False And Cred10.Visible = True Then
           If Cred10.Text = "" Then Cred10.Text = "X"
       ElseIf Cred11.Locked = False And Cred11.Visible = True Then
           If Cred11.Text = "" Then Cred11.Text = "X"
       ElseIf Cred12.Locked = False And Cred12.Visible = True Then
           If Cred12.Text = "" Then Cred12.Text = "X"
       End If
    End If
End Sub

