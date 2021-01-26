VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form AdmisionHospDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15750
   ControlBox      =   0   'False
   Icon            =   "AdmisionHospDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   15750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTriaje 
      Caption         =   "Triaje de ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Left            =   12120
      TabIndex        =   169
      Top             =   1230
      Width           =   3585
      Begin SISGalenPlus.ucTriajeVisor ucTriajeVisorCE 
         Height          =   3315
         Left            =   75
         TabIndex        =   170
         Top             =   285
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   5847
      End
   End
   Begin VB.Frame Frame 
      Height          =   675
      Index           =   0
      Left            =   9255
      TabIndex        =   106
      Top             =   -15
      Width           =   2775
      Begin VB.CommandButton btnBuscarPaciente 
         Height          =   315
         Left            =   75
         Picture         =   "AdmisionHospDetalle.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   210
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   1455
         Picture         =   "AdmisionHospDetalle.frx":3913
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   195
         Width           =   1305
      End
   End
   Begin VB.Frame FraProviene 
      Caption         =   "Procede de:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9255
      TabIndex        =   103
      Top             =   705
      Width           =   2760
      Begin VB.CommandButton btnProvCE 
         Caption         =   "ConsExt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1530
         TabIndex        =   105
         Top             =   210
         Width           =   1065
      End
      Begin VB.CommandButton btnProvEmergencia 
         Caption         =   "Emergencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   104
         Top             =   210
         Width           =   1065
      End
   End
   Begin SISGalenPlus.UcSISafiliacion UcSISafiliacion1 
      Height          =   615
      Left            =   5250
      TabIndex        =   5
      Top             =   210
      Visible         =   0   'False
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   1085
   End
   Begin UltraGrid.SSUltraGrid grdPacientesEncontrados 
      Height          =   255
      Left            =   105
      TabIndex        =   10
      Top             =   1110
      Visible         =   0   'False
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   450
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
      Caption         =   "Lista de pacientes encontrados"
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
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
      Left            =   20
      TabIndex        =   71
      Top             =   -15
      Width           =   9225
      Begin VB.CommandButton cmdSinApellidoMaterno 
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
         Left            =   3675
         Picture         =   "AdmisionHospDetalle.frx":3F3C
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   450
         Width           =   255
      End
      Begin VB.CommandButton cmdSinApellidoPaterno 
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
         Left            =   2700
         Picture         =   "AdmisionHospDetalle.frx":44C6
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   450
         Width           =   255
      End
      Begin VB.CheckBox chkMuestraHistorial 
         Alignment       =   1  'Right Justify
         Caption         =   "Muestra HISTORIAL al buscar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   165
         TabIndex        =   146
         Top             =   795
         Width           =   2505
      End
      Begin VB.Frame fraPacienteNuevo 
         Height          =   450
         Left            =   5235
         TabIndex        =   109
         Top             =   765
         Width           =   3900
         Begin VB.CheckBox chkPacienteNuevo 
            Caption         =   "Paciente &nuevo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   150
            TabIndex        =   111
            Top             =   135
            Width           =   1440
         End
         Begin VB.CheckBox chkBuscarEnSIS 
            Caption         =   "Buscar en SIS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2430
            TabIndex        =   110
            Top             =   135
            Visible         =   0   'False
            Width           =   1365
         End
      End
      Begin VB.TextBox txtNroHistoriaBusqueda 
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
         Left            =   960
         MaxLength       =   9
         TabIndex        =   12
         Top             =   465
         Width           =   945
      End
      Begin VB.TextBox txtApellidoPaternoBusqueda 
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
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   1
         Top             =   465
         Width           =   780
      End
      Begin VB.TextBox txtNroDNIBusqueda 
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
         Left            =   150
         TabIndex        =   11
         Top             =   465
         Width           =   825
      End
      Begin VB.TextBox txtSegundoNombreBusqueda 
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
         Left            =   4560
         MaxLength       =   40
         TabIndex        =   4
         Top             =   465
         Width           =   690
      End
      Begin VB.TextBox txtApellidoMaternoBusqueda 
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
         MaxLength       =   40
         TabIndex        =   2
         Top             =   450
         Width           =   705
      End
      Begin VB.TextBox txtPrimerNombreBusqueda 
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
         Left            =   3960
         MaxLength       =   40
         TabIndex        =   3
         Top             =   465
         Width           =   570
      End
      Begin VB.Label Label50 
         Caption         =   "DNI            Nº Historia    Apelli.Paterno ApelliMaterno 1rNom   2oNom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   72
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame Frame4 
      Height          =   885
      Left            =   -15
      TabIndex        =   56
      Top             =   8595
      Width           =   12045
      Begin VB.CommandButton cmdFiliacionCE 
         Caption         =   "Hoja Filiación Consultorio"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   9495
         Picture         =   "AdmisionHospDetalle.frx":4A50
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   150
         Width           =   1245
      End
      Begin VB.CommandButton btnBuscaHistoricos 
         Caption         =   "Históricos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   10770
         Picture         =   "AdmisionHospDetalle.frx":4F29
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   135
         Width           =   1155
      End
      Begin VB.Frame Frame 
         Height          =   315
         Index           =   1
         Left            =   9300
         TabIndex        =   135
         Top             =   465
         Visible         =   0   'False
         Width           =   525
         Begin VB.CommandButton btnQuitarMadre 
            DisabledPicture =   "AdmisionHospDetalle.frx":54B3
            DownPicture     =   "AdmisionHospDetalle.frx":583E
            Height          =   315
            Left            =   4950
            Picture         =   "AdmisionHospDetalle.frx":5BD1
            Style           =   1  'Graphical
            TabIndex        =   138
            Top             =   150
            Width           =   615
         End
         Begin VB.TextBox lblMadre 
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
            Left            =   1860
            TabIndex        =   137
            Top             =   135
            Width           =   3075
         End
         Begin VB.CommandButton cmdBuscaMadre 
            Caption         =   "..."
            Height          =   315
            Left            =   1620
            TabIndex        =   136
            TabStop         =   0   'False
            ToolTipText     =   "Busca a la Madre"
            Top             =   135
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "Madre (Apell.Nom)"
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
            Left            =   240
            TabIndex        =   139
            Top             =   285
            Width           =   1575
         End
      End
      Begin VB.CommandButton btnNuevoAdmisionHospDetalle 
         Caption         =   "Nuevo Reg."
         DisabledPicture =   "AdmisionHospDetalle.frx":5F62
         DownPicture     =   "AdmisionHospDetalle.frx":634B
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3720
         Picture         =   "AdmisionHospDetalle.frx":6757
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   150
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton btnImprimeFichaSIS 
         Caption         =   "FUA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   7320
         Picture         =   "AdmisionHospDetalle.frx":6B63
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   150
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton btnImprimeFiliacion 
         Caption         =   "Filiación Arch.Clínico"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2400
         Picture         =   "AdmisionHospDetalle.frx":703C
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   150
         Width           =   1185
      End
      Begin VB.CommandButton btnPreCuenta 
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1320
         Picture         =   "AdmisionHospDetalle.frx":7515
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   150
         Width           =   1065
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar"
         DisabledPicture =   "AdmisionHospDetalle.frx":79EE
         DownPicture     =   "AdmisionHospDetalle.frx":7EB2
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   6120
         Picture         =   "AdmisionHospDetalle.frx":839E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   150
         Width           =   1185
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AdmisionHospDetalle.frx":888A
         DownPicture     =   "AdmisionHospDetalle.frx":8CEA
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4920
         Picture         =   "AdmisionHospDetalle.frx":915F
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   150
         Width           =   1185
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Filiación (F3)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   75
         Picture         =   "AdmisionHospDetalle.frx":95D4
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   150
         Width           =   1185
      End
      Begin SISGalenPlus.UcEpisodioClinico UcEpisodioClinico1 
         Height          =   450
         Left            =   8505
         TabIndex        =   102
         Top             =   180
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   6059
         _ExtentY        =   1032
      End
   End
   Begin TabDlg.SSTab tabAdmision 
      Height          =   7275
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   12832
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "1. Datos del paciente (F10)"
      TabPicture(0)   =   "AdmisionHospDetalle.frx":9AAD
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ucPacientesDetalle1"
      Tab(0).Control(1)=   "ucMensajeParpadeando1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "2. Ingreso (F11)"
      TabPicture(1)   =   "AdmisionHospDetalle.frx":9AC9
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "TabIngreso"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin SISGalenPlus.ucPacientesDetalle ucPacientesDetalle1 
         Height          =   6465
         Left            =   -74880
         TabIndex        =   6
         Top             =   420
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   11404
      End
      Begin TabDlg.SSTab TabIngreso 
         Height          =   6840
         Left            =   90
         TabIndex        =   57
         Top             =   405
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   12065
         _Version        =   393216
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   520
         ForeColor       =   13653559
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "2.1. Ingreso"
         TabPicture(0)   =   "AdmisionHospDetalle.frx":9AE5
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "ucDiagnosticosIngreso"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "fraDatosReferenciaOrigen"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame7"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "fraNotas"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "2.2. Transferencias"
         TabPicture(1)   =   "AdmisionHospDetalle.frx":9B01
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraServicioActual"
         Tab(1).Control(1)=   "Frame9"
         Tab(1).Control(2)=   "ucTransferenciasDetalle1"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "2.3. Causas externas morbilidad"
         TabPicture(2)   =   "AdmisionHospDetalle.frx":9B1D
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame6"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "2.4. Recetas/Cpt"
         TabPicture(3)   =   "AdmisionHospDetalle.frx":9B39
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "UcPacientesSunasa1"
         Tab(3).Control(1)=   "SSCheck1"
         Tab(3).Control(2)=   "fraCPT"
         Tab(3).Control(3)=   "Frame1"
         Tab(3).Control(4)=   "Enfermeras"
         Tab(3).ControlCount=   5
         TabCaption(4)   =   "2.5. Nacimientos"
         TabPicture(4)   =   "AdmisionHospDetalle.frx":9B55
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "ucDiagnosticoNacimiento"
         Tab(4).Control(1)=   "ucNacimientoDetalle1"
         Tab(4).Control(2)=   "Frame5"
         Tab(4).ControlCount=   3
         TabCaption(5)   =   "2.6. Tratamiento"
         TabPicture(5)   =   "AdmisionHospDetalle.frx":9B71
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame10"
         Tab(5).ControlCount=   1
         Begin VB.Frame Frame10 
            Caption         =   "Indicaciones del Tratamiento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5085
            Left            =   -74865
            TabIndex        =   157
            Top             =   420
            Width           =   5775
            Begin VB.TextBox TxtCitaTratamiento 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4665
               Left            =   240
               MaxLength       =   1000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   158
               Top             =   300
               Width           =   5415
            End
         End
         Begin VB.Frame Frame5 
            Height          =   615
            Left            =   -74910
            TabIndex        =   129
            Top             =   435
            Width           =   11475
            Begin VB.CommandButton btnMedicoRespNacimiento 
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
               Left            =   3210
               Picture         =   "AdmisionHospDetalle.frx":9B8D
               Style           =   1  'Graphical
               TabIndex        =   165
               Top             =   180
               Width           =   360
            End
            Begin VB.TextBox txtIdMedicoNacimiento 
               Height          =   315
               Left            =   2220
               MaxLength       =   10
               TabIndex        =   131
               Top             =   195
               Width           =   945
            End
            Begin VB.TextBox lblNombreMedicoNacimiento 
               Height          =   315
               Left            =   3585
               TabIndex        =   130
               TabStop         =   0   'False
               Top             =   195
               Width           =   7755
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Medico responsable"
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
               TabIndex        =   132
               Top             =   240
               Width           =   1590
            End
         End
         Begin VB.Frame fraServicioActual 
            Caption         =   "Servicio Actual Transferido"
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
            Height          =   615
            Left            =   -74865
            TabIndex        =   123
            Top             =   5490
            Width           =   11535
            Begin VB.CommandButton cmdCamaTransf 
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
               Left            =   2460
               Picture         =   "AdmisionHospDetalle.frx":A117
               Style           =   1  'Graphical
               TabIndex        =   164
               Top             =   225
               Width           =   360
            End
            Begin VB.TextBox txtCamaTransf 
               BackColor       =   &H00FFFFFF&
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
               TabIndex        =   126
               Top             =   240
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.CheckBox chkLlegoSS 
               Alignment       =   1  'Right Justify
               Caption         =   "¿Llegó al 'Servicio Transferido'?"
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
               Left            =   8400
               TabIndex        =   125
               Top             =   240
               Visible         =   0   'False
               Width           =   2955
            End
            Begin VB.TextBox txtServicioTransf 
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
               Left            =   3960
               TabIndex        =   124
               TabStop         =   0   'False
               Top             =   240
               Width           =   4245
            End
            Begin VB.Label lblCamaTransf 
               Alignment       =   2  'Center
               Caption         =   "Código cama"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   128
               Top             =   270
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.Label lblServicioTransf 
               Alignment       =   2  'Center
               Caption         =   "Servicio"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3120
               TabIndex        =   127
               Top             =   270
               Width           =   855
            End
         End
         Begin VB.CommandButton Enfermeras 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   -74910
            Picture         =   "AdmisionHospDetalle.frx":A6A1
            Style           =   1  'Graphical
            TabIndex        =   122
            ToolTipText     =   "Visitas Enfermeras"
            Top             =   7050
            Width           =   525
         End
         Begin VB.Frame Frame1 
            Caption         =   "Ordenes Médicas (RECETAS)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4200
            Left            =   -74895
            TabIndex        =   117
            Top             =   390
            Width           =   11610
            Begin VB.CommandButton btnAgregaApoyoDx 
               DisabledPicture =   "AdmisionHospDetalle.frx":A9E6
               DownPicture     =   "AdmisionHospDetalle.frx":ADCF
               Height          =   390
               Left            =   10890
               Picture         =   "AdmisionHospDetalle.frx":B1DB
               Style           =   1  'Graphical
               TabIndex        =   118
               ToolTipText     =   "Agrega RECETA"
               Top             =   300
               Width           =   645
            End
            Begin UltraGrid.SSUltraGrid grdApoyoDx 
               Height          =   3540
               Left            =   195
               TabIndex        =   119
               Top             =   285
               Width           =   10650
               _ExtentX        =   18785
               _ExtentY        =   6244
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
               Caption         =   "grdApoyoDx"
            End
            Begin Threed.SSCommand btnModificar 
               Height          =   390
               Left            =   10890
               TabIndex        =   120
               ToolTipText     =   "Modifica RECETA"
               Top             =   735
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   688
               _Version        =   262144
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "AdmisionHospDetalle.frx":B5E7
               PictureAlignment=   9
            End
            Begin VB.Label Label3 
               Caption         =   "<Enter> = Detalle del RESULTADO de LABORATORIO, los que tengan la palabra SI en columna RESULTADO"
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
               Left            =   180
               TabIndex        =   121
               Top             =   3870
               Width           =   11250
            End
         End
         Begin VB.Frame fraCPT 
            Caption         =   "Procedimientos realizados en el SERVICIO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            Left            =   -74910
            TabIndex        =   113
            Top             =   4620
            Width           =   11610
            Begin VB.CommandButton btnQuitarCpt 
               DisabledPicture =   "AdmisionHospDetalle.frx":E573
               DownPicture     =   "AdmisionHospDetalle.frx":E8FE
               Height          =   390
               Left            =   10875
               Picture         =   "AdmisionHospDetalle.frx":EC91
               Style           =   1  'Graphical
               TabIndex        =   115
               ToolTipText     =   "Elimina CPT"
               Top             =   795
               Width           =   645
            End
            Begin VB.CommandButton btnAgregarCpt 
               DisabledPicture =   "AdmisionHospDetalle.frx":F022
               DownPicture     =   "AdmisionHospDetalle.frx":F40B
               Height          =   390
               Left            =   10875
               Picture         =   "AdmisionHospDetalle.frx":F817
               Style           =   1  'Graphical
               TabIndex        =   114
               ToolTipText     =   "Agrega CPT"
               Top             =   315
               Width           =   645
            End
            Begin UltraGrid.SSUltraGrid grdOtrosCpt 
               Height          =   1710
               Left            =   165
               TabIndex        =   116
               Top             =   315
               Width           =   10680
               _ExtentX        =   18838
               _ExtentY        =   3016
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
               Caption         =   "CPT"
            End
         End
         Begin Threed.SSCheck SSCheck1 
            Height          =   30
            Left            =   -73590
            TabIndex        =   99
            Top             =   1260
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   53
            _Version        =   262144
            Caption         =   "SSCheck1"
         End
         Begin VB.Frame fraNotas 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   75
            TabIndex        =   93
            Top             =   5625
            Width           =   11535
            Begin VB.TextBox txtNroOrdenPago 
               BackColor       =   &H8000000F&
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
               Height          =   360
               Left            =   8910
               TabIndex        =   149
               Text            =   "..."
               Top             =   165
               Width           =   690
            End
            Begin VB.TextBox txtNroCuenta 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   360
               Left            =   5745
               TabIndex        =   147
               Text            =   ".."
               Top             =   165
               Width           =   915
            End
            Begin VB.TextBox txtEmergenciaN 
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
               Left            =   10575
               MaxLength       =   10
               TabIndex        =   38
               Top             =   165
               Width           =   900
            End
            Begin VB.TextBox txtDNIacompaniante 
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
               Left            =   4440
               MaxLength       =   8
               TabIndex        =   37
               Top             =   165
               Width           =   1290
            End
            Begin VB.TextBox txtNombreAcompañante 
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
               Left            =   1275
               MaxLength       =   100
               TabIndex        =   36
               Top             =   165
               Width           =   2760
            End
            Begin VB.Label lblOrdenPago 
               AutoSize        =   -1  'True
               Caption         =   "Nº Ord.Pago"
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
               Left            =   7860
               TabIndex        =   150
               Top             =   240
               Width           =   1035
            End
            Begin VB.Label lblEstadoCta 
               AutoSize        =   -1  'True
               Caption         =   "..."
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
               Left            =   6690
               TabIndex        =   148
               Top             =   240
               Width           =   135
            End
            Begin VB.Label lblEmergenciaN 
               AutoSize        =   -1  'True
               Caption         =   "Emerg N°"
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
               Left            =   9765
               TabIndex        =   142
               Top             =   195
               Width           =   795
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "DNI"
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
               Left            =   4125
               TabIndex        =   141
               Top             =   195
               Width           =   300
            End
            Begin VB.Label lblNombreAcompañante 
               AutoSize        =   -1  'True
               Caption         =   "Acompañante"
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
               Left            =   90
               TabIndex        =   94
               Top             =   195
               Width           =   1140
            End
         End
         Begin VB.Frame Frame7 
            Height          =   1590
            Left            =   6495
            TabIndex        =   89
            Top             =   315
            Width           =   5190
            Begin VB.TextBox txtNroAfiliacionSis 
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
               Left            =   1860
               MaxLength       =   30
               TabIndex        =   143
               TabStop         =   0   'False
               Top             =   1215
               Width           =   3240
            End
            Begin SISGalenPlus.ucSISfuaCodPrestacion ucSISfuaCodPrestacion1 
               Height          =   345
               Left            =   120
               TabIndex        =   30
               Top             =   870
               Visible         =   0   'False
               Width           =   4890
               _ExtentX        =   8625
               _ExtentY        =   609
            End
            Begin MSDataListLib.DataCombo cmbFormaPago 
               Height          =   330
               Left            =   1860
               TabIndex        =   29
               Top             =   510
               Width           =   3270
               _ExtentX        =   5768
               _ExtentY        =   582
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   "DataCombo1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDataListLib.DataCombo cmbFuenteFinanciamiento 
               Height          =   330
               Left            =   1860
               TabIndex        =   23
               Top             =   150
               Width           =   3270
               _ExtentX        =   5768
               _ExtentY        =   582
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   "DataCombo1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label lblNafiliacionSIS 
               AutoSize        =   -1  'True
               Caption         =   "N° Afiliación (SIS)"
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
               TabIndex        =   144
               Top             =   1215
               Width           =   1440
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Producto/Plan"
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
               TabIndex        =   91
               Top             =   570
               Width           =   1155
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
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
               Height          =   210
               Left            =   120
               TabIndex        =   90
               Top             =   210
               Width           =   1575
            End
         End
         Begin VB.Frame Frame6 
            Height          =   6210
            Left            =   -74880
            TabIndex        =   76
            Top             =   435
            Width           =   11430
            Begin VB.ComboBox cmbIdTipoAgenteAGAN 
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
               Left            =   3075
               TabIndex        =   55
               Top             =   4395
               Width           =   4500
            End
            Begin VB.ComboBox cmbIdGrupoOcupacionalALAB 
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
               Left            =   3075
               TabIndex        =   54
               Top             =   4020
               Width           =   4500
            End
            Begin VB.ComboBox cmbIdPosicionLesionadoALAB 
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
               Left            =   3075
               TabIndex        =   53
               Top             =   3645
               Width           =   4500
            End
            Begin VB.ComboBox cmbIdUbicacionLesionado 
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
               Left            =   3075
               TabIndex        =   52
               Top             =   3270
               Width           =   4500
            End
            Begin VB.ComboBox cmbIdTipoTransporte 
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
               Left            =   3075
               TabIndex        =   51
               Top             =   2895
               Width           =   4500
            End
            Begin VB.ComboBox cmbIdTipoVehiculo 
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
               Left            =   3075
               TabIndex        =   50
               Top             =   2520
               Width           =   4500
            End
            Begin VB.ComboBox cmbIdClaseAccidente 
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
               Left            =   3075
               TabIndex        =   49
               Top             =   2145
               Width           =   4500
            End
            Begin VB.ComboBox cmbIdRelacionAgresorVictima 
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
               Left            =   3075
               TabIndex        =   48
               Top             =   1770
               Width           =   4500
            End
            Begin VB.ComboBox cmbIdSeguridad 
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
               Left            =   3075
               TabIndex        =   47
               Top             =   1395
               Width           =   4500
            End
            Begin VB.ComboBox cmbIdLugarEvento 
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
               Left            =   3075
               TabIndex        =   45
               Top             =   645
               Width           =   4500
            End
            Begin VB.ComboBox cmbIdCausaExternaMorbilidad 
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
               Left            =   3075
               TabIndex        =   43
               Top             =   270
               Width           =   4500
            End
            Begin VB.ComboBox cmbIdTipoEvento 
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
               Left            =   3075
               TabIndex        =   46
               Top             =   1020
               Width           =   4500
            End
            Begin VB.Label lblIdTipoAgenteAGAN 
               Caption         =   "Tipo de agente AGAN"
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
               Left            =   210
               TabIndex        =   88
               Top             =   4395
               Width           =   2250
            End
            Begin VB.Label lblIdGrupoOcupacionalALAB 
               Caption         =   "Grupo ocupacional ALAB"
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
               Left            =   210
               TabIndex        =   87
               Top             =   4020
               Width           =   2250
            End
            Begin VB.Label lblIdPosicionLesionadoALAB 
               Caption         =   "Posición del lesionado ALAB"
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
               Left            =   210
               TabIndex        =   86
               Top             =   3645
               Width           =   2250
            End
            Begin VB.Label lblIdUbicacionLesionado 
               Caption         =   "Ubicación del lesionado"
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
               Left            =   210
               TabIndex        =   85
               Top             =   3270
               Width           =   2250
            End
            Begin VB.Label lblIdTipoTransporte 
               Caption         =   "Tipo de transporte"
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
               Left            =   210
               TabIndex        =   84
               Top             =   2895
               Width           =   2250
            End
            Begin VB.Label lblIdTipoVehiculo 
               Caption         =   "Tipo de vehiculo"
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
               Left            =   210
               TabIndex        =   83
               Top             =   2520
               Width           =   2250
            End
            Begin VB.Label lblIdClaseAccidente 
               Caption         =   "Clase de accidente"
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
               Left            =   210
               TabIndex        =   82
               Top             =   2145
               Width           =   2250
            End
            Begin VB.Label lblIdRelacionAgresorVictima 
               Caption         =   "Relacion agresor víctima"
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
               Left            =   210
               TabIndex        =   81
               Top             =   1770
               Width           =   2250
            End
            Begin VB.Label lblIdSeguridad 
               Caption         =   "Seguridad"
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
               Left            =   210
               TabIndex        =   80
               Top             =   1395
               Width           =   2250
            End
            Begin VB.Label lblIdTipoEvento 
               Caption         =   "Tipo de evento"
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
               Left            =   210
               TabIndex        =   79
               Top             =   1020
               Width           =   2250
            End
            Begin VB.Label lblIdLugarEvento 
               Caption         =   "Lugar del evento"
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
               Left            =   210
               TabIndex        =   78
               Top             =   645
               Width           =   2250
            End
            Begin VB.Label lblIdCausaExternaMorbilidad 
               Caption         =   "Causa externa de morbilidad"
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
               Left            =   210
               TabIndex        =   77
               Top             =   270
               Width           =   2415
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Condición del paciente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   -74865
            TabIndex        =   68
            Top             =   6150
            Visible         =   0   'False
            Width           =   11505
            Begin VB.ComboBox cmbIdCondicionEnElEstablecimiento 
               Height          =   315
               Left            =   7350
               TabIndex        =   74
               Top             =   210
               Width           =   3480
            End
            Begin VB.ComboBox cmbIdCondicionEnElServicio 
               Height          =   315
               Left            =   1260
               TabIndex        =   73
               Top             =   240
               Width           =   3150
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "En el servicio"
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
               TabIndex        =   70
               Top             =   270
               Width           =   1050
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "En el establecimiento"
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
               Left            =   5565
               TabIndex        =   69
               Top             =   270
               Width           =   1740
            End
         End
         Begin SISGalenPlus.ucTransferenciasDetalle ucTransferenciasDetalle1 
            Height          =   4995
            Left            =   -74865
            TabIndex        =   40
            Top             =   405
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   8811
         End
         Begin VB.Frame fraDatosReferenciaOrigen 
            Caption         =   "Referencia Origen"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1620
            Left            =   6495
            TabIndex        =   65
            Top             =   1905
            Width           =   5160
            Begin VB.CommandButton btnBuscarEstablecimiento 
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
               Left            =   2175
               Picture         =   "AdmisionHospDetalle.frx":FC23
               Style           =   1  'Graphical
               TabIndex        =   160
               Top             =   570
               Width           =   330
            End
            Begin VB.TextBox txtMedicoRef 
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
               Left            =   1665
               MaxLength       =   8
               TabIndex        =   154
               ToolTipText     =   "Buscar por COLEGIATURA"
               Top             =   1230
               Width           =   780
            End
            Begin VB.ComboBox cmbMedicoRef 
               BackColor       =   &H00FFFFFF&
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
               Height          =   330
               Left            =   2445
               TabIndex        =   153
               Top             =   1230
               Width           =   2685
            End
            Begin VB.TextBox txtReferenciaO 
               Height          =   315
               Left            =   4200
               TabIndex        =   33
               Top             =   210
               Width           =   900
            End
            Begin VB.ComboBox cmbIdTipoReferenciaOrigen 
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
               Left            =   1650
               TabIndex        =   31
               Top             =   210
               Width           =   1620
            End
            Begin VB.TextBox lblNombreOrigenReferencia 
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
               Left            =   2520
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   570
               Width           =   2580
            End
            Begin VB.TextBox txtIdEstablecimientoOrigen 
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
               Left            =   1650
               TabIndex        =   32
               Top             =   566
               Width           =   525
            End
            Begin PVCOMBOLibCtl.PVComboBox cmbServicioReferenciaO 
               Height          =   330
               Left            =   1665
               TabIndex        =   34
               Top             =   900
               Width           =   3465
               _Version        =   524288
               _cx             =   6112
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
               ItemLabelWidth  =   40
               ItemLabelForeColor=   0
               ItemLabelBackColor=   13160660
               ColumnHeaderStyle=   1
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
               ColumnCount     =   2
               Column0.Heading =   "Código"
               Column0.Width   =   60
               Column0.Alignment=   0
               Column0.Hidden  =   0   'False
               Column0.Name    =   "codigo"
               Column0.Format  =   ""
               Column0.Bound   =   -1  'True
               Column0.Locked  =   0   'False
               Column0.HeaderAlignment=   0
               Column1.Heading =   "Descripción"
               Column1.Width   =   200
               Column1.Alignment=   0
               Column1.Hidden  =   0   'False
               Column1.Name    =   "descripcion"
               Column1.Format  =   ""
               Column1.Bound   =   -1  'True
               Column1.Locked  =   0   'False
               Column1.HeaderAlignment=   0
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
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Médico referencia"
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
               Left            =   105
               TabIndex        =   155
               Top             =   1305
               Width           =   1440
            End
            Begin VB.Label lblServicioReferencia 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Servicio referencia"
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
               Left            =   105
               TabIndex        =   140
               Top             =   945
               Width           =   1485
            End
            Begin VB.Label lblNreferencia 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N° refer"
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
               Left            =   3495
               TabIndex        =   95
               Top             =   240
               Width           =   660
            End
            Begin VB.Label lblIdTipoReferenciaOrigen 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo referencia"
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
               Left            =   105
               TabIndex        =   67
               Top             =   255
               Width           =   1230
            End
            Begin VB.Label lblIdEstablecimientoOrigen 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Estab. referencia"
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
               Left            =   105
               TabIndex        =   66
               Top             =   600
               Width           =   1380
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Atención"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3195
            Left            =   90
            TabIndex        =   58
            Top             =   315
            Width           =   6345
            Begin VB.CommandButton btnVerDisponibilidadDeCamas 
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
               Left            =   2235
               Picture         =   "AdmisionHospDetalle.frx":101AD
               Style           =   1  'Graphical
               TabIndex        =   163
               Top             =   2010
               Width           =   360
            End
            Begin VB.CommandButton btnBuscarMedicos 
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
               Left            =   4920
               Picture         =   "AdmisionHospDetalle.frx":10737
               Style           =   1  'Graphical
               TabIndex        =   162
               Top             =   1320
               Width           =   330
            End
            Begin VB.CommandButton btnBuscarServicios 
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
               Left            =   4890
               Picture         =   "AdmisionHospDetalle.frx":10CC1
               Style           =   1  'Graphical
               TabIndex        =   161
               Top             =   600
               Width           =   330
            End
            Begin VB.CheckBox chkLlegoSI 
               Alignment       =   1  'Right Justify
               Caption         =   "Llegó al 'Servicio Ingreso''"
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
               Left            =   3750
               TabIndex        =   159
               Top             =   2055
               Visible         =   0   'False
               Width           =   2415
            End
            Begin VB.ComboBox cmbEstadoLlegada 
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
               ItemData        =   "AdmisionHospDetalle.frx":1124B
               Left            =   1170
               List            =   "AdmisionHospDetalle.frx":11258
               TabIndex        =   22
               Top             =   2775
               Width           =   1920
            End
            Begin VB.ComboBox cmbComoLlego 
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
               ItemData        =   "AdmisionHospDetalle.frx":1127E
               Left            =   1170
               List            =   "AdmisionHospDetalle.frx":1128E
               TabIndex        =   20
               Top             =   2385
               Width           =   1920
            End
            Begin VB.ComboBox cmbTipoAtencion 
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
               ItemData        =   "AdmisionHospDetalle.frx":112BE
               Left            =   4410
               List            =   "AdmisionHospDetalle.frx":112C8
               TabIndex        =   21
               Top             =   2400
               Width           =   1860
            End
            Begin VB.Frame Frame8 
               Enabled         =   0   'False
               Height          =   465
               Left            =   4485
               TabIndex        =   96
               Top             =   1575
               Width           =   1755
               Begin VB.CheckBox chkRecienNacido 
                  Alignment       =   1  'Right Justify
                  Caption         =   "¿Recien nacido?"
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
                  Left            =   60
                  TabIndex        =   97
                  ToolTipText     =   "Indica al sistema que en el caso de los recien nacidos la edad se debe calcular como: (fecha de egreso - fecha de nacimiento)"
                  Top             =   135
                  Width           =   1605
               End
            End
            Begin VB.TextBox lblNombreMedico 
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
               Left            =   1170
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   1305
               Width           =   3735
            End
            Begin VB.TextBox lblNombreServicio 
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
               Left            =   1170
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   615
               Width           =   3705
            End
            Begin VB.ComboBox cmbIdTipoEdad 
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
               ItemData        =   "AdmisionHospDetalle.frx":112E2
               Left            =   2025
               List            =   "AdmisionHospDetalle.frx":112E4
               TabIndex        =   27
               Top             =   1665
               Width           =   1425
            End
            Begin VB.ComboBox cmbIdTipoGravedad 
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
               ItemData        =   "AdmisionHospDetalle.frx":112E6
               Left            =   4275
               List            =   "AdmisionHospDetalle.frx":112E8
               TabIndex        =   18
               Text            =   "cmbIdTipoGravedad"
               Top             =   960
               Width           =   1995
            End
            Begin VB.ComboBox cmbIdViasAdmision 
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
               Left            =   4125
               TabIndex        =   14
               Text            =   "cmbIdViasAdmision"
               Top             =   225
               Width           =   2175
            End
            Begin VB.ComboBox cmbIdTipoServicio 
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
               Left            =   1170
               TabIndex        =   13
               Top             =   255
               Width           =   2325
            End
            Begin VB.TextBox txtEdadEnDias 
               BackColor       =   &H00FFFFFF&
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
               Left            =   1170
               TabIndex        =   26
               Top             =   1665
               Width           =   855
            End
            Begin VB.TextBox txtIdServicioIngreso 
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
               Left            =   5235
               TabIndex        =   24
               Top             =   615
               Width           =   1020
            End
            Begin VB.TextBox txtIdMedicoIngreso 
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
               Left            =   5250
               TabIndex        =   25
               Top             =   1320
               Width           =   1005
            End
            Begin VB.TextBox txtNroCamaIngreso 
               BackColor       =   &H00FFFFFF&
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
               Left            =   1170
               TabIndex        =   28
               Top             =   2040
               Width           =   1035
            End
            Begin MSMask.MaskEdBox txtHoraIngreso 
               Height          =   315
               Left            =   2385
               TabIndex        =   17
               Top             =   930
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtFechaIngreso 
               Height          =   315
               Left            =   1170
               TabIndex        =   16
               Top             =   930
               Width           =   1215
               _ExtentX        =   2143
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
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Estado"
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
               TabIndex        =   156
               Top             =   2835
               Width           =   555
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Como llegó"
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
               TabIndex        =   152
               Top             =   2445
               Width           =   900
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Atención"
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
               Left            =   3210
               TabIndex        =   151
               Top             =   2475
               Width           =   1155
            End
            Begin VB.Label lblFecha 
               AutoSize        =   -1  'True
               Caption         =   "Fecha ingr"
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
               TabIndex        =   98
               Top             =   1005
               Width           =   840
            End
            Begin VB.Label lblGravedad 
               AutoSize        =   -1  'True
               Caption         =   "Gravedad"
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
               Left            =   3495
               TabIndex        =   75
               Top             =   1020
               Width           =   765
            End
            Begin VB.Label lblEdadEnDias 
               AutoSize        =   -1  'True
               Caption         =   "Edad"
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
               TabIndex        =   64
               Top             =   1710
               Width           =   405
            End
            Begin VB.Label lblIdTipoServicio 
               AutoSize        =   -1  'True
               Caption         =   "Tipo servicio"
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
               TabIndex        =   63
               Top             =   300
               Width           =   1005
            End
            Begin VB.Label lblIdServicioIngreso 
               AutoSize        =   -1  'True
               Caption         =   "Servicio ingr"
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
               Top             =   660
               Width           =   975
            End
            Begin VB.Label lblIdMedicoIngreso 
               AutoSize        =   -1  'True
               Caption         =   "Medico ing"
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
               TabIndex        =   61
               Top             =   1365
               Width           =   870
            End
            Begin VB.Label lblViaAdmision 
               AutoSize        =   -1  'True
               Caption         =   "Origen"
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
               TabIndex        =   60
               Top             =   285
               Width           =   540
            End
            Begin VB.Label lblNroCamaIngreso 
               AutoSize        =   -1  'True
               Caption         =   "Cama Ing"
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
               TabIndex        =   59
               Top             =   2055
               Width           =   765
            End
         End
         Begin SISGalenPlus.ucDiagnosticoDetalle ucDiagnosticosIngreso 
            Height          =   2070
            Left            =   45
            TabIndex        =   92
            Top             =   3480
            Width           =   11580
            _ExtentX        =   20426
            _ExtentY        =   3651
         End
         Begin SISGalenPlus.UcPacientesSunasa UcPacientesSunasa1 
            Height          =   225
            Left            =   -66030
            TabIndex        =   112
            Top             =   930
            Visible         =   0   'False
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   397
         End
         Begin SISGalenPlus.ucNacimientoDetalle ucNacimientoDetalle1 
            Height          =   3075
            Left            =   -74910
            TabIndex        =   133
            Top             =   1095
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   5424
         End
         Begin SISGalenPlus.ucDiagnosticoDetalle ucDiagnosticoNacimiento 
            Height          =   2445
            Left            =   -74940
            TabIndex        =   134
            Top             =   4200
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   4313
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nº Cuenta"
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
            Left            =   6015
            TabIndex        =   101
            Top             =   6330
            Width           =   855
         End
      End
      Begin SISGalenPlus.ucMensajeParpadeando ucMensajeParpadeando1 
         Height          =   315
         Left            =   -74820
         TabIndex        =   100
         Top             =   6840
         Visible         =   0   'False
         Width           =   11685
         _ExtentX        =   8017
         _ExtentY        =   1296
      End
   End
   Begin SISGalenPlus.ucPacientesCtasPDF ucPacientesCtasPDF1 
      Height          =   3675
      Left            =   12120
      TabIndex        =   171
      Top             =   4920
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   6482
   End
   Begin VB.Image pi_imagen 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   0
      MouseIcon       =   "AdmisionHospDetalle.frx":112EA
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      ToolTipText     =   "Pulsar Doble Click para ampliar Imagen"
      Top             =   0
      Visible         =   0   'False
      Width           =   2745
   End
End
Attribute VB_Name = "AdmisionHospDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Admisión del Paciente en Hospitalización/Emergencia
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lbHuboCambioEnDato As Boolean
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim mb_ExistenDatos As Boolean
Dim ml_TipoServicio As sghTipoServicio
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim ml_EstadoCuenta As Long
'
Dim mo_sighProxies As New SIGHProxies.Procesos
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim ms_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdminHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mo_SisConsumoWeb As New SIGHNegocios.SisConsumoWeb
Dim mo_Reniec As New ReniecGalenhos
Dim mo_AdminReglasHoteleria As New ReglasHoteleria
 
'
Dim mo_cmbIdTipoGravedad As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoServicio As New sighEntidades.ListaDespleglable
Dim mo_cmbIdViasAdmision As New sighEntidades.ListaDespleglable
Dim mo_cmbIdEspecialidadMedico As New sighEntidades.ListaDespleglable
Dim mo_cmbIdServicio As New sighEntidades.ListaDespleglable
Dim mo_cmbIdCondicionEnElServicio As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoReferenciaOrigen As New sighEntidades.ListaDespleglable
Dim mo_cmbIdCondicionEnElEstablecimiento As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoAgenteAGAN As New sighEntidades.ListaDespleglable
Dim mo_cmbIdGrupoOcupacionalALAB As New sighEntidades.ListaDespleglable
Dim mo_cmbIdPosicionLesionadoALAB As New sighEntidades.ListaDespleglable
Dim mo_cmbIdUbicacionLesionado As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoTransporte As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoVehiculo As New sighEntidades.ListaDespleglable
Dim mo_cmbIdClaseAccidente As New sighEntidades.ListaDespleglable
Dim mo_cmbIdRelacionAgresorVictima As New sighEntidades.ListaDespleglable
Dim mo_cmbIdSeguridad As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoEvento As New sighEntidades.ListaDespleglable
Dim mo_cmbIdLugarEvento As New sighEntidades.ListaDespleglable
Dim mo_cmbIdCausaExternaMorbilidad As New sighEntidades.ListaDespleglable
Dim mo_cmbIdTipoEdad As New sighEntidades.ListaDespleglable
'
Dim mo_DoUbicacionPaciente As New doPaciente
Dim mo_AtencionesEmergencia As New DOAtencionEmergencia
Dim mo_AtencionPadre As New DOAtencion
Dim mo_DoAtencionDatosAdicionales As New DoAtencionDatosAdicionales
Dim ldFechaEgresoMedicoAnterior As Date   'cuando se "modifique", generar "consumo por dias estancia"
Dim mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String
Dim wxParametroBusqRapida As String
'------------------------------------------------------------------------------------
'                               VARIABLES CUENTAS DE ATENCION
'------------------------------------------------------------------------------------
Dim mo_CuentasAtencion As New DOCuentaAtencion
Dim ml_idCuentaAtencion As Long

'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA ATENCION
'------------------------------------------------------------------------------------
Dim mo_Atenciones As New DOAtencion
Dim ml_idAtencion As Long
Dim ml_IdAtencionEmergencia As Long
Dim oRsEstancia As New Recordset
Dim mrs_OcupacionCamas As New ADODB.Recordset
Dim mo_OcupacionCamas As New Collection
Dim mrs_Interconsulta As New ADODB.Recordset
Dim ml_TipoAccionAdmision As sghTipoAccionEmergenciaYHospitalizacion
Dim ml_IdAtencionPadre As Long
'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA ATENCION HOSPITALIZACION
'------------------------------------------------------------------------------------
Dim ml_IdAtencionHosp As Long
Dim mo_Diagnosticos As New Collection
Dim mo_Procedimientos As New Collection
Dim mo_Examenes As New Collection
Dim mo_Nacimientos As New Collection
Dim mo_NroServiciosQuePasoElPaciente As Long
Dim lcUltimoCodigoDeServicioTransferido As String

'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA FILIACION
'------------------------------------------------------------------------------------
Dim ml_IdPaciente As Long
Dim mo_Pacientes  As New doPaciente
Dim ms_Autogenerado As String
Dim ml_TipoGeneracionHistoria As sghTipoNumeracionDeNroHistoria
Dim mo_Historia As New DOHistoriaClinica
Dim mo_DoPacientesDatosAdd As New DoPacienteDatosAdd

'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA CITA
'------------------------------------------------------------------------------------
Dim ml_IdMedico As Long
Dim ms_NombreMedico  As String
Dim mo_Especialidad As New DOEspecialidades
Dim mo_paciente As New doPaciente
Dim ml_IdPrestamo As Long
Dim ml_IdEspecialidad As Long

'------------------------------------------------------------------------------------
'                               PACIENTE NUEVO
'------------------------------------------------------------------------------------
Dim oRsFormaPago As New ADODB.Recordset
Dim oRsFuentesFinanciamiento As New ADODB.Recordset

Dim lcApP As String
Dim lcApM As String
Dim lcPnom As String
Dim lcSnombreReniec As String, ldFnacimientoReniec As Date, lnIdSexoReniec As Long
Dim lcDireccionReniec As String, mb_UsoWebReniec As Boolean
Dim lnIdDistritoSIS As Long, lnIdSexoSIS As Long, ldFechaNacimientoSIS As Date, lcSnombreSIS As String
Dim lnIdPlanSIS As Long, lcDniSIS As String, lnAfiliacionSIS1 As String, lnAfiliacionSIS2 As String, lnAfiliacionSIS5 As String
Dim lnAfiliacionSIS3 As String, lnAfiliacionSIS4 As Long, lcSIScodigo As String, lcTipoFormatoSIS As String
Dim lcCodigoEstablecimientoAdscripcionSIS As String, lbEncontroAfiliadoEnWebSIS As Boolean
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_IdServicioConCamaDisponible As Long
Dim lbPacienteNN As Boolean
Dim lnSecuenciaTransferencia As Long
Dim lnFocusCuandoCargeFrm As Long
Dim lbUsuarioConfirmaTransferencia As Boolean
Dim lbUsuarioConfirmaLlegada As Boolean
Dim lcCaptionTab2 As String
Dim lbUltimaTeclaPulsoENTER As Boolean
Dim lnIdPlanAnterior As Long
Dim lnIdTipoFinanciamientoAnterior As Long
Dim mo_lbCargaTablasUnaVez As Boolean
Dim mo_lbNuevoMovimiento As Boolean
Dim lnIdNacimientoSeleccionado As Long
Dim oDoSunasaPacientesHistoricos As New DoSunasaPacientesHistoricos
Dim mo_DOAtencionesCE As New DOAtencionesCE
Dim mb_EsObservacionEmergencia As Boolean
Dim lbBuscaDNIenReniec As Boolean
Dim lcElServicioUsaGalenHos As String
Dim ldFechaActualServidor As Date
Const lbCargaAlaVezCitaPacienteAtencionDA As Boolean = True
Dim ml_ldFechaEgreso As Date, ml_idServicioEgreso As Long, ml_lcServicioEgreso As String
Dim ml_lcCodigoServicioEgreso As String
Dim ml_idCamaEgreso As Long, ml_lcCamaEgreso As String, ml_lcHoraEgreso As String
Dim ml_idMedicoEgreso As Long, ml_lcMedicoEgreso As String
Dim lbElServicioRegistraFUA As String, lbCargaUnaVezVEntana As Boolean
Dim dxMorbilidadExterna342 As String
Dim lbProcedeDeConsExt As Boolean, lbProcedeDeEmergencia As Boolean     'debb-23/02/2015
Dim ml_idOrden As Long, ml_FechaReceta As Date
Dim mc_FuaVersionFormato As String
Dim lb_puedeCambiarFuenteFinanciamiento As Boolean      'debb-14/03/2015
Dim lcHistoriaYpaciente As String
Dim ml_idAtencionEmeg_CE As Long
Dim lnDocumentoTipoSIS As Long

Property Let lbNuevoMovimiento(lValue As Boolean)
   mo_lbNuevoMovimiento = lValue
End Property
Property Let lbCargaTablasUnaVez(lValue As Boolean)
   mo_lbCargaTablasUnaVez = lValue
End Property
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
   'mgaray09
   setListItemAControlDiagnosticos mo_lnIdTablaLISTBARITEMS
End Property
Property Let IdServicioConCamaDisponible(lValue As Long)
    ml_IdServicioConCamaDisponible = lValue
End Property
Property Get IdServicioConCamaDisponible() As Long
    IdServicioConCamaDisponible = ml_IdServicioConCamaDisponible
End Property


Property Let idMedico(lValue As Long)
   ml_IdMedico = lValue
End Property
Property Get idMedico() As Long
   idMedico = ml_IdMedico
End Property
Property Let IdPrestamo(lValue As Long)
   ml_IdPrestamo = lValue
End Property
Property Get IdPrestamo() As Long
   IdPrestamo = ml_IdPrestamo
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
Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let idAtencion(lValue As Long)
   ml_idAtencion = lValue
End Property
Property Get idAtencion() As Long
   idAtencion = ml_idAtencion
End Property
Property Let IdAtencionHosp(lValue As Long)
   ml_IdAtencionHosp = lValue
End Property
Property Get IdAtencionHosp() As Long
   IdAtencionHosp = ml_IdAtencionHosp
End Property
Property Let IdAtencionEmergencia(lValue As Long)
   ml_IdAtencionEmergencia = lValue
End Property
Property Get IdAtencionEmergencia() As Long
   IdAtencionEmergencia = ml_IdAtencionEmergencia
End Property
Property Let idPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_IdPaciente
End Property
Property Let TipoServicio(sValue As sghTipoServicio)
   ml_TipoServicio = sValue
End Property
Property Get TipoServicio() As sghTipoServicio
   TipoServicio = ml_TipoServicio
End Property
Property Let TipoGeneracionHistoria(lValue As Long)
   ml_TipoGeneracionHistoria = lValue
End Property
Property Get TipoGeneracionHistoria() As Long
   TipoGeneracionHistoria = ml_TipoGeneracionHistoria
End Property
Property Let IdEspecialidad(lValue As Long)
   ml_IdEspecialidad = lValue
End Property
Property Get IdEspecialidad() As Long
   IdEspecialidad = ml_IdEspecialidad
End Property
Property Let TipoAccionDeAdmision(lValue As sghTipoAccionEmergenciaYHospitalizacion)
    ml_TipoAccionAdmision = lValue
End Property
Property Get TipoAccionDeAdmision() As sghTipoAccionEmergenciaYHospitalizacion
    TipoAccionDeAdmision = ml_TipoAccionAdmision
End Property

Property Let IdAtencionPadre(lValue As Long)
   ml_IdAtencionPadre = lValue
End Property
Property Get IdAtencionPadre() As Long
   IdAtencionPadre = ml_IdAtencionPadre
End Property
Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

       mo_cmbIdTipoReferenciaOrigen.BoundColumn = "IdTipoReferencia"
       mo_cmbIdTipoReferenciaOrigen.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoReferenciaOrigen.RowSource = mo_AdminServiciosComunes.TiposReferenciaSeleccionarTodos
       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
       
       
       mo_cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
       mo_cmbIdTipoServicio.ListField = "DescripcionLarga"
       Select Case ml_TipoServicio
       Case sghConsultaExterna
            Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarAsistenciales
            mo_cmbIdTipoServicio.BoundText = "1"
            mo_Formulario.HabilitarDeshabilitar cmbIdTipoServicio, False
        Case sghHospitalizacion
            Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarAsistenciales
            mo_cmbIdTipoServicio.BoundText = "3"
            mo_Formulario.HabilitarDeshabilitar cmbIdTipoServicio, False
        Case sghEmergenciaConsultorios
            Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarDeEmergencia
            mo_Formulario.HabilitarDeshabilitar cmbIdTipoServicio, True
            mo_cmbIdTipoServicio.BoundText = "2"
        Case sghEmergenciaObservacion
            Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarDeEmergencia
            mo_Formulario.HabilitarDeshabilitar cmbIdTipoServicio, True
            mo_cmbIdTipoServicio.BoundText = "4"
        End Select
        
        
       
       If sMensaje <> "" Then
           MsgBox sMensaje, vbInformation, Me.Caption
       End If

       mo_cmbIdTipoEdad.BoundColumn = "IdTipoEdad"
       mo_cmbIdTipoEdad.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoEdad.RowSource = mo_AdminServiciosComunes.TiposEdadSeleccionarTodos
       sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
       
        Dim rsTiposGravedad As New ADODB.Recordset
        Set rsTiposGravedad = mo_AdminServiciosComunes.TipoGravedadAtencionSeleccionarTodos()
        mo_cmbIdTipoGravedad.CargarComboBoxDesdeRecordset cmbIdTipoGravedad, rsTiposGravedad, "IdTipoGravedad", "Descripcion"

        If ml_TipoServicio = sghEmergenciaConsultorios Or ml_TipoServicio = sghEmergenciaObservacion Then
            
            mo_cmbIdCausaExternaMorbilidad.BoundColumn = "IdCausaExternaMorbilidad"
            mo_cmbIdCausaExternaMorbilidad.ListField = "DescripcionLarga"
            Set mo_cmbIdCausaExternaMorbilidad.RowSource = mo_AdminAdmision.EmergenciaCausaExternaMorbilidadSeleccionarTodos()
            sMensaje = sMensaje + mo_AdminAdmision.MensajeError
                        
        End If
        
        Me.ucDiagnosticosIngreso.TipoDiagnostico = sghHospitalizacionIngreso

        Me.ucPacientesDetalle1.ConfigurarComboBoxes
        Me.ucPacientesDetalle1.Opcion = mi_Opcion
        Me.ucDiagnosticosIngreso.ConfigurarComboBoxes
        '
        Set oRsFormaPago = mo_AdminServiciosComunes.TiposFinanciamientoSegunFiltro("esFuenteFinanciamiento=1")
        Set cmbFormaPago.RowSource = oRsFormaPago
        cmbFormaPago.ListField = "Descripcion"
        cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
        mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, False
        '
        Set oRsFuentesFinanciamiento = mo_AdminServiciosComunes.FuentesFinanciamientoSegunFiltro("UtilizadoEn=2 or UtilizadoEn=3")
        Set cmbFuenteFinanciamiento.RowSource = oRsFuentesFinanciamiento
        cmbFuenteFinanciamiento.ListField = "Descripcion"
        cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"
        '
        'debb2009-desde la admision de Hospitalizacion o Emergencia
        'debb2009-se identifica el seguro del Paciente
        If wxParametro231 <> "S" Then
            mo_Formulario.HabilitarDeshabilitar cmbFuenteFinanciamiento, False
            cmbFuenteFinanciamiento.BoundText = "5"
            cmbFormaPago.BoundText = "1"
        End If
        'debb2009-fin
        mo_Formulario.HabilitarDeshabilitar txtEdadEnDias, False
        mo_Formulario.HabilitarDeshabilitar cmbIdTipoEdad, False
        mo_Formulario.HabilitarDeshabilitar txtNroCamaIngreso, False
        '
        Me.ucNacimientoDetalle1.ConfigurarComboBoxes
        Me.ucDiagnosticoNacimiento.TipoDiagnostico = sghHospitalizacionNacimiento
        Me.ucDiagnosticoNacimiento.IdListBarItem = mo_lnIdTablaLISTBARITEMS
        Me.ucDiagnosticoNacimiento.ConfigurarComboBoxes
        
        Set cmbServicioReferenciaO.ListSource = mo_AdminServiciosComunes.SuSalud_upsSeleccionarTodos   'debb-21/06/2016
End Sub





Private Sub btnAgregaApoyoDx_Click()
    Dim oReceta As New RecetaDetalle
    Dim lnIdMedico As Long
    lnIdMedico = Me.ucTransferenciasDetalle1.getIdMedicoUltimaTransferencia
    If lnIdMedico = 0 Then
       lnIdMedico = mo_Atenciones.IdMedicoIngreso
    End If
    oReceta.Opcion = sghAgregar
    oReceta.CargaDxParaFarmacia ucDiagnosticosIngreso.DevuelveDx
    oReceta.idTipoServicio = ml_TipoServicio
    oReceta.idUsuario = ml_idUsuario
    oReceta.idCuentaAtencion = ml_idCuentaAtencion
    oReceta.IdMedicoServicioActual = lnIdMedico
    oReceta.Show 1
    Set oReceta = Nothing
    CargaApoyoDx
End Sub

Private Sub btnBuscaHistoricos_Click()
   If lcHistoriaYpaciente <> "" Then
    Dim oBuscaHistoricos As New AdmisionCEhistorico
    oBuscaHistoricos.MuestraTab = 2
    oBuscaHistoricos.Paciente = lcHistoriaYpaciente
    oBuscaHistoricos.idPaciente = ml_IdPaciente
    oBuscaHistoricos.idTipoSexo = mo_paciente.idTipoSexo
    oBuscaHistoricos.NroHistoriaClinica = Val(Mid(lcHistoriaYpaciente, 2, InStr(lcHistoriaYpaciente, ")") - 2))
    oBuscaHistoricos.Show 1
    Set oBuscaHistoricos = Nothing
   End If
End Sub

Private Sub btnBuscarEstablecimiento_Click()
    If cmbIdTipoReferenciaOrigen.Text <> "" Then
        CompletarDatosDeEstablecimiento txtIdEstablecimientoOrigen, lblNombreOrigenReferencia, mo_cmbIdTipoReferenciaOrigen.BoundText
        grdPacientesEncontrados.Visible = False
    End If
End Sub
Private Sub btnBuscarEstablecimientoDestino_Click()

End Sub

Private Sub btnBuscarMedicos_Click()
    CompletarDatosDeMedico txtIdMedicoIngreso, lblNombreMedico, Val(Me.lblNombreServicio.Tag), "", CDate(Me.txtFechaIngreso.Text), Me.txtHoraIngreso.Text, ml_TipoServicio
    tabAdmision.Tab = 1
End Sub



Sub BuscarPacientesSISSegunFiltro()
    Dim lcSql As String, oRsBuscaPacientesSis As New Recordset
    If UCase(Trim(lnAfiliacionSIS2)) = "R" Then
       lnAfiliacionSIS2 = UCase(Trim(lnAfiliacionSIS2))
    End If
    If (lnAfiliacionSIS1 <> "" And lnAfiliacionSIS2 <> "" And lnAfiliacionSIS3 <> "") Or _
       (lnAfiliacionSIS2 = "R" And lnAfiliacionSIS3 <> "") Or _
       (txtNroDNIBusqueda.Text <> "") Or (txtApellidoPaternoBusqueda.Text <> "") Then
       lcSql = ""
       If (lnAfiliacionSIS1 <> "" And lnAfiliacionSIS2 <> "" And lnAfiliacionSIS3 <> "") Then
          lcSql = "  where afiliacionDisa='" & lnAfiliacionSIS1 & "' and AfiliacionTipoFormato='" & lnAfiliacionSIS2 & "' and AfiliacionNroFormato='" & lnAfiliacionSIS3 & "' order by paterno,materno,pnombre"
       ElseIf lnAfiliacionSIS2 = "R" And lnAfiliacionSIS3 <> "" Then
          lcSql = "  where AfiliacionTipoFormato='" & lnAfiliacionSIS2 & "' and AfiliacionNroFormato='" & lnAfiliacionSIS3 & "' order by paterno,materno,pnombre"
       ElseIf txtNroDNIBusqueda.Text <> "" Then
          lcSql = "   where  DocumentoTipo=1 and DocumentoNumero='" & txtNroDNIBusqueda.Text & "'"
       ElseIf txtApellidoPaternoBusqueda.Text <> "" Then
          lcSql = "   where  paterno like '%" & Trim(txtApellidoPaternoBusqueda.Text) & "%'"
          If txtApellidoMaternoBusqueda.Text <> "" Then
             lcSql = lcSql & " and  materno like '%" & Trim(txtApellidoMaternoBusqueda.Text) & "%'"
          End If
          If txtPrimerNombreBusqueda.Text <> "" Then
             lcSql = lcSql & " and  pnombre like '%" & Trim(txtPrimerNombreBusqueda.Text) & "%'"
          End If
          If txtSegundoNombreBusqueda.Text <> "" Then
             lcSql = lcSql & " and  onombres like '%" & Trim(txtSegundoNombreBusqueda.Text) & "%'"
          End If
       End If
       If lcSql <> "" Then
           lbEncontroAfiliadoEnWebSIS = False
           If wxParametro322 = "S" Then
              If (lnAfiliacionSIS1 <> "" And lnAfiliacionSIS2 <> "" And lnAfiliacionSIS3 <> "") Or _
                                             (lnAfiliacionSIS2 = "R" And lnAfiliacionSIS3 <> "") Then
                  '**************************Busca en Pag WEB del SIS x Nro Afiliado*******************
                  If Trim(lnAfiliacionSIS1) = "080" And Trim(lnAfiliacionSIS2) = "3" Then
                        lnAfiliacionSIS3 = Right("000000000" & Trim(lnAfiliacionSIS3), 9)
                  Else
                        lnAfiliacionSIS3 = Right("00000000" & Trim(lnAfiliacionSIS3), 8)
                  End If
'                  Set oRsBuscaPacientesSis = mo_ReglasSISgalenhos.ConsultaWebSisBuscarAfiliado("", Trim(lnAfiliacionSIS1), _
'                                                     Trim(lnAfiliacionSIS2), lnAfiliacionSIS3, _
'                                                     "", lcTipoFormatoSIS, wxParametro323)
                  'FCV 17072015
                  Set oRsBuscaPacientesSis = mo_SisConsumoWeb.WebServiceSISBuscarAfiliado("", Trim(lnAfiliacionSIS1), _
                                                     Trim(lnAfiliacionSIS2), lnAfiliacionSIS3, _
                                                     "", lcTipoFormatoSIS, wxParametro323)
                  If oRsBuscaPacientesSis.RecordCount > 0 Then
                       lbEncontroAfiliadoEnWebSIS = True
                  End If
              ElseIf txtNroDNIBusqueda.Text <> "" Then
                  '***************************Busca en Pag WEB del SIS x DNI***************************
'                  Set oRsBuscaPacientesSis = mo_ReglasSISgalenhos.ConsultaWebSisBuscarAfiliado(txtNroDNIBusqueda.Text, "", _
'                                                     "", "", "", "", wxParametro323)
                  'FCV 17072015
                  Set oRsBuscaPacientesSis = mo_SisConsumoWeb.WebServiceSISBuscarAfiliado(txtNroDNIBusqueda.Text, "", _
                                                     "", "", "", "", wxParametro323)
                  If oRsBuscaPacientesSis.RecordCount > 0 Then
                       lbEncontroAfiliadoEnWebSIS = True
                  End If
              End If
           End If
           If lbEncontroAfiliadoEnWebSIS = False Then
              If wxParametro322 = "S" And (txtNroDNIBusqueda.Text <> "" Or lnAfiliacionSIS3 <> "") And wxParametro526 <> "S" Then
                 Set oRsBuscaPacientesSis = mo_ReglasSISgalenhos.SisFiltraPacientesAfiliados(lcSql, wxParametroJAMO)
              ElseIf txtApellidoPaternoBusqueda.Text <> "" Or txtApellidoMaternoBusqueda.Text <> "" Then
                 Set oRsBuscaPacientesSis = mo_ReglasSISgalenhos.SisFiltraPacientesAfiliados(lcSql, wxParametroJAMO)
              End If
           End If
           If oRsBuscaPacientesSis.State = 0 Then
              Set oRsBuscaPacientesSis = Nothing
              Exit Sub
           End If
           
           Set grdPacientesEncontrados.DataSource = oRsBuscaPacientesSis.Clone
           With grdPacientesEncontrados
                .Left = 240
                .Top = 1080       'debb-23/02/2015
                .Width = 11700
                .Height = 4455
           End With
           If lbEncontroAfiliadoEnWebSIS = False Then
              If wxParametro322 = "S" And (txtNroDNIBusqueda.Text <> "" Or lnAfiliacionSIS3 <> "") Then
                 grdPacientesEncontrados.Caption = wxParametro312 & "  (Verificar en el AREA DEL SIS SI ESTA AFILIADO, porque se buscó en la WEB SIS y no se encontró al Paciente)"
                 grdPacientesEncontrados.CaptionAppearance.ForeColor = vbRed
              Else
                 grdPacientesEncontrados.Caption = wxParametro312
                 grdPacientesEncontrados.CaptionAppearance.ForeColor = vbBlack
              End If
           Else
              grdPacientesEncontrados.Caption = "Ubicado en WEB SIS"
           End If
           grdPacientesEncontrados.Bands(0).Columns("cAfiliacion").Width = 1700
           grdPacientesEncontrados.Bands(0).Columns("EstadoSis").Width = 500
           grdPacientesEncontrados.Bands(0).Columns("fBAjaOk").Format = sighEntidades.DevuelveFechaSoloFormato_DMY
           Me.grdPacientesEncontrados.Visible = True
           If oRsBuscaPacientesSis.RecordCount = 1 Then
              grdPacientesEncontrados.SetFocus
           End If
           oRsBuscaPacientesSis.Close
       End If
    End If
    Set oRsBuscaPacientesSis = Nothing
End Sub


Private Sub btnBuscarPaciente_Click()
    If Me.chkBuscarEnSIS.Value = 1 Then
       lcCodigoEstablecimientoAdscripcionSIS = ""
       Me.UcSISafiliacion1.TipoFormatoSISvisible False
       Me.UcSISafiliacion1.DevuelveValoresDeFiliacion lnAfiliacionSIS1, lnAfiliacionSIS2, lnAfiliacionSIS3, lcTipoFormatoSIS, lnAfiliacionSIS5
       BuscarPacientesSISSegunFiltro
       Exit Sub
    End If

    Dim RsHistorias As New Recordset
    Dim oDOPaciente As New doPaciente
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    

    lcApP = ""
    lcApM = ""
    lcPnom = ""
    lcSnombreReniec = "": ldFnacimientoReniec = 0: lnIdSexoReniec = 0: lcDireccionReniec = "": mb_UsoWebReniec = False
    
'<(Inicio) Modificado Por: WABG el 18/10/2020-04:53:51 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
    'oDOPaciente.NroHistoriaClinica = Val(HCigualDNI_AgregaNUEVEaLaHistoria(Me.txtNroHistoriaBusqueda.Text))
    oDOPaciente.NroHistoriaClinica = Val(Me.txtNroHistoriaBusqueda.Text)
'</(Fin) Modificado Por: Project Administrator el 18/10/2020-04:53:51 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
    oDOPaciente.ApellidoPaterno = Me.txtApellidoPaternoBusqueda
    oDOPaciente.ApellidoMaterno = Me.txtApellidoMaternoBusqueda
    oDOPaciente.PrimerNombre = Me.txtPrimerNombreBusqueda
    oDOPaciente.SegundoNombre = Me.txtSegundoNombreBusqueda
'    If lnAfiliacionSIS3 = "" Then
       oDOPaciente.IdDocIdentidad = 1
'    Else
'      oDOPaciente.IdDocIdentidad = lnDocumentoTipoSIS
'    End If

    If lnDocumentoTipoSIS = 2 Then          'carnet extranjeria
        oDOPaciente.nrodocumento = ""
        oDOPaciente.IdDocIdentidad = 0
        Me.txtNroDNIBusqueda.Text = ""
    Else
         oDOPaciente.nrodocumento = Me.txtNroDNIBusqueda
    End If
    'oDOPaciente.nrodocumento = Me.txtNroDNIBusqueda
    
    If (oDOPaciente.ApellidoPaterno = "" _
    ) And _
    (Val(Me.txtNroHistoriaBusqueda.Text) = 0) And _
    (oDOPaciente.nrodocumento = "") Then
        MsgBox "Ingrese alguno de los valores de búsqueda", vbInformation, Me.Caption
        Exit Sub
    End If
    If Val(Me.txtNroHistoriaBusqueda.Text) > 0 Then
       Set RsHistorias = mo_AdminAdmision.PacientesFiltraPorHistoriaClinicaDefinitiva(oDOPaciente, oConexion)
    ElseIf Val(Me.txtNroDNIBusqueda.Text) > 0 Then
       Set RsHistorias = mo_AdminAdmision.PacientesFiltraPorNroDocumentoYtipo(oDOPaciente.nrodocumento, oDOPaciente.IdDocIdentidad, oConexion)
    Else
       If chkMuestraHistorial.Value = 1 Then
          Set RsHistorias = mo_AdminAdmision.PacientesFiltrarTodosSoloHistoriasDefinitivas(oDOPaciente, wxSinApellido, oConexion)
       Else
          Set RsHistorias = mo_AdminAdmision.PacientesFiltrarTodosSoloHistoriasDefinit_rap(oDOPaciente, wxSinApellido, oConexion)
       End If
       
    End If
    Set grdPacientesEncontrados.DataSource = RsHistorias
    
    With grdPacientesEncontrados
        .Left = 240
        .Top = 1080                'debb-23/02/2015
        .Width = 11775
        .Height = 4455
        .Caption = ""              'debb-23/02/2015
    End With
    If RsHistorias.RecordCount = 0 Then
        lcApP = txtApellidoPaternoBusqueda
        lcApM = txtApellidoMaternoBusqueda
        lcPnom = txtPrimerNombreBusqueda
        
        Me.grdPacientesEncontrados.Visible = False
        LimpiarFormulario
        
        Me.ucPacientesDetalle1.TipoNumeracion = 0
        Me.ucPacientesDetalle1.NroHistoriaClinica = 0
        
        'txtApellidoMaternoBusqueda = ""
        'txtPrimerNombreBusqueda = ""
        'txtSegundoNombreBusqueda = ""
        txtNroDNIBusqueda = ""
        Me.tabAdmision.Tab = 0
       MsgBox "El Paciente  NO TIENE HISTORIA en el ESTABLECIMIENTO", vbInformation, ""
       Exit Sub
    End If
    'Si hay una sola coincidencia, además se buscó por DNI o HISTORIA
    If RsHistorias.RecordCount = 1 And (txtNroDNIBusqueda.Text <> "" Or txtNroHistoriaBusqueda.Text <> "") Then
        If ProvieneDeEmergencia_o_CE(RsHistorias!idPaciente, oConexion) = False Then         'debb-29/04/2016
            If mo_AdminAdmision.BuscaSiEstaHospitalizado(RsHistorias!idPaciente, oConexion, ml_TipoServicio) = False Then  'debb-05/12/2015
                Me.grdPacientesEncontrados.Visible = False
                RsHistorias.MoveFirst
                
                CargaHCyPaciente RsHistorias!NroHistoriaClinica, RsHistorias!ApellidoPaterno, RsHistorias!ApellidoMaterno, _
                                 RsHistorias!PrimerNombre
                
                chkPacienteNuevo.Value = 0
                Me.ucPacientesDetalle1.LimpiarDatosDePaciente wxParametro211, ldFechaActualServidor
    
                Me.ucPacientesDetalle1.idPaciente = RsHistorias!idPaciente
                Me.ucPacientesDetalle1.CargarDatosDePacienteALosControles oConexion, wxParametro242, wxParametro287
    
                Me.ucPacientesDetalle1.NroHistoriaClinica = RsHistorias!NroHistoriaClinica
                Me.ucPacientesDetalle1.TipoNumeracion = RsHistorias!idTipoNumeracion
                Me.idPaciente = RsHistorias!idPaciente
                
                'yamill palomino
                'Carga diagnosticos de la atencion de emergencia anterior a 24 horas de hospitalizado (Por DNI o HC)
                TraeDiagnosticosHasta24HorasDeEmergencia RsHistorias!idPaciente
    
                Me.tabAdmision.Tab = 0
                DeudasPendientesDeAnterioresAtenciones oConexion
                Me.ucPacientesDetalle1.TabEnNroHistoria
            End If
        End If                                   'debb-29/04/2016
    ElseIf RsHistorias.RecordCount > 0 Then
        Me.grdPacientesEncontrados.Visible = True
        RsHistorias.MoveFirst
        If RsHistorias.RecordCount = 1 Then
           On Error Resume Next
           grdPacientesEncontrados.SetFocus
        End If
        
    ElseIf RsHistorias.RecordCount = 0 Then
        lcApP = txtApellidoPaternoBusqueda
        lcApM = txtApellidoMaternoBusqueda
        lcPnom = txtPrimerNombreBusqueda
        
        Me.grdPacientesEncontrados.Visible = False
        LimpiarFormulario
        
        Me.ucPacientesDetalle1.TipoNumeracion = 0
        Me.ucPacientesDetalle1.NroHistoriaClinica = 0
        
        txtApellidoMaternoBusqueda = ""
        txtPrimerNombreBusqueda = ""
        txtSegundoNombreBusqueda = ""
        txtNroDNIBusqueda = ""
        Me.tabAdmision.Tab = 0
        
    End If
    oConexion.Close
    Set oConexion = Nothing
    
    Set RsHistorias = Nothing
    Set oDOPaciente = Nothing
    
    
End Sub

Sub DeudasPendientesDeAnterioresAtenciones(oConexion As Connection)
        '
        UcPacientesSunasa1.idPaciente = ml_IdPaciente
        UcPacientesSunasa1.CargarDatosDelUltimoSeguroDelPacienteALosControles oConexion
        'Deudas
        ms_MensajeError = mo_AdminFacturacion.DevuelveDeudaPacienteDeAntencionesAnteriores(ml_IdPaciente, oConexion, ml_idCuentaAtencion)
        If ms_MensajeError <> "" Then
            ucMensajeParpadeando1.Visible = True
            ucMensajeParpadeando1.MensajeDeTexto = "Deudas:  " & ms_MensajeError
            'debb-29/02/2016 (inicio)
            If mi_Opcion = sghAgregar And InStr(ms_MensajeError, "<FALLECIDO>") > 0 Then
               btnCancelar_Click
            End If
            'debb-29/02/2016 (fin)
        Else
           '
           ucMensajeParpadeando1.Visible = False
           ucMensajeParpadeando1.MensajeDeTexto = ""
        End If
        ms_MensajeError = ""

End Sub



Private Sub btnBuscarServicios_Click()
    CompletarDatosDeServicio txtIdServicioIngreso, lblNombreServicio, ""
    If txtIdServicioIngreso.Text <> "" And ml_ldFechaEgreso = 0 Then
        ml_idServicioEgreso = Val(txtIdServicioIngreso.Text)
        ml_lcServicioEgreso = lblNombreServicio.Text
    End If
    On Error Resume Next
    tabAdmision.Tab = 1
    If txtNroCamaIngreso.Visible = True Then
       txtNroCamaIngreso.Text = ""
       txtNroCamaIngreso.Tag = ""
    End If
End Sub




'mgaray20140926
Private Sub btnImprimeFichaSIS_Click()
    If mi_Opcion <> sghAgregar Then
        CargaDatosAlObjetosDeDatos
    End If
    If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghAdmisionEmergencia Then
        If ValidarDatosObligatorios = False Or ValidarReglas = False Or lcElServicioUsaGalenHos = "N" Then
            Exit Sub
        End If
    ElseIf mi_Opcion = sghModificar Then
        If ValidarDatosObligatorios = False Or ValidarReglas = False Or lcElServicioUsaGalenHos = "N" Then
            Exit Sub
        End If
    End If
'    Dim oFua As New SIGHSis.clFUA
'    oFua.idCuentaAtencion = mo_Atenciones.idCuentaAtencion
'    oFua.lcNombrePc = mo_lcNombrePc
'    oFua.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
'    oFua.idUsuario = ml_idUsuario
'    oFua.Opcion = mi_Opcion
'    oFua.MostrarFormulario
'    Set oFua = Nothing
    Dim ml_FuaTipoAnexo2015 As Integer
    Dim oFua As New SIGHSis.clFUA
    oFua.idCuentaAtencion = mo_Atenciones.idCuentaAtencion
    oFua.lcNombrePc = mo_lcNombrePc
    oFua.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    oFua.idUsuario = ml_idUsuario
    oFua.Opcion = mi_Opcion
    oFua.IdServicio = Val(Me.txtIdServicioIngreso.Tag)
    oFua.MostrarFormulario
    Set oFua = Nothing
End Sub

Private Sub btnImprimeFiliacion_Click()
    Dim oImprime As New RptHistoriaClinicaCE
    Dim oEdad As Edad
    If mi_Opcion = sghAgregar Then
       oEdad = sighEntidades.CalcularEdad(mo_Pacientes.FechaNacimiento, mo_Historia.fechacreacion)
    Else
       oEdad = sighEntidades.CalcularEdad(mo_Pacientes.FechaNacimiento, Me.ucPacientesDetalle1.DevuelveFechaCreacionHistoria)
    End If
    oImprime.ImprimeEnFormatoDeFiliacionParaHistoriaClinica mo_Atenciones.idPaciente, oEdad.Edad, oEdad.TipoEdad, Me.hwnd
    Set oImprime = Nothing
    Me.Visible = False
End Sub

Private Sub btnImprimir_Click()
  On Error GoTo ErrorI
  
  If Me.idAtencion = 0 Then
    MsgBox "Debe agregar la atención para poder imprimir", vbInformation, Me.Caption
    Exit Sub
  End If
  Dim lcDxIng As String
  Dim rsDX As Recordset
  Dim oConexion As New Connection
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open sighEntidades.CadenaConexion
  
  Set rsDX = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(Me.idAtencion, sghHospitalizacionIngreso, oConexion)
  lcDxIng = ""
  If rsDX.RecordCount > 0 Then lcDxIng = "Dx.Ing: (" & rsDX!CodigoCIE2004 & ") " & rsDX.Fields!descripcion
  Set rsDX = Nothing
  Select Case ml_TipoServicio
    Case sghConsultaExterna
      '
    Case sghHospitalizacion
      Dim oRptHistoriaClinicaHosp As New SIGHReportes.clHistoriaClinicaHosp
      oRptHistoriaClinicaHosp.idAtencion = Me.idAtencion
      oRptHistoriaClinicaHosp.idCuentaAtencion = Me.txtNroCuenta.Text
      oRptHistoriaClinicaHosp.CrearReporteHistoriaClinicaDeLaAtencion cmbFuenteFinanciamiento.Text, _
                              lblNombreOrigenReferencia.Text, lcDxIng, txtNroCamaIngreso.Text, ml_idUsuario, _
                              Me.hwnd, Mid(cmbIdTipoEdad.Text, InStr(cmbIdTipoEdad.Text, "=") + 1)
      Set oRptHistoriaClinicaHosp = Nothing
      Me.Visible = False
    Case sghEmergenciaObservacion, sghEmergenciaConsultorios
      Dim oRptHistoriaEmerg As New SIGHReportes.clHistoriaConsEmerg
      oRptHistoriaEmerg.idAtencion = Me.idAtencion
      oRptHistoriaEmerg.idCuentaAtencion = Me.txtNroCuenta.Text
      oRptHistoriaEmerg.Plan = cmbFuenteFinanciamiento.Text
      'debb-30/03/2016
      oRptHistoriaEmerg.CrearReporteHistoriaClinicaConsultorioEmerg Me.hwnd, Mid(cmbIdTipoEdad.Text, _
                                                                InStr(cmbIdTipoEdad.Text, "=") + 1), _
                                                                cmbIdTipoGravedad.Text, _
                                 Mid(lcBuscaParametro.RetornaFechaHoraServidorSQL, 7, 4) & "-" & txtEmergenciaN.Text, _
                                txtIdEstablecimientoOrigen.Text & "-" & lblNombreOrigenReferencia.Text, _
                                txtDNIacompaniante.Text, 2, txtNroAfiliacionSis.Text, "FORMA DE INGRESO: " & cmbComoLlego.Text, _
                                cmbComoLlego.ListIndex, cmbTipoAtencion.ListIndex, cmbIdTipoGravedad.ListIndex
                                                                
      Set oRptHistoriaEmerg = Nothing
      Me.Visible = False
  End Select
  oConexion.Close
  Set oConexion = Nothing
  Set rsDX = Nothing
  Exit Sub
    
ErrorI:
  MsgBox "Error Número: " & Err.Number & Chr(13) & "Descripción: " & Err.Description
End Sub

Private Sub btnLimpiar_Click()
    txtNroDNIBusqueda.Text = ""
    txtNroHistoriaBusqueda.Text = ""
    txtApellidoPaternoBusqueda.Text = ""
    txtApellidoMaternoBusqueda.Text = ""
    txtPrimerNombreBusqueda.Text = ""
    txtSegundoNombreBusqueda.Text = ""
    UcSISafiliacion1.Limpiar
    Me.grdPacientesEncontrados.Visible = False
    lbProcedeDeEmergencia = False: lbProcedeDeConsExt = False     'debb-23/02/2015
    fraBusqueda.Enabled = True: btnBuscarPaciente.Enabled = True  'debb-23/02/2015
    On Error Resume Next
    txtNroDNIBusqueda.SetFocus
End Sub






Private Sub btnMedicoRespNacimiento_Click()
    'Buscará Médicos por Especialidad y no los Programados, porque no se sabe la Fecha/hora Nacimiento
    CompletarDatosDeMedico Me.txtIdMedicoNacimiento, Me.lblNombreMedicoNacimiento, 0, "", CDate(txtFechaIngreso.Text), txtHoraIngreso.Text, 0
End Sub

'AYañez 06-11-2014 ************************
Private Sub btnNuevoAdmisionHospDetalle_Click()
    Me.txtNroDNIBusqueda = ""
    Me.btnAceptar.Enabled = True
    Me.btnNuevoAdmisionHospDetalle.Visible = False 'A.Yañez 10/11/2014
    Me.tabAdmision.Tab = 0
    Me.txtApellidoPaternoBusqueda.SetFocus
    LimpiarFormulario
    ucDiagnosticosIngreso.LimpiarDatos
End Sub

Private Sub btnPreCuenta_Click()
   If txtNroCuenta.Text <> "" Then
      ImprimePreCuenta
      Me.Visible = False
   End If
End Sub


Private Sub btnQuitarCpt_Click()
    Dim oCpt As New FacOrdenServicioDetalle
    oCpt.FormMostradoDesde = 1
    oCpt.lbNOValidaCodigoPrestacion = True
    oCpt.PuntoCarga = 1   'consumo en el servicio
    oCpt.Opcion = sghEliminar
    oCpt.IdOrden = ml_idOrden
    oCpt.idUsuario = ml_idUsuario
    oCpt.idCuentaAtencion = ml_idCuentaAtencion
    oCpt.Show 1
    Set oCpt = Nothing
    CargaCPTrealizadosEnElServicio
End Sub

Private Sub btnQuitarMadre_Click()
    lblMadre.Text = ""
    lnIdNacimientoSeleccionado = 0
End Sub

Private Sub btnVerDisponibilidadDeCamas_Click()
Dim oBusqueda As New CamasBusqueda
Dim oDOCama As New DOCama
Dim oConexion As New Connection
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    
    oBusqueda.idTipoServicio = Val(mo_cmbIdTipoServicio.BoundText)
    oBusqueda.IdServicio = Val(Me.txtIdServicioIngreso.Tag)
    
    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
       CargaCamaSeleccionada (oBusqueda.idRegistroSeleccionado)
        Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOCama Is Nothing Then
            If oDOCama.idPaciente = mo_Atenciones.idPaciente Or oDOCama.idPaciente = 0 Then
                Me.txtNroCamaIngreso.Text = oDOCama.Codigo
                Me.txtNroCamaIngreso.Tag = oDOCama.idCama
                If Me.txtIdServicioIngreso.Tag = ml_idServicioEgreso Then
                   ml_lcCamaEgreso = oDOCama.Codigo
                   ml_idCamaEgreso = oDOCama.idCama
                ElseIf mi_Opcion = sghAgregar And Val(mo_cmbIdTipoServicio.BoundText) = sghEmergenciaConsultorios And mb_EsObservacionEmergencia = True Then               '09/08/2011
                   ml_lcCamaEgreso = oDOCama.Codigo
                   ml_idCamaEgreso = oDOCama.idCama
                End If
                sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "AdmCama: " & oDOCama.Codigo
            Else
                MsgBox "La cama seleccionada no puede usarla", vbInformation, Me.Caption
                Me.txtNroCamaIngreso.Text = ""
                Me.txtNroCamaIngreso.Tag = ""
            End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDOCama = Nothing
End Sub

Sub CargaCamaSeleccionada(idCama As Long)
        Dim oDOCama As New DOCama
        Dim oConexion As New Connection
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorId(idCama, oConexion)
        If Not oDOCama Is Nothing Then
            If Val(Me.txtIdServicioIngreso.Tag) = oDOCama.IdServicioUbicacionActual Then
                Me.txtNroCamaIngreso.Text = oDOCama.Codigo
                Me.txtNroCamaIngreso.Tag = oDOCama.idCama
            Else
                MsgBox "La cama seleccionada no pertenece al mismo servicio de ingreso", vbInformation, Me.Caption
                Me.txtNroCamaIngreso.Text = ""
                Me.txtNroCamaIngreso.Tag = ""
            End If
        End If
        oConexion.Close
        Set oConexion = Nothing
        Set oDOCama = Nothing
        Set oConexion = Nothing
End Sub

Private Sub btnVerDisponibilidadDeCamas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        btnVerDisponibilidadDeCamas_Click
    Else
        AdministrarKeyPreview KeyCode
    End If
End Sub



Private Sub chkPacienteNuevo_Click()
    
    If chkPacienteNuevo.Value = 1 Then
        ucPacientesDetalle1.MarcoCheckPacienteNuevo = True
        mo_Formulario.HabilitarDeshabilitar chkBuscarEnSIS, False 'A.Yañez 13112014
        '
        LimpiarFormulario
        
        grdPacientesEncontrados.Visible = False
        
        txtApellidoPaternoBusqueda = ""
        txtApellidoMaternoBusqueda = ""
        txtPrimerNombreBusqueda = ""
        txtSegundoNombreBusqueda = ""
        txtNroHistoriaBusqueda.Text = ""
        txtNroDNIBusqueda = ""
        Me.tabAdmision.Tab = 0
        '
'<(Inicio) Añadido Por: WABG el: 27/10/2020-08:54:54 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
        Me.ucPacientesDetalle1.HabilitarControlesDeTextoRENIEC
'</(Fin) Añadido Por: WABG el: 27/10/2020-08:54:54 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
        
        Me.ucPacientesDetalle1.ConfigurarValoresPorDefecto
        '
        If lbBuscaDNIenReniec = True And Len(txtNroDNIBusqueda.Text) = 8 Then
           mo_Reniec.ConsultarDNIenReniec txtNroDNIBusqueda.Text
           If mo_Reniec.ApellidoPaterno <> "" Then
                 lcApP = mo_Reniec.ApellidoPaterno
                 lcApM = mo_Reniec.ApellidoMaterno
                 lcPnom = mo_Reniec.PrimerNombre
                 lcSnombreReniec = mo_Reniec.SegundoNombre
                 ldFnacimientoReniec = mo_Reniec.FechaNacimiento
                 lnIdSexoReniec = mo_Reniec.idTipoSexo
                 lcDireccionReniec = mo_Reniec.DireccionDomicilio
                 mb_UsoWebReniec = True
           End If
        End If
        '
        If txtNroDNIBusqueda.Text = "" Then
           txtNroDNIBusqueda.Text = lcDniSIS
        End If
        '
        If ml_TipoServicio = sghHospitalizacion Then
           Me.ucPacientesDetalle1.CargaDatosBasicosPacienteNuevo UCase(lcApP), UCase(lcApM), UCase(lcPnom), wxParametro212, lcSnombreReniec, ldFnacimientoReniec, lnIdSexoReniec, lcDireccionReniec, mb_UsoWebReniec, txtNroDNIBusqueda.Text, lcSnombreSIS, lnIdDistritoSIS, lnIdSexoSIS, ldFechaNacimientoSIS

        Else
           Me.ucPacientesDetalle1.CargaDatosBasicosPacienteNuevo UCase(lcApP), UCase(lcApM), UCase(lcPnom), wxParametro210, lcSnombreReniec, ldFnacimientoReniec, lnIdSexoReniec, lcDireccionReniec, mb_UsoWebReniec, txtNroDNIBusqueda.Text, lcSnombreSIS, lnIdDistritoSIS, lnIdSexoSIS, ldFechaNacimientoSIS
        End If
        txtNroDNIBusqueda.Text = ""
        '
        UcSISafiliacion1.InabilitaControles False
        If lnIdPlanSIS > 0 Then
             cmbFuenteFinanciamiento.BoundText = lnIdPlanSIS
             cmbFuenteFinanciamiento_Click 1
        End If
        '
        UcPacientesSunasa1.YaNoTieneSeguro
        '
        '
        Me.ucPacientesDetalle1.SetFocusEnDNI
    Else
        ucPacientesDetalle1.MarcoCheckPacienteNuevo = False
        UcSISafiliacion1.InabilitaControles True
        mo_Formulario.HabilitarDeshabilitar chkBuscarEnSIS, True 'A.Yañez 13112014
        On Error Resume Next
        Me.txtApellidoPaternoBusqueda.SetFocus
    End If

    mo_Formulario.HabilitarDeshabilitar Me.txtNroHistoriaBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtApellidoPaternoBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtApellidoMaternoBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtPrimerNombreBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtSegundoNombreBusqueda, Not (chkPacienteNuevo.Value = 1)
    'mo_Formulario.HabilitarDeshabilitar Me.cmbNroHistoriaBusqueda, Not (chkPacienteNuevo.Value = 1)
    mo_Formulario.HabilitarDeshabilitar Me.txtNroDNIBusqueda, Not (chkPacienteNuevo.Value = 1)


End Sub





Private Sub cmbComoLlego_Click()
lbHuboCambioEnDato = True
End Sub

Private Sub cmbComoLlego_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbComoLlego
    AdministrarKeyPreview KeyCode

End Sub



Private Sub cmbComoLlego_LostFocus()
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbComoLlego.Text
          lbHuboCambioEnDato = False
        End If
End Sub

Private Sub cmbEstadoLlegada_Click()
lbHuboCambioEnDato = True
End Sub

Private Sub cmbEstadoLlegada_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbEstadoLlegada
    AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbEstadoLlegada_LostFocus()
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbEstadoLlegada.Text
          lbHuboCambioEnDato = False
        End If
End Sub

Private Sub cmbFormaPago_Click(Area As Integer)
lbHuboCambioEnDato = True
End Sub

Private Sub cmbFormaPago_KeyDown(KeyCode As Integer, Shift As Integer)
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbFormaPago_LostFocus()
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbFormaPago.Text
          lbHuboCambioEnDato = False
        End If
End Sub

Private Sub cmbFuenteFinanciamiento_Click(Area As Integer)
        lbHuboCambioEnDato = True
        
        Dim oConexion As New Connection
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set oRsFormaPago = mo_AdminFacturacion.TiposFinanciamientosTarifaSeleccionarPorPlan(Val(cmbFuenteFinanciamiento.BoundText))
        Set cmbFormaPago.RowSource = oRsFormaPago
        cmbFormaPago.ListField = "Descripcion"
        cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
        mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, True
        If oRsFormaPago.RecordCount = 1 Then
           cmbFormaPago.BoundText = oRsFormaPago.Fields!idTipoFinanciamiento
        ElseIf Val(cmbFuenteFinanciamiento.BoundText) = 5 Then
           cmbFormaPago.BoundText = wxParametro259
        End If
        '
        Me.UcPacientesSunasa1.HabilitaFrame True
        Me.UcPacientesSunasa1.YaNoTieneSeguro
        If mo_AdminFacturacion.TiposFinanciamientoGeneraReciboPago(Val(cmbFormaPago.BoundText), oConexion) = True Then
           Me.UcPacientesSunasa1.HabilitaFrame False
        Else
           Me.UcPacientesSunasa1.idPaciente = ml_IdPaciente
           Me.UcPacientesSunasa1.CargarDatosDelUltimoSeguroDelPacienteALosControles oConexion
        End If
        UcPacientesSunasa1.idTipoFinanciamiento = Val(cmbFormaPago.BoundText)
        oConexion.Close
        Set oConexion = Nothing
        'mgaray20140926
'        If wxParametro302 = "S" And lbElServicioRegistraFUA = "S" Then
'            wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
'            InicializarFUA
'        End If
        Call BuscarDatosServicioYAsignarVariablesFUA(Val(txtIdServicioIngreso.Tag))
        If wxParametro302 = "N" Then
           Me.ucSISfuaCodPrestacion1.Visible = False
        End If
        If UcSISafiliacion1.Visible = True Then
            HaceVisibleOnoBotonFUA
            If lbElServicioRegistraFUA = "S" And cmbFuenteFinanciamiento.Locked = False And Val(cmbFuenteFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS And lnIdPlanSIS = 0 Then
               Dim lcDNI As String, lbPreguntar As Boolean
               'mgaray20140926
               If ucPacientesDetalle1.DevuelveFechaNacimiento <> sighEntidades.FECHA_VACIA_DMY Then
                    If mo_ReglasSISgalenhos.PacienteBuscadoEnTablaGalenHosTieneAfiliacionSIS(ucPacientesDetalle1.DevuelveDNI, _
                                                 ucPacientesDetalle1.DevuelveApaterno, ucPacientesDetalle1.DevuelveAmaterno, _
                                                 ucPacientesDetalle1.DevuelvePnombre, ucPacientesDetalle1.DevuelveSnombre, _
                                                 ucPacientesDetalle1.DevuelveSexo, ucPacientesDetalle1.DevuelveFechaNacimiento, _
                                                 wxParametroJAMO, ldFechaActualServidor, lnAfiliacionSIS4, lcSIScodigo, True) = False Then
                           cmbFuenteFinanciamiento.BoundText = ""
                           cmbFormaPago.BoundText = ""
                    End If
               End If
               On Error Resume Next
               ucDiagnosticosIngreso.SetFocus
            End If
        End If
        
        'debb-04/07/2016
        lblNafiliacionSIS.Visible = False
        txtNroAfiliacionSis.Visible = False
        If ml_TipoServicio = sghEmergenciaConsultorios And _
           Val(cmbFuenteFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS And _
                                                                wxParametro302 <> "S" Then
           lblNafiliacionSIS.Visible = True
           txtNroAfiliacionSis.Visible = True
        End If
        '
End Sub

Sub HaceVisibleOnoBotonFUA()
    'mgaray20140926
    btnImprimeFichaSIS.Visible = False
        
    If wxParametro302 = "S" Then
        ucSISfuaCodPrestacion1.Visible = False
        Me.ucSISfuaCodPrestacion1.CodigoPrestacion = ""
        If Val(cmbFuenteFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS Then
            'mgaray20140926
            If UCase(lbElServicioRegistraFUA) = "S" And ml_TipoServicio = sghTipoServicio.sghEmergenciaConsultorios And lcElServicioUsaGalenHos = "S" And mi_Opcion <> sghAgregar Then
                btnImprimeFichaSIS.Visible = True
            End If
        
           ucSISfuaCodPrestacion1.Visible = True
           Dim lcSexo As String, ml_edad_En_YYYYMMDD As String
           'mgaray20140926
           If sighEntidades.EsFecha(Me.ucPacientesDetalle1.DevuelveFechaNacimiento, "DD/MM/AAAA") = True Then
                ml_edad_En_YYYYMMDD = sighEntidades.EdadActualEnFormatoYYYYMMDD(CDate(Format(Me.ucPacientesDetalle1.DevuelveFechaNacimiento, "dd/mm/yyyy hh:mm")), CDate(Format(txtFechaIngreso.Text & " " & txtHoraIngreso.Text, "dd/mm/yyyy hh:mm")))
           End If
           lcSexo = IIf(Left(Me.ucPacientesDetalle1.DevuelveSexo, 1) = 1, "M", "F")
           Me.ucSISfuaCodPrestacion1.ReglasDeConsistenciasAntesDeCargarFormulario ml_TipoServicio, lcSexo, ml_edad_En_YYYYMMDD
           If mi_Opcion <> sghAgregar And mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion <> "" Then
              ucSISfuaCodPrestacion1.CodigoPrestacion = mo_DoAtencionDatosAdicionales.FuaCodigoPrestacion
           End If
        End If
    End If
End Sub

Private Sub cmbFuenteFinanciamiento_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbFuenteFinanciamiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmbFuenteFinanciamiento_Click 1
       ucDiagnosticosIngreso.SetFocus
    End If
End Sub

Private Sub cmbFuenteFinanciamiento_LostFocus()
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbFuenteFinanciamiento.Text
          lbHuboCambioEnDato = False
        End If
End Sub

Private Sub cmbIdCausaExternaMorbilidad_Click()
lbHuboCambioEnDato = True
Dim lCausaExternaMorbilidad As Long
Dim sMensaje As String

        lCausaExternaMorbilidad = Val(mo_cmbIdCausaExternaMorbilidad.BoundText)
        
        If ml_TipoServicio = sghEmergenciaConsultorios Or ml_TipoServicio = sghEmergenciaObservacion Then
            
            cmbIdLugarEvento.Visible = False
            Select Case lCausaExternaMorbilidad
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9
                cmbIdLugarEvento.Visible = True
                mo_cmbIdLugarEvento.BoundColumn = "IdLugarEvento"
                mo_cmbIdLugarEvento.ListField = "DescripcionLarga"
                Set mo_cmbIdLugarEvento.RowSource = mo_AdminAdmision.EmergenciaLugarEventoSeleccionarTodos()
                sMensaje = sMensaje + mo_AdminAdmision.MensajeError
            End Select
            
            cmbIdTipoEvento.Visible = False
            Select Case lCausaExternaMorbilidad
            Case 1, 2, 3, 5, 6, 7, 8, 9
                cmbIdTipoEvento.Visible = True
                mo_cmbIdTipoEvento.BoundColumn = "IdTipoEvento"
                mo_cmbIdTipoEvento.ListField = "DescripcionLarga"
                Set mo_cmbIdTipoEvento.RowSource = mo_AdminAdmision.EmergenciaTipoEventoSeleccionarTodos()
                sMensaje = sMensaje + mo_AdminAdmision.MensajeError
            End Select
                        
            cmbIdSeguridad.Visible = False
            Select Case lCausaExternaMorbilidad
            Case 2, 3, 5, 6, 7, 8, 9
                cmbIdSeguridad.Visible = True
                mo_cmbIdSeguridad.BoundColumn = "IdSeguridad"
                mo_cmbIdSeguridad.ListField = "DescripcionLarga"
                Set mo_cmbIdSeguridad.RowSource = mo_AdminAdmision.EmergenciaSeguridadSeleccionarTodos()
                sMensaje = sMensaje + mo_AdminAdmision.MensajeError
            End Select
                        
            cmbIdRelacionAgresorVictima.Visible = False
            Select Case lCausaExternaMorbilidad
            Case 1
                cmbIdRelacionAgresorVictima.Visible = True
                mo_cmbIdRelacionAgresorVictima.BoundColumn = "IdRelacionAgresorVictima"
                mo_cmbIdRelacionAgresorVictima.ListField = "DescripcionLarga"
                Set mo_cmbIdRelacionAgresorVictima.RowSource = mo_AdminAdmision.EmergenciaRelacionAgresorVictimaSeleccionarTodos()
                sMensaje = sMensaje + mo_AdminAdmision.MensajeError
            End Select
                        
                        
            cmbIdClaseAccidente.Visible = False
            cmbIdTipoVehiculo.Visible = False
            cmbIdTipoTransporte.Visible = False
            cmbIdUbicacionLesionado.Visible = False
            Select Case lCausaExternaMorbilidad
            Case 2, 3
            
                cmbIdClaseAccidente.Visible = True
                mo_cmbIdClaseAccidente.BoundColumn = "IdClaseAccidente"
                mo_cmbIdClaseAccidente.ListField = "DescripcionLarga"
                Set mo_cmbIdClaseAccidente.RowSource = mo_AdminAdmision.EmergenciaClaseAccidenteSeleccionarTodos()
                sMensaje = sMensaje + mo_AdminAdmision.MensajeError
                
                cmbIdTipoVehiculo.Visible = True
                mo_cmbIdTipoVehiculo.BoundColumn = "IdTipoVehiculo"
                mo_cmbIdTipoVehiculo.ListField = "DescripcionLarga"
                Set mo_cmbIdTipoVehiculo.RowSource = mo_AdminAdmision.EmergenciaTipoVehiculoSeleccionarTodos()
                sMensaje = sMensaje + mo_AdminAdmision.MensajeError
                
                cmbIdTipoTransporte.Visible = True
                mo_cmbIdTipoTransporte.BoundColumn = "IdTipoTransporte"
                mo_cmbIdTipoTransporte.ListField = "DescripcionLarga"
                Set mo_cmbIdTipoTransporte.RowSource = mo_AdminAdmision.EmergenciaTipoTransporteSeleccionarTodos()
                sMensaje = sMensaje + mo_AdminAdmision.MensajeError
                
                cmbIdUbicacionLesionado.Visible = True
                mo_cmbIdUbicacionLesionado.BoundColumn = "IdUbicacionLesionado"
                mo_cmbIdUbicacionLesionado.ListField = "DescripcionLarga"
                Set mo_cmbIdUbicacionLesionado.RowSource = mo_AdminAdmision.EmergenciaUbicacionLesionadoSeleccionarTodos()
                sMensaje = sMensaje + mo_AdminAdmision.MensajeError
                
            End Select
            
            cmbIdGrupoOcupacionalALAB.Visible = False
            cmbIdPosicionLesionadoALAB.Visible = False
            Select Case lCausaExternaMorbilidad
            Case 5
                cmbIdGrupoOcupacionalALAB.Visible = True
                mo_cmbIdGrupoOcupacionalALAB.BoundColumn = "IdGrupoOcupacionalALAB"
                mo_cmbIdGrupoOcupacionalALAB.ListField = "DescripcionLarga"
                Set mo_cmbIdGrupoOcupacionalALAB.RowSource = mo_AdminAdmision.EmergenciaGrupoOcupacionalALABSeleccionarTodos()
                sMensaje = sMensaje + mo_AdminAdmision.MensajeError
                
                cmbIdPosicionLesionadoALAB.Visible = True
                mo_cmbIdPosicionLesionadoALAB.BoundColumn = "IdPosicionLesionadoALAB"
                mo_cmbIdPosicionLesionadoALAB.ListField = "DescripcionLarga"
                Set mo_cmbIdPosicionLesionadoALAB.RowSource = mo_AdminAdmision.EmergenciaPosicionLesionadoSeleccionarTodos()
                sMensaje = sMensaje + mo_AdminAdmision.MensajeError
            End Select
                        
            cmbIdTipoAgenteAGAN.Visible = False
            Select Case lCausaExternaMorbilidad
            Case 6
                cmbIdTipoAgenteAGAN.Visible = True
                mo_cmbIdTipoAgenteAGAN.BoundColumn = "IdTipoAgenteAGAN"
                mo_cmbIdTipoAgenteAGAN.ListField = "DescripcionLarga"
                Set mo_cmbIdTipoAgenteAGAN.RowSource = mo_AdminAdmision.EmergenciaTipoAgenteAGANSeleccionarTodos()
                sMensaje = sMensaje + mo_AdminAdmision.MensajeError
            End Select
                        
        End If
    
End Sub

Private Sub cmbIdCausaExternaMorbilidad_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbIdCausaExternaMorbilidad.Text
      lbHuboCambioEnDato = False
    End If
   
   If cmbIdCausaExternaMorbilidad.Text <> "" Then
       mo_cmbIdCausaExternaMorbilidad.BoundText = Val(Split(cmbIdCausaExternaMorbilidad.Text, " = ")(0))
    Else
        cmbIdCausaExternaMorbilidad_Click
   End If
End Sub

Private Sub cmbIdClaseAccidente_LostFocus()
   If cmbIdClaseAccidente.Text <> "" Then
       mo_cmbIdClaseAccidente.BoundText = Val(Split(cmbIdClaseAccidente.Text, " = ")(0))
   End If

End Sub


Private Sub cmbIdGrupoOcupacionalALAB_LostFocus()
   If cmbIdGrupoOcupacionalALAB.Text <> "" Then
       mo_cmbIdGrupoOcupacionalALAB.BoundText = Val(Split(cmbIdGrupoOcupacionalALAB.Text, " = ")(0))
   End If
End Sub

Private Sub cmbIdLugarEvento_LostFocus()
   If cmbIdLugarEvento.Text <> "" Then
       mo_cmbIdLugarEvento.BoundText = Val(Split(cmbIdLugarEvento.Text, " = ")(0))
   End If

End Sub

Private Sub cmbIdPosicionLesionadoALAB_LostFocus()
   If cmbIdPosicionLesionadoALAB.Text <> "" Then
       mo_cmbIdPosicionLesionadoALAB.BoundText = Val(Split(cmbIdPosicionLesionadoALAB.Text, " = ")(0))
   End If
End Sub

Private Sub cmbIdRelacionAgresorVictima_LostFocus()
   If cmbIdRelacionAgresorVictima.Text <> "" Then
       mo_cmbIdRelacionAgresorVictima.BoundText = Val(Split(cmbIdRelacionAgresorVictima.Text, " = ")(0))
   End If

End Sub

Private Sub cmbIdSeguridad_LostFocus()
   If cmbIdSeguridad.Text <> "" Then
       mo_cmbIdSeguridad.BoundText = Val(Split(cmbIdSeguridad.Text, " = ")(0))
   End If
End Sub

Private Sub cmbIdTipoAgenteAGAN_LostFocus()
   If cmbIdTipoAgenteAGAN.Text <> "" Then
       mo_cmbIdTipoAgenteAGAN.BoundText = Val(Split(cmbIdTipoAgenteAGAN.Text, " = ")(0))
   End If
End Sub

Private Sub cmbIdTipoEdad_LostFocus()
Dim oDOTipoEdad As New DOTipoEdad

   If cmbIdTipoEdad.Text <> "" Then
     Set oDOTipoEdad = mo_AdminServiciosComunes.TiposEdadSeleccionarPorCodigo(Trim(Split(cmbIdTipoEdad.Text, " = ")(0)))
     If oDOTipoEdad.idTipoEdad <> 0 Then
         mo_cmbIdTipoEdad.BoundText = oDOTipoEdad.idTipoEdad
    End If
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoEdad
End Sub

Private Sub cmbIdTipoEvento_LostFocus()
   If cmbIdTipoEvento.Text <> "" Then
       mo_cmbIdTipoEvento.BoundText = Val(Split(cmbIdTipoEvento.Text, " = ")(0))
   End If
End Sub



Private Sub cmbIdTipoGravedad_Click()
lbHuboCambioEnDato = True
End Sub

Private Sub cmbIdTipoTransporte_LostFocus()
   If cmbIdTipoTransporte.Text <> "" Then
       mo_cmbIdTipoTransporte.BoundText = Val(Split(cmbIdTipoTransporte.Text, " = ")(0))
   End If

End Sub

Private Sub cmbIdTipoVehiculo_LostFocus()
   If cmbIdTipoVehiculo.Text <> "" Then
       mo_cmbIdTipoVehiculo.BoundText = Val(Split(cmbIdTipoVehiculo.Text, " = ")(0))
   End If
End Sub

Private Sub cmbIdUbicacionLesionado_LostFocus()
   If cmbIdUbicacionLesionado.Text <> "" Then
       mo_cmbIdUbicacionLesionado.BoundText = Val(Split(cmbIdUbicacionLesionado.Text, " = ")(0))
   End If

End Sub




Private Sub cmbServicioReferenciaO_Click()
lbHuboCambioEnDato = True
End Sub



Private Sub cmbServicioReferenciaO_LostFocus()
        If lbHuboCambioEnDato = True Then
            sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbServicioReferenciaO.Text
            lbHuboCambioEnDato = False
        End If
End Sub



Private Sub cmbTipoAtencion_Click()
lbHuboCambioEnDato = True
End Sub

Private Sub cmbTipoAtencion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoAtencion
    AdministrarKeyPreview KeyCode

End Sub

Private Sub cmbTipoAtencion_LostFocus()
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbTipoAtencion.Text
          lbHuboCambioEnDato = False
        End If
End Sub

Private Sub cmdBuscaMadre_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaMadre
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
       lnIdNacimientoSeleccionado = oBusqueda.IdNacimientoSeleccionado
       lblMadre.Text = mo_AdminAdmision.DevuelveDatosDeLaMadreDelPacienteActual(lnIdNacimientoSeleccionado, Me.ucPacientesDetalle1.idTipoSexo, oConexion)
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
End Sub




Private Sub cmdCamaTransf_Click()
Dim oBusqueda As New CamasBusqueda
Dim oDOCama As New DOCama
Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
     
    oBusqueda.idTipoServicio = Val(mo_cmbIdTipoServicio.BoundText)
    oBusqueda.IdServicio = Val(ml_idServicioEgreso)
    oBusqueda.Show 1
    If oBusqueda.BotonPresionado = sghAceptar Then
       
        Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        If Not oDOCama Is Nothing Then
            If oDOCama.idPaciente = mo_Atenciones.idPaciente Or oDOCama.idPaciente = 0 Then
                Me.txtCamaTransf.Text = oDOCama.Codigo
                Me.txtCamaTransf.Tag = oDOCama.idCama
                ml_lcCamaEgreso = oDOCama.Codigo
                ml_idCamaEgreso = oDOCama.idCama
                
                sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "Transf.Cama: " & oDOCama.Codigo
                
                On Error Resume Next
                Me.btnAceptar.SetFocus
            Else
                MsgBox "La cama seleccionada no puede usarla", vbInformation, Me.Caption
                Me.txtCamaTransf.Text = ""
                Me.txtCamaTransf.Tag = ""
            End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oDOCama = Nothing
    Set oBusqueda = Nothing
End Sub



Private Sub cmdFiliacionCE_Click()
    Dim oRptHistoriaConsultaExterna As New RptHistoriaClinicaCE
    If Me.idAtencion = 0 Then
        MsgBox "De agregar la atención para poder imprimir", vbInformation, Me.Caption
        Exit Sub
    End If
    oRptHistoriaConsultaExterna.idAtencion = Me.idAtencion
    oRptHistoriaConsultaExterna.idCuentaAtencion = Val(txtNroCuenta.Text)
    oRptHistoriaConsultaExterna.IdOrden = 0
    oRptHistoriaConsultaExterna.CrearReporteHistoriaClinicaDeLaAtencionCE Me.hwnd
    Set oRptHistoriaConsultaExterna = Nothing

End Sub

Private Sub cmdSinApellidoMaterno_Click()
    txtApellidoMaternoBusqueda.Text = wxSinApellido

End Sub

Private Sub cmdSinApellidoPaterno_Click()
        txtApellidoPaternoBusqueda.Text = wxSinApellido

End Sub






Private Sub Form_Unload(Cancel As Integer)
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub

Sub grdPacientesEncontradosSIS()
    On Error GoTo errGrdSis
    Dim oRecordset As New Recordset
    Dim rsPaciente As Recordset
   
    Dim lbValidaSiEsAfiliadoActualDelSIS As Boolean
    Dim lcSql As String
    
    Set oRecordset = grdPacientesEncontrados.DataSource
    lnIdDistritoSIS = 0: lnIdSexoSIS = 0: ldFechaNacimientoSIS = 0: lcSnombreSIS = "":  lcDniSIS = ""
    txtNroAfiliacionSis.Text = "" 'debb-04/07/2016
    lnIdPlanSIS = sghFuenteFinanciamiento.sghFFSIS
    lcCodigoEstablecimientoAdscripcionSIS = ""
    If oRecordset.RecordCount > 0 Then
        If mo_ReglasSISgalenhos.Sis_ValidaSiEsAfiliadoActualDelSIS(oRecordset, ldFechaActualServidor, True) = True Then
        
        
            lnAfiliacionSIS1 = oRecordset.Fields!cDisa
            lnAfiliacionSIS2 = oRecordset.Fields!cFormato
            lnAfiliacionSIS3 = oRecordset.Fields!cnumero
            lnAfiliacionSIS4 = oRecordset.Fields!idSiaSis
            
            lnDocumentoTipoSIS = oRecordset!DocumentoTipo

            
            txtNroAfiliacionSis.Text = lnAfiliacionSIS1 & "-" & lnAfiliacionSIS2 & "-" & lnAfiliacionSIS3 'debb-04/07/2016
            lcSIScodigo = oRecordset.Fields!Codigo
            lcDniSIS = IIf(IsNull(oRecordset.Fields!DNI), "", oRecordset.Fields!DNI)
            lcCodigoEstablecimientoAdscripcionSIS = oRecordset.Fields!CodigoEstablAdscripcion
            If Not IsNull(oRecordset.Fields!sNombre) Then
               lcSnombreSIS = oRecordset.Fields!sNombre
            End If
            If Not IsNull(oRecordset.Fields!DistritoDomicilio) Then
               lnIdDistritoSIS = Val(oRecordset.Fields!DistritoDomicilio)
            End If
            If Not IsNull(oRecordset.Fields!Sexo) Then
               lnIdSexoSIS = IIf(oRecordset.Fields!Sexo = "0", 2, 1)
            End If
            If Not IsNull(oRecordset.Fields!FNacimiento) Then
               ldFechaNacimientoSIS = oRecordset.Fields!FNacimiento
            End If
            txtNroDNIBusqueda.Text = IIf(IsNull(oRecordset.Fields!DNI), "", oRecordset.Fields!DNI)
            Me.txtApellidoPaternoBusqueda.Text = oRecordset.Fields!apPaterno
            Me.txtApellidoMaternoBusqueda.Text = oRecordset.Fields!apMaterno
            Me.txtPrimerNombreBusqueda.Text = oRecordset.Fields!Pnombre
            lcApP = oRecordset.Fields!apPaterno
            lcApM = oRecordset.Fields!apMaterno
            lcPnom = oRecordset.Fields!Pnombre
            lcSnombreReniec = lcSnombreSIS
            ldFnacimientoReniec = ldFechaNacimientoSIS
            lnIdSexoReniec = lnIdSexoSIS
            Me.txtSegundoNombreBusqueda.Text = IIf(IsNull(oRecordset.Fields!sNombre), "", oRecordset.Fields!sNombre)
            Me.chkBuscarEnSIS.Value = 0
            btnBuscarPaciente_Click
            Set rsPaciente = Me.grdPacientesEncontrados.DataSource
            If rsPaciente.RecordCount = 1 Then
               grdPacientesEncontrados_DblClick
            Else
               'MsgBox "El Paciente SIS NO TIENE HISTORIA en el ESTABLECIMIENTO", vbInformation, ""
            End If
            If Me.ucPacientesDetalle1.idPaciente > 0 Then
                If lnIdPlanSIS > 0 Then
                     cmbFuenteFinanciamiento.BoundText = lnIdPlanSIS
                     cmbFuenteFinanciamiento_Click 1
                     

'<(Inicio) Añadido Por: WABG el: 26/01/2021-11:54:25 a.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
                     Me.ucPacientesDetalle1.SetFocusEnHistoria
'</(Fin) Añadido Por: WABG el: 26/01/2021-11:54:25 a.m. en el Equipo: SISGALENPLUS-PC<CAMBIO-37>

                End If
            End If
        End If
    End If
    Set oRecordset = Nothing
    Set rsPaciente = Nothing
    
errGrdSis:
End Sub


'debb-29/04/2016
Function ProvieneDeEmergencia_o_CE(lnIdPaciente As Long, oConexion As Connection) As Boolean
    ProvieneDeEmergencia_o_CE = False
    If ml_TipoServicio = sghHospitalizacion And fraBusqueda.Enabled = True Then
       Dim oRecordset As New Recordset
       Set oRecordset = mo_AdminAdmision.AtencionesXidPaciente(lnIdPaciente, oConexion)
       If oRecordset.RecordCount > 0 Then
          oRecordset.MoveFirst
          If oRecordset!idTipoServicio = sghTipoServicio.sghConsultaExterna Then
             If Not IsNull(oRecordset!IdDestinoAtencion) Then
                If oRecordset!IdDestinoAtencion = 11 Then
                   MsgBox "El Paciente proviene de Consultorio Externo" & Chr(13) & Chr(13) & _
                          "Debe usar el botón 'ConstExt'", vbInformation, ""
                   ProvieneDeEmergencia_o_CE = True
                End If
             End If
          ElseIf oRecordset!idTipoServicio = sghTipoServicio.sghEmergenciaConsultorios Then
             If Not IsNull(oRecordset!IdDestinoAtencion) Then
                If oRecordset!IdDestinoAtencion = 21 Then
                   MsgBox "El Paciente proviene de Emergencia" & Chr(13) & Chr(13) & _
                          "Debe usar el botón 'Emergencia'", vbInformation, ""
                   ProvieneDeEmergencia_o_CE = True
                End If
             End If
           End If
       End If
       oRecordset.Close
       Set oRecordset = Nothing
    End If
End Function


Private Sub grdPacientesEncontrados_DblClick()
    If Me.chkBuscarEnSIS.Value = 1 Then
       grdPacientesEncontradosSIS
       Exit Sub
    End If
    '
    Dim rsPaciente As Recordset
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    On Error Resume Next
    Set rsPaciente = Me.grdPacientesEncontrados.DataSource

    CargaHCyPaciente rsPaciente!NroHistoriaClinica, rsPaciente!ApellidoPaterno, rsPaciente!ApellidoMaterno, _
                     rsPaciente!PrimerNombre

    If mo_AdminAdmision.BuscaSiEstaHospitalizado(rsPaciente!idPaciente, oConexion, ml_TipoServicio) = True Then   'debb-05/12/2015
       Exit Sub
    End If
    '
    If ProvieneDeEmergencia_o_CE(rsPaciente!idPaciente, oConexion) = True Then   'debb-29/04/2016
       Exit Sub                                 'debb-29/04/2016
    End If                                      'debb-29/04/2016
    '
    
    Me.ucPacientesDetalle1.LimpiarDatosDePaciente wxParametro211, ldFechaActualServidor
    Me.ucPacientesDetalle1.TipoNumeracion = rsPaciente!idTipoNumeracion
    Me.ucPacientesDetalle1.NroHistoriaClinica = rsPaciente!NroHistoriaClnica
    Me.ucPacientesDetalle1.idPaciente = rsPaciente!idPaciente
    Me.idPaciente = rsPaciente!idPaciente
    Me.ucPacientesDetalle1.CargarDatosDePacienteALosControles oConexion, wxParametro242, wxParametro287
    chkPacienteNuevo.Value = 0
    Me.tabAdmision.Tab = 0
    Me.grdPacientesEncontrados.Visible = False: DoEvents
    '
    DeudasPendientesDeAnterioresAtenciones oConexion
    Me.ucPacientesDetalle1.TabEnNroHistoria
    
    'yamill palomino
    'Carga diagnosticos de la atencion de emergencia anterior a 24 horas de hospitalizado (Busqueda por apellidos )
    TraeDiagnosticosHasta24HorasDeEmergencia rsPaciente!idPaciente
    'debb-23/02/2015 - inicio
    ml_idAtencionEmeg_CE = 0
    If lbProcedeDeConsExt = True Or lbProcedeDeEmergencia = True Then
        If rsPaciente!idServicioDestino > 0 Then
           If Not IsNull(rsPaciente!idServicioDestino) Then
                Dim oDoServicio As New doServicio
                Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(rsPaciente!idServicioDestino, _
                                                                                  oConexion)
                If Not oDoServicio Is Nothing Then
                     Me.txtIdServicioIngreso.Tag = oDoServicio.IdServicio
                     Me.txtIdServicioIngreso.Text = oDoServicio.Codigo
                     Me.lblNombreServicio = oDoServicio.nombre
                     Me.lblNombreServicio.Tag = oDoServicio.IdEspecialidad
                     mo_cmbIdTipoServicio.BoundText = oDoServicio.idTipoServicio
                     Me.IdEspecialidad = oDoServicio.IdEspecialidad
                End If
           End If
           Set oDoServicio = Nothing
        End If
        ml_idAtencionEmeg_CE = rsPaciente!idAtencion
        mo_cmbIdViasAdmision.BoundText = IIf(lbProcedeDeConsExt = True, "30", "31")
        If Not IsNull(rsPaciente!IdFuenteFinanciamiento) Then
           cmbFuenteFinanciamiento.BoundText = IIf(rsPaciente!IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente, _
                                                   sghFuenteFinanciamiento.sghFFParticularHospitalizado, _
                                                   Trim(Str(rsPaciente!IdFuenteFinanciamiento)))
        End If
        cmbFormaPago.BoundText = Trim(Str(rsPaciente!IdFormaPago))
        tabAdmision.Tab = 1
        TabIngreso.Tab = 0
        'debb-14/03/2015 (inicio)
        If lb_puedeCambiarFuenteFinanciamiento = False Then
           mo_Formulario.HabilitarDeshabilitar cmbFuenteFinanciamiento, False
           mo_Formulario.HabilitarDeshabilitar cmbFormaPago, False
        End If
        'debb-14/03/2015 (fin)
        If lblNombreServicio.Text = "" Then
           lblNombreServicio.SetFocus
        Else
           txtHoraIngreso.SetFocus
        End If
    End If
    'debb-24/02/2015 - fin
    oConexion.Close
    Set oConexion = Nothing
    Set rsPaciente = Nothing
    '


End Sub

Private Sub grdPacientesEncontrados_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    On Error Resume Next
    
    
    
    grdPacientesEncontrados.Bands(0).Columns("IdPaciente").Hidden = True
    grdPacientesEncontrados.Bands(0).Columns("IdTipoNumeracion").Hidden = True
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Hidden = True
    
    grdPacientesEncontrados.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "NroHistoria"
    grdPacientesEncontrados.Bands(0).Columns("NroHistoriaClinica").Width = 1000
    
    grdPacientesEncontrados.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdPacientesEncontrados.Bands(0).Columns("ApellidoPaterno").Width = 1200
    
    grdPacientesEncontrados.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdPacientesEncontrados.Bands(0).Columns("ApellidoMaterno").Width = 1200
    
    grdPacientesEncontrados.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdPacientesEncontrados.Bands(0).Columns("PrimerNombre").Width = 1200

    grdPacientesEncontrados.Bands(0).Columns("SegundoNombre").Header.Caption = "2do Nombre"
    grdPacientesEncontrados.Bands(0).Columns("SegundoNombre").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("FechaNacimiento").Header.Caption = "Fecha Nac."
    grdPacientesEncontrados.Bands(0).Columns("FechaNacimiento").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Header.Caption = "Tipo Numeración"
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").Width = 1500
    grdPacientesEncontrados.Bands(0).Columns("TipoNumeracion").CellAppearance.TextAlign = ssAlignRight

    grdPacientesEncontrados.Bands(0).Columns("TipoServicio").Header.Caption = "Ult. Tipo Serv."
    grdPacientesEncontrados.Bands(0).Columns("TipoServicio").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("FechaIngreso").Header.Caption = "Ult.Fec.Ing"
    grdPacientesEncontrados.Bands(0).Columns("FechaIngreso").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("FechaEgreso").Header.Caption = "Ult.Fec.Egr."
    grdPacientesEncontrados.Bands(0).Columns("FechaEgreso").Width = 1000

    grdPacientesEncontrados.Bands(0).Columns("ServicioIngreso").Header.Caption = "Ult. Serv. Ing."
    grdPacientesEncontrados.Bands(0).Columns("ServicioIngreso").Width = 2700
End Sub

Private Sub grdPacientesEncontrados_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
        grdPacientesEncontrados_DblClick
    End If
End Sub










Private Sub lblNombreMedico_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub lblNombreMedico_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, lblNombreMedico
   AdministrarKeyPreview KeyCode

End Sub

Private Sub lblNombreMedico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lbUltimaTeclaPulsoENTER = True
    Else
        lbUltimaTeclaPulsoENTER = False
    End If
End Sub

Private Sub lblNombreMedico_LostFocus()

        
        If lblNombreMedico.Locked = False And lbUltimaTeclaPulsoENTER = True Then
           lbUltimaTeclaPulsoENTER = False
           CompletarDatosDeMedico txtIdMedicoIngreso, lblNombreMedico, Val(Me.lblNombreServicio.Tag), lblNombreMedico.Text, CDate(Me.txtFechaIngreso.Text), Me.txtHoraIngreso.Text, ml_TipoServicio
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, lblNombreMedico.Text
          lbHuboCambioEnDato = False
        End If
           On Error Resume Next
           If ml_TipoServicio = sghHospitalizacion Then
              cmbFuenteFinanciamiento.SetFocus
           Else
              cmbComoLlego.SetFocus
           End If
        End If
End Sub











Private Sub lblNombreServicio_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub lblNombreServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, lblNombreServicio
   AdministrarKeyPreview KeyCode
End Sub

Private Sub lblNombreServicio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lbUltimaTeclaPulsoENTER = True
    Else
        lbUltimaTeclaPulsoENTER = False
    End If
End Sub

Private Sub lblNombreServicio_LostFocus()
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, lblNombreServicio.Text
          lbHuboCambioEnDato = False
        End If
        
        If lblNombreServicio.Locked = False And lbUltimaTeclaPulsoENTER = True Then
            lbUltimaTeclaPulsoENTER = False
            CompletarDatosDeServicio txtIdServicioIngreso, lblNombreServicio, lblNombreServicio.Text
        End If
End Sub






Private Sub optProCE_Click(Value As Integer)
    If lbProcedeDeConsExt = True Then
       chkPacienteNuevo.Value = 0
       chkBuscarEnSIS.Value = 0
       AtencionesSinAdmHospitalizacion False
    End If
End Sub
'debb-23/02/2015
Private Sub optProEmer_Click(Value As Integer)
    If lbProcedeDeEmergencia = True Then
       chkPacienteNuevo.Value = 0
       chkBuscarEnSIS.Value = 0
       AtencionesSinAdmHospitalizacion True
    End If
End Sub



Private Sub tabAdmision_Click(PreviousTab As Integer)
   Select Case tabAdmision.Tab
    Case 1
       lcCaptionTab2 = TabIngreso.Caption
       Select Case TabIngreso.Tab
       Case 0
          Me.ucDiagnosticosIngreso.SexoPaciente = Val(Left(Me.ucPacientesDetalle1.DevuelveSexo, 1))
          On Error Resume Next
          Me.cmbIdViasAdmision.SetFocus
       Case 1
       Case 2
       End Select
       LimpiaDatosDeBusqueda
    End Select
End Sub





Private Sub TabIngreso_Click(PreviousTab As Integer)
   Select Case TabIngreso.Tab
   Case 3
       Me.UcPacientesSunasa1.DatosDeCabecera Me.ucPacientesDetalle1.DevuelvePaciente, Me.ucPacientesDetalle1.DevuelveSexo, Me.ucPacientesDetalle1.DevuelveDocumento, Me.ucPacientesDetalle1.DevuelveNroDocumento, Me.ucPacientesDetalle1.DevuelvePaisDomicilio, Me.ucPacientesDetalle1.DevuelveFechaNacimiento, Me.ucPacientesDetalle1.DevuelveUbigeoDomicilio
   End Select
End Sub


Private Sub TxtCitaTratamiento_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub TxtCitaTratamiento_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, TxtCitaTratamiento.Text
      lbHuboCambioEnDato = False
    End If
End Sub

Private Sub txtDNIacompaniante_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub txtDNIacompaniante_LostFocus()
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtDNIacompaniante.Text
          lbHuboCambioEnDato = False
        End If
End Sub

Private Sub txtEmergenciaN_Change()
lbHuboCambioEnDato = True
End Sub

'debb-21/07/2016
Private Sub txtEmergenciaN_LostFocus()
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtEmergenciaN.Text
          lbHuboCambioEnDato = False
        End If
    If txtEmergenciaN.Text <> "" And mi_Opcion = sghAgregar And wxParametro506 = "S" Then
       If BuscarSiExisteNrocorrelativoEmergencia = True Then
       End If
    End If
End Sub
'debb-21/07/2016
Function BuscarSiExisteNrocorrelativoEmergencia() As Boolean
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim oConexion As New Connection
        Dim lnIdAtencion As Long, lcNroEmergencia3 As String
        lcNroEmergencia3 = Mid(lcBuscaParametro.RetornaFechaHoraServidorSQL, 7, 4) & Trim(txtEmergenciaN.Text)
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        Set oRsTmp1 = mo_AdminAdmision.atencionesDatosAdicionalesXfiltro(" dbo.AtencionesDatosAdicionales.emergenciaCorrelativo='" & lcNroEmergencia3 & "'", oConexion)
        If oRsTmp1.RecordCount > 0 Then
           If mi_Opcion = sghAgregar Then
              Set oRsTmp2 = mo_AdminAdmision.AtencionesSeleccionarPorIdAtencion(oRsTmp1!idAtencion)
              If oRsTmp2!IdEstadoAtencion <> 0 Then
                    MsgBox "Ese N° Emergencia YA EXISTE", vbInformation, Me.Caption
                    BuscarSiExisteNrocorrelativoEmergencia = True
              End If
              oRsTmp2.Close
           ElseIf mi_Opcion = sghModificar Then
              oRsTmp1.MoveFirst
              Do While Not oRsTmp1.EOF
                 If oRsTmp1!idAtencion <> ml_idAtencion Then
                    Set oRsTmp2 = mo_AdminAdmision.AtencionesSeleccionarPorIdAtencion(oRsTmp1!idAtencion)
                    If oRsTmp2!IdEstadoAtencion <> 0 Then
                        MsgBox "Ese N° Emergencia YA EXISTE", vbInformation, Me.Caption
                        BuscarSiExisteNrocorrelativoEmergencia = True
                        Exit Do
                    End If
                    oRsTmp2.Close
                 End If
                 oRsTmp1.MoveNext
              Loop
           End If
           
        End If
        oRsTmp1.Close
        oConexion.Close
        Set oRsTmp1 = Nothing
        Set oConexion = Nothing
        Set oRsTmp2 = Nothing
End Function






Private Sub txtHoraIngreso_Change()
    lbHuboCambioEnDato = True
    If mi_Opcion = sghAgregar Then
        lblNombreMedico.Text = ""
        txtIdMedicoIngreso.Text = ""
    End If
End Sub



Private Sub txtMedicoRef_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtMedicoRef
End Sub

Private Sub txtMedicoRef_LostFocus()
    BuscaMedicoRerencia ""
End Sub

Private Sub txtNombreAcompañante_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub txtNombreAcompañante_LostFocus()
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtNombreAcompañante.Text
          lbHuboCambioEnDato = False
        End If
    If Len(txtNombreAcompañante.Text) > 0 Then
       txtNombreAcompañante.Text = UCase(txtNombreAcompañante.Text)
    End If
End Sub

Private Sub txtNroAfiliacionSis_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub txtNroAfiliacionSis_LostFocus()
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtNroAfiliacionSis.Text
          lbHuboCambioEnDato = False
        End If
End Sub

Private Sub txtNroHistoriaBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoriaBusqueda
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroHistoriaBusqueda_KeyPress(KeyAscii As Integer)
    
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 And Len(txtNroHistoriaBusqueda.Text) > 0 Then
       btnBuscarPaciente_Click
    End If
    
End Sub






Sub HabilitarFrameDestino(bValue As Boolean)
        mo_Formulario.HabilitarDeshabilitar btnBuscarEstablecimiento, bValue
End Sub






Private Sub cmbIdTipoGravedad_Change()
    If cmbIdTipoGravedad.Text <> "" Then
        'mo_cmbIdTipoGravedad.BoundText = Val(Split(cmbIdTipoGravedad.Text, " = ")(0))
        'A.Yañez 30-10-2014*************************
        If mo_cmbIdTipoGravedad.BoundText = "5" Then
        '*******************************************
            Me.ucPacientesDetalle1.PacienteNoIdentificado = True
        Else
            Me.ucPacientesDetalle1.PacienteNoIdentificado = False
       End If
    End If
    'A.Yañez 30-10-2014 *********************
    'If cmbIdTipoGravedad.Enabled = True Then cmbIdTipoGravedad.SetFocus
    '***************************************
End Sub

Private Sub cmbIdTipoGravedad_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoGravedad
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoGravedad_LostFocus()
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbIdTipoGravedad.Text
          lbHuboCambioEnDato = False
        End If
   If cmbIdTipoGravedad.Text <> "" Then
      ' mo_cmbIdTipoGravedad.BoundText = Val(Split(cmbIdTipoGravedad.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoGravedad
   On Error Resume Next
   lblNombreMedico.SetFocus
End Sub
















Private Sub Form_Initialize()
    Set mo_cmbIdTipoServicio.MiComboBox = cmbIdTipoServicio
    Set mo_cmbIdViasAdmision.MiComboBox = cmbIdViasAdmision
    Set mo_cmbIdCondicionEnElServicio.MiComboBox = cmbIdCondicionEnElServicio
    Set mo_cmbIdTipoReferenciaOrigen.MiComboBox = cmbIdTipoReferenciaOrigen
    Set mo_cmbIdCondicionEnElEstablecimiento.MiComboBox = cmbIdCondicionEnElEstablecimiento
    Set mo_cmbIdTipoGravedad.MiComboBox = cmbIdTipoGravedad
    Set mo_cmbIdTipoAgenteAGAN.MiComboBox = cmbIdTipoAgenteAGAN
    Set mo_cmbIdGrupoOcupacionalALAB.MiComboBox = cmbIdGrupoOcupacionalALAB
    Set mo_cmbIdPosicionLesionadoALAB.MiComboBox = cmbIdPosicionLesionadoALAB
    Set mo_cmbIdUbicacionLesionado.MiComboBox = cmbIdUbicacionLesionado
    Set mo_cmbIdTipoTransporte.MiComboBox = cmbIdTipoTransporte
    Set mo_cmbIdTipoVehiculo.MiComboBox = cmbIdTipoVehiculo
    Set mo_cmbIdClaseAccidente.MiComboBox = cmbIdClaseAccidente
    Set mo_cmbIdRelacionAgresorVictima.MiComboBox = cmbIdRelacionAgresorVictima
    Set mo_cmbIdSeguridad.MiComboBox = cmbIdSeguridad
    Set mo_cmbIdTipoEvento.MiComboBox = cmbIdTipoEvento
    Set mo_cmbIdLugarEvento.MiComboBox = cmbIdLugarEvento
    Set mo_cmbIdCausaExternaMorbilidad.MiComboBox = cmbIdCausaExternaMorbilidad
    Set mo_cmbIdTipoEdad.MiComboBox = cmbIdTipoEdad
    
End Sub


Private Sub txtApellidoPaternoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaternoBusqueda
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoPaternoBusqueda_LostFocus()
txtApellidoPaternoBusqueda.Text = mo_Teclado.CapitalizarNombres(txtApellidoPaternoBusqueda.Text)
'   If Len(txtApellidoPaternoBusqueda.Text) > 0 Then
      'btnBuscarPaciente_Click
'   End If
End Sub
Private Sub txtApellidoPaternoBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
   If KeyAscii = 13 Then
        If mo_AdminAdmision.RealizarBusquedaPacienteSiNo(wxParametroBusqRapida, txtApellidoPaternoBusqueda.Text, _
            txtApellidoMaternoBusqueda.Text) = True Then
            btnBuscarPaciente_Click
'        Else
'            If Trim(txtApellidoPaternoBusqueda.Text) = "" Then
'                MsgBox "Ingrese Apellido Paterno a Buscar", vbInformation, "Mensaje"
'                txtApellidoPaternoBusqueda.SetFocus
'            Else
'                MsgBox "Ingrese Apellido Materno a Buscar", vbInformation, "Mensaje"
'                txtApellidoMaternoBusqueda.SetFocus
'            End If
        End If
   End If
End Sub

Private Sub txtIdDiagnostico_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdTipoServicio_Click()
    mo_cmbIdViasAdmision.BoundColumn = "IdOrigenAtencion"
    mo_cmbIdViasAdmision.ListField = "DescripcionLarga"
    
    Select Case mo_cmbIdTipoServicio.BoundText
    Case 2
        Set mo_cmbIdViasAdmision.RowSource = mo_AdminAdmision.TiposOrigenAtencionSeleccionarViasDeConsEmergencia
    Case 3
        Set mo_cmbIdViasAdmision.RowSource = mo_AdminAdmision.TiposOrigenAtencionSeleccionarViasDeHospitalizacion(sghSoloPacHospitalizados)
    Case 4
        mo_cmbIdViasAdmision.BoundText = ""
        Set mo_cmbIdViasAdmision.RowSource = mo_AdminAdmision.TiposOrigenAtencionSeleccionarViasDeObsEmergencia
    
    End Select
    
End Sub

Private Sub cmbIdViasAdmision_Click()
lbHuboCambioEnDato = True

Dim sCodigoOrigen As String
    If cmbIdViasAdmision.Text <> "" Then
       sCodigoOrigen = Trim(Split(cmbIdViasAdmision.Text, " = ")(0))
    Else
       sCodigoOrigen = ""
    End If
    If sCodigoOrigen <> "R" And sCodigoOrigen <> "C" Then
        mo_cmbIdTipoReferenciaOrigen.BoundText = ""
        Me.txtIdEstablecimientoOrigen = ""
        Me.txtIdEstablecimientoOrigen.Tag = ""
        Me.lblNombreOrigenReferencia = ""
        txtReferenciaO.Text = ""
        cmbServicioReferenciaO.Text = ""                      'debb-21/06/2016
        txtMedicoRef.Text = "": Me.cmbMedicoRef.Text = ""     'FRANKLIN 2017
    Else
        mo_cmbIdTipoReferenciaOrigen.BoundText = "1"
        CargarAutomaticamenteEstablecimientoReferenciaSIS
    End If
    HabilitarFrameOrigen False
    Select Case sCodigoOrigen
    Case "R"
        HabilitarFrameOrigen True
        Me.fraDatosReferenciaOrigen = "Refer. Origen"
        Me.lblIdTipoReferenciaOrigen = "Tipo Referencia"
        Me.lblIdEstablecimientoOrigen = "Estab.Referencia"
        mo_cmbIdTipoReferenciaOrigen.BoundText = "1"
    Case "C"
        HabilitarFrameOrigen True
        Me.fraDatosReferenciaOrigen = "Contraref.Origen"
        Me.lblIdTipoReferenciaOrigen = "Tipo Contraref"
        Me.lblIdEstablecimientoOrigen = "Estab.Contraref."
        mo_cmbIdTipoReferenciaOrigen.BoundText = "1"
    End Select
    '
    If sCodigoOrigen = "J" Or sCodigoOrigen = "N" Then
       cmdBuscaMadre.Enabled = True
    Else
       cmdBuscaMadre.Enabled = False
    End If
End Sub

'debb-21/06/2016
Sub HabilitarFrameOrigen(bValue As Boolean)
        mo_Formulario.HabilitarDeshabilitar fraDatosReferenciaOrigen, bValue
        mo_Formulario.HabilitarDeshabilitar fraDatosReferenciaOrigen, bValue
        mo_Formulario.HabilitarDeshabilitar lblIdTipoReferenciaOrigen, bValue
        mo_Formulario.HabilitarDeshabilitar cmbIdTipoReferenciaOrigen, bValue
        mo_Formulario.HabilitarDeshabilitar lblIdEstablecimientoOrigen, bValue
        mo_Formulario.HabilitarDeshabilitar btnBuscarEstablecimiento, bValue
        mo_Formulario.HabilitarDeshabilitar Me.txtReferenciaO, bValue
        mo_Formulario.HabilitarDeshabilitar cmbServicioReferenciaO, bValue      'debb-21/06/2016
        mo_Formulario.HabilitarDeshabilitar lblServicioReferencia, bValue       'debb-21/06/2016
        mo_Formulario.HabilitarDeshabilitar lblNreferencia, bValue
        'FRANKLIN 2017
        mo_Formulario.HabilitarDeshabilitar txtMedicoRef, bValue
        mo_Formulario.HabilitarDeshabilitar Me.cmbMedicoRef, bValue
        
End Sub
Private Sub cmbIdViasAdmision_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdViasAdmision
   AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdViasAdmision_LostFocus()
    If lbHuboCambioEnDato = True Then
      sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, cmbIdViasAdmision.Text
      lbHuboCambioEnDato = False
    End If

    Dim oDOTipoOrigenAtencion As New DOTipoOrigenAtencion

   If cmbIdViasAdmision.Text <> "" Then
     Set oDOTipoOrigenAtencion = mo_AdminAdmision.TiposOrigenAtencionSeleccionarPorCodigo(Trim(Split(cmbIdViasAdmision.Text, " = ")(0)), ml_TipoServicio)
     If oDOTipoOrigenAtencion.IdOrigenAtencion <> 0 Then
         mo_cmbIdViasAdmision.BoundText = oDOTipoOrigenAtencion.IdOrigenAtencion
    End If
   End If
   mo_Formulario.MarcarComoVacio cmbIdViasAdmision
   Set oDOTipoOrigenAtencion = Nothing
   On Error Resume Next
   lblNombreServicio.SetFocus
End Sub

Private Sub txtApellidoMaternoBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaternoBusqueda
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoMaternoBusqueda_LostFocus()
   txtApellidoMaternoBusqueda.Text = mo_Teclado.CapitalizarNombres(txtApellidoMaternoBusqueda.Text)
'   If Len(txtApellidoMaternoBusqueda.Text) > 0 Then
'      btnBuscarPaciente_Click
'   End If
End Sub

Private Sub txtApellidoMaternoBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
   If KeyAscii = 13 Then
        If mo_AdminAdmision.RealizarBusquedaPacienteSiNo(wxParametroBusqRapida, txtApellidoPaternoBusqueda.Text, _
            txtApellidoMaternoBusqueda.Text) = True Then
            btnBuscarPaciente_Click
'        Else
'            If Trim(txtApellidoMaternoBusqueda.Text) = "" Then
'                MsgBox "Ingrese Apellido Materno a Buscar", vbInformation, "Mensaje"
'                txtApellidoMaternoBusqueda.SetFocus
'            Else
'                MsgBox "Ingrese Apellido Paterno a Buscar", vbInformation, "Mensaje"
'                txtApellidoPaternoBusqueda.SetFocus
'            End If
        End If
   End If
End Sub

Private Sub txtFechaIngreso_Change()
    lbHuboCambioEnDato = True
    On Error Resume Next
    Me.txtEdadEnDias = ""
    Dim oEdad As Edad
    
        oEdad = sighEntidades.CalcularEdad(CDate(Me.ucPacientesDetalle1.FechaNacimiento & " " & Me.ucPacientesDetalle1.HoraNacimiento), CDate(txtFechaIngreso + " " + txtHoraIngreso.Text))
    
    Me.txtEdadEnDias = oEdad.Edad
    mo_cmbIdTipoEdad.BoundText = oEdad.TipoEdad
    '
    ActualizaCheckRecienNacido
End Sub


Private Sub txtFechaIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaIngreso
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaIngreso_LostFocus()
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtFechaIngreso.Text
          lbHuboCambioEnDato = False
        End If
        
    If txtFechaIngreso <> sighEntidades.FECHA_VACIA_DMY Then
            If Not EsFecha(txtFechaIngreso, "DD/MM/AAAA") Then
                MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
                 txtFechaIngreso = sighEntidades.FECHA_VACIA_DMY
                 Exit Sub
            End If
    End If
    '
    chequeaFechaIngresoQueNoSeaMENORalORIGEN       'debb-22/02/2016
    '
    ChequeaQueNoExistaPacienteServicioFecha        'debb-22/08/2016
   
    mo_Formulario.MarcarComoVacio txtFechaIngreso

End Sub



Private Sub txtFechaIngreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub



Private Sub txtHoraIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraIngreso
    AdministrarKeyPreview KeyCode
End Sub

'debb-12/08/2016
Sub chequeaFechaIngresoQueNoSeaMENORalORIGEN()
    If IsDate(txtFechaIngreso.Text & " " & txtHoraIngreso.Text) And ml_TipoServicio = sghHospitalizacion And _
                                                    (mi_Opcion = sghAgregar Or mi_Opcion = sghModificar) Then
        Dim oRsBuscaAtencPac  As New Recordset
        If mo_cmbIdViasAdmision.BoundText = "30" Then
            Set oRsBuscaAtencPac = mo_AdminAdmision.AtencionesSeleccionarPorIdPaciente(ml_IdPaciente, sghConsultaExterna)
            If oRsBuscaAtencPac.RecordCount > 0 Then
                    oRsBuscaAtencPac.MoveLast
                    If Not IsNull(oRsBuscaAtencPac!fechaEgreso) Then
                         If CDate(oRsBuscaAtencPac!fechaEgreso & " " & oRsBuscaAtencPac!HoraEgreso) >= CDate(txtFechaIngreso.Text & _
                                                                                              " " & txtHoraIngreso.Text) Then
                            MsgBox "En CONSULTA EXTERNA, tiene Cuenta: " & oRsBuscaAtencPac!idCuentaAtencion & " con FECHA ATENCION: " & _
                                   oRsBuscaAtencPac!fechaEgreso & " " & oRsBuscaAtencPac!HoraEgreso & _
                                   "  que no puede ser mayor a la FECHA/HORA_INGRESO", vbInformation, "mensaje "
                            txtHoraIngreso.Text = sighEntidades.HORA_VACIA_HM
                            txtFechaIngreso.Text = sighEntidades.FECHA_VACIA_DMY
                            txtFechaIngreso.SetFocus
                         End If
                    ElseIf CDate(oRsBuscaAtencPac!FechaIngreso & " " & oRsBuscaAtencPac!HoraIngreso) >= CDate(txtFechaIngreso.Text & _
                                                                                               " " & txtHoraIngreso.Text) Then
                         MsgBox "En CONSULTA EXTERNA, tiene Cuenta: " & oRsBuscaAtencPac!idCuentaAtencion & " con Fecha: " & _
                                oRsBuscaAtencPac!FechaIngreso & " " & oRsBuscaAtencPac!HoraIngreso & _
                                "  que no puede ser mayor a la FECHA/HORA_INGRESO", vbInformation, "mensaje "
                         txtHoraIngreso.Text = sighEntidades.HORA_VACIA_HM
                         txtFechaIngreso.Text = sighEntidades.FECHA_VACIA_DMY
                         txtFechaIngreso.SetFocus
                    End If
            End If
            oRsBuscaAtencPac.Close
        ElseIf mo_cmbIdViasAdmision.BoundText = "31" Or mo_cmbIdViasAdmision.BoundText = "32" Then
            Set oRsBuscaAtencPac = mo_AdminAdmision.AtencionesSeleccionarPorIdPaciente(ml_IdPaciente, sghEmergenciaConsultorios)
            If oRsBuscaAtencPac.RecordCount > 0 Then
                    oRsBuscaAtencPac.MoveLast
                    
                    If IsNull(oRsBuscaAtencPac!fechaEgreso) Or IsNull(oRsBuscaAtencPac!HoraEgreso) Then
                        If oRsBuscaAtencPac!idEstado <> 5 And oRsBuscaAtencPac!idEstado <> 9 And oRsBuscaAtencPac!idEstado <> 13 Then   'cerrado/anulado/cerrado automatico
                             MsgBox "En EMERGENCIA, tiene Cuenta: " & oRsBuscaAtencPac!idCuentaAtencion & " con FECHA INGRESO: " & _
                                    oRsBuscaAtencPac!FechaIngreso & " " & oRsBuscaAtencPac!HoraIngreso & _
                                    "  que no le ha dado ALTA MEDICA, fijarse opción FACTURACION->ESTADO DE CUENTA", vbInformation, "mensaje "
                        End If
                    ElseIf CDate(oRsBuscaAtencPac!fechaEgreso & " " & oRsBuscaAtencPac!HoraEgreso) >= CDate(txtFechaIngreso.Text & _
                                       " " & txtHoraIngreso.Text) Then
                             MsgBox "En EMERGENCIA, tiene Cuenta: " & oRsBuscaAtencPac!idCuentaAtencion & " con FECHA ALTA: " & _
                                    oRsBuscaAtencPac!fechaEgreso & " " & oRsBuscaAtencPac!HoraEgreso & _
                                    "  que no puede ser mayor a la FECHA/HORA_INGRESO", vbInformation, "mensaje "
                             txtHoraIngreso.Text = sighEntidades.HORA_VACIA_HM
                             txtFechaIngreso.Text = sighEntidades.FECHA_VACIA_DMY
                             txtFechaIngreso.SetFocus
                    End If
            End If
            oRsBuscaAtencPac.Close
        End If
        Set oRsBuscaAtencPac = Nothing
    End If

End Sub

Private Sub txtHoraIngreso_LostFocus()
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtHoraIngreso.Text
          lbHuboCambioEnDato = False
        End If
        
    If txtHoraIngreso <> sighEntidades.HORA_VACIA_HM Then
        If Not sighEntidades.ValidaHora(txtHoraIngreso) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
             txtHoraIngreso = sighEntidades.HORA_VACIA_HM
        End If
    End If
        
    'WCG comentado por facturacion
    On Error Resume Next
    '
    chequeaFechaIngresoQueNoSeaMENORalORIGEN                'debb-22/02/2016
    '
    mo_Formulario.MarcarComoVacio txtHoraIngreso
    '
    If ml_TipoServicio = sghHospitalizacion Then
        On Error Resume Next
        lblNombreMedico.SetFocus
    End If
End Sub

Private Sub txtHoraIngreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtIdMedicoIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdMedicoIngreso
    If KeyCode = vbKeyF1 Then
        btnBuscarMedicos_Click
    End If
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdMedicoIngreso_LostFocus()
    CompletarDatosDeMedicoEnElLostFocus txtIdMedicoIngreso, lblNombreMedico
    mo_Formulario.MarcarComoVacio txtIdMedicoIngreso
End Sub

Private Sub txtIdMedicoIngreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdServicioIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdServicioIngreso
    If KeyCode = vbKeyF1 Then
        btnBuscarServicios_Click
    End If
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdServicioIngreso_LostFocus()
    CompletarDatosDeServicioEnElLostFocus txtIdServicioIngreso, lblNombreServicio
    mo_Formulario.MarcarComoVacio txtIdServicioIngreso
End Sub

Private Sub txtIdServicioIngreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtNroDNIBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNroDNIBusqueda
AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroDNIBusqueda_LostFocus()
    txtNroDNIBusqueda.Text = mo_Teclado.CapitalizarNombres(txtNroDNIBusqueda.Text)
   If Len(txtNroDNIBusqueda.Text) > 0 Then
      btnBuscarPaciente_Click
   End If
End Sub

Private Sub txtNroDNIBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Sub CargarDatosAlFormulario()
    Dim lnIdCamaSeleccionada As Long
    mo_Formulario.HabilitarDeshabilitar Me.cmbIdTipoServicio, False
    mo_Formulario.HabilitarDeshabilitar Me.txtIdMedicoIngreso, False
    mo_Formulario.HabilitarDeshabilitar Me.txtIdServicioIngreso, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNroOrdenPago, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNroCuenta, False
    mo_Formulario.HabilitarDeshabilitar Me.txtCamaTransf, False
    
    mo_Formulario.HabilitarDeshabilitar Me.txtServicioTransf, False 'Frank 06052015
    
    mo_Formulario.HabilitarDeshabilitar Me.txtIdEstablecimientoOrigen, False
    mo_Formulario.HabilitarDeshabilitar txtNroOrdenPago, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbFuenteFinanciamiento, True
    
    'mgaray
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGravedad, True
    mo_Formulario.HabilitarDeshabilitar cmbIdViasAdmision, True
    mo_Formulario.HabilitarDeshabilitar txtFechaIngreso, True
    mo_Formulario.HabilitarDeshabilitar txtHoraIngreso, True
    
    mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, False
    mo_Formulario.HabilitarDeshabilitar lblMadre, False
    mo_Formulario.HabilitarDeshabilitar lblNombreOrigenReferencia, False 'debb-05/04/2011
    
'    If mi_Opcion = sghAgregar Then
'       btnBuscarServicios.Enabled = False
'       btnBuscarMedicos.Enabled = False
'    Else
'       mo_Formulario.HabilitarDeshabilitar Me.lblNombreServicio, False
'       mo_Formulario.HabilitarDeshabilitar Me.lblNombreMedico, False
'    End If
    
    If mi_Opcion = sghAgregar Then
           btnBuscarServicios.Enabled = False
           btnBuscarMedicos.Enabled = False
    Else
           mo_Formulario.HabilitarDeshabilitar Me.lblNombreServicio, True 'A.Yañez-13/10/2014 (False)
           mo_Formulario.HabilitarDeshabilitar Me.lblNombreMedico, True ' A.Yañez-13/10/2014 (False)
    End If
    
    Me.ucPacientesDetalle1.HacerVisibleCheckPacienteNoIdentificado IIf(ml_TipoServicio = sghEmergenciaConsultorios Or ml_TipoServicio = sghEmergenciaObservacion, True, False)
    Me.ucPacientesDetalle1.NotaSobreUbicacion = "(*) Datos del día de la atención del paciente"
    
    Me.ucTransferenciasDetalle1.TipoServicio = ml_TipoServicio
    
    Me.ucDiagnosticosIngreso.TituloFrame = "Diagnósticos de Ingreso     (F1=Todos Dx)"
    Me.ucDiagnosticoNacimiento.TituloFrame = "Diagnósticos de muerte fetal     (F1=Todos Dx)"
    
    lcUltimoCodigoDeServicioTransferido = ""
        
    Select Case ml_TipoAccionAdmision
    Case sghAdmisionNormal  'Si el una admisión normal de hospitalizacion o de emergencia
        Select Case mi_Opcion
            Case sghAgregar
                ValoresPorDefecto
                If ml_IdServicioConCamaDisponible > 0 Then
                   CargaNombreDelServicioIngreso ml_IdServicioConCamaDisponible
                   lnIdCamaSeleccionada = mo_AdminHoteleria.CargaCamaDisponible(ml_IdServicioConCamaDisponible)
                   If lnIdCamaSeleccionada > 0 Then
                      CargaCamaSeleccionada lnIdCamaSeleccionada
                   End If

                End If
            Case sghModificar
                CargarDatosAlosControles
            Case sghConsultar
                CargarDatosAlosControles
            Case sghEliminar
                CargarDatosAlosControles
        End Select
        
    Case sghEnviarAObservacion
'        mi_Opcion = sghAgregar
'        ValoresPorDefecto
'        CargarDatosParaEnviarAObservacion
        
    Case sghTrasladarAHospitalizacion
'        mi_Opcion = sghAgregar
'        ValoresPorDefecto
'        CargarDatosParaEnviarAHospitalizacion
        
    Case sghDarDeAlta
        mi_Opcion = sghModificar
        CargarDatosAlosControles
        
    Case sghIngresarUnAlojamientoConjunto
        mi_Opcion = sghAgregar
        ValoresPorDefecto
        CargarDatosParaAgregarUnAlojamientoConjunto
        
    Case sghTransferencias
        mi_Opcion = sghModificar
        CargarDatosAlosControles
        
    End Select
    
     Select Case mi_Opcion
     Case sghAgregar
        Me.btnImprimir.Enabled = False
     Case sghModificar
        fraBusqueda.Enabled = False
        Me.btnImprimir.Enabled = True
        Me.chkPacienteNuevo.Enabled = False
        If Not Me.ucPacientesDetalle1.PacienteNoIdentificado Then
            Me.ucPacientesDetalle1.HacerVisibleCheckPacienteNoIdentificado False
        End If
     
     Case sghConsultar
        DeshabilitarControlesParaEdicion
        Me.btnAceptar.Visible = False
        Me.btnImprimir.Enabled = True
        If Not Me.ucPacientesDetalle1.PacienteNoIdentificado Then
            Me.ucPacientesDetalle1.HacerVisibleCheckPacienteNoIdentificado False
        End If
    
    Case sghEliminar
        DeshabilitarControlesParaEdicion
        Me.btnImprimir.Enabled = True
        If Not Me.ucPacientesDetalle1.PacienteNoIdentificado Then
            Me.ucPacientesDetalle1.HacerVisibleCheckPacienteNoIdentificado False
        End If
    
    End Select
    
End Sub

'Sub CargarDatosParaEnviarAObservacion()
'
'        CargarDatosDeLasAtencionesPadres
'
'        '1ro:   CARGAR DATOS DEL PACIENTE
'        Me.ucPacientesDetalle1.idPaciente = Me.idPaciente
'        Me.ucPacientesDetalle1.CargarDatosDePacienteALosControles
'
'        mo_cmbIdViasAdmision.BoundText = 40   'indica que viene de consultorio de emergencia
'        mo_Formulario.HabilitarDeshabilitar cmbIdViasAdmision, False
'
'        lcBuscaParametro.RetornaHoraServidorSQL1
'        Me.txtFechaIngreso.Text = lcBuscaParametro.RetornaFechaServidorSQL 'Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
'        Me.txtHoraIngreso.Text = lcBuscaParametro.RetornaHoraServidorSQL 'Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
'
'End Sub
'Sub CargarDatosParaEnviarAHospitalizacion()
'
'        CargarDatosDeLasAtencionesPadres
'
'        '1ro:   CARGAR DATOS DEL PACIENTE
'        Me.ucPacientesDetalle1.idPaciente = Me.idPaciente
'        Me.ucPacientesDetalle1.CargarDatosDePacienteALosControles
'
'        If mo_AtencionPadre.idTipoServicio = sghEmergenciaConsultorios Then
'            mo_cmbIdViasAdmision.BoundText = 31   'indica que viene de consultorio de emergencia
'        End If
'        If mo_AtencionPadre.idTipoServicio = sghEmergenciaObservacion Then
'            mo_cmbIdViasAdmision.BoundText = 32   'indica que viene de consultorio de emergencia
'        End If
'
'        mo_Formulario.HabilitarDeshabilitar cmbIdViasAdmision, False
'
'
'        Me.txtFechaIngreso.Text = lcBuscaParametro.RetornaFechaServidorSQL 'Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
'        Me.txtHoraIngreso.Text = lcBuscaParametro.RetornaHoraServidorSQL 'Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
'
'End Sub
Sub CargarDatosParaAgregarUnAlojamientoConjunto()
        Dim oConexion As New Connection
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(Me.idCuentaAtencion, oConexion)
        
        ucPacientesDetalle1.TipoNumeracion = sghHistoriaTemporalAlojamiento
        
        mo_cmbIdViasAdmision.BoundText = 35
        mo_Formulario.HabilitarDeshabilitar cmbIdViasAdmision, False
        
        Me.txtFechaIngreso.Text = lcBuscaParametro.RetornaFechaServidorSQL 'Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
        Me.txtHoraIngreso.Text = lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos 'Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
        oConexion.Close
        Set oConexion = Nothing
End Sub

Sub DeshabilitarControlesParaEdicion()
    
    fraBusqueda.Enabled = False
    fraDatosReferenciaOrigen.Enabled = False
    Me.chkPacienteNuevo.Enabled = False
    Me.ucPacientesDetalle1.DeshabilitarFrames

End Sub

Sub ValoresPorDefecto()

    Me.txtFechaIngreso.Text = lcBuscaParametro.RetornaFechaServidorSQL 'Format(Now, sighEntidades.DevuelveFechaSoloFormato_DMY)
    Me.txtHoraIngreso.Text = lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos  'Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
    Me.ucPacientesDetalle1.Opcion = mi_Opcion
    Me.ucPacientesDetalle1.TipoServicio = ml_TipoServicio
    Me.ucPacientesDetalle1.ConfigurarValoresPorDefecto
    Me.ucNacimientoDetalle1.FechaIngreso = CDate(Format(txtFechaIngreso.Text, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & txtHoraIngreso.Text)
End Sub


Sub Form_Load()
    If mo_lbCargaTablasUnaVez = True Then
    
        lbCargaTablasUnaVez = False
        
        
        InicilizarParametros
        Me.ucPacientesDetalle1.Inicializar
        Me.ucDiagnosticosIngreso.Inicializar
        Me.ucTransferenciasDetalle1.idUsuario = ml_idUsuario
        Me.ucTransferenciasDetalle1.Inicializar
        UcPacientesSunasa1.Inicializar
        UcPacientesSunasa1.YaNoTieneSeguro
        Me.ucNacimientoDetalle1.Inicializar
        Me.ucDiagnosticoNacimiento.Inicializar
        CargarComboBoxes
        '
        lbBuscaDNIenReniec = IIf(wxParametro296 = "S", True, False)
        If lbBuscaDNIenReniec = True Then
           mo_Reniec.SeAccesaAlaWebDesdeGalenhos = True
           mo_Reniec.Inicializar
        End If
        '
        'InicializarFUA
        '
        mo_Apariencia.ConfigurarFilasBiColores Me.grdPacientesEncontrados, sighEntidades.GrillaConFilasBicolor
        mo_Apariencia.ConfigurarFilasBiColores Me.grdOtrosCpt, sighEntidades.GrillaConFilasBicolor
        mo_Apariencia.ConfigurarFilasBiColores Me.grdApoyoDx, sighEntidades.GrillaConFilasBicolor
        '

    End If
    '
    SiempreCargaPorMovimiento
    
End Sub




Sub SiempreCargaPorMovimiento()
    If mo_lbNuevoMovimiento = True Then
        sighEntidades.ParaAuditoria = ""
        If Val(wxParametro208) = 1910 Then   'sullana
           lblGravedad.Caption = "Prioridad"
        End If
        
        Me.TxtCitaTratamiento.Text = ""
        txtNroAfiliacionSis.Text = ""   'debb-04/07/2016
        lbCargaUnaVezVEntana = True
        mo_lbNuevoMovimiento = False
        lbPacienteNN = False
        lcCaptionTab2 = ""
        lnFocusCuandoCargeFrm = 0
        lnIdNacimientoSeleccionado = 0
        Me.grdPacientesEncontrados.Visible = False
        fraTriaje.Visible = False
        FraProviene.Visible = False: lbProcedeDeConsExt = False: lbProcedeDeEmergencia = False       'debb-23/02/2015
        '
        
        Select Case ml_TipoServicio
        Case sghConsultaExterna
            Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarAsistenciales
            mo_cmbIdTipoServicio.BoundText = "1"
            mo_Formulario.HabilitarDeshabilitar cmbIdTipoServicio, False
            
        Case sghHospitalizacion
            FraProviene.Visible = IIf(mi_Opcion = sghAgregar, True, False) 'debb-23/02/2015
            Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarAsistenciales
            mo_cmbIdTipoServicio.BoundText = "3"
            mo_Formulario.HabilitarDeshabilitar cmbIdTipoServicio, False
            lblOrdenPago.Visible = False
            txtNroOrdenPago.Visible = False
            If mi_Opcion = sghModificar Then
               fraTriaje.Visible = True
            End If
            lblEmergenciaN.Visible = False: txtEmergenciaN.Visible = False   'debb-21/06/2016
            btnImprimir.Caption = "HojaFiliac Hosp"
            cmbComoLlego.Visible = False
            cmbTipoAtencion.Visible = False
            Label12.Visible = False
            Label11.Visible = False
            Label10.Visible = False: cmbEstadoLlegada.Visible = False: cmbEstadoLlegada.Text = ""
        Case sghEmergenciaConsultorios
            Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarDeEmergencia
            mo_Formulario.HabilitarDeshabilitar cmbIdTipoServicio, True
            mo_cmbIdTipoServicio.BoundText = "2"
            '
            mo_cmbIdCausaExternaMorbilidad.BoundColumn = "IdCausaExternaMorbilidad"
            mo_cmbIdCausaExternaMorbilidad.ListField = "DescripcionLarga"
            Set mo_cmbIdCausaExternaMorbilidad.RowSource = mo_AdminAdmision.EmergenciaCausaExternaMorbilidadSeleccionarTodos()
            '
            If mi_Opcion = sghModificar Then
               fraTriaje.Visible = True
            End If
            btnImprimir.Caption = "HistClin Emerg"
            
            lblEmergenciaN.Visible = True: txtEmergenciaN.Visible = True   'debb-21/06/2016
            cmbComoLlego.Text = ""
            cmbTipoAtencion.Text = ""
            cmbComoLlego.Visible = True
            cmbTipoAtencion.Visible = True
            Label12.Visible = True
            Label11.Visible = True
            Label10.Visible = True: cmbEstadoLlegada.Visible = True: cmbEstadoLlegada.Text = ""
           If Len(wxParametro546) = 7 Then    'debb-03/04/2018
              cmbComoLlego.ListIndex = Val(Mid(wxParametro546, 2, 1))
              cmbTipoAtencion.ListIndex = Val(Mid(wxParametro546, 4, 1))
              Me.cmbEstadoLlegada.ListIndex = Val(Mid(wxParametro546, 6, 1))
           End If
            
        Case sghEmergenciaObservacion
            Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarDeEmergencia
            mo_Formulario.HabilitarDeshabilitar cmbIdTipoServicio, True
            mo_cmbIdTipoServicio.BoundText = "4"
            '
            mo_cmbIdCausaExternaMorbilidad.BoundColumn = "IdCausaExternaMorbilidad"
            mo_cmbIdCausaExternaMorbilidad.ListField = "DescripcionLarga"
            Set mo_cmbIdCausaExternaMorbilidad.RowSource = mo_AdminAdmision.EmergenciaCausaExternaMorbilidadSeleccionarTodos()
            btnImprimir.Caption = "HistClin Emerg"
        End Select
        '
        Me.UcPacientesSunasa1.idSunasaPacienteHistorico_idPaciente_ConValorCero
        Me.UcPacientesSunasa1.LimpiarDatos
        Me.UcPacientesSunasa1.PaisTitularDefault
        UcPacientesSunasa1.YaNoTieneSeguro
        UcPacientesSunasa1.Opcion = mi_Opcion
        '
        fraBusqueda.Enabled = True
        btnAceptar.Enabled = True: btnAceptar.Visible = True
        cmdBuscaMadre.Enabled = False
        '
        LimpiaTodosControles
        ConfiguraTABSsegunPermisosDelUsuario
        ConfigurarControles
        CargarDatosAlFormulario
        
        'mgaray20141008
        Call BloquearEdicionAdmisionSegunReglas
        
        mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
        If mi_Opcion = sghAgregar Then
            mo_Formulario.HabilitarDeshabilitar lblNombreServicio, True
            mo_Formulario.HabilitarDeshabilitar lblNombreMedico, True
            tabAdmision.Tab = 0
            On Error Resume Next
            Me.txtApellidoPaternoBusqueda.SetFocus
        Else
            cmdFiliacionCE.Enabled = True
        End If
        grdPacientesEncontrados.Visible = False
        grdPacientesEncontrados.Height = 0
        '
        ucPacientesDetalle1.Opcion = mi_Opcion
        '
        If mi_Opcion = sghAgregar Then
            'mgaray20140926
'           lbElServicioRegistraFUA = "N"
'           wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
           Call SetVariableServicioUsaFUA(Nothing)
        End If
        InicializarFUA
        '
        'CargaDefaultVentanaDelCursor
        '
        CalculaEmergenciaNumero      'debb-21/06/2016
        '
        If mi_Opcion = sghAgregar Then
           'btnAceptar.Enabled = Not true    'licencia
        End If
        
    End If
End Sub
'mgaray20140926
Private Sub BuscarDatosServicioYAsignarVariablesFUA(lIdServicio As Long)
    Dim oServicio As Servicios
    Dim oDoServicio As doServicio
    Set oDoServicio = New doServicio
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    oDoServicio.IdServicio = lIdServicio
    Set oServicio = New Servicios
    Set oServicio.Conexion = oConexion
    If oServicio.SeleccionarPorId(oDoServicio) = True Then
        Call SetVariableServicioUsaFUA(oDoServicio)
    Else
        Call SetVariableServicioUsaFUA(Nothing)
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub
'mgaray20140926
Private Sub SetVariableServicioUsaFUA(oDoServicio As doServicio)
    lcElServicioUsaGalenHos = "N"
    lbElServicioRegistraFUA = "N"
    wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
    
    If Not (oDoServicio Is Nothing) Then
        If oDoServicio.IdServicio > 0 Then
            lcElServicioUsaGalenHos = IIf(oDoServicio.UsaGalenHos = True, "S", "N")
            If oDoServicio.UsaFUA = True Then
                lbElServicioRegistraFUA = "S"
            Else
                 wxParametro302 = "N"
            End If
        End If
    End If
    InicializarFUA
End Sub


Sub InicializarFUA()
        If wxParametro302 = "S" Then
           ucSISfuaCodPrestacion1.Visible = True
           Me.chkBuscarEnSIS.Visible = True
           UcSISafiliacion1.Visible = True
           UcSISafiliacion1.Inicializar
        Else
           'ucSISfuaCodPrestacion1.Visible = False A.Yañez 11/11/2014
           'Me.chkBuscarEnSIS.Visible = False
           'UcSISafiliacion1.Visible = False
        End If
End Sub

Sub LimpiaTodosControles()
    If mi_Opcion = sghAgregar Then
            Me.ucPacientesDetalle1.LimpiarDatosDePaciente wxParametro211, ldFechaActualServidor
            Me.ucPacientesDetalle1.HabilitarFrames
            '
            fraPacienteNuevo.Enabled = True
            Me.chkPacienteNuevo.Enabled = True
            '
            chkPacienteNuevo.Value = 0
            '
            cmbIdViasAdmision.Text = ""
            lblNombreServicio.Text = ""
            txtIdServicioIngreso.Text = ""
            lblNombreMedico.Text = ""
            txtIdMedicoIngreso.Text = ""
            chkRecienNacido.Value = 0
            txtEdadEnDias.Text = ""
            txtNroCamaIngreso.Text = ""
            chkLlegoSI.Value = 0
            cmbIdTipoGravedad.Text = ""
            txtNombreAcompañante.Text = ""
            txtDNIacompaniante.Text = "": txtEmergenciaN.Text = ""  'debb-21/06/2016
            cmbFuenteFinanciamiento.Text = ""
            cmbFormaPago.Text = ""
            txtNroCuenta.Text = ""
            txtNroOrdenPago.Text = ""
            ucMensajeParpadeando1.MensajeDeTexto = ""
            cmbIdTipoReferenciaOrigen.Text = ""
            txtIdEstablecimientoOrigen.Text = ""
            lblNombreOrigenReferencia.Text = ""
            txtReferenciaO.Text = ""
            cmbServicioReferenciaO.Text = ""            'debb-21/06/2016
            Me.txtIdMedicoIngreso.Tag = ""
            Me.txtIdServicioIngreso.Tag = ""
            Me.txtIdEstablecimientoOrigen.Tag = ""
            Me.txtNroCamaIngreso.Tag = ""
            lblMadre.Text = ""
            txtIdMedicoNacimiento.Text = ""
            '
            chkLlegoSS.Value = 0
            txtCamaTransf.Text = ""
            txtServicioTransf.Text = ""
            '
            cmbIdCausaExternaMorbilidad.Text = ""
            cmbIdLugarEvento.Text = ""
            cmbIdTipoEvento.Text = ""
            cmbIdSeguridad.Text = ""
            cmbIdRelacionAgresorVictima.Text = ""
            cmbIdClaseAccidente.Text = ""
            cmbIdTipoVehiculo.Text = ""
            cmbIdTipoTransporte.Text = ""
            cmbIdUbicacionLesionado.Text = ""
            cmbIdPosicionLesionadoALAB.Text = ""
            cmbIdGrupoOcupacionalALAB.Text = ""
            cmbIdTipoAgenteAGAN.Text = ""
            '
            '
            'mo_cmbIdTipoGravedad.BoundText = ""
            mo_cmbIdViasAdmision.BoundText = ""
            mo_cmbIdEspecialidadMedico.BoundText = ""
            mo_cmbIdServicio.BoundText = ""
            mo_cmbIdCondicionEnElServicio.BoundText = ""
            mo_cmbIdTipoReferenciaOrigen.BoundText = ""
            mo_cmbIdCondicionEnElEstablecimiento.BoundText = ""
            mo_cmbIdTipoAgenteAGAN.BoundText = ""
            mo_cmbIdGrupoOcupacionalALAB.BoundText = ""
            mo_cmbIdPosicionLesionadoALAB.BoundText = ""
            mo_cmbIdUbicacionLesionado.BoundText = ""
            mo_cmbIdTipoTransporte.BoundText = ""
            mo_cmbIdTipoVehiculo.BoundText = ""
            mo_cmbIdClaseAccidente.BoundText = ""
            mo_cmbIdRelacionAgresorVictima.BoundText = ""
            mo_cmbIdSeguridad.BoundText = ""
            mo_cmbIdTipoEvento.BoundText = ""
            mo_cmbIdLugarEvento.BoundText = ""
            'mo_cmbIdCausaExternaMorbilidad.BoundText = ""
            mo_cmbIdTipoEdad.BoundText = ""
            '
            txtNroHistoriaBusqueda.Text = ""
            txtApellidoPaternoBusqueda.Text = ""
            txtApellidoMaternoBusqueda.Text = ""
            txtPrimerNombreBusqueda.Text = ""
            txtSegundoNombreBusqueda.Text = ""
            txtNroDNIBusqueda.Text = ""
            lnIdDistritoSIS = 0: lnIdSexoSIS = 0: ldFechaNacimientoSIS = 0: lcSnombreSIS = "": lnIdPlanSIS = 0
            UcSISafiliacion1.InabilitaControles True
            Me.ucSISfuaCodPrestacion1.Visible = False
            If wxParametro302 = "S" Then
               Me.ucSISfuaCodPrestacion1.CodigoPrestacion = ""
            End If
            '
            ml_EstadoCuenta = 0
            ml_idCuentaAtencion = 0
            ml_idAtencion = 0
            ml_IdAtencionPadre = 0
            mo_NroServiciosQuePasoElPaciente = 0
            ml_IdPaciente = 0
            ms_Autogenerado = ""
            ml_IdMedico = 0
            ms_NombreMedico = ""
            ml_IdPrestamo = 0
            lcApP = ""
            lcApM = ""
            lcPnom = ""
            lcSnombreReniec = "": ldFnacimientoReniec = 0: lnIdSexoReniec = 0: lcDireccionReniec = "": mb_UsoWebReniec = False
            ml_IdServicioConCamaDisponible = 0
            lbPacienteNN = False
            lnFocusCuandoCargeFrm = 0
            lnIdPlanAnterior = 0
            lnIdTipoFinanciamientoAnterior = 0
    End If
    txtCamaTransf.Visible = False
    txtServicioTransf.Visible = False
    lblServicioTransf.Visible = False
    fraServicioActual.Visible = False
    fraServicioActual.Top = 5520
    ucTransferenciasDetalle1.Top = 360
    chkLlegoSS.Visible = False
    lblCamaTransf.Visible = False
    cmdCamaTransf.Visible = False
    chkLlegoSI.Visible = False
    lcUltimoCodigoDeServicioTransferido = ""
    lblNroCamaIngreso.Visible = True: txtNroCamaIngreso.Visible = True: btnVerDisponibilidadDeCamas.Visible = True
    
    txtMedicoRef.Text = "": Me.cmbMedicoRef.Text = ""  'franklin 2017
    '
    Me.ucDiagnosticosIngreso.LimpiarDatos
    Me.ucTransferenciasDetalle1.LimpiarDatos
    Me.ucNacimientoDetalle1.LimpiarDatos
    Me.ucDiagnosticoNacimiento.LimpiarDatos
    'mgaray20140926
    Me.btnImprimeFichaSIS.Visible = False
    '
'    Me.txtPeso.Text = ""
'    Me.txtPresion.Text = "___/___"
'    Me.txtTemperatura.Text = ""
'    Me.txtTalla.Text = ""
'    Me.txtPulso.Text = ""
'    Me.txtFrespiratoria.Text = ""
    lcHistoriaYpaciente = ""
    '
    Set mo_Diagnosticos = Nothing
    ml_idAtencionEmeg_CE = 0
End Sub


Sub ConfiguraTABSsegunPermisosDelUsuario()
    
    Dim oRsPermisosTabs As New Recordset
    lb_puedeCambiarFuenteFinanciamiento = False          'debb-14/03/2015
    
    Me.tabAdmision.TabsPerRow = 2
    Me.tabAdmision.TabVisible(0) = False
    Me.tabAdmision.TabVisible(1) = False
    Me.TabIngreso.TabVisible(5) = False
    If mi_Opcion <> sghAgregar Then
       Me.TabIngreso.TabsPerRow = 5
       If Val(mo_cmbIdTipoServicio.BoundText) = sghEmergenciaConsultorios And mi_Opcion = sghModificar Then
          Me.TabIngreso.TabsPerRow = 6
       End If
    Else
       Me.TabIngreso.TabsPerRow = 4
    End If
    Me.TabIngreso.TabVisible(0) = False
    Me.TabIngreso.TabVisible(1) = False
    Me.TabIngreso.TabVisible(2) = False
    Me.TabIngreso.TabVisible(3) = False
    Me.TabIngreso.TabVisible(4) = False
    Set oRsPermisosTabs = ms_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_idUsuario)
    If oRsPermisosTabs.RecordCount > 0 Then
       Do While Not oRsPermisosTabs.EOF
          Select Case oRsPermisosTabs.Fields!IdPermiso
          Case 350    'Admision Hosp/Emerg - Ver TAB 1.   Datos del Paciente
               Me.tabAdmision.TabVisible(0) = True
          Case 351    'Admisión Hosp/Emerg - Ver TAB 2.1 Ingreso
               Me.tabAdmision.TabVisible(1) = True
               Me.TabIngreso.TabVisible(0) = True
          Case 352    'Admisión Hosp/Emerg - Ver TAB 2.2 Transferencia
               Me.tabAdmision.TabVisible(1) = True
               Me.TabIngreso.TabVisible(1) = True
          Case 353    'Admisión Hosp/Emerg - Ver TAB 2.3 Causas Externas morbilidad Ing.
               Me.tabAdmision.TabVisible(1) = True
               Me.TabIngreso.TabVisible(2) = True
          Case 354    'Admisión Hosp/Emerg - Ver TAB 2.4 Diagnósticos Ing.
               Me.tabAdmision.TabVisible(1) = True
               'If mi_Opcion <> sghAgregar Then
                  Me.TabIngreso.TabVisible(3) = True
              ' End If
          Case 355    'Admisión Hosp/Emerg - Ver TAB 3.1 Egreso
               
          Case 356    'Admisión Hosp/Emerg - Ver TAB 3.2 Dx y Complicaciones Egr
          Case 357    'Admisión Hosp/Emerg - Ver TAB 3.3 Nacimientos Egr
               Me.TabIngreso.TabVisible(4) = True
          Case 358    'Admisión Hosp/Emerg - Ver TAB 3.4 Mortalidad Egr
          Case 362    'Admisión Hosp - Confirmar llegada de Paciente desde Adm.Emerg
               lbUsuarioConfirmaLlegada = True
          Case 363    'Admisión Hosp - Confirmar llegada de Paciente Transferido
               lbUsuarioConfirmaTransferencia = True
          Case 370    'Puede cambiar la FUENTE DE FINANCIAMIENTO (SOLO HOSPITALIZACION) PROVENIENTE DE EMERGENCIA O CE    'debb-14/03/2015
               lb_puedeCambiarFuenteFinanciamiento = True                   'debb-14/03/2015
          Case 371    'Admisión Emerg - Ver TAB 2.6 Tratamiento
               If Val(mo_cmbIdTipoServicio.BoundText) = sghEmergenciaConsultorios And mi_Opcion = sghModificar Then
                  Me.TabIngreso.TabVisible(5) = True
               End If
          End Select
          oRsPermisosTabs.MoveNext
       Loop
    End If
    Set oRsPermisosTabs = Nothing
   ' If ml_TipoServicio <> sghHospitalizacion Then
   '    Me.TabEgresos.TabVisible(2) = False
   ' End If
End Sub
Sub ConfigurarControles()
Dim oDOCuentaAtencionPadre As New DOCuentaAtencion
Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
     
    Me.ucDiagnosticosIngreso.TituloFrame = "Diagnósticos ingreso     (F1=Todos Dx)"
    HabilitarFrameOrigen False
    HabilitarFrameDestino False

    Select Case ml_TipoServicio
    Case sghHospitalizacion
        
        TituloDeForm "hospitalización"
        cmbIdTipoGravedad.Visible = False
        lblGravedad.Visible = False
    
        
        'Causa externa de morbilidad no se debe ver
        TabIngreso.TabVisible(2) = False
        TabIngreso.TabsPerRow = 4
                
        If ml_TipoAccionAdmision = sghTrasladarAHospitalizacion Then
            'No se debe ver datos de egreso
            Me.tabAdmision.TabVisible(2) = False
            Me.tabAdmision.TabsPerRow = 2
            
            'No se debe ver las transferencias
            Me.TabIngreso.TabVisible(1) = False
            Me.TabIngreso.TabsPerRow = 1
        
            
            'Me.ucDiagnosticosIngreso.IdAtencion = oDOCuentaAtencionPadre.IdAtencion
            Me.ucDiagnosticosIngreso.TipoDiagnostico = sghHospitalizacionIngreso
            Me.ucDiagnosticosIngreso.CargarDatosDeDiagnosticos oConexion
        
        End If
                
        If mi_Opcion = sghAgregar Then
            'No se debe ver datos de egreso
            Me.tabAdmision.TabsPerRow = 2
            
            'No se debe ver las transferencias
            Me.TabIngreso.TabVisible(1) = False
            Me.TabIngreso.TabVisible(3) = False
            Me.TabIngreso.TabsPerRow = 2
            
            'No se debe ver NACIMIENTOS
            Me.TabIngreso.TabVisible(4) = False
        End If
        
        If ml_TipoAccionAdmision = sghIngresarUnAlojamientoConjunto Then
            'No se debe ver datos de egreso
            Me.tabAdmision.TabVisible(2) = False
            Me.tabAdmision.TabsPerRow = 2
            
            'No se debe ver las transferencias
            Me.TabIngreso.TabVisible(1) = False
            Me.TabIngreso.TabsPerRow = 1
        End If
                
        If ml_TipoAccionAdmision = sghTransferencias Then
                'No se debe ver datos de egreso
                Me.tabAdmision.TabVisible(2) = False
                Me.tabAdmision.TabsPerRow = 2
        
                'Se debe ver las transferencias y lo presenta primero
                Me.tabAdmision.Tab = 1
                Me.TabIngreso.Tab = 1
        End If
        
        If ml_TipoAccionAdmision = sghDarDeAlta Then
                'Se ubica en el tab de egreso
                Me.tabAdmision.Tab = 2
        End If
                
                
    Case sghEmergenciaConsultorios
        
        TituloDeForm "emergencia"
        
        lblNroCamaIngreso.Visible = False: txtNroCamaIngreso.Visible = False: btnVerDisponibilidadDeCamas.Visible = False
        Me.ucTransferenciasDetalle1.OcultarDatosDeCama
        
        
        cmbIdTipoGravedad.Visible = True
        lblGravedad.Visible = True
    
        If mi_Opcion = sghAgregar Then
            'No se debe ver datos de egreso
            'Me.tabAdmision.TabVisible(2) = False
            'Me.tabAdmision.TabsPerRow = 2
            
            'No se debe ver las transferencias
            Me.TabIngreso.TabVisible(3) = False
            Me.TabIngreso.TabVisible(1) = False
            Me.TabIngreso.TabsPerRow = 2
            
            'No se debe ver NACIMIENTOS
            Me.TabIngreso.TabVisible(4) = False
        End If
    
        If ml_TipoAccionAdmision = sghTransferencias Then
                'No se debe ver datos de egreso
                Me.tabAdmision.TabVisible(2) = False
                Me.tabAdmision.TabsPerRow = 2
        
                'Se debe ver las transferencias y lo presenta primero
                Me.tabAdmision.Tab = 1
                Me.TabIngreso.Tab = 1
        End If
        
        If ml_TipoAccionAdmision = sghDarDeAlta Then
                'Se ubica en el tab de egreso
                Me.tabAdmision.Tab = 2
        End If
    
    Case sghEmergenciaObservacion
    
        
        TituloDeForm "emergencia"
        cmbIdTipoGravedad.Visible = False
        lblGravedad.Visible = False
        
        'Causa externa de morbilidad no se debe ver
        TabIngreso.TabVisible(2) = False
        TabIngreso.TabsPerRow = 3
        
        'El acompañante no se debe ver
        Me.txtNombreAcompañante.Visible = False
        Me.lblNombreAcompañante.Visible = False
        fraNotas.Visible = False   'debb-21/06/2016
        
        If ml_TipoAccionAdmision = sghEnviarAObservacion Then
            If mi_Opcion = sghAgregar Then
                'No se debe ver datos de egreso
                Me.tabAdmision.TabVisible(2) = False
                Me.tabAdmision.TabsPerRow = 2
                
                'No se debe ver las transferencias
                Me.TabIngreso.TabVisible(1) = False
                Me.TabIngreso.TabsPerRow = 2
            
                Me.ucDiagnosticosIngreso.TipoDiagnostico = sghHospitalizacionIngreso
                Me.ucDiagnosticosIngreso.CargarDatosDeDiagnosticos oConexion
            
            End If
        End If
        
        If ml_TipoAccionAdmision = sghTransferencias Then
                'No se debe ver datos de egreso
                Me.tabAdmision.TabVisible(2) = False
                Me.tabAdmision.TabsPerRow = 2
        
                'Se debe ver las transferencias
                Me.TabIngreso.TabsPerRow = 1
        End If
        
        If ml_TipoAccionAdmision = sghDarDeAlta Then
                'Se ubica en el tab de egreso
                Me.tabAdmision.Tab = 2
        End If
        
    End Select

    oConexion.Close
    Set oConexion = Nothing
End Sub
Sub TituloDeForm(sTitulo As String)
        
        Select Case mi_Opcion
        Case sghAgregar
            Me.Caption = "Agrega admisión " & sTitulo
        Case sghModificar
            Me.Caption = "Modifica admisión " & sTitulo
        Case sghConsultar
            Me.Caption = "Consulta admisión " & sTitulo
        Case sghEliminar
            Me.Caption = "Elimina admisión " & sTitulo
        End Select

End Sub
'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If lcBuscaParametro.RetornaHoraServidorSQL < wxHoraMadrugada Then
       mo_lbCargaTablasUnaVez = True
   End If
   SiempreCargaPorMovimiento
   setToolTipText
   If mi_Opcion <> sghAgregar Then
        'Me.txtApellidoPaternoBusqueda.SetFocus
        If Not mb_ExistenDatos Then
           Me.Visible = False
           LimpiarVariablesDeMemoria
        End If
        '
        If lbElServicioRegistraFUA = "S" And wxParametro302 = "S" And Val(cmbFuenteFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS Then
        Else
           wxParametro302 = "N"
           'Me.ucSISfuaCodPrestacion1.Visible = False
           'Me.UcSISafiliacion1.Visible = False
           'Me.chkBuscarEnSIS.Visible = False
        End If
        '
        On Error Resume Next
        Select Case lnFocusCuandoCargeFrm
        Case 0  'No se ha realizado ninguna transferencia
             tabAdmision.Tab = 1
             lcCaptionTab2 = TabIngreso.Caption
        Case 1   'Confirmacion de Transferencias
             tabAdmision.Tab = 1
             TabIngreso.Tab = 1
             lcCaptionTab2 = TabIngreso.Caption
        Case 2   'confirmacion de llegada al Servicio
             tabAdmision.Tab = 1
             TabIngreso.Tab = 0
             lcCaptionTab2 = TabIngreso.Caption
        End Select
        lnFocusCuandoCargeFrm = 100
   Else
        If ml_TipoServicio = sghEmergenciaConsultorios Then
               mo_cmbIdTipoGravedad.BoundText = ""
               If wxParametro316 <> "" Then
                  mo_cmbIdTipoGravedad.BoundText = wxParametro316
               End If
               mo_cmbIdCausaExternaMorbilidad.BoundText = ""
               If wxParametro317 <> "" Then
                  mo_cmbIdCausaExternaMorbilidad.BoundText = wxParametro317
               End If
               'Actualizado 31102014
               HabilitarControlesAdmision
        End If
        '
        If lbCargaUnaVezVEntana = True Then
             'mgaray20141023
            Call LimpiarVariablesEnMemoria
             lbCargaUnaVezVEntana = False
             On Error Resume Next
             If ml_TipoServicio = sghHospitalizacion Then
                 Select Case WxDEFAULT_BUSQ_HOSPITALIZ
                 Case sghDefaultVentana.sighApellidoPaterno
                      txtApellidoPaternoBusqueda.SetFocus
                 Case sghDefaultVentana.sighDNI
                      txtNroDNIBusqueda.SetFocus
                 Case sghDefaultVentana.sighHistoria
                      txtNroHistoriaBusqueda.SetFocus
                 End Select
             Else
                 Select Case WxDEFAULT_BUSQ_EMERGENCIA
                 Case sghDefaultVentana.sighApellidoPaterno
                      txtApellidoPaternoBusqueda.SetFocus
                 Case sghDefaultVentana.sighDNI
                      txtNroDNIBusqueda.SetFocus
                 Case sghDefaultVentana.sighHistoria
                      txtNroHistoriaBusqueda.SetFocus
                 End Select
             End If
        End If
   End If
   
End Sub



Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    'Case vbKeyEscape
    '    btnCancelar_Click
    Case vbKeyF2
        btnAceptar_Click
    Case vbKeyF3
        btnImprimir_Click
    Case vbKeyF5
        btnLimpiar_Click
    Case vbKeyF6
            btnBuscarPaciente_Click
     Case vbKeyF7
         Me.tabAdmision.Tab = 0
         Me.ucPacientesDetalle1.SetPestaniaTabPaciente 0
         On Error Resume Next
         Me.ucPacientesDetalle1.SetFocusOnDepartamentoDomicilio
     Case vbKeyF8
         Me.tabAdmision.Tab = 0
         Me.ucPacientesDetalle1.SetPestaniaTabPaciente 1
         On Error Resume Next
         Me.ucPacientesDetalle1.SetFocusOnDepartamentoProcedencia
     Case vbKeyF9
         Me.tabAdmision.Tab = 0
         On Error Resume Next
         Me.ucPacientesDetalle1.SetPestaniaTabPaciente 2
         Me.ucPacientesDetalle1.SetFocusOnDepartamentoNacimiento
     Case vbKeyF10
         Me.tabAdmision.Tab = 0
         On Error Resume Next
         Me.ucPacientesDetalle1.SetFocusOnApellidoPaterno
     Case vbKeyF11
         LimpiaDatosDeBusqueda
         Me.tabAdmision.Tab = 1
         On Error Resume Next
         tabAdmision_Click 0
    End Select
       
End Sub




Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Dim oConexion As New Connection
   oConexion.CommandTimeout = 300
   oConexion.CursorLocation = adUseClient
   oConexion.Open sighEntidades.CadenaConexion
   
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If Not ValidarDiasInternamiento() Then Exit Sub
               If AgregarDatos() Then
               
               
                    Me.idAtencion = mo_Atenciones.idAtencion
                    Me.txtNroCuenta = mo_Atenciones.idCuentaAtencion
                    lcApP = VerSiTieneServicioAutomaticoPorEstancia(oConexion)
                    MsgBox "Los datos se agregaron correctamente para la Historia Nª:  " & _
                    HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(mo_Pacientes.NroHistoriaClinica)), False) & _
                    Chr(13) & Chr(13) & "N° Cuenta: " & Me.txtNroCuenta.Text & Chr(13) & Chr(13) & lcApP, vbInformation, Me.Caption
                    mo_AdminArchivoClinico.generadorNroHistoriaClinicaActualizaNroAutomaticoDeHistoriaClinica oConexion
                    Me.btnImprimir.Enabled = True
                    Me.btnAceptar.Enabled = False
                    'A.Yañez  06-01-2014 ********************************
                    Me.btnNuevoAdmisionHospDetalle.Visible = True
                    '****************************************************
                    If Me.chkPacienteNuevo.Value = 1 Then
                       btnImprimeFiliacion.Enabled = True
                    End If
                    Me.btnImprimir.SetFocus
                    If wxParametro357 = "S" Then
                        If txtNroCuenta.Text <> "" Then ImprimePreCuenta 'Impresion directa del ticket de admision en Hospitalizacion y emergencia
                    End If
                    If ml_TipoServicio = sghEmergenciaConsultorios Then
                        If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS And lcElServicioUsaGalenHos = "S" Then
                             btnImprimeFichaSIS_Click
                             btnImprimeFichaSIS.Visible = True
                        End If
                    End If
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
       
   Case sghModificar
        'A.Yañez  06-01-2014 ********************************
        Me.btnNuevoAdmisionHospDetalle.Visible = False
        '***************************************************
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
           If ValidarReglas() Then
               If Not ValidarDiasInternamiento() Then Exit Sub
               If ModificarDatos() Then
                   
                   'ImprimeFormularioEnPDF
                   
                   MsgBox " Los datos se modificaron correctamente, para la Cuenta N° " & Me.txtNroCuenta.Text & VerSiTieneServicioAutomaticoPorEstancia(oConexion), vbInformation, Me.Caption
                   If wxParametro357 = "S" Then
                        If txtNroCuenta.Text <> "" Then ImprimePreCuenta 'Impresion directa del ticket de admision en Hospitalizacion y Emergencia
                   End If
                   Me.Visible = False
                   LimpiarVariablesDeMemoria
                   If ml_TipoServicio = sghEmergenciaConsultorios Then
                        If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS And lcElServicioUsaGalenHos = "S" Then
                             btnImprimeFichaSIS_Click
                        End If
                    End If
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
       HabilitarControlesAdmision
   Case sghEliminar
        'A.Yañez  06-01-2014 ********************************
         Me.btnNuevoAdmisionHospDetalle.Visible = False
        '****************************************************
               CargaDatosAlObjetosDeDatos
               If EliminarDatos(oConexion) Then
                   MsgBox " Los datos se eliminaron correctamente, para la Cuenta N° " & Me.txtNroCuenta.Text, vbInformation, Me.Caption
                   Me.Visible = False
                   LimpiarVariablesDeMemoria
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
               End If
   End Select
   ActualizaCodigoPrestacionEnCE oConexion
   oConexion.Close
   Set oConexion = Nothing
End Sub

Function VerSiTieneServicioAutomaticoPorEstancia(oConexion As Connection) As String
    Dim lcSql As String
    Dim oRsTmp As New ADODB.Recordset
    VerSiTieneServicioAutomaticoPorEstancia = ""
    Set oRsTmp = mo_AdminFacturacion.FactOrdenServicioPagosPorIdAtencion(mo_Atenciones.idAtencion, oConexion)
    If oRsTmp.RecordCount > 0 Then
       If mo_Atenciones.idTipoServicio = 3 Then
          oRsTmp.Filter = "idPuntoCarga=9"
       Else
          oRsTmp.Filter = "idPuntoCarga=10"
       End If
    End If
    
    If oRsTmp.RecordCount > 0 Then
       VerSiTieneServicioAutomaticoPorEstancia = Chr(13) & "(Ord.Pago)= "
       oRsTmp.MoveFirst
       txtNroOrdenPago.Text = oRsTmp.Fields!IdOrdenPago
       Do While Not oRsTmp.EOF
          VerSiTieneServicioAutomaticoPorEstancia = VerSiTieneServicioAutomaticoPorEstancia & Trim(Str(oRsTmp.Fields!IdOrdenPago)) & " , "
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    If VerSiTieneServicioAutomaticoPorEstancia = "" Then
       txtNroOrdenPago.Text = "Pagó Consulta"
    End If
End Function

Function ValidarDiasInternamiento() As Boolean
    ValidarDiasInternamiento = False

    ValidarDiasInternamiento = True
End Function

Private Sub btnCancelar_Click()
   Dim lbSale As Boolean
   If sighEntidades.ParaAuditoria = "" Then
      lbSale = True
   ElseIf MsgBox("Hubo cambios, desea salir de todas maneras ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
      lbSale = True
   End If
   If lbSale Then
   
   Me.Visible = False
   'mgaray20141003
   lcDniSIS = ""
   LimpiarVariablesDeMemoria
   cmbIdViasAdmision.Enabled = True
   lblNombreServicio.Enabled = True
   lblNombreMedico.Enabled = True
   txtNroCamaIngreso.Enabled = True
   cmbIdTipoGravedad.Enabled = True
   cmbFuenteFinanciamiento.Enabled = True
   cmbFormaPago.Enabled = True
   txtFechaIngreso.Enabled = True
   txtHoraIngreso.Enabled = True
   'A.Yañez  06-01-2014 ********************************
   Me.btnNuevoAdmisionHospDetalle.Visible = False
   '****************************************************
   End If
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   
   '-------------------------------------------------------------------------
   '                VALIDA DATOS DE LA CUENTA DE ATENCION
   '-------------------------------------------------------------------------
   If Me.txtIdMedicoIngreso.Text = "" Then
        sMensaje = sMensaje + "Ingrese el médico de ingreso" + Chr(13)
   End If
   If Me.cmbFuenteFinanciamiento.Text = "" Then
       sMensaje = sMensaje + "Elija el Plan de Atención" + Chr(13)
   End If
   
    '---------------------------------------------------------------------------------
    '           VALIDA DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
   If Me.txtIdServicioIngreso = "" Then
       sMensaje = sMensaje + "Ingrese el valor del servicio de ingreso" + Chr(13)
   End If
   If Me.txtHoraIngreso.Text = "" Then
       sMensaje = sMensaje + "Ingrese el valor de HoraIngreso" + Chr(13)
   End If
   If Me.txtFechaIngreso.Text = sighEntidades.FECHA_VACIA_DMY Then
       sMensaje = sMensaje + "Ingrese el valor de la fecha de ingreso" + Chr(13)
   End If
   If Val(mo_cmbIdTipoServicio.BoundText) = 0 Then
       sMensaje = sMensaje + "Ingrese el valor del tipo de servicio" + Chr(13)
   End If
   '
   CambioFechaNacimiento Me.ucPacientesDetalle1.DevuelveFechaNacimiento, Me.ucPacientesDetalle1.DevuelveHoraNacimiento
   If Val(Me.txtEdadEnDias.Text) = 0 Then
       sMensaje = sMensaje + "Ingrese el valor de la edad" + Chr(13)
   End If
   
   
   If mi_Opcion = sghModificar Then
        If chkLlegoSI.Visible = True And txtNroCamaIngreso.Text = "" Then
           sMensaje = sMensaje + "Por favor asigne la Cama (Ficha 2.1)" + Chr(13)
        End If
        If txtCamaTransf.Visible = True And txtCamaTransf.Text = "" Then
           sMensaje = sMensaje + "Por favor asigne la Cama que fué transferido (Ficha 2.2)" + Chr(13)
        End If
   End If
   If cmbFormaPago.Text = "" Then
      sMensaje = sMensaje + "Por favor elija el Tipo Financiamiento (Ficha 2.1)" + Chr(13)
      Me.tabAdmision.Tab = 1
      TabIngreso.Tab = 0
   End If
   
   If ml_TipoServicio = sghEmergenciaConsultorios Then
      If cmbComoLlego.Text = "" Then
         sMensaje = sMensaje + "Por favor elija COMO LLEGO EL PACIENTE (Ficha 2.1)" + Chr(13)
         Me.tabAdmision.Tab = 1
         Me.TabIngreso.Tab = 0
      End If
      If cmbTipoAtencion.Text = "" Then
         sMensaje = sMensaje + "Por favor elija TIPO ATENCION (Ficha 2.1)" + Chr(13)
         Me.tabAdmision.Tab = 1
         Me.TabIngreso.Tab = 0
      End If
   End If
   
   If dxMorbilidadExterna342 = "S" And Trim(cmbIdCausaExternaMorbilidad.Text) = "" And ml_TipoServicio = sghEmergenciaConsultorios Then
        sMensaje = sMensaje + "Por favor elija Causa Externa de Morbilidad (Ficha 2.3)" + Chr(13)
         Me.tabAdmision.Tab = 1
         Me.TabIngreso.Tab = 2
    End If
    '
    If ml_TipoServicio = sghEmergenciaConsultorios And wxParametro547 = "S" And _
                             Val(cmbFuenteFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS And _
                                                wxParametro302 <> "S" And txtNroAfiliacionSis.Text = "" Then
      sMensaje = sMensaje + "Por favor registre el N° FILIACION (SIS) (Ficha 2.1)" + Chr(13)
      Me.tabAdmision.Tab = 1
      TabIngreso.Tab = 0
    
    End If
   '
   '---------------------------------------------------------------------------------
   '           VALIDA DATOS DE PACIENTES
   '---------------------------------------------------------------------------------
   sMensaje = sMensaje + ucPacientesDetalle1.ValidarDatosObligatorios(wxParametro282, wxParametro333)
'   If sMensaje <> "" Then
'        If ucPacientesDetalle1.DevuelveEtnia = "" Then
'           Me.tabAdmision.Tab = 0
'           ucPacientesDetalle1.SetFocusEnEtnia
'        ElseIf ucPacientesDetalle1.DevuelveIdioma = "" Then
'           Me.tabAdmision.Tab = 0
'           ucPacientesDetalle1.SetFocusEnIdioma
'        End If
'   End If
   
   If wxParametro521 = "S" And ml_TipoServicio = sghEmergenciaConsultorios And cmbEstadoLlegada.Text = "" Then
      sMensaje = sMensaje + "Por favor elija ESTADO de llegada del Paciente (Ficha 2.1)" + Chr(13)
   End If


   '9
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
    Dim lcMensaje As String
    Dim rsCitas As New Recordset
    ValidarReglas = False
    
    If mo_Pacientes.idTipoNumeracion > 3 And ml_TipoServicio = sghHospitalizacion Then
       MsgBox "Solo puede usar HISTORIA FINAL", vbInformation, Me.Caption
       Exit Function
    End If
    
    If mi_Opcion = sghAgregar And mo_Pacientes.idPaciente > 0 Then
             lcMensaje = mo_AdminFacturacion.DevuelveSiElPacienteFallecioOhistoriaPasoPasivo(mo_Pacientes.idPaciente)
             If lcMensaje <> "" Then
                MsgBox lcMensaje, vbInformation, Me.Caption
                Me.tabAdmision.Tab = 0
                Exit Function
             End If
    End If
    
    If Not Me.ucPacientesDetalle1.ValidarReglas(mo_Pacientes) Then
        Me.tabAdmision.Tab = 0
        Exit Function
    End If
    
    If ChequeaQueNoExistaPacienteServicioFecha = False Then
       Exit Function
    End If
    
    'Menor SIN ACOMPAÑANTE
    If ml_TipoServicio = sghEmergenciaConsultorios And sighEntidades.DevuelveEdadEnMeses(mo_Pacientes.FechaNacimiento, mo_Atenciones.FechaIngreso) <= (18 * 12 - 1) Then
       If txtNombreAcompañante.Text = "" Then
            MsgBox "El Paciente es MENOR DE EDAD, registre el ACOMPAÑANTE (ficha 2.1)", vbInformation, Me.Caption
            Exit Function
'       ElseIf Me.txtDNIacompaniante.Text = "" Then
'            MsgBox "El Paciente es MENOR DE EDAD, registre el DNI del ACOMPAÑANTE (ficha 2.1)", vbInformation, Me.Caption
'            Exit Function
       End If
    End If
    
    If Not mo_AdminAdmision.ValidaEdadMaximaYSexoSegunServicioHosp(Val(txtEdadEnDias.Text), _
                  Val(mo_cmbIdTipoEdad.BoundText), mo_Pacientes.idTipoSexo, mo_Atenciones.IdServicioEgreso, True) Then
        Exit Function
    End If
    
   
    If Val(Me.txtEdadEnDias) > 130 Then
        If MsgBox("La edad es de mas de 130 años, ¿es correcto?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Function
        End If
    End If
    
    'debb-16/03/2016 (inicio)
    
    If CDate(Me.txtFechaIngreso.Text) > ldFechaActualServidor Then
       ldFechaActualServidor = lcBuscaParametro.RetornaFechaServidorSQL
       If CDate(Me.txtFechaIngreso.Text) > ldFechaActualServidor Then
            MsgBox "La FECHA DE INGRESO, no puede ser mayor a HOY", vbInformation, Me.Caption
            Exit Function
       End If
    End If
    'debb-16/03/2016 (fin)
   
    Dim lIdCausaBasica As Long
    Dim lIdCausaIntermedia As Long
    Dim lIdCausaFinal As Long
    Dim lIdDxPrincipal As Long
    Dim lIdDxIngreso As Long
    
    ObtenerDiagnosticos lIdDxPrincipal, lIdCausaBasica, lIdCausaIntermedia, lIdCausaFinal, lIdDxIngreso
      
    '
    If ml_TipoServicio = sghHospitalizacion Then
       If mi_Opcion = sghModificar And chkLlegoSI.Visible = True And mo_Diagnosticos.Count = 0 Then  'debb-18/05/2016
          MsgBox "Por favor asigne el Diagnóstico de INGRESO", vbInformation, Me.Caption
          Me.tabAdmision.Tab = 1
          Me.TabIngreso.Tab = 0
          Exit Function
       End If
    Else
       If wxParametro506 = "S" Then
            If BuscarSiExisteNrocorrelativoEmergencia = True Then
               Exit Function
            End If
       End If
    End If
    
    'Verifica si algunos de los servicios es de cirugia, ginecologia u obstetricia
    Dim bServPerteneceACirugiaOGinecologia As Boolean
    Dim bServicioPerteneceAPediatria As Boolean
    bServPerteneceACirugiaOGinecologia = False
    bServicioPerteneceAPediatria = False
    Dim oDOOcupacion As New DOEstanciaHospitalaria
    Dim oDoServicio As New doServicio
    Dim lIdDepartamento As Long
    
    For Each oDOOcupacion In mo_OcupacionCamas
        lIdDepartamento = mo_AdminServiciosHosp.ServiciosSeleccionarIdDepartamento(oDOOcupacion.IdServicio)
        If lIdDepartamento = 3 Or lIdDepartamento = 4 Then  'Cirugia y (Ginecologia y Obstetricia)
            bServPerteneceACirugiaOGinecologia = True
        End If
        If lIdDepartamento = 2 Then  'Pediatria
            bServicioPerteneceAPediatria = True
        End If
        
    Next
    If mi_Opcion = sghModificar Then
        If mo_NroServiciosQuePasoElPaciente < mo_OcupacionCamas.Count And _
                          lcUltimoCodigoDeServicioTransferido = ml_lcServicioEgreso Then
           MsgBox "No puede transferir al mismo SERVICIO", vbInformation, Me.Caption
           Exit Function
        End If
    End If
    
    'Validar que la cama no este ocupada
    Dim rsCamas As New Recordset
    If mi_Opcion = sghModificar And mo_NroServiciosQuePasoElPaciente = 1 Then
        If Me.txtNroCamaIngreso.Tag <> "" Then
            Set rsCamas = mo_AdminReglasHoteleria.CamasSeleccionarDisponibilidadPorServicioUbicacionActual(CLng(Val(Me.txtIdServicioIngreso.Tag)))
            If rsCamas.RecordCount > 0 Then
                rsCamas.MoveFirst
                Do While Not rsCamas.EOF
                    If Val(Me.txtNroCamaIngreso.Tag) = rsCamas.Fields!idCama Then
                        If mo_Pacientes.NroHistoriaClinica <> rsCamas.Fields!NroHistoriaClinica Then
                            MsgBox "La cama " + Trim(rsCamas.Fields!Codigo) + " se encuentra ocupada por (HC: " + CStr(rsCamas.Fields!NroHistoriaClinica) + " " + rsCamas.Fields!ApellidoPaterno + " " + rsCamas.Fields!ApellidoMaterno + " " + _
                                    rsCamas.Fields!PrimerNombre + " " + IIf(IsNull(rsCamas.Fields!SegundoNombre), "", rsCamas.Fields!SegundoNombre) + ")", vbInformation, Me.Caption
                            Set rsCamas = Nothing
                            Exit Function
                        End If
                    End If
                    rsCamas.MoveNext
                Loop
            End If
        End If
    '        Me.idPaciente
    End If
    
    'Si tiene algun servicio del dpto de cirugia , ginecologia y obstetricia
    If bServPerteneceACirugiaOGinecologia Then
        If mo_Procedimientos.Count = 0 Then
            

        End If
    End If
    
    'Valida la infeccion intrahospitalaria
    If Not HuboDiagnosticoInfeccionAlIngreso() Then
        If HayDiagnosticosInfeccionAlEgreso() Then
            If MsgBox("Los diagnosticos muestran la existencia de infección intrahospitalaria, ¿Es correcto?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                mo_DoAtencionDatosAdicionales.HuboInfeccionIntraHospitalaria = False
            Else
                mo_DoAtencionDatosAdicionales.HuboInfeccionIntraHospitalaria = True
            End If
        End If
    End If
    
    If bServicioPerteneceAPediatria Then
        If Val(mo_cmbIdTipoEdad.BoundText) = 1 Then
            If Val(Me.txtEdadEnDias) >= 18 Then
                If MsgBox("El paciente es mayor de 18 años, y esta en el servicio de Pediatría, ¿es correcto?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End If
    'refencia INGRESO
    If cmbIdViasAdmision.Text <> "" Then
        Dim sCodigoOrigen As String
        sCodigoOrigen = Trim(Split(cmbIdViasAdmision.Text, " = ")(0))
        If sCodigoOrigen = "R" Or sCodigoOrigen = "C" Then
           If lblNombreOrigenReferencia.Text = "" Then
              MsgBox "Debe elejir el ESTABLEC.REFERIDO (Origen)(ficha 2.1)", vbInformation, Me.Caption
              Me.tabAdmision.Tab = 1
              Me.TabIngreso.Tab = 0
              Exit Function
           End If
           If txtReferenciaO.Text = "" Then
              MsgBox "Debe registrar el N°REFERENCIA (Origen)(ficha 2.1)", vbInformation, Me.Caption
              Me.tabAdmision.Tab = 1
              Me.TabIngreso.Tab = 0
              Exit Function
           End If
           If mo_Diagnosticos.Count = 0 Then
              MsgBox "Debe registrar el DX DE INGRESO (de la REFERENCIA) (ficha 2.1)", vbInformation, Me.Caption
              Me.tabAdmision.Tab = 1
              Me.TabIngreso.Tab = 0
              Exit Function
           End If
           If cmbServicioReferenciaO.Text = "" Then
              MsgBox "Debe elegir el SERVICIO DE LA REFERENCIA (ficha 2.1)", vbInformation, Me.Caption
              Me.tabAdmision.Tab = 1
              Me.TabIngreso.Tab = 0
              Exit Function
           End If
           'FRANKLIN 2017
           If ml_TipoServicio = sghEmergenciaConsultorios And lcBuscaParametro.SeleccionaFilaParametro(516) = "S" And _
                                                             (txtMedicoRef.Text = "" Or Me.cmbMedicoRef.Text = "") Then
               MsgBox "Por favor debe ingresar: COLEGIATURA, APELLIDOS Y NOMBRES DEL MEDICO QUE REFIERE (Ficha 'Cita')" + Chr(13), vbInformation, Me.Caption
           End If
           
        End If
        'Valida que el recien nacido, si ha nacido en el Hospital,  tenga asociada la Cuenta de la MADRE
        If ml_TipoServicio = sghHospitalizacion And (sCodigoOrigen = "J" Or sCodigoOrigen = "N") And Year(mo_Atenciones.FechaIngreso) > 2010 Then
                                                          '**ojo***eliminar esta validacion 2010 cuando sea necesaria
           If lnIdNacimientoSeleccionado = 0 And (lblMadre.Text = "" And mo_Pacientes.Nombremadre = "") Then       'debb-05/03/12
                 MsgBox "Por favor ingrese el 'Nombre de la Madre'", vbInformation, Me.Caption
                 Me.tabAdmision.Tab = 1
                 Me.TabIngreso.Tab = 0
                 Exit Function
           End If
        End If
        
    End If
    If mi_Opcion = sghAgregar And mb_EsObservacionEmergencia = True And ml_TipoServicio = sghEmergenciaConsultorios And Me.txtNroCamaIngreso.Text = "" Then  '09/08/2011
        MsgBox "Por favor seleccione la CAMA de Observación de Emergencia", vbInformation, Me.Caption
        Me.tabAdmision.Tab = 1
        Me.TabIngreso.Tab = 0
        Exit Function
    End If
    
    If cmbIdTipoGravedad.Text = "" And ml_TipoServicio = sghEmergenciaConsultorios And Val(wxParametro316) > 0 Then
        Me.tabAdmision.Tab = 1
        Me.TabIngreso.Tab = 0
        MsgBox "Debe elegir la GRAVEDAD" & Chr(13) & "ficha '2.1-Ingreso'", vbExclamation, Me.Caption
        Exit Function
    End If
    
    If mi_Opcion = sghModificar And ml_TipoServicio = sghEmergenciaConsultorios Then   '09/08/2011
       Dim lcSistolica As String, lcDiastolica As String
'       If Me.txtPresion.Text <> "" Then
'            lcSistolica = Left(Me.txtPresion.Text, InStr(Me.txtPresion.Text, "/") - 1)
'            lcDiastolica = Mid(Me.txtPresion.Text, InStr(Me.txtPresion.Text, "/") + 1, 100)
'            If Val(lcSistolica) < Val(lcDiastolica) Then
'                MsgBox "En la Presión: Sistolica debe ser mayor a la Diastólica", vbInformation, Me.Caption
'                Exit Function
'            End If
'       End If
'       If Me.txtTemperatura.Text <> "" And Not (Val(Me.txtTemperatura.Text) >= 35 And Val(Me.txtTemperatura.Text) <= 42) Then
'            MsgBox "La Temperatura debe estar entre 35 y 42 °C ", vbInformation, Me.Caption
'            Exit Function
'       End If
'       If Val(Me.txtPulso.Text) > 250 Then
'            Me.tabAdmision.Tab = 1
'            Me.TabIngreso.Tab = 0
'            MsgBox "El PULSO no debe pasar de 250" & Chr(13) & "ficha '2.1-Ingreso'", vbExclamation, Me.Caption
'            Exit Function
'       End If
'       If Val(Me.txtFrespiratoria.Text) > 70 Then
'            Me.tabAdmision.Tab = 1
'            Me.TabIngreso.Tab = 0
'            MsgBox "La FRECUENCIA RESPIRATORIA no debe pasar de 70" & Chr(13) & "ficha '2.1-Ingreso'", vbExclamation, Me.Caption
'            Exit Function
'       End If
       If Trim(cmbIdCausaExternaMorbilidad.Text) = "" Then
            Me.tabAdmision.Tab = 1
            Me.TabIngreso.Tab = 2
            MsgBox "Debe elegir CAUSA EXTERNA DE MORBILIDAD" & Chr(13) & "ficha '2.3-Causas Externas morbilidad'", vbExclamation, Me.Caption
            Exit Function
       End If
    End If
    If mi_Opcion = sghAgregar And lcCodigoEstablecimientoAdscripcionSIS <> "" _
                              And Val(cmbFuenteFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS _
                              And ml_TipoServicio <> sghHospitalizacion Then
       If mo_cmbIdViasAdmision.BoundText = "21" Then
            lcMensaje = mo_ReglasSISgalenhos.ChequeaCodigoEstablecimientoAdscripcion(lcCodigoEstablecimientoAdscripcionSIS, _
                                                Val(mo_cmbIdTipoServicio.BoundText), _
                                                mo_AdminAdmision.TiposOrigenAtencionDevuelveIdSis(Val(mo_cmbIdViasAdmision.BoundText)), _
                                                "")
            If lcMensaje <> "" Then
                  MsgBox lcMensaje, vbInformation, Me.Caption
                 ''''' Frank 2608
                 CargarAutomaticamenteEstablecimientoReferenciaSIS
                 '''
                  
                 Exit Function
            End If
       End If
    End If
    If wxParametro302 = "S" And Val(cmbFuenteFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS And _
                                                                                                mi_Opcion = sghEliminar Then
            Set rsCitas = mo_ReglasSISgalenhos.SisFuaAtencionSeleccionarPorCuenta(Val(Me.txtNroCuenta.Text))
            If rsCitas.RecordCount > 0 Then
               MsgBox "El formato FUA ya fué generado: " & rsCitas.Fields!fuaDisa & "-" & rsCitas!fuaLote & "-" & _
                      rsCitas!FuaNumero & Chr(13) & "Debe eliminar el formato FUA (módulo: SIS, opción: Formato FUA)", _
                      vbInformation, Me.Caption
               Exit Function
            End If
    End If
    
    'kike 2017
    
    If mo_Pacientes.IdPaisDomicilio = 166 And (mo_Pacientes.IdDistritoDomicilio = 0 Or mo_Pacientes.DireccionDomicilio = "") And _
                                 Me.ucPacientesDetalle1.PacienteNoIdentificado = False Then
      MsgBox "Por favor elija el DISTRITO DEL DOMICILIO y/o DIRECCION  (Ficha 1)", vbInformation, Me.Caption
      Me.tabAdmision.Tab = 0
      Me.ucPacientesDetalle1.SetPestaniaTabPaciente 0
      On Error Resume Next
      Me.ucPacientesDetalle1.SetFocusOnDepartamentoDomicilio
      SendKeys "{tab}"
      Exit Function
    End If
    '
    Set oDOOcupacion = Nothing
    Set oDoServicio = Nothing
    Set rsCitas = Nothing
    ValidarReglas = True
End Function

Public Sub CargarAutomaticamenteEstablecimientoReferenciaSIS() 'Frank 2808
    If lcBuscaParametro.SeleccionaFilaParametro(326) = "S" And lcCodigoEstablecimientoAdscripcionSIS <> "" _
       And ml_TipoServicio = sghEmergenciaConsultorios Then
       Dim lcCodigoSis As String
       Dim lcEstablecimientoOrigen As String
       Dim DOEstablecimiento As New DOEstablecimiento
       Dim oRsEstabNoMINSA As Recordset
       Dim lnIdOrigenDelPacienteDesdeFUA As Long
       lnIdOrigenDelPacienteDesdeFUA = mo_AdminAdmision.TiposOrigenAtencionDevuelveIdSis(Val(mo_cmbIdViasAdmision.BoundText))
       
       If Val(lcBuscaParametro.SeleccionaFilaParametro(280)) <> Val(lcCodigoEstablecimientoAdscripcionSIS) Then
          If lcBuscaParametro.SeleccionaFilaParametro(282) <> "S" Then 'Hospital
               'If Not (lnIdOrigenDelPacienteDesdeFUA = "4" Or lnIdOrigenDelPacienteDesdeFUA = "6") Then 'Referido CE, ContraReferido
                    'mo_cmbIdViasAdmision.BoundText = "21"
                    If mo_AdminServiciosComunes.EstablecimientosSeleccionarPorCodigo(Right(lcCodigoEstablecimientoAdscripcionSIS, 5), DOEstablecimiento) = True Then
                        mo_cmbIdTipoReferenciaOrigen.BoundText = 1 'MINSA
                        txtIdEstablecimientoOrigen.Text = DOEstablecimiento.Codigo
                        txtIdEstablecimientoOrigen.Tag = DOEstablecimiento.IdEstablecimiento
                        lblNombreOrigenReferencia.Text = DOEstablecimiento.nombre
                    Else
                        Set oRsEstabNoMINSA = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorCodigo(Right(lcCodigoEstablecimientoAdscripcionSIS, 5))
                        If oRsEstabNoMINSA.RecordCount > 0 Then
                            oRsEstabNoMINSA.MoveFirst
                            mo_cmbIdTipoReferenciaOrigen.BoundText = 2 'NO MINSA
                            txtIdEstablecimientoOrigen.Text = oRsEstabNoMINSA.Fields!Codigo
                            txtIdEstablecimientoOrigen.Tag = oRsEstabNoMINSA.Fields!IdEstablecimientoNoMINSA
                            lblNombreOrigenReferencia.Text = oRsEstabNoMINSA.Fields!nombre
                        End If
                        Set oRsEstabNoMINSA = Nothing
                    End If
               'End If
          End If
       End If
       Set DOEstablecimiento = Nothing
    End If
End Sub


Function HayDiagnosticosInfeccionAlEgreso() As Boolean
Dim oDOAtencionDiagnostico As DOAtencionDiagnostico
Dim oDODiagnostico As DODiagnostico

    HayDiagnosticosInfeccionAlEgreso = False

    For Each oDOAtencionDiagnostico In mo_Diagnosticos
        If oDOAtencionDiagnostico.IdClasificacionDx = 3 Then    'Diagnostico de egreso
            Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oDOAtencionDiagnostico.idDiagnostico)
            If oDODiagnostico.Intrahospitalario Then
                HayDiagnosticosInfeccionAlEgreso = True
            End If
        End If
    Next

End Function
Function HuboDiagnosticoInfeccionAlIngreso() As Boolean
Dim oDODiagnostico As New DODiagnostico
Dim oDOAtencionDiagnostico As DOAtencionDiagnostico

    HuboDiagnosticoInfeccionAlIngreso = False
    
    For Each oDOAtencionDiagnostico In mo_Diagnosticos
        If oDOAtencionDiagnostico.IdClasificacionDx = 2 Then    'Diagnostico de ingreso
            Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oDOAtencionDiagnostico.idDiagnostico)
            If oDODiagnostico.Intrahospitalario Then
                HuboDiagnosticoInfeccionAlIngreso = True
            End If
        End If
    Next
    Set oDODiagnostico = Nothing

End Function
Function ObtenerDiagnosticos(lIdDxPrincipal As Long, lIdCausaBasica As Long, lIdCausaIntermedia As Long, lIdCausaFinal As Long, lIdDxIngreso)
    
    lIdDxPrincipal = 0
    lIdCausaBasica = 0
    lIdCausaIntermedia = 0
    lIdCausaFinal = 0
    lIdDxIngreso = 0
    
    Dim oDODiagnostico As DOAtencionDiagnostico
    For Each oDODiagnostico In mo_Diagnosticos
        If oDODiagnostico.IdSubclasificacionDx = 301 Then
            lIdDxPrincipal = oDODiagnostico.idDiagnostico
            lIdDxIngreso = oDODiagnostico.idDiagnostico
        End If
        If oDODiagnostico.IdSubclasificacionDx = 303 Then lIdCausaFinal = oDODiagnostico.idDiagnostico
        If oDODiagnostico.IdSubclasificacionDx = 304 Then lIdCausaIntermedia = oDODiagnostico.idDiagnostico
        If oDODiagnostico.IdSubclasificacionDx = 305 Then lIdCausaBasica = oDODiagnostico.idDiagnostico
        If oDODiagnostico.IdSubclasificacionDx = 0 Then lIdDxIngreso = oDODiagnostico.idDiagnostico
    Next
    
End Function


Sub CargaDatosAtencionJamo()
    Dim oDODiagnostico As New DODiagnostico
    Set oDODiagnostico = mo_AdminFacturacion.DevuelveDxAltaMedica(mo_Atenciones.idAtencion, 1)
    With mo_DOAtencionesCE
        .IdUsuarioAuditoria = mo_Atenciones.IdUsuarioAuditoria
        .NroHistoriaClinica = mo_Pacientes.NroHistoriaClinica
        .TriajeEdad = txtEdadEnDias.Text
        If IsNull(.idAtencion) Then
           .TriajeFecha = lcBuscaParametro.RetornaFechaHoraServidorSQL
           .TriajeIdUsuario = mo_Atenciones.IdUsuarioAuditoria
           
        End If
        .CitaTratamiento = TxtCitaTratamiento.Text
        Call mo_AdminServiciosComunes.cargarDatosTriajeAObjetoDatos(mo_DOAtencionesCE, ucTriajeVisorCE.DOAtencionCE)
'        .TriajePeso = Me.txtPeso.Text
'        .TriajePresion = Me.txtPresion.Text
'        .TriajeTalla = Me.txtTalla.Text
'        .TriajeTemperatura = Me.txtTemperatura.Text
'        .TriajePulso = Val(Me.txtPulso.Text)
'        .TriajeFrecRespiratoria = Val(Me.txtFrespiratoria.Text)
    End With
    Set oDODiagnostico = Nothing
End Sub



'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Sub CargaDatosAlObjetosDeDatos()
    'Limpia Dx
    Set mo_Diagnosticos = Nothing
    'Limpia Nacimientos
    Set mo_Nacimientos = Nothing
    'Limpia Transferencias
    Set mo_OcupacionCamas = Nothing
    '
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA CUENTA ATENCION
        '---------------------------------------------------------------------------------
    Select Case ml_TipoAccionAdmision
    Case sghAdmisionNormal  'Si el una admisión normal de hospitalizacion o de emergencia
        Select Case mi_Opcion
        Case sghAgregar
            With mo_CuentasAtencion
                        .idCuentaAtencion = Me.idCuentaAtencion
                        .idPaciente = Me.idPaciente
                        .TotalAsegurado = 0
                        .TotalExonerado = 0
                        .TotalPagado = 0
                        .TotalPorPagar = 0
                        If ml_TipoServicio = sghHospitalizacion Then
                             If wxParametro232 = "S" Then
                                .idEstado = sghEstadoCuenta.sghNoLlegaAlServicioHospitalizado
                             Else
                                .idEstado = sghEstadoCuenta.sghAbierto
                             End If
                        Else
                             .idEstado = sghEstadoCuenta.sghAbierto   'Estado de la cuenta = ABIERTO
                        End If
                        'WCG 10/06
                        .FechaApertura = IIf(Me.txtFechaIngreso.Text = sighEntidades.HORA_VACIA_HM, "", Me.txtFechaIngreso.Text)
                        .HoraApertura = IIf(Me.txtHoraIngreso.Text = sighEntidades.HORA_VACIA_HM, "", Me.txtHoraIngreso.Text)
                        .fechaCierre = 0
                        .HoraCierre = ""
                        .IdUsuarioAuditoria = ml_idUsuario
            End With
        Case sghModificar
            '******confirmaron que llego al Servicio (desde Adm.Emerg)
            If chkLlegoSI.Visible = True And chkLlegoSI.Value = 1 Then
               mo_CuentasAtencion.idEstado = sghEstadoCuenta.sghAbierto
            End If
            '******Transferencias
            If lcUltimoCodigoDeServicioTransferido <> "" Then
               If ml_TipoServicio = sghEmergenciaConsultorios Then
                    Dim oDOServicioH1 As New doServicio
                    Set oDOServicioH1 = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(lcUltimoCodigoDeServicioTransferido)
                    If oDOServicioH1.EsObservacionEmergencia = True Then
                       mo_CuentasAtencion.idEstado = sghEstadoCuenta.sghNoLlegaAlServicioHospitalizado
                    Else
                       mo_CuentasAtencion.idEstado = 1
                    End If
                    Set oDOServicioH1 = Nothing
               Else
                    mo_CuentasAtencion.idEstado = sghEstadoCuenta.sghNoLlegaAlServicioHospitalizado
               End If
            ElseIf chkLlegoSS.Visible = True Then
               If chkLlegoSS.Value = 1 And txtCamaTransf.Text <> "" Then
                  mo_CuentasAtencion.idEstado = sghEstadoCuenta.sghAbierto
               End If
            End If
        End Select
    End Select
   
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
   With mo_Atenciones
           .idAtencion = Me.idAtencion
           .IdEspecialidadMedico = ml_IdEspecialidad
           .IdMedicoIngreso = Val(Me.txtIdMedicoIngreso.Tag)
           .IdMedicoEgreso = 0
           .IdServicioIngreso = Val(Me.txtIdServicioIngreso.Tag)
           .IdOrigenAtencion = Val(mo_cmbIdViasAdmision.BoundText)
           
           .IdDestinoAtencion = 0
           
           .HoraIngreso = IIf(Me.txtHoraIngreso.Text = sighEntidades.HORA_VACIA_HM, "", Me.txtHoraIngreso.Text)
           .FechaIngreso = IIf(Me.txtFechaIngreso.Text = sighEntidades.HORA_VACIA_HM, "", Me.txtFechaIngreso.Text)
           .fechaEgreso = 0
           .HoraEgreso = ""
           .idTipoServicio = mo_cmbIdTipoServicio.BoundText
           .Edad = Me.txtEdadEnDias.Text
           .idTipoEdad = Val(mo_cmbIdTipoEdad.BoundText)
           .idPaciente = Me.idPaciente
           .IdUsuarioAuditoria = Me.idUsuario
           'Estos datos llenaran  en el modulo de registro de atenciones
            If Me.chkPacienteNuevo = 1 Then
                .IdTipoCondicionALEstab = 1
                .IdTipoCondicionAlServicio = 1
            Else
                mo_AdminServiciosComunes.TiposCondicionPacienteCondicionAlEstablecimientoYservicio .IdTipoCondicionALEstab, .IdTipoCondicionAlServicio, Me.idPaciente, Format(Me.txtFechaIngreso, sighEntidades.DevuelveFechaSoloFormato_DMY), Me.idAtencion, Me.txtIdServicioIngreso.Tag
            End If
   
            .FechaEgresoAdministrativo = 0
            .HoraEgresoAdministrativo = ""
            .IdCamaIngreso = Val(Me.txtNroCamaIngreso.Tag)
            If Me.txtCamaTransf.Visible = False Then
               .IdCamaEgreso = Val(Me.txtNroCamaIngreso.Tag)
            Else
               .IdCamaEgreso = Val(Me.txtCamaTransf.Tag)
            End If
            .IdCondicionAlta = 0
            If mi_Opcion = sghModificar Then
               If lcUltimoCodigoDeServicioTransferido <> "" Then
                    Dim oDOServicioH As New doServicio
                    Set oDOServicioH = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(lcUltimoCodigoDeServicioTransferido)
                    .IdServicioEgreso = oDOServicioH.IdServicio
                    .IdCamaEgreso = 0
                    Set oDOServicioH = Nothing
               Else
                   '.IdServicioEgreso = Val(Me.txtIdServicioIngreso.Tag)
               End If
            ElseIf mi_Opcion = sghAgregar Then
               .IdServicioEgreso = Val(Me.txtIdServicioIngreso.Tag)
            End If
               
            .IdTipoAlta = 0
            
            
            .IdTipoGravedad = Val(mo_cmbIdTipoGravedad.BoundText)
            .IdFormaPago = Val(cmbFormaPago.BoundText)            'Tiposfinanciamiento-->Credito Hospitalario
            .IdFuenteFinanciamiento = Val(cmbFuenteFinanciamiento.BoundText) 'Fuentefinanciamiento-->Credito hospitalario
            .IdEstadoAtencion = 1
            
            If mi_Opcion = sghModificar And Me.ucTransferenciasDetalle1.getIdServicioUltimaTransferencia() = 0 Then
                .IdServicioEgreso = .IdServicioIngreso
            End If
   End With
   
   


    '---------------------------------------------------------------------------------
    '           CARGA DATOS DEL PACIENTE
    '---------------------------------------------------------------------------------
    Me.ucPacientesDetalle1.idUsuario = ml_idUsuario
    Me.ucPacientesDetalle1.CargarDatosAlObjetoDatos mo_Pacientes, mo_Historia, mo_DoPacientesDatosAdd

    '---------------------------------------------------------------------------------
    '           COMPLETA LOS DATOS DE LA ATENCION
    '---------------------------------------------------------------------------------
    With mo_DoAtencionDatosAdicionales
        .IdMedicoRespNacimiento = Val(Me.txtIdMedicoNacimiento.Tag)
        .IdTipoReferenciaOrigen = Val(mo_cmbIdTipoReferenciaOrigen.BoundText)
        If .IdTipoReferenciaOrigen = 1 Then
            .IdEstablecimientoOrigen = Val(Me.txtIdEstablecimientoOrigen.Tag)
            .IdEstablecimientoNoMinsaOrigen = 0
        Else
            .IdEstablecimientoOrigen = 0
            .IdEstablecimientoNoMinsaOrigen = Val(Me.txtIdEstablecimientoOrigen.Tag)
        End If
        '.IdTipoReferenciaDestino = Val(mo_cmbIdTipoReferenciaDestino.BoundText)
        If .IdTipoReferenciaDestino = 1 Then
             '.IdEstablecimientoDestino = Val(Me.txtIdEstablecimientoDestino.Tag)
             '.IdEstablecimientoNoMinsaDestino = 0
        Else
             '.IdEstablecimientoDestino = 0
             '.IdEstablecimientoNoMinsaDestino = Val(Me.txtIdEstablecimientoDestino.Tag)
        End If
        .RecienNacido = (chkRecienNacido.Value = 1)
        '.TieneNecropsia = IIf(Me.chkSeRealizoNecropsia.Value, True, False)
        .HuboInfeccionIntraHospitalaria = False
        .NroReferenciaOrigen = txtReferenciaO.Text
        '.NroReferenciaDestino = txtReferenciaD.Text
        
        .DireccionDomicilio = mo_Pacientes.DireccionDomicilio
        .NombreAcompaniante = Me.txtNombreAcompañante.Text
        .AcompanianteDNI = Me.txtDNIacompaniante.Text      'debb-21/06/2016
        .emergenciaCorrelativo = Mid(lcBuscaParametro.RetornaFechaHoraServidorSQL, 7, 4) & Me.txtEmergenciaN.Text    'debb-21/06/2016
        If mi_Opcion = sghAgregar Or mi_Opcion = sghModificar Then
           If Val(cmbFuenteFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS Then
                If lnAfiliacionSIS4 = 0 Or lcCodigoEstablecimientoAdscripcionSIS = "" Then
                   mo_ReglasSISgalenhos.SisFiliacionesDevuelveKEY lnAfiliacionSIS4, lcSIScodigo, _
                                        mo_Pacientes.ApellidoPaterno, mo_Pacientes.ApellidoMaterno, _
                                        mo_Pacientes.PrimerNombre, mo_Pacientes.FechaNacimiento, _
                                        lcCodigoEstablecimientoAdscripcionSIS
                End If
                .idSiaSis = lnAfiliacionSIS4
                If wxParametro302 = "S" Then
                   .FuaCodigoPrestacion = Me.ucSISfuaCodPrestacion1.CodigoPrestacion
                End If
                .SisCodigo = lcSIScodigo
                .sisAfiliacion = txtNroAfiliacionSis.Text
           Else
                .idSiaSis = 0
                .FuaCodigoPrestacion = ""
                .SisCodigo = ""
                .sisAfiliacion = ""
           End If
        End If
        'debb-21/06/2016 (inicio)
        .referenciaOservicio = PVcomboBoxDevuelveEleccion(cmbServicioReferenciaO)
        
        .idAtencionEmeg_CE = ml_idAtencionEmeg_CE
        '.referenciaDservicio
        '.referenciaDfextension
        '.referenciaDftramite
        'debb-21/06/2016 (fin)
        'FRANKLIN 2017
        .ReferenciaMedicoOColeg = Me.txtMedicoRef.Text
        If Trim(Me.cmbMedicoRef.Text) <> "" Then
           .ReferenciaMedicoOIdcolegio = Trim(Split(Me.cmbMedicoRef.Text, " = ")(0))
        Else
           .ReferenciaMedicoOIdcolegio = ""
        End If
        '
        
    End With


    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE DIAGNOSTICOS DE INGRESO
    '---------------------------------------------------------------------------------
    Me.ucDiagnosticosIngreso.idUsuario = ml_idUsuario
    ucDiagnosticosIngreso.TipoDiagnostico = sghHospitalizacionIngreso
    Me.ucDiagnosticosIngreso.CargarDiagnosticosAlObjetoDatos mo_Diagnosticos
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE DIAGNOSTICOS DE EGRESO
    '---------------------------------------------------------------------------------
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE DIAGNOSTICOS DE MORTALIDAD
    '---------------------------------------------------------------------------------
    
    'If ml_TipoServicio = sghHospitalizacion Then
        '---------------------------------------------------------------------------------
        '           CARGA DATOS DE NACIMIENTO
        '---------------------------------------------------------------------------------
        Me.ucNacimientoDetalle1.idUsuario = ml_idUsuario
        Me.ucNacimientoDetalle1.CargarNacimientosAlObjetoDatos mo_Nacimientos
        
        '---------------------------------------------------------------------------------
        '           CARGA DATOS DE DIAGNOSTICOS DE NACIMIENTO
        '---------------------------------------------------------------------------------
        Me.ucDiagnosticoNacimiento.idUsuario = ml_idUsuario
        Me.ucDiagnosticoNacimiento.TipoDiagnostico = sghHospitalizacionNacimiento
        Me.ucDiagnosticoNacimiento.CargarDiagnosticosAlObjetoDatos mo_Diagnosticos
    'End If
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE DIAGNOSTICOS DE COMPLICACIONES
    '---------------------------------------------------------------------------------
    
    
    Dim oDOOcupacion As New DOEstanciaHospitalaria
    oDOOcupacion.IdServicio = Val(Me.txtIdServicioIngreso.Tag)
    oDOOcupacion.IdMedicoOrdena = Val(Me.txtIdMedicoIngreso.Tag)
    oDOOcupacion.FechaOcupacion = Me.txtFechaIngreso
    oDOOcupacion.HoraOcupacion = Me.txtHoraIngreso
    oDOOcupacion.idCama = Val(Me.txtNroCamaIngreso.Tag)
    oDOOcupacion.IdUsuarioAuditoria = ml_idUsuario
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE TRANSFERENCIAS
    '---------------------------------------------------------------------------------
    Me.ucTransferenciasDetalle1.idUsuario = ml_idUsuario
    Me.ucTransferenciasDetalle1.CargaTransferenciasAlObjetosDatos mo_OcupacionCamas, oDOOcupacion, _
                                 Format(ml_ldFechaEgreso, sighEntidades.DevuelveFechaSoloFormato_DMY), ml_lcHoraEgreso, chkLlegoSI.Value, _
                                 chkLlegoSS.Value, lnSecuenciaTransferencia, Val(txtCamaTransf.Tag), 0
    
    
    If ml_TipoServicio = sghEmergenciaConsultorios Or ml_TipoServicio = sghEmergenciaObservacion Then
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE ATENCIONES DE EMERGENCIA
    '---------------------------------------------------------------------------------
        With mo_AtencionesEmergencia
            .idAtencion = ml_idAtencion
            .IdAtencionEmergencia = ml_IdAtencionEmergencia
            .IdCausaExternaMorbilidad = Val(mo_cmbIdCausaExternaMorbilidad.BoundText)
            .IdClaseAccidente = Val(mo_cmbIdClaseAccidente.BoundText)
            .IdGrupoOcupacionalALAB = Val(mo_cmbIdGrupoOcupacionalALAB.BoundText)
            .IdLugarEvento = Val(mo_cmbIdLugarEvento.BoundText)
            .IdPosicionLesionadoALAB = Val(mo_cmbIdPosicionLesionadoALAB.BoundText)
            .IdRelacionAgresorVictima = Val(mo_cmbIdRelacionAgresorVictima.BoundText)
            .IdSeguridad = Val(mo_cmbIdSeguridad.BoundText)
            .IdTipoAgenteAGAN = Val(mo_cmbIdTipoAgenteAGAN.BoundText)
            .IdTipoEvento = Val(mo_cmbIdTipoEvento.BoundText)
            .IdTipoTransporte = Val(mo_cmbIdTipoTransporte.BoundText)
            .IdTipoVehiculo = Val(mo_cmbIdTipoVehiculo.BoundText)
            .IdUbicacionLesionado = Val(mo_cmbIdUbicacionLesionado.BoundText)
            .IdUsuarioAuditoria = ml_idUsuario
            .comoLlego = Me.cmbComoLlego.ListIndex + 1
            .tipoAtencion = Me.cmbTipoAtencion.ListIndex + 1
            .idEstadoLlegada = IIf(Me.cmbEstadoLlegada.Text = "", 0, Me.cmbEstadoLlegada.ListIndex + 1)
        End With
    End If
    Select Case tabAdmision.Tab
    Case 1
       lcCaptionTab2 = TabIngreso.Caption
    End Select
    '
    Me.UcPacientesSunasa1.idUsuario = ml_idUsuario
    Me.UcPacientesSunasa1.CargarDatosAlObjetoDatos oDoSunasaPacientesHistoricos
End Sub
Sub GrabaImagenesEnRutaDelServidor()
    Dim lcArchivoElegido As String
    Dim lcArchivoImagenFinal As String
    lcArchivoElegido = ucPacientesDetalle1.ArchivoElegido
    lcArchivoImagenFinal = wxParametro237 & "\" & Trim(Str(mo_Pacientes.NroHistoriaClinica)) & ".JPG"
    If lcArchivoElegido = "DEL" Then
       Kill lcArchivoImagenFinal
    ElseIf lcArchivoElegido <> "" Then
        pi_imagen.Picture = LoadPicture(lcArchivoElegido)
        SavePicture pi_imagen, lcArchivoImagenFinal
    End If
End Sub
'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    Dim esPacienteNuevo As Boolean
    esPacienteNuevo = False
    If mo_Pacientes.idPaciente = 0 Then
        esPacienteNuevo = True
    End If
    AgregarDatos = mo_AdminAdmision.AdmisionHospAgregar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mo_Historia, _
                                                        Me.ucPacientesDetalle1.TipoNumeracionAnterior, mo_OcupacionCamas, _
                                                        mo_Diagnosticos, mo_Procedimientos, mo_Examenes, mo_Nacimientos, _
                                                        Me.ucPacientesDetalle1.IdHistoriaClinicaAnterior, _
                                                        mo_AtencionesEmergencia, mo_AtencionPadre, lbPacienteNN, _
                                                        mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                                        Trim(tabAdmision.Caption) & "/" & Trim(lcCaptionTab2), _
                                                        lnIdNacimientoSeleccionado, oDoSunasaPacientesHistoricos, _
                                                        mo_DoAtencionDatosAdicionales, mb_EsObservacionEmergencia, _
                                                        mo_DoPacientesDatosAdd)
    ms_MensajeError = mo_AdminAdmision.MensajeError
    If ms_MensajeError = "" Then
'        If Val(wxParametro208) <> 7686 Then
'            If esPacienteNuevo = True Then
'                Dim o_ReglasIntegracion As New ReglasIntegracion
'                Call o_ReglasIntegracion.EnviarDatosPacienteRisPacs(mo_Pacientes)
'            End If
'            GrabaImagenesEnRutaDelServidor
'        End If
        mo_AdminArchivoClinico.ActualizaIdRecienNacidoEnTablaAtenciones mo_Atenciones.idAtencion
        If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS And _
                                                                             lbEncontroAfiliadoEnWebSIS = True Then
           lcTipoFormatoSIS = IIf(Trim(lcTipoFormatoSIS) = "", Trim(Str(sghSIScodigo.sghAfiliacionAUXgratis)), lcTipoFormatoSIS)
           mo_ReglasSISgalenhos.SisFiliacionesActualizarAfiliadoDesdeWEB lcDniSIS, lnAfiliacionSIS1, lnAfiliacionSIS2, _
                                            lnAfiliacionSIS3, lnAfiliacionSIS5, lcTipoFormatoSIS, _
                                            wxParametro323
        End If
    End If
End Function



'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    If CalculaEstanciaParaPacienteConAltaMedica(mo_Atenciones) = True Then
        Dim oEpisodioClinico As EpisodioClinico
        oEpisodioClinico = EpisodioClinicoDevuelveDatos
        Dim esPacienteNuevo As Boolean
        esPacienteNuevo = False
        If mo_Pacientes.idPaciente = 0 Then
            esPacienteNuevo = True
        End If
        ModificarDatos = mo_AdminAdmision.AdmisionHospModificar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mo_Historia, _
                                                                Me.ucPacientesDetalle1.TipoNumeracionAnterior, mo_OcupacionCamas, _
                                                                mo_Diagnosticos, mo_Procedimientos, mo_Examenes, mo_Nacimientos, _
                                                                Me.ucPacientesDetalle1.IdHistoriaClinicaAnterior, mo_AtencionesEmergencia, _
                                                                ml_TipoAccionAdmision, ldFechaEgresoMedicoAnterior, lbPacienteNN, _
                                                                mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                                                Trim(tabAdmision.Caption) & "/" & Trim(lcCaptionTab2), lnIdPlanAnterior, _
                                                                lnIdTipoFinanciamientoAnterior, oRsEstancia, lnIdNacimientoSeleccionado, _
                                                                oDoSunasaPacientesHistoricos, mo_DoAtencionDatosAdicionales, _
                                                                lcUltimoCodigoDeServicioTransferido, mb_EsObservacionEmergencia, _
                                                                oEpisodioClinico, mo_DoPacientesDatosAdd)
        ms_MensajeError = mo_AdminAdmision.MensajeError
        If ms_MensajeError = "" Then
            If Val(wxParametro208) <> 7686 Then
                If esPacienteNuevo = True Then
                    Dim o_ReglasIntegracion As New ReglasIntegracion
                    Call o_ReglasIntegracion.EnviarDatosPacienteRisPacs(mo_Pacientes)
                End If
                GrabaImagenesEnRutaDelServidor
            End If
            '
            If Val(wxParametro208) <> 7686 Then
                Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
                mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar mo_Atenciones.idCuentaAtencion, False, 0
                Set mo_ReglasFacturacion = Nothing
            End If
            '
            If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
               mo_ReglasSISgalenhos.SisFuaAtencionActualizaDatosDesdeHospEmegCE mo_Atenciones.idCuentaAtencion, _
                                                                      mo_Atenciones.idTipoServicio, mo_Atenciones.idAtencion, _
                                                                      mo_lnIdTablaLISTBARITEMS, ml_idUsuario
            End If
        End If
    End If
    
End Function

Sub ActualizaCodigoPrestacionEnCE(oConexion As Connection)
    On Error GoTo errActCP
    Dim oDoAtencionDatosAdicionales9 As New DoAtencionDatosAdicionales
    Dim oAtencionesDatosAdicionales9 As New AtencionesDatosAdicionales
    Dim oAtenciones9 As New Atenciones
    Dim oDOAtencion9 As New DOAtencion
    If wxParametro302 = "S" And mo_Atenciones.IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS And _
                                            mo_Atenciones.IdOrigenAtencion = 30 And wxParametro553 = "S" Then
       Dim ldFechaAdmisionHosp As Date, lnIdAtencionEnCE As Long
       ldFechaAdmisionHosp = CDate(mo_Atenciones.FechaIngreso)
       lnIdAtencionEnCE = mo_DoAtencionDatosAdicionales.idAtencionEmeg_CE
       If lnIdAtencionEnCE > 0 Then
            Set oAtencionesDatosAdicionales9.Conexion = oConexion
            Set oAtenciones9.Conexion = oConexion
            oDOAtencion9.idAtencion = lnIdAtencionEnCE
            oDOAtencion9.IdUsuarioAuditoria = sighEntidades.Usuario
            If oAtenciones9.SeleccionarPorId(oDOAtencion9) = True Then
               If CDate(oDOAtencion9.FechaIngreso = ldFechaAdmisionHosp) Then
                    oDoAtencionDatosAdicionales9.idAtencion = lnIdAtencionEnCE
                    oDoAtencionDatosAdicionales9.IdUsuarioAuditoria = sighEntidades.Usuario
                    If oAtencionesDatosAdicionales9.SeleccionarPorId(oDoAtencionDatosAdicionales9) = True Then
                       oDoAtencionDatosAdicionales9.FuaCodigoPrestacion = "056"
                       If oAtencionesDatosAdicionales9.Modificar(oDoAtencionDatosAdicionales9) = True Then
                       End If
                    End If
               End If
            End If
       End If
    End If
errActCP:
    Set oDoAtencionDatosAdicionales9 = Nothing
    Set oAtencionesDatosAdicionales9 = Nothing
    Set oAtenciones9 = Nothing
    Set oDOAtencion9 = Nothing
End Sub



Function EpisodioClinicoDevuelveDatos() As EpisodioClinico
        Dim oEpisodioClinico As EpisodioClinico
        oEpisodioClinico.idEpisodio = Me.UcEpisodioClinico1.idEpisodio
        oEpisodioClinico.lbCierreEpisodio = Me.UcEpisodioClinico1.lbCierreEpisodio
        oEpisodioClinico.lbNuevoEpisodio = Me.UcEpisodioClinico1.lbNuevoEpisodio
        EpisodioClinicoDevuelveDatos = oEpisodioClinico
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos(oConexion As Connection) As Boolean
    ms_MensajeError = mo_AdminAdmision.VerificaSiTieneMovimientoFarmaciaOservicio(mo_CuentasAtencion.idCuentaAtencion, _
                                                   mo_Atenciones.idTipoServicio, oConexion)
    If ms_MensajeError = "" Then
        Dim oEpisodioClinico As EpisodioClinico
        oEpisodioClinico = EpisodioClinicoDevuelveDatos
        '
        mo_CuentasAtencion.idEstado = 9 'anulado
        mo_Atenciones.IdEstadoAtencion = 0  'anulado
        EliminarDatos = mo_AdminAdmision.AdmisionHospAnular(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mo_Historia, _
                                               Me.ucPacientesDetalle1.TipoNumeracionAnterior, mo_OcupacionCamas, _
                                               mo_Diagnosticos, mo_Procedimientos, mo_Examenes, mo_Nacimientos, _
                                               Me.ucPacientesDetalle1.IdHistoriaClinicaAnterior, mo_AtencionesEmergencia, _
                                               ml_TipoAccionAdmision, ldFechaEgresoMedicoAnterior, False, _
                                               mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                               Trim(tabAdmision.Caption) & "/" & Trim(lcCaptionTab2), lnIdNacimientoSeleccionado, _
                                               oEpisodioClinico)
        ms_MensajeError = mo_AdminAdmision.MensajeError
    Else
        MsgBox ms_MensajeError & Chr(13) & "La Anulación tendrá que realizarlo FACTURACION ", vbInformation, "Admisión"
    End If
End Function

'debb-Jamo
Sub CargaAtencionCEJamo()
       On Error GoTo ErrJamo
       If ml_TipoServicio = sghEmergenciaConsultorios Or ml_TipoServicio = sghHospitalizacion Then
            Dim oAtencionesCE As New AtencionesCE
            Dim oConexion As New Connection
            oConexion.Open wxParametroJAMO
            mo_DOAtencionesCE.idAtencion = ml_idAtencion
            Set oAtencionesCE.Conexion = oConexion
            
            'mgaray
            ucTriajeVisorCE.Origen = IIf(ml_TipoServicio = sghHospitalizacion, sightriajeorigen.Hospitalizacion, sightriajeorigen.Emergencia)
            ucTriajeVisorCE.EstadoPaciente = 0
            ucTriajeVisorCE.OpcionFormulario = mi_Opcion
            ucTriajeVisorCE.AsignarIdAtencionYLlenarControles (mo_Atenciones.idAtencion)
            
            If oAtencionesCE.SeleccionarPorId(mo_DOAtencionesCE) = False Then
                mo_DOAtencionesCE.idAtencion = 0
                Exit Sub
            End If
'            mo_Formulario.HabilitarDeshabilitar Me.txtPeso, True
'            mo_Formulario.HabilitarDeshabilitar Me.txtPresion, True
'            mo_Formulario.HabilitarDeshabilitar Me.txtTalla, True
'            mo_Formulario.HabilitarDeshabilitar Me.txtTemperatura, True
'            mo_Formulario.HabilitarDeshabilitar Me.txtFrespiratoria, True
'            mo_Formulario.HabilitarDeshabilitar Me.txtPulso, True
            If Not mo_DOAtencionesCE Is Nothing Then
                 With mo_DOAtencionesCE
'                      Me.txtPeso.Text = .TriajePeso
'                      Me.txtPresion.Text = .TriajePresion
'                      Me.txtTalla.Text = .TriajeTalla
'                      Me.txtTemperatura.Text = .TriajeTemperatura
'                      Me.txtFrespiratoria.Text = .TriajeFrecRespiratoria
'                      Me.txtPulso.Text = .TriajePulso
                       Me.TxtCitaTratamiento.Text = .CitaTratamiento
                 End With
            Else
                mo_DOAtencionesCE.idAtencion = 0
            End If
            Set oAtencionesCE = Nothing
            Set oConexion = Nothing
       End If
       Exit Sub
ErrJamo:
'       mo_Formulario.HabilitarDeshabilitar Me.txtPeso, False
'       mo_Formulario.HabilitarDeshabilitar Me.txtPresion, False
'       mo_Formulario.HabilitarDeshabilitar Me.txtTalla, False
'       mo_Formulario.HabilitarDeshabilitar Me.txtTemperatura, False
'       mo_Formulario.HabilitarDeshabilitar Me.txtFrespiratoria, False
'       mo_Formulario.HabilitarDeshabilitar Me.txtPulso, False
End Sub


Sub CargarDatosAlosControles()
        Dim oConexion As New Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighEntidades.CadenaConexion
        
        'CargaFormaPago
        If oRsFormaPago.State = adStateOpen Then oRsFormaPago.Close
        Set oRsFormaPago = mo_AdminServiciosComunes.TiposFinanciamientoSegunFiltro("esFuenteFinanciamiento=1")
        Set cmbFormaPago.RowSource = oRsFormaPago
        cmbFormaPago.ListField = "Descripcion"
        cmbFormaPago.BoundColumn = "idTipoFinanciamiento"
        mo_Formulario.HabilitarDeshabilitar Me.cmbFormaPago, False
        '
        '1do:   CARGAR DATOS DE LA ATENCION
        CargarDatosDelaAtencion oConexion
        
        If mo_Atenciones.idAtencion = 0 Then
             mb_ExistenDatos = False
             Exit Sub
        End If
        '
        If lbCargaAlaVezCitaPacienteAtencionDA = False Then
           Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(Me.idCuentaAtencion, oConexion)
        End If
        lblEstadoCta = mo_ReglasFarmacia.DevuelveEstadoActualDeEstadoCuenta("idEstado=" & mo_CuentasAtencion.idEstado, oConexion)
        If mo_CuentasAtencion.idEstado <> 1 And mo_CuentasAtencion.idEstado <> 12 Then
            btnAceptar.Enabled = False
            
            'Actualizado 16102014 A.Yañez *****************************
            cmbIdViasAdmision.Enabled = False
            lblNombreServicio.Enabled = False
            lblNombreMedico.Enabled = False
            txtNroCamaIngreso.Enabled = False
            cmbIdTipoGravedad.Enabled = False
            cmbFuenteFinanciamiento.Enabled = False
            cmbFormaPago.Enabled = False
            txtFechaIngreso.Enabled = False
            txtHoraIngreso.Enabled = False
            '**************************************
        End If
        'debb-03/08/2016
'        If mo_CuentasAtencion.IdEstado = 12 Then
            btnBuscarServicios.Enabled = True
            btnBuscarMedicos.Enabled = True
'        Else
'            btnBuscarServicios.Enabled = False
'            btnBuscarMedicos.Enabled = False
'        End If
        txtNroCuenta.Text = mo_CuentasAtencion.idCuentaAtencion
        
        '
        ms_MensajeError = VerSiTieneServicioAutomaticoPorEstancia(oConexion)
        
        '4to:   PARA VISUALIZAR LA UBICACION DEL PACIENTE AL DIA DE LA ATENCION
        mo_DoUbicacionPaciente.DireccionDomicilio = mo_DoAtencionDatosAdicionales.DireccionDomicilio
        Me.ucPacientesDetalle1.ReemplazarDatosDeUbicacion mo_DoUbicacionPaciente
        
        '4to:   CARGAR DATOS DE LOS DIAGNOSTICOS INGRESO POR ATENCION
        Me.ucDiagnosticosIngreso.idAtencion = Me.idAtencion
        Me.ucDiagnosticosIngreso.TipoDiagnostico = sghHospitalizacionIngreso
        Me.ucDiagnosticosIngreso.CargarDatosDeDiagnosticos oConexion
        
        
        '5to:   CARGAR DATOS DE LOS DIAGNOSTICOS EGRESO POR ATENCION
        
        '6to:   CARGAR DATOS DE LOS DIAGNOSTICOS MORTALIDAD POR ATENCION
        
        '7to:   CARGAR DATOS DE LOS DIAGNOSTICOS NACIMIENTO POR ATENCION
        Me.ucDiagnosticoNacimiento.idAtencion = Me.idAtencion
        Me.ucDiagnosticoNacimiento.SexoPaciente = mo_Pacientes.idTipoSexo
        Me.ucDiagnosticoNacimiento.CargarDatosDeDiagnosticos oConexion
        
        '8to:   CARGAR DATOS DE LOS DIAGNOSTICOS COMPLICACIONES POR ATENCION
        
        '11to:    CARGAR DATOS DE OCUPACION DE EGRESOS
        Dim rsOcupacion As New Recordset
        lnSecuenciaTransferencia = 0
        Me.ucTransferenciasDetalle1.FechaIngreso = CDate(Me.txtFechaIngreso.Text & " " & Me.txtHoraIngreso.Text)
        Me.ucTransferenciasDetalle1.idAtencion = Me.idAtencion
        Me.ucTransferenciasDetalle1.idCuentaAtencion = Me.txtNroCuenta.Text
        Me.ucTransferenciasDetalle1.CargarDatosDeTransferencias oConexion
        If Me.ucTransferenciasDetalle1.IdServicioUltimaTransferencia = 0 Then
            mo_NroServiciosQuePasoElPaciente = 1
            CompletarDatosDeEgreso mo_Atenciones.IdServicioEgreso, mo_Atenciones.IdCamaEgreso
            'debb2009-el registro de la admision de un paciente en hospitalizacion
            'debb2009-se realiza desde Emergencia, cuando llega al Servicio de Hospitalizacion
            'debb2009-deberan 'CONFIRMAR SI LLEGO AL SERVICIO, sino se cierra la Cuenta despues
            'debb2009-de N horas desde su Admision.
            'debb2009-No hay transferencias hasta que confirme que llego.
            chkLlegoSI.Value = 1
            If mo_CuentasAtencion.idEstado = 12 And ml_TipoServicio = sghHospitalizacion And mi_Opcion = sghModificar And lbUsuarioConfirmaLlegada = True Then
                MsgBox "Debe confirmar que el Paciente 'Llegó al Servicio de Ingreso'" & Chr(13) & "sino la Cuenta se ANULARA en  " & wxParametro233 & " Horas", vbInformation, Me.Caption
                lnFocusCuandoCargeFrm = 2
                'A.Yañez *****************************
                chkLlegoSI.Visible = True
                chkLlegoSI.Value = 0
                chkLlegoSI.Enabled = True
                Me.TabIngreso.TabVisible(1) = False   'no hay TRANSFERENCIAS
                btnAceptar.Enabled = True
            End If
            'debb2009-fin
        Else
            CompletarDatosDeEgreso Me.ucTransferenciasDetalle1.IdServicioUltimaTransferencia, Me.ucTransferenciasDetalle1.IdCamaUltimaTransferencia
            
            
            lnFocusCuandoCargeFrm = 3
            'debb2009-como hay TRANSFERENCIAS, se debe chequear si llegó al ultimo Servicio Transferido
            'debb2009-deberan 'CONFIRMAR SI LLEGO AL SERVICIO TRANSFERIDO, sino se cierra la Cuenta despues
            'debb2009-de N horas desde su Admision
            chkLlegoSI.Value = 1
            chkLlegoSS.Value = 1
            'If wxParametro232="S" then
                Set rsOcupacion = mo_AdminAdmision.EstanciaHospitalariaSeleccionarPorAtencion(idAtencion, 0, oConexion)
                mo_NroServiciosQuePasoElPaciente = rsOcupacion.RecordCount
                If rsOcupacion.RecordCount > 0 Then
                    rsOcupacion.MoveLast
                    lnSecuenciaTransferencia = rsOcupacion.Fields!Secuencia
                    txtCamaTransf.Tag = Me.ucTransferenciasDetalle1.IdCamaUltimaTransferencia
                    txtCamaTransf.Text = ml_lcCamaEgreso
                    
                    'Frank 06052015
                    Dim oDoServicio As New doServicio
                    Set oDoServicio = getDatosDeServicio(Me.ucTransferenciasDetalle1.IdServicioUltimaTransferencia)
                    txtServicioTransf.Text = oDoServicio.Codigo & " - " & oDoServicio.nombre
                    Set oDoServicio = Nothing
                    
                    If (ml_TipoServicio = sghHospitalizacion Or mb_EsObservacionEmergencia = True) And mi_Opcion = sghModificar And mo_CuentasAtencion.idEstado = sghEstadoCuenta.sghNoLlegaAlServicioHospitalizado And rsOcupacion.Fields!LlegoAlServicio <> 1 And lbUsuarioConfirmaTransferencia = True Then
                        MsgBox "Debe confirmar que el Paciente 'Llegó al Servicio Transferido'" & Chr(13) & "sino la Cuenta se ANULARA en  " & wxParametro233 & " Horas", vbInformation, Me.Caption
                        lnFocusCuandoCargeFrm = 1
                        chkLlegoSS.Visible = True
'                        lblCamaTransf.Visible = True
'                        txtCamaTransf.Visible = True
'                        cmdCamaTransf.Visible = True
                        chkLlegoSS.Value = 0
                    End If
                    
                    If (ml_TipoServicio = sghHospitalizacion Or mb_EsObservacionEmergencia = True) Then
                        lblCamaTransf.Visible = True
                        txtCamaTransf.Visible = True
                        lblServicioTransf.Visible = True
                        txtServicioTransf.Visible = True
                        cmdCamaTransf.Visible = True
                        fraServicioActual.Visible = True
                        fraServicioActual.Top = 360
                        ucTransferenciasDetalle1.Top = fraServicioActual.Top + 650
                        lnFocusCuandoCargeFrm = 1
                    End If
                    rsOcupacion.MoveFirst
                End If
                'FRANK
                Me.ucTransferenciasDetalle1.ColocarsePrimerRegistroTransferencia
            'End If
            'debb2009-fin
        End If
        Set rsOcupacion = Nothing
        
        'If ml_TipoServicio = sghHospitalizacion Then
            '12to:    CARGAR DATOS DE OCUPACION DE CAMAS
            Me.ucNacimientoDetalle1.idAtencion = Me.idAtencion
            Me.ucNacimientoDetalle1.CargarDatosDeNacimientos oConexion
            If Left(Me.ucPacientesDetalle1.DevuelveSexo, 1) <> "2" Then
                  Me.TabIngreso.TabVisible(4) = False
            End If
        'End If
        
        '13to:    CARGAR DATOS DE ATENCION DE EMREGENCIA
        If ml_TipoServicio = sghEmergenciaConsultorios Or ml_TipoServicio = sghEmergenciaObservacion Then
            CargarDatosDeLaAtencionDeEmergencia oConexion
        End If
        
        '14avo:    CARGAR FECHA DE EGRESO MEDICO, si lo tuviese, para "MODIFICAR"--> con el fin de generar CONSUMO POR DIAS DE ESTANCIA
        ldFechaEgresoMedicoAnterior = ml_ldFechaEgreso
        
        '
        'Ya tuvo movimientos(Farmacia/servicios), no podrá cambiar de plan
        If mi_Opcion = sghModificar Then
            ms_MensajeError = mo_AdminAdmision.VerificaSiTieneMovimientoFarmaciaOservicio(mo_Atenciones.idCuentaAtencion, mo_Atenciones.idTipoServicio, oConexion)
            If ms_MensajeError <> "" Then
               mo_Formulario.HabilitarDeshabilitar Me.cmbFuenteFinanciamiento, False
               mo_Formulario.HabilitarDeshabilitar cmbFormaPago, False
               Me.ucMensajeParpadeando1.MensajeDeTexto = ms_MensajeError
               Me.ucMensajeParpadeando1.Visible = True
            End If
        End If
        '
        DeudasPendientesDeAnterioresAtenciones oConexion
        '
        CargaDatosDeLaMadre oConexion
        '
        Me.UcPacientesSunasa1.YaNoTieneSeguro
        Me.UcPacientesSunasa1.HabilitaFrame False
        If mo_Atenciones.idSunasaPacienteHistorico > 0 Then
            If mo_AdminFacturacion.TiposFinanciamientoGeneraReciboPago(Val(cmbFormaPago.BoundText), oConexion) = False Then
                Me.UcPacientesSunasa1.HabilitaFrame True
                Me.UcPacientesSunasa1.idSunasaPacienteHistorico = mo_Atenciones.idSunasaPacienteHistorico
                Me.UcPacientesSunasa1.CargarDatosPorId
            End If
        End If
        UcPacientesSunasa1.idTipoFinanciamiento = Val(cmbFormaPago.BoundText)
        '
        CargaAtencionCEJamo
        '
        If mi_Opcion <> sghAgregar Then
           Me.UcEpisodioClinico1.idPaciente = mo_Atenciones.idPaciente
           Me.UcEpisodioClinico1.idAtencion = mo_Atenciones.idAtencion
           Me.UcEpisodioClinico1.Inicializar
           Me.UcEpisodioClinico1.Limpiar
           Me.UcEpisodioClinico1.CargaEpisodiosHistoricos
           If ml_ldFechaEgreso <> 0 Then
              Me.UcEpisodioClinico1.CargarDatosAlosControles oConexion
           End If
        End If
        '
        oConexion.Close
        Set oConexion = Nothing
        '
        'If wxParametro302 = "S" And Val(cmbFuenteFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS Then
           HaceVisibleOnoBotonFUA
        'Else
        '   wxParametro302 = "N"
        '   Me.ucSISfuaCodPrestacion1.Visible = False
        '   Me.UcSISafiliacion1.Visible = False
        '   Me.chkBuscarEnSIS.Visible = False
        'End If
        '
        ReglasDeConsistenciasDespuesDeElegirCodigoPrestacion
        If mo_Atenciones.IdOrigenAtencion > 0 Then
          mo_cmbIdViasAdmision.BoundText = mo_Atenciones.IdOrigenAtencion
        End If
        '
        '
        mo_Formulario.HabilitarDeshabilitar Me.lblNombreServicio, True
        If mi_Opcion = sghModificar And Me.ucTransferenciasDetalle1.getIdServicioUltimaTransferencia() = 0 Then
             mo_Formulario.HabilitarDeshabilitar Me.lblNombreServicio, False
        End If
        
        Me.btnImprimeFiliacion.Enabled = True
        'mgaray
        'Actualizado 14/10/2014
        'bloqueoControlesImpresionFicha
        '
        CargaCPTrealizadosEnElServicio
        CargaApoyoDx
        
        ucPacientesCtasPDF1.Inicializar mo_Atenciones.idPaciente, mo_Atenciones.idCuentaAtencion
        'debb-14/03/2015 (inicio)
        If lb_puedeCambiarFuenteFinanciamiento = False Then
           mo_Formulario.HabilitarDeshabilitar cmbFuenteFinanciamiento, False
           mo_Formulario.HabilitarDeshabilitar cmbFormaPago, False
        End If
        'debb-14/03/2015 (fin)
        If wxParametro552 = "S" Then
        ElseIf wxParametro525 = "S" Then
           If Me.ucDiagnosticosIngreso.rsDiagnosticos.RecordCount > 0 Then
              mo_Formulario.HabilitarDeshabilitar lblNombreMedico, False
              Me.btnBuscarMedicos.Enabled = False
              Me.btnBuscarServicios.Enabled = False
           End If
        End If
        lbHuboCambioEnDato = False
End Sub

Sub CargaDatosDeLaMadre(oConexion As Connection)
    Dim oRsTmp As New Recordset
    Set oRsTmp = mo_AdminAdmision.AtencionesNacimientosSeleccionarPorFiltro("idPacienteNacido=" & Trim(Str(mo_Atenciones.idPaciente)), oConexion)
    If oRsTmp.RecordCount > 0 Then
       lnIdNacimientoSeleccionado = oRsTmp.Fields!idNacimiento
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    '
    If lnIdNacimientoSeleccionado > 0 Then
       lblMadre.Text = mo_AdminAdmision.DevuelveDatosDeLaMadreDelPacienteActual(lnIdNacimientoSeleccionado, Me.ucPacientesDetalle1.idTipoSexo, oConexion)
    End If
End Sub

Sub CargarDatosDeLasAtencionesPadres()
    
    '---------------------------------------------------------------------------------
    'CARGAR DATOS DE LA ATENCION PADRE SI ES QUE HUBIERA
    '---------------------------------------------------------------------------------
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Select Case ml_TipoAccionAdmision

    Case sghIngresarUnAlojamientoConjunto
    
        Set mo_AtencionPadre = mo_AdminAdmision.AtencionesSeleccionarPorId(ml_IdAtencionPadre, oConexion)
        Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(mo_AtencionPadre.idCuentaAtencion, oConexion)
        
    Case sghEnviarAObservacion
        
        Set mo_AtencionPadre = mo_AdminAdmision.AtencionesSeleccionarPorId(ml_IdAtencionPadre, oConexion)
        mo_AtencionPadre.fechaEgreso = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
        mo_AtencionPadre.HoraEgreso = Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
        mo_AtencionPadre.IdTipoAlta = 5         'Paciente trasladado a observacion de emergencia
        mo_AtencionPadre.IdCondicionAlta = 3    'Paciente inalterado
        'Se supone que a observacion siempre llegan desde el consultorio
        mo_AtencionPadre.IdDestinoAtencion = 22
        
        Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(mo_AtencionPadre.idCuentaAtencion, oConexion)
        
    Case sghTrasladarAHospitalizacion
        
        Set mo_AtencionPadre = mo_AdminAdmision.AtencionesSeleccionarPorId(ml_IdAtencionPadre, oConexion)
        mo_AtencionPadre.fechaEgreso = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY)
        mo_AtencionPadre.HoraEgreso = Format(Now, sighEntidades.DevuelveHoraSoloFormato_HM)
        mo_AtencionPadre.IdTipoAlta = 6         'Paciente trasladado a hospitalizaciòn
        mo_AtencionPadre.IdCondicionAlta = 3    'Paciente inalterado
        
        'Si viene a hospitalizarse desde de observacion de emergencia
        If mo_AtencionPadre.idTipoServicio = sghEmergenciaConsultorios Then
            mo_AtencionPadre.IdDestinoAtencion = 21
        End If
        
        'Si viene a hospitalizarse desde consultorio de emergencia
        If mo_AtencionPadre.idTipoServicio = sghEmergenciaObservacion Then
            mo_AtencionPadre.IdDestinoAtencion = 41
        End If
        
        Set mo_CuentasAtencion = mo_AdminFacturacion.CuentasAtencionSeleccionarPorId(mo_AtencionPadre.idCuentaAtencion, oConexion)
        
    End Select
    oConexion.Close
    Set oConexion = Nothing
End Sub

Sub CompletarDatosDeEgreso(lIdServicioEgreso As Long, lIdCamaEgreso As Long)
Dim oDOCama As New DOCama
Dim oDoServicio As New doServicio
Dim oConexion As New Connection
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorId(lIdCamaEgreso, oConexion)
        ml_lcCamaEgreso = oDOCama.Codigo
        ml_idCamaEgreso = oDOCama.idCama
        mb_EsObservacionEmergencia = False
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(lIdServicioEgreso, oConexion)
        If Not oDoServicio Is Nothing Then
            ml_idServicioEgreso = oDoServicio.IdServicio
            ml_lcCodigoServicioEgreso = oDoServicio.Codigo
            ml_lcServicioEgreso = oDoServicio.nombre
            mb_EsObservacionEmergencia = oDoServicio.EsObservacionEmergencia
        Else
            ml_idServicioEgreso = 0
            ml_lcCodigoServicioEgreso = ""
            ml_lcServicioEgreso = ""
        End If
        
        oConexion.Close
        Set oConexion = Nothing
        Set oDOCama = Nothing
        Set oDoServicio = Nothing
End Sub
Sub CargarDatosDeLaAtencionDeEmergencia(oConexion As Connection)

        Me.IdAtencionEmergencia = mo_AdminAdmision.AtencionesEmergenciaSeleccionarIdPorIdAtencion(ml_idAtencion, oConexion)
        If Me.IdAtencionEmergencia = 0 Then
             Exit Sub
        End If
        
        Set mo_AtencionesEmergencia = mo_AdminAdmision.AtencionesEmergenciaSeleccionarPorId(Me.IdAtencionEmergencia, oConexion)
        If mo_AdminAdmision.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption"
             Exit Sub
        End If
        
        If Not mo_AtencionesEmergencia Is Nothing Then
            With mo_AtencionesEmergencia
                Me.IdAtencionEmergencia = .IdAtencionEmergencia
                mo_cmbIdCausaExternaMorbilidad.BoundText = .IdCausaExternaMorbilidad
                mo_cmbIdClaseAccidente.BoundText = .IdClaseAccidente
                mo_cmbIdGrupoOcupacionalALAB.BoundText = .IdGrupoOcupacionalALAB
                mo_cmbIdLugarEvento.BoundText = .IdLugarEvento
                mo_cmbIdPosicionLesionadoALAB.BoundText = .IdPosicionLesionadoALAB
                mo_cmbIdRelacionAgresorVictima.BoundText = .IdRelacionAgresorVictima
                mo_cmbIdSeguridad.BoundText = .IdSeguridad
                mo_cmbIdTipoAgenteAGAN.BoundText = .IdTipoAgenteAGAN
                mo_cmbIdTipoEvento.BoundText = .IdTipoEvento
                mo_cmbIdTipoTransporte.BoundText = .IdTipoTransporte
                mo_cmbIdTipoVehiculo.BoundText = .IdTipoVehiculo
                mo_cmbIdUbicacionLesionado.BoundText = .IdUbicacionLesionado
                If .comoLlego > 0 Then
                   cmbComoLlego.ListIndex = .comoLlego - 1
                End If
                If .tipoAtencion > 0 Then
                   Me.cmbTipoAtencion.ListIndex = .tipoAtencion - 1
                End If
                If .idEstadoLlegada > 0 Then
                   Me.cmbEstadoLlegada.ListIndex = .idEstadoLlegada - 1
                End If
            End With
        End If
                

End Sub
Sub CargarDatosDelaAtencion(oConexion As Connection)
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oDoServicio As New doServicio
Dim lcEstadoAtencion As String
        If lbCargaAlaVezCitaPacienteAtencionDA = False Then
           Set mo_Atenciones = mo_AdminAdmision.AtencionesSeleccionarPorId(Me.idAtencion, oConexion)
           If mo_Atenciones.idAtencion = 0 Then
                'El registro ha sido eliminado, pero no se hizo el refresh
                 mb_ExistenDatos = False
                 Exit Sub
           End If
        Else
           mo_Atenciones.idAtencion = Me.idAtencion
           mb_ExistenDatos = mo_AdminAdmision.AtencionesPacientesCitasDatosadicionalesSeleccionarPorId(mo_Pacientes, _
                                                mo_Atenciones, mo_DoAtencionDatosAdicionales, _
                                                oConexion, mo_CuentasAtencion, False)
           If mo_Atenciones.idAtencion = 0 Then
                'El registro ha sido eliminado, pero no se hizo el refresh
                 mb_ExistenDatos = False
                 Exit Sub
           End If
        End If
        If mo_AdminAdmision.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbInformation, Me.Caption"
             mb_ExistenDatos = False
             Exit Sub
        End If

        If Not mo_Atenciones Is Nothing Then
           '
           Me.ucNacimientoDetalle1.Visible = True
           Frame5.Visible = True
           ucDiagnosticoNacimiento.Visible = True
           If mo_AdminServiciosHosp.ServiciosPuedeUsarFichaNacimientos(IIf(mo_Atenciones.IdServicioEgreso = 0, mo_Atenciones.IdServicioIngreso, mo_Atenciones.IdServicioEgreso), oConexion) = False Then
               Me.ucNacimientoDetalle1.Visible = False
               Frame5.Visible = False
               ucDiagnosticoNacimiento.Visible = False
           End If
           '
           With mo_Atenciones
                Me.idAtencion = .idAtencion
                Me.idPaciente = .idPaciente
                
                mo_cmbIdTipoServicio.BoundText = .idTipoServicio
                
                Me.txtIdServicioIngreso.Tag = .IdServicioIngreso
                Me.txtIdMedicoIngreso.Tag = .IdMedicoIngreso
                
                Me.IdEspecialidad = .IdEspecialidadMedico
'                Me.chkRecienNacido.Value = IIf(.RecienNacido, 1, 0)
                
                mo_cmbIdViasAdmision.BoundText = .IdOrigenAtencion
'                mo_cmbIdTipoReferenciaOrigen.BoundText = .IdTipoReferenciaOrigen
'                CompletarDatosDelEstablecimientoEnElLoad .idEstablecimientoOrigen, .IdEstablecimientoNoMinsaOrigen, txtIdEstablecimientoOrigen, lblNombreOrigenReferencia, .IdTipoReferenciaOrigen
'                txtReferenciaO.Text = IIf(IsNull(.nroReferenciaOrigen), "", .nroReferenciaOrigen)
                
'                mo_cmbIdTipoReferenciaDestino.BoundText = .IdTipoReferenciaDestino
'                CompletarDatosDelEstablecimientoEnElLoad .idEstablecimientoDestino, .IdEstablecimientoNoMinsaDestino, txtIdEstablecimientoDestino, lblNombreDestinoReferencia, .IdTipoReferenciaDestino
'                txtReferenciaD.Text = IIf(IsNull(.nroReferenciaDestino), "", .nroReferenciaDestino)
                
                Me.txtHoraIngreso.Text = IIf(.HoraIngreso = "", sighEntidades.HORA_VACIA_HM, .HoraIngreso)
                Me.txtFechaIngreso.Text = IIf(.FechaIngreso = 0, sighEntidades.FECHA_VACIA_DMY, .FechaIngreso)
                
                
                'Se guarda en estas variables para validar si el paciente ya esta de alta o no
                ml_lcHoraEgreso = IIf(.HoraEgreso = "", sighEntidades.HORA_VACIA_HM, .HoraEgreso)
                ml_ldFechaEgreso = .fechaEgreso
                
                Me.txtEdadEnDias.Text = .Edad
                Me.txtEdadEnDias.Tag = .Edad
                
                mo_cmbIdTipoEdad.BoundText = .idTipoEdad
                cmbIdTipoEdad.Tag = .idTipoEdad
                
                If mo_AdminProgramacion.MedicosSeleccionarPorId(.IdMedicoIngreso, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
                    Me.txtIdMedicoIngreso = oDOEmpleado.CodigoPlanilla
                    Me.lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                Else
                    Me.lblNombreMedico = ""
                End If
                
                
                
                
                Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(.IdServicioIngreso, oConexion)
                If Not oDoServicio Is Nothing Then
                    'mgaray20140926
'                    lcElServicioUsaGalenHos = IIf(oDoServicio.UsaGalenHos = True, "S", "N")

                    Me.txtIdServicioIngreso.Tag = oDoServicio.IdServicio
                    Me.txtIdServicioIngreso.Text = oDoServicio.Codigo
                    Me.lblNombreServicio = oDoServicio.nombre
                    Me.lblNombreServicio.Tag = oDoServicio.IdEspecialidad
                    mo_cmbIdTipoServicio.BoundText = oDoServicio.idTipoServicio
                    Me.IdEspecialidad = oDoServicio.IdEspecialidad
                    If oDoServicio.EsObservacionEmergencia = True And ml_TipoServicio = sghEmergenciaConsultorios Then '09/08/2011
                       Me.lblNroCamaIngreso.Visible = True
                       Me.txtNroCamaIngreso.Visible = True
                       Me.btnVerDisponibilidadDeCamas.Visible = True 'A.Yañez Actualizado 16102014
                    End If
                    'mgaray20140926
'                    If oDoServicio.UsaFUA = False Then
'                       lbElServicioRegistraFUA = "N"
'                       wxParametro302 = "N"
'                    Else
'                       lbElServicioRegistraFUA = "S"
'                       wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
'                    End If
                Else
                    Me.txtIdServicioIngreso.Tag = ""
                    Me.lblNombreServicio = ""
                    mo_cmbIdTipoServicio.BoundText = ""
                    Me.lblNombreServicio.Tag = ""
                End If
                'mgaray20140926
                Call SetVariableServicioUsaFUA(oDoServicio)
                
                
                'Cama de ingreso
                Me.txtNroCamaIngreso.Tag = .IdCamaIngreso
                Dim oDOCama As New DOCama
                Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorId(.IdCamaIngreso, oConexion)
                Me.txtNroCamaIngreso = oDOCama.Codigo
                                         
                mo_cmbIdTipoGravedad.BoundText = .IdTipoGravedad
                
                
                mo_cmbIdCondicionEnElServicio.BoundText = .IdTipoCondicionAlServicio
                mo_cmbIdCondicionEnElEstablecimiento.BoundText = .IdTipoCondicionALEstab
                cmbFormaPago.BoundText = .IdFormaPago
                cmbFuenteFinanciamiento.BoundText = .IdFuenteFinanciamiento
                
                Me.ucNacimientoDetalle1.FechaIngreso = IIf(.FechaIngreso = 0, sighEntidades.FECHA_VACIA_DMY, .FechaIngreso)
                
                Select Case .IdEstadoAtencion
                Case 0
                    lcEstadoAtencion = "Anulado"
                    btnAceptar.Enabled = False
                Case 1
                    lcEstadoAtencion = "Registrado"
                Case 2
                    lcEstadoAtencion = "Cerrado"
                    btnAceptar.Enabled = False
                End Select
                lnIdPlanAnterior = .IdFuenteFinanciamiento
                lnIdTipoFinanciamientoAnterior = .IdFormaPago
                Set oDOCama = Nothing
                ml_ldFechaEgreso = .fechaEgreso
                ml_idServicioEgreso = .IdServicioEgreso
                ml_lcServicioEgreso = mo_AdminFacturacion.BuscaServicioActualDelPaciente(ml_idServicioEgreso)
                ml_idCamaEgreso = .IdCamaEgreso
                ml_lcHoraEgreso = .HoraEgreso
                ml_idMedicoEgreso = .IdMedicoEgreso
                ml_lcMedicoEgreso = mo_AdminProgramacion.MedicosDevuelveNombre(.IdMedicoIngreso, oConexion)
                mb_ExistenDatos = True
           End With
           Me.ucNacimientoDetalle1.FechaIngreso = CDate(Format(txtFechaIngreso.Text, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & txtHoraIngreso.Text)
           
           '
           If lbCargaAlaVezCitaPacienteAtencionDA = False Then
              Set mo_DoAtencionDatosAdicionales = mo_AdminAdmision.AtencionesDatosAdicionalesSeleccionarPorId(Me.idAtencion, oConexion)
           End If
           With mo_DoAtencionDatosAdicionales
                Me.txtIdMedicoNacimiento.Tag = .IdMedicoRespNacimiento
                Me.chkRecienNacido.Value = IIf(.RecienNacido, 1, 0)
                mo_cmbIdTipoReferenciaOrigen.BoundText = .IdTipoReferenciaOrigen
                CompletarDatosDelEstablecimientoEnElLoad .IdEstablecimientoOrigen, .IdEstablecimientoNoMinsaOrigen, txtIdEstablecimientoOrigen, lblNombreOrigenReferencia, .IdTipoReferenciaOrigen
                txtReferenciaO.Text = IIf(IsNull(.NroReferenciaOrigen), "", .NroReferenciaOrigen)
                
                Me.txtNombreAcompañante.Text = .NombreAcompaniante      'debb-21/06/2016
                Me.txtDNIacompaniante.Text = .AcompanianteDNI           'debb-21/06/2016
                Me.txtEmergenciaN.Text = Mid(.emergenciaCorrelativo, 5, 10)       'debb-21/06/2016
                lnAfiliacionSIS4 = .idSiaSis
                lcSIScodigo = .SisCodigo
                If mo_AdminProgramacion.MedicosSeleccionarPorId(mo_DoAtencionDatosAdicionales.IdMedicoRespNacimiento, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
                    Me.txtIdMedicoNacimiento = oDOEmpleado.CodigoPlanilla
                    Me.lblNombreMedicoNacimiento = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
                Else
                    Me.lblNombreMedicoNacimiento = ""
                End If
                'debb-21/06/2016 (inicio)
                PVcomboBoxUbicaPosicion mo_DoAtencionDatosAdicionales.referenciaOservicio, cmbServicioReferenciaO
                txtNroAfiliacionSis.Text = mo_DoAtencionDatosAdicionales.sisAfiliacion
                ml_idAtencionEmeg_CE = mo_DoAtencionDatosAdicionales.idAtencionEmeg_CE
                'debb-21/06/2016 (fin)
                'franklin 2017
                Me.txtMedicoRef.Text = .ReferenciaMedicoOColeg
                BuscaMedicoRerencia .ReferenciaMedicoOIdcolegio
                
           End With
           'ESTOS DATOS SE UTILIZARAN MAS ADELANTE PARA ACTUALIZAR LA UBICACION DE PACIENTE
           Dim oPacientesTmp As New SIGHComun.doPaciente
           If lbCargaAlaVezCitaPacienteAtencionDA = False Then
              Set oPacientesTmp = mo_AdminAdmision.PacientesSeleccionarPorId(ml_IdPaciente, oConexion)
           Else
              Set oPacientesTmp = mo_Pacientes
           End If
           If mo_AdminAdmision.MensajeError <> "" Then
                 MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbInformation, "Datos de paciente"
                 Exit Sub
           End If
           If Not oPacientesTmp Is Nothing Then
                'lcHistoriaYpaciente = "(" & Trim(Str(oPacientesTmp.NroHistoriaClinica)) & ") " & _
                'Trim(oPacientesTmp.ApellidoPaterno) & " " & Trim(oPacientesTmp.ApellidoMaterno) & " " & _
                'Trim(oPacientesTmp.PrimerNombre)
                CargaHCyPaciente oPacientesTmp.NroHistoriaClinica, oPacientesTmp.ApellidoPaterno, oPacientesTmp.ApellidoMaterno, _
                                 oPacientesTmp.PrimerNombre
                Me.Caption = Trim(Me.Caption) & "  (HC: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oPacientesTmp.NroHistoriaClinica)), False) & " " & _
                             Trim(oPacientesTmp.ApellidoPaterno) & " " & Trim(oPacientesTmp.ApellidoMaterno) & _
                            " " & Trim(oPacientesTmp.PrimerNombre) & ") (Estado Atenc: " & lcEstadoAtencion & _
                            ")(Edad: " & Trim(txtEdadEnDias.Text) & " " & Left(Trim(cmbIdTipoEdad.Text), 1) & _
                            ")(Gs: " & IIf(IsNull(oPacientesTmp.GrupoSanguineo), "", oPacientesTmp.GrupoSanguineo) & _
                            ", Frh: " & IIf(IsNull(oPacientesTmp.FactorRh), "", oPacientesTmp.FactorRh) & ")"
                With oPacientesTmp
                    mo_DoUbicacionPaciente.IdPaisDomicilio = .IdPaisDomicilio
                    mo_DoUbicacionPaciente.IdCentroPobladoDomicilio = .IdCentroPobladoDomicilio
                    
                    mo_DoUbicacionPaciente.IdPaisProcedencia = .IdPaisProcedencia
                    mo_DoUbicacionPaciente.IdCentroPobladoProcedencia = .IdCentroPobladoProcedencia
                    
                    mo_DoUbicacionPaciente.DireccionDomicilio = .DireccionDomicilio
                End With
                Me.ucPacientesDetalle1.CargarDatosDePacienteALosControlesSinBuscar oPacientesTmp, wxParametro242, wxParametro287
                lbPacienteNN = Me.ucPacientesDetalle1.DevuelveSiElPacienteEsNN
           End If
           Set oPacientesTmp = Nothing
           '
           cmbFormaPago.BoundText = mo_Atenciones.IdFormaPago
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
       
        Set oDoMedico = Nothing
        Set oDOEmpleado = Nothing
        Set oDOEspecialidades = Nothing
        Set oDoServicio = Nothing

End Sub
'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla CuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           'LIMPIAR DATOS DE LA CUENTA DE ATENCION
           Me.idCuentaAtencion = 0
           Me.idAtencion = 0
           
           'LIMPIAR DATOS DE LA ATENCION
           mo_cmbIdTipoReferenciaOrigen.BoundText = ""
           Me.txtIdEstablecimientoOrigen.Text = ""
           Me.txtEdadEnDias.Text = ""
           
           Me.ucPacientesDetalle1.LimpiarDatosDePaciente wxParametro211, ldFechaActualServidor
            
           Me.idAtencion = 0
           Me.IdAtencionEmergencia = 0
           Me.cmbIdTipoAgenteAGAN.Text = ""
           Me.cmbIdGrupoOcupacionalALAB.Text = ""
           Me.cmbIdPosicionLesionadoALAB.Text = ""
           Me.cmbIdUbicacionLesionado.Text = ""
           Me.cmbIdTipoTransporte.Text = ""
           Me.cmbIdTipoVehiculo.Text = ""
           Me.cmbIdClaseAccidente.Text = ""
           Me.cmbIdRelacionAgresorVictima.Text = ""
           Me.cmbIdSeguridad.Text = ""
           Me.cmbIdTipoEvento.Text = ""
           Me.cmbIdLugarEvento.Text = ""
           Me.cmbIdCausaExternaMorbilidad.Text = ""
           
           btnImprimeFiliacion.Enabled = False
           
           'A.Yañez 06-11-2014 ********************************
           Me.cmbIdViasAdmision.Text = ""
           Me.lblNombreServicio.Text = ""
           Me.lblNombreMedico.Text = ""
           Me.txtIdServicioIngreso.Text = ""
           Me.txtIdMedicoIngreso.Text = ""
           Me.txtEdadEnDias.Text = ""
           Me.cmbIdTipoEdad.Text = ""
           Me.txtNroCamaIngreso = ""
           Me.lblMadre.Text = ""
           Me.cmbFuenteFinanciamiento.Text = ""
           Me.cmbFormaPago.Text = ""
           Me.cmbIdTipoReferenciaOrigen.Text = ""
           Me.txtIdEstablecimientoOrigen.Text = ""
           Me.lblNombreOrigenReferencia.Text = ""
           Me.txtReferenciaO.Text = ""
           Me.ucDiagnosticosIngreso.limpiacampos
           lbProcedeDeEmergencia = False: lbProcedeDeConsExt = False    'debb-23/02/2015
           '******************************************
           cmbComoLlego.Text = ""
           cmbTipoAtencion.Text = ""
           Me.cmbEstadoLlegada.Text = ""
End Sub
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'*****************************************************************************************
'                               EVENTOS DE LA ATENCION
'*****************************************************************************************
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------



Private Sub cmbIdTipoReferenciaOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoReferenciaOrigen
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoReferenciaOrigen_LostFocus()
   If cmbIdTipoReferenciaOrigen.Text <> "" Then
       mo_cmbIdTipoReferenciaOrigen.BoundText = Val(Split(cmbIdTipoReferenciaOrigen.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoReferenciaOrigen
End Sub

Private Sub cmbIdTipoReferenciaOrigen_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIdEstablecimientoOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdEstablecimientoOrigen
    If KeyCode = vbKeyF1 Then
        btnBuscarEstablecimiento_Click
    End If
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdEstablecimientoOrigen_LostFocus()
    If mo_cmbIdTipoReferenciaOrigen.BoundText <> "" Then
        CompletarDatosDelEstablecimientoEnElLostFocus txtIdEstablecimientoOrigen, lblNombreOrigenReferencia, mo_cmbIdTipoReferenciaOrigen.BoundText
        mo_Formulario.MarcarComoVacio txtIdEstablecimientoOrigen
    End If
End Sub

Private Sub txtIdEstablecimientoOrigen_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoServicio
AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoServicio_LostFocus()
   If cmbIdTipoServicio.Text <> "" Then
       mo_cmbIdTipoServicio.BoundText = Val(Split(cmbIdTipoServicio.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoServicio
End Sub

Private Sub cmbIdTipoServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtEdadEnDias_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtEdadEnDias
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtEdadEnDias_LostFocus()
   mo_Formulario.MarcarComoVacio txtEdadEnDias
End Sub

Private Sub txtEdadEnDias_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub





'Private Sub txtPeso_KeyDown(KeyCode As Integer, Shift As Integer)
'    mo_Teclado.RealizarNavegacion KeyCode, txtPeso
'    AdministrarKeyPreview KeyCode
'
'End Sub

'Private Sub txtPresion_KeyDown(KeyCode As Integer, Shift As Integer)
'    mo_Teclado.RealizarNavegacion KeyCode, txtPresion
'    AdministrarKeyPreview KeyCode
'
'End Sub

Private Sub txtPrimerNombreBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombreBusqueda
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtPrimerNombreBusqueda_LostFocus()
   txtPrimerNombreBusqueda.Text = mo_Teclado.CapitalizarNombres(txtPrimerNombreBusqueda.Text)
   If Len(txtPrimerNombreBusqueda.Text) > 0 Then
      btnBuscarPaciente_Click
   End If
End Sub

Private Sub txtPrimerNombreBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub




Private Sub txtReferenciaD_Change()
End Sub





Private Sub txtReferenciaO_Change()
lbHuboCambioEnDato = True
End Sub

Private Sub txtReferenciaO_LostFocus()
        If lbHuboCambioEnDato = True Then
          sighEntidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, txtReferenciaO.Text
          lbHuboCambioEnDato = False
        End If
End Sub

Private Sub txtSegundoNombreBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtSegundoNombreBusqueda
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtSegundoNombreBusqueda_LostFocus()
   txtSegundoNombreBusqueda.Text = mo_Teclado.CapitalizarNombres(txtSegundoNombreBusqueda.Text)
   If Len(txtSegundoNombreBusqueda.Text) > 0 Then
      btnBuscarPaciente_Click
   End If
End Sub

Private Sub txtSegundoNombreBusqueda_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub CompletarDatosDeDiagnostico(txtCodigoDx As TextBox, lblDescripcionDx As TextBox)
Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
Dim oDODiagnostico As DODiagnostico

    oBusqueda.MostrarFormulario
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            txtCodigoDx.Text = oDODiagnostico.CodigoCIE2004
            txtCodigoDx.Tag = oDODiagnostico.idDiagnostico
            lblDescripcionDx = oDODiagnostico.descripcion
        Else
            txtCodigoDx.Text = ""
            txtCodigoDx.Tag = ""
            lblDescripcionDx = ""
        End If
    End If
    Set oBusqueda = Nothing
End Sub
Sub CompletarDatosDeDiagnosticoEnELLostFocus(txtCodigoDx As TextBox, lblDescripcionDx As TextBox)
    
    txtCodigoDx.Text = UCase(txtCodigoDx.Text)

   If txtCodigoDx.Text <> "" Then
    Dim oDODiagnostico As DODiagnostico
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorCodigoCIE2004(txtCodigoDx.Text, True)
        If Not oDODiagnostico Is Nothing Then
            txtCodigoDx.Tag = oDODiagnostico.idDiagnostico
            lblDescripcionDx = oDODiagnostico.descripcion
        Else
            txtCodigoDx.Tag = ""
            lblDescripcionDx = ""
        End If
   End If

End Sub

Sub CompletarDatosDeMedico(txtMedico As TextBox, lblNombreMedico As TextBox, lIdEspecialidad As Long, lcFiltraMedico As String, ldFechaProgramada As Date, lcHoraProgramada As String, lnIdTipoServicio As Long)
'Dim oBusqueda As New MedicosBusqueda
Dim oBusqueda As New SIGHNegocios.BuscaMedicos
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
     
    oBusqueda.IdEspecialidad = lIdEspecialidad
    oBusqueda.FechaProgramada = ldFechaProgramada
    oBusqueda.HoraProgramada = lcHoraProgramada
    oBusqueda.idTipoServicio = lnIdTipoServicio
    If mi_Opcion = sghAgregar Then
        oBusqueda.NombreMedico = lcFiltraMedico
    End If
    oBusqueda.NoMuestraInactivos = True
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
            txtMedico.Text = oDOEmpleado.CodigoPlanilla
            txtMedico.Tag = oDoMedico.idMedico
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDoMedico = Nothing
    Set oDOEmpleado = Nothing
    Set oDOEspecialidades = Nothing
End Sub

Sub CompletarDatosDeMedicoEgreso(txtMedico As TextBox, lblNombreMedico As TextBox, lIdEspecialidad As Long, lcFiltraMedico As String, ldFechaProgramada As Date, lcHoraProgramada As String, lnIdTipoServicio As Long)
'Dim oBusqueda As New MedicosBusqueda
Dim oBusqueda As New SIGHNegocios.BuscaMedicos
Dim oDoMedico As New DOMedico
Dim oDOEmpleado As New dOEmpleado
Dim oDOEspecialidades As New Collection
Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oBusqueda.IdEspecialidad = lIdEspecialidad
    oBusqueda.NombreMedico = lcFiltraMedico
    oBusqueda.FechaProgramada = ldFechaProgramada
    oBusqueda.HoraProgramada = lcHoraProgramada
    oBusqueda.idTipoServicio = lnIdTipoServicio
    'oBusqueda.Show 1
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
'        txtMedico.Tag = ""
'        lblNombreMedico.Text = ""
        If mo_AdminProgramacion.MedicosSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oDoMedico, oDOEmpleado, oDOEspecialidades, oConexion) Then
            txtMedico.Text = oDOEmpleado.CodigoPlanilla
            txtMedico.Tag = oDoMedico.idMedico
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDoMedico = Nothing
    Set oDOEmpleado = Nothing
    Set oDOEspecialidades = Nothing
End Sub


Sub CompletarDatosDeMedicoEnElLostFocus(txtMedico As TextBox, lblNombreMedico As TextBox)
Dim oMedicosEspecialidad As New Collection
    txtMedico = Trim(txtMedico)
    If txtMedico <> "" Then
        Dim oDOEmpleado As New dOEmpleado
        Dim oDoMedico As New DOMedico
        If mo_AdminProgramacion.MedicosSeleccionarPorCodigo1(CStr(txtMedico), oDoMedico, oDOEmpleado, oMedicosEspecialidad) Then
            txtMedico.Tag = oDoMedico.idMedico
            Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(oDoMedico.IdEmpleado)
            lblNombreMedico = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        Else
            txtMedico.Tag = ""
            lblNombreMedico = ""
        End If
        Set oDOEmpleado = Nothing
        Set oDoMedico = Nothing
    Else
    End If
    Set oMedicosEspecialidad = Nothing
End Sub

Sub CompletarDatosDeServicio(txtIdServicio As TextBox, lblDescripcionServicio As TextBox, lcFiltraServicio As String)
Dim oBusqueda As New SIGHNegocios.BuscaServicioHosp
Dim oDoServicio As New doServicio
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    oBusqueda.idTipoServicio = Val(mo_cmbIdTipoServicio.BoundText)
    oBusqueda.HabilitarTipoServicio = False
    If mi_Opcion = sghAgregar Then
        oBusqueda.NombreServicio = lcFiltraServicio
    End If
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.idRegistroSeleccionado, oConexion)
        
        'mgaray20140926
        Call SetVariableServicioUsaFUA(oDoServicio)
        If Not oDoServicio Is Nothing Then
            If Val(mo_cmbIdTipoServicio.BoundText) = oDoServicio.idTipoServicio Then
                'mgaray20140926
'                lcElServicioUsaGalenHos = IIf(oDoServicio.UsaGalenHos = True, "S", "N")

                txtIdServicio.Text = oDoServicio.Codigo
                txtIdServicio.Tag = oDoServicio.IdServicio
                lblDescripcionServicio = oDoServicio.nombre
                lblDescripcionServicio.Tag = oDoServicio.IdEspecialidad
                'mgaray20140926
'                If oDoServicio.UsaFUA = False Then
'                   lbElServicioRegistraFUA = "N"
'                   wxParametro302 = "N"
'                Else
'                   wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
'                   lbElServicioRegistraFUA = "S"
'                End If
'                InicializarFUA
                '
                txtIdMedicoIngreso.Text = ""
                lblNombreMedico.Text = ""
                If Val(mo_cmbIdTipoServicio.BoundText) = sghEmergenciaConsultorios Then   '09/08/2011
                    lblNroCamaIngreso.Visible = False
                    txtNroCamaIngreso.Visible = False
                    btnVerDisponibilidadDeCamas.Visible = False
                    mb_EsObservacionEmergencia = False
                    If oDoServicio.EsObservacionEmergencia = True Then
                        lblNroCamaIngreso.Visible = True
                        txtNroCamaIngreso.Visible = True
                        btnVerDisponibilidadDeCamas.Visible = True
                        mb_EsObservacionEmergencia = True
                    End If
                End If
                CalculaEmergenciaNumero
            Else
                MsgBox "El servicio seleccionado no pertenece a emergencia", vbInformation, Me.Caption
                txtIdServicio.Text = ""
                txtIdServicio.Tag = ""
                lblDescripcionServicio = ""
                lblDescripcionServicio.Tag = ""
                lcElServicioUsaGalenHos = "N"
                'mgaray20140926
                Call SetVariableServicioUsaFUA(Nothing)
            End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDoServicio = Nothing
    
End Sub

Sub CargaNombreDelServicioIngreso(IdServicio As Long)   'solo cuando se pulsa AGREGAR
        Dim oDoServicio As New doServicio
        Dim oConexion As New Connection
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(IdServicio, oConexion)
        If Not oDoServicio Is Nothing Then
            If Val(mo_cmbIdTipoServicio.BoundText) = oDoServicio.idTipoServicio Then
                txtIdServicioIngreso.Text = oDoServicio.Codigo
                txtIdServicioIngreso.Tag = oDoServicio.IdServicio
                lblNombreServicio = oDoServicio.nombre
                lblNombreServicio.Tag = oDoServicio.IdEspecialidad
            Else
                'MsgBox "El servicio seleccionado no pertenece a emergencia", vbInformation, Me.Caption
                txtIdServicioIngreso.Text = ""
                txtIdServicioIngreso.Tag = ""
                lblNombreServicio = ""
                lblNombreServicio.Tag = ""
            End If
        End If
        oConexion.Close
        Set oConexion = Nothing
        Set oDoServicio = Nothing
 
End Sub


Sub CompletarDatosDeServicioEnElLostFocus(txtIdServicio As TextBox, lblDescripcionServicio As TextBox)
    
    txtIdServicio.Text = UCase(txtIdServicio.Text)
    If txtIdServicio.Text <> "" Then
        Dim oDoServicio As doServicio
        Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(txtIdServicio.Text)
        'mgaray20140926
        Call SetVariableServicioUsaFUA(oDoServicio)
        
        If Not oDoServicio Is Nothing Then
            If mo_cmbIdTipoServicio.BoundText = oDoServicio.idTipoServicio Then
                txtIdServicio.Tag = oDoServicio.IdServicio
                lblDescripcionServicio = oDoServicio.nombre
                lblDescripcionServicio.Tag = oDoServicio.IdEspecialidad
            Else
                MsgBox "El servicio ingresado no pertenece es de emergencia", vbInformation, Me.Caption
                txtIdServicio.Tag = ""
                lblDescripcionServicio = ""
                lblDescripcionServicio.Tag = ""
                'mgaray20140926
                Call SetVariableServicioUsaFUA(Nothing)
            End If
        Else
            txtIdServicio.Tag = ""
            lblDescripcionServicio = ""
            lblDescripcionServicio.Tag = ""
        End If
   End If

End Sub
Sub CompletarDatosDeEstablecimiento(txtIdEstablecimiento As TextBox, lblNombreEstablecimiento As TextBox, lTipoReferencia As Long)
    
    If lTipoReferencia = 1 Then
        Dim oBusqueda As New SIGHNegocios.BuscaEstablecimientos
        Dim oDoEstablecimiento As New DOEstablecimiento
        oBusqueda.MostrarFormulario
        If oBusqueda.BotonPresionado = sghAceptar Then
        
            Set oDoEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
            If Not oDoEstablecimiento Is Nothing Then
                txtIdEstablecimiento.Tag = oDoEstablecimiento.IdEstablecimiento
                txtIdEstablecimiento.Text = oDoEstablecimiento.Codigo
                lblNombreEstablecimiento = oDoEstablecimiento.nombre
            Else
                txtIdEstablecimiento.Tag = ""
                txtIdEstablecimiento.Text = ""
                lblNombreEstablecimiento = ""
            End If
        End If
        Set oBusqueda = Nothing
        Set oDoEstablecimiento = Nothing
    Else
        Dim oBusquedaNM As New SIGHNegocios.BuscaEstablecNoMinsa
        Dim oDoEstablecimientoNM As New DOEstablecimientoNoMinsa
        oBusquedaNM.lcNombrePc = mo_lcNombrePc
        oBusquedaNM.idUsuario = ml_idUsuario
        oBusquedaNM.MostrarFormulario
        If oBusquedaNM.BotonPresionado = sghAceptar Then
            Set oDoEstablecimientoNM = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(oBusquedaNM.idRegistroSeleccionado)
            If Not oDoEstablecimientoNM Is Nothing Then
                txtIdEstablecimiento.Tag = oDoEstablecimientoNM.IdEstablecimientoNoMINSA
                txtIdEstablecimiento.Text = oDoEstablecimientoNM.IdEstablecimientoNoMINSA
                lblNombreEstablecimiento = oDoEstablecimientoNM.nombre
            Else
                txtIdEstablecimiento.Tag = ""
                txtIdEstablecimiento.Text = ""
                lblNombreEstablecimiento = ""
            End If
        End If
        Set oBusquedaNM = Nothing
        Set oDoEstablecimientoNM = Nothing
    End If

End Sub
Sub CompletarDatosDelEstablecimientoEnElLostFocus(txtIdEstablecimiento As TextBox, lblNombreEstablecimiento As TextBox, lTipoReferencia As Long)
    
    If txtIdEstablecimiento <> "" Then
        If lTipoReferencia = 1 Then
                Dim oDoEstablecimiento As New DOEstablecimiento
                If mo_AdminServiciosComunes.EstablecimientosSeleccionarPorCodigo(txtIdEstablecimiento.Text, oDoEstablecimiento) Then
                    txtIdEstablecimiento.Tag = oDoEstablecimiento.IdEstablecimiento
                    txtIdEstablecimiento.Text = oDoEstablecimiento.Codigo
                    lblNombreEstablecimiento = oDoEstablecimiento.nombre
                Else
                    txtIdEstablecimiento.Tag = ""
                    txtIdEstablecimiento = ""
                    lblNombreEstablecimiento = ""
                End If
                Set oDoEstablecimiento = Nothing
        Else
                Dim oDOEstablecimientoNoMinsa As New DOEstablecimientoNoMinsa
                Set oDOEstablecimientoNoMinsa = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(txtIdEstablecimiento.Text)
                If Not oDOEstablecimientoNoMinsa Is Nothing Then
                    txtIdEstablecimiento.Tag = oDOEstablecimientoNoMinsa.IdEstablecimientoNoMINSA
                    txtIdEstablecimiento.Text = oDOEstablecimientoNoMinsa.IdEstablecimientoNoMINSA
                    lblNombreEstablecimiento = oDOEstablecimientoNoMinsa.nombre
                Else
                    txtIdEstablecimiento.Tag = ""
                    txtIdEstablecimiento = ""
                    lblNombreEstablecimiento = ""
                End If
                Set oDOEstablecimientoNoMinsa = Nothing
        End If
    End If

End Sub
Sub CompletarDatosDelEstablecimientoEnElLoad(lIdEstablecimiento As Long, lIdEstablecimientoNoMinsa As Long, txtIdEstablecimiento As TextBox, lblNombreEstablecimiento As TextBox, lTipoReferencia As Long)
                
    If lTipoReferencia = 1 Then
        Dim oDoEstablecimiento As New DOEstablecimiento
         Set oDoEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(lIdEstablecimiento)
         If Not oDoEstablecimiento Is Nothing Then
             txtIdEstablecimiento.Text = oDoEstablecimiento.Codigo
             txtIdEstablecimiento.Tag = oDoEstablecimiento.IdEstablecimiento
             lblNombreEstablecimiento = oDoEstablecimiento.nombre
        Else
             txtIdEstablecimiento.Text = ""
             txtIdEstablecimiento.Tag = ""
             lblNombreEstablecimiento = ""
        End If
        Set oDoEstablecimiento = Nothing
    Else
        Dim oDOEstablecimientoNoMinsa As New DOEstablecimientoNoMinsa
         Set oDOEstablecimientoNoMinsa = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(lIdEstablecimientoNoMinsa)
         If Not oDOEstablecimientoNoMinsa Is Nothing Then
             txtIdEstablecimiento.Text = oDOEstablecimientoNoMinsa.IdEstablecimientoNoMINSA
             txtIdEstablecimiento.Tag = oDOEstablecimientoNoMinsa.IdEstablecimientoNoMINSA
             lblNombreEstablecimiento = oDOEstablecimientoNoMinsa.nombre
        Else
             txtIdEstablecimiento.Text = ""
             txtIdEstablecimiento.Tag = ""
             lblNombreEstablecimiento = ""
         End If
         Set oDOEstablecimientoNoMinsa = Nothing
    End If

End Sub


Sub SePrecionoF2EnDx(lnKeyCode As Integer)
    On Error Resume Next
    Select Case lnKeyCode
    Case vbKeyTab
       If btnAceptar.Enabled Then btnAceptar.SetFocus
    Case vbKeyF2
       If btnAceptar.Enabled Then btnAceptar_Click
    Case Else
       AdministrarKeyPreview lnKeyCode
    End Select

End Sub







Private Sub ucDiagnosticosIngreso_SePresionoTeclaEspecial(KeyCode As Integer)
   SePrecionoF2EnDx (KeyCode)
End Sub









Private Sub ucPacientesDetalle1_SeModificoFechaNacimiento(sFechaNacimiento As String, sHoraNacimiento As String)
    
    
    CambioFechaNacimiento sFechaNacimiento, sHoraNacimiento
    
    
    
End Sub

Sub CambioFechaNacimiento(sFechaNacimiento As String, sHoraNacimiento As String)
    On Error Resume Next
    Me.txtEdadEnDias = ""
    Dim oEdad As Edad
    oEdad = sighEntidades.CalcularEdad(CDate(sFechaNacimiento & " " & sHoraNacimiento), CDate(txtFechaIngreso & " " & txtHoraIngreso.Text))
    Me.txtEdadEnDias = oEdad.Edad
    mo_cmbIdTipoEdad.BoundText = oEdad.TipoEdad
    If Me.txtEdadEnDias.Text = "" Then
        Me.txtEdadEnDias.Text = Me.txtEdadEnDias.Tag
    End If
    Me.ucDiagnosticosIngreso.EdadPaciente = EdadEnDias(oEdad)
    Me.ucDiagnosticoNacimiento.EdadPaciente = EdadEnDias(oEdad)
    ActualizaCheckRecienNacido
End Sub

'Actualiza Check 'Recien Nacido', al ingresar al SERVICIO de Hospitalización
Sub ActualizaCheckRecienNacido()
    chkRecienNacido.Value = sighEntidades.CalculaSiEsRecienNacido(CDate(Me.ucPacientesDetalle1.FechaNacimiento & " " & Me.ucPacientesDetalle1.HoraNacimiento), CDate(txtFechaIngreso & " " & txtHoraIngreso.Text))
End Sub



Private Sub ucPacientesDetalle1_SeModificoPacienteNoIdentificado(bPacienteNoIdentificado As Boolean)

    If bPacienteNoIdentificado = True Then
        If mi_Opcion = sghAgregar Then
           chkPacienteNuevo.Value = 1
        End If
        chkPacienteNuevo.Enabled = False
        fraNotas.Caption = "Datos del acompañante"
        lbPacienteNN = True
    Else
        If mi_Opcion = sghAgregar Then
            chkPacienteNuevo.Enabled = True
            chkPacienteNuevo.Value = 1
        End If
        fraNotas.Caption = "Notas"
        lbPacienteNN = False
    End If

End Sub

Private Sub ucPacientesDetalle1_SeModificoSexo(lIdTipoSexo As Long)
    
    Me.ucDiagnosticosIngreso.SexoPaciente = lIdTipoSexo
    
End Sub

Private Sub ucPacientesDetalle1_SePresionoTeclaEspecial(KeyCode As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoAgenteAGAN_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoAgenteAGAN
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoAgenteAGAN_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdGrupoOcupacionalALAB_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdGrupoOcupacionalALAB
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdGrupoOcupacionalALAB_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdPosicionLesionadoALAB_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdPosicionLesionadoALAB
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdPosicionLesionadoALAB_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdUbicacionLesionado_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdUbicacionLesionado
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdUbicacionLesionado_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdTipoTransporte_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoTransporte
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoTransporte_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdTipoVehiculo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoVehiculo
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoVehiculo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdClaseAccidente_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdClaseAccidente
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdClaseAccidente_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdRelacionAgresorVictima_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdRelacionAgresorVictima
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdRelacionAgresorVictima_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdSeguridad_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdSeguridad
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdSeguridad_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdTipoEvento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoEvento
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoEvento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdLugarEvento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdLugarEvento
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdLugarEvento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub cmbIdCausaExternaMorbilidad_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdCausaExternaMorbilidad
AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdCausaExternaMorbilidad_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtNroCamaIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroCamaIngreso
    If KeyCode = vbKeyF1 Then
        btnVerDisponibilidadDeCamas_Click
    End If
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroCamaIngreso_LostFocus()
    txtNroCamaIngreso = UCase(Trim(txtNroCamaIngreso))
    CompletarDatosDeCamasEnElLostFocus txtNroCamaIngreso
    mo_Formulario.MarcarComoVacio txtNroCamaIngreso
End Sub

Private Sub txtNroCamaIngreso_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub CompletarDatosDeCamasEnElLostFocus(txtNroCama As TextBox)
    
'    txtNroCama.Tag = ""
'    txtNroCama.Text = UCase(txtNroCama.Text)
    
    If txtNroCama.Text <> "" Then
        Dim oDOCama As DOCama
        Set oDOCama = mo_AdminHoteleria.CamasSeleccionarPorCodigo(txtNroCama.Text)
            If Not oDOCama Is Nothing Then
                If Val(Me.txtIdServicioIngreso.Tag) = oDOCama.IdServicioUbicacionActual Then
                    txtNroCama.Tag = oDOCama.idCama
                Else
                    MsgBox "La cama seleccionada no pertenece al mismo servicio de ingreso", vbInformation, Me.Caption
                    txtNroCama.Tag = ""
                End If
            End If
    Else
    
    End If

End Sub

Sub ImprimePreCuenta()
    Dim oReporte As New RptCaja
    Dim lcPaciente As String
    Dim lcMedico As String
    If mi_Opcion <> sghAgregar Then
       Me.ucPacientesDetalle1.CargarDatosAlObjetoDatos mo_Pacientes, mo_Historia, mo_DoPacientesDatosAdd
    End If
    lcPaciente = Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & " " & Trim(mo_Pacientes.PrimerNombre)
    If mo_Pacientes.SegundoNombre <> "" Then
       lcPaciente = lcPaciente & " " & Trim(mo_Pacientes.SegundoNombre)
    End If
    If mo_Pacientes.TercerNombre <> "" Then
      lcPaciente = lcPaciente & " " & Trim(mo_Pacientes.TercerNombre)
    End If
    lcMedico = lblNombreMedico.Text

    oReporte.ImpresionPreCuenta txtFechaIngreso.Text, txtHoraIngreso.Text, lcPaciente, mo_Pacientes.NroHistoriaClinica, _
                                lblNombreServicio.Text, lcMedico, IIf(ml_TipoServicio = sghHospitalizacion, _
                                "HOSPITALIZACION", "EMERGENCIA"), mo_Atenciones.idAtencion, txtNroOrdenPago.Text, _
                                mo_Atenciones.idCuentaAtencion, cmbFuenteFinanciamiento.Text, "", ml_idUsuario, _
                                "Cama: " & txtNroCamaIngreso.Text, mo_Pacientes.FichaFamiliar, mo_Pacientes.idTipoNumeracion, _
                                wxParametro216, wxParametro306, False, _
                                IIf(txtIdMedicoIngreso.Tag = "", 0, CLng(txtIdMedicoIngreso.Tag))
    Set oReporte = Nothing
'    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub






Private Sub UcPacientesSunasa1_SePresionoTeclaEspecial(KeyCode As Integer)
    AdministrarKeyPreview KeyCode
End Sub














Private Sub ucSISfuaCodPrestacion1_GotFocus()
'    MsgBox "test"
End Sub

Private Sub ucTransferenciasDetalle1_SePresionoTeclaEspecial(KeyCode As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub ucTransferenciasDetalle1_UltimoServicioTransferido(lcUltimoCodigoServicio As String)
    ' Me.txtIdServicioEgreso.Tag = lcUltimoCodigoServicio
     lcUltimoCodigoDeServicioTransferido = lcUltimoCodigoServicio

End Sub

Sub LimpiaDatosDeBusqueda()
    If mi_Opcion = sghAgregar Then
        txtNroHistoriaBusqueda.Text = ""
        txtApellidoPaternoBusqueda.Text = ""
        txtApellidoMaternoBusqueda.Text = ""
        txtPrimerNombreBusqueda.Text = ""
        txtSegundoNombreBusqueda.Text = ""
        txtNroDNIBusqueda.Text = ""
    End If
End Sub

Sub LimpiarVariablesDeMemoria()

End Sub

Function CalculaEstanciaParaPacienteConAltaMedica(oDOAtencion As DOAtencion) As Boolean
        Dim lnDiasEstancia As Integer: Dim lnHorasEstancia As Integer
        Dim lcCodServicio As String: Dim lcHoraEstanciaMax  As String
        Dim lbEliminaConsultaEmerg As Boolean, lnIdServicio As Long
        Dim lcCodigoNombre As String, lnPrecioUnitario As Double
        Dim lnIdTipoFinanciamiento As Long, lnIdProducto As Long
        Dim lnTotalDiasEstancia As Long, lnTotalPagarEstancia As Double
        Dim oBuscaDiasPaciente As New SIGHDatos.Parametros
        Dim oRsTmp As New Recordset
        Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
        Dim oGenerarRecordsetProductos As New SighFacturacion.dllFactUcEstadoCuenta
        Dim oConexion As New Connection
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        If sighEntidades.EsFecha(oDOAtencion.fechaEgreso, "DD/MM/AAAA") Then
            lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
            'Dias de Estancia
            lnDiasEstancia = oBuscaDiasPaciente.DiasDelPacienteEnHospitalizacionEmergencia(oDOAtencion.FechaIngreso, oDOAtencion.HoraIngreso, oDOAtencion.fechaEgreso, oDOAtencion.HoraEgreso, lcHoraEstanciaMax)
    
            'Horas de Estancia
            lnHorasEstancia = sighEntidades.HorasDelPacienteEnHospitalizacionEmergencia(oDOAtencion.FechaIngreso, oDOAtencion.HoraIngreso, oDOAtencion.fechaEgreso, oDOAtencion.HoraEgreso)
            
            'Codigo Servicio
            If oDOAtencion.idTipoServicio = sghHospitalizacion Then
               lnIdServicio = Val(wxParametro202)
            ElseIf lnHorasEstancia < 12 Then   'Emergencia Menor a 12 horas
               lnIdServicio = Val(wxParametro203)
            ElseIf lnHorasEstancia <= 24 Then  'Emergencia Entre 12 y 24 horas
               lnIdServicio = Val(wxParametro204)
            Else                               'Emergencia >24 horas
               lnIdServicio = Val(wxParametro204)
            End If
            oGenerarRecordsetProductos.GenerarRecordsetProductos oRsEstancia
            If lnIdServicio > 0 Then
                If oDOAtencion.idTipoServicio = sghHospitalizacion Then
                     mo_AdminAdmision.GeneraEstanciaPorCadaServicioHospitalizado oDOAtencion.idCuentaAtencion, _
                                      oDOAtencion.fechaEgreso, oDOAtencion.HoraEgreso, oRsEstancia, _
                                      lnTotalPagarEstancia, lnTotalDiasEstancia, oConexion, True, False
                     With oRsTmp
                        .Fields.Append "IdProducto", adInteger
                        .Fields.Append "CantidadEstancia", adInteger
                        .Fields.Append "PrecioEstancia", adCurrency
                        .CursorType = adOpenDynamic
                        .LockType = adLockOptimistic
                        .Open
                     End With
                     If oRsEstancia.RecordCount > 0 Then
                        oRsEstancia.Sort = "idProducto"
                        oRsEstancia.MoveFirst
                        Do While Not oRsEstancia.EOF
                           lnIdProducto = oRsEstancia.Fields!idProducto
                           lnPrecioUnitario = oRsEstancia.Fields!PrecioEstancia
                           lnDiasEstancia = 0
                           Do While Not oRsEstancia.EOF And lnIdProducto = oRsEstancia.Fields!idProducto
                                lnDiasEstancia = lnDiasEstancia + oRsEstancia.Fields!CantidadEstancia
                                oRsEstancia.MoveNext
                                If oRsEstancia.EOF Then
                                   Exit Do
                                End If
                           Loop
                           oRsTmp.AddNew
                           oRsTmp.Fields!idProducto = lnIdProducto
                           oRsTmp.Fields!CantidadEstancia = lnDiasEstancia
                           oRsTmp.Fields!PrecioEstancia = lnPrecioUnitario
                           oRsTmp.Update
                        Loop
                        oRsEstancia.Close
                        Set oRsEstancia = oRsTmp
                     End If
                Else
                    '********cabecera/detalle
                     lnPrecioUnitario = 0
                     lnIdTipoFinanciamiento = oDOAtencion.IdFormaPago
                     oDoCatalogoServicioHosp.PrecioUnitario = 0
                     Set oDoCatalogoServicioHosp = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarPorId(lnIdServicio, lnIdTipoFinanciamiento)
                     If oDoCatalogoServicioHosp.PrecioUnitario = 0 And oDoCatalogoServicioHosp.SeUsaSinPrecio = False Then
                        lcCodigoNombre = mo_AdminServiciosComunes.DevuelveNombreMedicamentoOServicioSegunId(lnIdServicio, sghServicio)
                        MsgBox "Tiene problemas con el ID SERVICIO: " & lcCodigoNombre & Chr(13) & "para el Producto/Plan: " & oDOAtencion.IdFormaPago & Chr(13) & "...consulte al ADMINISTRADOR DEL SISTEMA....", vbInformation, "Error"
                        Exit Function
                     End If
                     lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
                     oRsEstancia.AddNew
                     oRsEstancia.Fields!idProducto = lnIdServicio
                     oRsEstancia.Fields!CantidadEstancia = lnDiasEstancia
                     oRsEstancia.Fields!PrecioEstancia = lnPrecioUnitario
                     oRsEstancia.Update
                End If
            End If
        End If
        CalculaEstanciaParaPacienteConAltaMedica = True
        Set oRsTmp = Nothing
        Set oDoCatalogoServicioHosp = Nothing
        Set oGenerarRecordsetProductos = Nothing
        oConexion.Close
        Set oConexion = Nothing
End Function

Sub InicilizarParametros()
    wxParametro202 = lcBuscaParametro.SeleccionaFilaParametro(202)
    wxParametro203 = lcBuscaParametro.SeleccionaFilaParametro(203)
    wxParametro204 = lcBuscaParametro.SeleccionaFilaParametro(204)
    wxParametro208 = lcBuscaParametro.SeleccionaFilaParametro(208)
    wxParametro210 = lcBuscaParametro.SeleccionaFilaParametro(210)
    wxParametro211 = lcBuscaParametro.SeleccionaFilaParametro(211)
    wxParametro212 = lcBuscaParametro.SeleccionaFilaParametro(212)
    wxParametro215 = lcBuscaParametro.SeleccionaFilaParametro(215)
    wxParametro216 = lcBuscaParametro.SeleccionaFilaParametro(216)
    wxParametro231 = lcBuscaParametro.SeleccionaFilaParametro(231)
    wxParametro232 = lcBuscaParametro.SeleccionaFilaParametro(232)
    wxParametro233 = lcBuscaParametro.SeleccionaFilaParametro(233)
    wxParametro237 = lcBuscaParametro.SeleccionaFilaParametro(237)
    wxParametro242 = lcBuscaParametro.SeleccionaFilaParametro(242)
    wxParametro259 = lcBuscaParametro.SeleccionaFilaParametro(259)
    wxParametro282 = lcBuscaParametro.SeleccionaFilaParametro(282)
    wxParametro287 = lcBuscaParametro.SeleccionaFilaParametro(287)
    wxParametro290 = lcBuscaParametro.SeleccionaFilaParametro(290)
    wxParametro291 = lcBuscaParametro.SeleccionaFilaParametro(291)
    wxParametro292 = lcBuscaParametro.SeleccionaFilaParametro(292)
    wxParametro296 = lcBuscaParametro.SeleccionaFilaParametro(296)
    wxParametro302 = lcBuscaParametro.SeleccionaFilaParametro(302)
    wxParametro306 = lcBuscaParametro.SeleccionaFilaParametro(306)
    wxParametro312 = lcBuscaParametro.SeleccionaFilaParametro(312)
    wxParametro316 = lcBuscaParametro.SeleccionaFilaParametro(316): wxParametro316 = IIf(IsNull(wxParametro316), "", wxParametro316)
    wxParametro317 = Trim(lcBuscaParametro.SeleccionaFilaParametro(317)): wxParametro317 = IIf(IsNull(wxParametro317), "", wxParametro317)
    wxParametro322 = lcBuscaParametro.SeleccionaFilaParametro(322)
    wxParametro323 = lcBuscaParametro.SeleccionaFilaParametro(323)
    wxParametro324 = lcBuscaParametro.SeleccionaFilaParametro(324)
    wxParametro333 = lcBuscaParametro.SeleccionaFilaParametro(333)
    wxParametro336 = lcBuscaParametro.SeleccionaFilaParametro(336)
    wxParametro351 = lcBuscaParametro.SeleccionaFilaParametro(351)
    wxParametro353 = lcBuscaParametro.SeleccionaFilaParametro(353)
    wxParametro357 = lcBuscaParametro.SeleccionaFilaParametro(357)
    wxParametro358 = lcBuscaParametro.SeleccionaFilaParametro(358)
    wxParametro359 = lcBuscaParametro.SeleccionaFilaParametro(359)
    wxParametro506 = UCase(lcBuscaParametro.SeleccionaFilaParametro(506))
    wxParametro521 = UCase(lcBuscaParametro.SeleccionaFilaParametro(521))
    wxParametro525 = lcBuscaParametro.SeleccionaFilaParametro(525)
    wxParametro526 = lcBuscaParametro.SeleccionaFilaParametro(526)
    wxParametro536 = lcBuscaParametro.SeleccionaFilaParametro(536)
    wxParametro545 = lcBuscaParametro.SeleccionaFilaParametro(545)
    wxParametro546 = lcBuscaParametro.SeleccionaFilaParametro(546)    'debb-03/04/2018
    wxParametro547 = lcBuscaParametro.SeleccionaFilaParametro(547)
    wxParametro552 = lcBuscaParametro.SeleccionaFilaParametro(552)
    wxParametro553 = lcBuscaParametro.SeleccionaFilaParametro(553)
    wxParametro559 = lcBuscaParametro.SeleccionaFilaParametro(559)
    wxParametroSIS = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    ldFechaActualServidor = lcBuscaParametro.RetornaFechaServidorSQL
    wxParametroJAMO = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    dxMorbilidadExterna342 = UCase(lcBuscaParametro.SeleccionaFilaParametro(342))
    wxParametroBusqRapida = UCase(lcBuscaParametro.SeleccionaFilaParametro(344))
End Sub


Private Sub chkBuscarEnSIS_Click()
    If chkBuscarEnSIS.Value = 1 Then
       fraBusqueda.Caption = "Solo se puede buscar en la WEB DEL SIS por DNI o por N°AFiliación"
       fraBusqueda.ForeColor = vbRed
       Me.txtNroDNIBusqueda.SetFocus
       mo_Formulario.HabilitarDeshabilitar txtNroHistoriaBusqueda, False
       mo_Formulario.HabilitarDeshabilitar chkPacienteNuevo, False 'A.Yañez 13112014
       btnBuscarPaciente_Click
       
    Else
       fraBusqueda.Caption = "Búsqueda"
       fraBusqueda.ForeColor = vbBlack
       Me.txtApellidoPaternoBusqueda.SetFocus
       mo_Formulario.HabilitarDeshabilitar txtNroHistoriaBusqueda, True
       mo_Formulario.HabilitarDeshabilitar chkPacienteNuevo, True 'A. Yañez 13112014
       chkPacienteNuevo.Enabled = True
    End If
End Sub

Private Sub UcSISafiliacion1_OnLostFocus(lcDisa As String, lcLote As String, lcNumero As String)
   lnAfiliacionSIS1 = lcDisa
   lnAfiliacionSIS2 = lcLote
   lnAfiliacionSIS3 = lcNumero
   Me.chkBuscarEnSIS.Value = 1
   btnBuscarPaciente_Click
   If lnAfiliacionSIS3 <> "" Then
      On Error Resume Next
      Me.grdPacientesEncontrados.SetFocus
   End If
End Sub

Sub ReglasDeConsistenciasDespuesDeElegirCodigoPrestacion()
   If wxParametro302 = "S" Then
      If Val(cmbFuenteFinanciamiento.BoundText) = sghFuenteFinanciamiento.sghFFSIS Then
            Dim oRsDestinoAtencion As New Recordset
            Dim lbReglasDeConsistenciaSISestanOK As Boolean
            lbReglasDeConsistenciaSISestanOK = mo_ReglasSISgalenhos.ReglasDeConsistenciaSISestanOK(mo_lnIdTablaLISTBARITEMS, wxParametro302, _
                                                                    Val(cmbFuenteFinanciamiento.BoundText), ms_MensajeError, mo_Atenciones.idCuentaAtencion, _
                                                                    Me.ucSISfuaCodPrestacion1.CodigoPrestacion, lcElServicioUsaGalenHos, _
                                                                    ml_TipoServicio, mi_Opcion, False, True, _
                                                                    Me.ucPacientesDetalle1.NroHistoriaClinica, Me.txtFechaIngreso.Text, _
                                                                    Me.txtHoraIngreso.Text, , oRsDestinoAtencion)
            If lbReglasDeConsistenciaSISestanOK = False Then
                 ucSISfuaCodPrestacion1.CodigoPrestacion = ""
                 ucSISfuaCodPrestacion1.Prestacion = ""
            End If
        Else
            cmbIdTipoServicio_Click
        End If
   End If
End Sub

Private Sub ucSISfuaCodPrestacion1_LostFocus()
     'debb-02/05/2016 (inicio)
     If mo_AdminServiciosComunes.FUAvalidaCodigoPrestacionSegunAdmision(mo_lnIdTablaLISTBARITEMS, ucSISfuaCodPrestacion1.CodigoPrestacion) = False Then
        ucSISfuaCodPrestacion1.CodigoPrestacion = ""
        Exit Sub
     End If
     'debb-02/05/2016 (fin)

    ReglasDeConsistenciasDespuesDeElegirCodigoPrestacion
End Sub

Private Sub setToolTipText()
    lblNombreServicio.ToolTipText = ""
    lblNombreMedico.ToolTipText = ""
    txtIdServicioIngreso.ToolTipText = ""
    txtIdMedicoIngreso.ToolTipText = ""
    
    If mi_Opcion = sghAgregar Then
        lblNombreServicio.ToolTipText = "Presionar Enter para buscar servicio"
        lblNombreMedico.ToolTipText = "Presionar Enter para buscar Médico"
        txtIdServicioIngreso.ToolTipText = "Presionar F1 para buscar servicio"
        txtIdMedicoIngreso.ToolTipText = "Presionar F1 para buscar Médico"
    End If
End Sub

Private Sub bloqueoControlesImpresionFicha()
    If Not (mo_DoAtencionDatosAdicionales Is Nothing) Then
            With mo_DoAtencionDatosAdicionales
    ''        If mo_DoAtencionDatosAdicionales.SeImprimioFicha = True Then
                mo_Formulario.HabilitarDeshabilitar cmbFuenteFinanciamiento, Not .SeImprimioFicha
'                mo_Formulario.HabilitarDeshabilitar cmbFormaPago, Not .SeImprimioFicha
    
                mo_Formulario.HabilitarDeshabilitar cmbIdTipoGravedad, Not .SeImprimioFicha
                mo_Formulario.HabilitarDeshabilitar cmbIdViasAdmision, Not .SeImprimioFicha
                mo_Formulario.HabilitarDeshabilitar txtFechaIngreso, Not .SeImprimioFicha
                mo_Formulario.HabilitarDeshabilitar txtHoraIngreso, Not .SeImprimioFicha
'            End If
        End With
    End If
    
End Sub

'yamill palomino
'Colocar en Entidades
Public Function ConvertirHoraEnMinutos(lcHora As String) As Long
Dim lnPosicion As Integer
Dim lnHora As Integer
Dim lnMinutos As Integer
ConvertirHoraEnMinutos = 0

'    lnPosicion = InStr(lcHora, ":") '12:05
'    lnHora = Mid(lcHora, 1, lnPosicion - 1)
'    lnMinutos = Mid(lcHora, lnPosicion + 1, 2)

    lnHora = CInt(Mid(lcHora, 1, 2))
    lnMinutos = CInt(Mid(lcHora, 4, 2))

    ConvertirHoraEnMinutos = 60 * lnHora + lnMinutos

End Function

'mgaray20140926
Private Function setListItemAControlDiagnosticos(IdListBarItem As Long)
    ucDiagnosticosIngreso.IdListBarItem = IdListBarItem
    ucDiagnosticoNacimiento.IdListBarItem = IdListBarItem
End Function

'debb-23/02/2015
Public Sub TraeDiagnosticosHasta24HorasDeEmergencia(lnIdPaciente As Long)
    If ml_TipoServicio = sghHospitalizacion Then
        'yamill palomino
        'CARGAR DIAGNOSICOS DE EMERGENCIA INGRESO-EGRESO
        
        Dim oRsBuscaAtencPac As New Recordset
        Dim lnMinutosTranscurridos As Long
        Dim lcDiasAtencion As String
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        Me.ucDiagnosticosIngreso.LimpiarDatos
        If lbProcedeDeConsExt = True Then
            '****proviene de Consultorios Externos, busca Dx de Atencion del Medico en CE
            Set oRsBuscaAtencPac = mo_AdminAdmision.AtencionesSeleccionarPorIdPaciente(lnIdPaciente, sghTipoServicio.sghConsultaExterna)
            If oRsBuscaAtencPac.RecordCount > 0 Then
                oRsBuscaAtencPac.MoveLast
                If oRsBuscaAtencPac.Fields!idEstado <> sghEstadoCuenta.sghAnulado Then
                        Me.ucDiagnosticosIngreso.idAtencion = oRsBuscaAtencPac.Fields!idAtencion
                        Me.ucDiagnosticosIngreso.TipoDiagnostico = sghHospitalizacionIngreso
                        Me.ucDiagnosticosIngreso.CargarDatosDeDiagnosticosEmergCE oConexion, 1
                End If
            End If
        Else
            '****proviene de Emergencia, busca Dx de Alta Médica
            Set oRsBuscaAtencPac = mo_AdminAdmision.AtencionesSeleccionarPorIdPaciente(lnIdPaciente, sghTipoServicio.sghEmergenciaConsultorios)
            If oRsBuscaAtencPac.RecordCount > 0 Then
                oRsBuscaAtencPac.MoveLast
                If Not (oRsBuscaAtencPac.Fields!idEstado = sghEstadoCuenta.sghAbierto Or oRsBuscaAtencPac.Fields!idEstado = sghEstadoCuenta.sghAnulado) Then
                        If Not IsNull(oRsBuscaAtencPac.Fields!fechaEgreso) Then 'FCV 14042015
                            lcDiasAtencion = Round(DateDiff("d", oRsBuscaAtencPac.Fields!fechaEgreso, lcBuscaParametro.RetornaFechaServidorSQL))
                        Else
                            lcDiasAtencion = 0
                        End If
                        If lcDiasAtencion = 0 Then
                                'llenar la grilla
                                Me.ucDiagnosticosIngreso.idAtencion = oRsBuscaAtencPac.Fields!idAtencion
                                Me.ucDiagnosticosIngreso.TipoDiagnostico = sghHospitalizacionIngreso
                                Me.ucDiagnosticosIngreso.CargarDatosDeDiagnosticosEmergCE oConexion, 3
                        Else
                            If lcDiasAtencion = 1 Then
                                lnMinutosTranscurridos = 1440 + ConvertirHoraEnMinutos(lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos) - ConvertirHoraEnMinutos(oRsBuscaAtencPac.Fields!HoraEgreso)
                                If lnMinutosTranscurridos <= 1440 Then
                                    'llenar la grilla
                                    Me.ucDiagnosticosIngreso.idAtencion = oRsBuscaAtencPac.Fields!idAtencion
                                    Me.ucDiagnosticosIngreso.TipoDiagnostico = sghHospitalizacionIngreso
                                    Me.ucDiagnosticosIngreso.CargarDatosDeDiagnosticosEmergCE oConexion, 3
                                End If
                            End If
                        End If
                End If
            End If
        End If
        oConexion.Close
        Set oRsBuscaAtencPac = Nothing
        Set oConexion = Nothing
    End If
End Sub

'mgaray20141008
Private Function BloquearEdicionAdmisionSegunReglas()
    Me.ucDiagnosticosIngreso.HabilitarEdicionDatos
    HabilitarControlesAdmision
    'If mi_Opcion = sghModificar Then
    '    Me.ucPacientesDetalle1.HabilitarFrames
        'If ms_ReglasSeguridad.TieneRolAdministrador(ml_idUsuario) = False Then
        '    Dim bEsMedico As Boolean
        '    bEsMedico = ms_ReglasSeguridad.UsuarioEsMedico(ml_idUsuario)
        '    If bEsMedico = True Then
        '        Me.ucPacientesDetalle1.DeshabilitarFrames
        '        DeshabilitarControlesAdmision
        '    Else
        '        Me.ucDiagnosticosIngreso.DeshabilitarEdicionDatos
        '    End If
        'End If
    'End If
    If mi_Opcion = sghModificar Then
         If wxParametro353 = "S" Then   'CASIMIRO ULLOA, JAMO
            If ms_ReglasSeguridad.TieneRolAdministrador(ml_idUsuario) = False Then
               'A.Yañez 30-10-2014 ***********************
               Dim bEsMedico As String
               '******************************************
               'A.Yañez 30-10-2014********************************************
               bEsMedico = ms_ReglasSeguridad.UsuarioEsMedicomejorado(ml_idUsuario)
               If bEsMedico = "1" Then
                   DeshabilitarControlesAdmision
                   Me.ucPacientesDetalle1.DeshabilitarFrames
                   Me.ucDiagnosticosIngreso.HabilitarEdicionDatos
               Else
                  If bEsMedico = "2" Then
                     DeshabilitarControlesAdmision
                     Me.ucPacientesDetalle1.DeshabilitarFrames
                     Me.ucDiagnosticosIngreso.DeshabilitarEdicionDatos
                  Else
                     Me.ucDiagnosticosIngreso.DeshabilitarEdicionDatos
                  End If
               End If
            End If
         End If
    End If
    '************************************************************************
End Function

Private Function DeshabilitarControlesAdmision()
    Frame2.Enabled = False
    Frame7.Enabled = False
    btnQuitarMadre.Enabled = False
    'controles que se habilitan segun datos elegido en otros controles
    HabilitarFrameOrigen False
    cmdBuscaMadre.Enabled = False
End Function

Private Function HabilitarControlesAdmision()
    Frame2.Enabled = True
    Frame7.Enabled = True
    btnQuitarMadre.Enabled = True
End Function
'mgaray20141023
Private Function LimpiarVariablesEnMemoria()
    mo_Atenciones.idPaciente = 0
End Function


'debb-23/02/2015
Private Sub btnProvCE_Click()
       fraBusqueda.Enabled = False
       btnBuscarPaciente.Enabled = False
       lbProcedeDeEmergencia = False
       lbProcedeDeConsExt = True
       AtencionesSinAdmHospitalizacion False
End Sub
'debb-23/02/2015
Private Sub btnProvEmergencia_Click()
       fraBusqueda.Enabled = False
       btnBuscarPaciente.Enabled = False
       AtencionesSinAdmHospitalizacion True
       lbProcedeDeEmergencia = True
       lbProcedeDeConsExt = False
End Sub


'debb-01/04/2015
Sub AtencionesSinAdmHospitalizacion(lbProcedeDeEmergencia As Boolean)
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Me.MousePointer = 11
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
           Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSinAdmHospitalizacion"
        Set oParameter = .CreateParameter("@ProvieneDeEmergencia", adInteger, adParamInput, 0, IIf(lbProcedeDeEmergencia = True, 2, 1)): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
   '
   With grdPacientesEncontrados
         .Left = 240
         .Top = 1080
         .Width = 11775
         .Height = 4455
   End With
   grdPacientesEncontrados.Caption = "Pacientes que provenien de " & IIf(lbProcedeDeEmergencia = True, "Emergencia", "Consultorios Externos")
   Set grdPacientesEncontrados.DataSource = oRecordset
   Me.grdPacientesEncontrados.Visible = True
   If oRecordset.RecordCount > 0 Then

        grdPacientesEncontrados.SetFocus
   End If
   Me.MousePointer = 1
Exit Sub
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
   Me.MousePointer = 1
Exit Sub
End Sub






Sub CargaCPTrealizadosEnElServicio()
'    Set oRsServiciosIntermedios = mo_AdminAdmision.BuscaAtencionesCptCEparaFormatoHIS(ml_idCuentaAtencion)
'    Set grdOtrosCpt.DataSource = oRsServiciosIntermedios
    Set grdOtrosCpt.DataSource = mo_AdminAdmision.BuscaAtencionesCptCEparaFormatoHIS(ml_idCuentaAtencion, sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
End Sub
Private Sub grdOtrosCpt_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
    
End Sub

Private Sub grdOtrosCpt_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
        grdOtrosCpt.Bands(0).Columns("IdPuntoCarga").Hidden = True
        grdOtrosCpt.Bands(0).Columns("IdCuentaAtencion").Hidden = True
        'mgaray20141022
        grdOtrosCpt.Bands(0).Columns("IdProducto").Hidden = True
        grdOtrosCpt.Bands(0).Columns("IdOrden").Width = 900
        grdOtrosCpt.Bands(0).Columns("codigo").Width = 800
        
        grdOtrosCpt.Bands(0).Columns("cantidad").Width = 300
        grdOtrosCpt.Bands(0).Columns("precio").Width = 600
        grdOtrosCpt.Bands(0).Columns("total").Width = 600
        
        grdOtrosCpt.Bands(0).Columns("labConfHIS").Header.Caption = "Lab"
        grdOtrosCpt.Bands(0).Columns("labConfHIS").Width = 600
        
        'mgaray201412a
        grdOtrosCpt.Bands(0).Columns("labConfHIS").Hidden = True
        grdOtrosCpt.Bands(0).Columns("Nombre").Width = grdOtrosCpt.Width - grdOtrosCpt.Bands(0).Columns("codigo").Width _
                                                    - grdOtrosCpt.Bands(0).Columns("cantidad").Width _
                                                    - grdOtrosCpt.Bands(0).Columns("precio").Width _
                                                    - grdOtrosCpt.Bands(0).Columns("total").Width _
                                                    - 700 'grdOtrosCpt.Bands(0).Columns("labConfHIS").Width - 700
End Sub

Private Sub btnAgregarCpt_Click()
    Dim oCpt As New FacOrdenServicioDetalle
    oCpt.FormMostradoDesde = 1
    oCpt.lbNOValidaCodigoPrestacion = True
    oCpt.PuntoCarga = 1   'consumo en el servicio
    oCpt.Opcion = sghAgregar
    oCpt.idUsuario = ml_idUsuario
    oCpt.idCuentaAtencion = ml_idCuentaAtencion
    oCpt.Show 1
    Set oCpt = Nothing
    CargaCPTrealizadosEnElServicio
End Sub




Private Sub grdApoyoDx_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdApoyoDx.Override.RowSizing = ssRowSizingFree
    ConfiguraGrillas 3

End Sub



Private Sub grdApoyoDx_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    On Error GoTo Fin
    If KeyAscii = 13 Then
  
        Dim ml_IdPruebaSeleccionada As String
        Dim ml_NombrePruebaSeleccionada As String
        Dim ml_nombrePaciente As String
        Dim ml_idOrden As Long
        Dim ml_IdProducto As Long
        Dim ml_NombreMedico As String
        Dim ml_areaTrabajo As Long
        Dim ml_idOrdenLab As Long
        Dim oRsTmp As New Recordset
        Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
        
        'Cargar los formularios para el resultado
        Set oRsTmp = grdApoyoDx.DataSource
        If oRsTmp.Fields!resultado <> "SI" Then
           Set oRsTmp = Nothing
           Set mo_ReglasLaboratorio = Nothing
           Exit Sub
        End If
        ml_IdPruebaSeleccionada = oRsTmp("codigo")
        ml_NombrePruebaSeleccionada = oRsTmp("item")
        ml_nombrePaciente = Me.ucPacientesDetalle1.DevuelvePaciente  '   txtPaciente.Text
        ml_idOrden = oRsTmp("idOrden")
        ml_IdProducto = oRsTmp("idProducto")
        '************(inicio) usa el nuevo formulario para llenar e imprimir RESULTADOS **********************
        Dim oRsTmp1 As New Recordset
        Set oRsTmp1 = mo_ReglasLaboratorio.LabItemsCptSeleccionarXfiltro("dbo.FactCatalogoServicios.Codigo='" & ml_IdPruebaSeleccionada & "'")
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.Close
           Dim oResultadoXitems As New SIGHLaboratorio.ResultadoXitems
           oResultadoXitems.IdOrden = ml_idOrden
           oResultadoXitems.idProductoCpt = ml_IdProducto
           oResultadoXitems.NoMuestraBotonGrabar = True
           oResultadoXitems.MostrarFormulario
           Set oResultadoXitems = Nothing
           Set oRsTmp1 = Nothing
           Set oRsTmp = Nothing
           Set mo_ReglasLaboratorio = Nothing
           Exit Sub
        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
        '************(fin) usa el nuevo formulario para llenar e imprimir RESULTADOS **********************
        
        Dim oMuestraResultado As New SIGHLaboratorio.Ingresos
        Dim ml_idTipoSexo As Long
        ml_idTipoSexo = mo_Pacientes.idTipoSexo
        oMuestraResultado.MuestraResultadoDelExamen ml_IdPruebaSeleccionada, ml_NombrePruebaSeleccionada, _
                                                    ml_nombrePaciente, ml_idOrden, ml_IdPaciente, ml_NombreMedico, _
                                                    ml_areaTrabajo, ml_idOrdenLab, ml_idTipoSexo, True
        Set oMuestraResultado = Nothing
        Set oRsTmp = Nothing
        Set mo_ReglasLaboratorio = Nothing

    End If
Fin:
End Sub

Sub ConfiguraGrillas(lnGrilla As Long)
    
    If lnGrilla = 3 Then
        grdApoyoDx.Bands(0).Columns("idCuentaAtencion").Hidden = True
        grdApoyoDx.Bands(0).Columns("IdOrden").Hidden = True
        grdApoyoDx.Bands(0).Columns("idProducto").Hidden = True
        grdApoyoDx.Bands(0).Columns("Codigo").Hidden = True
        grdApoyoDx.Bands(0).Columns("Fecha").Width = 800
        grdApoyoDx.Bands(0).Columns("hora").Width = 400
        grdApoyoDx.Bands(0).Columns("servicioApDx").Width = 700
        grdApoyoDx.Bands(0).Columns("item").Width = 3000
        grdApoyoDx.Bands(0).Columns("cantidad").Width = 600
        grdApoyoDx.Bands(0).Columns("cantidad").Header.Caption = "CantDespachada"
        grdApoyoDx.Bands(0).Columns("resultado").Width = 3400
        grdApoyoDx.Bands(0).Columns("especialista").Width = 1500
        grdApoyoDx.Bands(0).Columns("resultado").CellMultiLine = ssCellMultiLineTrue
        grdApoyoDx.Bands(0).Columns("item").CellMultiLine = ssCellMultiLineTrue
        grdApoyoDx.Bands(0).Columns("Receta").Width = 800
        grdApoyoDx.Bands(0).Columns("FReceta").Width = 1200
        grdApoyoDx.Bands(0).Columns("docDespacho").Width = 900
        grdApoyoDx.Bands(0).Columns("CantPedida").Width = 600
        grdApoyoDx.Bands(0).Columns("CantPedida").Header.Caption = "CantRecetada"
        grdApoyoDx.Caption = ""
    End If
End Sub

Sub CargaApoyoDx()
   Dim oRsTmp1 As New Recordset
   Set oRsTmp1 = mo_AdminAdmision.ServiciosIntermediosSeleccionarPorPaciente(ml_IdPaciente, True, False, ml_idCuentaAtencion, _
                                  False, False)
   Set grdApoyoDx.DataSource = oRsTmp1
   Set oRsTmp1 = Nothing
End Sub

Private Sub grdOtrosCpt_Click()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdOtrosCpt.DataSource
    On Error Resume Next
    ml_idOrden = rsRecordset("IdOrden")
    Set rsRecordset = Nothing
End Sub

Private Sub grdOtrosCpt_AfterRowActivate()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdOtrosCpt.DataSource
    On Error Resume Next
    ml_idOrden = rsRecordset("IdOrden")
    Set rsRecordset = Nothing
End Sub
Private Sub grdApoyoDx_Click()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdApoyoDx.DataSource
    On Error Resume Next
    ml_FechaReceta = rsRecordset("fReceta")
    Set rsRecordset = Nothing
End Sub
Private Sub grdApoyoDx_AfterRowActivate()
    Dim rsRecordset As ADODB.Recordset
    Set rsRecordset = grdApoyoDx.DataSource
    On Error Resume Next
    ml_FechaReceta = rsRecordset("fReceta")
    Set rsRecordset = Nothing
End Sub

Private Sub btnModificar_Click()
    Dim oReceta As New RecetaDetalle
    oReceta.Opcion = sghModificar
    oReceta.FechaReceta = ml_FechaReceta
    oReceta.idTipoServicio = ml_TipoServicio
    oReceta.idUsuario = ml_idUsuario
    oReceta.idCuentaAtencion = ml_idCuentaAtencion
    oReceta.IdMedicoServicioActual = mo_Atenciones.IdMedicoIngreso
    oReceta.Show 1
    Set oReceta = Nothing
    CargaApoyoDx
End Sub

Private Sub Enfermeras_Click()
    Dim mo_VisitasEnfermeras As New VisitasEnfermeras
    mo_VisitasEnfermeras.Opcion = sghConsultar
    mo_VisitasEnfermeras.idCuentaAtencion = ml_idCuentaAtencion
    mo_VisitasEnfermeras.TipoServicio = ml_TipoServicio
    mo_VisitasEnfermeras.lcNombrePc = ""
    mo_VisitasEnfermeras.lnIdTablaLISTBARITEMS = 302
    mo_VisitasEnfermeras.lbNuevoMovimiento = False
    mo_VisitasEnfermeras.CargaUnaSolaVez = True
    mo_VisitasEnfermeras.idUsuario = 0
    mo_VisitasEnfermeras.Show 1
    Set mo_VisitasEnfermeras = Nothing
End Sub

Private Function getDatosDeServicio(lIdServicio As Long) As doServicio
    Dim oDoServicio As New doServicio
    Dim oConexion As New Connection
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oDoServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(lIdServicio, oConexion)
    
    Set getDatosDeServicio = oDoServicio
    oConexion.Close
    Set oConexion = Nothing
End Function


Private Sub ucNacimientoDetalle1_SePresionoTeclaEspecial(KeyCode As Integer)
    Select Case KeyCode
    Case 1000  'Se pulso boton AGREGAR
        Dim oEdad As Edad
        oEdad = sighEntidades.CalcularEdad(ucNacimientoDetalle1.FechaNacimiento, _
                         CDate(Format(txtFechaIngreso.Text, sighEntidades.DevuelveFechaSoloFormato_DMY) & " " & txtHoraIngreso.Text))
        Me.ucDiagnosticoNacimiento.EdadPaciente = sighEntidades.EdadEnDias(oEdad)
        Me.ucDiagnosticoNacimiento.SexoPaciente = ucNacimientoDetalle1.idTipoSexo
    End Select
End Sub

Private Sub txtIdMedicoNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdMedicoNacimiento
    If KeyCode = vbKeyF1 Then
        btnMedicoRespNacimiento_Click
    End If
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdMedicoNacimiento_LostFocus()
    CompletarDatosDeMedicoEnElLostFocus txtIdMedicoNacimiento, lblNombreMedicoNacimiento
    mo_Formulario.MarcarComoVacio txtIdMedicoNacimiento
End Sub
Private Sub txtIdMedicoNacimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'debb-21/06/2016
Private Sub txtReferenciaO_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtReferenciaO
End Sub
'debb-21/06/2016
Private Sub txtNombreAcompañante_KeyDown(KeyCode As Integer, Shift As Integer)
     mo_Teclado.RealizarNavegacion KeyCode, txtNombreAcompañante
End Sub
'debb-21/06/2016
Private Sub txtDNIacompaniante_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDNIacompaniante
End Sub
'debb-21/06/2016
Private Sub txtEmergenciaN_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtEmergenciaN
End Sub
'debb-21/06/2016
Sub CalculaEmergenciaNumero()
    If mi_Opcion = sghAgregar And ml_TipoServicio = sghEmergenciaConsultorios Then
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        If sighEntidades.Parametro550 = "S" Then
            Dim lcNuevoCorrelativo As String
            
            Call mo_AdminServiciosComunes.ServiciosGrabaYdevueveCorrelativoEmergencia(True, lcNuevoCorrelativo, Val(Me.txtIdServicioIngreso.Tag), oConexion)
            txtEmergenciaN.Text = lcNuevoCorrelativo
        Else
    
            Dim oRsTmp As New Recordset
            Set oRsTmp = mo_AdminAdmision.atencionesDatosAdicionalesXfiltro("  left(dbo.AtencionesDatosAdicionales.emergenciaCorrelativo,4)='" & _
                         Mid(lcBuscaParametro.RetornaFechaHoraServidorSQL, 7, 4) & "' order by dbo.AtencionesDatosAdicionales.emergenciaCorrelativo desc", _
                         oConexion)
            txtEmergenciaN.Text = Trim(Str(oRsTmp.RecordCount + 1))
'            If oRsTmp.RecordCount = 0 Then
'               txtEmergenciaN.Text = "1"
'            Else
'               txtEmergenciaN.Text = Trim(Str(Val(Mid(oRsTmp!emergenciaCorrelativo, 5, 10)) + 1))
'            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
        End If
        oConexion.Close
        Set oConexion = Nothing
    End If
End Sub

'debb-22/08/2016
Function ChequeaQueNoExistaPacienteServicioFecha() As Boolean
    ChequeaQueNoExistaPacienteServicioFecha = True
    If wxParametro536 = "S" Then Exit Function
    
    If mi_Opcion = sghAgregar And ml_IdPaciente > 0 And Val(txtIdServicioIngreso.Tag) > 0 And _
                          sighEntidades.EsFecha(txtFechaIngreso.Text, "DD/MM/AAAA") Then
        Dim oRsTmp87 As New Recordset
        Set oRsTmp87 = mo_AdminAdmision.AtencionesSeleccionarPorIdPaciente(ml_IdPaciente, ml_TipoServicio)
        oRsTmp87.Filter = "idServicioIngreso=" & txtIdServicioIngreso.Tag & " and fechaIngreso='" & txtFechaIngreso.Text & "'"
        If oRsTmp87.RecordCount > 0 And Me.txtNroCuenta.Text = "" Then
            MsgBox "YA EXISTE una ADMISION para ese PACIENTE/SERVICIO/FECHA con la Cuenta N° " & oRsTmp87!idCuentaAtencion, vbInformation, Me.Caption
            ChequeaQueNoExistaPacienteServicioFecha = False
        End If
        oRsTmp87.Close
        Set oRsTmp87 = Nothing
    End If
End Function

Sub CargaHCyPaciente(lnHistoria As Long, lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String)
    lcHistoriaYpaciente = "(" & Trim(Str(lnHistoria)) & ") " & Trim(lcApellidoPaterno) & _
    " " & Trim(lcApellidoMaterno) & " " & Trim(lcPrimerNombre)
End Sub

'franklin 2017
Sub BuscaMedicoRerencia(lcIdColegio As String)
    If Len(txtMedicoRef.Text) >= 1 Then
        Dim oRsTmp112 As New Recordset
        Dim lnId As Integer, lnIdex As Integer
        Set oRsTmp112 = mo_ReglasSISgalenhos.a_resatencionSeleccionarPorColegiatura(txtMedicoRef.Text)
        cmbMedicoRef.Clear
        lnId = 1: lnIdex = 0
        If oRsTmp112.RecordCount > 0 Then
           oRsTmp112.MoveFirst
           Do While Not oRsTmp112.EOF
              If Val(oRsTmp112!pers_IdTipoPersonalSalud) = Val(lcIdColegio) Then
                 lnIdex = lnId
              End If
              cmbMedicoRef.AddItem oRsTmp112!Medico
              oRsTmp112.MoveNext
              lnId = lnId + 1
           Loop
           If oRsTmp112.RecordCount = 1 Then
              cmbMedicoRef.ListIndex = 0
           ElseIf lnIdex > 0 Then
              cmbMedicoRef.ListIndex = lnIdex - 1
           End If
        End If
        oRsTmp112.Close
        Set oRsTmp112 = Nothing
    End If
End Sub

