VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Object = "{0002E558-0000-0000-C000-000000000046}#1.1#0"; "OWC11.DLL"
Begin VB.Form VisitasEnfermeras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VISITAS - ENFERMERAS"
   ClientHeight    =   10395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VistasEnfermeras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   16380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraHistoricoControles 
      Height          =   9540
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   2295
      Begin UltraGrid.SSUltraGrid grdHistoricoVisitas 
         Height          =   8850
         Left            =   0
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   120
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   15610
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         RowConnectorColor=   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Histórico de visitas"
      End
      Begin Threed.SSCommand btnAgregarVisita 
         Height          =   465
         Left            =   30
         TabIndex        =   65
         ToolTipText     =   "Agregar visita"
         Top             =   9000
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   820
         _Version        =   262144
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "VistasEnfermeras.frx":0CCA
         Caption         =   "Agregar"
         PictureAlignment=   9
      End
      Begin Threed.SSCommand btnQuitarVisita 
         Height          =   465
         Left            =   1200
         TabIndex        =   66
         ToolTipText     =   "Quitar Visita"
         Top             =   9000
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   820
         _Version        =   262144
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "VistasEnfermeras.frx":3C56
         Caption         =   "Quitar"
         PictureAlignment=   9
         ShapeSize       =   1
      End
   End
   Begin VB.Frame FrameVisita 
      Caption         =   "VISITA Nº 1"
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
      Height          =   9540
      Left            =   2370
      TabIndex        =   16
      Top             =   30
      Width           =   13965
      Begin VB.Frame frmTratamiento 
         Caption         =   "Administración de medicamentos por parte del profesional de salud"
         Height          =   2340
         Left            =   90
         TabIndex        =   68
         Top             =   5760
         Width           =   9495
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1950
            Left            =   10
            ScaleHeight     =   1950
            ScaleWidth      =   9405
            TabIndex        =   69
            Top             =   290
            Width           =   9400
            Begin VB.PictureBox pcTratamiento 
               Height          =   1935
               Left            =   120
               ScaleHeight     =   1875
               ScaleWidth      =   8835
               TabIndex        =   71
               Top             =   0
               Width           =   8895
               Begin VB.TextBox txtDosisProrenata 
                  Height          =   315
                  Index           =   0
                  Left            =   4680
                  MaxLength       =   2
                  TabIndex        =   85
                  Top             =   630
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.CheckBox chkDosis6 
                  Caption         =   "6"
                  Height          =   315
                  Index           =   0
                  Left            =   6720
                  TabIndex        =   81
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.CheckBox chkDosis5 
                  Caption         =   "5"
                  Height          =   315
                  Index           =   0
                  Left            =   6120
                  TabIndex        =   80
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.CheckBox chkDosis4 
                  Caption         =   "4"
                  Height          =   315
                  Index           =   0
                  Left            =   5520
                  TabIndex        =   79
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.CheckBox chkDosis3 
                  Caption         =   "3"
                  Height          =   315
                  Index           =   0
                  Left            =   4920
                  TabIndex        =   78
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.CheckBox chkDosis2 
                  Caption         =   "2"
                  Height          =   315
                  Index           =   0
                  Left            =   4320
                  TabIndex        =   77
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.CheckBox chkDosis1 
                  Caption         =   "1"
                  Height          =   315
                  Index           =   0
                  Left            =   3720
                  TabIndex        =   72
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.Label lblDosis6 
                  Caption         =   "6"
                  Height          =   315
                  Index           =   0
                  Left            =   7000
                  TabIndex        =   91
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.Label lblDosis5 
                  Caption         =   "5"
                  Height          =   315
                  Index           =   0
                  Left            =   6400
                  TabIndex        =   90
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.Label lblDosis4 
                  Caption         =   "4"
                  Height          =   315
                  Index           =   0
                  Left            =   5800
                  TabIndex        =   89
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.Label lblDosis3 
                  Caption         =   "3"
                  Height          =   315
                  Index           =   0
                  Left            =   5200
                  TabIndex        =   88
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.Label lblDosis2 
                  Caption         =   "2"
                  Height          =   315
                  Index           =   0
                  Left            =   4600
                  TabIndex        =   87
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.Label lblDosis1 
                  Caption         =   "1"
                  Height          =   315
                  Index           =   0
                  Left            =   4000
                  TabIndex        =   86
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.Label lblProrenata 
                  Caption         =   "PRORENATA"
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
                  Index           =   0
                  Left            =   3720
                  TabIndex        =   84
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label Label9 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000C&
                  Caption         =   "Observación"
                  ForeColor       =   &H8000000E&
                  Height          =   255
                  Left            =   7320
                  TabIndex        =   83
                  Top             =   0
                  Width           =   1545
               End
               Begin VB.Label lblObsMedicamento 
                  Caption         =   "Anulado por Médico"
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
                  Index           =   0
                  Left            =   7330
                  TabIndex        =   82
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   1695
               End
               Begin VB.Label lblMedicamento 
                  Caption         =   "MEDICAMENTO"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   76
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   3660
               End
               Begin VB.Label Label6 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000C&
                  Caption         =   "Dosis"
                  ForeColor       =   &H8000000E&
                  Height          =   255
                  Left            =   3720
                  TabIndex        =   75
                  Top             =   0
                  Width           =   3585
               End
               Begin VB.Label Label7 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000C&
                  Caption         =   "Medicamento"
                  ForeColor       =   &H8000000E&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   74
                  Top             =   0
                  Width           =   3705
               End
            End
            Begin VB.VScrollBar vsTratamiento 
               Height          =   1935
               Left            =   9050
               TabIndex        =   70
               Top             =   0
               Width           =   375
            End
         End
      End
      Begin VB.ListBox LisBoxVariable 
         Height          =   780
         Index           =   0
         ItemData        =   "VistasEnfermeras.frx":60D8
         Left            =   6840
         List            =   "VistasEnfermeras.frx":60DF
         Style           =   1  'Checkbox
         TabIndex        =   64
         Top             =   2760
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.TextBox txtServicio 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11160
         TabIndex        =   6
         Top             =   240
         Width           =   2715
      End
      Begin VB.TextBox txtNroHistoria 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7275
         TabIndex        =   2
         Top             =   240
         Width           =   1290
      End
      Begin VB.TextBox txtNroCuenta 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1170
         TabIndex        =   1
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox txtNroCama 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9495
         MaxLength       =   9
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtPrimerNombre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2265
         TabIndex        =   3
         Top             =   240
         Width           =   3930
      End
      Begin VB.VScrollBar scrollGraficos 
         Height          =   6000
         Left            =   13560
         TabIndex        =   57
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtVariable 
         Height          =   360
         Index           =   0
         Left            =   7080
         MaxLength       =   9
         TabIndex        =   53
         Top             =   3000
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.CheckBox chkVariable 
         Height          =   360
         Index           =   0
         Left            =   7080
         TabIndex        =   52
         Top             =   3000
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.ComboBox cmbVariable 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   3000
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.CommandButton btnBuscarEmpleado 
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
         Height          =   360
         Left            =   6270
         TabIndex        =   13
         Top             =   2130
         Width           =   315
      End
      Begin VB.TextBox txtNombreEmpleado 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6585
         TabIndex        =   36
         ToolTipText     =   "Profesional de salud que realizo la visita"
         Top             =   2145
         Width           =   7215
      End
      Begin VB.Frame Frame16 
         Caption         =   "Antecedentes del paciente"
         Height          =   1455
         Left            =   135
         TabIndex        =   28
         Top             =   660
         Width           =   13800
         Begin VB.TextBox txtantecedObstetrico 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   550
            Left            =   5550
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   810
            Width           =   3560
         End
         Begin VB.TextBox txtantecedFamiliar 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   550
            Left            =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   810
            Width           =   3560
         End
         Begin VB.TextBox txtantecedPatologico 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   550
            Left            =   10120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   240
            Width           =   3560
         End
         Begin VB.TextBox txtantecedAlergico 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   550
            Left            =   5550
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   240
            Width           =   3560
         End
         Begin VB.TextBox txtantecedQuirurgico 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   550
            Left            =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   240
            Width           =   3560
         End
         Begin VB.TextBox txtAntecedentes 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   550
            Left            =   10120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   810
            Width           =   3560
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Obstétricos"
            Height          =   210
            Left            =   4620
            TabIndex        =   34
            Top             =   840
            Width           =   930
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Familiares"
            Height          =   210
            Left            =   90
            TabIndex        =   33
            Top             =   840
            Width           =   750
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Patológicos"
            Height          =   210
            Left            =   9200
            TabIndex        =   32
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alergias "
            Height          =   210
            Left            =   4620
            TabIndex        =   31
            Top             =   360
            Width           =   675
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quirúrgicos"
            Height          =   210
            Left            =   90
            TabIndex        =   30
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Otros"
            Height          =   210
            Left            =   9200
            TabIndex        =   29
            Top             =   840
            Width           =   450
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Observaciones"
         Height          =   1320
         Left            =   90
         TabIndex        =   26
         Top             =   8115
         Width           =   9525
         Begin VB.TextBox txtObservaciones 
            ForeColor       =   &H00000000&
            Height          =   1020
            Left            =   80
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   240
            Width           =   9375
         End
      End
      Begin VB.Frame FraValorizacion 
         Caption         =   "Valoración de Enfermería/Obstetricia"
         Height          =   3180
         Left            =   90
         TabIndex        =   25
         Top             =   2535
         Width           =   9525
         Begin VB.CheckBox chkValorizacion 
            Caption         =   "Ingresar valoración del profesional de salud"
            Height          =   210
            Left            =   0
            TabIndex        =   73
            Top             =   0
            Value           =   1  'Checked
            Width           =   3855
         End
         Begin TabDlg.SSTab TabsDominios 
            Height          =   2775
            Left            =   120
            TabIndex        =   37
            Top             =   330
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   4895
            _Version        =   393216
            Tabs            =   14
            TabsPerRow      =   14
            TabHeight       =   520
            TabCaption(0)   =   "P1"
            TabPicture(0)   =   "VistasEnfermeras.frx":60F3
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "frameTab(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Tab 1"
            TabPicture(1)   =   "VistasEnfermeras.frx":610F
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "frameTab(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Tab 2"
            TabPicture(2)   =   "VistasEnfermeras.frx":612B
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "frameTab(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Tab 3"
            TabPicture(3)   =   "VistasEnfermeras.frx":6147
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "frameTab(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Tab 4"
            TabPicture(4)   =   "VistasEnfermeras.frx":6163
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "frameTab(4)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Tab 5"
            TabPicture(5)   =   "VistasEnfermeras.frx":617F
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "frameTab(5)"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "Tab 6"
            TabPicture(6)   =   "VistasEnfermeras.frx":619B
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "frameTab(6)"
            Tab(6).ControlCount=   1
            TabCaption(7)   =   "Tab 7"
            TabPicture(7)   =   "VistasEnfermeras.frx":61B7
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "frameTab(7)"
            Tab(7).ControlCount=   1
            TabCaption(8)   =   "Tab 8"
            TabPicture(8)   =   "VistasEnfermeras.frx":61D3
            Tab(8).ControlEnabled=   0   'False
            Tab(8).Control(0)=   "frameTab(8)"
            Tab(8).ControlCount=   1
            TabCaption(9)   =   "Tab 9"
            TabPicture(9)   =   "VistasEnfermeras.frx":61EF
            Tab(9).ControlEnabled=   0   'False
            Tab(9).Control(0)=   "frameTab(9)"
            Tab(9).ControlCount=   1
            TabCaption(10)  =   "Tab 10"
            TabPicture(10)  =   "VistasEnfermeras.frx":620B
            Tab(10).ControlEnabled=   0   'False
            Tab(10).Control(0)=   "frameTab(10)"
            Tab(10).ControlCount=   1
            TabCaption(11)  =   "Tab 11"
            TabPicture(11)  =   "VistasEnfermeras.frx":6227
            Tab(11).ControlEnabled=   0   'False
            Tab(11).Control(0)=   "frameTab(11)"
            Tab(11).ControlCount=   1
            TabCaption(12)  =   "Tab 12"
            TabPicture(12)  =   "VistasEnfermeras.frx":6243
            Tab(12).ControlEnabled=   0   'False
            Tab(12).Control(0)=   "frameTab(12)"
            Tab(12).ControlCount=   1
            TabCaption(13)  =   "Tab 13"
            TabPicture(13)  =   "VistasEnfermeras.frx":625F
            Tab(13).ControlEnabled=   0   'False
            Tab(13).Control(0)=   "frameTab(13)"
            Tab(13).ControlCount=   1
            Begin VB.Frame frameTab 
               Height          =   2295
               Index           =   0
               Left            =   0
               TabIndex        =   56
               Top             =   360
               Width           =   9100
            End
            Begin VB.Frame frameTab 
               Height          =   2295
               Index           =   13
               Left            =   -74880
               TabIndex        =   50
               Top             =   240
               Width           =   9100
            End
            Begin VB.Frame frameTab 
               Height          =   2295
               Index           =   12
               Left            =   -74880
               TabIndex        =   49
               Top             =   240
               Width           =   9100
            End
            Begin VB.Frame frameTab 
               Height          =   2295
               Index           =   11
               Left            =   -74880
               TabIndex        =   48
               Top             =   240
               Width           =   9100
            End
            Begin VB.Frame frameTab 
               Height          =   2295
               Index           =   10
               Left            =   -74880
               TabIndex        =   47
               Top             =   240
               Width           =   9100
            End
            Begin VB.Frame frameTab 
               Height          =   2295
               Index           =   9
               Left            =   -74880
               TabIndex        =   46
               Top             =   240
               Width           =   9100
            End
            Begin VB.Frame frameTab 
               Height          =   2295
               Index           =   8
               Left            =   -74880
               TabIndex        =   45
               Top             =   240
               Width           =   9100
            End
            Begin VB.Frame frameTab 
               Height          =   2295
               Index           =   7
               Left            =   -74880
               TabIndex        =   44
               Top             =   240
               Width           =   9100
            End
            Begin VB.Frame frameTab 
               Height          =   2295
               Index           =   6
               Left            =   -74880
               TabIndex        =   43
               Top             =   240
               Width           =   9100
            End
            Begin VB.Frame frameTab 
               Height          =   2295
               Index           =   5
               Left            =   -74880
               TabIndex        =   42
               Top             =   240
               Width           =   9100
            End
            Begin VB.Frame frameTab 
               Height          =   2295
               Index           =   4
               Left            =   -74880
               TabIndex        =   41
               Top             =   240
               Width           =   9100
            End
            Begin VB.Frame frameTab 
               Height          =   2295
               Index           =   3
               Left            =   -74880
               TabIndex        =   40
               Top             =   240
               Width           =   9100
            End
            Begin VB.Frame frameTab 
               Height          =   2295
               Index           =   2
               Left            =   -75000
               TabIndex        =   39
               Top             =   240
               Width           =   9100
            End
            Begin VB.Frame frameTab 
               Height          =   2295
               Index           =   1
               Left            =   -74950
               TabIndex        =   38
               Top             =   310
               Width           =   9100
            End
         End
      End
      Begin VB.Frame FrameGrafico3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   9720
         TabIndex        =   21
         Top             =   7470
         Width           =   3810
         Begin OWC11.ChartSpace CSGrafico3 
            Height          =   1950
            Left            =   0
            OleObjectBlob   =   "VistasEnfermeras.frx":627B
            TabIndex        =   22
            Top             =   0
            Width           =   3780
         End
      End
      Begin VB.Frame FrameGrafico2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   9720
         TabIndex        =   19
         Top             =   5420
         Width           =   3810
         Begin OWC11.ChartSpace CSGrafico2 
            Height          =   1950
            Left            =   0
            OleObjectBlob   =   "VistasEnfermeras.frx":6E6F
            TabIndex        =   20
            Top             =   0
            Width           =   3780
         End
      End
      Begin VB.Frame FrameGrafico1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   9720
         TabIndex        =   17
         Top             =   3360
         Width           =   3810
         Begin OWC11.ChartSpace CSGrafico1 
            Height          =   1950
            Left            =   0
            OleObjectBlob   =   "VistasEnfermeras.frx":7A63
            TabIndex        =   18
            Top             =   0
            Width           =   3780
         End
      End
      Begin MSMask.MaskEdBox TxtVariableFormato 
         Height          =   360
         Index           =   0
         Left            =   6960
         TabIndex        =   54
         Tag             =   "__/__/____ __:__"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1900
         _ExtentX        =   3360
         _ExtentY        =   635
         _Version        =   393216
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox txtFechaControl 
         Height          =   360
         Left            =   1320
         TabIndex        =   4
         Tag             =   "__/__/____ __:__"
         Top             =   2145
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl2Puntos 
         Caption         =   ":"
         Height          =   255
         Index           =   0
         Left            =   9600
         TabIndex        =   67
         Top             =   3000
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Servicio"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   10560
         TabIndex        =   63
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Historia"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   6405
         TabIndex        =   62
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Cuenta"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   195
         TabIndex        =   61
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Cama"
         Height          =   210
         Left            =   8745
         TabIndex        =   60
         Top             =   315
         Width           =   735
      End
      Begin VB.Label lblFechaVisita 
         Caption         =   "Fecha de visita"
         Height          =   255
         Left            =   105
         TabIndex        =   59
         Top             =   2205
         Width           =   1275
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         ForeColor       =   &H8000000E&
         Height          =   360
         Left            =   9720
         TabIndex        =   58
         Top             =   2980
         Width           =   4215
      End
      Begin VB.Label lblVariable 
         Caption         =   "Fecha de visita"
         Height          =   255
         Index           =   0
         Left            =   7200
         TabIndex        =   55
         Top             =   3000
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Profesional de Salud"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4590
         TabIndex        =   35
         Top             =   2205
         Width           =   1635
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   0
      TabIndex        =   0
      Top             =   9480
      Width           =   16365
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "VistasEnfermeras.frx":8657
         DownPicture     =   "VistasEnfermeras.frx":8B1B
         Height          =   700
         Left            =   8670
         Picture         =   "VistasEnfermeras.frx":9007
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "VistasEnfermeras.frx":94F3
         DownPicture     =   "VistasEnfermeras.frx":9953
         Height          =   700
         Left            =   7125
         Picture         =   "VistasEnfermeras.frx":9DC8
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   180
         Width           =   1365
      End
   End
End
Attribute VB_Name = "VisitasEnfermeras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento del Módulo de Enfermería
'        Programado por: Cachay F
'        Fecha: Enero 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim ml_IdUsuario As Long
Dim ml_TipoServicio As sghTipoServicio
Dim ml_TotalDeGraficos As Integer
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim ml_EstadoCuenta As Long
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_DoAtencionDatosAdicionales As New DoAtencionDatosAdicionales
Dim oDoPacienteDatosAdd As New DoPacienteDatosAdd
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ldFechaEgresoMedicoAnterior As Date   'cuando se "modifique", generar "consumo por dias estancia"
Dim mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String
Dim oRsEnfermeriaCatalogoDominios As New Recordset
Dim oRsEnfermeriaCatalogoVariables As New Recordset
Dim oRsEnfermeriaHistoricoVisitas As New Recordset
Dim oRsEnfermeriaCatalogoGraficos As New Recordset
Dim oDOEnfermeria_Visitas As New DOEnfermeria_Visitas
Dim oRsEnfermeriaVariables As New Recordset
Dim oRsEnfermeriaValoresCombo As New Recordset
Dim oRsEnfermeriaMedicamentos As New Recordset
Dim oRsEnfermeriaTratamientoDosis As New Recordset
'------------------------------------------------------------------------------------
'                               VARIABLES CUENTAS DE ATENCION
'------------------------------------------------------------------------------------
Dim mo_lbNuevoMovimiento As Boolean
Dim ml_idCuentaAtencion As Long
Dim mo_Atenciones As New DOAtencion
Dim ml_idAtencion As Long
Dim ml_idPaciente As Long
Dim mo_Pacientes  As New doPaciente
Dim mo_Especialidad As New DOEspecialidades
Dim ml_IdEspecialidad As Long
Dim ml_IdVisita As Integer
Dim ml_IdDiaVisita As Integer
Dim ml_IdUltimaVisitaRegistrada As Integer
Dim ml_IdServicio As Integer
Dim ml_idCama As Integer
Dim ml_IdEmpleado As Integer
Dim mb_EsNuevaVisita As Boolean
Dim mb_CargaUnaSolaVez As Boolean
Dim mb_ValidaCheckValorizacion As Boolean
'------------------------------------------------------------------------------------
'                               LISTAS DESPLEGABLES PARA LOS COMBOS
'------------------------------------------------------------------------------------
Dim mo_cmbVariable(1 To 150) As New ListaDespleglable
Dim mo_listprueba As New ListaDespleglable
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'                               GRAFICOS
'------------------------------------------------------------------------------------
Dim xValues As Variant, yValues As Variant
Dim owcChart As OWC11.ChChart
Dim owcSeries As OWC11.ChSeries
Dim lnNroPuntosGraficos As Integer
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------


Property Let lbNuevoMovimiento(lValue As Boolean)
   mo_lbNuevoMovimiento = lValue
End Property
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
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
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let idPaciente(lValue As Long)
   ml_idPaciente = lValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property
Property Let TipoServicio(sValue As sghTipoServicio)
   ml_TipoServicio = sValue
End Property
Property Get TipoServicio() As sghTipoServicio
   TipoServicio = ml_TipoServicio
End Property
Property Let IdEspecialidad(lValue As Long)
   ml_IdEspecialidad = lValue
End Property
Property Get IdEspecialidad() As Long
   IdEspecialidad = ml_IdEspecialidad
End Property
Property Let CargaUnaSolaVez(lValue As Boolean)
   mb_CargaUnaSolaVez = lValue
End Property
Property Get CargaUnaSolaVez() As Boolean
   CargaUnaSolaVez = mb_CargaUnaSolaVez
End Property

Private Sub btnAgregarVisita_Click()
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    If MsgBox("¿Desea registrar una nueva visita?", vbYesNo, Me.Caption) = vbYes Then
        If mb_EsNuevaVisita = True Then
            MsgBox "El registro de la nueva visita ya esta disponible", vbInformation, Me.Caption
        Else
            oRsEnfermeriaHistoricoVisitas.MoveFirst
            oRsEnfermeriaHistoricoVisitas.Find "IdVisita=" & CInt(ml_IdUltimaVisitaRegistrada + 1)
            If Not oRsEnfermeriaHistoricoVisitas.EOF Then
                ml_IdVisita = oRsEnfermeriaHistoricoVisitas.Fields!IdVisita
            Else
                ml_IdVisita = CInt(ml_IdUltimaVisitaRegistrada + 1)
                oRsEnfermeriaHistoricoVisitas.AddNew
                oRsEnfermeriaHistoricoVisitas.Fields!IdVisita = ml_IdVisita
                oRsEnfermeriaHistoricoVisitas.Fields!Descripcion = "Nº " & CStr(ml_IdVisita)
                oRsEnfermeriaHistoricoVisitas.Fields!FechaHoraVisita = "Nuevo"
                oRsEnfermeriaHistoricoVisitas.Update
            End If
            Me.FrameVisita.Caption = "Visita Nº " & CStr(ml_IdVisita) & " (Nuevo)"
            mo_Formulario.HabilitarDeshabilitar Me.txtFechaControl, True
            mb_EsNuevaVisita = True
            LimpiaTodosControles False
            BuscaDatosActualesPaciente oConexion
            ConfiguraListadoMedicamentos
            MuestraGraficosPorHoja
            Me.TabsDominios.Tab = 0
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub btnCpt_Click()

End Sub

Private Sub btnQuitarVisita_Click()
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    If MsgBox("¿Desea eliminar el registro de la visita?", vbYesNo, Me.Caption) = vbYes Then
        If mb_EsNuevaVisita = True Then
            If ml_IdUltimaVisitaRegistrada = 0 Then
                MsgBox "No se puede eliminar el primer registro", vbInformation, Me.Caption
            Else
                oRsEnfermeriaHistoricoVisitas.MoveFirst
                oRsEnfermeriaHistoricoVisitas.Find "IdVisita=" & ml_IdVisita
                If Not oRsEnfermeriaHistoricoVisitas.EOF Then
                    oRsEnfermeriaHistoricoVisitas.Delete
                    oRsEnfermeriaHistoricoVisitas.Update
                End If
                oRsEnfermeriaHistoricoVisitas.MoveFirst
                oRsEnfermeriaHistoricoVisitas.Find "IdVisita=" & ml_IdUltimaVisitaRegistrada
                ml_IdVisita = ml_IdUltimaVisitaRegistrada
                Me.FrameVisita.Caption = "Visita Nº " & CStr(ml_IdVisita)
                mo_Formulario.HabilitarDeshabilitar Me.txtFechaControl, False
                mb_EsNuevaVisita = False
                LimpiaTodosControles False
                CargaDatosVisita oConexion
                ConfiguraListadoMedicamentos
                MuestraGraficosPorHoja
            End If
        Else
            If ml_IdVisita = ml_IdUltimaVisitaRegistrada Then
                oRsEnfermeriaHistoricoVisitas.MoveFirst
                oRsEnfermeriaHistoricoVisitas.Find "IdVisita=" & CInt(ml_IdUltimaVisitaRegistrada + 1)
                If Not oRsEnfermeriaHistoricoVisitas.EOF Then
                    MsgBox "Primero debe eliminar el nuevo registro", vbInformation, Me.Caption
                Else
                    oRsEnfermeriaHistoricoVisitas.MoveFirst
                    oRsEnfermeriaHistoricoVisitas.Find "IdVisita=" & CInt(ml_IdVisita)
                    mo_Formulario.HabilitarDeshabilitar Me.txtFechaControl, False
                    If Not oRsEnfermeriaHistoricoVisitas.EOF Then
                        If EliminarDatos() Then
                            MsgBox "Los datos de la visita Nº " & CStr(ml_IdVisita) & " se eliminaron correctamente", vbInformation, Me.Caption
                            Me.Visible = False
                        Else
                            MsgBox "No se pudo eliminar los datos de la visita" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                        End If
                    End If
                End If
            Else
                MsgBox "Si desea eliminar, debe empezar desde la última visita", vbInformation, Me.Caption
            End If
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub chkValorizacion_Click()
    If chkValorizacion.Value Then
        HabilitarDehabilitarVariables True
    Else
        If mb_ValidaCheckValorizacion = True Then
            If MsgBox("¿Esta seguro de no ingresar la valorización?", vbYesNo, Me.Caption) = vbYes Then
                LimpiarDatosVariables
                HabilitarDehabilitarVariables False
            Else
                chkValorizacion.Value = 1
                HabilitarDehabilitarVariables True
            End If
       End If
    End If
End Sub

Sub HabilitarDehabilitarVariables(Habilitar As Boolean)
    oRsEnfermeriaCatalogoVariables.MoveFirst
    Do While Not oRsEnfermeriaCatalogoVariables.EOF
        Select Case oRsEnfermeriaCatalogoVariables.Fields!tipo
            Case "ValorEntero"
                mo_Formulario.HabilitarDeshabilitar txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable), Habilitar
            Case "ValorTexto"
                mo_Formulario.HabilitarDeshabilitar TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable), Habilitar
            Case "ValorDouble"
                mo_Formulario.HabilitarDeshabilitar TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable), Habilitar
            Case "ValorCombo"
                mo_Formulario.HabilitarDeshabilitar Me.cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable), Habilitar
            Case "ValorMultiple"
                If Habilitar Then
                    Me.LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).BackColor = &HFFFFFF
                    Me.LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).ForeColor = &H0&
'                    Me.LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).BackColor = &H8000000F
                    Me.LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Enabled = True
                Else
                    Me.LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).BackColor = &HF9EADF
                    Me.LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).ForeColor = &H808080
                    Me.LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Enabled = False
                End If
'                mo_Formulario.HabilitarDeshabilitar Me.LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable), Habilitar
            Case "ValorCheck"
                mo_Formulario.HabilitarDeshabilitar Me.chkVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable), Habilitar
        End Select
        oRsEnfermeriaCatalogoVariables.MoveNext
    Loop
End Sub

Sub LimpiarDatosVariables()
    oRsEnfermeriaCatalogoVariables.MoveFirst
    Do While Not oRsEnfermeriaCatalogoVariables.EOF
        Select Case oRsEnfermeriaCatalogoVariables.Fields!tipo
            Case "ValorEntero"
                txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text = ""
            Case "ValorTexto"
                If oRsEnfermeriaCatalogoVariables.Fields!TieneFormatoMask Then
                    TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text = Replace(oRsEnfermeriaCatalogoVariables.Fields!FormatoMask, "#", "_")
                Else
                    TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text = ""
                End If
            Case "ValorDouble"
                TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text = ""
            Case "ValorCombo"
                mo_cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).BoundText = 1
            Case "ValorMultiple"
                Dim lnItem As Integer
                For lnItem = 0 To LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).ListCount - 1
                    LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Selected(lnItem) = False
                Next
            Case "ValorCheck"
                chkVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Value = 0
        End Select
        oRsEnfermeriaCatalogoVariables.MoveNext
    Loop
End Sub

Private Sub chkVariable_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, chkVariable(Index)
End Sub

Private Sub cmbVariable_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbVariable(Index)
End Sub

Private Sub CSGrafico1_DblClick()
    VistaPreviaGrafico Val(CSGrafico1.Tag), CSGrafico1.ToolTipText
End Sub

Private Sub CSGrafico2_DblClick()
    VistaPreviaGrafico Val(CSGrafico2.Tag), CSGrafico2.ToolTipText
End Sub

Private Sub CSGrafico3_DblClick()
    VistaPreviaGrafico Val(CSGrafico3.Tag), CSGrafico3.ToolTipText
End Sub

Public Sub VistaPreviaGrafico(ByVal liIdVariable As Integer, ByVal lcTitulo As String)
    Dim oVisitEnferGraficos As New VisitEnferGraficos
    oVisitEnferGraficos.idCuentaAtencion = ml_idCuentaAtencion
    oVisitEnferGraficos.IdVisita = ml_IdVisita
    oVisitEnferGraficos.IdVariable = liIdVariable
    oVisitEnferGraficos.TituloGrafico = lcTitulo
    oVisitEnferGraficos.EsNuevaVisita = mb_EsNuevaVisita
    oRsEnfermeriaCatalogoVariables.MoveFirst
    Do While Not oRsEnfermeriaCatalogoVariables.EOF
        If oRsEnfermeriaCatalogoVariables.Fields!IdVariable = liIdVariable Then
            If oRsEnfermeriaCatalogoVariables.Fields!tipo = "ValorEntero" Then
                oVisitEnferGraficos.TextoVariable = Trim(Me.txtVariable(liIdVariable).Text)
            Else
                oVisitEnferGraficos.TextoVariable = Trim(Me.TxtVariableFormato(liIdVariable).Text)
            End If
        End If
        oRsEnfermeriaCatalogoVariables.MoveNext
    Loop
'    oVisitEnferGraficos.TextoVariable = Trim(Me.txtVariable(liIdVariable).Text)
    oVisitEnferGraficos.MostrarFormulario
    mb_CargaUnaSolaVez = False
End Sub

Private Sub grdHistoricoVisitas_Click()
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    If grdHistoricoVisitas.Selected.Rows.Count > 0 Then
        ml_IdVisita = grdHistoricoVisitas.Selected.Rows(0).Cells("IdVisita").Value
        If grdHistoricoVisitas.Selected.Rows(0).Cells("FechaHoraVisita").Value = "Nuevo" Then
            mb_EsNuevaVisita = True
            Me.FrameVisita.Caption = "Visita " & grdHistoricoVisitas.Selected.Rows(0).Cells("Descripcion").Value & " (Nuevo)"
            LimpiaTodosControles False
            BuscaDatosActualesPaciente oConexion
            mo_Formulario.HabilitarDeshabilitar Me.txtFechaControl, True
        Else
            mb_EsNuevaVisita = False
            Me.FrameVisita.Caption = "Visita " & grdHistoricoVisitas.Selected.Rows(0).Cells("Descripcion").Value
            LimpiaTodosControles False
            CargaDatosVisita oConexion
            mo_Formulario.HabilitarDeshabilitar Me.txtFechaControl, False
        End If
        ConfiguraListadoMedicamentos
        MuestraGraficosPorHoja
    End If
    
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub Form_Activate()
    If mb_CargaUnaSolaVez = False Then Exit Sub
    TituloDeForm
    CargarDatosAlosControles
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub

Sub Form_Load()
    If mb_CargaUnaSolaVez = False Then Exit Sub
    ConfigurarPantallaVisitaEnfermera
    TituloDeForm
    CargarDatosAlosControles
    
    pcTratamiento.Top = 0
    pcTratamiento.Left = 0
    
    With vsTratamiento 'Si vas a utilizar el Vertical
        .Min = 0
        .SmallChange = 90
        .LargeChange = 300
        .Top = 0
        .ZOrder 0
    End With
End Sub



Private Sub txtDosisProrenata_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtFechaControl_KeyUp(KeyCode As Integer, Shift As Integer)
    ConfiguraListadoMedicamentos
End Sub

Private Sub vsTratamiento_Change()
    pcTratamiento.Top = -vsTratamiento.Value
End Sub

Sub ConfigurarPantallaVisitaEnfermera()
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    'Bloquea datos paciente
    mo_Formulario.HabilitarDeshabilitar Me.txtNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNroCuenta, False
    mo_Formulario.HabilitarDeshabilitar Me.txtPrimerNombre, False
    mo_Formulario.HabilitarDeshabilitar Me.txtServicio, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNroCama, False
    mo_Formulario.HabilitarDeshabilitar Me.txtNombreEmpleado, False
'    mo_Formulario.HabilitarDeshabilitar Me.txtantecedAlergico, False
'    mo_Formulario.HabilitarDeshabilitar Me.txtAntecedentes, False
'    mo_Formulario.HabilitarDeshabilitar Me.txtantecedFamiliar, False
'    mo_Formulario.HabilitarDeshabilitar Me.txtantecedObstetrico, False
'    mo_Formulario.HabilitarDeshabilitar Me.txtantecedPatologico, False
'    mo_Formulario.HabilitarDeshabilitar Me.txtantecedQuirurgico, False
'
    CreaTemporalCatalogoDominios 'Crea temporal catalogo de dominios
    CreaTemporalCatalogoVariables 'Crea temporal catalogo de variables
    CreaTemporalCatalogoGraficos 'Crea temporal catalogo de graficos
    CreaTemporalMedicamentoRecetados ' Crea Temporal CatalogoMedicamentos
    
    ConfiguraDominios oConexion 'Configura dominios
    ConfiguraVariables oConexion 'Configura variables
    
    CreaTemporalHistoricoVisitas 'Inicializa CatalogoVisitas
    CalculaCantidadGraficos oConexion 'Calcula cantidad graficos

    oConexion.Close
    Set oConexion = Nothing
End Sub

Sub CreaTemporalHistoricoVisitas()
    If oRsEnfermeriaHistoricoVisitas.State = 1 Then
       Set oRsEnfermeriaHistoricoVisitas = Nothing
    End If
    With oRsEnfermeriaHistoricoVisitas
          .Fields.Append "IdVisita", adInteger, 4, adFldIsNullable
          .Fields.Append "Descripcion", adVarChar, 100, adFldIsNullable
          .Fields.Append "FechaHoraVisita", adVarChar, 100, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdHistoricoVisitas.DataSource = oRsEnfermeriaHistoricoVisitas
    mo_Apariencia.ConfigurarFilasBiColores grdHistoricoVisitas, sighEntidades.GrillaConFilasBicolor
End Sub

Sub CreaTemporalCatalogoDominios()
    If oRsEnfermeriaCatalogoDominios.State = 1 Then
       Set oRsEnfermeriaCatalogoDominios = Nothing
    End If
    With oRsEnfermeriaCatalogoDominios
        .Fields.Append "IdDominio", adInteger
        .Fields.Append "CodDominio", adVarChar, 255, adFldIsNullable
        .Fields.Append "DominioTexto", adVarChar, 255, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Sub CreaTemporalCatalogoVariables()
    If oRsEnfermeriaCatalogoVariables.State = 1 Then
       Set oRsEnfermeriaCatalogoVariables = Nothing
    End If
    With oRsEnfermeriaCatalogoVariables
          .Fields.Append "IdVariable", adInteger
          .Fields.Append "IdDominio", adInteger
          .Fields.Append "OrdernDominio", adInteger
          .Fields.Append "Texto", adVarChar, 255, adFldIsNullable
          .Fields.Append "Tipo", adVarChar, 255, adFldIsNullable
          .Fields.Append "Ancho", adInteger
          .Fields.Append "EsDatoObligatorio", adBoolean
          .Fields.Append "TextoToolTip", adVarChar, 255, adFldIsNullable
          .Fields.Append "EsDatoGrafico", adBoolean
          .Fields.Append "TieneFormatoMask", adBoolean
          .Fields.Append "FormatoMask", adVarChar, 255, adFldIsNullable
          .Fields.Append "TieneRango", adBoolean
          .Fields.Append "RangoInicial", adInteger
          .Fields.Append "RangoFinal", adInteger
          .Fields.Append "PosicionFila", adInteger
          .Fields.Append "PosicionColumna", adInteger
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

Sub CreaTemporalCatalogoGraficos()
    If oRsEnfermeriaCatalogoGraficos.State = 1 Then
       Set oRsEnfermeriaCatalogoGraficos = Nothing
    End If
    With oRsEnfermeriaCatalogoGraficos
        .Fields.Append "IdVariable", adInteger
        .Fields.Append "IdDominio", adInteger
        .Fields.Append "Texto", adVarChar, 255, adFldIsNullable
        .Fields.Append "NroHoja", adInteger
        .Fields.Append "NroGraficoHoja", adInteger
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Sub CreaTemporalMedicamentoRecetados()
    If oRsEnfermeriaMedicamentos.State = 1 Then
       Set oRsEnfermeriaMedicamentos = Nothing
    End If
    With oRsEnfermeriaMedicamentos
          .Fields.Append "IdControles", adInteger
          .Fields.Append "idReceta", adInteger
          .Fields.Append "idMedicoReceta", adInteger
          .Fields.Append "idEstadoDetalle", adInteger
          .Fields.Append "Estado", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdEmpleado", adInteger
          .Fields.Append "ApellidoPaterno", adVarChar, 255, adFldIsNullable
          .Fields.Append "ApellidoMaterno", adVarChar, 255, adFldIsNullable
          .Fields.Append "Nombres", adVarChar, 255, adFldIsNullable
          .Fields.Append "idItem", adInteger
          .Fields.Append "Codigo", adVarChar, 255, adFldIsNullable
          .Fields.Append "Medicamento", adVarChar, 255, adFldIsNullable
          .Fields.Append "CantidadPedida", adInteger
          .Fields.Append "CantidadDespachada", adInteger
          .Fields.Append "idDosisRecetada", adInteger
          .Fields.Append "MotivoAnulacionMedico", adVarChar, 255, adFldIsNullable
          .Fields.Append "Dosis", adInteger, 0, adFldIsNullable
          .Fields.Append "DatoProrenata", adInteger, 0, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

Sub ConfiguraDominios(oConexion As Connection)
   Dim ml_CuentaTabs As Integer
   Dim oRsTmp1 As New Recordset
   'Configura dominios
    
    Set oRsTmp1 = mo_AdminAdmision.Enfermeria_CatalogoDominiosSeleccionar(oConexion)
    If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
              oRsEnfermeriaCatalogoDominios.AddNew
              oRsEnfermeriaCatalogoDominios.Fields!IdDominio = oRsTmp1.Fields!IdDominio
              oRsEnfermeriaCatalogoDominios.Fields!CodDominio = oRsTmp1.Fields!CodDominio
              oRsEnfermeriaCatalogoDominios.Fields!DominioTexto = oRsTmp1.Fields!DominioTexto
              oRsEnfermeriaCatalogoDominios.Update
              oRsTmp1.MoveNext
        Loop
        oRsEnfermeriaCatalogoDominios.MoveFirst
        For ml_CuentaTabs = 0 To 13
            Me.frameTab(ml_CuentaTabs).Left = 50
            Me.frameTab(ml_CuentaTabs).Top = 310
            Me.TabsDominios.TabVisible(ml_CuentaTabs) = False
        Next
        Do While Not oRsEnfermeriaCatalogoDominios.EOF
            TabsDominios.TabVisible(oRsEnfermeriaCatalogoDominios.Fields!IdDominio) = True
            Me.TabsDominios.Tab = oRsEnfermeriaCatalogoDominios.Fields!IdDominio
            TabsDominios.Caption = oRsEnfermeriaCatalogoDominios.Fields!CodDominio
            Me.frameTab(oRsEnfermeriaCatalogoDominios.Fields!IdDominio).Caption = oRsEnfermeriaCatalogoDominios.Fields!DominioTexto
            oRsEnfermeriaCatalogoDominios.MoveNext
        Loop
        TabsDominios.TabsPerRow = oRsEnfermeriaCatalogoDominios.RecordCount
        TabsDominios.Width = 9210
        TabsDominios.Height = 2775
    End If
    If oRsTmp1.State = 1 Then oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Sub

Sub LimpiaListadoMedicamentos()
        If oRsEnfermeriaMedicamentos.RecordCount > 0 Then
            oRsEnfermeriaMedicamentos.MoveFirst
            Do While Not oRsEnfermeriaMedicamentos.EOF
'                Unload lblObsRecetado(oRsEnfermeriaMedicamentos.Fields!IdControles)
                Unload lblMedicamento(oRsEnfermeriaMedicamentos.Fields!IdControles)
                Unload lblObsMedicamento(oRsEnfermeriaMedicamentos.Fields!IdControles)
                If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada = 0 Then
                    Unload lblProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    Unload txtDosisProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles)
                Else
                    If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 1 Then
                        Unload lblDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles)
                        Unload chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    End If
                    If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 2 Then
                        Unload lblDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles)
                        Unload chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    End If
                    If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 3 Then
                        Unload lblDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles)
                        Unload chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    End If
                    If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 4 Then
                        Unload lblDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles)
                        Unload chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    End If
                    If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 5 Then
                        Unload lblDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles)
                        Unload chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    End If
                    If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 6 Then
                        Unload lblDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles)
                        Unload chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    End If
                End If
                oRsEnfermeriaMedicamentos.MoveNext
            Loop
        End If
    
    'Limpia Medicamentos visita
    With oRsEnfermeriaMedicamentos
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              .Delete
              .Update
              .MoveNext
           Loop
        End If
    End With
End Sub

Sub ConfiguraListadoMedicamentos()
    Dim ml_CuentaTabs As Integer
    Dim oRsTmp1 As New Recordset
    Dim lnIdControles As Integer
    Dim lnTopMedicamento As Long
    Dim lbDosis1Dada As Boolean
    Dim lbDosis2Dada As Boolean
    Dim lbDosis3Dada As Boolean
    Dim lbDosis4Dada As Boolean
    Dim lbDosis5Dada As Boolean
    Dim lbDosis6Dada As Boolean
    Dim lbDosisVisita As Boolean
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    CalcularIdDiaVisita
    LimpiaListadoMedicamentos
    If mb_EsNuevaVisita Then
        If IsDate(Me.txtFechaControl.Text) = False Then Exit Sub
        Set oRsTmp1 = mo_AdminAdmision.Enfermeria_RecetasDespachadasPorCuenta(oConexion, ml_idCuentaAtencion, Me.txtFechaControl.Text)
    Else
        Set oRsTmp1 = mo_AdminAdmision.Enfermeria_ConsultarMedicamentosVisita(oConexion, ml_idCuentaAtencion, ml_IdVisita)
    End If
'    ml_IdDiaVisita
    If oRsTmp1.RecordCount > 0 Then
        lnIdControles = 0
        pcTratamiento.Height = 600 * oRsTmp1.RecordCount
        vsTratamiento.Max = IIf(pcTratamiento.Height > 32000, 32000, pcTratamiento.Height)  'debb2017
        
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
            oRsEnfermeriaMedicamentos.AddNew
            lnIdControles = lnIdControles + 1
            oRsEnfermeriaMedicamentos.Fields!IdControles = lnIdControles
            oRsEnfermeriaMedicamentos.Fields!idReceta = oRsTmp1.Fields!idReceta
            oRsEnfermeriaMedicamentos.Fields!idMedicoREceta = oRsTmp1.Fields!idMedicoREceta
            oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = oRsTmp1.Fields!idEstadoDetalle
            oRsEnfermeriaMedicamentos.Fields!estado = oRsTmp1.Fields!estado
            oRsEnfermeriaMedicamentos.Fields!IdEmpleado = oRsTmp1.Fields!IdEmpleado
            oRsEnfermeriaMedicamentos.Fields!ApellidoPaterno = oRsTmp1.Fields!ApellidoPaterno
            oRsEnfermeriaMedicamentos.Fields!ApellidoMaterno = oRsTmp1.Fields!ApellidoMaterno
            oRsEnfermeriaMedicamentos.Fields!Nombres = oRsTmp1.Fields!Nombres
            oRsEnfermeriaMedicamentos.Fields!idItem = oRsTmp1.Fields!idItem
            oRsEnfermeriaMedicamentos.Fields!Codigo = oRsTmp1.Fields!Codigo
            oRsEnfermeriaMedicamentos.Fields!Medicamento = oRsTmp1.Fields!Medicamento
            oRsEnfermeriaMedicamentos.Fields!CantidadPedida = oRsTmp1.Fields!CantidadPedida
            oRsEnfermeriaMedicamentos.Fields!CantidadDespachada = IIf(IsNull(oRsTmp1.Fields!CantidadDespachada), 0, oRsTmp1.Fields!CantidadDespachada)
            oRsEnfermeriaMedicamentos.Fields!idDosisRecetada = oRsTmp1.Fields!idDosisRecetada
            oRsEnfermeriaMedicamentos.Fields!MotivoAnulacionMedico = oRsTmp1.Fields!MotivoAnulacionMedico
            oRsEnfermeriaMedicamentos.Fields!Dosis = IIf(oRsTmp1.Fields!Dosis = "", Null, oRsTmp1.Fields!Dosis)
            oRsEnfermeriaMedicamentos.Fields!DatoProrenata = IIf(oRsTmp1.Fields!DatoProrenata = "", Null, oRsTmp1.Fields!DatoProrenata)
            oRsEnfermeriaMedicamentos.Update
            oRsTmp1.MoveNext
        Loop
        oRsEnfermeriaMedicamentos.MoveFirst
        Do While Not oRsEnfermeriaMedicamentos.EOF
            lnTopMedicamento = 300 + 330 * (CInt(IIf(oRsEnfermeriaMedicamentos.Fields!IdControles > 98, 98, oRsEnfermeriaMedicamentos.Fields!IdControles)) - 1) 'debb2017
            Load lblMedicamento(oRsEnfermeriaMedicamentos.Fields!IdControles)
            lblMedicamento(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
            lblMedicamento(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 10
            lblMedicamento(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
            lblMedicamento(oRsEnfermeriaMedicamentos.Fields!IdControles).Caption = oRsEnfermeriaMedicamentos.Fields!Medicamento
            lblMedicamento(oRsEnfermeriaMedicamentos.Fields!IdControles).ToolTipText = "Recetado por: " & oRsEnfermeriaMedicamentos.Fields!Nombres & " " & oRsEnfermeriaMedicamentos.Fields!ApellidoPaterno & " " & oRsEnfermeriaMedicamentos.Fields!ApellidoMaterno
            
            Load lblObsMedicamento(oRsEnfermeriaMedicamentos.Fields!IdControles)
            lblObsMedicamento(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
            lblObsMedicamento(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 7330
            lblObsMedicamento(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
            lblObsMedicamento(oRsEnfermeriaMedicamentos.Fields!IdControles).Caption = oRsEnfermeriaMedicamentos.Fields!estado
            
            If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada = 0 Then
                Load lblProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles)
                lblProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
                lblProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 3720
                lblProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
                
                Load txtDosisProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles)
                txtDosisProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
                txtDosisProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 4680
                txtDosisProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
                
                If Not mb_EsNuevaVisita Then
                    txtDosisProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles).Text = IIf(IsNull(oRsEnfermeriaMedicamentos.Fields!DatoProrenata), "", oRsEnfermeriaMedicamentos.Fields!DatoProrenata)
                End If
                
                If oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = 4 Or oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = 0 Then
                    mo_Formulario.HabilitarDeshabilitar txtDosisProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                End If
            Else
                lbDosis1Dada = False
                lbDosis2Dada = False
                lbDosis3Dada = False
                lbDosis4Dada = False
                lbDosis5Dada = False
                lbDosis6Dada = False
                lbDosisVisita = False
                If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 1 Then
                    Load lblDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    lblDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
                    lblDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 4000
                    lblDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
                    
                    Load chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
                    chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 3720
                    chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
                    mo_Formulario.HabilitarDeshabilitar chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                    Set oRsTmp1 = mo_AdminAdmision.Enfermeria_ConsultarMedicamentoDosisDia(oConexion, ml_idCuentaAtencion, oRsEnfermeriaMedicamentos.Fields!idReceta, oRsEnfermeriaMedicamentos.Fields!idItem, ml_IdDiaVisita, 1)
                    If oRsTmp1.RecordCount > 0 Then
                        chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles).Value = 1
                        lblDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles).ToolTipText = "Dosificado el :" & oRsTmp1.Fields!FechaHoraVisita & " por el Prof. de Salud:  " & oRsTmp1.Fields!Nombres & " " & oRsTmp1.Fields!ApellidoPaterno & " " & oRsTmp1.Fields!ApellidoMaterno
                        chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles).ToolTipText = "Dosificado el :" & oRsTmp1.Fields!FechaHoraVisita & " por el Prof. de Salud:  " & oRsTmp1.Fields!Nombres & " " & oRsTmp1.Fields!ApellidoPaterno & " " & oRsTmp1.Fields!ApellidoMaterno
                        If ml_IdVisita = oRsTmp1.Fields!IdVisita Then
                            mo_Formulario.HabilitarDeshabilitar chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles), True
                            lbDosisVisita = True
                        End If
                        lbDosis1Dada = True
                    Else
                        mo_Formulario.HabilitarDeshabilitar chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles), True
                        lbDosis1Dada = False
                    End If
                    
                    If oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = 4 Or oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = 0 Then
                        mo_Formulario.HabilitarDeshabilitar chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                    End If
                End If
                If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 2 Then
                    Load lblDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    lblDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
                    lblDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 4600
                    lblDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
                    
                    Load chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
                    chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 4320
                    chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
                    If lbDosis1Dada = False Then
                        mo_Formulario.HabilitarDeshabilitar chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                        lbDosis2Dada = False
                    Else
                        Set oRsTmp1 = mo_AdminAdmision.Enfermeria_ConsultarMedicamentoDosisDia(oConexion, ml_idCuentaAtencion, oRsEnfermeriaMedicamentos.Fields!idReceta, oRsEnfermeriaMedicamentos.Fields!idItem, ml_IdDiaVisita, 2)
                        If oRsTmp1.RecordCount > 0 Then
                            chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles).Value = 1
                            mo_Formulario.HabilitarDeshabilitar chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            lblDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles).ToolTipText = "Dosificado el :" & oRsTmp1.Fields!FechaHoraVisita & " por el Prof. de Salud:  " & oRsTmp1.Fields!Nombres & " " & oRsTmp1.Fields!ApellidoPaterno & " " & oRsTmp1.Fields!ApellidoMaterno
                            chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles).ToolTipText = "Dosificado el :" & oRsTmp1.Fields!FechaHoraVisita & " por el Prof. de Salud:  " & oRsTmp1.Fields!Nombres & " " & oRsTmp1.Fields!ApellidoPaterno & " " & oRsTmp1.Fields!ApellidoMaterno
                            mo_Formulario.HabilitarDeshabilitar chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            If ml_IdVisita = oRsTmp1.Fields!IdVisita Then
                                mo_Formulario.HabilitarDeshabilitar chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles), True
                                lbDosisVisita = True
                            End If
                            lbDosis2Dada = True
                        Else
                            mo_Formulario.HabilitarDeshabilitar chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            If Not chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles).Enabled Then
                                If Not lbDosisVisita Then
                                    mo_Formulario.HabilitarDeshabilitar chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles), True
                                End If
                            End If
                            lbDosis2Dada = False
                        End If
                    End If
                    If oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = 4 Or oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = 0 Then
                        mo_Formulario.HabilitarDeshabilitar chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                    End If
                End If
                If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 3 Then
                    Load lblDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    lblDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
                    lblDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 5200
                    lblDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
                
                    Load chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
                    chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 4920
                    chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
'                    mo_Formulario.HabilitarDeshabilitar chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                    If lbDosis2Dada = False Then
                        mo_Formulario.HabilitarDeshabilitar chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                        lbDosis3Dada = False
                    Else
                        Set oRsTmp1 = mo_AdminAdmision.Enfermeria_ConsultarMedicamentoDosisDia(oConexion, ml_idCuentaAtencion, oRsEnfermeriaMedicamentos.Fields!idReceta, oRsEnfermeriaMedicamentos.Fields!idItem, ml_IdDiaVisita, 3)
                        If oRsTmp1.RecordCount > 0 Then
                            chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles).Value = 1
                            mo_Formulario.HabilitarDeshabilitar chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            lblDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles).ToolTipText = "Dosificado el :" & oRsTmp1.Fields!FechaHoraVisita & " por el Prof. de Salud:  " & oRsTmp1.Fields!Nombres & " " & oRsTmp1.Fields!ApellidoPaterno & " " & oRsTmp1.Fields!ApellidoMaterno
                            chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles).ToolTipText = "Dosificado el :" & oRsTmp1.Fields!FechaHoraVisita & " por el Prof. de Salud:  " & oRsTmp1.Fields!Nombres & " " & oRsTmp1.Fields!ApellidoPaterno & " " & oRsTmp1.Fields!ApellidoMaterno
                            mo_Formulario.HabilitarDeshabilitar chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            If ml_IdVisita = oRsTmp1.Fields!IdVisita Then
                                mo_Formulario.HabilitarDeshabilitar chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles), True
                                lbDosisVisita = True
                            End If
                            lbDosis3Dada = True
                        Else
                            mo_Formulario.HabilitarDeshabilitar chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            If Not chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles).Enabled Then
                                If Not lbDosisVisita Then
                                    mo_Formulario.HabilitarDeshabilitar chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles), True
                                End If
                            End If
                            lbDosis3Dada = False
                        End If
                    End If
                    If oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = 4 Or oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = 0 Then
                        mo_Formulario.HabilitarDeshabilitar chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                    End If
                End If
                If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 4 Then
                    Load lblDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    lblDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
                    lblDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 5800
                    lblDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
                    
                    Load chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
                    chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 5520
                    chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
                    If lbDosis3Dada = False Then
                        mo_Formulario.HabilitarDeshabilitar chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                        lbDosis4Dada = False
                    Else
                        Set oRsTmp1 = mo_AdminAdmision.Enfermeria_ConsultarMedicamentoDosisDia(oConexion, ml_idCuentaAtencion, oRsEnfermeriaMedicamentos.Fields!idReceta, oRsEnfermeriaMedicamentos.Fields!idItem, ml_IdDiaVisita, 4)
                        If oRsTmp1.RecordCount > 0 Then
                            chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles).Value = 1
                            mo_Formulario.HabilitarDeshabilitar chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            lblDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles).ToolTipText = "Dosificado el :" & oRsTmp1.Fields!FechaHoraVisita & " por el Prof. de Salud:  " & oRsTmp1.Fields!Nombres & " " & oRsTmp1.Fields!ApellidoPaterno & " " & oRsTmp1.Fields!ApellidoMaterno
                            chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles).ToolTipText = "Dosificado el :" & oRsTmp1.Fields!FechaHoraVisita & " por el Prof. de Salud:  " & oRsTmp1.Fields!Nombres & " " & oRsTmp1.Fields!ApellidoPaterno & " " & oRsTmp1.Fields!ApellidoMaterno
                            mo_Formulario.HabilitarDeshabilitar chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            If ml_IdVisita = oRsTmp1.Fields!IdVisita Then
                                mo_Formulario.HabilitarDeshabilitar chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles), True
                                lbDosisVisita = True
                            End If
                            lbDosis4Dada = True
                        Else
                            mo_Formulario.HabilitarDeshabilitar chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            If Not chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles).Enabled Then
                                If Not lbDosisVisita Then
                                    mo_Formulario.HabilitarDeshabilitar chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles), True
                                End If
                            End If
                            lbDosis4Dada = False
                        End If
                    End If
                    If oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = 4 Or oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = 0 Then
                        mo_Formulario.HabilitarDeshabilitar chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                    End If
                End If
                If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 5 Then
                    Load lblDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    lblDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
                    lblDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 6400
                    lblDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
                    
                    Load chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
                    chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 6120
                    chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
                    If lbDosis4Dada = False Then
                        mo_Formulario.HabilitarDeshabilitar chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                        lbDosis5Dada = False
                    Else
                        Set oRsTmp1 = mo_AdminAdmision.Enfermeria_ConsultarMedicamentoDosisDia(oConexion, ml_idCuentaAtencion, oRsEnfermeriaMedicamentos.Fields!idReceta, oRsEnfermeriaMedicamentos.Fields!idItem, ml_IdDiaVisita, 5)
                        If oRsTmp1.RecordCount > 0 Then
                            chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles).Value = 1
                            mo_Formulario.HabilitarDeshabilitar chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            lblDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles).ToolTipText = "Dosificado el :" & oRsTmp1.Fields!FechaHoraVisita & " por el Prof. de Salud:  " & oRsTmp1.Fields!Nombres & " " & oRsTmp1.Fields!ApellidoPaterno & " " & oRsTmp1.Fields!ApellidoMaterno
                            chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles).ToolTipText = "Dosificado el :" & oRsTmp1.Fields!FechaHoraVisita & " por el Prof. de Salud:  " & oRsTmp1.Fields!Nombres & " " & oRsTmp1.Fields!ApellidoPaterno & " " & oRsTmp1.Fields!ApellidoMaterno
                            mo_Formulario.HabilitarDeshabilitar chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            If ml_IdVisita = oRsTmp1.Fields!IdVisita Then
                                mo_Formulario.HabilitarDeshabilitar chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles), True
                                lbDosisVisita = True
                            End If
                            lbDosis5Dada = True
                        Else
                            mo_Formulario.HabilitarDeshabilitar chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            If Not chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles).Enabled Then
                                If Not lbDosisVisita Then
                                    mo_Formulario.HabilitarDeshabilitar chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles), True
                                End If
                            End If
                            lbDosis5Dada = False
                        End If
                    End If
                    If oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = 4 Or oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = 0 Then
                        mo_Formulario.HabilitarDeshabilitar chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                    End If
                End If
                If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 6 Then
                    Load lblDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    lblDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
                    lblDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 7000
                    lblDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
                    
                    Load chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles)
                    chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles).Visible = True
                    chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles).Left = 6720
                    chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles).Top = lnTopMedicamento
                    If lbDosis5Dada = False Then
                        mo_Formulario.HabilitarDeshabilitar chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                        lbDosis6Dada = False
                    Else
                        Set oRsTmp1 = mo_AdminAdmision.Enfermeria_ConsultarMedicamentoDosisDia(oConexion, ml_idCuentaAtencion, oRsEnfermeriaMedicamentos.Fields!idReceta, oRsEnfermeriaMedicamentos.Fields!idItem, ml_IdDiaVisita, 6)
                        If oRsTmp1.RecordCount > 0 Then
                            chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles).Value = 1
                            mo_Formulario.HabilitarDeshabilitar chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            lblDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles).ToolTipText = "Dosificado el :" & oRsTmp1.Fields!FechaHoraVisita & " por el Prof. de Salud:  " & oRsTmp1.Fields!Nombres & " " & oRsTmp1.Fields!ApellidoPaterno & " " & oRsTmp1.Fields!ApellidoMaterno
                            chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles).ToolTipText = "Dosificado el :" & oRsTmp1.Fields!FechaHoraVisita & " por el Prof. de Salud:  " & oRsTmp1.Fields!Nombres & " " & oRsTmp1.Fields!ApellidoPaterno & " " & oRsTmp1.Fields!ApellidoMaterno
                            mo_Formulario.HabilitarDeshabilitar chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            If ml_IdVisita = oRsTmp1.Fields!IdVisita Then mo_Formulario.HabilitarDeshabilitar chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles), True
                            lbDosis6Dada = True
                        Else
                            mo_Formulario.HabilitarDeshabilitar chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                            If Not chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles).Enabled Then
                                If Not lbDosisVisita Then
                                    mo_Formulario.HabilitarDeshabilitar chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles), True
                                End If
                            End If
                            lbDosis6Dada = False
                        End If
                    End If
                    If oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = 4 Or oRsEnfermeriaMedicamentos.Fields!idEstadoDetalle = 0 Then
                        mo_Formulario.HabilitarDeshabilitar chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles), False
                    End If
                End If
            
            End If
            oRsEnfermeriaMedicamentos.MoveNext
        Loop
    End If
    If oRsTmp1.State = 1 Then oRsTmp1.Close
    Set oRsTmp1 = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub

Sub CalculaCantidadGraficos(oConexion As Connection)
    Dim oRsTmp1 As New Recordset
    Dim lnNroGrafico As Integer
    Dim lnNroHoja As Integer
    Dim lnNroGraficoHoja As Integer
    Set oRsTmp1 = mo_AdminAdmision.Enfermeria_TotalDatosGrafico(oConexion)
    lnNroGrafico = 0
    If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
            lnNroGrafico = lnNroGrafico + 1
            lnNroHoja = IIf(Round(lnNroGrafico / 3) < lnNroGrafico / 3, Round(lnNroGrafico / 3) + 1, Round(lnNroGrafico / 3))
            lnNroGraficoHoja = IIf(lnNroGrafico Mod 3 = 0, 3, lnNroGrafico Mod 3)
            oRsEnfermeriaCatalogoGraficos.AddNew
            oRsEnfermeriaCatalogoGraficos.Fields!IdVariable = oRsTmp1.Fields!IdVariable
            oRsEnfermeriaCatalogoGraficos.Fields!IdDominio = oRsTmp1.Fields!IdDominio
            oRsEnfermeriaCatalogoGraficos.Fields!Texto = oRsTmp1.Fields!Texto
            oRsEnfermeriaCatalogoGraficos.Fields!NroHoja = lnNroHoja
            oRsEnfermeriaCatalogoGraficos.Fields!NroGraficoHoja = lnNroGraficoHoja
            oRsEnfermeriaCatalogoGraficos.Update
            oRsTmp1.MoveNext
        Loop
    End If
   ml_TotalDeGraficos = oRsEnfermeriaCatalogoGraficos.RecordCount
   scrollGraficos.Max = IIf(Round(ml_TotalDeGraficos / 3) < ml_TotalDeGraficos / 3, Round(ml_TotalDeGraficos / 3) + 1, Round(ml_TotalDeGraficos / 3))
   scrollGraficos.Min = 1
   If oRsTmp1.State = 1 Then oRsTmp1.Close
   Set oRsTmp1 = Nothing
End Sub

Private Sub grdHistoricoVisitas_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    InicializarLaGrilla grdHistoricoVisitas
End Sub

Private Sub InicializarLaGrilla(oGrilla As SSUltraGrid)
    Select Case oGrilla.Name
    Case "grdHistoricoVisitas"
        oGrilla.Bands(0).Columns("IdVisita").Hidden = True
        oGrilla.Bands(0).Columns("Descripcion").Header.Caption = "Visita"
        oGrilla.Bands(0).Columns("Descripcion").Width = 490
        oGrilla.Bands(0).Columns("Descripcion").Activation = ssActivationActivateNoEdit
        oGrilla.Bands(0).Columns("FechaHoraVisita").Header.Caption = "Fecha"
        oGrilla.Bands(0).Columns("FechaHoraVisita").Width = 1410
        oGrilla.Bands(0).Columns("FechaHoraVisita").Activation = ssActivationActivateNoEdit
    Case "grdPruebas"
    End Select
End Sub

Private Sub LisBoxVariable_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, LisBoxVariable(Index)
End Sub

Private Sub scrollGraficos_Change()
    MuestraGraficosPorHoja
End Sub

Private Sub scrollGraficos_Scroll()
    MuestraGraficosPorHoja
End Sub

Public Sub MuestraGraficosPorHoja()
    Dim lnNroHoja As Integer
    Dim lnNroTotalHojas As Integer
    lnNroHoja = scrollGraficos.Value
    lnNroTotalHojas = scrollGraficos.Max
    'Limpiar Graficos
    'Ocultar Graficos
    Me.FrameGrafico1.Visible = False
    Me.FrameGrafico2.Visible = False
    Me.FrameGrafico3.Visible = False
    
    If oRsEnfermeriaCatalogoGraficos.RecordCount > 0 Then
        oRsEnfermeriaCatalogoGraficos.MoveFirst
        Do While Not oRsEnfermeriaCatalogoGraficos.EOF
            If oRsEnfermeriaCatalogoGraficos.Fields!NroHoja = lnNroHoja Then
                Select Case oRsEnfermeriaCatalogoGraficos.Fields!NroGraficoHoja
                Case 1
                    Me.FrameGrafico1.Visible = True
                    CSGrafico1.ToolTipText = oRsEnfermeriaCatalogoGraficos.Fields!Texto
                    CSGrafico1.Tag = oRsEnfermeriaCatalogoGraficos.Fields!IdVariable
                    CargaGraficoChartSpace CSGrafico1, oRsEnfermeriaCatalogoGraficos.Fields!IdVariable, oRsEnfermeriaCatalogoGraficos.Fields!Texto
                Case 2
                    Me.FrameGrafico2.Visible = True
                    CSGrafico2.ToolTipText = oRsEnfermeriaCatalogoGraficos.Fields!Texto
                    CSGrafico2.Tag = oRsEnfermeriaCatalogoGraficos.Fields!IdVariable
                    CargaGraficoChartSpace CSGrafico2, oRsEnfermeriaCatalogoGraficos.Fields!IdVariable, oRsEnfermeriaCatalogoGraficos.Fields!Texto
                Case 3
                    Me.FrameGrafico3.Visible = True
                    CSGrafico3.ToolTipText = oRsEnfermeriaCatalogoGraficos.Fields!Texto
                    CSGrafico3.Tag = oRsEnfermeriaCatalogoGraficos.Fields!IdVariable
                    CargaGraficoChartSpace CSGrafico3, oRsEnfermeriaCatalogoGraficos.Fields!IdVariable, oRsEnfermeriaCatalogoGraficos.Fields!Texto
                End Select
            End If
            oRsEnfermeriaCatalogoGraficos.MoveNext
        Loop
    End If
    lblMensaje.Caption = "Pagina " & CStr(lnNroHoja) & " de " & CStr(lnNroTotalHojas)
End Sub

Sub ConfiguraVariables(oConexion As Connection)
   Dim ml_FilaDatoVariable As Integer
   Dim ml_ColumnaDatoVariable As Integer
   Dim lnTabIndex As Integer
   Dim oRsTmp1 As New Recordset
   'Configura controles para los datos de cabecera
    Set oRsTmp1 = mo_AdminAdmision.Enfermeria_CatalogoVariablesSeleccionar(oConexion)
    If oRsTmp1.RecordCount > 0 Then
      oRsTmp1.MoveFirst
      Do While Not oRsTmp1.EOF
            oRsEnfermeriaCatalogoVariables.AddNew
            oRsEnfermeriaCatalogoVariables.Fields!IdVariable = oRsTmp1.Fields!IdVariable
            oRsEnfermeriaCatalogoVariables.Fields!IdDominio = oRsTmp1.Fields!IdDominio
            oRsEnfermeriaCatalogoVariables.Fields!OrdernDominio = oRsTmp1.Fields!OrdernDominio
            oRsEnfermeriaCatalogoVariables.Fields!Texto = oRsTmp1.Fields!Texto
            oRsEnfermeriaCatalogoVariables.Fields!tipo = oRsTmp1.Fields!tipo
            oRsEnfermeriaCatalogoVariables.Fields!Ancho = oRsTmp1.Fields!Ancho
            oRsEnfermeriaCatalogoVariables.Fields!EsDatoObligatorio = oRsTmp1.Fields!EsDatoObligatorio
            oRsEnfermeriaCatalogoVariables.Fields!TextoToolTip = oRsTmp1.Fields!TextoToolTip
            oRsEnfermeriaCatalogoVariables.Fields!EsDatoGrafico = oRsTmp1.Fields!EsDatoGrafico
            oRsEnfermeriaCatalogoVariables.Fields!TieneFormatoMask = oRsTmp1.Fields!TieneFormatoMask
            oRsEnfermeriaCatalogoVariables.Fields!FormatoMask = oRsTmp1.Fields!FormatoMask
            oRsEnfermeriaCatalogoVariables.Fields!TieneRango = oRsTmp1.Fields!TieneRango
            oRsEnfermeriaCatalogoVariables.Fields!RangoInicial = IIf(IsNull(oRsTmp1.Fields!RangoInicial), 0, oRsTmp1.Fields!RangoInicial)
            oRsEnfermeriaCatalogoVariables.Fields!rangoFinal = IIf(IsNull(oRsTmp1.Fields!rangoFinal), 0, oRsTmp1.Fields!rangoFinal)
            oRsEnfermeriaCatalogoVariables.Fields!PosicionFila = oRsTmp1.Fields!PosicionFila
            oRsEnfermeriaCatalogoVariables.Fields!PosicionColumna = oRsTmp1.Fields!PosicionColumna
            oRsEnfermeriaCatalogoVariables.Update
            oRsTmp1.MoveNext
      Loop
      oRsEnfermeriaCatalogoVariables.MoveFirst
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
    
    lnTabIndex = 15
    oRsEnfermeriaCatalogoVariables.MoveFirst
    Do While Not oRsEnfermeriaCatalogoVariables.EOF
        'Calcula coordenadas datos de cabecera
        ml_ColumnaDatoVariable = 120 + 4460 * (oRsEnfermeriaCatalogoVariables.Fields!PosicionColumna - 1)
        ml_FilaDatoVariable = 300 + 400 * (oRsEnfermeriaCatalogoVariables.Fields!PosicionFila - 1)
        
        'Visualiza los label para el ingreso de datos de las variables
        Load lblVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable)
        lblVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Visible = True
'       Set lblVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Container = frameTab(1)
        Set lblVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Container = frameTab(Val(oRsEnfermeriaCatalogoVariables.Fields!IdDominio))
        lblVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Left = ml_ColumnaDatoVariable
        lblVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Caption = Mid(oRsEnfermeriaCatalogoVariables.Fields!Texto, 1, 25)
        lblVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Top = ml_FilaDatoVariable + 20

        Load lbl2Puntos(oRsEnfermeriaCatalogoVariables.Fields!IdVariable)
        lbl2Puntos(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Visible = True
        Set lbl2Puntos(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Container = frameTab(Val(oRsEnfermeriaCatalogoVariables.Fields!IdDominio))
        lbl2Puntos(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Top = ml_FilaDatoVariable + 20
        lbl2Puntos(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Left = ml_ColumnaDatoVariable + 2300
        
        Select Case oRsEnfermeriaCatalogoVariables.Fields!tipo
            Case "ValorEntero"
                Load txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable)
                txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Visible = True
                Set txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Container = frameTab(Val(oRsEnfermeriaCatalogoVariables.Fields!IdDominio))
                txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Left = ml_ColumnaDatoVariable + 2400
                txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Top = ml_FilaDatoVariable
                txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).ToolTipText = oRsEnfermeriaCatalogoVariables.Fields!TextoToolTip
                txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).MaxLength = oRsEnfermeriaCatalogoVariables.Fields!Ancho
                txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).TabIndex = lnTabIndex
                
            Case "ValorTexto", "ValorDouble"
                Load TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable)
                TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Visible = True
                Set TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Container = frameTab(Val(oRsEnfermeriaCatalogoVariables.Fields!IdDominio))
                TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Left = ml_ColumnaDatoVariable + 2400
                TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Top = ml_FilaDatoVariable
                TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).ToolTipText = oRsEnfermeriaCatalogoVariables.Fields!TextoToolTip
                TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).MaxLength = Val(oRsEnfermeriaCatalogoVariables.Fields!Ancho)
                If oRsEnfermeriaCatalogoVariables.Fields!TieneFormatoMask Then
                     TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Mask = oRsEnfermeriaCatalogoVariables.Fields!FormatoMask
                     TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Tag = Replace(oRsEnfermeriaCatalogoVariables.Fields!FormatoMask, "#", "_")
                End If
                TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).TabIndex = lnTabIndex
                
            Case "ValorCombo"
                Load cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable)
                cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Visible = True
                Set cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Container = frameTab(Val(oRsEnfermeriaCatalogoVariables.Fields!IdDominio))
                cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Left = ml_ColumnaDatoVariable + 2400
                cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Top = ml_FilaDatoVariable
                cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).ToolTipText = oRsEnfermeriaCatalogoVariables.Fields!TextoToolTip
                
                If oRsTmp1.State = 1 Then oRsTmp1.Close
                Set mo_cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).MiComboBox = cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable)
                mo_cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).BoundColumn = "IdValorCombo"
                mo_cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).ListField = "ComboTexto"
                Set oRsTmp1 = mo_AdminAdmision.Enfermeria_VariablesComboSeleccionarValores(oConexion, oRsEnfermeriaCatalogoVariables.Fields!IdVariable)
                Set mo_cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).RowSource = oRsTmp1
                If oRsTmp1.RecordCount > 0 Then mo_cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).BoundText = 1
                oRsTmp1.Close
                Set oRsTmp1 = Nothing
                cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).TabIndex = lnTabIndex
                
            Case "ValorMultiple"
                Load LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable)
                LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Visible = True
                Set LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Container = frameTab(Val(oRsEnfermeriaCatalogoVariables.Fields!IdDominio))
                LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Left = ml_ColumnaDatoVariable + 2400
                LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Top = ml_FilaDatoVariable
                LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).ToolTipText = oRsEnfermeriaCatalogoVariables.Fields!TextoToolTip
                
                If oRsTmp1.State = 1 Then oRsTmp1.Close
                Set mo_cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).MiComboBox = LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable)
                mo_cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).BoundColumn = "IdValorCombo"
                mo_cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).ListField = "ComboTexto"
                Set oRsTmp1 = mo_AdminAdmision.Enfermeria_VariablesComboSeleccionarValores(oConexion, oRsEnfermeriaCatalogoVariables.Fields!IdVariable)
                Set mo_cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).RowSource = oRsTmp1
                
                oRsTmp1.Close
                Set oRsTmp1 = Nothing
                LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).TabIndex = lnTabIndex
                
            Case "ValorCheck"
                Load chkVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable)
                chkVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Visible = True
                Set chkVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Container = frameTab(Val(oRsEnfermeriaCatalogoVariables.Fields!IdDominio))
                chkVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Left = ml_ColumnaDatoVariable + 2400
                chkVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Top = ml_FilaDatoVariable
                chkVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Caption = oRsEnfermeriaCatalogoVariables.Fields!TextoToolTip
                chkVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).ToolTipText = oRsEnfermeriaCatalogoVariables.Fields!TextoToolTip
                chkVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).TabIndex = lnTabIndex
        
        End Select
        lnTabIndex = lnTabIndex + 1
        oRsEnfermeriaCatalogoVariables.MoveNext
    Loop
    txtObservaciones.TabIndex = lnTabIndex
End Sub

Sub CargarDatosAlosControles()
    Dim oRsTmp As New Recordset
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    LimpiaTodosControles True
    'Consultar ultima visita al paciente
    mb_EsNuevaVisita = False
    ml_IdVisita = 0
    ml_IdDiaVisita = 0
    ml_IdServicio = 0
    ml_idCama = 0
    ml_IdEmpleado = 0
    ml_IdUltimaVisitaRegistrada = 0
    Set oRsTmp = mo_AdminAdmision.Enfermeria_ConsultarUltimaVisita(oConexion, ml_idCuentaAtencion)
    If oRsTmp.RecordCount = 0 Then
        mb_ValidaCheckValorizacion = True
        mb_EsNuevaVisita = True
        ml_IdVisita = 1
        ml_IdDiaVisita = 1
    Else
        ml_IdVisita = oRsTmp.Fields!IdVisita
        ml_IdUltimaVisitaRegistrada = oRsTmp.Fields!IdVisita
        mb_EsNuevaVisita = False
    End If
    
    If mb_EsNuevaVisita = True Then
        oRsEnfermeriaHistoricoVisitas.AddNew
        oRsEnfermeriaHistoricoVisitas.Fields!IdVisita = 1
        oRsEnfermeriaHistoricoVisitas.Fields!Descripcion = "Nº 1"
        oRsEnfermeriaHistoricoVisitas.Fields!FechaHoraVisita = "Nuevo"
        oRsEnfermeriaHistoricoVisitas.Update
        BuscaDatosActualesPaciente oConexion
        Me.FrameVisita.Caption = "Visita Nº 1 (Nuevo)"
        mo_Formulario.HabilitarDeshabilitar Me.txtFechaControl, True
        ConfiguraListadoMedicamentos ' Frank 0608
    Else
        'Carga Historico visita
        If oRsTmp.State = 1 Then oRsTmp.Close
        Set oRsTmp = mo_AdminAdmision.Enfermeria_CatalogoVisitas(oConexion, ml_idCuentaAtencion)
        If oRsTmp.RecordCount > 0 Then
            oRsTmp.MoveFirst
            Do While Not oRsTmp.EOF
                oRsEnfermeriaHistoricoVisitas.AddNew
                oRsEnfermeriaHistoricoVisitas.Fields!IdVisita = oRsTmp.Fields!IdVisita
                oRsEnfermeriaHistoricoVisitas.Fields!Descripcion = "Nº " & CStr(oRsTmp.Fields!IdVisita)
                oRsEnfermeriaHistoricoVisitas.Fields!FechaHoraVisita = Format(oRsTmp.Fields!FechaHoraVisita, sighEntidades.DevuelveFechaSoloFormato_DMY_HM)
                oRsEnfermeriaHistoricoVisitas.Update
                oRsTmp.MoveNext
            Loop
        End If
        'Cargar datos visita
        CargaDatosVisita oConexion
        Me.FrameVisita.Caption = "Visita Nº " & CStr(ml_IdVisita)
        mo_Formulario.HabilitarDeshabilitar Me.txtFechaControl, False
        ConfiguraListadoMedicamentos
    End If
    Me.TabsDominios.Tab = 0
    MuestraGraficosPorHoja 'Carga los graficos
    oConexion.Close
    Set oConexion = Nothing
End Sub

Sub CalcularIdDiaVisita()
    Dim oRsTmp As New Recordset
    Dim oConexion As New Connection
    Dim lcFecha As String
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    
    ml_IdDiaVisita = 1
    Set oRsTmp = mo_AdminAdmision.Enfermeria_CatalogoVisitas(oConexion, ml_idCuentaAtencion)
    If oRsTmp.RecordCount > 0 Then
        oRsTmp.MoveFirst
        lcFecha = Format(oRsTmp.Fields!FechaHoraVisita, sighEntidades.DevuelveFechaSoloFormato_DMY)
        Do While Not oRsTmp.EOF
            If Format(oRsTmp.Fields!FechaHoraVisita, sighEntidades.DevuelveFechaSoloFormato_DMY) <= Format(Me.txtFechaControl.Text, sighEntidades.DevuelveFechaSoloFormato_DMY) Then
                If lcFecha <> Format(oRsTmp.Fields!FechaHoraVisita, sighEntidades.DevuelveFechaSoloFormato_DMY) Then
                    lcFecha = Format(oRsTmp.Fields!FechaHoraVisita, sighEntidades.DevuelveFechaSoloFormato_DMY)
                    ml_IdDiaVisita = ml_IdDiaVisita + 1
                End If
            End If
            oRsTmp.MoveNext
        Loop
        If mb_EsNuevaVisita Then
            If lcFecha <> Format(Me.txtFechaControl.Text, sighEntidades.DevuelveFechaSoloFormato_DMY) Then
                ml_IdDiaVisita = ml_IdDiaVisita + 1
            End If
        End If
    End If
    
    oRsTmp.Close
    Set oRsTmp = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub

Sub CargaDatosVisita(ByVal oConexion As Connection)
        Dim oRsTmp As New Recordset
        Dim oRsTmp1 As New Recordset
        Dim lnItem As Long
        Dim lbSeleccionado As Boolean
    
        'Cargar cabecera visita
        If oRsTmp.State = 1 Then oRsTmp.Close
        Set oRsTmp = mo_AdminAdmision.EnfermeriaCargarDatosPacientePorVisita(oConexion, ml_idCuentaAtencion, ml_IdVisita)    'ml_IdVisita
        If oRsTmp.RecordCount > 0 Then
            oRsTmp.MoveFirst
            Me.txtNroHistoria.Text = oRsTmp.Fields!NroHistoriaClinica
            Me.txtNroCuenta.Text = oRsTmp.Fields!idCuentaAtencion
            Me.txtPrimerNombre.Text = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & ", " & Trim(oRsTmp.Fields!PrimerNombre) & " " & Trim(oRsTmp.Fields!SegundoNombre) & " " & Trim(oRsTmp.Fields!TercerNombre)
            Me.txtFechaControl.Text = Format(oRsTmp.Fields!FechaHoraVisita, sighEntidades.DevuelveFechaSoloFormato_DMY_HM)
            Me.txtNroCama.Text = IIf(oRsTmp.Fields!idCama = 0, "", oRsTmp.Fields!cama)
            ml_idCama = IIf(IsNull(oRsTmp.Fields!idCama), 0, oRsTmp.Fields!idCama)
            Me.txtServicio.Text = IIf(IsNull(oRsTmp.Fields!ServicioIngreso), "", oRsTmp.Fields!ServicioIngreso)
            Me.txtObservaciones.Text = IIf(IsNull(oRsTmp.Fields!Observaciones), "", oRsTmp.Fields!Observaciones)
            ml_IdServicio = IIf(IsNull(oRsTmp.Fields!IdServicio), 0, oRsTmp.Fields!IdServicio)
            ml_idPaciente = oRsTmp.Fields!idPaciente
            ml_IdEmpleado = oRsTmp.Fields!IdEmpleadoEnfermera
            Me.txtNombreEmpleado.Text = oRsTmp.Fields!Empleado
            mb_ValidaCheckValorizacion = False
            Me.chkValorizacion.Value = IIf(oRsTmp.Fields!IngresoValorizacion = True, 1, 0)
            mb_ValidaCheckValorizacion = True
            Me.btnAceptar.Enabled = True
            If oRsTmp.Fields!IdEstado <> 1 Then
                Me.btnAceptar.Enabled = False
            End If
            If chkValorizacion.Value Then
                HabilitarDehabilitarVariables True
            Else
                LimpiarDatosVariables
                HabilitarDehabilitarVariables False
            End If
            PacienteDatosAdicionalesCargar oConexion
        End If

        'cargar variables visita
        If oRsTmp.State = 1 Then oRsTmp.Close
        Set oRsTmp = mo_AdminAdmision.EnfermeriaCargarDatosVariablePorVisita(oConexion, ml_idCuentaAtencion, ml_IdVisita)
        Set oRsTmp1 = mo_AdminAdmision.EnfermeriaDatosComboPorVisita(oConexion, ml_idCuentaAtencion, ml_IdVisita)
        If oRsTmp.RecordCount > 0 Then
            oRsTmp.MoveFirst
            Do While Not oRsTmp.EOF
                Select Case oRsTmp.Fields!tipo
                    Case "ValorEntero"
                        txtVariable(oRsTmp.Fields!IdVariable).Text = IIf(IsNull(oRsTmp.Fields!VariableDato) = True, "", oRsTmp.Fields!VariableDato)
                    Case "ValorTexto", "ValorDouble"
                        TxtVariableFormato(oRsTmp.Fields!IdVariable).Text = IIf(IsNull(oRsTmp.Fields!VariableDato) = True, "", oRsTmp.Fields!VariableDato)
                    Case "ValorCombo"
                        mo_cmbVariable(oRsTmp.Fields!IdVariable).BoundText = Val(oRsTmp.Fields!VariableDato)
                    Case "ValorMultiple"
                        For lnItem = 0 To LisBoxVariable(oRsTmp.Fields!IdVariable).ListCount - 1
                            LisBoxVariable(oRsTmp.Fields!IdVariable).ListIndex = lnItem
                            lbSeleccionado = False
                            If oRsTmp1.RecordCount > 0 Then
                                oRsTmp1.MoveFirst
                                Do While Not oRsTmp1.EOF
                                    If oRsTmp1.Fields!IdVariable = oRsTmp.Fields!IdVariable And _
                                        oRsTmp1.Fields!IdValorCombo = Val(mo_cmbVariable(oRsTmp.Fields!IdVariable).BoundText) Then
                                        lbSeleccionado = True
                                        Exit Do
                                    End If
                                    oRsTmp1.MoveNext
                                Loop
                            End If
                            If lbSeleccionado = True Then LisBoxVariable(oRsTmp.Fields!IdVariable).Selected(lnItem) = True
                         Next
                         LisBoxVariable(oRsTmp.Fields!IdVariable).ListIndex = 0
                    Case "ValorCheck"
                        chkVariable(oRsTmp.Fields!IdVariable).Value = Val(oRsTmp.Fields!VariableDato)
                End Select
                oRsTmp.MoveNext
            Loop
        End If
End Sub

Sub BuscaDatosActualesPaciente(oConexion As Connection)
    Dim oRsTmp As New Recordset
    Dim oRsFiltrados As New Recordset
    Set oRsTmp = mo_AdminAdmision.Enfermeria_ConsultarDatosActualesPaciente(oConexion, ml_idCuentaAtencion)
    If oRsTmp.RecordCount > 0 Then
        oRsTmp.MoveFirst
        Me.txtNroHistoria.Text = oRsTmp.Fields!Historia
        Me.txtNroCuenta.Text = oRsTmp.Fields!idCuentaAtencion
        Me.txtPrimerNombre.Text = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & ", " & Trim(oRsTmp.Fields!PrimerNombre) & " " & Trim(oRsTmp.Fields!SegundoNombre) & " " & Trim(oRsTmp.Fields!TercerNombre)
        'Me.txtFechaControl.Text = lcBuscaParametro.RetornaFechaHoraServidorSQL
        Me.txtFechaControl.Text = lcBuscaParametro.RetornaFechaServidorSQL + " " + lcBuscaParametro.RetornaHoraServidorSQLserverFormatoGalenhos
        Me.txtNroCama.Text = IIf(IsNull(oRsTmp.Fields!idCama), "", oRsTmp.Fields!cama)
        ml_idCama = IIf(IsNull(oRsTmp.Fields!idCama), 0, oRsTmp.Fields!idCama)
        
        If IsNull(oRsTmp.Fields!IdServicioEgreso) Then
            Me.txtServicio.Text = IIf(IsNull(oRsTmp.Fields!ServicioIngreso), "", oRsTmp.Fields!ServicioIngreso)
            ml_IdServicio = IIf(IsNull(oRsTmp.Fields!IdServicioIngreso), 0, oRsTmp.Fields!IdServicioIngreso)
        Else
            Me.txtServicio.Text = IIf(IsNull(oRsTmp.Fields!ServicioEgreso), "", oRsTmp.Fields!ServicioEgreso)
            ml_IdServicio = IIf(IsNull(oRsTmp.Fields!IdServicioEgreso), 0, oRsTmp.Fields!IdServicioEgreso)
        End If
        
        ml_idPaciente = oRsTmp.Fields!idPaciente
        Me.btnAceptar.Enabled = True
        If oRsTmp.Fields!IdEstado <> 1 Then
            Me.btnAceptar.Enabled = False
        End If
        PacienteDatosAdicionalesCargar oConexion
    End If
End Sub

Sub PacienteDatosAdicionalesCargar(oConexion As Connection)
    'Dim oDoPacienteDatosAdd As New DoPacienteDatosAdd
    Set oDoPacienteDatosAdd = mo_AdminAdmision.PacientesDatosAdicionalesSeleccionarPorId(ml_idPaciente, oConexion)
    If oDoPacienteDatosAdd.idPaciente > 0 Then
       With oDoPacienteDatosAdd
          Me.txtAntecedentes.Text = .antecedentes
          Me.txtantecedAlergico.Text = .antecedAlergico
          Me.txtantecedObstetrico.Text = .antecedObstetrico
          Me.txtantecedQuirurgico.Text = .antecedQuirurgico
          Me.txtantecedFamiliar.Text = .antecedFamiliar
          Me.txtantecedPatologico.Text = .antecedPatologico
       End With
    End If
    'Set oDoPacienteDatosAdd = Nothing
End Sub

Sub LimpiaTodosControles(ByVal lbIncluyeHistoricoVisitas As Boolean)
    Dim lnItem As Long
    Me.chkValorizacion.Value = 1
    Me.txtNroHistoria.Text = ""
    Me.txtNroCuenta.Text = ""
    Me.txtPrimerNombre.Text = ""
    Me.txtFechaControl.Text = sighEntidades.FECHA_VACIA_DMY_HM
    Me.txtNroCama.Text = ""
    Me.txtServicio.Text = ""
    Me.txtantecedAlergico.Text = ""
    Me.txtAntecedentes.Text = ""
    Me.txtantecedFamiliar.Text = ""
    Me.txtantecedObstetrico.Text = ""
    Me.txtantecedPatologico.Text = ""
    Me.txtantecedQuirurgico.Text = ""
    Me.txtNombreEmpleado.Text = ""
    Me.txtObservaciones.Text = ""
    
    oRsEnfermeriaCatalogoVariables.MoveFirst
    Do While Not oRsEnfermeriaCatalogoVariables.EOF
        Select Case oRsEnfermeriaCatalogoVariables.Fields!tipo
            Case "ValorEntero"
                txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text = ""
            Case "ValorTexto"
                If oRsEnfermeriaCatalogoVariables.Fields!TieneFormatoMask Then
                    TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text = Replace(oRsEnfermeriaCatalogoVariables.Fields!FormatoMask, "#", "_")
                Else
                    TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text = ""
                End If
            Case "ValorDouble"
                TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text = ""
            Case "ValorCombo"
                mo_cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).BoundText = 1
            Case "ValorMultiple"
                For lnItem = 0 To LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).ListCount - 1
                    LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Selected(lnItem) = False
                Next
            Case "ValorCheck"
                chkVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Value = 0
        End Select
        oRsEnfermeriaCatalogoVariables.MoveNext
    Loop
    
    'Limpiando datos Historicos
    If lbIncluyeHistoricoVisitas = True Then
        With oRsEnfermeriaHistoricoVisitas
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
End Sub

Sub TituloDeForm()
    Select Case ml_TipoServicio
        Case sghHospitalizacion
                Me.Caption = "Visita de Enfermera - Hospitalización"
        Case sghEmergenciaConsultorios
                Me.Caption = "Visita de Enfermera - Emergencia"
        Case sghEmergenciaObservacion
                Me.Caption = "Visita de Enfermera - Emergencia"
    End Select
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
    Select Case KeyCode
    'Case vbKeyEscape
    '    btnCancelar_Click
    Case vbKeyF2
        If Me.btnAceptar.Enabled = True Then btnAceptar_Click
    End Select
End Sub

Private Sub btnAceptar_Click()
    If ValidarDatosObligatorios() Then
        If ValidarReglas() Then
            CargaDatosAlObjetosDeDatos
            If mb_EsNuevaVisita Then
                If AgregarDatos() Then
                    MsgBox "Los datos de la visita se registraron satisfactoriamente", vbInformation, Me.Caption
                    Me.Visible = False
                    LimpiarVariablesDeMemoria
                Else
                    MsgBox "No se pudo registrar los datos de la visita" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                End If
            Else
                If ModificarDatos() Then
                    MsgBox "Los datos de la visita se modificaron satisfactoriamente", vbInformation, Me.Caption
                    Me.Visible = False
                    LimpiarVariablesDeMemoria
                Else
                    MsgBox "No se pudo modificar los datos de la visita" + Chr(13) + ms_MensajeError, vbExclamation, Me.Caption
                End If
            End If
        End If
    End If
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
   LimpiarVariablesDeMemoria
End Sub

Function ValidarDatosObligatorios() As Boolean
    Dim sMensaje As String
    Dim lnIndiceLista As Integer
    Dim lbSeleccionado As Boolean
    sMensaje = ""
    ValidarDatosObligatorios = False
    If Me.txtFechaControl.Text = sighEntidades.FECHA_VACIA_DMY_HM Then
        sMensaje = sMensaje + "Ingrese la fecha de la visita." + Chr(13)
    Else
        If IsDate(txtFechaControl.Text) = False Then
            sMensaje = sMensaje + "La fecha de visita no tiene el formato correcto." + Chr(13)
        End If
    End If
    If ml_IdEmpleado = 0 Or Me.txtNombreEmpleado.Text = "" Then
        sMensaje = sMensaje + "Ingrese el nombre del profesional de salud." + Chr(13)
    End If
   
   If chkValorizacion.Value Then
        oRsEnfermeriaCatalogoVariables.MoveFirst
        Do While Not oRsEnfermeriaCatalogoVariables.EOF
            If oRsEnfermeriaCatalogoVariables.Fields!EsDatoObligatorio Then
                Select Case oRsEnfermeriaCatalogoVariables.Fields!tipo
                    Case "ValorEntero"
                        If txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text = "" Then
                            sMensaje = sMensaje + "Ingrese el valor de '" + oRsEnfermeriaCatalogoVariables.Fields!Texto + "'. " + Chr(13)
                        End If
                    Case "ValorTexto"
                        If TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text = "" Then
                            sMensaje = sMensaje + "Ingrese el valor de '" + oRsEnfermeriaCatalogoVariables.Fields!Texto + "'. " + Chr(13)
                        Else
                            If oRsEnfermeriaCatalogoVariables.Fields!TieneFormatoMask Then
                                If TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text = TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Tag Then
                                    sMensaje = sMensaje + "Ingrese el valor de '" + oRsEnfermeriaCatalogoVariables.Fields!Texto + "'. " + Chr(13)
                                End If
                            End If
                        End If
                    Case "ValorDouble"
                        If TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text = "" Then
                            sMensaje = sMensaje + "Ingrese el valor de '" + oRsEnfermeriaCatalogoVariables.Fields!Texto + "'. " + Chr(13)
                        Else
                            If InStr(TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable), ".") = Len(TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable)) Then
                                sMensaje = sMensaje + "El valor de '" + oRsEnfermeriaCatalogoVariables.Fields!Texto + "' no tiene el formato correcto." + Chr(13)
                            End If
                        End If
                    Case "ValorMultiple"
                        lbSeleccionado = False
                        For lnIndiceLista = 0 To LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).ListCount - 1
                            If LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Selected(lnIndiceLista) = True Then
                                lbSeleccionado = True
                            End If
                        Next
                        If lbSeleccionado = False Then
                            sMensaje = sMensaje + "Ingrese por lo menos un valor de '" + oRsEnfermeriaCatalogoVariables.Fields!Texto + "'. " + Chr(13)
                        End If
                End Select
            End If
            oRsEnfermeriaCatalogoVariables.MoveNext
        Loop
    End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   
   ValidarDatosObligatorios = True
End Function

Function ValidarReglas() As Boolean
    Dim oRsTmp As New ADODB.Recordset
    ValidarReglas = False
    If chkValorizacion.Value Then
        oRsEnfermeriaCatalogoVariables.MoveFirst
        Do While Not oRsEnfermeriaCatalogoVariables.EOF
            If oRsEnfermeriaCatalogoVariables.Fields!EsDatoObligatorio Then
                Select Case oRsEnfermeriaCatalogoVariables.Fields!tipo
                    Case "ValorEntero"
                        If txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text <> "" Then
                            If oRsEnfermeriaCatalogoVariables.Fields!TieneRango Then
                                If Val(txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text) < oRsEnfermeriaCatalogoVariables.Fields!RangoInicial Or _
                                    Val(txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text) > oRsEnfermeriaCatalogoVariables.Fields!rangoFinal Then
                                        MsgBox "El valor de '" + oRsEnfermeriaCatalogoVariables.Fields!Texto + "' debe estar entre " + Str(oRsEnfermeriaCatalogoVariables.Fields!RangoInicial) + " y " + Str(oRsEnfermeriaCatalogoVariables.Fields!rangoFinal) + ". ", vbExclamation, Me.Caption
                                        Exit Function
                                End If
                            End If
                        End If
                    Case "ValorDouble"
                        If TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text <> "" Then
                            If oRsEnfermeriaCatalogoVariables.Fields!TieneRango Then
                                If Val(TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text) < oRsEnfermeriaCatalogoVariables.Fields!RangoInicial Or _
                                    Val(TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text) > oRsEnfermeriaCatalogoVariables.Fields!rangoFinal Then
                                        MsgBox "El valor de '" + oRsEnfermeriaCatalogoVariables.Fields!Texto + "' debe estar entre " + Str(oRsEnfermeriaCatalogoVariables.Fields!RangoInicial) + " y " + Str(oRsEnfermeriaCatalogoVariables.Fields!rangoFinal) + ". ", vbExclamation, Me.Caption
                                        Exit Function
                                End If
                            End If
                        End If
                End Select
            End If
            oRsEnfermeriaCatalogoVariables.MoveNext
        Loop
    End If
    If mb_EsNuevaVisita Then
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        Set oRsTmp = mo_AdminAdmision.Enfermeria_ConsultarUltimaVisita(oConexion, ml_idCuentaAtencion)
        If oRsTmp.RecordCount > 0 Then
            If Format(oRsTmp.Fields!FechaHoraVisita, sighEntidades.DevuelveFechaSoloFormato_DMY_HM) >= txtFechaControl.Text Then
                MsgBox "La fecha de visita debe ser mayor a la última visita", vbExclamation, Me.Caption
                Exit Function
            End If
        End If
        oConexion.Close
        Set oConexion = Nothing
    End If
    
    If oRsEnfermeriaMedicamentos.RecordCount > 0 Then
        oRsEnfermeriaMedicamentos.MoveFirst
        Do While Not oRsEnfermeriaMedicamentos.EOF
           If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada = 0 Then
                If IsNumeric(Me.txtDosisProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles).Text) Then
                    Dim lnCantidadDosificada As Integer
                    Dim oConexion2 As New Connection
                    oConexion2.CommandTimeout = 300
                    oConexion2.CursorLocation = adUseClient
                    oConexion2.Open sighEntidades.CadenaConexion
                    Set oRsTmp = mo_AdminAdmision.Enfermeria_ConsultarTotalDosificadoProrenata(oConexion2, oRsEnfermeriaMedicamentos.Fields!idReceta, oRsEnfermeriaMedicamentos.Fields!idItem)
                    lnCantidadDosificada = 0
                    If oRsTmp.RecordCount > 0 Then
                        lnCantidadDosificada = IIf(IsNull(oRsTmp.Fields!CantidadDosificada), 0, oRsTmp.Fields!CantidadDosificada)
                    End If
                    If oRsEnfermeriaMedicamentos.Fields!CantidadDespachada - lnCantidadDosificada < CInt(Me.txtDosisProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles).Text) Then
                        MsgBox "La cantidad de dosis prorenata ingresada no debe ser mayor que " & oRsEnfermeriaMedicamentos.Fields!CantidadDespachada - lnCantidadDosificada & " [Cantidad Despachada - Cantidad Dosificada]", vbExclamation, Me.Caption
                        Exit Function
                    End If
                    oConexion2.Close
                    Set oConexion2 = Nothing
                End If
           End If
           oRsEnfermeriaMedicamentos.MoveNext
        Loop
    End If
    
    ValidarReglas = True
End Function


Sub CargaDatosAlObjetosDeDatos()
    Dim lnItem As Integer
    Dim lbCargoDatoDosis As Boolean
    With oDOEnfermeria_Visitas
        .FechaHoraVisita = Format(Me.txtFechaControl.Text, sighEntidades.DevuelveFechaSoloFormato_DMY_HM)
        .idCama = ml_idCama
        .idCuentaAtencion = ml_idCuentaAtencion
        .IdEmpleadoEnfermera = ml_IdEmpleado
        .IdServicio = ml_IdServicio
        .IdUsuarioAuditoria = ml_IdUsuario
        .IdVisita = ml_IdVisita
        .Observaciones = Me.txtObservaciones.Text
        .IngresoValorizacion = Me.chkValorizacion.Value
    End With
    
    If oRsEnfermeriaVariables.State = 1 Then
       Set oRsEnfermeriaVariables = Nothing
    End If
    With oRsEnfermeriaVariables
          .Fields.Append "IdCuentaAtencion", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdVisita", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdVariable", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "VariableDato", adVarChar, 255, adFldIsNullable + adFldUpdatable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    
    If oRsEnfermeriaValoresCombo.State = 1 Then
        Set oRsEnfermeriaValoresCombo = Nothing
    End If
    With oRsEnfermeriaValoresCombo
          .Fields.Append "IdCuentaAtencion", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdVisita", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdVariable", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdValorCombo", adInteger, 0, adFldIsNullable + adFldUpdatable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    
    If oRsEnfermeriaTratamientoDosis.State = 1 Then
        Set oRsEnfermeriaTratamientoDosis = Nothing
    End If
    With oRsEnfermeriaTratamientoDosis
          .Fields.Append "IdCuentaAtencion", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdVisita", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdDiaVisita", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdReceta", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "IdItem", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "Dosis", adInteger, 0, adFldIsNullable + adFldUpdatable
          .Fields.Append "DatoProrenata", adInteger, 0, adFldIsNullable + adFldUpdatable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With

    If oRsEnfermeriaCatalogoVariables.RecordCount > 0 Then
        oRsEnfermeriaCatalogoVariables.MoveFirst
        Do While Not oRsEnfermeriaCatalogoVariables.EOF
           oRsEnfermeriaVariables.AddNew
           oRsEnfermeriaVariables.Fields!idCuentaAtencion = ml_idCuentaAtencion
           oRsEnfermeriaVariables.Fields!IdVisita = ml_IdVisita
           oRsEnfermeriaVariables.Fields!IdVariable = oRsEnfermeriaCatalogoVariables.Fields!IdVariable
           
           Select Case oRsEnfermeriaCatalogoVariables.Fields!tipo
                Case "ValorEntero"
                    oRsEnfermeriaVariables.Fields!VariableDato = txtVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text
                Case "ValorTexto", "ValorDouble"
                    oRsEnfermeriaVariables.Fields!VariableDato = TxtVariableFormato(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Text
                Case "ValorCombo"
                    oRsEnfermeriaVariables.Fields!VariableDato = mo_cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).BoundText
                Case "ValorMultiple"
                    For lnItem = 0 To LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).ListCount - 1
                        If LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Selected(lnItem) = True Then
                            LisBoxVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).ListIndex = lnItem
                            oRsEnfermeriaValoresCombo.AddNew
                            oRsEnfermeriaValoresCombo.Fields!idCuentaAtencion = ml_idCuentaAtencion
                            oRsEnfermeriaValoresCombo.Fields!IdVisita = ml_IdVisita
                            oRsEnfermeriaValoresCombo.Fields!IdVariable = oRsEnfermeriaCatalogoVariables.Fields!IdVariable
                            oRsEnfermeriaValoresCombo.Fields!IdValorCombo = mo_cmbVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).BoundText
                            oRsEnfermeriaValoresCombo.Update
                        End If
                    Next
                Case "ValorCheck"
                    oRsEnfermeriaVariables.Fields!VariableDato = CStr(chkVariable(oRsEnfermeriaCatalogoVariables.Fields!IdVariable).Value)
           End Select
           oRsEnfermeriaVariables.Update
           oRsEnfermeriaCatalogoVariables.MoveNext
        Loop
        oRsEnfermeriaCatalogoVariables.MoveFirst
      End If
      
    If oRsEnfermeriaMedicamentos.RecordCount > 0 Then
        oRsEnfermeriaMedicamentos.MoveFirst
        Do While Not oRsEnfermeriaMedicamentos.EOF
           oRsEnfermeriaTratamientoDosis.AddNew
           oRsEnfermeriaTratamientoDosis.Fields!idCuentaAtencion = ml_idCuentaAtencion
           oRsEnfermeriaTratamientoDosis.Fields!IdVisita = ml_IdVisita
           oRsEnfermeriaTratamientoDosis.Fields!IdDiaVisita = ml_IdDiaVisita
           oRsEnfermeriaTratamientoDosis.Fields!idReceta = oRsEnfermeriaMedicamentos.Fields!idReceta
           oRsEnfermeriaTratamientoDosis.Fields!idItem = oRsEnfermeriaMedicamentos.Fields!idItem
           oRsEnfermeriaTratamientoDosis.Fields!Dosis = Null
           oRsEnfermeriaTratamientoDosis.Fields!DatoProrenata = Null
           If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada <> 0 Then
                lbCargoDatoDosis = False
                If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 1 Then
                    If Me.chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles).Enabled Then
                       If Me.chkDosis1(oRsEnfermeriaMedicamentos.Fields!IdControles).Value = 1 Then
                            oRsEnfermeriaTratamientoDosis.Fields!Dosis = 1
                        Else
                             oRsEnfermeriaTratamientoDosis.Fields!Dosis = Null
                        End If
                       lbCargoDatoDosis = True
                    Else
                        If Not IsNull(oRsEnfermeriaMedicamentos.Fields!Dosis) Then oRsEnfermeriaTratamientoDosis.Fields!Dosis = oRsEnfermeriaMedicamentos.Fields!Dosis
                    End If
                End If
                If lbCargoDatoDosis = False Then
                    If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 2 Then
                        If Me.chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles).Enabled Then
                           If Me.chkDosis2(oRsEnfermeriaMedicamentos.Fields!IdControles).Value = 1 Then
                                oRsEnfermeriaTratamientoDosis.Fields!Dosis = 2
                           Else
                                oRsEnfermeriaTratamientoDosis.Fields!Dosis = Null
                           End If
                           lbCargoDatoDosis = True
                        Else
                            If Not IsNull(oRsEnfermeriaMedicamentos.Fields!Dosis) Then oRsEnfermeriaTratamientoDosis.Fields!Dosis = oRsEnfermeriaMedicamentos.Fields!Dosis
                        End If
                    End If
                End If
                If lbCargoDatoDosis = False Then
                    If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 3 Then
                        If Me.chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles).Enabled Then
                           If Me.chkDosis3(oRsEnfermeriaMedicamentos.Fields!IdControles).Value = 1 Then
                                oRsEnfermeriaTratamientoDosis.Fields!Dosis = 3
                           Else
                                oRsEnfermeriaTratamientoDosis.Fields!Dosis = Null
                           End If
                           lbCargoDatoDosis = True
                        Else
                            If Not IsNull(oRsEnfermeriaMedicamentos.Fields!Dosis) Then oRsEnfermeriaTratamientoDosis.Fields!Dosis = oRsEnfermeriaMedicamentos.Fields!Dosis
                        End If
                    End If
                End If
                If lbCargoDatoDosis = False Then
                    If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 4 Then
                        If Me.chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles).Enabled Then
                           If Me.chkDosis4(oRsEnfermeriaMedicamentos.Fields!IdControles).Value = 1 Then
                                oRsEnfermeriaTratamientoDosis.Fields!Dosis = 4
                           Else
                                oRsEnfermeriaTratamientoDosis.Fields!Dosis = Null
                           End If
                           lbCargoDatoDosis = True
                        Else
                            If Not IsNull(oRsEnfermeriaMedicamentos.Fields!Dosis) Then oRsEnfermeriaTratamientoDosis.Fields!Dosis = oRsEnfermeriaMedicamentos.Fields!Dosis
                        End If
                    End If
                End If
                If lbCargoDatoDosis = False Then
                    If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 5 Then
                        If Me.chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles).Enabled Then
                           If Me.chkDosis5(oRsEnfermeriaMedicamentos.Fields!IdControles).Value = 1 Then
                                oRsEnfermeriaTratamientoDosis.Fields!Dosis = 5
                           Else
                                oRsEnfermeriaTratamientoDosis.Fields!Dosis = Null
                           End If
                           lbCargoDatoDosis = True
                        Else
                            If Not IsNull(oRsEnfermeriaMedicamentos.Fields!Dosis) Then oRsEnfermeriaTratamientoDosis.Fields!Dosis = oRsEnfermeriaMedicamentos.Fields!Dosis
                        End If
                    End If
                End If
                If lbCargoDatoDosis = False Then
                    If oRsEnfermeriaMedicamentos.Fields!idDosisRecetada >= 6 Then
                        If Me.chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles).Enabled Then
                           If Me.chkDosis6(oRsEnfermeriaMedicamentos.Fields!IdControles).Value = 1 Then
                               oRsEnfermeriaTratamientoDosis.Fields!Dosis = 6
                           Else
                                oRsEnfermeriaTratamientoDosis.Fields!Dosis = Null
                           End If
                           lbCargoDatoDosis = True
                        Else
                            If Not IsNull(oRsEnfermeriaMedicamentos.Fields!Dosis) Then oRsEnfermeriaTratamientoDosis.Fields!Dosis = oRsEnfermeriaMedicamentos.Fields!Dosis
                        End If
                    End If
                End If
           Else
                If Me.txtDosisProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles).Text <> "" Then oRsEnfermeriaTratamientoDosis.Fields!DatoProrenata = Me.txtDosisProrenata(oRsEnfermeriaMedicamentos.Fields!IdControles).Text
           End If
           oRsEnfermeriaTratamientoDosis.Update
           oRsEnfermeriaMedicamentos.MoveNext
        Loop
    End If
    With oDoPacienteDatosAdd
        .antecedentes = Me.txtAntecedentes.Text
        .antecedAlergico = Me.txtantecedAlergico.Text
        .antecedObstetrico = Me.txtantecedObstetrico.Text
        .antecedQuirurgico = Me.txtantecedQuirurgico.Text
        .antecedFamiliar = Me.txtantecedFamiliar.Text
        .antecedPatologico = Me.txtantecedPatologico.Text
    End With
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------
Function AgregarDatos() As Boolean
    AgregarDatos = False
    AgregarDatos = mo_AdminAdmision.EnfermeriaVisitasAgregar(oDOEnfermeria_Visitas, oRsEnfermeriaVariables, oRsEnfermeriaValoresCombo, oRsEnfermeriaTratamientoDosis, oDoPacienteDatosAdd)
    ms_MensajeError = mo_AdminAdmision.MensajeError
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean
    ModificarDatos = False
    ModificarDatos = mo_AdminAdmision.EnfermeriaVisitasModificar(oDOEnfermeria_Visitas, oRsEnfermeriaVariables, oRsEnfermeriaValoresCombo, oRsEnfermeriaTratamientoDosis)
    ms_MensajeError = mo_AdminAdmision.MensajeError
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    EliminarDatos = False
    EliminarDatos = mo_AdminAdmision.EnfermeriaVisitasEliminar(ml_idCuentaAtencion, ml_IdVisita)
    ms_MensajeError = mo_AdminAdmision.MensajeError
End Function


Sub LimpiarVariablesDeMemoria()
End Sub

Sub InicilizarParametros()
End Sub

Private Sub btnBuscarEmpleado_Click()
    CompletarDatosResponsable
End Sub

Sub CompletarDatosResponsable()
    'Dim oBusqueda As New EmpleadosBusqueda
    Dim oBusqueda As New SIGHNegocios.BuscaEmpleados
    Dim oDOEmpleado As New dOEmpleado

    'oBusqueda.Show 1
    Me.txtNombreEmpleado.Text = ""
    oBusqueda.MostrarFormulario
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOEmpleado = mo_AdminServiciosComunes.EmpleadosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDOEmpleado Is Nothing Then
            ml_IdEmpleado = oDOEmpleado.IdEmpleado
            Me.txtNombreEmpleado.Text = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If
End Sub

Private Sub txtVariable_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtVariable(Index)
End Sub

Private Sub txtVariable_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
            KeyAscii = 0
        End If
    End If
End Sub

Sub CargaGraficoChartSpace(ByVal oCsGrafico As ChartSpace, ByVal mlIdVariable As Integer, ByVal lcTituloGrafico As String)
    Dim lnFor As Integer
    Dim lnUltimoPunto As Integer
    Dim oConexion As New Connection
    Dim orsTemp As New ADODB.Recordset
    Dim lnRangoMaximo As Long
    Dim lcTipoVariable As String
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion

    xValues = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250, 280)
    yValues = Array(10, 30, 50, 80, 100, 120, 150, 160, 180, 190, 200, 210, 220, 230, 250, 280)

    If ml_IdVisita = 0 Then Exit Sub
    Set orsTemp = mo_AdminAdmision.enfermeria_ConsultaValoresVariableGrafico(oConexion, ml_idCuentaAtencion, mlIdVariable, ml_IdVisita)
    
    If mb_EsNuevaVisita = True Then
        lnUltimoPunto = orsTemp.RecordCount
    Else
        If orsTemp.RecordCount = 0 Then
            lnUltimoPunto = 0
        Else
            lnUltimoPunto = orsTemp.RecordCount - 1
        End If
    End If
    ReDim xValues(lnUltimoPunto)
    lnRangoMaximo = 110
    If orsTemp.RecordCount > 0 Then
        orsTemp.MoveLast
        For lnFor = (orsTemp.RecordCount - 1) To 0 Step -1
           xValues(lnFor) = orsTemp.Fields!IdVisita
           yValues(lnFor) = orsTemp.Fields!VariableDato
           If Val(orsTemp.Fields!VariableDato) > lnRangoMaximo Then
                lnRangoMaximo = Val(orsTemp.Fields!VariableDato) + 50
           End If
           orsTemp.MovePrevious
        Next
    End If
    If Me.chkValorizacion.Value = 1 Then
        xValues(lnUltimoPunto) = ml_IdVisita
        oRsEnfermeriaCatalogoVariables.MoveFirst
        Do While Not oRsEnfermeriaCatalogoVariables.EOF
            If oRsEnfermeriaCatalogoVariables.Fields!IdVariable = mlIdVariable Then
                lcTipoVariable = oRsEnfermeriaCatalogoVariables.Fields!tipo
            End If
            oRsEnfermeriaCatalogoVariables.MoveNext
        Loop
        If lcTipoVariable = "ValorEntero" Then
            yValues(lnUltimoPunto) = IIf(Trim(Me.txtVariable(mlIdVariable).Text) = "", 0, Val(Me.txtVariable(mlIdVariable).Text))
        Else
            yValues(lnUltimoPunto) = IIf(Trim(Me.TxtVariableFormato(mlIdVariable).Text) = "", 0, Val(Me.TxtVariableFormato(mlIdVariable).Text))
        End If
    End If
    '
    oCsGrafico.Clear
    oCsGrafico.DisplayToolbar = False
    Set owcChart = oCsGrafico.Charts.Add
    owcChart.HasTitle = True
    owcChart.Title.Caption = lcTituloGrafico '+ " vs Visita"
    owcChart.Title.Font.Name = "Arial Narrow"
    owcChart.Title.Font.Size = 10
    owcChart.Title.Font.Color = vbBlue
    owcChart.Axes(chAxisPositionBottom).Font.Name = "Arial narrow"
    owcChart.Axes(chAxisPositionBottom).Font.Size = 8
    owcChart.Axes(chAxisPositionBottom).Font.Color = vbBlue
    owcChart.Axes(chAxisPositionBottom).Scaling.Minimum = 0
    owcChart.Axes(chAxisPositionLeft).Font.Name = "Arial narrow"
    owcChart.Axes(chAxisPositionLeft).Font.Size = "8"
    owcChart.Axes(chAxisPositionLeft).Font.Color = vbBlue
    owcChart.Axes(chAxisPositionLeft).Scaling.Minimum = 0
    owcChart.Axes(chAxisPositionLeft).Scaling.Maximum = lnRangoMaximo '110
'    owcChart.Axes(1).HasTitle = 1
'    owcChart.Axes(1).Font.Name = "Arial Narrow"
'    owcChart.Axes(1).Font.Size = 8
'    owcChart.Axes(1).Font.Color = vbBlue
'    owcChart.Axes(1).Title.Caption = ""
'    owcChart.Axes(1).Title.Font.Name = "Arial Narrow"
'    owcChart.Axes(1).Title.Font.Size = 8
'    owcChart.Axes(1).Title.Font.Color = vbBlue
    '
    Set owcSeries = owcChart.SeriesCollection.Add
    With owcSeries
        .Caption = ""
        .SetData chDimCategories, chDataLiteral, xValues
        .SetData chDimValues, chDataLiteral, yValues
        .Type = chChartTypeLineMarkers
        .Line.Color = vbRed
        .Line.Weight = 3
        .Marker.Style = chMarkerStyleCircle
        .Line.DashStyle = chLineSolid
        .DataLabelsCollection.Add
    End With
    
    oConexion.Close
    Set oConexion = Nothing
End Sub

Private Sub txtVariable_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim lbActualizoVariableGrafico As Boolean
    lbActualizoVariableGrafico = False
    oRsEnfermeriaCatalogoVariables.MoveFirst
    Do While Not oRsEnfermeriaCatalogoVariables.EOF
        If oRsEnfermeriaCatalogoVariables.Fields!IdVariable = Index Then
            If oRsEnfermeriaCatalogoVariables.Fields!EsDatoGrafico = True Then
                lbActualizoVariableGrafico = True
            End If
        End If
        oRsEnfermeriaCatalogoVariables.MoveNext
    Loop
    If lbActualizoVariableGrafico Then MuestraGraficosPorHoja
End Sub

Private Sub TxtVariableFormato_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lbValidaDouble As Boolean
    lbValidaDouble = False
    oRsEnfermeriaCatalogoVariables.MoveFirst
    Do While Not oRsEnfermeriaCatalogoVariables.EOF
        If oRsEnfermeriaCatalogoVariables.Fields!IdVariable = Index Then
            If oRsEnfermeriaCatalogoVariables.Fields!tipo = "ValorDouble" Then
                lbValidaDouble = True
                Exit Do
            End If
        End If
        oRsEnfermeriaCatalogoVariables.MoveNext
    Loop
    If lbValidaDouble = True Then
        If Not (mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtVariableFormato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim lbActualizoVariableGrafico As Boolean
    lbActualizoVariableGrafico = False
    oRsEnfermeriaCatalogoVariables.MoveFirst
    Do While Not oRsEnfermeriaCatalogoVariables.EOF
        If oRsEnfermeriaCatalogoVariables.Fields!IdVariable = Index Then
            If oRsEnfermeriaCatalogoVariables.Fields!EsDatoGrafico = True Then
                lbActualizoVariableGrafico = True
            End If
        End If
        oRsEnfermeriaCatalogoVariables.MoveNext
    Loop
    If lbActualizoVariableGrafico Then MuestraGraficosPorHoja
End Sub

Private Sub TxtVariableFormato_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, TxtVariableFormato(Index)
End Sub





