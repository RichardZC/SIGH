VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form HerrExportaCitasWeb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exporta/Importa datos al Sistema Citas Web  (debe tener INTERNET activo)"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14250
   ControlBox      =   0   'False
   Icon            =   "HerrExportaCitasWeb.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8835
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   15584
      _Version        =   393216
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
      TabCaption(0)   =   "Importar Datos"
      TabPicture(0)   =   "HerrExportaCitasWeb.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tabImportar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Exportar datos"
      TabPicture(1)   =   "HerrExportaCitasWeb.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "CS/PS que podrán Referir"
      TabPicture(2)   =   "HerrExportaCitasWeb.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label8"
      Tab(2).Control(1)=   "grdConsultorios"
      Tab(2).Control(2)=   "grdCSreferir"
      Tab(2).Control(3)=   "fraAgregar"
      Tab(2).Control(4)=   "btnAgregarDx"
      Tab(2).Control(5)=   "Frame5"
      Tab(2).Control(6)=   "cmdEliminar"
      Tab(2).Control(7)=   "Frame(0)"
      Tab(2).ControlCount=   8
      Begin TabDlg.SSTab tabImportar 
         Height          =   6975
         Left            =   75
         TabIndex        =   47
         Top             =   405
         Width           =   14160
         _ExtentX        =   24977
         _ExtentY        =   12303
         _Version        =   393216
         Tabs            =   2
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
         TabCaption(0)   =   "Por IMPORTAR desde la WEB"
         TabPicture(0)   =   "HerrExportaCitasWeb.frx":0D1E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label11"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "grdCitasWebPorImportar"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chkSeImportacionAutomatica"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "chkTodos"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "btnEliminaSolicitudCita"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtCitasRechazadas"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Ya con CITAS asignadas"
         TabPicture(1)   =   "HerrExportaCitasWeb.frx":0D3A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FraFiltrar"
         Tab(1).Control(1)=   "btnQuitarDx"
         Tab(1).Control(2)=   "grdCitasWeb"
         Tab(1).ControlCount=   3
         Begin VB.TextBox txtCitasRechazadas 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Left            =   5400
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   69
            Text            =   "HerrExportaCitasWeb.frx":0D56
            Top             =   5880
            Width           =   8235
         End
         Begin VB.Frame FraFiltrar 
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
            Left            =   -74925
            TabIndex        =   62
            Top             =   405
            Width           =   13470
            Begin VB.CommandButton btnBuscar 
               Height          =   315
               Left            =   6030
               Picture         =   "HerrExportaCitasWeb.frx":0D69
               Style           =   1  'Graphical
               TabIndex        =   65
               Top             =   285
               Width           =   1305
            End
            Begin VB.CommandButton btnLimpiar 
               Height          =   315
               Left            =   6045
               Picture         =   "HerrExportaCitasWeb.frx":39B2
               Style           =   1  'Graphical
               TabIndex        =   67
               Top             =   645
               Width           =   1275
            End
            Begin VB.TextBox txtBusqAMaterno 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2190
               TabIndex        =   64
               Top             =   555
               Width           =   2010
            End
            Begin VB.TextBox txtBusqApaterno 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   165
               TabIndex        =   63
               Top             =   555
               Width           =   2010
            End
            Begin VB.Label Label 
               Caption         =   "Apellido Paterno            Apellido Materno"
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
               Left            =   195
               TabIndex        =   66
               Top             =   210
               Width           =   4020
            End
         End
         Begin VB.CommandButton btnEliminaSolicitudCita 
            DisabledPicture =   "HerrExportaCitasWeb.frx":658E
            DownPicture     =   "HerrExportaCitasWeb.frx":6919
            Height          =   315
            Left            =   13635
            Picture         =   "HerrExportaCitasWeb.frx":6CAC
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Libera CITA"
            Top             =   840
            Width           =   420
         End
         Begin VB.CheckBox chkTodos 
            Caption         =   "Seleccionar Todos/ninguno"
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
            Left            =   105
            TabIndex        =   60
            Top             =   5880
            Width           =   2595
         End
         Begin VB.CheckBox chkSeImportacionAutomatica 
            Caption         =   "Se importa Automáticamente"
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
            Left            =   135
            TabIndex        =   56
            Top             =   465
            Width           =   3135
         End
         Begin VB.Frame Frame2 
            Caption         =   "Importar Citas Web (solo cupos CONFIRMADOS)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   13155
            TabIndex        =   50
            Top             =   315
            Visible         =   0   'False
            Width           =   885
            Begin VB.TextBox txtArchivoCuposConfirmados 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   9585
               TabIndex        =   52
               Text            =   "c:\cuposWeb.txt"
               Top             =   330
               Width           =   1995
            End
            Begin Threed.SSOption optSisGalenPlus 
               Height          =   345
               Left            =   6135
               TabIndex        =   51
               Top             =   300
               Width           =   2790
               _ExtentX        =   4921
               _ExtentY        =   609
               _Version        =   262144
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Desde Citas Web SisGalenPlus"
            End
            Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar7 
               Height          =   300
               Left            =   11640
               TabIndex        =   53
               Top             =   360
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BrushStyle      =   0
               Color           =   6956042
            End
            Begin Threed.SSOption optMINSA 
               Height          =   345
               Left            =   90
               TabIndex        =   54
               Top             =   300
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   609
               _Version        =   262144
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "En forma automática"
               Value           =   -1
            End
            Begin Threed.SSOption optManual 
               Height          =   345
               Left            =   3330
               TabIndex        =   59
               Top             =   300
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   609
               _Version        =   262144
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "En forma Manual"
            End
            Begin VB.Label Label2 
               Caption         =   "Ruta"
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
               Left            =   9000
               TabIndex        =   55
               Top             =   360
               Width           =   555
            End
         End
         Begin VB.CommandButton btnQuitarDx 
            DisabledPicture =   "HerrExportaCitasWeb.frx":703D
            DownPicture     =   "HerrExportaCitasWeb.frx":73C8
            Height          =   315
            Left            =   -61395
            Picture         =   "HerrExportaCitasWeb.frx":775B
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Libera CITA"
            Top             =   1845
            Width           =   420
         End
         Begin UltraGrid.SSUltraGrid grdCitasWeb 
            Height          =   5220
            Left            =   -74925
            TabIndex        =   49
            Top             =   1515
            Width           =   13470
            _ExtentX        =   23760
            _ExtentY        =   9208
            _Version        =   131072
            GridFlags       =   17040388
            UpdateMode      =   2
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
            Caption         =   "Citas Web con CITAS asignadas"
         End
         Begin UltraGrid.SSUltraGrid grdCitasWebPorImportar 
            Height          =   4995
            Left            =   90
            TabIndex        =   57
            Top             =   780
            Width           =   13530
            _ExtentX        =   23865
            _ExtentY        =   8811
            _Version        =   131072
            GridFlags       =   17040388
            UpdateMode      =   2
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
            Caption         =   "Citas Web Solicitadas"
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Citas Rechazadas"
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
            Left            =   3990
            TabIndex        =   68
            Top             =   5910
            Width           =   1365
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Ordenado por"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   0
         Left            =   -67455
         TabIndex        =   43
         Top             =   6690
         Width           =   6570
         Begin VB.CommandButton cmdAdicionaConsultoriosNuevos 
            Caption         =   "..."
            Height          =   255
            Left            =   6030
            TabIndex        =   70
            Top             =   240
            Width           =   285
         End
         Begin Threed.SSOption optPorConsultorio 
            Height          =   285
            Left            =   585
            TabIndex        =   44
            Top             =   225
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Consultorio"
            Value           =   -1
         End
         Begin Threed.SSOption optCsPS 
            Height          =   345
            Left            =   2535
            TabIndex        =   45
            Top             =   180
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "CsPs"
         End
      End
      Begin VB.CommandButton cmdEliminar 
         DisabledPicture =   "HerrExportaCitasWeb.frx":7AEC
         DownPicture     =   "HerrExportaCitasWeb.frx":7E77
         Height          =   315
         Left            =   -68100
         Picture         =   "HerrExportaCitasWeb.frx":820A
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1065
         Width           =   540
      End
      Begin VB.Frame Frame5 
         Height          =   1215
         Left            =   -74925
         TabIndex        =   37
         Top             =   7515
         Width           =   14055
         Begin VB.CommandButton cmdGrabaReferidos 
            Caption         =   "Grabar en Web"
            DisabledPicture =   "HerrExportaCitasWeb.frx":859B
            DownPicture     =   "HerrExportaCitasWeb.frx":89FB
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   5745
            Picture         =   "HerrExportaCitasWeb.frx":8E70
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   225
            Width           =   1365
         End
         Begin VB.CommandButton cmdCancelaReferidos 
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "HerrExportaCitasWeb.frx":92E5
            DownPicture     =   "HerrExportaCitasWeb.frx":97A9
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   7189
            Picture         =   "HerrExportaCitasWeb.frx":9C95
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   210
            Width           =   1335
         End
      End
      Begin VB.CommandButton btnAgregarDx 
         DisabledPicture =   "HerrExportaCitasWeb.frx":A181
         DownPicture     =   "HerrExportaCitasWeb.frx":A56A
         Height          =   315
         Left            =   -68115
         Picture         =   "HerrExportaCitasWeb.frx":A976
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   690
         Width           =   555
      End
      Begin VB.Frame fraAgregar 
         Caption         =   "Adicionar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2925
         Left            =   -74925
         TabIndex        =   24
         Top             =   4350
         Width           =   5040
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "Cancelar"
            DisabledPicture =   "HerrExportaCitasWeb.frx":AD82
            DownPicture     =   "HerrExportaCitasWeb.frx":B246
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3045
            Picture         =   "HerrExportaCitasWeb.frx":B732
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1905
            Width           =   1335
         End
         Begin VB.CommandButton cmdAdicionar 
            Caption         =   "Adicionar"
            DisabledPicture =   "HerrExportaCitasWeb.frx":BC1E
            DownPicture     =   "HerrExportaCitasWeb.frx":C07E
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1650
            Picture         =   "HerrExportaCitasWeb.frx":C4F3
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1920
            Width           =   1365
         End
         Begin VB.TextBox txtEmail 
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
            Left            =   1635
            MaxLength       =   50
            TabIndex        =   33
            Top             =   1515
            Width           =   3315
         End
         Begin VB.TextBox txtCSclave 
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
            Left            =   1635
            MaxLength       =   20
            TabIndex        =   31
            Top             =   1080
            Width           =   3315
         End
         Begin VB.TextBox txtCSusuario 
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
            Left            =   1635
            MaxLength       =   20
            TabIndex        =   29
            Top             =   675
            Width           =   3315
         End
         Begin VB.CommandButton btnBuscarEstablecimientoDestino 
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
            Left            =   825
            Picture         =   "HerrExportaCitasWeb.frx":C968
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   300
            Width           =   330
         End
         Begin VB.TextBox txtCS 
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
            Left            =   1635
            TabIndex        =   26
            Top             =   285
            Width           =   3315
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Email"
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
            TabIndex        =   32
            Top             =   1590
            Width           =   405
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Clave"
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
            TabIndex        =   30
            Top             =   1170
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
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
            TabIndex        =   28
            Top             =   765
            Width           =   585
         End
         Begin VB.Label lblCs 
            AutoSize        =   -1  'True
            Caption         =   "CS o PS"
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
            Top             =   375
            Width           =   645
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   30
         TabIndex        =   17
         Top             =   7440
         Width           =   14175
         Begin VB.CommandButton cmdImportarWeb 
            Caption         =   "Importar Citas Web"
            DisabledPicture =   "HerrExportaCitasWeb.frx":CEF2
            DownPicture     =   "HerrExportaCitasWeb.frx":D352
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   5745
            Picture         =   "HerrExportaCitasWeb.frx":D7C7
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   210
            Width           =   1365
         End
         Begin VB.CommandButton cmdSalir 
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "HerrExportaCitasWeb.frx":DC3C
            DownPicture     =   "HerrExportaCitasWeb.frx":E100
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   7189
            Picture         =   "HerrExportaCitasWeb.frx":E5EC
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   210
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1080
         Left            =   -74880
         TabIndex        =   6
         Top             =   7590
         Width           =   13935
         Begin VB.CommandButton btnCancelar 
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "HerrExportaCitasWeb.frx":EAD8
            DownPicture     =   "HerrExportaCitasWeb.frx":EF9C
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   7170
            Picture         =   "HerrExportaCitasWeb.frx":F488
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   135
            Width           =   1335
         End
         Begin VB.CommandButton btnAceptar 
            Caption         =   "Exporta Citas Web"
            DisabledPicture =   "HerrExportaCitasWeb.frx":F974
            DownPicture     =   "HerrExportaCitasWeb.frx":FDD4
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   5715
            Picture         =   "HerrExportaCitasWeb.frx":10249
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   150
            Width           =   1365
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Exportar Programación Médica con cupos disponibles, que podrán ser usados en Citas Web"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7005
         Left            =   -74880
         TabIndex        =   1
         Top             =   600
         Width           =   13920
         Begin VB.CheckBox chkExportaTurnos 
            Caption         =   "Exporta Turnos"
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
            Left            =   12120
            TabIndex        =   22
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkFteFinanciamiento 
            Caption         =   "Exp. FteFinanciam"
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
            Left            =   10200
            TabIndex        =   20
            Top             =   360
            Width           =   1785
         End
         Begin VB.CheckBox chkExpPacientes 
            Caption         =   "Exporta Pacientes"
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
            Left            =   8280
            TabIndex        =   15
            Top             =   330
            Width           =   1845
         End
         Begin VB.CheckBox chkExpMedicos 
            Caption         =   "Exporta Médicos"
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
            Left            =   6450
            TabIndex        =   14
            Top             =   330
            Width           =   1725
         End
         Begin VB.CheckBox chkExpServicios 
            Caption         =   "Exporta Servicios"
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
            Left            =   4680
            TabIndex        =   13
            Top             =   330
            Width           =   1725
         End
         Begin VB.ComboBox cmbRangoMeses 
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
            ItemData        =   "HerrExportaCitasWeb.frx":106BE
            Left            =   1950
            List            =   "HerrExportaCitasWeb.frx":106E6
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   300
            Width           =   2475
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
            ItemData        =   "HerrExportaCitasWeb.frx":1074F
            Left            =   720
            List            =   "HerrExportaCitasWeb.frx":10751
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   300
            Width           =   1215
         End
         Begin UltraGrid.SSUltraGrid grdProgramacionDelMes 
            Height          =   5295
            Left            =   120
            TabIndex        =   4
            Top             =   1050
            Width           =   13665
            _ExtentX        =   24104
            _ExtentY        =   9340
            _Version        =   131072
            GridFlags       =   17040388
            UpdateMode      =   2
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
            Caption         =   "grdAnteriores"
         End
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar1 
            Height          =   300
            Left            =   750
            TabIndex        =   9
            Top             =   660
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   6956042
         End
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar2 
            Height          =   300
            Left            =   2580
            TabIndex        =   10
            Top             =   660
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   6956042
         End
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar3 
            Height          =   300
            Left            =   4710
            TabIndex        =   11
            Top             =   660
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   6956042
         End
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar4 
            Height          =   300
            Left            =   6450
            TabIndex        =   12
            Top             =   660
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   6956042
         End
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar5 
            Height          =   300
            Left            =   8280
            TabIndex        =   16
            Top             =   630
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   6956042
         End
         Begin SISGalenPlus.XP_ProgressBar XP_ProgressBar6 
            Height          =   300
            Left            =   10200
            TabIndex        =   19
            Top             =   660
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BrushStyle      =   0
            Color           =   6956042
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   $"HerrExportaCitasWeb.frx":10753
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
            Height          =   195
            Left            =   90
            TabIndex        =   40
            Top             =   6660
            Width           =   12360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "* Pulse doble clic sobre las X, para asignar CUPO WEB"
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
            Height          =   195
            Left            =   90
            TabIndex        =   21
            Top             =   6405
            Width           =   3870
         End
         Begin VB.Label Label1 
            Caption         =   "Desde"
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
            TabIndex        =   5
            Top             =   360
            Width           =   555
         End
      End
      Begin UltraGrid.SSUltraGrid grdCSreferir 
         Height          =   3825
         Left            =   -74925
         TabIndex        =   23
         Top             =   465
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   6747
         _Version        =   131072
         GridFlags       =   17040388
         UpdateMode      =   2
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
         Caption         =   "Lista de CS/PS"
      End
      Begin UltraGrid.SSUltraGrid grdConsultorios 
         Height          =   6285
         Left            =   -67440
         TabIndex        =   41
         Top             =   375
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   11086
         _Version        =   131072
         GridFlags       =   17040388
         UpdateMode      =   2
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
         Caption         =   "Consultorios"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   $"HerrExportaCitasWeb.frx":107EE
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
         Height          =   195
         Left            =   -74925
         TabIndex        =   46
         Top             =   7320
         Width           =   12360
      End
   End
End
Attribute VB_Name = "HerrExportaCitasWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Configura CITAS WEB
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighEntidades.Teclado
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_Formulario As New sighEntidades.Formulario
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_Procesos As New SIGHProxies.Procesos
Dim oRsCSreferencias As New Recordset
Dim oRsCSreferConsul As New Recordset
Dim oRsProgramacionDelMes As New Recordset
Dim oRsCitasWebPorImportar As New Recordset
Dim lcFechaInicio As String, lcFechaFinal As String, ldFechaHoy As Date
Dim oRsCuposWeb As New Recordset
Dim lcSql As String, lcMesAnio As String
Dim lnMinutosTranscurridos As Integer
Const lcEquix As String = "x"
Dim ml_idUsuario As Long
Dim ml_IdPaciente As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_EsOpcionCitaWebConfigurar As Boolean
Dim lbEstaBloquedaCitasWeb As Boolean
Dim lbEsPrimeraVezQgrabaElMes As Boolean

Property Let EsCitaWebConfigurar(lValue As Boolean)
    ml_EsOpcionCitaWebConfigurar = lValue
End Property

'MARIO
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
'MARIO
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let idUsuario(lIdValue As Long)
    ml_idUsuario = lIdValue
End Property

Private Sub btnAceptar_Click()
    
    If CDate(lcFechaFinal) < CDate("01/" & Format(ldFechaHoy, "mm/yyyy")) Then
       MsgBox "No puede EXPORTAR CITAS menores al mes actual", vbInformation, ""
       Exit Sub
    End If
    If MsgBox("Esta seguro", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       Dim oExportaCitasWeb As New Procesos
       Set oExportaCitasWeb.progressRpt1 = Me.XP_ProgressBar1
       Set oExportaCitasWeb.progressRpt2 = Me.XP_ProgressBar2
       Set oExportaCitasWeb.progressRpt3 = Me.XP_ProgressBar3
       Set oExportaCitasWeb.progressRpt4 = Me.XP_ProgressBar4
       Set oExportaCitasWeb.progressRpt5 = Me.XP_ProgressBar5
       Set oExportaCitasWeb.progressRpt6 = Me.XP_ProgressBar6
       If oRsCuposWeb.RecordCount > 0 Then
          If lbEsPrimeraVezQgrabaElMes = True Then
            If oExportaCitasWeb.ExportaCitasWeb(CDate(lcFechaInicio), CDate(lcFechaFinal), oRsCuposWeb, ml_idUsuario, _
                      ml_IdPaciente, IIf(Me.chkExpServicios.Value = 1, True, False), IIf(Me.chkExpPacientes.Value = 1, True, False), _
                      IIf(Me.chkExpMedicos.Value = 1, True, False), IIf(Me.chkFteFinanciamiento.Value = 1, True, False), _
                      IIf(chkExportaTurnos.Value = 1, True, False), 0) = True Then
                Set oExportaCitasWeb = Nothing
                If CitasWebQuitarBloqueo = False Then   'debb-18/05/2019
            '       MsgBox "Problemas en la WEB, chequee la BD SOMEE para el HOSPITAL", vbInformation, ""
            '       Exit Sub
                End If
                Me.Visible = False
            Else
                Set oExportaCitasWeb = Nothing
            End If
          Else
            If CitasWebQuitarBloqueo = False Then   'debb-18/05/2019
        '       MsgBox "Problemas en la WEB, chequee la BD SOMEE para el HOSPITAL", vbInformation, ""
        '       Exit Sub
            End If
            Me.Visible = False
          End If
       Else
            If oExportaCitasWeb.ExportaCitasWebSoloTablas(IIf(Me.chkExpServicios.Value = 1, True, False), _
                                                          IIf(Me.chkExpPacientes.Value = 1, True, False), _
                                                          IIf(Me.chkExpMedicos.Value = 1, True, False), _
                                                          IIf(Me.chkFteFinanciamiento.Value = 1, True, False), _
                                                          IIf(chkExportaTurnos.Value = 1, True, False)) = True Then
                Set oExportaCitasWeb = Nothing
                Me.Visible = False
            Else
                Set oExportaCitasWeb = Nothing
            End If
       End If
       Me.MousePointer = 1
    End If
End Sub

Private Sub btnAgregarDx_Click()
    mo_Formulario.HabilitarDeshabilitar FraAgregar, True
    btnBuscarEstablecimientoDestino.SetFocus
End Sub

Private Sub btnBuscar_Click()
    If Me.txtBusqApaterno.Text = "" And Me.txtBusqAMaterno.Text = "" Then
       MsgBox "Tiene que ingresar APELLIDO PATERNO completo y/o APELLIDO MATERNO completo", vbInformation, ""
    Else
       Dim lcFiltro As String
       Dim rsRecordset As New Recordset
       lcFiltro = ""
       If Me.txtBusqApaterno.Text <> "" Then
          lcFiltro = lcFiltro & "ApellidoPaterno='" & Me.txtBusqApaterno.Text & "'"
       End If
       If Me.txtBusqAMaterno.Text <> "" Then
          lcFiltro = lcFiltro & IIf(lcFiltro = "", "", " and ") & "ApellidoMaterno='" & Me.txtBusqAMaterno.Text & "'"
       End If
       Set rsRecordset = grdCitasWeb.DataSource
       rsRecordset.Filter = lcFiltro
       Set grdCitasWeb.DataSource = rsRecordset
    End If
End Sub

Private Sub btnBuscarEstablecimientoDestino_Click()
        Dim oBusqueda As New SIGHNegocios.BuscaEstablecimientos
        Dim oDoEstablecimiento As New DOEstablecimiento
        oBusqueda.MostrarFormulario
        If oBusqueda.BotonPresionado = sghAceptar Then
            Set oDoEstablecimiento = mo_ReglasComunes.EstablecimientosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
            If Not oDoEstablecimiento Is Nothing Then
                txtCS.Tag = oDoEstablecimiento.Codigo
                txtCS.Text = oDoEstablecimiento.nombre
            End If
        End If
        Set oBusqueda = Nothing
        Set oDoEstablecimiento = Nothing
        txtCSusuario.SetFocus
End Sub

'debb-18/05/2019
Private Sub btnCancelar_Click()
    If ml_EsOpcionCitaWebConfigurar = True Then
        If CitasWebQuitarBloqueo = False Then
           MsgBox "Problemas en la WEB, chequee la BD SOMEE para el HOSPITAL", vbInformation, ""
        End If
    End If
    Me.Visible = False
End Sub



Private Sub btnEliminaSolicitudCita_Click()
    If CitasWebEstaBloqueda = True Then
       Exit Sub
    End If
    On Error GoTo ErrQu
    Dim rsRecordset As ADODB.Recordset
    Dim mo_DOCuentaAtencion As New DOCuentaAtencion
    Dim oDOCitasWebCupos As New DOCitasWebCupos
    Dim oCitasWebCupos As New CitasWebCupos
    Dim oConexionExterna As New Connection
    Dim mo_Atenciones As New DOAtencion, oAtenciones As New Atenciones
    Dim mo_Pacientes As New doPaciente, oPacientes As New Pacientes
    Dim oConexion As New Connection
    Dim oMensajeCelular As New SIGHProxies.Procesos
    Dim lcMensajeCelular As String
    Dim lcCuentaSeleccionada As String, lcFechaSeleccionada As String, lcIdCitaBloqueada As String
    
    Set rsRecordset = grdCitasWebPorImportar.DataSource
    lcIdCitaBloqueada = rsRecordset("IdCitaBloqueada")
    If MsgBox("Esta seguro de LIBERAR CITA ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        Me.MousePointer = 11
        oConexionExterna.CommandTimeout = 900
        oConexionExterna.CursorLocation = adUseClient
        oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        Set oCitasWebCupos.Conexion = oConexionExterna
        Set oDOCitasWebCupos = oCitasWebCupos.SeleccionarPorIdCitaBloqueada(Val(lcIdCitaBloqueada))
        oDOCitasWebCupos.idEstadoCitaWeb = sghCitaWebEstados.CupoDisponibleEnCitaWeb
        If oCitasWebCupos.Modificar(oDOCitasWebCupos) = True Then
              mo_Procesos.HistoriaWebEliminar rsRecordset!DNI
              
              Dim lcMensaje As String, lbSeTerminaSistema As Boolean
              lcMensaje = ""
              If lcMensaje <> "" Then
                 MsgBox lcMensaje, vbInformation, ""
              Else
                 rsRecordset.Delete
                 rsRecordset.Update
                 grdCitasWebPorImportar.Refresh
                 MsgBox "Se liberó CITA WEB", vbInformation, ""
              End If
        End If
        oConexionExterna.Close
        Me.MousePointer = 1
    End If
ErrQu:
    Set mo_DOCuentaAtencion = Nothing
    Set oDOCitasWebCupos = Nothing
    Set oCitasWebCupos = Nothing
    Set oConexionExterna = Nothing
    Set mo_Atenciones = Nothing
    Set oAtenciones = Nothing
    Set mo_Pacientes = Nothing
    Set oPacientes = Nothing
    Set oConexion = Nothing
    Set oMensajeCelular = Nothing

End Sub

Private Sub btnLimpiar_Click()
    txtBusqApaterno.Text = ""
    txtBusqAMaterno.Text = ""
    '
    Dim rsRecordset As New Recordset
    Set rsRecordset = grdCitasWeb.DataSource
    rsRecordset.Filter = ""
    Set grdCitasWeb.DataSource = rsRecordset
End Sub

Private Sub btnQuitarDx_Click()
    If CitasWebEstaBloqueda = True Then
       Exit Sub
    End If
    On Error GoTo ErrQu
    Dim rsRecordset As ADODB.Recordset
    Dim mo_DOCuentaAtencion As New DOCuentaAtencion
    Dim oDOCitasWebCupos As New DOCitasWebCupos
    Dim oCitasWebCupos As New CitasWebCupos
    Dim oConexionExterna As New Connection
    Dim mo_Atenciones As New DOAtencion, oAtenciones As New Atenciones
    Dim mo_Pacientes As New doPaciente, oPacientes As New Pacientes
    Dim oConexion As New Connection
    Dim oMensajeCelular As New SIGHProxies.Procesos
    Dim lcMensajeCelular As String
    Dim lcCuentaSeleccionada As String, lcFechaSeleccionada As String, lcIdCitaBloqueada As String
    
    Set rsRecordset = grdCitasWeb.DataSource
    lcCuentaSeleccionada = rsRecordset("nCuenta")
    lcFechaSeleccionada = rsRecordset("fecha")
    lcIdCitaBloqueada = rsRecordset("IdCitaBloqueada")
    If ldFechaHoy <= CDate(lcFechaSeleccionada) Then
        If MsgBox("Esta seguro de LIBERAR CITA ?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
            Me.MousePointer = 11
            oConexion.CommandTimeout = 900
            oConexion.CursorLocation = adUseClient
            oConexion.Open sighEntidades.CadenaConexion
            oConexionExterna.CommandTimeout = 900
            oConexionExterna.CursorLocation = adUseClient
            oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
            Set oCitasWebCupos.Conexion = oConexionExterna
            Set oAtenciones.Conexion = oConexion
            Set oPacientes.Conexion = oConexion
            mo_DOCuentaAtencion.idCuentaAtencion = Val(lcCuentaSeleccionada)
            mo_DOCuentaAtencion.IdUsuarioAuditoria = sighEntidades.Usuario
            mo_DOCuentaAtencion.TotalPorPagar = 0
            If mo_ReglasFacturacion.CuentasAtencionAnulada(mo_DOCuentaAtencion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") = True Then
               mo_ReglasFacturacion.FacturacionCuentasAtencionPtosActualizar Val(lcCuentaSeleccionada), False, 0
               Set oDOCitasWebCupos = oCitasWebCupos.SeleccionarPorIdCitaBloqueada(Val(lcIdCitaBloqueada))
               If oDOCitasWebCupos.idWeb > 0 Then
                  oDOCitasWebCupos.idEstadoCitaWeb = sghCitaWebEstados.CupoANULADO
                  If oCitasWebCupos.Eliminar(oDOCitasWebCupos) = True Then
                     mo_Pacientes.idPaciente = oDOCitasWebCupos.idPaciente
                     If oPacientes.SeleccionarPorId(mo_Pacientes) = True Then
                        mo_Atenciones.idCuentaAtencion = Val(lcCuentaSeleccionada)
                        If oAtenciones.SeleccionarPorIdCuentaDeAtencion(mo_Atenciones) = True Then
                            lcMensajeCelular = "La CITA para " & mo_Atenciones.FechaIngreso & " " & mo_Atenciones.HoraIngreso & _
                                       " (N° Cuenta: " & mo_Atenciones.idCuentaAtencion & ") " & _
                                       " (a sido ANULADA, comuniquese o vaya al " & wxParametro205
                            oMensajeCelular.MensajeCelularEnviar mo_Pacientes, mo_Atenciones.idCuentaAtencion, lcMensajeCelular, _
                                    "EXPORTACITAWEB", oConexion
                        End If
                     End If
                     Dim lcMensaje As String, lbSeTerminaSistema As Boolean
                     lcMensaje = ""
                     If lcMensaje <> "" Then
                       MsgBox lcMensaje, vbInformation, ""
                     End If
                     
                     mo_Procesos.HistoriaWebEliminar rsRecordset!DNI
                  End If
               End If
                
            Else
                MsgBox "No se pudo ANULAR EL CUPO", vbInformation, "Facturación"
            End If
            oConexionExterna.Close
            oConexion.Close
            Me.MousePointer = 1
            Me.Visible = False
        End If
    End If
ErrQu:
    Set mo_DOCuentaAtencion = Nothing
    Set oDOCitasWebCupos = Nothing
    Set oCitasWebCupos = Nothing
    Set oConexionExterna = Nothing
    Set mo_Atenciones = Nothing
    Set oAtenciones = Nothing
    Set mo_Pacientes = Nothing
    Set oPacientes = Nothing
    Set oConexion = Nothing
    Set oMensajeCelular = Nothing
End Sub

Private Sub chkSeImportacionAutomatica_Click()
    ManualOautomatica
End Sub

Private Sub chkTodos_Click()
    On Error Resume Next
    oRsCitasWebPorImportar.MoveFirst
    Do While Not oRsCitasWebPorImportar.EOF
        If chkTodos.Value = 1 Then
           oRsCitasWebPorImportar!seleccionar = True
        Else
           oRsCitasWebPorImportar!seleccionar = False
        End If
        oRsCitasWebPorImportar.Update
        oRsCitasWebPorImportar.MoveNext
    Loop
End Sub

Private Sub cmbRangoMeses_Click()
    If cmbRangoMeses.ListIndex >= 0 Then
       lbEsPrimeraVezQgrabaElMes = True
       CargaCitasWebSolicitadas           'debb-18/05/2019
       CargaDatosDeCuposWeb
       'Busca Programacion del Mes y Año
       Dim lnIdServicio As Long, lcServicio As String, lcDNIcitaWeb As String
       Dim oRsTmp1 As New Recordset
       Dim lcDia1 As String, lcDia2 As String, lcDia3 As String, lcDia4 As String, lcDia5 As String
       Dim lcDia6 As String, lcDia7 As String, lcDia8 As String, lcDia9 As String, lcDia10 As String
       Dim lcDia11 As String, lcDia12 As String, lcDia13 As String, lcDia14 As String, lcDia15 As String
       Dim lcDia16 As String, lcDia17 As String, lcDia18 As String, lcDia19 As String, lcDia20 As String
       Dim lcDia21 As String, lcDia22 As String, lcDia23 As String, lcDia24 As String, lcDia25 As String
       Dim lcDia26 As String, lcDia27 As String, lcDia28 As String, lcDia29 As String, lcDia30 As String, lcDia31 As String
       
       lcMesAnio = "/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text
       lcFechaInicio = "01/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text
       lcFechaFinal = sighEntidades.DevuelveUltimoDiaDelMes(Me.cmbRangoMeses.ListIndex + 1, Val(Me.cmbAnio.Text)) & _
                       "/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text
       Set oRsTmp1 = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarPorFechas(CDate(lcFechaInicio), CDate(lcFechaFinal))
       If oRsTmp1.RecordCount = 0 Then
          MsgBox "Aún no hay programación Médica", vbInformation, Me.Caption
       Else
          oRsTmp1.MoveFirst
          Do While Not oRsTmp1.EOF
             lnIdServicio = oRsTmp1.Fields!IdServicio
             lcServicio = oRsTmp1.Fields!nombre
             lcDia1 = "": lcDia2 = "": lcDia3 = "": lcDia4 = "": lcDia5 = ""
             lcDia6 = "": lcDia7 = "": lcDia8 = "": lcDia9 = "": lcDia10 = ""
             lcDia11 = "": lcDia12 = "": lcDia13 = "": lcDia14 = "": lcDia15 = ""
             lcDia16 = "": lcDia17 = "": lcDia18 = "": lcDia19 = "": lcDia20 = ""
             lcDia21 = "": lcDia22 = "": lcDia23 = "": lcDia24 = "": lcDia25 = ""
             lcDia26 = "": lcDia27 = "": lcDia28 = "": lcDia29 = "": lcDia30 = "": lcDia31 = ""
             Do While Not oRsTmp1.EOF And lnIdServicio = oRsTmp1.Fields!IdServicio
                Select Case Day(oRsTmp1.Fields!fecha)
                Case 1
                     lcDia1 = lcEquix
                Case 2
                     lcDia2 = lcEquix
                Case 3
                     lcDia3 = lcEquix
                Case 4
                     lcDia4 = lcEquix
                Case 5
                     lcDia5 = lcEquix
                Case 6
                     lcDia6 = lcEquix
                Case 7
                     lcDia7 = lcEquix
                Case 8
                     lcDia8 = lcEquix
                Case 9
                     lcDia9 = lcEquix
                Case 10
                     lcDia10 = lcEquix
                Case 11
                     lcDia11 = lcEquix
                Case 12
                     lcDia12 = lcEquix
                Case 13
                     lcDia13 = lcEquix
                Case 14
                     lcDia14 = lcEquix
                Case 15
                     lcDia15 = lcEquix
                Case 16
                     lcDia16 = lcEquix
                Case 17
                     lcDia17 = lcEquix
                Case 18
                     lcDia18 = lcEquix
                Case 19
                     lcDia19 = lcEquix
                Case 20
                     lcDia20 = lcEquix
                Case 21
                     lcDia21 = lcEquix
                Case 22
                     lcDia22 = lcEquix
                Case 23
                     lcDia23 = lcEquix
                Case 24
                     lcDia24 = lcEquix
                Case 25
                     lcDia25 = lcEquix
                Case 26
                     lcDia26 = lcEquix
                Case 27
                     lcDia27 = lcEquix
                Case 28
                     lcDia28 = lcEquix
                Case 29
                     lcDia29 = lcEquix
                Case 30
                     lcDia30 = lcEquix
                Case 31
                     lcDia31 = lcEquix
                End Select
                oRsTmp1.MoveNext
                If oRsTmp1.EOF Then
                   Exit Do
                End If
             Loop
             oRsProgramacionDelMes.AddNew
             oRsProgramacionDelMes.Fields!IdServicio = lnIdServicio
             oRsProgramacionDelMes.Fields!Servicio = lcServicio
             oRsProgramacionDelMes.Fields!dia1 = lcDia1
             oRsProgramacionDelMes.Fields!dia2 = lcDia2
             oRsProgramacionDelMes.Fields!dia3 = lcDia3
             oRsProgramacionDelMes.Fields!dia4 = lcDia4
             oRsProgramacionDelMes.Fields!dia5 = lcDia5
             oRsProgramacionDelMes.Fields!dia6 = lcDia6
             oRsProgramacionDelMes.Fields!dia7 = lcDia7
             oRsProgramacionDelMes.Fields!dia8 = lcDia8
             oRsProgramacionDelMes.Fields!dia9 = lcDia9
             oRsProgramacionDelMes.Fields!dia10 = lcDia10
             oRsProgramacionDelMes.Fields!dia11 = lcDia11
             oRsProgramacionDelMes.Fields!dia12 = lcDia12
             oRsProgramacionDelMes.Fields!dia13 = lcDia13
             oRsProgramacionDelMes.Fields!dia14 = lcDia14
             oRsProgramacionDelMes.Fields!dia15 = lcDia15
             oRsProgramacionDelMes.Fields!dia16 = lcDia16
             oRsProgramacionDelMes.Fields!dia17 = lcDia17
             oRsProgramacionDelMes.Fields!dia18 = lcDia18
             oRsProgramacionDelMes.Fields!dia19 = lcDia19
             oRsProgramacionDelMes.Fields!dia20 = lcDia20
             oRsProgramacionDelMes.Fields!dia21 = lcDia21
             oRsProgramacionDelMes.Fields!dia22 = lcDia22
             oRsProgramacionDelMes.Fields!dia23 = lcDia23
             oRsProgramacionDelMes.Fields!dia24 = lcDia24
             oRsProgramacionDelMes.Fields!dia25 = lcDia25
             oRsProgramacionDelMes.Fields!dia26 = lcDia26
             oRsProgramacionDelMes.Fields!dia27 = lcDia27
             oRsProgramacionDelMes.Fields!dia28 = lcDia28
             oRsProgramacionDelMes.Fields!dia29 = lcDia29
             oRsProgramacionDelMes.Fields!dia30 = lcDia30
             oRsProgramacionDelMes.Fields!dia31 = lcDia31
             oRsProgramacionDelMes.Update
          Loop
          oRsProgramacionDelMes.Sort = "Servicio"
          '
          Dim oRsCuposWebSeleccionado As New Recordset
          Set oRsCuposWebSeleccionado = mo_ReglasDeProgMedica.CitasWebCuposSeleccionarPorFechas(CDate(lcFechaInicio), CDate(lcFechaFinal))
          If oRsCuposWebSeleccionado.RecordCount > 0 Then
             lbEsPrimeraVezQgrabaElMes = False
             oRsCuposWebSeleccionado.MoveFirst
             Do While Not oRsCuposWebSeleccionado.EOF
                 'Actualiza temporal con CITAS YA SOLICITADAS EN LA WEB  debb-18/05/2019
                 lcDNIcitaWeb = ""
                 If Not IsNull(oRsCuposWebSeleccionado.Fields!idCitaBloqueada) Then
                    oRsCitasWebPorImportar.Filter = "idCitaBloqueada=" & oRsCuposWebSeleccionado.Fields!idCitaBloqueada
                    If oRsCitasWebPorImportar.RecordCount > 0 Then
                       lcDNIcitaWeb = oRsCitasWebPorImportar!DNI
                    End If
                 End If
                 '
                 oRsCuposWeb.AddNew
                 oRsCuposWeb.Fields!HoraInicio = oRsCuposWebSeleccionado.Fields!HoraInicio
                 oRsCuposWeb.Fields!HoraFinal = oRsCuposWebSeleccionado.Fields!HoraFinal
                 oRsCuposWeb.Fields!IdServicio = oRsCuposWebSeleccionado.Fields!IdServicio
                 oRsCuposWeb.Fields!fecha = oRsCuposWebSeleccionado.Fields!fecha
                 oRsCuposWeb.Fields!idMedico = oRsCuposWebSeleccionado.Fields!idMedico
                 oRsCuposWeb.Fields!idEstadoCitaWeb = IIf(lcDNIcitaWeb <> "", sghCitaWebEstados.CupoConfirmadoEnCitaWeb, _
                                                      oRsCuposWebSeleccionado.Fields!idEstadoCitaWeb)    'debb-18-05/2019
                 oRsCuposWeb.Fields!idCitaBloqueada = oRsCuposWebSeleccionado.Fields!idCitaBloqueada
                 oRsCuposWeb.Fields!DNI = IIf(IsNull(oRsCuposWebSeleccionado.Fields!DNI), lcDNIcitaWeb, _
                                          oRsCuposWebSeleccionado.Fields!DNI)   'debb-18/05/2019
                 oRsCuposWeb.Fields!ApellidoPaterno = IIf(IsNull(oRsCuposWebSeleccionado.Fields!ApellidoPaterno), "", oRsCuposWebSeleccionado.Fields!ApellidoPaterno)
                 oRsCuposWeb.Fields!ApellidoMaterno = IIf(IsNull(oRsCuposWebSeleccionado.Fields!ApellidoMaterno), "", oRsCuposWebSeleccionado.Fields!ApellidoMaterno)
                 oRsCuposWeb.Fields!PrimerNombre = IIf(IsNull(oRsCuposWebSeleccionado.Fields!PrimerNombre), "", oRsCuposWebSeleccionado.Fields!PrimerNombre)
                 oRsCuposWeb.Fields!SegundoNombre = IIf(IsNull(oRsCuposWebSeleccionado.Fields!SegundoNombre), "", oRsCuposWebSeleccionado.Fields!SegundoNombre)
                 oRsCuposWeb.Fields!idTipoSexo = IIf(IsNull(oRsCuposWebSeleccionado.Fields!idTipoSexo), 0, oRsCuposWebSeleccionado.Fields!idTipoSexo)
                 oRsCuposWeb.Fields!FechaNacimiento = IIf(IsNull(oRsCuposWebSeleccionado.Fields!FechaNacimiento), 0, oRsCuposWebSeleccionado.Fields!FechaNacimiento)
                 oRsCuposWeb.Fields!IdTurno = oRsCuposWebSeleccionado.Fields!IdTurno
                 oRsCuposWeb.Fields!idPaciente = oRsCuposWebSeleccionado.Fields!idPaciente
                 oRsCuposWeb.Update
                 If oRsCuposWebSeleccionado.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoDisponibleEnCitaWeb Then
                    oRsProgramacionDelMes.MoveFirst
                    oRsProgramacionDelMes.Filter = "idServicio=" & oRsCuposWebSeleccionado.Fields!IdServicio
                    If Not oRsProgramacionDelMes.EOF Then
                        Select Case Day(oRsCuposWebSeleccionado.Fields!fecha)
                        Case 1
                             If Val(oRsProgramacionDelMes.Fields!dia1) > 0 Then
                                oRsProgramacionDelMes.Fields!dia1 = oRsProgramacionDelMes.Fields!dia1 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia1 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 2
                             If Val(oRsProgramacionDelMes.Fields!dia2) > 0 Then
                                oRsProgramacionDelMes.Fields!dia2 = oRsProgramacionDelMes.Fields!dia2 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia2 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 3
                             If Val(oRsProgramacionDelMes.Fields!dia3) > 0 Then
                                oRsProgramacionDelMes.Fields!dia3 = oRsProgramacionDelMes.Fields!dia3 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia3 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 4
                             If Val(oRsProgramacionDelMes.Fields!dia4) > 0 Then
                                oRsProgramacionDelMes.Fields!dia4 = oRsProgramacionDelMes.Fields!dia4 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia4 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 5
                             If Val(oRsProgramacionDelMes.Fields!dia5) > 0 Then
                                oRsProgramacionDelMes.Fields!dia5 = oRsProgramacionDelMes.Fields!dia5 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia5 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 6
                             If Val(oRsProgramacionDelMes.Fields!dia6) > 0 Then
                                oRsProgramacionDelMes.Fields!dia6 = oRsProgramacionDelMes.Fields!dia6 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia6 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 7
                             If Val(oRsProgramacionDelMes.Fields!dia7) > 0 Then
                                oRsProgramacionDelMes.Fields!dia7 = oRsProgramacionDelMes.Fields!dia7 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia7 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 8
                             If Val(oRsProgramacionDelMes.Fields!dia8) > 0 Then
                                oRsProgramacionDelMes.Fields!dia8 = oRsProgramacionDelMes.Fields!dia8 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia8 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 9
                             If Val(oRsProgramacionDelMes.Fields!dia9) > 0 Then
                                oRsProgramacionDelMes.Fields!dia9 = oRsProgramacionDelMes.Fields!dia9 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia9 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 10
                             If Val(oRsProgramacionDelMes.Fields!dia10) > 0 Then
                                oRsProgramacionDelMes.Fields!dia10 = oRsProgramacionDelMes.Fields!dia10 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia10 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 11
                             If Val(oRsProgramacionDelMes.Fields!dia11) > 0 Then
                                oRsProgramacionDelMes.Fields!dia11 = oRsProgramacionDelMes.Fields!dia11 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia11 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 12
                             If Val(oRsProgramacionDelMes.Fields!dia12) > 0 Then
                                oRsProgramacionDelMes.Fields!dia12 = oRsProgramacionDelMes.Fields!dia12 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia12 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 13
                             If Val(oRsProgramacionDelMes.Fields!dia13) > 0 Then
                                oRsProgramacionDelMes.Fields!dia13 = oRsProgramacionDelMes.Fields!dia13 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia13 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 14
                             If Val(oRsProgramacionDelMes.Fields!dia14) > 0 Then
                                oRsProgramacionDelMes.Fields!dia14 = oRsProgramacionDelMes.Fields!dia14 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia14 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 15
                             If Val(oRsProgramacionDelMes.Fields!dia15) > 0 Then
                                oRsProgramacionDelMes.Fields!dia15 = oRsProgramacionDelMes.Fields!dia15 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia15 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 16
                             If Val(oRsProgramacionDelMes.Fields!dia16) > 0 Then
                                oRsProgramacionDelMes.Fields!dia16 = oRsProgramacionDelMes.Fields!dia16 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia16 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 17
                             If Val(oRsProgramacionDelMes.Fields!dia17) > 0 Then
                                oRsProgramacionDelMes.Fields!dia17 = oRsProgramacionDelMes.Fields!dia17 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia17 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 18
                             If Val(oRsProgramacionDelMes.Fields!dia18) > 0 Then
                                oRsProgramacionDelMes.Fields!dia18 = oRsProgramacionDelMes.Fields!dia18 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia18 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 19
                             If Val(oRsProgramacionDelMes.Fields!dia19) > 0 Then
                                oRsProgramacionDelMes.Fields!dia19 = oRsProgramacionDelMes.Fields!dia19 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia19 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 20
                             If Val(oRsProgramacionDelMes.Fields!dia20) > 0 Then
                                oRsProgramacionDelMes.Fields!dia20 = oRsProgramacionDelMes.Fields!dia20 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia20 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 21
                             If Val(oRsProgramacionDelMes.Fields!dia21) > 0 Then
                                oRsProgramacionDelMes.Fields!dia21 = oRsProgramacionDelMes.Fields!dia21 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia21 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 22
                             If Val(oRsProgramacionDelMes.Fields!dia22) > 0 Then
                                oRsProgramacionDelMes.Fields!dia22 = oRsProgramacionDelMes.Fields!dia22 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia22 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 23
                             If Val(oRsProgramacionDelMes.Fields!dia23) > 0 Then
                                oRsProgramacionDelMes.Fields!dia23 = oRsProgramacionDelMes.Fields!dia23 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia23 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 24
                             If Val(oRsProgramacionDelMes.Fields!dia24) > 0 Then
                                oRsProgramacionDelMes.Fields!dia24 = oRsProgramacionDelMes.Fields!dia24 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia24 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 25
                             If Val(oRsProgramacionDelMes.Fields!dia25) > 0 Then
                                oRsProgramacionDelMes.Fields!dia25 = oRsProgramacionDelMes.Fields!dia25 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia25 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 26
                             If Val(oRsProgramacionDelMes.Fields!dia26) > 0 Then
                                oRsProgramacionDelMes.Fields!dia26 = oRsProgramacionDelMes.Fields!dia26 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia26 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 27
                             If Val(oRsProgramacionDelMes.Fields!dia27) > 0 Then
                                oRsProgramacionDelMes.Fields!dia27 = oRsProgramacionDelMes.Fields!dia27 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia27 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 28
                             If Val(oRsProgramacionDelMes.Fields!dia28) > 0 Then
                                oRsProgramacionDelMes.Fields!dia28 = oRsProgramacionDelMes.Fields!dia28 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia28 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 29
                             If Val(oRsProgramacionDelMes.Fields!dia29) > 0 Then
                                oRsProgramacionDelMes.Fields!dia29 = oRsProgramacionDelMes.Fields!dia29 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia29 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 30
                             If Val(oRsProgramacionDelMes.Fields!dia30) > 0 Then
                                oRsProgramacionDelMes.Fields!dia30 = oRsProgramacionDelMes.Fields!dia30 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia30 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        Case 31
                             If Val(oRsProgramacionDelMes.Fields!dia31) > 0 Then
                                oRsProgramacionDelMes.Fields!dia31 = oRsProgramacionDelMes.Fields!dia31 + 1
                             Else
                                oRsProgramacionDelMes.Fields!dia31 = 1
                             End If
                             oRsProgramacionDelMes.Update
                        End Select
                   End If
                 End If
                 oRsCuposWebSeleccionado.MoveNext
             Loop
             oRsProgramacionDelMes.Filter = ""
          End If
          Set oRsCuposWebSeleccionado = Nothing
       End If
       oRsTmp1.Close
       Set oRsTmp1 = Nothing
    End If

End Sub


Private Sub cmbRangoMeses_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbRangoMeses_Click
    End If
End Sub


Sub LimpiarDatos()
    txtCS.Text = ""
    txtCSusuario.Text = ""
    txtCSclave.Text = ""
    txtEmail.Text = ""
End Sub

Function ExisteUsuarioClave() As Boolean
       On Error GoTo ErrCargCS1
       ExisteUsuarioClave = False
       If oRsCSreferencias.RecordCount > 0 Then
          oRsCSreferencias.MoveFirst
          Do While Not oRsCSreferencias.EOF
               If UCase(oRsCSreferencias!Usuario) = UCase(txtCSusuario.Text) And UCase(oRsCSreferencias!Clave) = UCase(txtCSclave.Text) Then
                  ExisteUsuarioClave = True
                  MsgBox "Ya existe ese USUARIO y CLAVE ", vbInformation, ""
                  Exit Function
               End If
                oRsCSreferencias.MoveNext
          Loop
       End If
       Dim lcMensaje As String, lbSeTerminaSistema As Boolean
       Dim oRsTmp1 As New Recordset
       lcMensaje = ""
       oRsTmp1.Filter = "eess<>" & lcBuscaParametro.SeleccionaFilaParametro(280)
       If oRsTmp1.RecordCount > 0 Then
          ExisteUsuarioClave = True
          MsgBox "Ya existe ese USUARIO y CLAVE en otro HOSPITAL que usa CITAS WEB", vbInformation, ""
       End If
       oRsTmp1.Close
ErrCargCS1:
       Set oRsTmp1 = Nothing
End Function

Private Sub cmdAdicionaConsultoriosNuevos_Click()
    On Error GoTo cmdElm
    
    Dim oRsTmp2 As New Recordset
    If MsgBox("Agrega los NUEVOS CONSULTORIOS", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
            Me.MousePointer = 11
            Dim lbAgrega As Boolean
            Set oRsTmp2 = mo_ReglasFacturacion.ServiciosSeleccionarPorFiltro("idTipoServicio=1", sghPorCodigo)
            If oRsTmp2.RecordCount > 0 Then
               oRsTmp2.MoveFirst
               Do While Not oRsTmp2.EOF
                  oRsCSreferencias.MoveFirst
                  Do While Not oRsCSreferencias.EOF
                     lbAgrega = True
                     lcSql = "refEess='" & oRsCSreferencias!EessCodigo & "' and idServicio=" & oRsTmp2!IdServicio
                     oRsCSreferConsul.Filter = lcSql
                     If oRsCSreferConsul.RecordCount > 0 Then
                       lbAgrega = False
                     End If
                     If lbAgrega = True Then
                        oRsCSreferConsul.AddNew
                        oRsCSreferConsul!Eess = oRsCSreferencias!Eess
                        oRsCSreferConsul!refEess = oRsCSreferencias!EessCodigo
                        oRsCSreferConsul!Servicio = Left(oRsTmp2!nombre, 100)
                        oRsCSreferConsul!IdServicio = oRsTmp2!IdServicio
                        oRsCSreferConsul!maximoCupos = 0
                        oRsCSreferConsul.Update
                     End If
                     oRsCSreferencias.MoveNext
                  Loop
                  oRsTmp2.MoveNext
               Loop
               oRsCSreferConsul.Filter = ""
            End If
            MsgBox "Ya agregó nuevos Consultorios", vbInformation, ""
       End If
cmdElm:
       Set oRsTmp2 = Nothing
       Me.MousePointer = 1
End Sub

Private Sub cmdAdicionar_Click()
    If txtCS.Text = "" Then
       MsgBox "Tiene que elegir el CS o PS", vbInformation, ""
       Exit Sub
    End If
    If txtCSusuario.Text = "" Then
       MsgBox "Tiene que registrar el USUARIO", vbInformation, ""
       Exit Sub
    End If
    If txtCSclave.Text = "" Then
       MsgBox "Tiene que regsitrar la CLAVE", vbInformation, ""
       Exit Sub
    End If
    If txtEmail.Text <> "" Then
       If InStr(txtEmail.Text, "@") = 0 Then
            MsgBox "Debe tener   @    en el EMAIL", vbInformation, ""
            Exit Sub
       End If
    End If
    If oRsCSreferencias.RecordCount > 0 Then
       oRsCSreferencias.MoveFirst
       oRsCSreferencias.Find "eessCodigo='" & txtCS.Tag & "'"
       If Not oRsCSreferencias.EOF Then
          MsgBox "Ya está registrado ese EESS", vbInformation, ""
          Exit Sub
       End If
    End If
    If ExisteUsuarioClave = True Then
       Exit Sub
    End If
    Me.MousePointer = 11
    
    oRsCSreferencias.AddNew
    oRsCSreferencias!EessCodigo = txtCS.Tag
    oRsCSreferencias!Eess = Left(txtCS.Text, 100)
    oRsCSreferencias!Usuario = txtCSusuario.Text
    oRsCSreferencias!Clave = txtCSclave.Text
    oRsCSreferencias!Email = txtEmail.Text
    oRsCSreferencias.Update
    mo_Formulario.HabilitarDeshabilitar FraAgregar, False
    
    Dim oRsTmp2 As New Recordset
    Set oRsTmp2 = mo_ReglasFacturacion.ServiciosSeleccionarPorFiltro("idTipoServicio=1", sghPorDescripcion)
    If oRsTmp2.RecordCount > 0 Then
       oRsTmp2.MoveFirst
       Do While Not oRsTmp2.EOF
            oRsCSreferConsul.AddNew
            oRsCSreferConsul!Eess = Left(txtCS.Text, 100)
            oRsCSreferConsul!refEess = txtCS.Tag
            oRsCSreferConsul!Servicio = Left(oRsTmp2!nombre, 100)
            oRsCSreferConsul!IdServicio = oRsTmp2!IdServicio
            oRsCSreferConsul!maximoCupos = 0
            oRsCSreferConsul.Update
            oRsTmp2.MoveNext
       Loop
    End If
    oRsTmp2.Close
    Set oRsTmp2 = Nothing
    
    oRsCSreferConsul.MoveFirst
    Set grdConsultorios.DataSource = oRsCSreferConsul
    
    LimpiarDatos
    Me.MousePointer = 1
End Sub

Private Sub cmdCancelar_Click()
    LimpiarDatos
    mo_Formulario.HabilitarDeshabilitar FraAgregar, False
End Sub

'debb-18/05/2019
Private Sub cmdCancelaReferidos_Click()
    If ml_EsOpcionCitaWebConfigurar = True Then
        If CitasWebQuitarBloqueo = False Then
           MsgBox "Problemas en la WEB, chequee la BD SOMEE para el HOSPITAL", vbInformation, ""
        End If
    End If
    Me.Visible = False
End Sub

Private Sub cmdEliminar_Click()
    On Error GoTo cmdElm
    If MsgBox("Esta seguro", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        oRsCSreferConsul.Filter = "refEess='" & oRsCSreferencias!EessCodigo & "'"
        If oRsCSreferConsul.RecordCount > 0 Then
           oRsCSreferConsul.MoveFirst
           Do While Not oRsCSreferConsul.EOF
              oRsCSreferConsul.Delete
              oRsCSreferConsul.Update
              oRsCSreferConsul.MoveNext
           Loop
        End If
        oRsCSreferConsul.Filter = ""
        oRsCSreferencias.Delete
        oRsCSreferencias.Update
        Set grdConsultorios.DataSource = oRsCSreferConsul
        Set grdCSreferir.DataSource = oRsCSreferencias
    End If
cmdElm:
End Sub

Private Sub cmdGrabaReferidos_Click()
        
        Dim lcMensaje As String, lbSeTerminaSistema As Boolean
        lcMensaje = ""
        If lcMensaje <> "" Then
           MsgBox lcMensaje, vbInformation, ""
        End If
        
        oRsCSreferConsul.Filter = ""
        lcMensaje = ""
        If lcMensaje <> "" Then
           MsgBox lcMensaje, vbInformation, ""
        End If
        If CitasWebQuitarBloqueo = False Then   'debb-18/05/2019
           'MsgBox "Problemas en la WEB, chequee la BD SOMEE para el HOSPITAL", vbInformation, ""
           'exit sub
        End If
        Me.Visible = False
End Sub
Sub EliminaHistoriasQueSeCrearon(lbDesdeBotonAceptar As Boolean)
       If lbDesdeBotonAceptar = True Then
          oRsCitasWebPorImportar.Filter = "seleccionar=false"
       Else
          oRsCitasWebPorImportar.Filter = ""
       End If
       If oRsCitasWebPorImportar.RecordCount > 0 Then
          oRsCitasWebPorImportar.MoveFirst
          Do While Not oRsCitasWebPorImportar.EOF
             mo_Procesos.HistoriaWebEliminar oRsCitasWebPorImportar!DNI
             oRsCitasWebPorImportar.MoveNext
          Loop
       End If
       oRsCitasWebPorImportar.Filter = ""

End Sub
Private Sub cmdImportarWeb_Click()
    If CitasWebEstaBloqueda = True Then   'debb-18/05/2019
       Exit Sub
    End If
    
    If MsgBox("Esta seguro", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Me.MousePointer = 11
       Dim oImportaCitasWeb As New Procesos
       Set oImportaCitasWeb.progressRpt1 = Me.XP_ProgressBar7
       oImportaCitasWeb.idUsuario = ml_idUsuario
       oImportaCitasWeb.lnIdTablaLISTBARITEMS = 500 + 183
       If oImportaCitasWeb.ImportaCitasWeb(oRsCuposWeb, Me.txtArchivoCuposConfirmados.Text, Me.optSisGalenPlus.Value, _
                                           Me.hwnd, IIf(Me.chkSeImportacionAutomatica.Value = 0, True, False), _
                                           oRsCitasWebPorImportar) Then
       End If
       Set oImportaCitasWeb = Nothing
       
       EliminaHistoriasQueSeCrearon True
       'enviar EMAIL a CS/PS con las CITAS ya GENERADAS y sin problemas
       On Error GoTo errEmail
       Dim oRsEmailMensajes As New Recordset
       Dim oConexion As New Connection
       Dim oImprimeTicketCita As New LoginActualizaClave
       Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
       lcParametro524 = lcBuscaParametro.SeleccionaFilaParametro(524)
       lcParametro523 = lcBuscaParametro.SeleccionaFilaParametro(523)
       sighEntidades.AbreConexionSIGH oConexion
       lcSql = "select * from Reporte_cabecera where idUsuario=" & sighEntidades.Usuario
       oRsEmailMensajes.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       oRsEmailMensajes.Filter = "nroCuenta=1"
       If oRsEmailMensajes.RecordCount > 0 Then
           oRsEmailMensajes.MoveFirst
           Do While Not oRsEmailMensajes.EOF
              mo_AdminAdmision.CitasFormaCitaActualiza oRsEmailMensajes!IdServicio, "W", oConexion
              oRsEmailMensajes.MoveNext
           Loop
           oRsEmailMensajes.MoveFirst
           Do While Not oRsEmailMensajes.EOF
                 oImprimeTicketCita.idUsuario = sighEntidades.Usuario
                 oImprimeTicketCita.ImprimeCuenta = True
                 oImprimeTicketCita.CuentaDesdeOtroFormulario = oRsEmailMensajes!IdServicio      'N cuenta
                 oImprimeTicketCita.Show 1
                 mo_ReglasComunes.WaitSeconds 5
                 EnviaEmailPacientesOKEY lcParametro524, lcParametro523, _
                                         "Cta:" & Trim(Str(oRsEmailMensajes!IdServicio)) & " " & oRsEmailMensajes!Estancia, _
                                         App.Path & "\" & Trim(Str(oRsEmailMensajes!IdServicio)) & ".pdf", _
                                         oRsEmailMensajes!Motivo, oRsEmailMensajes!destino
              oRsEmailMensajes.MoveNext
           Loop
       End If
       oConexion.Close
       Set oRsEmailMensajes = Nothing
       Set oConexion = Nothing
       Set oImprimeTicketCita = Nothing
       Set mo_AdminAdmision = Nothing
       '
       Me.MousePointer = 1
       Me.Visible = False
    End If
    Exit Sub
errEmail:
    Me.Visible = False
    Set oRsEmailMensajes = Nothing
    Set oConexion = Nothing
    Set oImprimeTicketCita = Nothing
    Set mo_AdminAdmision = Nothing
End Sub

Sub EnviaEmailPacientesOKEY(lcParametro524 As String, lcParametro523 As String, lcAsunto As String, lcArchivoPDF As String, _
                            lcEmailDestino As String, lcMensajeDetalle As String)
        Dim mo_Procesos As New SIGHProxies.Procesos
        mo_Procesos.EnviaEmail lcParametro524, lcParametro523, lcAsunto, _
                            lcArchivoPDF, _
                            lcEmailDestino, lcMensajeDetalle
       Set mo_Procesos = Nothing
End Sub


Private Sub cmdSalir_Click()
    If MsgBox("Esta seguro", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       EliminaHistoriasQueSeCrearon False
       Me.Visible = False
    End If
End Sub






'debb-18/05/2019
Private Sub Form_Activate()
       If ml_EsOpcionCitaWebConfigurar = True Then
          If lbEstaBloquedaCitasWeb = False Then
             If CitasWebGenerarBloqueo = True Then
                lbEstaBloquedaCitasWeb = True
             End If
          End If
       Else
          If CitasWebEstaBloqueda = True Then
             Me.Visible = False
          End If
       End If
End Sub

Private Sub Form_Load()
       ldFechaHoy = lcBuscaParametro.RetornaFechaServidorSQL
       mo_Formulario.LlenaComboConAnios cmbAnio
       mo_Formulario.HabilitarDeshabilitar Frame2, False
       Me.chkSeImportacionAutomatica.Value = IIf(sighEntidades.Parametro503valorInt = "1", 1, 0)
       CargaDatosDeCuposWeb
       
       'debb-18/05/2019
       If ml_EsOpcionCitaWebConfigurar = True Then
          SSTab1.TabVisible(0) = False
          CargaCSreferencias
          Me.Caption = "Configura Citas Web y datos de los Cs/Ps que harán la Referencia"
       Else
          SSTab1.TabVisible(1) = False
          SSTab1.TabVisible(2) = False
          CargaCitasWEBYaAsignadasEnGalenhos
          ManualOautomatica
          Me.Caption = "Importa Citas WEB solicitadas y les crea Cita en el Hospital"
       End If
       
       
End Sub

'debb-18/05/2019
Function CitasWebGenerarBloqueo() As Boolean
    CitasWebGenerarBloqueo = False
    Dim mo_Procesos As New SIGHProxies.Procesos
    Dim lcMensaje As String, lbSeTerminaSistema As Boolean, oRsTmp1 As New Recordset
    lcMensaje = ""
    Set mo_Procesos = Nothing
    Set oRsTmp1 = Nothing
    If lcMensaje <> "" Then
       MsgBox lcMensaje, vbInformation, ""
    Else
       CitasWebGenerarBloqueo = True
    End If
End Function
'debb-18/05/2019
Function CitasWebQuitarBloqueo() As Boolean
    CitasWebQuitarBloqueo = False
    Dim mo_Procesos As New SIGHProxies.Procesos
    Dim lcMensaje As String, lbSeTerminaSistema As Boolean, oRsTmp1 As New Recordset
    lcMensaje = ""
    Set mo_Procesos = Nothing
    Set oRsTmp1 = Nothing
    If lcMensaje <> "" Then
       MsgBox lcMensaje, vbInformation, ""
    Else
       CitasWebQuitarBloqueo = True
    End If
End Function
'debb-18/05/2019
Function CitasWebEstaBloqueda() As Boolean
    CitasWebEstaBloqueda = False
    Dim mo_Procesos As New SIGHProxies.Procesos
    Dim lcMensaje As String, lbSeTerminaSistema As Boolean, oRsTmp1 As New Recordset
    lcMensaje = ""
    Set mo_Procesos = Nothing
    Set oRsTmp1 = Nothing
    If lcMensaje <> "" Then
       MsgBox lcMensaje, vbInformation, ""
       CitasWebEstaBloqueda = True
    End If

End Function


Sub CargaCitasWebSolicitadas()
    

    Dim oExportaCitasWeb As New Procesos
'    Dim oRsTmp1 As New Recordset
'    Set oRsTmp1 = oExportaCitasWeb.CitasWebLista(CupoConfirmadoEnCitaWeb)
    If oRsCitasWebPorImportar.State = 1 Then
       Set oRsCitasWebPorImportar = Nothing
    End If
    With oRsCitasWebPorImportar
          .Fields.Append "Seleccionar", adBoolean
          .Fields.Append "Fecha", adDate
          .Fields.Append "HoraInicio", adVarChar, 5
          .Fields.Append "Paciente", adVarChar, 100
          .Fields.Append "Dni", adVarChar, 8
          .Fields.Append "Consultorio", adVarChar, 150
          .Fields.Append "ver", adVarChar, 20                          'debb-21/02/2019
          .Fields.Append "EessReferencia", adVarChar, 150
          .Fields.Append "idCitaBloqueada", adInteger
          .Fields.Append "dx", adVarChar, 250
          .LockType = adLockOptimistic
          .Open
    End With
    
    Dim mo_Procesos As New SIGHProxies.Procesos
    Dim lcMensaje As String, lbSeTerminaSistema As Boolean, oRsCitasWeb As New Recordset, lcUsuario As String
    lcUsuario = "soloCarga"   'Esto indica que solo hay que poner "sigh_Externa.citasWebCupos.idEstadoCitaWeb=3  (con Cita en la WEB)
    lcMensaje = ""
    Me.txtCitasRechazadas.Text = mo_Procesos.MensajeError
    Set oRsCitasWeb = Nothing
    Set mo_Procesos = Nothing
    
    
    Set grdCitasWebPorImportar.DataSource = oRsCitasWebPorImportar
    mo_Apariencia.ConfigurarFilasBiColores grdCitasWebPorImportar, sighEntidades.GrillaConFilasBicolor
    Set oExportaCitasWeb = Nothing
 '   Set oRsTmp1 = Nothing
End Sub

Sub ManualOautomatica()
    If Me.chkSeImportacionAutomatica.Value = 1 Then
       optMINSA.Value = True
       grdCitasWebPorImportar.Visible = False
       Set oRsCitasWebPorImportar = Nothing
    Else
       optManual.Value = True
       grdCitasWebPorImportar.Visible = True
       CargaCitasWebSolicitadas
    End If
    sighEntidades.Parametro503valorInt = IIf(Me.chkSeImportacionAutomatica.Value = 1, "1", "0")
End Sub

Sub CargaCSreferencias()
    On Error GoTo ErrCargCS
    mo_Formulario.HabilitarDeshabilitar FraAgregar, False
    mo_Formulario.HabilitarDeshabilitar txtCS, False
    
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oRsTmp3 As New Recordset
    Dim lbAgrega As Boolean
    If oRsCSreferencias.State = 1 Then Set oRsCSreferencias = Nothing
    With oRsCSreferencias
          .Fields.Append "EessCodigo", adVarChar, 20
          .Fields.Append "Eess", adVarChar, 100
          .Fields.Append "Usuario", adVarChar, 20
          .Fields.Append "Clave", adVarChar, 20
          .Fields.Append "Email", adVarChar, 50
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdCSreferir.DataSource = oRsCSreferencias
    mo_Apariencia.ConfigurarFilasBiColores Me.grdCSreferir, sighEntidades.GrillaConFilasBicolor
    
    With oRsCSreferConsul
          
          .Fields.Append "refEess", adVarChar, 20
          .Fields.Append "Servicio", adVarChar, 100
          .Fields.Append "eess", adVarChar, 100
          .Fields.Append "idServicio", adInteger
          .Fields.Append "maximoCupos", adInteger
          .LockType = adLockOptimistic
          .Open
    End With
    oRsCSreferConsul.Sort = "Servicio,eess"
    Set grdConsultorios.DataSource = oRsCSreferConsul
    mo_Apariencia.ConfigurarFilasBiColores Me.grdConsultorios, sighEntidades.GrillaConFilasBicolor
    
    Dim lcMensaje As String, lbSeTerminaSistema As Boolean
    lcMensaje = ""
    If lcMensaje <> "" Then
       MsgBox lcMensaje, vbInformation, ""
    End If
    If oRsTmp1.RecordCount > 0 Then
       lbAgrega = True
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          Set oRsTmp2 = mo_ReglasComunes.EstablecimientosSeleccionarPorFiltro("codigo='" & oRsTmp1!refEess & "'")
          If oRsTmp2.RecordCount > 0 Then
             oRsCSreferencias.AddNew
             oRsCSreferencias!EessCodigo = oRsTmp1!refEess
             oRsCSreferencias!Eess = Left(oRsTmp2!nombre, 100)
             oRsCSreferencias!Usuario = oRsTmp1!refEessUsuario
             oRsCSreferencias!Clave = oRsTmp1!refEessClave
             oRsCSreferencias!Email = IIf(IsNull(oRsTmp1!refEessEmail), "", oRsTmp1!refEessEmail)
             oRsCSreferencias.Update
          End If
          oRsTmp2.Close
          oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    
    lcMensaje = ""
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          Set oRsTmp3 = mo_ReglasComunes.EstablecimientosSeleccionarPorFiltro("codigo='" & oRsTmp1!refEess & "'")
          Set oRsTmp2 = mo_ReglasFacturacion.ServiciosSeleccionarPorFiltro("idServicio=" & oRsTmp1!IdServicio, sghPorCodigo)
          If oRsTmp2.RecordCount > 0 And oRsTmp3.RecordCount > 0 Then
             oRsCSreferConsul.AddNew
             oRsCSreferConsul!Eess = Left(oRsTmp3!nombre, 100)
             oRsCSreferConsul!refEess = oRsTmp1!refEess
             oRsCSreferConsul!Servicio = Left(oRsTmp2!nombre, 100)
             oRsCSreferConsul!IdServicio = oRsTmp1!IdServicio
             oRsCSreferConsul!maximoCupos = oRsTmp1!maximoCupos
             oRsCSreferConsul.Update
          End If
          oRsTmp2.Close
          oRsTmp3.Close
          oRsTmp1.MoveNext
       Loop

    End If
    oRsTmp1.Close
    optPorConsultorio_Click 1
    
ErrCargCS:
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oRsTmp3 = Nothing
End Sub

Sub CargaDatosDeCuposWeb()
    If oRsProgramacionDelMes.State = 1 Then Set oRsProgramacionDelMes = Nothing
    With oRsProgramacionDelMes
          .Fields.Append "idServicio", adInteger, 4, adFldIsNullable
          .Fields.Append "Servicio", adVarChar, 100, adFldIsNullable
          .Fields.Append "dia1", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia2", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia3", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia4", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia5", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia6", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia7", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia8", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia9", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia10", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia11", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia12", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia13", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia14", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia15", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia16", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia17", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia18", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia19", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia20", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia21", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia22", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia23", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia24", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia25", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia26", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia27", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia28", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia29", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia30", adVarChar, 2, adFldIsNullable
          .Fields.Append "dia31", adVarChar, 2, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdProgramacionDelMes.DataSource = oRsProgramacionDelMes
    mo_Apariencia.ConfigurarFilasBiColores Me.grdProgramacionDelMes, sighEntidades.GrillaConFilasBicolor
    grdProgramacionDelMes.Caption = ""
    '
    Set oRsCuposWeb = Nothing
    With oRsCuposWeb
          .Fields.Append "Fecha", adDate
          .Fields.Append "idServicio", adInteger, 4, adFldIsNullable
          .Fields.Append "idMedico", adInteger, 4, adFldIsNullable
          .Fields.Append "HoraInicio", adVarChar, 5, adFldIsNullable
          .Fields.Append "HoraFinal", adVarChar, 5, adFldIsNullable
          .Fields.Append "idEstadoCitaWeb", adInteger, 4, adFldIsNullable
          .Fields.Append "idCitaBloqueada", adInteger, 4, adFldIsNullable
          .Fields.Append "DNI", adVarChar, 8, adFldIsNullable
          .Fields.Append "apellidoPaterno", adVarChar, 40, adFldIsNullable
          .Fields.Append "apellidoMaterno", adVarChar, 40, adFldIsNullable
          .Fields.Append "PrimerNombre", adVarChar, 40, adFldIsNullable
          .Fields.Append "SegundoNombre", adVarChar, 40, adFldIsNullable
          .Fields.Append "idTipoSexo", adInteger, 4, adFldIsNullable
          .Fields.Append "FechaNacimiento", adDate, 8, adFldIsNullable
          .Fields.Append "Ubigeo", adInteger, 4, adFldIsNullable
          .Fields.Append "FechaConfirmacion", adDate, 8, adFldIsNullable
          .Fields.Append "HoraConfirmacion", adVarChar, 5, adFldIsNullable
          .Fields.Append "idFuenteFinanciamiento", adInteger, 4, adFldIsNullable
          .Fields.Append "idTurno", adInteger, 4, adFldIsNullable
          .Fields.Append "idPaciente", adInteger, 4, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    
    'CargaCitasWEBYaAsignadasEnGalenhos                'debb-18/05/2019
End Sub

Sub CargaCitasWEBYaAsignadasEnGalenhos()
    Dim oExportaCitasWeb As New Procesos
    Set Me.grdCitasWeb.DataSource = oExportaCitasWeb.CitasWebLista(CupoConfirmadoYconCitaEnGalenhos)
    mo_Apariencia.ConfigurarFilasBiColores Me.grdCitasWeb, sighEntidades.GrillaConFilasBicolor
    Set oExportaCitasWeb = Nothing
End Sub



Private Sub grdCitasWeb_ClickCellButton(ByVal Cell As UltraGrid.SSCell)   'debb-21/02/2019
       On Error GoTo errVerImg
       Me.MousePointer = 11
       Dim lcArchivoImagenFinal  As String, lcParametro236 As String
        lcParametro236 = lcBuscaParametro.SeleccionaFilaParametro(236)
        lcArchivoImagenFinal = lcParametro236 & IIf(Right(lcParametro236, 1) = "\", "", "\") & "jpg" & Trim(Str(grdCitasWeb.ActiveRow.Cells("idCitaBloqueada").Value)) & ".jpg"
        If Len(lcArchivoImagenFinal) > 0 Then
           FileCopy lcArchivoImagenFinal, "c:\dibujo1.jpg"
           Dim oCargaImg As Long
           oCargaImg = Shell("rundll32.exe url.dll,FileProtocolHandler " & "c:\dibujo1.jpg", vbMaximizedFocus)
        End If
        Me.MousePointer = 1
        Exit Sub
errVerImg:
       Me.MousePointer = 1
       MsgBox "No se registro IMAGEN en la CITA WEB", vbInformation, ""
End Sub


Private Sub grdCitasWeb_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdCitasWeb.Bands(0).Columns("ver").Style = ssStyleButton     'debb-21/02/2019
    grdCitasWeb.Bands(0).Columns("ver").Width = 500     'debb-21/02/2019
End Sub

'debb-18/05/2019
Private Sub grdCitasWebPorImportar_ClickCellButton(ByVal Cell As UltraGrid.SSCell)    'debb-21/02/2019
       On Error GoTo errVerImg
       Me.MousePointer = 11
       Dim lcArchivoImagenFinal  As String, lcParametro236 As String
        lcParametro236 = lcBuscaParametro.SeleccionaFilaParametro(236)
        lcArchivoImagenFinal = lcParametro236 & IIf(Right(lcParametro236, 1) = "\", "", "\") & "jpg" & Trim(Str(grdCitasWebPorImportar.ActiveRow.Cells("idCitaBloqueada").Value)) & ".jpg"
        If Len(lcArchivoImagenFinal) > 0 Then
           FileCopy lcArchivoImagenFinal, "c:\dibujo1.jpg"
           Dim oCargaImg As Long
           oCargaImg = Shell("rundll32.exe url.dll,FileProtocolHandler " & "c:\dibujo1.jpg", vbMaximizedFocus)
        Else
           MsgBox "No se registro IMAGEN en la CITA WEB", vbInformation, ""
        End If
        Me.MousePointer = 1
        Exit Sub
errVerImg:
        Me.MousePointer = 1
        MsgBox "No se registro IMAGEN en la CITA WEB", vbInformation, ""
End Sub

Private Sub grdCitasWebPorImportar_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    'grdCitasWebPorImportar.Bands(0).Columns("idCitaBloqueada").Hidden = True
    grdCitasWebPorImportar.Bands(0).Columns("Paciente").Width = 3000
    grdCitasWebPorImportar.Bands(0).Columns("Consultorio").Width = 3000
    grdCitasWebPorImportar.Bands(0).Columns("EessReferencia").Width = 3000
    grdCitasWebPorImportar.Bands(0).Columns("Seleccionar").Width = 500
    grdCitasWebPorImportar.Bands(0).Columns("HoraInicio").Width = 700
    grdCitasWebPorImportar.Bands(0).Columns("ver").Style = ssStyleButton     'debb-21/02/2019
    grdCitasWebPorImportar.Bands(0).Columns("ver").Width = 500     'debb-21/02/2019
End Sub

Private Sub grdConsultorios_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
Cancel = True
End Sub

Private Sub grdConsultorios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdConsultorios.Bands(0).Columns("refEess").Hidden = True
    grdConsultorios.Bands(0).Columns("IdServicio").Hidden = True
    grdConsultorios.Bands(0).Columns("eess").Width = 2300
    grdConsultorios.Bands(0).Columns("eess").Activation = ssActivationActivateNoEdit
    grdConsultorios.Bands(0).Columns("Servicio").Width = 2700
    grdConsultorios.Bands(0).Columns("Servicio").Activation = ssActivationActivateNoEdit
    grdConsultorios.Bands(0).Columns("Servicio").Header.Caption = "Consultorios"
    grdConsultorios.Bands(0).Columns("maximoCupos").Width = 900
    grdConsultorios.Bands(0).Columns("maximoCupos").Header.Caption = "Max.Cupos"
    
End Sub

Private Sub grdCSreferir_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
Cancel = True
End Sub

Private Sub grdCSreferir_DblClick()
'    On Error GoTo errGrCS
'    grdConsultorios.Caption = "CUPOS para " & oRsCSreferencias!Eess
'    oRsCSreferConsul.Filter = "refEss='" & oRsCSreferencias!eessCodigo & "'"
'    Set grdConsultorios.DataSource = oRsCSreferConsul
'errGrCS:
    
End Sub

Private Sub grdCSreferir_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdCSreferir.Bands(0).Columns("EessCodigo").Hidden = True
    '
    grdCSreferir.Bands(0).Columns("Eess").Width = 3000
    grdCSreferir.Bands(0).Columns("Eess").Activation = ssActivationActivateNoEdit
    grdCSreferir.Bands(0).Columns("Usuario").Width = 1000
    grdCSreferir.Bands(0).Columns("Usuario").Activation = ssActivationActivateNoEdit
    grdCSreferir.Bands(0).Columns("Clave").Width = 1000
    grdCSreferir.Bands(0).Columns("Clave").Activation = ssActivationActivateNoEdit
    grdCSreferir.Bands(0).Columns("Email").Width = 3000
    
End Sub

Private Sub grdProgramacionDelMes_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub


Function ElConsultorioTieneTODosLosPreciosDeConsulta(lnIdServicio As Long) As Boolean
        ElConsultorioTieneTODosLosPreciosDeConsulta = True
        Dim oRs1 As New Recordset
        Dim oRs As New Recordset
        Dim oRs2 As New Recordset
        Dim lcMensaje1 As String
        Set oRs1 = mo_ReglasComunes.FuentesFinanciamientoTarifasSoloTarifa
        Set oRs = mo_ReglasServiciosHosp.EspecialidadCESeleccionarPorIdServicio(lnIdServicio)
        lcMensaje1 = ""
        If oRs.RecordCount = 0 Then
           lcMensaje1 = lcMensaje1 & "PARA EL CONSULTORIO no hay TARIFA de PAGO DE CITAS EN CONSULTORIOS EXTERNOS" & Chr(13) & _
                                     "Chequee opción:   general->especialidades" & Chr(13)
        Else
           oRs.MoveFirst
           Do While Not oRs.EOF
              oRs1.MoveFirst
              Do While Not oRs1.EOF
                 Set oRs2 = mo_ReglasFacturacion.FactCatalogoServiciosHospSeleccionarPorIdYtipoFinanciamiento(oRs!idProductoConsulta, _
                                                                                                              oRs1!idTipoFinanciamiento)
                 If oRs2.RecordCount = 0 Then
                    lcMensaje1 = lcMensaje1 & "(cpt: " & oRs!Codigo & ") falta TARIFA (opción: FACT-CONFIG->CAT.SERVICIOS)" & Chr(13)
                 Else
                    If oRs2!PrecioUnitario <= 0 Then
                       If oRs2!SeUsaSinPrecio = 0 Then
                          lcMensaje1 = lcMensaje1 & "(cpt: " & oRs!Codigo & ") El PRECIO UNITARIO es menor o igual a CERO o no marcó SE USA SIN PRECIO (opción: FACT-CONFIG->CAT.SERVICIOS)" & Chr(13)
                       End If
                    End If
                 End If
                 oRs1.MoveNext
              Loop
              oRs.MoveNext
           Loop
        End If
        Set oRs1 = Nothing
        Set oRs = Nothing
        Set oRs2 = Nothing
        If lcMensaje1 <> "" Then
           MsgBox "El CONSULTORIO tiene problemas:" & Chr(13) & lcMensaje1, vbInformation, ""
           ElConsultorioTieneTODosLosPreciosDeConsulta = False
        End If
End Function

Private Sub grdProgramacionDelMes_DblClick()
    On Error Resume Next
    Dim oRow As SSRow, oHerrExportaCitasDia As New HerrExportaCitasDia
    Dim oRsTmp1 As New Recordset, ldFecha As Date, lnNroCuposElegidos As Long
    Set oRow = grdProgramacionDelMes.ActiveCell.Row
    oHerrExportaCitasDia.IdServicio = oRow.Cells("idServicio").Value
    
    If ElConsultorioTieneTODosLosPreciosDeConsulta(Val(oRow.Cells("idServicio").Value)) = False Then
       Set oHerrExportaCitasDia = Nothing
       Set oRsTmp1 = Nothing
       Exit Sub
    End If
    
    Dim lnIdServicio As Long
    lnIdServicio = oRow.Cells("idServicio").Value
    
    
    Select Case grdProgramacionDelMes.ActiveCell.Column.Key
    Case "dia1"
        If (oRow.Cells("dia1").Value = lcEquix Or Val(oRow.Cells("dia1").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("01" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("01/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia1").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia2"
        If (oRow.Cells("dia2").Value = lcEquix Or Val(oRow.Cells("dia2").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("02" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("02/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia2").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia3"
        If (oRow.Cells("dia3").Value = lcEquix Or Val(oRow.Cells("dia3").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("03" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("03/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia3").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia4"
        If (oRow.Cells("dia4").Value = lcEquix Or Val(oRow.Cells("dia4").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("04" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("04/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia4").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia5"
        If (oRow.Cells("dia5").Value = lcEquix Or Val(oRow.Cells("dia5").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("05" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("05/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia5").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia6"
        If (oRow.Cells("dia6").Value = lcEquix Or Val(oRow.Cells("dia6").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("06" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("06/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia6").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia7"
        If (oRow.Cells("dia7").Value = lcEquix Or Val(oRow.Cells("dia7").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("07" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("07/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia7").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia8"
        If (oRow.Cells("dia8").Value = lcEquix Or Val(oRow.Cells("dia8").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("08" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("08/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia8").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia9"
        If (oRow.Cells("dia9").Value = lcEquix Or Val(oRow.Cells("dia9").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("09" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("09/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia9").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia10"
        If (oRow.Cells("dia10").Value = lcEquix Or Val(oRow.Cells("dia10").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("10" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("10/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia10").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia11"
        If (oRow.Cells("dia11").Value = lcEquix Or Val(oRow.Cells("dia11").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("11" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("11/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia11").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia12"
        If (oRow.Cells("dia12").Value = lcEquix Or Val(oRow.Cells("dia12").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("12" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("12/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia12").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia13"
        If (oRow.Cells("dia13").Value = lcEquix Or Val(oRow.Cells("dia13").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("13" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("13/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia13").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia14"
        If (oRow.Cells("dia14").Value = lcEquix Or Val(oRow.Cells("dia14").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("14" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("14/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia14").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia15"
        If (oRow.Cells("dia15").Value = lcEquix Or Val(oRow.Cells("dia15").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("15" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("15/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia15").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia16"
        If (oRow.Cells("dia16").Value = lcEquix Or Val(oRow.Cells("dia16").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("16" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("16/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia16").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia17"
        If (oRow.Cells("dia17").Value = lcEquix Or Val(oRow.Cells("dia17").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("17" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("17/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia17").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia18"
        If (oRow.Cells("dia18").Value = lcEquix Or Val(oRow.Cells("dia18").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("18" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("18/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia18").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia19"
        If (oRow.Cells("dia19").Value = lcEquix Or Val(oRow.Cells("dia19").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("19" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("19/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia19").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia20"
        If (oRow.Cells("dia20").Value = lcEquix Or Val(oRow.Cells("dia20").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("20" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("20/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia20").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia21"
        If (oRow.Cells("dia21").Value = lcEquix Or Val(oRow.Cells("dia21").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("21" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("21/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia21").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia22"
        If (oRow.Cells("dia22").Value = lcEquix Or Val(oRow.Cells("dia22").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("22" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("22/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia22").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia23"
        If (oRow.Cells("dia23").Value = lcEquix Or Val(oRow.Cells("dia23").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("23" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("23/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia23").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia24"
        If (oRow.Cells("dia24").Value = lcEquix Or Val(oRow.Cells("dia24").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("24" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("24/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia24").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia25"
        If (oRow.Cells("dia25").Value = lcEquix Or Val(oRow.Cells("dia25").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("25" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("25/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia25").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia26"
        If (oRow.Cells("dia26").Value = lcEquix Or Val(oRow.Cells("dia26").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("26" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("26/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia26").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia27"
        If (oRow.Cells("dia27").Value = lcEquix Or Val(oRow.Cells("dia27").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("27" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("27/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia27").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia28"
        If (oRow.Cells("dia28").Value = lcEquix Or Val(oRow.Cells("dia28").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("28" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("28/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia28").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia29"
        If (oRow.Cells("dia29").Value = lcEquix Or Val(oRow.Cells("dia29").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("29" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("29/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia29").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia30"
        If (oRow.Cells("dia30").Value = lcEquix Or Val(oRow.Cells("dia30").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("30" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("30/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia30").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    Case "dia31"
        If (oRow.Cells("dia31").Value = lcEquix Or Val(oRow.Cells("dia31").Value) > 0) And DateDiff("d", ldFechaHoy, CDate("31" & lcMesAnio)) >= 0 Then
           ldFecha = CDate("31/" & Right("0" & Trim(Str(Me.cmbRangoMeses.ListIndex + 1)), 2) & "/" & Me.cmbAnio.Text)
           oRsCuposWeb.Filter = ""
           oRsCuposWeb.Filter = "fecha='" & Format(ldFecha, "dd/mm/yyyy") & "' and idServicio=" & oRow.Cells("idServicio").Value
           oHerrExportaCitasDia.fecha = ldFecha
           Set oHerrExportaCitasDia.CuposDelDia = oRsCuposWeb
           oHerrExportaCitasDia.Show 1
           lnNroCuposElegidos = oHerrExportaCitasDia.NroCuposElegidos
           oRow.Cells("dia31").Value = IIf(lnNroCuposElegidos > 0, Trim(Str(lnNroCuposElegidos)), lcEquix)
        End If
    End Select
    '
    If lbEsPrimeraVezQgrabaElMes = False And oHerrExportaCitasDia.BotonPresionado = sghAceptar Then
        Me.MousePointer = 11
        Dim oExportaCitasWeb As New Procesos
        Set oExportaCitasWeb.progressRpt1 = Me.XP_ProgressBar1
        Set oExportaCitasWeb.progressRpt2 = Me.XP_ProgressBar2
        Set oExportaCitasWeb.progressRpt3 = Me.XP_ProgressBar3
        Set oExportaCitasWeb.progressRpt4 = Me.XP_ProgressBar4
        Set oExportaCitasWeb.progressRpt5 = Me.XP_ProgressBar5
        Set oExportaCitasWeb.progressRpt6 = Me.XP_ProgressBar6
        If oRsCuposWeb.RecordCount > 0 Then
             If oExportaCitasWeb.ExportaCitasWeb(ldFecha, ldFecha, oRsCuposWeb, ml_idUsuario, _
                                                    ml_IdPaciente, False, False, False, False, False, lnIdServicio) = True Then
             End If
        End If
        Set oExportaCitasWeb = Nothing
        Me.MousePointer = 1
    End If
    '
    Set oHerrExportaCitasDia = Nothing

End Sub

Private Sub grdProgramacionDelMes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdProgramacionDelMes.Bands(0).Columns("idServicio").Hidden = True
    '
    grdProgramacionDelMes.Bands(0).Columns("servicio").Width = 3700
    grdProgramacionDelMes.Bands(0).Columns("servicio").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia1").Header.Caption = "1"
    grdProgramacionDelMes.Bands(0).Columns("dia1").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia1").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia2").Header.Caption = "2"
    grdProgramacionDelMes.Bands(0).Columns("dia2").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia2").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia3").Header.Caption = "3"
    grdProgramacionDelMes.Bands(0).Columns("dia3").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia3").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia4").Header.Caption = "4"
    grdProgramacionDelMes.Bands(0).Columns("dia4").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia4").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia5").Header.Caption = "5"
    grdProgramacionDelMes.Bands(0).Columns("dia5").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia5").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia6").Header.Caption = "6"
    grdProgramacionDelMes.Bands(0).Columns("dia6").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia6").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia7").Header.Caption = "7"
    grdProgramacionDelMes.Bands(0).Columns("dia7").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia7").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia8").Header.Caption = "8"
    grdProgramacionDelMes.Bands(0).Columns("dia8").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia8").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia9").Header.Caption = "9"
    grdProgramacionDelMes.Bands(0).Columns("dia9").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia9").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia10").Header.Caption = "10"
    grdProgramacionDelMes.Bands(0).Columns("dia10").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia10").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia11").Header.Caption = "11"
    grdProgramacionDelMes.Bands(0).Columns("dia11").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia11").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia12").Header.Caption = "12"
    grdProgramacionDelMes.Bands(0).Columns("dia12").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia12").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia13").Header.Caption = "13"
    grdProgramacionDelMes.Bands(0).Columns("dia13").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia13").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia14").Header.Caption = "14"
    grdProgramacionDelMes.Bands(0).Columns("dia14").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia14").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia15").Header.Caption = "15"
    grdProgramacionDelMes.Bands(0).Columns("dia15").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia15").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia16").Header.Caption = "16"
    grdProgramacionDelMes.Bands(0).Columns("dia16").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia16").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia17").Header.Caption = "17"
    grdProgramacionDelMes.Bands(0).Columns("dia17").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia17").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia18").Header.Caption = "18"
    grdProgramacionDelMes.Bands(0).Columns("dia18").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia18").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia19").Header.Caption = "19"
    grdProgramacionDelMes.Bands(0).Columns("dia19").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia19").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia20").Header.Caption = "20"
    grdProgramacionDelMes.Bands(0).Columns("dia20").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia20").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia21").Header.Caption = "21"
    grdProgramacionDelMes.Bands(0).Columns("dia21").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia21").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia22").Header.Caption = "22"
    grdProgramacionDelMes.Bands(0).Columns("dia22").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia22").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia23").Header.Caption = "23"
    grdProgramacionDelMes.Bands(0).Columns("dia23").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia23").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia24").Header.Caption = "24"
    grdProgramacionDelMes.Bands(0).Columns("dia24").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia24").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia25").Header.Caption = "25"
    grdProgramacionDelMes.Bands(0).Columns("dia25").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia25").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia26").Header.Caption = "26"
    grdProgramacionDelMes.Bands(0).Columns("dia26").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia26").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia27").Header.Caption = "27"
    grdProgramacionDelMes.Bands(0).Columns("dia27").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia27").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia28").Header.Caption = "28"
    grdProgramacionDelMes.Bands(0).Columns("dia28").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia28").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia29").Header.Caption = "29"
    grdProgramacionDelMes.Bands(0).Columns("dia29").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia29").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia30").Header.Caption = "30"
    grdProgramacionDelMes.Bands(0).Columns("dia30").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia30").Activation = ssActivationActivateNoEdit
    grdProgramacionDelMes.Bands(0).Columns("dia31").Header.Caption = "31"
    grdProgramacionDelMes.Bands(0).Columns("dia31").Width = 300
    grdProgramacionDelMes.Bands(0).Columns("dia31").Activation = ssActivationActivateNoEdit
End Sub

Private Sub grdProgramacionDelMes_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdProgramacionDelMes_DblClick
    End If
End Sub





Private Sub optCsPS_Click(Value As Integer)
    If optCsPS.Value = False Then
       oRsCSreferConsul.Sort = "Servicio,eess"
    Else
       oRsCSreferConsul.Sort = "eess,Servicio"
    End If
    oRsCSreferConsul.MoveFirst
    Set grdConsultorios.DataSource = oRsCSreferConsul

End Sub

Private Sub optPorConsultorio_Click(Value As Integer)
    If optPorConsultorio.Value = True Then
       oRsCSreferConsul.Sort = "Servicio,eess"
    Else
       oRsCSreferConsul.Sort = "eess,Servicio"
    End If
    oRsCSreferConsul.MoveFirst
    Set grdConsultorios.DataSource = oRsCSreferConsul
End Sub

Private Sub txtBusqAMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtBusqAMaterno
End Sub

Private Sub txtBusqApaterno_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtBusqApaterno
End Sub

Private Sub txtCSclave_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCSclave
End Sub



Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtEmail
End Sub



Private Sub txtCSusuario_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCSusuario
End Sub



